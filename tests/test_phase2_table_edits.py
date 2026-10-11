"""ทดสอบเฟส 2 PR2 (แก้โครงสร้าง Excel: เปลี่ยนอัตรา / แก้-ลบ-เพิ่มรายการในตาราง / ขั้นยืนยัน) และ 7 จุดที่แก้
เพิ่มจากรายงานบั๊ก ใช้ fixture workbook สังเคราะห์ (ไม่ใช่ไฟล์ลูกค้าจริง) และ fake DB ในหน่วยความจำล้วนๆ
(ไม่ต่อ DB จริงเลย) ผ่าน tests/_harness.py"""
import io
import datetime
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import pytest
import openpyxl

from _harness import (
    build_namespace, build_fixture_workbook, inject_formula_cache,
    FIXTURE_TOTAL, FIXTURE_VAT, FIXTURE_GRAND_TOTAL,
)


@pytest.fixture
def ns():
    return build_namespace()


@pytest.fixture
def doc1(ns):
    """สร้างเอกสารทดสอบ (document_id=1, chat_id=1, user_id=1) จาก fixture workbook ไว้ใน fake db"""
    raw = build_fixture_workbook()
    label_map = ns["_extract_excel_labels"](raw)
    ns["_fake_db"].add_document(1, "quote.xlsx", raw, label_map, chat_id=1, user_id=1)
    return ns["_fake_db"].documents[1]


# ---------- change_rate ----------

def test_change_rate_to_10_percent(ns, doc1):
    result = ns["propose_change_rate"](1, 1, 1, "ภาษีมูลค่าเพิ่ม 7%", 10)
    assert result["proposed"] is True
    assert "10%" in result["message"]

    confirmed = ns["apply_change_rate"](1, 1, "ภาษีมูลค่าเพิ่ม 7%", 10)
    assert confirmed["matched"] is True
    expected_vat = round(FIXTURE_TOTAL * 0.10, 2)
    assert confirmed["new_value"] == expected_vat
    assert str(FIXTURE_VAT) in confirmed["message"] or str(expected_vat) in confirmed["message"]
    # รวมราคาทั้งสิ้น (side effect) ต้องเปลี่ยนตามไปด้วย
    expected_total = FIXTURE_TOTAL + expected_vat
    assert str(expected_total) in confirmed["message"]


def test_change_rate_unsupported_formula_is_rejected(ns, doc1):
    # แก้สูตรของ "รวมเงิน" (=SUM(G16:G25)) ให้ไม่ตรงรูปแบบอัตราที่รองรับเลย
    result = ns["propose_change_rate"](1, 1, 1, "รวมเงิน", 10)
    assert result["proposed"] is False
    assert "ไม่ใช่รูปแบบอัตราที่รองรับ" in result["message"]


def test_change_rate_position_based_lookup_matches_real_file_scale(ns):
    """ทวนด้วยตัวเลขจากไฟล์จริง (ป่าไม้เขียว, PR #25): รวมเงิน 58150, VAT 7%->10% ต้องได้ 5815 และรวม
    ราคาทั้งสิ้น 58150+5815=63965 — ยืนยันว่าหา "ค่าใหม่" ถูกต้องด้วยตำแหน่งเซลล์ แม้ label มี % อยู่ในชื่อ
    (ไม่ต้อง inject cache เพราะ apply_change_rate คำนวณค่าใหม่ผ่าน pycel เสมอ ไม่พึ่งค่า cache เริ่มต้นเลย)"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Quotation"
    ws["A26"] = "รวมเงิน"
    ws["G26"] = 58150  # ค่าคงที่ตรงๆ (ไม่ใช่สูตร) พอสำหรับทดสอบจุดนี้
    ws["A27"] = "ภาษีมูลค่าเพิ่ม 7%"
    ws["G27"] = "=G26*7/100"
    ws["A28"] = "รวมราคาทั้งสิ้น"
    ws["G28"] = "=SUM(G26:G27)"
    buf = io.BytesIO()
    wb.save(buf)
    # G27/G28 เป็นสูตรที่ไม่เคยถูกคำนวณจริง (openpyxl ล้วนๆ ไม่มี cache) — ต้อง inject cache ก่อน ไม่งั้น
    # _extract_excel_labels (data_only=True) อ่านได้ None ทำให้แถว 27/28 ถูกเข้าใจผิดเป็น "หัวข้อหมวด"
    # (มีแค่ A27/A28 ไม่ว่าง) แทนที่จะเป็นคู่ label:value ปกติ
    raw = inject_formula_cache(buf.getvalue(), {"G27": 4070.5, "G28": 62220.5})
    label_map = ns["_extract_excel_labels"](raw)
    ns["_fake_db"].add_document(2, "real_scale.xlsx", raw, label_map, chat_id=1, user_id=1)

    confirmed = ns["apply_change_rate"](2, 1, "ภาษีมูลค่าเพิ่ม 7%", 10)
    assert confirmed["matched"] is True
    assert confirmed["new_value"] == 5815.0
    assert "63965" in confirmed["message"]


# ---------- update_table_item ----------

def test_update_table_item_description(ns, doc1):
    result = ns["apply_update_table_item"](1, 1, 3, {"description": "Item C renamed"})
    assert result["matched"] is True
    assert "Item C renamed" in result["message"]


# ---------- delete_table_item ----------

def test_delete_middle_item_shifts_and_renumbers_continuously(ns, doc1):
    """ลบรายการที่ 3 จาก 8 รายการ: ลำดับต้องเป็น 1..7 ต่อเนื่อง (ไม่ใช่ 1,2,1,2,3,4,5 แบบบั๊กเดิม) และ
    รายละเอียดต้องเลื่อนขึ้นถูกแถว (รายการที่ 3 ใหม่ต้องเป็นของเดิมรายการที่ 4)"""
    propose = ns["propose_delete_table_item"](1, 1, 1, 3)
    assert propose["proposed"] is True

    result = ns["apply_delete_table_item"](1, 1, 3)
    assert result["matched"] is True

    updated_doc = ns["get_editable_document_raw"](1)
    data = ns["read_document_values"](updated_doc)
    table = data["tables"][0]
    filled = table["rows"]  # read_document_values คืนเฉพาะแถวที่ไม่ว่างอยู่แล้ว (กรองให้แล้ว ไม่มี is_empty)

    assert len(filled) == 7
    assert [r["values"]["no"] for r in filled] == [1, 2, 3, 4, 5, 6, 7]
    # รายการที่ 3 (index 2) ตอนนี้ควรเป็นของเดิม "Item D" (เดิมรายการที่ 4)
    assert filled[2]["values"]["description"] == "Item D"
    assert filled[2]["values"]["qty"] == 2
    assert filled[2]["values"]["unit_price"] == 300


def test_undo_after_delete_restores_original(ns, doc1):
    ns["apply_delete_table_item"](1, 1, 3)
    undo_result = ns["undo_last_edit"](1, 1)
    assert undo_result["matched"] is True

    updated_doc = ns["get_editable_document_raw"](1)
    data = ns["read_document_values"](updated_doc)
    table = data["tables"][0]
    filled = table["rows"]
    assert len(filled) == 8
    assert filled[2]["values"]["description"] == "Item C"


def test_delete_rejected_when_amount_not_formula(ns):
    raw_wb = openpyxl.Workbook()
    ws = raw_wb.active
    ws.title = "Quotation"
    ws["A15"] = "ลำดับ"
    ws["B15"] = "รายการ"
    ws["D15"] = "จำนวน"
    ws["E15"] = "หน่วย"
    ws["F15"] = "ราคา/หน่วย"
    ws["G15"] = "จำนวนเงิน"
    ws["A16"] = 1
    ws["B16"] = "Item X"
    ws["D16"] = 1
    ws["E16"] = "ชิ้น"
    ws["F16"] = 100
    ws["G16"] = 100  # ค่าคงที่ ไม่ใช่สูตร — ต้องถูกปฏิเสธ
    buf = io.BytesIO()
    raw_wb.save(buf)
    raw = buf.getvalue()
    label_map = ns["_extract_excel_labels"](raw)
    ns["_fake_db"].add_document(3, "bad_table.xlsx", raw, label_map, chat_id=1, user_id=1)

    result = ns["propose_delete_table_item"](3, 1, 1, 1)
    assert result["proposed"] is False
    assert result["message"] == ns["_TABLE_NON_FORMULA_AMOUNT_ERROR"]


# ---------- add_table_item ----------

def test_add_table_item_until_full(ns, doc1):
    result = ns["propose_add_table_item"](1, 1, 1, {"description": "Item I", "qty": 2, "unit": "ชิ้น", "unit_price": 20})
    assert result["proposed"] is True

    confirmed = ns["apply_add_table_item"](1, 1, {"description": "Item I", "qty": 2, "unit": "ชิ้น", "unit_price": 20})
    assert confirmed["matched"] is True

    # ตอนนี้ตารางเต็มพอดี (10 แถว: 8 เดิม + 1 ที่เพิ่ง add = 9 ยังไม่เต็ม แต่เหลือ 1 ว่าง — เติมอีกให้เต็ม)
    fill_more = ns["apply_add_table_item"](1, 1, {"description": "Item J", "qty": 1, "unit": "ชิ้น", "unit_price": 5})
    assert fill_more["matched"] is True

    # ตอนนี้เต็มแล้ว (10/10) — ลองเพิ่มอีกต้องถูกปฏิเสธ
    rejected = ns["apply_add_table_item"](1, 1, {"description": "Item K", "qty": 1, "unit": "ชิ้น", "unit_price": 1})
    assert rejected["matched"] is False
    assert "เต็มแล้ว" in rejected["message"]
    assert "10" in rejected["message"]


# ---------- ขั้นยืนยัน (ChatPendingAction) ----------

def test_confirm_without_pending_does_nothing(ns, doc1):
    pending = ns["get_chat_pending_action"](1)
    assert pending is None
    # ไม่มี pending เลย -> _is_confirm_command เป็น True ก็ไม่ควรมีอะไรให้ execute (caller เช็ค pending is None เอง)


def test_pending_action_expires_after_ttl(ns, doc1):
    ns["propose_change_rate"](1, 1, 1, "ภาษีมูลค่าเพิ่ม 7%", 10)
    pending = ns["get_chat_pending_action"](1)
    assert pending is not None

    stale_created_at = datetime.datetime.utcnow() - datetime.timedelta(seconds=ns["PENDING_ACTION_TTL_SECONDS"] + 1)
    assert ns["_is_pending_action_fresh"](stale_created_at) is False
    assert ns["_is_pending_action_fresh"](pending["created_at"]) is True


def test_execute_pending_action_rejects_if_document_changed_since_proposal(ns, doc1):
    ns["propose_delete_table_item"](1, 1, 1, 3)
    pending = ns["get_chat_pending_action"](1)

    # เอกสารถูกแก้ไปแล้วระหว่างรอยืนยัน (เช่นแก้ผ่านทางอื่น) -> version_id ไม่ตรงกับตอนเสนอ preview อีกต่อไป
    ns["apply_update_table_item"](1, 1, 1, {"description": "changed elsewhere"})

    answer = ns["_execute_pending_action"](pending["action"], 1, 1)
    assert answer == "เอกสารเปลี่ยนไปหลังจากเสนอการแก้ไข กรุณาสั่งใหม่อีกครั้งครับ"


def test_execute_pending_action_rejects_if_document_closed(ns, doc1):
    ns["propose_change_rate"](1, 1, 1, "ภาษีมูลค่าเพิ่ม 7%", 10)
    pending = ns["get_chat_pending_action"](1)

    ns["_fake_db"].close_document(1)

    answer = ns["_execute_pending_action"](pending["action"], 1, 1)
    assert answer == "เอกสารเปลี่ยนไปหลังจากเสนอการแก้ไข กรุณาสั่งใหม่อีกครั้งครับ"


def test_execute_pending_action_succeeds_when_version_unchanged(ns, doc1):
    ns["propose_change_rate"](1, 1, 1, "ภาษีมูลค่าเพิ่ม 7%", 10)
    pending = ns["get_chat_pending_action"](1)

    answer = ns["_execute_pending_action"](pending["action"], 1, 1)
    assert "เปลี่ยนอัตราของ" in answer
    assert "เอกสารเปลี่ยนไป" not in answer


# ---------- จุดที่ 1: _is_pending_action_fresh รองรับ tz-aware ----------

def test_is_pending_action_fresh_handles_naive_and_aware(ns):
    naive_now = datetime.datetime.utcnow()
    aware_now = datetime.datetime.now(datetime.timezone.utc)
    assert ns["_is_pending_action_fresh"](naive_now) is True
    assert ns["_is_pending_action_fresh"](aware_now) is True

    naive_stale = naive_now - datetime.timedelta(seconds=ns["PENDING_ACTION_TTL_SECONDS"] + 5)
    aware_stale = aware_now - datetime.timedelta(seconds=ns["PENDING_ACTION_TTL_SECONDS"] + 5)
    assert ns["_is_pending_action_fresh"](naive_stale) is False
    assert ns["_is_pending_action_fresh"](aware_stale) is False


# ---------- จุดที่ 2: confirm/cancel ปฏิเสธรูปแบบปฏิเสธ ----------

@pytest.mark.parametrize("text", ["ไม่ยืนยัน", "ยังไม่ยืนยัน"])
def test_confirm_command_rejects_negated_forms(ns, text):
    assert ns["_is_confirm_command"](text) is False


@pytest.mark.parametrize("text", ["ยืนยัน", "ยืนยันครับ", "ตกลง ยืนยัน"])
def test_confirm_command_accepts_short_real_confirm_texts(ns, text):
    assert ns["_is_confirm_command"](text) is True


def test_cancel_command_rejects_negated_form(ns):
    assert ns["_is_cancel_command"]("ไม่ยกเลิก") is False


@pytest.mark.parametrize("text", ["ยกเลิก", "ยกเลิกครับ"])
def test_cancel_command_accepts_real_cancel_texts(ns, text):
    assert ns["_is_cancel_command"](text) is True


# ---------- จุดที่ 4: คำถามเชิงอ่าน ไม่พึ่ง prompt อย่างเดียว ----------

def test_open_library_file_sets_focus_and_appears_in_active_docs(ns):
    raw = build_fixture_workbook()
    ns["_fake_db"].add_library_file(file_id=42, filename="quote.xlsx", storage_key="lib/quote.xlsx")
    ns["_fake_db"].add_storage_blob("lib/quote.xlsx", raw)

    result = ns["execute_open_library_file"]({"file_id": 42}, user_id=1, chat_id=1)
    assert result.get("opened") is True
    document_id = result["document_id"]

    assert ns["_fake_db"].focus.get(1) == document_id
    active_docs = ns["list_active_editable_documents_by_chat"](1, 1)
    assert any(d["id"] == document_id for d in active_docs)


def test_reading_question_without_legal_keyword_routes_to_document_values(ns):
    assert ns["_is_document_reading_question"]("มีรายการอะไรบ้าง") is True
    assert ns["_has_legal_keyword"]("มีรายการอะไรบ้าง") is False


def test_reading_question_with_legal_keyword_routes_to_rag(ns):
    query = "ธุรกิจบัญชีสามมีอะไรบ้าง"
    assert ns["_is_document_reading_question"](query) is True
    assert ns["_has_legal_keyword"](query) is True
    # document_content_question = reading_question and not _has_legal_keyword(query) -> False -> ไป RAG ปกติ
    document_content_question = ns["_is_document_reading_question"](query) and not ns["_has_legal_keyword"](query)
    assert document_content_question is False


def test_answer_document_reading_question_uses_read_document_values_not_rag(ns, doc1):
    active_docs = [doc1]
    answer, used_doc_id = ns["_answer_document_reading_question"]("มีรายการอะไรบ้าง", active_docs, 1, 1)
    assert used_doc_id == 1
    # client ปลอม (ดู _harness._FakeAnthropicClient) ถูกเรียกจริง และได้รับ tool_text ที่มาจาก
    # read_document_values ตรงๆ (มีตารางรายการ + รวมเงิน/VAT/รวมทั้งสิ้น) ไม่ใช่ RAG context ใดๆ เลย
    sent_messages = ns["_fake_client"].last_call_kwargs["messages"]
    sent_content = sent_messages[0]["content"]
    assert "ตารางรายการ" in sent_content
    assert "รวมเงิน" in sent_content
    assert "ภาษีมูลค่าเพิ่ม 7%" in sent_content
    assert answer == ns["_fake_client"].response_text


# ---------- จุดที่ 7 (ทวนอีกครั้งสำหรับ add_table_item) ----------

def test_add_rejected_when_amount_not_formula(ns):
    raw_wb_obj = openpyxl.Workbook()
    ws = raw_wb_obj.active
    ws.title = "Quotation"
    ws["A15"] = "ลำดับ"
    ws["B15"] = "รายการ"
    ws["D15"] = "จำนวน"
    ws["E15"] = "หน่วย"
    ws["F15"] = "ราคา/หน่วย"
    ws["G15"] = "จำนวนเงิน"
    ws["A16"] = 1
    ws["B16"] = "Item X"
    ws["D16"] = 1
    ws["E16"] = "ชิ้น"
    ws["F16"] = 100
    ws["G16"] = 100  # ค่าคงที่ ไม่ใช่สูตร
    buf = io.BytesIO()
    raw_wb_obj.save(buf)
    raw = buf.getvalue()
    label_map = ns["_extract_excel_labels"](raw)
    ns["_fake_db"].add_document(4, "bad_table2.xlsx", raw, label_map, chat_id=1, user_id=1)

    result = ns["propose_add_table_item"](4, 1, 1, {"description": "New", "qty": 1, "unit": "ชิ้น", "unit_price": 1})
    assert result["proposed"] is False
    assert result["message"] == ns["_TABLE_NON_FORMULA_AMOUNT_ERROR"]


# ---------- Review fix: โครงสร้างต้องไปทาง agent ไม่ใช่ _match_and_apply_excel_edit ----------

@pytest.mark.parametrize("query", [
    "เปลี่ยน VAT เป็น 10%",
    "เปลี่ยนอัตราภาษีเป็น 10 เปอร์เซ็นต์",
])
def test_structural_command_detects_rate_change_wording(ns, query):
    assert ns["_is_structural_edit_command"](query) is True


@pytest.mark.parametrize("query", [
    "ลบรายการที่ 3",
    "เพิ่มรายการ Box 2x2 5 ชิ้น 30 บาท",
])
def test_structural_command_detects_table_item_wording(ns, query):
    assert ns["_is_structural_edit_command"](query) is True


def test_edit_command_with_question_word_is_not_reading_question(ns):
    # มี "แก้" (คำสั่งแก้) ปนกับ "ไหม" (คำเชิงคำถาม) — ต้องไม่ใช่ reading question เด็ดขาด
    assert ns["_is_document_reading_question"]("แก้ชื่อลูกค้าเป็น บริษัท ดูดี จำกัด ได้ไหม") is False


def test_match_and_apply_excel_edit_refuses_percent_on_formula_cell(ns, doc1):
    """safety net: เซลล์เป้าหมายเป็นสูตร + ค่าใหม่มี "%" -> ห้ามเขียนทับ ต้องบอกให้ใช้ change_rate แทน"""
    ns["_fake_client"].response_text = '{"matched": true, "label": "ภาษีมูลค่าเพิ่ม 7%", "new_value": "10%"}'
    result = ns["_match_and_apply_excel_edit"](1, 1, "เปลี่ยน VAT เป็น 10%")
    assert result["matched"] is False
    assert "change_rate" in result["message"] or "เปลี่ยนอัตรา" in result["message"]

    # ไม่มีเวอร์ชันใหม่ถูกสร้างเลย (ไม่ได้แก้จริง) สูตรเดิมยังอยู่ในสำเนาต้นฉบับ
    assert ns["_fake_db"].versions.get(1, []) == []
    vat_info = doc1["label_map"]["ภาษีมูลค่าเพิ่ม 7%"]
    wb_raw = openpyxl.load_workbook(io.BytesIO(doc1["original_bytes"]), data_only=False)
    cell_value = wb_raw[vat_info["sheet"]].cell(row=vat_info["row"], column=vat_info["col"]).value
    assert isinstance(cell_value, str) and cell_value.startswith("=")


def test_match_and_apply_excel_edit_still_works_for_plain_values(ns, doc1):
    """ยืนยันว่า safety net ไม่กระทบการแก้ค่าปกติ (ไม่ใช่สูตร หรือไม่มี "%" ในค่าใหม่)"""
    ns["_fake_client"].response_text = '{"matched": true, "label": "รวมเงิน", "new_value": "99999"}'
    result = ns["_match_and_apply_excel_edit"](1, 1, "เปลี่ยนรวมเงินเป็น 99999")
    assert result["matched"] is True


def test_document_reading_question_fallback_when_client_fails(ns, doc1):
    """client.messages.create ล้ม -> ตอบ _format_document_values_text(data) ตรงๆ แทน ไม่ raise ออกไป"""
    ns["_fake_client"].raise_on_create = RuntimeError("simulated API failure")
    active_docs = [doc1]
    answer, used_doc_id = ns["_answer_document_reading_question"]("มีรายการอะไรบ้าง", active_docs, 1, 1)
    assert used_doc_id == 1

    expected = ns["_format_document_values_text"](ns["read_document_values"](doc1))
    assert answer == expected
