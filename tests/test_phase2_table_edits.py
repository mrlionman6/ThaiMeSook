"""ทดสอบเฟส 2 PR2 (แก้โครงสร้าง Excel: เปลี่ยนอัตรา / แก้-ลบ-เพิ่มรายการในตาราง) + เฟส 3 (คิว pending action
หลายรายการต่อแชท + กัน AI แต่งผลการแก้เอง) ใช้ fixture workbook สังเคราะห์ (ไม่ใช่ไฟล์ลูกค้าจริง) และ fake DB/
LLM client ในหน่วยความจำล้วนๆ (ไม่ต่อ DB จริงเลย) ผ่าน tests/_harness.py"""
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


def _confirm_single(ns, document_id=1, user_id=1, chat_id=1):
    """ช่วย apply รายการล่าสุดเดี่ยวๆ ในคิว (คิวต้องมีพอดี 1 รายการ) — เทียบเท่าผู้ใช้พิมพ์ 'ยืนยัน'"""
    queue = ns["get_chat_pending_actions"](chat_id)
    assert len(queue) == 1, f"คาดว่าคิวมี 1 รายการ แต่มี {len(queue)}"
    return ns["_handle_pending_queue_command"]("confirm_single", queue, user_id, chat_id)


# ---------- change_rate ----------

def test_change_rate_to_10_percent(ns, doc1):
    result = ns["propose_change_rate"](1, 1, 1, "ภาษีมูลค่าเพิ่ม 7%", 10)
    assert result["proposed"] is True
    assert "10%" in result["message"]
    assert result["queue_count"] == 1

    confirmed = _confirm_single(ns)
    assert "เปลี่ยนอัตราของ" in confirmed
    expected_vat = round(FIXTURE_TOTAL * 0.10, 2)
    expected_total = FIXTURE_TOTAL + expected_vat
    assert str(expected_vat) in confirmed
    assert str(expected_total) in confirmed


def test_change_rate_unsupported_formula_is_rejected(ns, doc1):
    # แก้สูตรของ "รวมเงิน" (=SUM(G16:G25)) ให้ไม่ตรงรูปแบบอัตราที่รองรับเลย
    result = ns["propose_change_rate"](1, 1, 1, "รวมเงิน", 10)
    assert result["proposed"] is False
    assert "ไม่ใช่รูปแบบอัตราที่รองรับ" in result["message"]
    assert ns["get_chat_pending_actions"](1) == []


def test_change_rate_position_based_lookup_matches_real_file_scale(ns):
    """ทวนด้วยตัวเลขจากไฟล์จริง (ป่าไม้เขียว, PR #25): รวมเงิน 58150, VAT 7%->10% ต้องได้ 5815 และรวม
    ราคาทั้งสิ้น 58150+5815=63965 — ยืนยันว่าหา "ค่าใหม่" ถูกต้องด้วยตำแหน่งเซลล์ แม้ label มี % อยู่ในชื่อ"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Quotation"
    ws["A26"] = "รวมเงิน"
    ws["G26"] = 58150
    ws["A27"] = "ภาษีมูลค่าเพิ่ม 7%"
    ws["G27"] = "=G26*7/100"
    ws["A28"] = "รวมราคาทั้งสิ้น"
    ws["G28"] = "=SUM(G26:G27)"
    buf = io.BytesIO()
    wb.save(buf)
    raw = inject_formula_cache(buf.getvalue(), {"G27": 4070.5, "G28": 62220.5})
    label_map = ns["_extract_excel_labels"](raw)
    ns["_fake_db"].add_document(2, "real_scale.xlsx", raw, label_map, chat_id=1, user_id=1)

    result = ns["propose_change_rate"](2, 1, 1, "ภาษีมูลค่าเพิ่ม 7%", 10)
    assert result["proposed"] is True
    confirmed = _confirm_single(ns, document_id=2)
    assert "5815" in confirmed
    assert "63965" in confirmed


# ---------- update_table_item (เปลี่ยนมาต้องยืนยันด้วยแล้ว — ส่วน 1 ของ task) ----------

def test_update_table_item_description(ns, doc1):
    propose = ns["propose_update_table_item"](1, 1, 1, 3, {"description": "Item C renamed"})
    assert propose["proposed"] is True
    confirmed = _confirm_single(ns)
    assert "Item C renamed" in confirmed


# ---------- delete_table_item ----------

def test_delete_middle_item_shifts_and_renumbers_continuously(ns, doc1):
    """ลบรายการที่ 3 จาก 8 รายการ: ลำดับต้องเป็น 1..7 ต่อเนื่อง และรายละเอียดต้องเลื่อนขึ้นถูกแถว (รายการที่ 3
    ใหม่ต้องเป็นของเดิมรายการที่ 4) — Box 6x6 (รายการที่ 2) ต้องยังอยู่ ไม่ถูกลบผิดแถว"""
    propose = ns["propose_delete_table_item"](1, 1, 1, 3)
    assert propose["proposed"] is True
    _confirm_single(ns)

    updated_doc = ns["get_editable_document_raw"](1)
    data = ns["read_document_values"](updated_doc)
    table = data["tables"][0]
    filled = table["rows"]  # read_document_values คืนเฉพาะแถวที่ไม่ว่างอยู่แล้ว

    assert len(filled) == 7
    assert [r["values"]["no"] for r in filled] == [1, 2, 3, 4, 5, 6, 7]
    assert filled[1]["values"]["description"] == "Item B"  # Box 6x6 เทียบเท่า (รายการที่ 2 เดิม) ยังอยู่
    assert filled[2]["values"]["description"] == "Item D"
    assert filled[2]["values"]["qty"] == 2
    assert filled[2]["values"]["unit_price"] == 300


def test_undo_after_delete_restores_original(ns, doc1):
    ns["propose_delete_table_item"](1, 1, 1, 3)
    _confirm_single(ns)

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
    _confirm_single(ns)

    result2 = ns["propose_add_table_item"](1, 1, 1, {"description": "Item J", "qty": 1, "unit": "ชิ้น", "unit_price": 5})
    assert result2["proposed"] is True
    _confirm_single(ns)

    # ตอนนี้เต็มแล้ว (10/10) — ลองเพิ่มอีกต้องถูกปฏิเสธทันที (ไม่ต่อคิวเลย)
    rejected = ns["propose_add_table_item"](1, 1, 1, {"description": "Item K", "qty": 1, "unit": "ชิ้น", "unit_price": 1})
    assert rejected["proposed"] is False
    assert "เต็มแล้ว" in rejected["message"]
    assert "10" in rejected["message"]
    assert ns["get_chat_pending_actions"](1) == []


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


# ---------- คิว pending action: คำสั่งพื้นฐาน ----------

@pytest.mark.parametrize("text", ["ยืนยันทั้งหมด", "ยืนยันหมด", "ยืนยันทุกรายการ"])
def test_classify_confirm_all(ns, text):
    assert ns["_classify_pending_queue_command"](text) == "confirm_all"


def test_classify_confirm_latest(ns):
    assert ns["_classify_pending_queue_command"]("ยืนยันล่าสุด") == "confirm_latest"


@pytest.mark.parametrize("text", ["ยืนยัน", "ยืนยันครับ", "ตกลง ยืนยัน"])
def test_classify_confirm_single(ns, text):
    assert ns["_classify_pending_queue_command"](text) == "confirm_single"


@pytest.mark.parametrize("text", ["ไม่ยืนยัน", "ยังไม่ยืนยัน"])
def test_classify_confirm_rejects_negated_forms(ns, text):
    assert ns["_classify_pending_queue_command"](text) is None


def test_classify_cancel_latest(ns):
    assert ns["_classify_pending_queue_command"]("ยกเลิกล่าสุด") == "cancel_latest"


@pytest.mark.parametrize("text", ["ยกเลิก", "ยกเลิกทั้งหมด"])
def test_classify_cancel_all(ns, text):
    assert ns["_classify_pending_queue_command"](text) == "cancel_all"


def test_classify_cancel_rejects_negated_form(ns):
    assert ns["_classify_pending_queue_command"]("ไม่ยกเลิก") is None


# ---------- คิว pending action: คำสั่งทนรูปแบบ (ตัดคำลงท้าย/สร้อยคำพูดก่อนเทียบ) ----------

@pytest.mark.parametrize("text,expected", [
    ("ยืนยันทั้งหมดครับ", "confirm_all"),
    ("ยืนยันทั้งหมดเลย", "confirm_all"),
    ("ยืนยันหมดเลยนะ", "confirm_all"),
    ("ยกเลิกทั้งหมดเลย", "cancel_all"),
    ("ยืนยันด้วย", "confirm_single"),
    ("ยืนยันล่าสุดครับ", "confirm_latest"),
])
def test_classify_tolerates_trailing_particles(ns, text, expected):
    """เหตุการณ์จริงที่แก้: "ยืนยันทั้งหมดเลย" ไม่ตรง _CONFIRM_ALL_TEXTS เป๊ะ (ขาดแค่คำว่า "เลย") เลยตกไปให้ LLM
    ตีความแทน ซึ่งไปเรียก propose ซ้ำของเดิมในคิว — ตอนนี้ต้องตัดคำลงท้าย/สร้อยคำพูดออกก่อนเทียบเสมอ"""
    assert ns["_classify_pending_queue_command"](text) == expected


@pytest.mark.parametrize("text", ["ไม่ยืนยันนะ", "อย่าเพิ่งยืนยัน", "ยังไม่ยืนยันครับ"])
def test_classify_still_rejects_negated_forms_with_particles(ns, text):
    """คำปฏิเสธต้องไม่ถูกตัดคำลงท้ายจนกลายเป็นคำสั่งยืนยันจริงๆ ไปได้ — "ไม่ยืนยันนะ" ตัด "นะ" เหลือ "ไม่ยืนยัน"
    ซึ่งยังไม่ตรง _CONFIRM_SINGLE_TEXTS อยู่ดี (ไม่ใช่ "ยืนยัน" เฉยๆ)"""
    assert ns["_classify_pending_queue_command"](text) is None


@pytest.mark.parametrize("text,expected", [
    ("ยืนยันทั้งหมดเลย", "ยืนยันทั้งหมด"),
    ("ยืนยันหมดเลยนะ", "ยืนยันหมด"),
    ("ไม่ยืนยันนะ", "ไม่ยืนยัน"),
    ("ยืนยันได้เลย", "ยืนยัน"),
])
def test_normalize_pending_queue_command_text(ns, text, expected):
    assert ns["_normalize_pending_queue_command_text"](text) == expected


# ---------- คิว pending action: ดักคำสั่งที่ "ดูเหมือน" ตั้งใจจะสั่งคิวแต่พิมพ์แปลกไปจากรูปแบบที่รู้จัก ----------

def test_mentions_pending_queue_intent_catches_unrecognized_confirm_phrasing(ns):
    assert ns["_mentions_pending_queue_intent"]("ยืนยันแบบนั้นแหละ") is True


@pytest.mark.parametrize("text", ["ไม่ยืนยันนะ", "อย่าเพิ่งยืนยัน", "ไม่ยกเลิกนะ"])
def test_mentions_pending_queue_intent_excludes_negated_forms(ns, text):
    assert ns["_mentions_pending_queue_intent"](text) is False


@pytest.mark.parametrize("text", [
    "ต้องยืนยันตัวตนก่อนยื่นภาษีไหม",
    "ยืนยันตัวตนต้องทำอย่างไรบ้าง",
])
def test_mentions_pending_queue_intent_excludes_real_questions(ns, text):
    """กันคำถามกฎหมายทั่วไปที่บังเอิญมีคำว่า "ยืนยัน" ปนอยู่ (เช่น "ต้องยืนยันตัวตนก่อนยื่นภาษีไหม") ถูกเข้าใจผิด
    เป็นความตั้งใจจะสั่งคิว — ใช้คำเชิงคำถามเดียวกับ _OPEN_COMMAND_QUESTION_WORDS กัน"""
    assert ns["_mentions_pending_queue_intent"](text) is False


def test_mentions_pending_queue_intent_false_when_no_confirm_cancel_word(ns):
    assert ns["_mentions_pending_queue_intent"]("เปลี่ยน Box 4x4 เป็น 7x7") is False


# ---------- คิว pending action: กันรายการซ้ำ/การแก้ที่ไม่เปลี่ยนอะไร (propose_* ทุกตัว) ----------

def test_propose_change_rate_duplicate_in_queue_is_rejected(ns, doc1):
    first = ns["propose_change_rate"](1, 1, 1, "ภาษีมูลค่าเพิ่ม 7%", 10)
    assert first["proposed"] is True

    second = ns["propose_change_rate"](1, 1, 1, "ภาษีมูลค่าเพิ่ม 7%", 10)
    assert second["proposed"] is False
    assert "อยู่ในคิวรอยืนยันแล้วครับ (ข้อ 1)" in second["message"]
    assert len(ns["get_chat_pending_actions"](1)) == 1


def test_propose_change_rate_noop_when_already_at_target_rate(ns, doc1):
    """ไฟล์ fixture เริ่มต้นที่ VAT 7% อยู่แล้ว — เสนอเปลี่ยนเป็น 7% (เท่าเดิม) ต้องไม่ต่อคิว"""
    result = ns["propose_change_rate"](1, 1, 1, "ภาษีมูลค่าเพิ่ม 7%", 7)
    assert result["proposed"] is False
    assert "อยู่แล้ว" in result["message"]
    assert ns["get_chat_pending_actions"](1) == []


def test_propose_update_table_item_duplicate_in_queue_is_rejected(ns, doc1):
    """เปลี่ยนคอลัมน์อื่นที่ไม่ใช่ "รายละเอียด" (เช่น จำนวน) ซ้ำ — รายละเอียดเดิมยังเหมือนเดิมในสถานะเสมือนด้วย
    (ไม่ใช่ฟิลด์ที่เพิ่งแก้) identity จึงยังตรงกับรายการที่คิวไว้แล้ว ตรวจจับเป็นรายการซ้ำได้ถูกต้อง (ต่างจากกรณี
    แก้ "รายละเอียด" เองซ้ำ ซึ่งจะเข้าทาง no-op แทน เพราะสถานะเสมือนหลังคิวแรกมีค่าใหม่อยู่แล้ว — ดู
    test_propose_update_table_item_noop_when_value_matches_already_queued_change ด้านล่าง)"""
    first = ns["propose_update_table_item"](1, 1, 1, 1, {"qty": 99})
    assert first["proposed"] is True

    second = ns["propose_update_table_item"](1, 1, 1, 1, {"qty": 99})
    assert second["proposed"] is False
    assert "อยู่ในคิวรอยืนยันแล้วครับ (ข้อ 1)" in second["message"]
    assert len(ns["get_chat_pending_actions"](1)) == 1


def test_propose_update_table_item_noop_when_value_unchanged(ns, doc1):
    result = ns["propose_update_table_item"](1, 1, 1, 1, {"description": "Item A"})
    assert result["proposed"] is False
    assert "อยู่แล้ว" in result["message"]
    assert ns["get_chat_pending_actions"](1) == []


def test_propose_update_table_item_noop_when_value_matches_already_queued_change(ns, doc1):
    """เสนอแก้ "รายละเอียด" เป็นค่าเดียวกันซ้ำ (ฟิลด์เดียวกับที่ใช้จับคู่ identity เอง) — สถานะเสมือนหลังรายการ
    แรกในคิวมีค่าใหม่อยู่แล้ว การเสนอซ้ำจึงเข้าทาง "no-op" (ค่าที่ต้องการอยู่แล้วเมื่อนับรวมคิว) แทนที่จะเป็น
    "รายการซ้ำ" ตรงๆ — แต่ผลลัพธ์ที่ผู้ใช้ต้องการ (ไม่ต่อคิวซ้ำ) เหมือนกัน"""
    first = ns["propose_update_table_item"](1, 1, 1, 1, {"description": "Box 7x7"})
    assert first["proposed"] is True

    second = ns["propose_update_table_item"](1, 1, 1, 1, {"description": "Box 7x7"})
    assert second["proposed"] is False
    assert "อยู่แล้ว" in second["message"]
    assert len(ns["get_chat_pending_actions"](1)) == 1


def test_find_duplicate_queue_action_detects_delete_table_item_duplicate(ns):
    """delete_table_item ไม่สามารถเสนอซ้ำผ่าน propose_delete_table_item ตามธรรมชาติได้จริง (รายการที่ถูกลบไปแล้ว
    ในคิวจะหายไปจากสถานะเสมือนเลย ทำให้ item_no เดิมชี้ไปรายการอื่นที่เลื่อนมาแทนเสมอ ไม่ใช่รายการซ้ำ) — ทดสอบ
    ตรรกะการตรวจจับรายการซ้ำของ _find_duplicate_queue_action ตรงๆ แทน (จำลองคิวที่มี delete ของ "Item C" อยู่
    แล้ว แล้วเสนอ delete ของ "Item C" เดิมซ้ำด้วย identity เดียวกันเป๊ะ)"""
    pending_queue = [{
        "id": 1,
        "action": {"type": "delete_table_item", "document_id": 1, "identity": {"row": 18, "description": "Item C"}},
    }]
    candidate = {"type": "delete_table_item", "document_id": 1, "identity": {"row": 18, "description": "Item C"}}
    assert ns["_find_duplicate_queue_action"](pending_queue, candidate) == 1

    different = {"type": "delete_table_item", "document_id": 1, "identity": {"row": 19, "description": "Item D"}}
    assert ns["_find_duplicate_queue_action"](pending_queue, different) is None


def test_propose_add_table_item_duplicate_in_queue_is_rejected(ns, doc1):
    fields = {"description": "Item I", "qty": 2, "unit": "ชิ้น", "unit_price": 20}
    first = ns["propose_add_table_item"](1, 1, 1, dict(fields))
    assert first["proposed"] is True

    second = ns["propose_add_table_item"](1, 1, 1, dict(fields))
    assert second["proposed"] is False
    assert "อยู่ในคิวรอยืนยันแล้วครับ (ข้อ 1)" in second["message"]
    assert len(ns["get_chat_pending_actions"](1)) == 1


# ---------- คิว pending action: ชื่อฟิลด์ภาษาไทยใน preview/receipt ----------

def test_update_table_item_preview_uses_thai_field_names(ns, doc1):
    result = ns["propose_update_table_item"](1, 1, 1, 1, {"description": "Box 7x7", "qty": 15})
    assert result["proposed"] is True
    assert "รายละเอียด จะเปลี่ยนจาก" in result["message"]
    assert "จำนวน จะเปลี่ยนจาก" in result["message"]
    assert "description" not in result["message"]
    assert "qty" not in result["message"]


def test_add_table_item_preview_uses_thai_field_names(ns, doc1):
    result = ns["propose_add_table_item"](1, 1, 1, {"description": "Item I", "qty": 2, "unit": "ชิ้น", "unit_price": 20})
    assert result["proposed"] is True
    assert "รายละเอียด=Item I" in result["message"]
    assert "จำนวน=2" in result["message"]
    assert "ราคา/หน่วย=20" in result["message"]


def test_apply_update_table_item_receipt_uses_thai_field_names(ns, doc1):
    ns["propose_update_table_item"](1, 1, 1, 1, {"description": "Box 7x7"})
    queue = ns["get_chat_pending_actions"](1)
    confirmed = ns["_handle_pending_queue_command"]("confirm_single", queue, 1, 1)
    assert "รายละเอียด เปลี่ยนเป็น Box 7x7" in confirmed


# ---------- คำสั่งรูปแบบชัดเจน: parse ด้วยโค้ดตรงๆ ไม่ผ่าน LLM (_parse_direct_structural_command) ----------

def test_find_single_rate_label_finds_the_one_percent_label(ns, doc1):
    base_bytes, label_map = ns["get_working_copy"](doc1)
    assert ns["_find_single_rate_label"](base_bytes, label_map) == "ภาษีมูลค่าเพิ่ม 7%"


@pytest.mark.parametrize("query", ["เปลี่ยน VAT เป็น 10%", "แก้ภาษีเป็น10%", "เปลี่ยนภาษีมูลค่าเพิ่มเป็น 10%"])
def test_parse_direct_structural_command_detects_rate_change(ns, doc1, query):
    base_bytes, label_map = ns["get_working_copy"](doc1)
    action = ns["_parse_direct_structural_command"](query, base_bytes, label_map)
    assert action == {"type": "change_rate", "label": "ภาษีมูลค่าเพิ่ม 7%", "new_rate": 10.0}


def test_parse_direct_structural_command_detects_delete(ns, doc1):
    base_bytes, label_map = ns["get_working_copy"](doc1)
    action = ns["_parse_direct_structural_command"]("ลบรายการที่ 3", base_bytes, label_map)
    assert action == {"type": "delete_table_item", "item_no": 3}


def test_parse_direct_structural_command_returns_none_for_ambiguous_command(ns, doc1):
    base_bytes, label_map = ns["get_working_copy"](doc1)
    assert ns["_parse_direct_structural_command"]("เปลี่ยน Box 4x4 เป็น 7x7", base_bytes, label_map) is None


def test_parse_direct_structural_command_returns_none_when_multiple_rate_labels(ns):
    """ไฟล์มี label ที่มี "%" และเป็นสูตรอัตราที่รู้จักมากกว่า 1 ตัว — กำกวมเกินจะเดา ต้องคืน None (ปล่อยให้ agent/
    LLM ตัดสินใจแทน ไม่ใช่เดาสุ่มตัวใดตัวหนึ่ง)"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Quotation"
    ws["A1"] = "ภาษี A 7%"
    ws["B1"] = "=C1*7/100"
    ws["A2"] = "ภาษี B 7%"
    ws["B2"] = "=C2*7/100"
    buf = io.BytesIO()
    wb.save(buf)
    raw = inject_formula_cache(buf.getvalue(), {"B1": 70, "B2": 70})
    label_map = ns["_extract_excel_labels"](raw)

    action = ns["_parse_direct_structural_command"]("เปลี่ยน VAT เป็น 10%", raw, label_map)
    assert action is None


# ---------- สถานการณ์จริง (รอบ 2): เปลี่ยน VAT -> เปลี่ยนชื่อ -> "ลบรายการที่ 3" (parse ตรง ไม่ใช้ LLM) ->
# "ยืนยัน" (ถามกลับ) -> "ยืนยันทั้งหมดเลย" (normalize แล้วจึง apply ครบ) ----------

def test_full_scenario_direct_parse_delete_then_confirm_all_with_particle(ns, doc1):
    r1 = ns["propose_change_rate"](1, 1, 1, "ภาษีมูลค่าเพิ่ม 7%", 10)
    assert r1["proposed"] is True

    r2 = ns["propose_update_table_item"](1, 1, 1, 1, {"description": "Box 7x7"})
    assert r2["proposed"] is True

    # "ลบรายการที่ 3" ต้อง parse ได้ด้วยโค้ดตรงๆ (เหมือนที่ ask_question() เรียกจริง) ไม่ต้องพึ่ง LLM เลือก tool
    base_bytes, label_map = ns["get_working_copy"](ns["get_editable_document_raw"](1))
    direct_action = ns["_parse_direct_structural_command"]("ลบรายการที่ 3", base_bytes, label_map)
    assert direct_action == {"type": "delete_table_item", "item_no": 3}
    r3 = ns["propose_delete_table_item"](1, 1, 1, direct_action["item_no"])
    assert r3["proposed"] is True

    queue = ns["get_chat_pending_actions"](1)
    assert len(queue) == 3

    # "ยืนยัน" เฉยๆ กับคิว 3 รายการ -> ถามกลับ ไม่ apply (ลำดับเดิมจาก PR #28)
    ask_back = ns["_handle_pending_queue_command"]("confirm_single", queue, 1, 1)
    assert "ยืนยันทั้งหมด" in ask_back
    assert ns["get_chat_pending_actions"](1) == queue

    # "ยืนยันทั้งหมดเลย" ต้อง normalize แล้วจัดเป็น confirm_all เหมือน "ยืนยันทั้งหมด" เป๊ะ (ไม่ตกไปให้ LLM)
    command = ns["_classify_pending_queue_command"]("ยืนยันทั้งหมดเลย")
    assert command == "confirm_all"
    ns["_handle_pending_queue_command"](command, queue, 1, 1)
    assert ns["get_chat_pending_actions"](1) == []

    updated_doc = ns["get_editable_document_raw"](1)
    data = ns["read_document_values"](updated_doc)
    filled = data["tables"][0]["rows"]
    assert len(filled) == 7
    assert filled[0]["values"]["description"] == "Box 7x7"
    assert not any(r["values"]["description"] == "Item C" for r in filled)
    label_values = {l["label"]: l["value"] for l in data["labels"]}
    assert label_values["ภาษีมูลค่าเพิ่ม 7%"] != FIXTURE_VAT  # VAT เปลี่ยนไปจากเดิมแล้ว (เป็น 10% ไม่ใช่ 7%)


def test_is_pending_queue_fresh_handles_naive_and_aware_and_empty(ns):
    naive_now = datetime.datetime.utcnow()
    aware_now = datetime.datetime.now(datetime.timezone.utc)
    stale = naive_now - datetime.timedelta(seconds=ns["PENDING_ACTION_TTL_SECONDS"] + 5)

    assert ns["_is_pending_queue_fresh"]([]) is False
    assert ns["_is_pending_queue_fresh"]([{"created_at": naive_now}]) is True
    assert ns["_is_pending_queue_fresh"]([{"created_at": aware_now}]) is True
    assert ns["_is_pending_queue_fresh"]([{"created_at": stale}]) is False


# ---------- สถานการณ์จริงเต็มรูปแบบ: VAT -> เปลี่ยนชื่อสินค้า -> ลบรายการ -> ยืนยันทั้งหมด -> undo ----------

def test_full_scenario_vat_rename_delete_confirm_all_then_undo(ns, doc1):
    # 1) เปลี่ยน VAT เป็น 10%
    r1 = ns["propose_change_rate"](1, 1, 1, "ภาษีมูลค่าเพิ่ม 7%", 10)
    assert r1["proposed"] is True
    assert r1["queue_count"] == 1

    # 2) เปลี่ยน Box 4x4 (รายการที่ 1) เป็น Box 7x7
    r2 = ns["propose_update_table_item"](1, 1, 1, 1, {"description": "Box 7x7"})
    assert r2["proposed"] is True
    assert r2["queue_count"] == 2
    # preview ของรายการนี้ต้องคำนวณต่อจาก VAT 10% แล้ว (ยอดในข้อความ side-effect ควรสะท้อนอัตราใหม่ ถ้ามี
    # side effect ปรากฏ) — อย่างน้อยต้องไม่ error และต่อคิวสำเร็จ

    # 3) ลบรายการที่ 3 (ไส้ไก่ — เทียบเท่า Item C ในฟิกซ์เจอร์นี้)
    r3 = ns["propose_delete_table_item"](1, 1, 1, 3)
    assert r3["proposed"] is True
    assert r3["queue_count"] == 3
    assert "Item C" in r3["message"]

    # คิวมี 3 รายการ — พิมพ์ "ยืนยัน" เฉยๆ ต้องไม่ apply อะไรเลย แค่แสดงรายการแล้วถาม
    queue = ns["get_chat_pending_actions"](1)
    assert len(queue) == 3
    ask_back = ns["_handle_pending_queue_command"]("confirm_single", queue, 1, 1)
    assert "ยืนยันทั้งหมด" in ask_back and "ยืนยันล่าสุด" in ask_back
    assert ns["get_chat_pending_actions"](1) == queue  # ยังไม่ถูกลบออกจากคิวเลย (ไม่ได้ apply อะไร)

    # "ยืนยันทั้งหมด" -> apply ทุกรายการตามลำดับ
    result_text = ns["_handle_pending_queue_command"]("confirm_all", queue, 1, 1)
    assert ns["get_chat_pending_actions"](1) == []

    updated_doc = ns["get_editable_document_raw"](1)
    data = ns["read_document_values"](updated_doc)
    table = data["tables"][0]
    filled = table["rows"]

    assert len(filled) == 7
    assert filled[0]["values"]["description"] == "Box 7x7"  # รายการที่ 1 เปลี่ยนชื่อแล้ว
    assert [r["values"]["no"] for r in filled] == [1, 2, 3, 4, 5, 6, 7]  # ลำดับต่อเนื่อง ไม่ใช่ 1,2,1,2,...
    assert filled[1]["values"]["description"] == "Item B"  # Box 6x6 เทียบเท่า ยังอยู่ (ไม่ถูกลบผิดแถว)
    assert not any(r["values"]["description"] == "Item C" for r in filled)  # ไส้ไก่ (Item C) หายไปแล้ว

    label_values = {l["label"]: l["value"] for l in data["labels"]}
    assert label_values["ภาษีมูลค่าเพิ่ม 7%"] != FIXTURE_VAT  # VAT เปลี่ยนไปจากเดิมแล้ว (เป็น 10% ไม่ใช่ 7%)

    # undo 1 ครั้ง -> ย้อนเฉพาะรายการล่าสุด (ลบ) รายการที่ 3 (ไส้ไก่/Item C) ต้องกลับมา แต่ Box 7x7 และ VAT 10%
    # ยังอยู่ (คนละเวอร์ชันกัน — undo ทีละขั้นได้จริง)
    undo_result = ns["undo_last_edit"](1, 1)
    assert undo_result["matched"] is True

    after_undo_doc = ns["get_editable_document_raw"](1)
    after_undo_data = ns["read_document_values"](after_undo_doc)
    after_undo_table = after_undo_data["tables"][0]
    after_undo_filled = after_undo_table["rows"]
    assert len(after_undo_filled) == 8
    assert any(r["values"]["description"] == "Item C" for r in after_undo_filled)
    assert after_undo_filled[0]["values"]["description"] == "Box 7x7"  # ยังเป็นชื่อใหม่อยู่ (คนละขั้นกับลบ)


def test_queue_preview_computes_on_top_of_prior_queue_items(ns, doc1):
    """preview ของรายการที่ 2 ในคิว (ลบรายการที่ 3) ต้องคำนวณยอดต่อจาก VAT 10% ที่เสนอไปก่อนหน้าแล้ว ไม่ใช่ยอด
    ดั้งเดิม 7%"""
    ns["propose_change_rate"](1, 1, 1, "ภาษีมูลค่าเพิ่ม 7%", 10)
    r2 = ns["propose_delete_table_item"](1, 1, 1, 3)
    assert r2["proposed"] is True
    # ถ้า preview คำนวณถูกต้อง (ต่อจาก VAT 10%) side effect ของรวมราคาทั้งสิ้นที่โชว์ต้องไม่ใช่เลขที่มาจาก VAT 7%
    assert str(FIXTURE_GRAND_TOTAL) not in r2["message"]


def test_confirm_latest_applies_only_last_item(ns, doc1):
    ns["propose_change_rate"](1, 1, 1, "ภาษีมูลค่าเพิ่ม 7%", 10)
    ns["propose_delete_table_item"](1, 1, 1, 3)
    queue = ns["get_chat_pending_actions"](1)
    assert len(queue) == 2

    ns["_handle_pending_queue_command"]("confirm_latest", queue, 1, 1)

    remaining = ns["get_chat_pending_actions"](1)
    assert len(remaining) == 1
    assert remaining[0]["action"]["type"] == "change_rate"  # รายการแรก (VAT) ยังอยู่ในคิว ไม่ได้ apply

    updated_doc = ns["get_editable_document_raw"](1)
    data = ns["read_document_values"](updated_doc)
    filled = data["tables"][0]["rows"]
    assert len(filled) == 7  # รายการลบ (ล่าสุด) ถูก apply ไปแล้วจริง


def test_confirm_latest_then_confirm_single_applies_remaining_item(ns, doc1):
    """บั๊กจริงจาก review PR #28: propose VAT 10% -> propose ลบรายการที่ 3 -> "ยืนยันล่าสุด" (ลบ) ->
    "ยืนยัน" (คิวเหลือรายการเดียว คือ VAT) ต้อง apply รายการ VAT สำเร็จจริง ไม่ใช่โดนเข้าใจผิดว่ามีการแก้นอกคิว
    แทรกเข้ามา (เพราะ apply รายการแรกในคิวทำให้เวอร์ชันเอกสารขยับไปจริง ต้องอัปเดต version_id_at_propose ของ
    รายการที่เหลือในคิวตามไปด้วย ไม่งั้น baseline check รอบถัดไปจะเทียบกับเวอร์ชันเก่าที่ไม่ตรงแล้วเสมอ)"""
    r1 = ns["propose_change_rate"](1, 1, 1, "ภาษีมูลค่าเพิ่ม 7%", 10)
    assert r1["proposed"] is True
    r2 = ns["propose_delete_table_item"](1, 1, 1, 3)
    assert r2["proposed"] is True

    queue = ns["get_chat_pending_actions"](1)
    latest_result = ns["_handle_pending_queue_command"]("confirm_latest", queue, 1, 1)
    assert latest_result.startswith("✓")
    assert "เปลี่ยนไป" not in latest_result

    remaining = ns["get_chat_pending_actions"](1)
    assert len(remaining) == 1
    assert remaining[0]["action"]["type"] == "change_rate"

    final_result = ns["_handle_pending_queue_command"]("confirm_single", remaining, 1, 1)
    assert "เปลี่ยนไป" not in final_result  # ไม่ใช่ถูกเข้าใจผิดว่ามีการแก้นอกคิวแทรก
    assert final_result.startswith("✓ เปลี่ยนอัตรา")
    assert ns["get_chat_pending_actions"](1) == []

    updated_doc = ns["get_editable_document_raw"](1)
    data = ns["read_document_values"](updated_doc)
    filled = data["tables"][0]["rows"]
    assert len(filled) == 7
    assert not any(r["values"]["description"] == "Item C" for r in filled)

    label_values = {l["label"]: l["value"] for l in data["labels"]}
    expected_total = FIXTURE_TOTAL - 150  # Item C ที่ถูกลบ: qty 3 * unit_price 50
    expected_vat = round(expected_total * 0.10, 2)
    expected_grand_total = expected_total + expected_vat
    assert label_values["รวมเงิน"] == expected_total
    assert label_values["ภาษีมูลค่าเพิ่ม 7%"] == expected_vat  # key ชื่อเดิมแช่แข็งไว้ แต่ค่าเป็น 10% แล้ว
    assert label_values["รวมราคาทั้งสิ้น"] == expected_grand_total


def test_confirm_latest_then_external_edit_still_clears_remaining_queue(ns, doc1):
    """การอัปเดต version_id_at_propose ของรายการที่เหลือหลัง apply บางส่วนสำเร็จ (บั๊กที่เพิ่งแก้ข้างบน) ต้องไม่
    ทำให้การตรวจจับ "มีการแก้จากนอกคิวแทรก" เสียไป — ยืนยันล่าสุดสำเร็จ แล้วมีการแก้จากนอกคิวแทรกเข้ามาอีกที (ไม่
    ผ่านกลไกคิวเลย) ก่อนยืนยันรายการที่เหลือ ต้องยังล้างคิวทั้งหมดทิ้งและแจ้งผู้ใช้เหมือนเดิม ไม่ใช่ apply ทับไป
    เงียบๆ"""
    ns["propose_change_rate"](1, 1, 1, "ภาษีมูลค่าเพิ่ม 7%", 10)
    ns["propose_delete_table_item"](1, 1, 1, 3)
    queue = ns["get_chat_pending_actions"](1)
    ns["_handle_pending_queue_command"]("confirm_latest", queue, 1, 1)

    doc = ns["get_editable_document_raw"](1)
    base_bytes, label_map = ns["get_working_copy"](doc)

    def _external_mutate(wb):
        ws = wb["Quotation"]
        ws.cell(row=16, column=2, value="Changed from outside the queue again")

    ns["create_new_version"](1, _external_mutate, change_summary="external edit 2")

    remaining = ns["get_chat_pending_actions"](1)
    assert len(remaining) == 1
    answer = ns["_handle_pending_queue_command"]("confirm_single", remaining, 1, 1)
    assert "เปลี่ยนไป" in answer
    assert ns["get_chat_pending_actions"](1) == []


def test_identity_based_resolution_survives_prior_delete_in_queue(ns, doc1):
    """เสนอแก้รายการที่ 5 (Item E) ก่อน แล้วเสนอลบรายการที่ 1 (Item A) ทีหลัง — ตอน apply ทั้งคิว การแก้ Item E
    ต้องยัง "เจอ Item E ถูกตัว" แม้ว่ารายการก่อนหน้ามันจะถูกลบไปจริงแล้ว (เลื่อนตำแหน่งแถว) ไม่ใช่ไปแก้แถวอื่นผิด
    (เพราะ re-resolve ด้วย identity/description ไม่ใช่เลขแถวเดิม)"""
    r1 = ns["propose_update_table_item"](1, 1, 1, 5, {"description": "Item E renamed"})
    assert r1["proposed"] is True
    r2 = ns["propose_delete_table_item"](1, 1, 1, 1)
    assert r2["proposed"] is True
    assert "Item A" in r2["message"]

    queue = ns["get_chat_pending_actions"](1)
    ns["_handle_pending_queue_command"]("confirm_all", queue, 1, 1)

    updated_doc = ns["get_editable_document_raw"](1)
    data = ns["read_document_values"](updated_doc)
    filled = data["tables"][0]["rows"]
    assert len(filled) == 7
    assert any(r["values"]["description"] == "Item E renamed" for r in filled)
    assert not any(r["values"]["description"] == "Item A" for r in filled)
    assert not any(r["values"]["description"] == "Item E" for r in filled)  # ไม่มีชื่อเดิมค้างอยู่ (ไม่ได้แก้แถวอื่นผิด)


def test_cancel_latest_removes_only_last_item(ns, doc1):
    ns["propose_change_rate"](1, 1, 1, "ภาษีมูลค่าเพิ่ม 7%", 10)
    ns["propose_delete_table_item"](1, 1, 1, 3)
    queue = ns["get_chat_pending_actions"](1)

    ns["_handle_pending_queue_command"]("cancel_latest", queue, 1, 1)
    remaining = ns["get_chat_pending_actions"](1)
    assert len(remaining) == 1
    assert remaining[0]["action"]["type"] == "change_rate"


def test_cancel_all_clears_whole_queue(ns, doc1):
    ns["propose_change_rate"](1, 1, 1, "ภาษีมูลค่าเพิ่ม 7%", 10)
    ns["propose_delete_table_item"](1, 1, 1, 3)
    queue = ns["get_chat_pending_actions"](1)

    ns["_handle_pending_queue_command"]("cancel_all", queue, 1, 1)
    assert ns["get_chat_pending_actions"](1) == []


def test_queue_baseline_invalidated_by_external_edit_clears_queue(ns, doc1):
    """มีการแก้ไฟล์จากนอกคิวแทรกเข้ามาระหว่างที่คิวค้างอยู่ (เช่น undo หรือทางอื่น) -> ยืนยันต้องล้างคิวทิ้งแทน
    ที่จะเสี่ยงแก้ทับของเก่าที่ไม่ตรงกับ preview ที่เคยโชว์ไปแล้ว"""
    ns["propose_change_rate"](1, 1, 1, "ภาษีมูลค่าเพิ่ม 7%", 10)
    queue = ns["get_chat_pending_actions"](1)

    # จำลองการแก้จากนอกคิว: สร้างเวอร์ชันใหม่ตรงๆ ผ่าน create_new_version (ไม่ผ่านกลไกคิวเลย)
    doc = ns["get_editable_document_raw"](1)
    base_bytes, label_map = ns["get_working_copy"](doc)

    def _external_mutate(wb):
        ws = wb["Quotation"]
        ws.cell(row=16, column=2, value="Changed from outside the queue")

    ns["create_new_version"](1, _external_mutate, change_summary="external edit")

    answer = ns["_handle_pending_queue_command"]("confirm_all", queue, 1, 1)
    assert "เปลี่ยนไป" in answer
    assert ns["get_chat_pending_actions"](1) == []


# ---------- routing: โครงสร้างต้องไปทาง agent ไม่ใช่ _match_and_apply_excel_edit ----------

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
    assert ns["_is_document_reading_question"]("แก้ชื่อลูกค้าเป็น บริษัท ดูดี จำกัด ได้ไหม") is False


def test_match_and_apply_excel_edit_refuses_percent_on_formula_cell(ns, doc1):
    """safety net: เซลล์เป้าหมายเป็นสูตร + ค่าใหม่มี "%" -> ห้ามเขียนทับ ต้องบอกให้ใช้ change_rate แทน"""
    ns["_fake_client"].response_text = '{"matched": true, "label": "ภาษีมูลค่าเพิ่ม 7%", "new_value": "10%"}'
    result = ns["_match_and_apply_excel_edit"](1, 1, "เปลี่ยน VAT เป็น 10%")
    assert result["matched"] is False
    assert "change_rate" in result["message"] or "เปลี่ยนอัตรา" in result["message"]

    assert ns["_fake_db"].versions.get(1, []) == []
    vat_info = doc1["label_map"]["ภาษีมูลค่าเพิ่ม 7%"]
    wb_raw = openpyxl.load_workbook(io.BytesIO(doc1["original_bytes"]), data_only=False)
    cell_value = wb_raw[vat_info["sheet"]].cell(row=vat_info["row"], column=vat_info["col"]).value
    assert isinstance(cell_value, str) and cell_value.startswith("=")


def test_match_and_apply_excel_edit_still_works_for_plain_values(ns, doc1):
    ns["_fake_client"].response_text = '{"matched": true, "label": "รวมเงิน", "new_value": "99999"}'
    result = ns["_match_and_apply_excel_edit"](1, 1, "เปลี่ยนรวมเงินเป็น 99999")
    assert result["matched"] is True


def test_document_reading_question_fallback_when_client_fails(ns, doc1):
    ns["_fake_client"].raise_on_create = RuntimeError("simulated API failure")
    active_docs = [doc1]
    answer, used_doc_id = ns["_answer_document_reading_question"]("มีรายการอะไรบ้าง", active_docs, 1, 1)
    assert used_doc_id == 1
    expected = ns["_format_document_values_text"](ns["read_document_values"](doc1))
    assert answer == expected


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
    document_content_question = ns["_is_document_reading_question"](query) and not ns["_has_legal_keyword"](query)
    assert document_content_question is False


def test_answer_document_reading_question_uses_read_document_values_not_rag(ns, doc1):
    active_docs = [doc1]
    answer, used_doc_id = ns["_answer_document_reading_question"]("มีรายการอะไรบ้าง", active_docs, 1, 1)
    assert used_doc_id == 1
    sent_messages = ns["_fake_client"].last_call_kwargs["messages"]
    sent_content = sent_messages[0]["content"]
    assert "ตารางรายการ" in sent_content
    assert "รวมเงิน" in sent_content
    assert "ภาษีมูลค่าเพิ่ม 7%" in sent_content
    assert answer == ns["_fake_client"].response_text


# ---------- ส่วน 2: ห้าม AI แต่งผลการแก้เอง (ทดสอบ guard logic ของ rag_answer โดยตรง ด้วย fake agentic loop) ----------

def test_rag_answer_blocks_structural_route_with_no_tool_called(ns):
    """structural_to_agent แต่ไม่มี structural tool ถูกเรียกเลย (LLM mock เขียน "Preview...พิมพ์ยืนยัน" เอง) ->
    ผู้ใช้ต้องได้ข้อความคงที่ ไม่ใช่ข้อความที่ LLM แต่งขึ้น"""
    ns["_fake_agentic_loop"].set_response("Preview: จะเปลี่ยนราคาเป็น 70 บาท พิมพ์ยืนยันเพื่อดำเนินการ", [], False)
    answer, sources = ns["rag_answer"]("เปลี่ยน Box 4x4 เป็น 7x7", user_id=1, chat_id=1, structural_route=True)
    assert answer == ns["_STRUCTURAL_NO_TOOL_FALLBACK_MESSAGE"]
    assert "Preview" not in answer


def test_rag_answer_allows_structural_route_when_tool_called(ns):
    ns["_fake_agentic_loop"].set_response("จะเปลี่ยนราคาเป็น 70 บาท\n\nพิมพ์ 'ยืนยัน' เพื่อดำเนินการ หรือ 'ยกเลิก'", [], True)
    answer, sources = ns["rag_answer"]("เปลี่ยน Box 4x4 เป็น 7x7", user_id=1, chat_id=1, structural_route=True)
    assert "ยืนยัน" in answer  # ข้อความจริงจาก tool (จำลองผ่าน fake loop) ไม่ถูก block


def test_rag_answer_blocks_fabricated_edit_claim_without_receipts(ns, doc1):
    """LLM ตอบ "✓ แก้แล้ว" โดยไม่มี receipt จาก tool เลย (และมีเอกสารเปิดอยู่ในแชทจริง) -> ต้องถูก block"""
    ns["_fake_agentic_loop"].set_response("✓ แก้รายการเรียบร้อยแล้วครับ", [], False)
    answer, sources = ns["rag_answer"]("เปลี่ยน Box 4x4 เป็น 7x7", user_id=1, chat_id=1, structural_route=False)
    assert answer == ns["_STRUCTURAL_NO_TOOL_FALLBACK_MESSAGE"]


def test_rag_answer_blocks_claim_with_filler_words_between_verb_and_done(ns, doc1):
    """"ลบรายการที่ 3 แล้วครับ" (มีข้อความคั่นกลางระหว่าง "ลบ" กับ "แล้ว" ไม่ใช่ "ลบแล้ว" ติดกันเป๊ะ) ต้องยังถูก
    จับได้เหมือนเดิม แม้ structural_route=False — ไม่ได้พึ่งคำว่า "ยืนยัน"/"preview"/"✓" เลย"""
    ns["_fake_agentic_loop"].set_response("ลบรายการที่ 3 แล้วครับ", [], False)
    answer, sources = ns["rag_answer"]("ลบรายการที่ 3", user_id=1, chat_id=1, structural_route=False)
    assert answer == ns["_STRUCTURAL_NO_TOOL_FALLBACK_MESSAGE"]


def test_rag_answer_does_not_block_legal_confirm_word_even_with_active_docs(ns, doc1):
    """review PR #28 จุดที่ 3: คำตอบกฎหมายปกติที่มีคำว่า "ยืนยัน" ปนอยู่ (เช่น "ต้องยืนยันตัวตนก่อนยื่นภาษี") ต้อง
    ไม่ถูก block แม้มีเอกสารเปิดอยู่ในแชทจริงก็ตาม (เดิมกว้างไป บล็อกผิดเพราะเช็ค "ยืนยัน" ทุกเส้นทางที่มีเอกสาร
    เปิดอยู่ — "ยืนยัน"/"preview"/"✓" ควรเช็คเฉพาะตอน structural_route=True เท่านั้น เพราะมีแต่เส้นทางนั้นที่คำพวก
    นี้บ่งบอกการแต่ง preview/ยืนยันปลอมได้จริง)"""
    ns["_fake_agentic_loop"].set_response("ต้องยืนยันตัวตนก่อนยื่นภาษีครับ", [], False)
    answer, sources = ns["rag_answer"]("ต้องทำอะไรก่อนยื่นภาษี", user_id=1, chat_id=1, structural_route=False)
    assert answer != ns["_STRUCTURAL_NO_TOOL_FALLBACK_MESSAGE"]
    assert "ยืนยันตัวตน" in answer


@pytest.mark.parametrize("query,answer_text", [
    ("เพิ่มทุนจดทะเบียนต้องทำอย่างไร", "ถ้าต้องการเพิ่มทุนจดทะเบียน ต้องยื่นคำขอแล้วรออนุมัติ"),
    ("แก้ไขหนังสือบริคณห์สนธิต้องทำอย่างไร", "หากแก้ไขหนังสือบริคณห์สนธิแล้ว ต้องแจ้งกรมพัฒน์ภายใน 14 วัน"),
    ("ผู้ถือหุ้นต่างด้าวเปลี่ยนสัญชาติต้องตรวจอะไรบ้าง", "บริษัทที่เปลี่ยนสัญชาติผู้ถือหุ้นแล้วต้องตรวจ FBA"),
])
def test_rag_answer_does_not_block_normal_legal_sentences_with_active_docs(ns, doc1, query, answer_text):
    """review PR #28 รอบ 2: ประโยคกฎหมายปกติที่มีกริยา (เพิ่ม/แก้ไข/เปลี่ยน) + "แล้ว" ปนอยู่ในประโยคเดียวกัน แต่
    กริยาไม่ได้ติดกับคำนามของสิ่งที่ถูกแก้ในเอกสารเลย (เช่น "เพิ่มทุน" ไม่ใช่ "เพิ่มรายการ", "แก้ไขหนังสือบริคณห์
    สนธิ" ไม่ใช่ "แก้ไขไฟล์") ต้องไม่ถูก guard บล็อก แม้มีเอกสารเปิดอยู่ในแชทจริงก็ตาม (เดิม regex กว้างไปจับ
    กริยา+แล้ว ทั้งประโยคโดยไม่สนว่าติดกับคำนามอะไร)"""
    ns["_fake_agentic_loop"].set_response(answer_text, [], False)
    answer, sources = ns["rag_answer"](query, user_id=1, chat_id=1, structural_route=False)
    assert answer != ns["_STRUCTURAL_NO_TOOL_FALLBACK_MESSAGE"]


@pytest.mark.parametrize("answer_text", [
    "เปลี่ยน VAT เป็น 10% ให้แล้วครับ",
    "แก้ไขเรียบร้อยครับ",
    "ไฟล์ถูกแก้ไขแล้วครับ",
    "แก้ชื่อลูกค้าเป็น บริษัท ทดสอบ แล้วครับ",
])
def test_rag_answer_blocks_additional_fabricated_claim_phrasings(ns, doc1, answer_text):
    """review PR #28 รอบ 2: วลีอ้างว่าแก้สำเร็จแล้วอีกหลายแบบ (กริยาติดกับคำนามของสิ่งที่ถูกแก้ในเอกสารจริง หรือ
    พูดถึงไฟล์/เอกสารถูกแก้ไขตรงๆ หรือ "แก้ไขเรียบร้อย") ต้องยังถูก block เหมือนเดิมหลังจากทำ regex ให้แคบลง"""
    ns["_fake_agentic_loop"].set_response(answer_text, [], False)
    answer, sources = ns["rag_answer"]("แก้ไฟล์ให้หน่อย", user_id=1, chat_id=1, structural_route=False)
    assert answer == ns["_STRUCTURAL_NO_TOOL_FALLBACK_MESSAGE"]


def test_rag_answer_blocks_fabricated_claim_even_with_nonempty_top_chunks(ns, doc1):
    """review PR #28 รอบ 3: ctx["top_chunks"] ไม่ว่างเกือบทุกครั้งใน production จริง (rerank_with_scores คืน
    top_k=5 เสมอ ไม่มี threshold กรองทิ้ง) ไม่ได้แปลว่าคำตอบมาจากฐานความรู้จริงเสมอไป — guard ต้องไม่พึ่งมันเป็น
    เงื่อนไขเลย คำถามที่ไม่มีคำเชิงกฎหมายปนมา + คำตอบที่ยังอ้างว่าแก้สำเร็จ ต้องยังถูก block แม้ top_chunks ไม่ว่าง"""
    ns["_fake_rag_context"].top_chunks = ["chunk 1 เนื้อหาอะไรสักอย่างจากฐานความรู้", "chunk 2"]
    ns["_fake_agentic_loop"].set_response("ลบรายการที่ 3 แล้วครับ", [], False)
    answer, sources = ns["rag_answer"]("ลบรายการที่ 3", user_id=1, chat_id=1, structural_route=False)
    assert answer == ns["_STRUCTURAL_NO_TOOL_FALLBACK_MESSAGE"]


@pytest.mark.parametrize("query,answer_text", [
    ("เพิ่มทุนจดทะเบียนต้องทำอย่างไร", "ถ้าต้องการเพิ่มทุนจดทะเบียน ต้องยื่นคำขอแล้วรออนุมัติ"),
    ("แก้ไขหนังสือบริคณห์สนธิต้องทำอย่างไร", "หากแก้ไขหนังสือบริคณห์สนธิแล้ว ต้องแจ้งกรมพัฒน์ภายใน 14 วัน"),
    ("ผู้ถือหุ้นต่างด้าวเปลี่ยนสัญชาติต้องตรวจอะไรบ้าง", "บริษัทที่เปลี่ยนสัญชาติผู้ถือหุ้นแล้วต้องตรวจ FBA"),
])
def test_rag_answer_does_not_block_legal_sentences_with_nonempty_top_chunks(ns, doc1, query, answer_text):
    """ประโยคกฎหมายปกติทั้ง 3 ข้างบน (ที่ regex ไม่จับอยู่แล้วจากรอบ 2) ต้องไม่ถูก block แม้ ctx["top_chunks"] ไม่
    ว่างเหมือนสถานการณ์จริงใน production ก็ตาม (ยืนยันว่าไม่ถูก block ด้วย _has_legal_keyword(query) อย่างเดียว
    ไม่ใช่เพราะ top_chunks ว่างเปล่าในเทสต์)"""
    ns["_fake_rag_context"].top_chunks = ["chunk 1 เนื้อหาอะไรสักอย่างจากฐานความรู้", "chunk 2"]
    ns["_fake_agentic_loop"].set_response(answer_text, [], False)
    answer, sources = ns["rag_answer"](query, user_id=1, chat_id=1, structural_route=False)
    assert answer != ns["_STRUCTURAL_NO_TOOL_FALLBACK_MESSAGE"]


def test_rag_answer_does_not_block_normal_legal_answer_without_active_docs(ns):
    """คำตอบกฎหมายปกติที่บังเอิญมีคำว่า "ยืนยัน" ปนอยู่ (เช่น "ต้องยืนยันตัวตนก่อนยื่นภาษี") แต่ไม่มีเอกสารเปิดอยู่
    เลย -> ไม่ควรถูก guard บล็อก (กันคำว่า "ยืนยัน" ที่เป็นคำปกติถูกเหมาเข่งผิดๆ)"""
    ns["_fake_agentic_loop"].set_response("ต้องยืนยันตัวตนก่อนยื่นภาษีครับ", [], False)
    answer, sources = ns["rag_answer"]("ต้องทำอะไรก่อนยื่นภาษี", user_id=None, chat_id=None, structural_route=False)
    assert answer != ns["_STRUCTURAL_NO_TOOL_FALLBACK_MESSAGE"]
    assert "ยืนยันตัวตน" in answer
