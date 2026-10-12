"""ทดสอบเฟส 2 PR2 (แก้โครงสร้าง Excel: เปลี่ยนอัตรา / แก้-ลบ-เพิ่มรายการในตาราง — แก้ทันที ไม่มีขั้นยืนยัน/คิว
อีกต่อไป ดู feat/apply-immediately) ใช้ fixture workbook สังเคราะห์ (ไม่ใช่ไฟล์ลูกค้าจริง) และ fake DB/LLM client
ในหน่วยความจำล้วนๆ (ไม่ต่อ DB จริงเลย) ผ่าน tests/_harness.py"""
import io
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import pytest
import openpyxl

from _harness import (
    build_namespace, build_fixture_workbook, inject_formula_cache,
    FIXTURE_TOTAL, FIXTURE_VAT, FIXTURE_GRAND_TOTAL,
    _FakeCitation, _FakeAgenticTextBlock, _FakeAgenticResponse,
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


# ---------- ฟิกซ์เจอร์ที่สอง: รายละเอียดสินค้าแบบเหตุการณ์จริงที่ใช้ทดสอบ direct parse (_parse_direct_structural_
# commands) ซึ่งต้องจับคู่ "รายละเอียด" ของรายการในตารางด้วยข้อความจริง ("Box 4x4"/"ไส้ไก่") ไม่ใช่แค่ item_no —
# build_fixture_workbook() ของ _harness.py ใช้ชื่อ "Item A".."Item H" เฉยๆ ไม่พอทดสอบกรณีนี้ ----------
_BOX_ITEMS = [
    ("Box 4x4", 5, "ชิ้น", 50),
    ("Box 6x6", 3, "ชิ้น", 70),
    ("ไส้ไก่", 10, "ชิ้น", 20),
    ("Item D", 2, "ชุด", 300),
    ("Item E", 1, "เครื่อง", 5000),
    ("Item F", 4, "ชิ้น", 25),
    ("Item G", 8, "ชิ้น", 10),
    ("Item H", 6, "ชิ้น", 15),
]
_BOX_TOTAL = sum(qty * price for _, qty, _, price in _BOX_ITEMS)
_BOX_VAT_7 = round(_BOX_TOTAL * 0.07, 2)
_BOX_GRAND_7 = _BOX_TOTAL + _BOX_VAT_7


def _build_box_quote_workbook():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Quotation"
    ws["A15"] = "ลำดับ"
    ws["B15"] = "รายการ"
    ws["D15"] = "จำนวน"
    ws["E15"] = "หน่วย"
    ws["F15"] = "ราคา/หน่วย"
    ws["G15"] = "จำนวนเงิน"

    cache_values = {}
    for i, (desc, qty, unit, price) in enumerate(_BOX_ITEMS):
        row = 16 + i
        ws.cell(row=row, column=1, value=i + 1)
        ws.cell(row=row, column=2, value=desc)
        ws.cell(row=row, column=4, value=qty)
        ws.cell(row=row, column=5, value=unit)
        ws.cell(row=row, column=6, value=price)
        ws.cell(row=row, column=7, value=f"=D{row}*F{row}")
        cache_values[f"G{row}"] = qty * price

    last_item_row = 16 + len(_BOX_ITEMS) - 1
    ws["A26"] = "รวมเงิน"
    ws["G26"] = f"=SUM(G16:G{last_item_row})"
    cache_values["G26"] = _BOX_TOTAL

    ws["A27"] = "ภาษีมูลค่าเพิ่ม 7%"
    ws["G27"] = "=SUM(G26)*7/100"
    cache_values["G27"] = _BOX_VAT_7

    ws["A28"] = "รวมราคาทั้งสิ้น"
    ws["G28"] = "=SUM(G26:G27)"
    cache_values["G28"] = _BOX_GRAND_7

    buf = io.BytesIO()
    wb.save(buf)
    return inject_formula_cache(buf.getvalue(), cache_values)


@pytest.fixture
def box_doc(ns):
    raw = _build_box_quote_workbook()
    label_map = ns["_extract_excel_labels"](raw)
    ns["_fake_db"].add_document(1, "quote.xlsx", raw, label_map, chat_id=1, user_id=1)
    return ns["_fake_db"].documents[1]


def _build_quote_workbook_with_items(items):
    """ทั่วไปกว่า _build_box_quote_workbook — รับรายการ (desc, qty, unit, price) เองได้ ใช้ทดสอบว่า advisor
    ส่งตารางรายการเป็น context ได้กับสินค้า "ทุกประเภท" ไม่ใช่เฉพาะกรณี Box/ไส้ไก่ (ดู feat/advisor-route:
    เกณฑ์เปรียบเทียบ/ผลกระทบ/หน่วยขายต้องเป็นกติกาทั่วไป ไม่ hardcode ตามสินค้า)"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Quotation"
    ws["A15"] = "ลำดับ"
    ws["B15"] = "รายการ"
    ws["D15"] = "จำนวน"
    ws["E15"] = "หน่วย"
    ws["F15"] = "ราคา/หน่วย"
    ws["G15"] = "จำนวนเงิน"

    cache_values = {}
    for i, (desc, qty, unit, price) in enumerate(items):
        row = 16 + i
        ws.cell(row=row, column=1, value=i + 1)
        ws.cell(row=row, column=2, value=desc)
        ws.cell(row=row, column=4, value=qty)
        ws.cell(row=row, column=5, value=unit)
        ws.cell(row=row, column=6, value=price)
        ws.cell(row=row, column=7, value=f"=D{row}*F{row}")
        cache_values[f"G{row}"] = qty * price

    last_item_row = 16 + len(items) - 1
    total = sum(qty * price for _, qty, _, price in items)
    ws["A26"] = "รวมเงิน"
    ws["G26"] = f"=SUM(G16:G{last_item_row})"
    cache_values["G26"] = total

    buf = io.BytesIO()
    wb.save(buf)
    return inject_formula_cache(buf.getvalue(), cache_values)


def _current(ns, document_id=1):
    """อ่านค่าปัจจุบันของเอกสาร (สำเนาทำงานล่าสุด) — ใช้ตรวจผลหลัง apply ทันที"""
    doc = ns["get_editable_document_raw"](document_id)
    return ns["read_document_values"](doc)


# ---------- change_rate (แก้ทันที) ----------

def test_apply_change_rate_to_10_percent(ns, doc1):
    result = ns["apply_change_rate"](1, 1, "ภาษีมูลค่าเพิ่ม 7%", 10)
    assert result["matched"] is True
    assert "เปลี่ยนอัตราของ" in result["message"]
    assert "10%" in result["message"]
    assert "พิมพ์ 'ย้อนการแก้'" in result["message"]

    expected_vat = round(FIXTURE_TOTAL * 0.10, 2)
    expected_total = FIXTURE_TOTAL + expected_vat
    assert str(expected_vat) in result["message"]
    assert str(expected_total) in result["message"]

    data = _current(ns)
    label_values = {l["label"]: l["value"] for l in data["labels"]}
    # label เปลี่ยนชื่อตามข้อความจริงในเซลล์ป้ายหลังแก้ (ดู _current_label_display_name) — ไม่ใช่ key เดิม
    # "ภาษีมูลค่าเพิ่ม 7%" ที่แช่แข็งไว้ตั้งแต่ extract ครั้งแรกอีกต่อไป
    assert "ภาษีมูลค่าเพิ่ม 7%" not in label_values
    assert label_values["ภาษีมูลค่าเพิ่ม 10%"] == expected_vat
    assert label_values["รวมราคาทั้งสิ้น"] == expected_total
    assert len(ns["_fake_db"].versions.get(1, [])) == 1  # แก้ทันทีเป็น 1 เวอร์ชัน ไม่มีคิวมาเกี่ยวข้องเลย


def test_apply_change_rate_message_and_values_show_current_label_not_stale_name(ns, doc1):
    """เหตุการณ์จริงที่แก้: เปลี่ยน VAT จาก 7% เป็น 25% สำเร็จ แต่ข้อความผลลัพธ์/read_document_values ยังโชว์ชื่อ
    เดิม "ภาษีมูลค่าเพิ่ม 7%" ค้างอยู่ตลอดไป (label_map key แช่แข็งไว้ตั้งแต่ extract ครั้งแรก ไม่เคยอัปเดตตาม
    ข้อความจริงในเซลล์ป้าย) ตอนนี้ต้องโชว์ "ภาษีมูลค่าเพิ่ม 25%" ทั้งในข้อความผลลัพธ์และ read_document_values"""
    result = ns["apply_change_rate"](1, 1, "ภาษีมูลค่าเพิ่ม 7%", 25)
    assert result["matched"] is True
    assert "'ภาษีมูลค่าเพิ่ม 25%'" in result["message"]
    assert "ภาษีมูลค่าเพิ่ม 7%" not in result["message"]
    assert "ยอดล่าสุด" in result["message"]
    assert "ภาษีมูลค่าเพิ่ม 25%:" in result["message"]

    data = _current(ns)
    label_values = {l["label"]: l["value"] for l in data["labels"]}
    assert "ภาษีมูลค่าเพิ่ม 7%" not in label_values
    expected_vat = round(FIXTURE_TOTAL * 0.25, 2)
    assert label_values["ภาษีมูลค่าเพิ่ม 25%"] == expected_vat


def test_apply_change_rate_noop_when_already_at_target_rate(ns, doc1):
    """fixture เริ่มต้นที่ VAT 7% อยู่แล้ว — เสนอเปลี่ยนเป็น 7% (เท่าเดิม) ต้องไม่สร้างเวอร์ชันใหม่เลย"""
    result = ns["apply_change_rate"](1, 1, "ภาษีมูลค่าเพิ่ม 7%", 7)
    assert result["matched"] is False
    assert "อยู่แล้ว" in result["message"]
    assert ns["_fake_db"].versions.get(1, []) == []


def test_apply_change_rate_unsupported_formula_is_rejected(ns, doc1):
    # "รวมเงิน" เป็นสูตร =SUM(...) ไม่ตรงรูปแบบอัตราที่รองรับเลย
    result = ns["apply_change_rate"](1, 1, "รวมเงิน", 10)
    assert result["matched"] is False
    assert "ไม่ใช่รูปแบบอัตราที่รองรับ" in result["message"]
    assert ns["_fake_db"].versions.get(1, []) == []


def test_apply_change_rate_position_based_lookup_matches_real_file_scale(ns):
    """ทวนด้วยตัวเลขจากไฟล์จริง (ป่าไม้เขียว, PR #25): รวมเงิน 58150, VAT 7%->10% ต้องได้ 5815 และรวมราคา
    ทั้งสิ้น 58150+5815=63965 — ยืนยันว่าหา "ค่าใหม่" ถูกต้องด้วยตำแหน่งเซลล์ แม้ label มี % อยู่ในชื่อ"""
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

    result = ns["apply_change_rate"](2, 1, "ภาษีมูลค่าเพิ่ม 7%", 10)
    assert result["matched"] is True
    assert "5815" in result["message"]
    assert "63965" in result["message"]


# ---------- update_table_item (แก้ทันที) ----------

def test_apply_update_table_item_description(ns, doc1):
    result = ns["apply_update_table_item"](1, 1, 3, {"description": "Item C renamed"})
    assert result["matched"] is True
    assert "Item C renamed" in result["message"]
    assert _current(ns)["tables"][0]["rows"][2]["values"]["description"] == "Item C renamed"


def test_apply_update_table_item_noop_when_value_unchanged(ns, doc1):
    result = ns["apply_update_table_item"](1, 1, 1, {"description": "Item A"})
    assert result["matched"] is False
    assert "อยู่แล้ว" in result["message"]
    assert ns["_fake_db"].versions.get(1, []) == []


def test_apply_update_table_item_uses_thai_field_names(ns, doc1):
    result = ns["apply_update_table_item"](1, 1, 1, {"description": "Box 7x7", "qty": 15})
    assert result["matched"] is True
    assert "รายละเอียด เปลี่ยนเป็น Box 7x7" in result["message"]
    assert "จำนวน เปลี่ยนเป็น 15" in result["message"]
    assert "description" not in result["message"]
    assert "qty" not in result["message"]


# ---------- delete_table_item (แก้ทันที) ----------

def test_apply_delete_middle_item_shifts_and_renumbers_continuously(ns, doc1):
    """ลบรายการที่ 3 จาก 8 รายการ: ลำดับต้องเป็น 1..7 ต่อเนื่อง และรายละเอียดต้องเลื่อนขึ้นถูกแถว (รายการที่ 3
    ใหม่ต้องเป็นของเดิมรายการที่ 4) — รายการที่ 2 ต้องยังอยู่ ไม่ถูกลบผิดแถว"""
    result = ns["apply_delete_table_item"](1, 1, 3)
    assert result["matched"] is True

    filled = _current(ns)["tables"][0]["rows"]
    assert len(filled) == 7
    assert [r["values"]["no"] for r in filled] == [1, 2, 3, 4, 5, 6, 7]
    assert filled[1]["values"]["description"] == "Item B"
    assert filled[2]["values"]["description"] == "Item D"
    assert filled[2]["values"]["qty"] == 2
    assert filled[2]["values"]["unit_price"] == 300


def test_undo_after_apply_delete_restores_original(ns, doc1):
    ns["apply_delete_table_item"](1, 1, 3)

    undo_result = ns["undo_last_edit"](1, 1)
    assert undo_result["matched"] is True

    filled = _current(ns)["tables"][0]["rows"]
    assert len(filled) == 8
    assert filled[2]["values"]["description"] == "Item C"


def test_apply_delete_rejected_when_amount_not_formula(ns):
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

    result = ns["apply_delete_table_item"](3, 1, 1)
    assert result["matched"] is False
    assert result["message"] == ns["_TABLE_NON_FORMULA_AMOUNT_ERROR"]


# ---------- add_table_item (แก้ทันที) ----------

def test_apply_add_table_item_until_full(ns, doc1):
    result = ns["apply_add_table_item"](1, 1, {"description": "Item I", "qty": 2, "unit": "ชิ้น", "unit_price": 20})
    assert result["matched"] is True
    assert "รายละเอียด=Item I" in result["message"]

    result2 = ns["apply_add_table_item"](1, 1, {"description": "Item J", "qty": 1, "unit": "ชิ้น", "unit_price": 5})
    assert result2["matched"] is True

    # ตอนนี้เต็มแล้ว (10/10) — ลองเพิ่มอีกต้องถูกปฏิเสธทันที ไม่สร้างเวอร์ชัน
    versions_before = len(ns["_fake_db"].versions.get(1, []))
    rejected = ns["apply_add_table_item"](1, 1, {"description": "Item K", "qty": 1, "unit": "ชิ้น", "unit_price": 1})
    assert rejected["matched"] is False
    assert "เต็มแล้ว" in rejected["message"]
    assert "10" in rejected["message"]
    assert len(ns["_fake_db"].versions.get(1, [])) == versions_before


def test_apply_add_rejected_when_amount_not_formula(ns):
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

    result = ns["apply_add_table_item"](4, 1, {"description": "New", "qty": 1, "unit": "ชิ้น", "unit_price": 1})
    assert result["matched"] is False
    assert result["message"] == ns["_TABLE_NON_FORMULA_AMOUNT_ERROR"]


# ---------- คำสั่งยืนยัน/ยกเลิกแบบเดิม — ตอนนี้แค่ "ตรวจจับ" ว่าควรตอบคงที่ ไม่ใช่จัดการคิวอีกต่อไป ----------

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


@pytest.mark.parametrize("text,expected", [
    ("ยืนยันทั้งหมดครับ", "confirm_all"),
    ("ยืนยันทั้งหมดเลย", "confirm_all"),
    ("ยืนยันหมดเลยนะ", "confirm_all"),
    ("ยกเลิกทั้งหมดเลย", "cancel_all"),
    ("ยืนยันด้วย", "confirm_single"),
    ("ยืนยันล่าสุดครับ", "confirm_latest"),
])
def test_classify_tolerates_trailing_particles(ns, text, expected):
    assert ns["_classify_pending_queue_command"](text) == expected


@pytest.mark.parametrize("text", ["ไม่ยืนยันนะ", "อย่าเพิ่งยืนยัน", "ยังไม่ยืนยันครับ"])
def test_classify_still_rejects_negated_forms_with_particles(ns, text):
    assert ns["_classify_pending_queue_command"](text) is None


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
    เป็นความตั้งใจจะยืนยัน/ยกเลิก — ใช้คำเชิงคำถามเดียวกับ _OPEN_COMMAND_QUESTION_WORDS กัน"""
    assert ns["_mentions_pending_queue_intent"](text) is False


def test_mentions_pending_queue_intent_false_when_no_confirm_cancel_word(ns):
    assert ns["_mentions_pending_queue_intent"]("เปลี่ยน Box 4x4 เป็น 7x7") is False


def test_ask_question_would_show_fixed_message_for_confirm_word_with_doc_open(ns, doc1):
    """"ยืนยัน" ขณะมีเอกสารเปิดอยู่ (ไม่มีขั้นยืนยัน/คิวแล้ว) ต้องถูกตัวตรวจจับจับได้ (เพื่อให้ ask_question() ตอบ
    ข้อความคงที่ "ตอนนี้ระบบแก้ทันทีแล้วครับ..." แทนการส่งเข้า LLM — ข้อความคงที่นี้เขียนอยู่ใน ask_question() เอง
    ซึ่งหนักเกินจะ extract ทั้งฟังก์ชันผ่าน harness นี้ (เหมือน ask_question ทั้งฟังก์ชันในรอบก่อนๆ) ทดสอบแค่กลไก
    ตรวจจับที่ ask_question() พึ่งพาอยู่โดยตรงแทน"""
    assert ns["_classify_pending_queue_command"]("ยืนยัน") is not None


# ---------- direct parse: รูปแบบชัดเจนด้วยโค้ดล้วนๆ ไม่ผ่าน LLM (_parse_direct_structural_commands) ----------

def test_find_single_rate_label_finds_the_one_percent_label(ns, doc1):
    base_bytes, label_map = ns["get_working_copy"](doc1)
    assert ns["_find_single_rate_label"](base_bytes, label_map) == "ภาษีมูลค่าเพิ่ม 7%"


@pytest.mark.parametrize("query", ["เปลี่ยน VAT เป็น 10%", "แก้ภาษีเป็น10%", "เปลี่ยนภาษีมูลค่าเพิ่มเป็น 10%"])
def test_parse_direct_structural_commands_detects_rate_change(ns, doc1, query):
    base_bytes, label_map = ns["get_working_copy"](doc1)
    commands = ns["_parse_direct_structural_commands"](query, base_bytes, label_map)
    assert commands == [{"type": "change_rate", "label": "ภาษีมูลค่าเพิ่ม 7%", "new_rate": 10.0}]


def test_parse_direct_structural_commands_detects_delete(ns, doc1):
    base_bytes, label_map = ns["get_working_copy"](doc1)
    commands = ns["_parse_direct_structural_commands"]("ลบรายการที่ 3", base_bytes, label_map)
    assert commands == [{"type": "delete_table_item", "item_no": 3}]


def test_parse_direct_structural_commands_returns_none_for_multiple_rate_labels(ns):
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

    commands = ns["_parse_direct_structural_commands"]("เปลี่ยน VAT เป็น 10%", raw, label_map)
    assert commands is None


@pytest.mark.parametrize("query", ["เปลี่ยน Box 4x4 เป็น 7x7", "เปลี่ยน box 4X4 เป็น Box 7x7"])
def test_parse_direct_structural_commands_detects_description_change(ns, box_doc, query):
    """"Box 4x4" -> "7x7" (ส่วนท้ายสั้นๆ) ประกอบเป็น "Box 7x7" / "box 4X4" (ตัวพิมพ์ต่าง) -> "Box 7x7" (ชื่อเต็ม)
    ใช้ตรงๆ — ทั้งสองแบบต้อง parse ได้โดยไม่ต้องผ่าน agent เลย"""
    base_bytes, label_map = ns["get_working_copy"](box_doc)
    commands = ns["_parse_direct_structural_commands"](query, base_bytes, label_map)
    assert commands == [{"type": "update_table_item", "item_no": 1, "fields": {"description": "Box 7x7"}}]


def test_parse_direct_structural_commands_detects_qty_field_update(ns, box_doc):
    base_bytes, label_map = ns["get_working_copy"](box_doc)
    commands = ns["_parse_direct_structural_commands"]("เปลี่ยนจำนวน Box 4x4 เป็น 30", base_bytes, label_map)
    assert commands == [{"type": "update_table_item", "item_no": 1, "fields": {"qty": 30.0}}]


def test_parse_direct_structural_commands_detects_unit_price_field_update_by_item_no(ns, box_doc):
    base_bytes, label_map = ns["get_working_copy"](box_doc)
    commands = ns["_parse_direct_structural_commands"]("เปลี่ยนราคารายการที่ 2 เป็น 80", base_bytes, label_map)
    assert commands == [{"type": "update_table_item", "item_no": 2, "fields": {"unit_price": 80.0}}]


def test_parse_direct_structural_commands_returns_none_when_target_not_found(ns, box_doc):
    base_bytes, label_map = ns["get_working_copy"](box_doc)
    assert ns["_parse_direct_structural_commands"]("เปลี่ยน Box 9x9 เป็น 1x1", base_bytes, label_map) is None


def test_resolve_table_target_returns_none_when_description_matches_multiple_rows(ns):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Quotation"
    ws["A15"] = "ลำดับ"
    ws["B15"] = "รายการ"
    ws["D15"] = "จำนวน"
    ws["E15"] = "หน่วย"
    ws["F15"] = "ราคา/หน่วย"
    ws["G15"] = "จำนวนเงิน"
    for i, row in enumerate((16, 17)):
        ws.cell(row=row, column=1, value=i + 1)
        ws.cell(row=row, column=2, value="Widget")  # รายละเอียดซ้ำกันทั้งสองแถวโดยตั้งใจ
        ws.cell(row=row, column=4, value=1)
        ws.cell(row=row, column=5, value="ชิ้น")
        ws.cell(row=row, column=6, value=10)
        ws.cell(row=row, column=7, value=f"=D{row}*F{row}")
    buf = io.BytesIO()
    wb.save(buf)
    raw = inject_formula_cache(buf.getvalue(), {"G16": 10, "G17": 10})

    table = ns["_get_single_item_table"](raw)
    assert ns["_resolve_table_target"]("Widget", table) is None


def test_parse_direct_structural_commands_returns_none_for_ambiguous_box_change(ns, doc1):
    """fixture ปกติ ("Item A".."Item H") ไม่มีรายการชื่อ "Box 4x4" เลย — กำกวมเกินจะเดา ต้องปล่อยให้ agent
    ตัดสินใจแทน (คืน None)"""
    base_bytes, label_map = ns["get_working_copy"](doc1)
    assert ns["_parse_direct_structural_commands"]("เปลี่ยน Box 4x4 เป็น 7x7", base_bytes, label_map) is None


def test_parse_direct_structural_commands_handles_two_commands_joined_by_และ(ns, doc1):
    base_bytes, label_map = ns["get_working_copy"](doc1)
    commands = ns["_parse_direct_structural_commands"](
        "เปลี่ยน VAT เป็น 10% และลบรายการที่ 3", base_bytes, label_map
    )
    assert commands == [
        {"type": "change_rate", "label": "ภาษีมูลค่าเพิ่ม 7%", "new_rate": 10.0},
        {"type": "delete_table_item", "item_no": 3},
    ]


def test_parse_direct_structural_commands_none_if_any_clause_ambiguous(ns, doc1):
    """ถ้ามีประโยคย่อยไหนกำกวม/ไม่ตรงรูปแบบเลย แม้ประโยคอื่นจะ parse ได้ ต้องปล่อยทั้งข้อความให้ agent ตัดสินใจ
    (ไม่ทำแค่บางส่วนตามที่เข้าใจไปเอง)"""
    base_bytes, label_map = ns["get_working_copy"](doc1)
    commands = ns["_parse_direct_structural_commands"](
        "เปลี่ยน VAT เป็น 10% และเปลี่ยน Box 9x9 เป็น 1x1", base_bytes, label_map
    )
    assert commands is None


# ---------- สถานการณ์จริงเต็มรูปแบบ: parse + apply ทันที (จำลองสิ่งที่ ask_question() ทำจริงทีละคำสั่ง) ----------

def _apply_query(ns, document_id, query):
    """จำลองสิ่งที่ ask_question() ทำจริงสำหรับคำสั่งแก้โครงสร้างแบบชัดเจน — parse แล้ว apply ทันที คืนข้อความ
    ผลลัพธ์รวม (ใช้ในเทสต์เท่านั้น ไม่ใช่โค้ดจริงใน main.py)"""
    doc = ns["get_editable_document_raw"](document_id)
    base_bytes, label_map = ns["get_working_copy"](doc)
    commands = ns["_parse_direct_structural_commands"](query, base_bytes, label_map)
    assert commands is not None, f"คาดว่า '{query}' ต้อง parse ได้ด้วยโค้ดตรงๆ ไม่ต้องผ่าน agent"
    return ns["_apply_structural_commands"](document_id, 1, commands)


def test_full_scenario_vat_rename_delete_then_undo_three_times(ns, box_doc):
    """เหตุการณ์จริงทั้งชุด: "เปลี่ยน VAT เป็น 10%" -> "เปลี่ยน Box 4x4 เป็น 7x7" -> "ลบรายการที่ 3" (ไส้ไก่) —
    ทุกคำสั่ง parse ได้ด้วยโค้ดตรงๆ ไม่ต้องผ่าน agent แก้ไฟล์ทันทีทีละคำสั่ง (3 เวอร์ชันแยกกัน) แล้ว undo 3 ครั้ง
    ต้องกลับไปเป็นไฟล์เดิมทุกค่าเป๊ะ"""
    _apply_query(ns, 1, "เปลี่ยน VAT เป็น 10%")
    _apply_query(ns, 1, "เปลี่ยน Box 4x4 เป็น 7x7")
    _apply_query(ns, 1, "ลบรายการที่ 3")

    assert len(ns["_fake_db"].versions[1]) == 3

    data = _current(ns)
    filled = data["tables"][0]["rows"]
    assert len(filled) == 7
    assert filled[0]["values"]["description"] == "Box 7x7"
    assert not any(r["values"]["description"] == "ไส้ไก่" for r in filled)
    label_values = {l["label"]: l["value"] for l in data["labels"]}
    # label เปลี่ยนชื่อตามข้อความจริงในเซลล์ป้ายหลังแก้ (ดู _current_label_display_name) — "7%" เดิมไม่มีอีกแล้ว
    assert "ภาษีมูลค่าเพิ่ม 7%" not in label_values
    assert label_values["ภาษีมูลค่าเพิ่ม 10%"] != _BOX_VAT_7

    for _ in range(3):
        undo_result = ns["undo_last_edit"](1, 1)
        assert undo_result["matched"] is True

    assert ns["_fake_db"].versions.get(1, []) == []
    final_data = _current(ns)
    final_filled = final_data["tables"][0]["rows"]
    assert len(final_filled) == 8
    assert final_filled[0]["values"]["description"] == "Box 4x4"
    assert any(r["values"]["description"] == "ไส้ไก่" for r in final_filled)
    final_label_values = {l["label"]: l["value"] for l in final_data["labels"]}
    assert final_label_values["ภาษีมูลค่าเพิ่ม 7%"] == _BOX_VAT_7
    assert final_label_values["รวมราคาทั้งสิ้น"] == _BOX_GRAND_7


def test_full_scenario_field_updates_apply_immediately_with_correct_totals(ns, box_doc):
    _apply_query(ns, 1, "เปลี่ยนจำนวน Box 4x4 เป็น 30")
    _apply_query(ns, 1, "เปลี่ยนราคารายการที่ 2 เป็น 80")

    data = _current(ns)
    filled = data["tables"][0]["rows"]
    assert filled[0]["values"]["qty"] == 30
    assert filled[1]["values"]["unit_price"] == 80

    expected_total = _BOX_TOTAL + (30 - 5) * 50 + (80 - 70) * 3
    label_values = {l["label"]: l["value"] for l in data["labels"]}
    assert label_values["รวมเงิน"] == expected_total


def test_full_scenario_multi_command_single_message_creates_two_versions(ns, doc1):
    answer = _apply_query(ns, 1, "เปลี่ยน VAT เป็น 10% และลบรายการที่ 3")
    assert len(ns["_fake_db"].versions[1]) == 2
    assert answer.count("✓") == 2

    data = _current(ns)
    filled = data["tables"][0]["rows"]
    assert len(filled) == 7
    label_values = {l["label"]: l["value"] for l in data["labels"]}
    assert "ภาษีมูลค่าเพิ่ม 7%" not in label_values
    assert label_values["ภาษีมูลค่าเพิ่ม 10%"] != FIXTURE_VAT


def test_apply_structural_commands_reports_failure_without_stopping_remaining_commands(ns, doc1):
    """คำสั่งหนึ่งล้มเหลว (เช่น label ไม่เจอ) ไม่ควรทำให้คำสั่งอื่นในข้อความเดียวกันไม่ได้ทำไปด้วย — แต่ละคำสั่ง
    เป็นอิสระจากกัน"""
    commands = [
        {"type": "change_rate", "label": "label ไม่มีจริง", "new_rate": 10},
        {"type": "delete_table_item", "item_no": 3},
    ]
    answer = ns["_apply_structural_commands"](1, 1, commands)
    assert "✗" in answer
    assert "✓" in answer
    assert len(ns["_fake_db"].versions.get(1, [])) == 1  # เฉพาะคำสั่งที่สำเร็จเท่านั้นที่สร้างเวอร์ชัน


# ---------- routing ระดับ ask_question(): direct parse ต้องถูกลองก่อน _is_structural_edit_command เสมอ ----------
# (review PR #31 — เดิม direct parse ถูกเรียกเฉพาะเมื่อ _is_structural_edit_command(query) เป็น True เท่านั้น
# "เปลี่ยน Box 4x4 เป็น 7x7"/"เปลี่ยนจำนวน Box 4x4 เป็น 30" ไม่มีคำที่ฟังก์ชันนั้นรู้จักเลย เลยไม่เคยไปถึง
# _parse_direct_structural_commands — ตกไป classifier/agent ทั้งที่ parse ได้ชัดเจนด้วยโค้ดเฉยๆ)

def test_resolve_direct_structural_dispatch_for_description_change_bypasses_is_structural_edit_command(ns, box_doc):
    """"เปลี่ยน Box 4x4 เป็น 7x7" ไม่มีคำไหนที่ _is_structural_edit_command รู้จักเลย (ยืนยันด้วย assert ข้อแรก)
    แต่ _resolve_direct_structural_dispatch ต้องยังจับได้และ apply จริง ไม่ต้องรอผ่าน gate นั้นก่อน"""
    query = "เปลี่ยน Box 4x4 เป็น 7x7"
    assert ns["_is_structural_edit_command"](query) is False

    result = ns["_resolve_direct_structural_dispatch"](query, [box_doc], 1, 1)
    assert result is not None
    assert result["document_id"] == 1
    assert "✓" in result["answer"]

    filled = _current(ns)["tables"][0]["rows"]
    assert filled[0]["values"]["description"] == "Box 7x7"


@pytest.mark.parametrize("query,field,expected_value", [
    ("เปลี่ยนจำนวน Box 4x4 เป็น 30", "qty", 30),
    ("แก้หน่วยของ Box 6x6 เป็น กล่อง", "unit", "กล่อง"),
])
def test_resolve_direct_structural_dispatch_for_field_updates_bypasses_is_structural_edit_command(
    ns, box_doc, query, field, expected_value
):
    assert ns["_is_structural_edit_command"](query) is False

    result = ns["_resolve_direct_structural_dispatch"](query, [box_doc], 1, 1)
    assert result is not None

    filled = _current(ns)["tables"][0]["rows"]
    target_index = 0 if "4x4" in query else 1
    assert filled[target_index]["values"][field] == expected_value


def test_resolve_direct_structural_dispatch_returns_none_for_non_table_label_edit(ns, doc1):
    """"เปลี่ยนชื่อลูกค้าเป็น บริษัท ทดสอบ จำกัด" ไม่ใช่รายการในตาราง ("ชื่อลูกค้า" ไม่ใช่รายละเอียดสินค้าแถวไหน
    เลย) — ต้องคืน None ปล่อยให้ ask_question() ไปทาง classifier (intent="edit") ตามเดิม ไม่ใช่เดา apply ผิดที่"""
    result = ns["_resolve_direct_structural_dispatch"]("เปลี่ยนชื่อลูกค้าเป็น บริษัท ทดสอบ จำกัด", [doc1], 1, 1)
    assert result is None
    assert ns["_fake_db"].versions.get(1, []) == []


def test_resolve_direct_structural_dispatch_returns_none_for_document_reading_question(ns, doc1):
    """"มีรายการอะไรบ้าง" ไม่มีกริยาแก้ไฟล์เลย (ไม่มี "เปลี่ยน"/"แก้"/"ลบ") — parse ไม่ได้ตามธรรมชาติอยู่แล้ว คืน
    None เสมอ ไม่ต้องพึ่ง gate อะไรเพิ่ม (ใน ask_question() จริง คำถามแบบนี้ถูกดักไปตอบด้วย
    read_document_values ตั้งแต่ก่อนถึงจุดนี้อยู่แล้วเช่นกัน — ดู reading_question/document_content_question)"""
    assert ns["_resolve_direct_structural_dispatch"]("มีรายการอะไรบ้าง", [doc1], 1, 1) is None


def test_resolve_direct_structural_dispatch_returns_none_when_no_target_document(ns):
    assert ns["_resolve_direct_structural_dispatch"]("เปลี่ยน VAT เป็น 10%", [], 1, 1) is None


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
    """"ธุรกิจบัญชีสามมีอะไรบ้าง" ไม่มีคำอ้างอิงไฟล์เลย (ดู _DOCUMENT_REFERENCE_KEYWORDS หลัง tighten) —
    _is_document_reading_question จึงเป็น False อยู่แล้วด้วยตัวมันเอง แต่ที่สำคัญกว่าคือ _has_legal_keyword
    เป็น True ("บัญชี") ซึ่งทำให้ _classify_legal_or_advisor_route ตัดสินเป็น legal ก่อนเช็คอะไรอื่นเสมอ ไม่ว่า
    _is_document_reading_question จะตอบอย่างไรก็ตาม"""
    query = "ธุรกิจบัญชีสามมีอะไรบ้าง"
    assert ns["_is_document_reading_question"](query) is False
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
    จับได้เหมือนเดิม แม้ structural_route=False — ไม่ได้พึ่งคำว่า "ยืนยัน"/"preview"/"✓" เลย — นี่คือเหตุการณ์จริง
    ที่โจทย์ขอให้ยืนยันซ้ำ: agent (mock) อ้างว่าแก้แล้วโดยไม่มี receipt ต้องถูกบล็อก"""
    ns["_fake_agentic_loop"].set_response("ลบรายการที่ 3 แล้วครับ", [], False)
    answer, sources = ns["rag_answer"]("ลบรายการที่ 3", user_id=1, chat_id=1, structural_route=False)
    assert answer == ns["_STRUCTURAL_NO_TOOL_FALLBACK_MESSAGE"]


def test_rag_answer_does_not_block_legal_confirm_word_even_with_active_docs(ns, doc1):
    """คำตอบกฎหมายปกติที่มีคำว่า "ยืนยัน" ปนอยู่ (เช่น "ต้องยืนยันตัวตนก่อนยื่นภาษี") ต้องไม่ถูก block แม้มีเอกสาร
    เปิดอยู่ในแชทจริงก็ตาม — "ยืนยัน"/"preview"/"✓" เช็คเฉพาะตอน structural_route=True เท่านั้น"""
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
    """ประโยคกฎหมายปกติที่มีกริยา (เพิ่ม/แก้ไข/เปลี่ยน) + "แล้ว" ปนอยู่ในประโยคเดียวกัน แต่กริยาไม่ได้ติดกับคำนาม
    ของสิ่งที่ถูกแก้ในเอกสารเลย (เช่น "เพิ่มทุน" ไม่ใช่ "เพิ่มรายการ") ต้องไม่ถูก guard บล็อก"""
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
    ns["_fake_agentic_loop"].set_response(answer_text, [], False)
    answer, sources = ns["rag_answer"]("แก้ไฟล์ให้หน่อย", user_id=1, chat_id=1, structural_route=False)
    assert answer == ns["_STRUCTURAL_NO_TOOL_FALLBACK_MESSAGE"]


def test_rag_answer_blocks_fabricated_claim_even_with_nonempty_top_chunks(ns, doc1):
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
    ns["_fake_rag_context"].top_chunks = ["chunk 1 เนื้อหาอะไรสักอย่างจากฐานความรู้", "chunk 2"]
    ns["_fake_agentic_loop"].set_response(answer_text, [], False)
    answer, sources = ns["rag_answer"](query, user_id=1, chat_id=1, structural_route=False)
    assert answer != ns["_STRUCTURAL_NO_TOOL_FALLBACK_MESSAGE"]


def test_rag_answer_does_not_block_normal_legal_answer_without_active_docs(ns):
    ns["_fake_agentic_loop"].set_response("ต้องยืนยันตัวตนก่อนยื่นภาษีครับ", [], False)
    answer, sources = ns["rag_answer"]("ต้องทำอะไรก่อนยื่นภาษี", user_id=None, chat_id=None, structural_route=False)
    assert answer != ns["_STRUCTURAL_NO_TOOL_FALLBACK_MESSAGE"]
    assert "ยืนยันตัวตน" in answer


# ---------- feat/advisor-route: แยกคำถามไฟล์/กฎหมาย/advisor ออกจากกัน + advisor_answer ----------
# เหตุการณ์จริงที่แก้: คำถามแนะนำสินค้า (เช่น "แนะนำ Rack case server ที่คุ้มกว่า 9U") ขณะมีไฟล์เปิดอยู่ ถูก
# _is_document_reading_question เดิม (เช็คแค่คำเชิงคำถามทั่วไปอย่าง "บ้าง"/"ไหม") ตีความผิดเป็นคำถามเกี่ยวกับไฟล์
# ตอบ "ไฟล์ไม่มีข้อมูลรุ่นให้เลือก" หรือหลุดไปทาง rag_answer เดิม (มี context ฐานความรู้กฎหมายเสมอ) ได้คำตอบไม่
# เกี่ยวข้องเรื่อง FBA/BOI แทนที่จะแนะนำสินค้าจริง

@pytest.mark.parametrize("query", ["มีรายการอะไรบ้าง", "ยอดรวมเท่าไร"])
def test_document_reading_keywords_still_match_quotation_questions(ns, query):
    assert ns["_is_document_reading_question"](query) is True


@pytest.mark.parametrize("query", ["มีรุ่นไหนให้เลือกบ้าง พร้อมบอกราคามาเลย", "แนะนำ Rack case server 12U"])
def test_advisor_style_questions_are_not_document_reading_questions(ns, doc1, query):
    """คำถามแนะนำสินค้าทั่วไปขณะมีไฟล์เปิดอยู่ ต้องไม่ถูกตีความเป็นคำถามเกี่ยวกับไฟล์เลย ทั้งทางคำเชิงคำถามทั่วไป
    (ไม่มีคำอ้างอิงไฟล์ใน _DOCUMENT_REFERENCE_KEYWORDS) และทางชื่อรายการ/label ในไฟล์ (fixture ไม่มี "server"/
    "รุ่น" อยู่ในตารางหรือ label เลย)"""
    assert ns["_is_document_reading_question"](query) is False
    assert ns["_query_mentions_table_item_or_label"](query, doc1) is False


def test_classify_question_route_legal_keyword_short_circuits_without_llm_call(ns):
    """"ธุรกิจบัญชีสามมีอะไรบ้าง" มีคำเชิงความรู้กฎหมาย ("บัญชี") ปนมา -> ต้องตัดสินเป็น legal ทันทีด้วยโค้ด
    (_has_legal_keyword) ไม่ต้องเรียก Haiku classifier เลยสักครั้ง (ประหยัด API call)"""
    result = ns["_classify_question_route"]("ธุรกิจบัญชีสามมีอะไรบ้าง")
    assert result == "legal"
    assert ns["_fake_client"].last_call_kwargs is None


@pytest.mark.parametrize("query", ["มีรุ่นไหนให้เลือกบ้าง พร้อมบอกราคามาเลย", "แนะนำ Rack case server 12U"])
def test_classify_question_route_falls_back_to_llm_classifier_for_advisor_questions(ns, query):
    """ไม่มีคำเชิงความรู้กฎหมายปนมาเลย -> ต้องถาม Haiku classifier (ดู _classify_legal_or_advisor_route) ซึ่ง
    ในเทสต์นี้ตั้งค่าให้ตอบ "advisor" กลับมา"""
    ns["_fake_client"].response_text = '{"category": "advisor"}'
    result = ns["_classify_question_route"](query)
    assert result == "advisor"
    assert ns["_fake_client"].last_call_kwargs is not None


def test_classify_legal_or_advisor_route_defaults_to_advisor_on_api_failure(ns):
    ns["_fake_client"].raise_on_create = RuntimeError("simulated API failure")
    assert ns["_classify_legal_or_advisor_route"]("คำถามอะไรสักอย่างที่กำกวม") == "advisor"


def test_advisor_answer_sends_chat_history_and_has_no_legal_kb_text_in_system_prompt(ns):
    """advisor_answer ต้องส่งประวัติสนทนาเข้าไปด้วย (ให้ AI เข้าใจว่าผู้ใช้เพิ่งตอบคำถามที่เคยถูกถามกลับไปแล้ว)
    และ system prompt ต้องไม่มีคำเชิงความรู้ฐาน KB กฎหมายเฉพาะทาง (BOI/FBA/พ.ร.บ.ต่างด้าว) ปนเลยสักคำ — ต่างจาก
    system prompt ของ rag_answer (_prepare_rag_context) ที่มีคำพวกนี้อยู่เสมอ"""
    ns["_fake_agentic_loop"].set_response("แนะนำรุ่น A, B, C ครับ", [], False)
    history = [
        {"role": "user", "content": "ฉันกำลังหา Rack case server ที่น่าจะคุ้มกว่า 9U แนะนำหน่อย"},
        {"role": "assistant", "content": "งบประมาณเท่าไรครับ"},
        {"role": "user", "content": "งบไม่จำกัด ส่วนอย่างอื่นเอาแบบดีที่สุด"},
    ]
    answer, sources = ns["advisor_answer"](
        "มีรุ่นไหนให้เลือกบ้าง พร้อมบอกราคามาเลย", history=history, user_id=1, chat_id=1, active_docs=[],
    )
    assert sources == []
    call_args = ns["_fake_agentic_loop"].last_call_args
    assert call_args["tools_override"] == ns["ADVISOR_TOOLS"]
    sent_messages = call_args["messages"]
    assert len(sent_messages) == len(history) + 1
    assert sent_messages[-1]["content"] == "มีรุ่นไหนให้เลือกบ้าง พร้อมบอกราคามาเลย"
    system_prompt = call_args["system_prompt"]
    for forbidden in ("BOI", "FBA", "พ.ร.บ.การประกอบธุรกิจของคนต่างด้าว"):
        assert forbidden not in system_prompt


def test_advisor_answer_offers_replacement_when_asked_item_exists_in_open_quotation_table(ns, box_doc):
    """มีใบเสนอราคาเปิดอยู่ และคำถามเอ่ยถึงรายการที่มีอยู่จริงในตารางตรงๆ ("Box 4x4") -> ต่อท้ายคำตอบด้วยข้อความ
    เสนอ "แทนรายการที่ N ด้วยรุ่นที่เลือก" (การแทนจริงยังไม่ต่อสายใน PR นี้ — แค่ข้อความเสนอ)"""
    ns["_fake_agentic_loop"].set_response("Box 4x4 รุ่นที่คุ้มกว่าคือ ... ครับ", [], False)
    answer, _ = ns["advisor_answer"](
        "Box 4x4 มีรุ่นที่คุ้มกว่านี้ไหม", history=[], user_id=1, chat_id=1, active_docs=[box_doc],
    )
    assert ns["_ADVISOR_REPLACEMENT_OFFER_LINE"] in answer


def test_advisor_answer_does_not_offer_replacement_when_question_unrelated_to_open_table(ns, box_doc):
    ns["_fake_agentic_loop"].set_response("แนะนำรุ่น A, B, C ครับ", [], False)
    answer, _ = ns["advisor_answer"](
        "แนะนำ Rack case server 12U", history=[], user_id=1, chat_id=1, active_docs=[box_doc],
    )
    assert ns["_ADVISOR_REPLACEMENT_OFFER_LINE"] not in answer


# ---------- feat/advisor-route (review fix): web_search แยกคำตอบเป็นหลาย text block กลางประโยค + pause_turn ----------
# เหตุการณ์จริงที่แก้: web_search ทำให้ Claude แบ่งคำตอบเป็นหลาย text block ตรงจุดอ้างอิงแหล่งที่มา (มักกลาง
# ประโยค) เดิมต่อ text block ด้วย "\n\n" ทำให้ประโยคแตกเป็นย่อหน้าใหม่กลางคัน — กระทบเส้นทาง advisor ที่ค้นเว็บบ่อย
# ที่สุด แต่ใช้ฟังก์ชันร่วมกับเส้นทางกฎหมายเดิมด้วย (ดู _build_final_answer_with_citations) และ stop_reason
# "pause_turn" (server tool อย่าง web_search ใช้เวลานาน) เดิมถูกตีความเป็น "จบเทิร์นแล้ว" (break) ทั้งที่ต้องส่ง
# messages เดิมกลับไปเรียกต่อในรอบถัดไป (ดู run_agentic_tool_loop)

def test_build_final_answer_with_citations_joins_text_blocks_without_breaking_sentences(ns):
    blocks = [
        _FakeAgenticTextBlock("กล้องรุ่น A ราคาประมาณ"),
        _FakeAgenticTextBlock(" 1,500 บาท", citations=[_FakeCitation("https://a.com", "ร้าน A")]),
        _FakeAgenticTextBlock(" ตามที่ประกาศไว้", citations=[_FakeCitation("https://a.com", "ร้าน A")]),
        _FakeAgenticTextBlock(" และยังมีอีกร้านที่เสนอราคา"),
        _FakeAgenticTextBlock(" 1,450 บาท", citations=[_FakeCitation("https://b.com", "ร้าน B")]),
    ]
    text = ns["_build_final_answer_with_citations"](blocks)

    # ต่อกันเป็นประโยคเดียวต่อเนื่อง ไม่มี "\n\n" คั่นกลางประโยคจาก citation เลย ("\n\n" ก่อนหัวข้อ "แหล่งที่มา:"
    # เป็นตัวคั่นที่ตั้งใจเติมเอง ไม่ใช่สิ่งที่ต้องเช็ค — ตัดส่วนท้ายนั้นออกก่อน)
    body = text.split("\n\nแหล่งที่มา:")[0]
    assert body == "กล้องรุ่น A ราคาประมาณ 1,500 บาท ตามที่ประกาศไว้ และยังมีอีกร้านที่เสนอราคา 1,450 บาท"
    assert "\n\n" not in body

    # citation ตัดซ้ำตาม url (https://a.com เจอ 2 ครั้งแต่นับครั้งเดียว) ไม่เกิน 5 แหล่ง
    assert text.count("แหล่งที่มา:") == 1
    assert "1. ร้าน A — https://a.com" in text
    assert "2. ร้าน B — https://b.com" in text
    assert text.count("ร้าน A —") == 1


def test_build_final_answer_with_citations_caps_at_five_sources(ns):
    blocks = [
        _FakeAgenticTextBlock(f"ข้อความที่ {i}", citations=[_FakeCitation(f"https://site{i}.com", f"แหล่ง {i}")])
        for i in range(8)
    ]
    text = ns["_build_final_answer_with_citations"](blocks)
    assert text.count(" — https://site") == 5
    assert "https://site5.com" not in text  # index 0-4 เท่านั้น (5 แหล่งแรก)


def test_build_final_answer_with_citations_no_source_heading_without_any_citation(ns):
    blocks = [_FakeAgenticTextBlock("คำตอบธรรมดา"), _FakeAgenticTextBlock(" ไม่มีการอ้างอิงเว็บเลย")]
    text = ns["_build_final_answer_with_citations"](blocks)
    assert text == "คำตอบธรรมดา ไม่มีการอ้างอิงเว็บเลย"
    assert "แหล่งที่มา" not in text


def test_run_agentic_tool_loop_continues_after_pause_turn_then_ends_on_second_call(ns):
    """stop_reason="pause_turn" (server tool เช่น web_search ยังทำงานไม่เสร็จ) ต้องไม่ break ทันที ต้องเรียก API
    ต่อในรอบถัดไปจนกว่าจะได้ stop_reason อื่น (ในเทสต์นี้คือ end_turn ในรอบที่สอง) และข้อความ/citation ของรอบ
    pause_turn (ส่วนแรกของคำตอบ) ต้องไม่หายไป — ต้องถูกสะสมไว้แล้วต่อกับรอบสุดท้ายตอนประกอบคำตอบ (เหตุการณ์จริงที่
    แก้: เดิมใช้แค่ response.content ของรอบสุดท้ายรอบเดียว ทำให้ข้อความ/citation ของรอบ pause_turn ก่อนหน้าหายไป)"""
    real_loop = ns["_real_run_agentic_tool_loop"]
    fake_client = ns["_fake_client"]
    fake_client.response_queue = [
        _FakeAgenticResponse(
            [_FakeAgenticTextBlock("ส่วนแรก ", citations=[_FakeCitation("https://a.com", "แหล่ง A")])],
            stop_reason="pause_turn",
        ),
        _FakeAgenticResponse(
            [_FakeAgenticTextBlock("ส่วนหลัง", citations=[_FakeCitation("https://b.com", "แหล่ง B")])],
            stop_reason="end_turn",
        ),
    ]
    answer, receipts, structural = real_loop(
        "system prompt ทดสอบ", [{"role": "user", "content": "แนะนำกล้องหน่อย"}],
        user_id=None, chat_id=None, tools_override=[],
    )
    assert answer.startswith("ส่วนแรก ส่วนหลัง")
    assert "แหล่ง A — https://a.com" in answer
    assert "แหล่ง B — https://b.com" in answer
    assert fake_client.call_count == 2
    assert receipts == []
    assert structural is False


# ---------- feat/advisor-route: ส่งตารางรายการเป็น context + เกณฑ์เปรียบเทียบ/ผลกระทบ/หน่วยขาย ----------
# ต้องเป็นกติกาทั่วไปที่ใช้ได้กับสินค้า "ทุกประเภท" ไม่ hardcode ตามสินค้าใดสินค้าหนึ่ง — ทดสอบด้วย 3 ประเภทสินค้า
# ที่ต่างกันมาก (กล้องวงจรปิด/ฮาร์ดดิสก์/สายสัญญาณ) ยืนยันว่า system prompt มีตารางรายการจริง + คำสั่งเรื่องเกณฑ์
# เปรียบเทียบและผลกระทบต่อรายการอื่นเหมือนกันทุกประเภท (ไม่ใช่ข้อความเฉพาะกิจของสินค้าใดสินค้าหนึ่ง)

_CAMERA_ITEMS = [("กล้องวงจรปิด 4MP รุ่น A", 8, "ตัว", 1200), ("Switch PoE 8 พอร์ต", 1, "เครื่อง", 2500)]
_HDD_ITEMS = [("ฮาร์ดดิสก์ 4TB สำหรับ NVR", 2, "ลูก", 2800), ("NVR 8 ช่อง", 1, "เครื่อง", 4500)]
_CABLE_ITEMS = [("สาย UTP Cat6 กล่อง 305 ม.", 3, "กล่อง", 1800), ("หัวต่อ RJ45", 100, "ชิ้น", 3)]


@pytest.fixture
def camera_doc(ns):
    raw = _build_quote_workbook_with_items(_CAMERA_ITEMS)
    label_map = ns["_extract_excel_labels"](raw)
    ns["_fake_db"].add_document(1, "quote.xlsx", raw, label_map, chat_id=1, user_id=1)
    return ns["_fake_db"].documents[1]


@pytest.fixture
def hdd_doc(ns):
    raw = _build_quote_workbook_with_items(_HDD_ITEMS)
    label_map = ns["_extract_excel_labels"](raw)
    ns["_fake_db"].add_document(1, "quote.xlsx", raw, label_map, chat_id=1, user_id=1)
    return ns["_fake_db"].documents[1]


@pytest.fixture
def cable_doc(ns):
    raw = _build_quote_workbook_with_items(_CABLE_ITEMS)
    label_map = ns["_extract_excel_labels"](raw)
    ns["_fake_db"].add_document(1, "quote.xlsx", raw, label_map, chat_id=1, user_id=1)
    return ns["_fake_db"].documents[1]


@pytest.mark.parametrize("doc_fixture_name, query, item_text", [
    ("camera_doc", "มีกล้องวงจรปิด 4MP รุ่นอื่นที่คุ้มกว่านี้ไหม", "กล้องวงจรปิด 4MP รุ่น A"),
    ("hdd_doc", "ฮาร์ดดิสก์สำหรับ NVR รุ่นนี้มีตัวเลือกอื่นไหม", "ฮาร์ดดิสก์ 4TB สำหรับ NVR"),
    ("cable_doc", "สาย UTP แบบนี้มีเกรดอื่นให้เลือกไหม", "สาย UTP Cat6 กล่อง 305 ม."),
])
def test_advisor_system_prompt_includes_quotation_table_for_any_product_category(
    ns, request, doc_fixture_name, query, item_text,
):
    """ไม่ว่าจะเป็นสินค้าประเภทไหน (กล้อง/HDD/สาย) advisor_answer ต้องส่งตารางรายการจริงของใบเสนอราคาที่เปิดอยู่
    เป็น context เสมอ (ดู _advisor_document_context) และ system prompt ต้องมีกติกาทั่วไป 3 ข้อ: เกณฑ์เปรียบเทียบ
    3-5 ข้อ, หัวข้อ "ผลกระทบต่อรายการอื่น", และหน่วยขายของรุ่นที่แนะนำ — ข้อความเดียวกันทุกประเภทสินค้า ไม่ใช่
    ข้อความเฉพาะกิจของสินค้าใดสินค้าหนึ่ง"""
    doc = request.getfixturevalue(doc_fixture_name)
    ns["_fake_agentic_loop"].set_response("แนะนำรุ่น A, B, C ครับ", [], False)
    ns["advisor_answer"](query, history=[], user_id=1, chat_id=1, active_docs=[doc])

    system_prompt = ns["_fake_agentic_loop"].last_call_args["system_prompt"]
    assert item_text in system_prompt
    assert "เกณฑ์เปรียบเทียบ 3-5 ข้อ" in system_prompt
    assert "ผลกระทบต่อรายการอื่น" in system_prompt
    assert "หน่วยขาย" in system_prompt
    # ห้ามบอกว่าแก้ไฟล์ไปแล้ว/กำลังจะแก้ — แจ้งผลกระทบเท่านั้น
    assert "ยังไม่มีการแก้ไฟล์ใดๆ เกิดขึ้นจริง" in system_prompt


def test_advisor_system_prompt_has_no_document_context_section_without_open_document(ns):
    ns["_fake_agentic_loop"].set_response("แนะนำรุ่น A, B, C ครับ", [], False)
    ns["advisor_answer"]("แนะนำ Rack case server 12U", history=[], user_id=1, chat_id=1, active_docs=[])

    system_prompt = ns["_fake_agentic_loop"].last_call_args["system_prompt"]
    assert "เกณฑ์เปรียบเทียบ 3-5 ข้อ" not in system_prompt
    assert "ผลกระทบต่อรายการอื่น" not in system_prompt


# ---------- feat/advisor-route: export บทสนทนาเป็นไฟล์ .txt ----------

def test_is_chat_export_command_matches_known_phrases(ns):
    for phrase in ["ขอบทสนทนา", "export แชท", "บันทึกแชทเป็นไฟล์", "ดาวน์โหลดบทสนทนา", "ช่วยขอบทสนทนานี้หน่อย"]:
        assert ns["_is_chat_export_command"](phrase) is True
    assert ns["_is_chat_export_command"]("ลบรายการที่ 3") is False


def test_build_chat_export_answer_is_pure_text_without_calling_llm(ns):
    """คำสั่ง export ต้องตอบลิงก์ดาวน์โหลดตรงๆ โดยไม่เรียก LLM เลยสักครั้ง (ดู ask_question — ดักก่อนเรียก AI
    ใดๆ เสมอ)"""
    answer = ns["_build_chat_export_answer"](42)
    assert "/api/chat/42/export.txt" in answer
    assert ns["_fake_client"].last_call_kwargs is None


def test_build_chat_export_text_contains_all_messages_in_chronological_order_with_thai_role_labels(ns):
    messages = [
        {"role": "user", "content": "สวัสดีครับ", "created_at": "2026-10-12T12:30:00"},
        {"role": "assistant", "content": "สวัสดีครับ มีอะไรให้ช่วยไหมครับ", "created_at": "2026-10-12T12:30:05"},
        {"role": "user", "content": "ขอบทสนทนานี้หน่อย", "created_at": "2026-10-12T12:35:00"},
    ]
    text = ns["_build_chat_export_text"]("แชททดสอบ", messages)

    assert "แชททดสอบ" in text
    # UTC naive -> เวลาไทย (+7) เสมอ (ดู _format_chat_export_timestamp)
    assert "[2026-10-12 19:30] ผู้ใช้: สวัสดีครับ" in text
    assert "[2026-10-12 19:30] ThaiMeSook: สวัสดีครับ มีอะไรให้ช่วยไหมครับ" in text
    assert "[2026-10-12 19:35] ผู้ใช้: ขอบทสนทนานี้หน่อย" in text

    # เรียงตามเวลาจริง (ลำดับเดียวกับที่ get_chat_messages คืนมา ไม่ใช่สลับเอง)
    first_idx = text.index("สวัสดีครับ")
    second_idx = text.index("มีอะไรให้ช่วยไหมครับ")
    third_idx = text.index("ขอบทสนทนานี้หน่อย")
    assert first_idx < second_idx < third_idx


def test_build_chat_export_text_handles_missing_or_unparsable_timestamp(ns):
    messages = [{"role": "user", "content": "ข้อความไม่มีเวลา", "created_at": None}]
    text = ns["_build_chat_export_text"]("แชททดสอบ", messages)
    assert "[-] ผู้ใช้: ข้อความไม่มีเวลา" in text
