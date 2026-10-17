"""feat/edit-any-field: "ทุกช่องในเอกสารแก้ได้ด้วยกลไกเดียว" ไม่ hardcode ชื่อป้าย ใช้ได้กับทุกแม่แบบ reuse
get_working_copy/create_new_version/_get_single_item_table ฯลฯ ของ main.py — import ชื่อเหล่านี้จาก main ตรงๆ
ที่หัวไฟล์ (ไม่ใช่ deferred import) เพราะ main.py เองก็ import build_editable_fields/apply_field_edits กลับไป
ด้วย — main.py ตั้งใจวาง `from document_fields import ...` ไว้ "หลัง" นิยามของทุกชื่อที่ไฟล์นี้ต้องใช้แล้วเท่านั้น
(ดูคอมเมนต์ตรงนั้นในตัวมัน) กัน circular import"""
import io

import openpyxl
from openpyxl.styles.numbers import is_date_format

from db import get_editable_document, log_document_edit, link_edit_log_to_version
from main import (
    get_working_copy, create_new_version, _get_single_item_table, _merged_cell_top_left,
    _coerce_value_for_cell, _document_totals_lines, _UNDO_HINT_LINE, _TABLE_FIELD_THAI_NAMES,
    _RATE_FORMULA_PATTERNS, set_chat_document_focus, _format_display_value,
)

# คอลัมน์ตารางรายการที่แก้ตรงได้ (ไม่รวม "amount" — เป็นสูตรเสมอ)
_TABLE_EDITABLE_FIELDS = ("no", "description", "qty", "unit", "unit_price")

_RATE_FIELD_REASON = "เป็นอัตราที่คำนวณในสูตร ใช้คำสั่งแก้อัตราตรงๆ เช่น \"เปลี่ยน <ชื่อป้าย> เป็น 10%\" แทน"
_FORMULA_FIELD_REASON = "คำนวณอัตโนมัติจากสูตรในไฟล์"
_AMOUNT_FIELD_REASON = "จำนวน × ราคา/หน่วย"


def _guess_value_type(value, number_format) -> str:
    """เดาชนิดข้อมูลของช่องจากค่าปัจจุบัน + number_format (text/number/date/percent)"""
    if number_format and is_date_format(number_format):
        return "date"
    if number_format and "%" in number_format:
        return "percent"
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return "number"
    return "text"


def _is_rate_formula_text(formula_text: str) -> bool:
    return any(pattern.search(formula_text) for pattern, _kind in _RATE_FORMULA_PATTERNS)


_TEXT_FIELD_LABEL_MAX_CHARS = 20


def _short_text_field_label(value: str, cell: str) -> str:
    """ช่องข้อความลอย (kind="text") ใช้ค่าในเซลล์เองเป็น label — ถ้ายาวเกินไป/มีตัวเลข (เช่น "โทร : 02-888-8888
    โทรสาร : ") ใช้คำนำหน้าแรกก่อน ":" แทน (เช่น "โทร") ไม่มี ":" ให้ใช้ "ข้อความในช่อง <cell>" แทนทั้งค่า"""
    if len(value) <= _TEXT_FIELD_LABEL_MAX_CHARS and not any(ch.isdigit() for ch in value):
        return value
    prefix = value.split(":", 1)[0].strip()
    if prefix and len(prefix) <= _TEXT_FIELD_LABEL_MAX_CHARS:
        return prefix
    return f"ข้อความในช่อง {cell}"


def build_editable_fields(document: dict) -> list[dict]:
    """สแกนเอกสารทั้งไฟล์ (label ทุกตัว + ตารางรายการ + ข้อความเดี่ยวไม่มีป้าย เช่น ชื่อ/ที่อยู่บริษัทเรา/
    หมายเหตุ/ชื่อผู้เซ็น) คืนรายการช่องทั้งหมด {field_id, kind, sheet, cell, label, current_value, value_type,
    editable, reason} kind: "label"/"table_cell"/"text" — ใช้กลไกเดียวกันทุกแม่แบบ ไม่ hardcode ชื่อป้าย คืนรายการ
    เต็มไม่ตัดขนาด (ตัด/จำกัดจำนวนเป็นหน้าที่ของผู้แสดงผล เช่น _build_router_document_context — apply_field_edits
    ต้องเห็นช่องทั้งหมดจริงเพื่อตรวจสอบ)"""
    base_bytes, label_map = get_working_copy(document)
    wb_raw = openpyxl.load_workbook(io.BytesIO(base_bytes), data_only=False)

    fields: list[dict] = []
    occupied_positions = set()  # (sheet, row, col) ที่เป็น label/value cell ของ label_map ไปแล้ว กันซ้ำกับ text field

    # ---- label fields (label_map เดิมของ main.py ตรงๆ) ----
    for label, info in label_map.items():
        sheet, row, col = info["sheet"], info["row"], info["col"]
        occupied_positions.add((sheet, row, col))
        if info.get("label_row") is not None and info.get("label_col") is not None:
            occupied_positions.add((sheet, info["label_row"], info["label_col"]))

        raw_value = wb_raw[sheet].cell(row=row, column=col).value
        is_formula = isinstance(raw_value, str) and raw_value.startswith("=")
        if is_formula:
            editable = False
            reason = _RATE_FIELD_REASON if _is_rate_formula_text(raw_value) else _FORMULA_FIELD_REASON
        else:
            editable, reason = True, None

        fields.append({
            "field_id": f"label|{sheet}|{row}|{col}",
            "kind": "label",
            "sheet": sheet,
            "row": row,
            "col": col,
            "cell": f"{openpyxl.utils.get_column_letter(col)}{row}",
            "label": label,
            "current_value": info["current_value"],
            "value_type": _guess_value_type(info["current_value"], info.get("number_format")),
            "number_format": info.get("number_format"),
            "editable": editable,
            "reason": reason,
        })

    # ---- table_cell fields (ตารางรายการเดียวที่มีอยู่ — ดู _get_single_item_table) ----
    table = _get_single_item_table(base_bytes)
    excluded_rows = set()
    if table is not None:
        excluded_rows = set(range(table["header_row"], table["last_row"] + 1))
        filled_rows = [r for r in table["rows"] if not r["is_empty"]]
        for item_no, row_info in enumerate(filled_rows, start=1):
            row = row_info["row"]
            for field in _TABLE_EDITABLE_FIELDS:
                col = table["columns"].get(field)
                if col is None:
                    continue
                occupied_positions.add((table["sheet"], row, col))
                raw_value = wb_raw[table["sheet"]].cell(row=row, column=col).value
                is_formula = isinstance(raw_value, str) and raw_value.startswith("=")
                fields.append({
                    "field_id": f"table|{table['sheet']}|{row}|{field}",
                    "kind": "table_cell",
                    "sheet": table["sheet"],
                    "row": row,
                    "col": col,
                    "cell": f"{openpyxl.utils.get_column_letter(col)}{row}",
                    "label": f"รายการที่ {item_no} - {_TABLE_FIELD_THAI_NAMES.get(field, field)}",
                    "current_value": row_info["values"].get(field),
                    "value_type": "number" if field in ("qty", "unit_price") else "text",
                    "number_format": None,
                    "editable": not is_formula,
                    "reason": _FORMULA_FIELD_REASON if is_formula else None,
                })
            amount_col = table["columns"].get("amount")
            if amount_col is not None:
                occupied_positions.add((table["sheet"], row, amount_col))
                fields.append({
                    "field_id": f"table|{table['sheet']}|{row}|amount",
                    "kind": "table_cell",
                    "sheet": table["sheet"],
                    "row": row,
                    "col": amount_col,
                    "cell": f"{openpyxl.utils.get_column_letter(amount_col)}{row}",
                    "label": f"รายการที่ {item_no} - {_TABLE_FIELD_THAI_NAMES.get('amount', 'จำนวนเงิน')}",
                    "current_value": row_info["values"].get("amount"),
                    "value_type": "number",
                    "number_format": None,
                    "editable": False,
                    "reason": _AMOUNT_FIELD_REASON,
                })

    # ---- text fields: แถวที่มีเซลล์ไม่ว่างเซลล์เดียว ไม่ใช่สูตร ไม่ซ้ำกับ label/ตาราง (เช่น ชื่อ/ที่อยู่
    # บริษัทเรา, หมายเหตุ, ชื่อผู้เซ็น — มักวางเป็นข้อความลอยไม่มีป้ายกำกับในเอกสารจริง) ----
    for sheet_name in wb_raw.sheetnames:
        ws = wb_raw[sheet_name]
        sheet_excluded_rows = excluded_rows if (table is not None and table["sheet"] == sheet_name) else set()
        for row in ws.iter_rows():
            if row[0].row in sheet_excluded_rows:
                continue
            non_empty = [c for c in row if c.value is not None and str(c.value).strip() != ""]
            if len(non_empty) != 1:
                continue
            cell = non_empty[0]
            if (sheet_name, cell.row, cell.column) in occupied_positions:
                continue
            value = cell.value
            if not isinstance(value, str) or not value.strip() or value.strip().startswith("="):
                continue
            cell_ref = f"{openpyxl.utils.get_column_letter(cell.column)}{cell.row}"
            fields.append({
                "field_id": f"text|{sheet_name}|{cell.row}|{cell.column}",
                "kind": "text",
                "sheet": sheet_name,
                "row": cell.row,
                "col": cell.column,
                "cell": cell_ref,
                "label": _short_text_field_label(value.strip(), cell_ref),
                "current_value": value,
                "value_type": "text",
                "number_format": None,
                "editable": True,
                "reason": None,
            })

    return fields


def apply_field_edits(document_id: int, user_id: int, field_edits: list[dict]) -> dict:
    """แก้ช่องใดก็ได้ในเอกสาร (label/table_cell/text) ด้วยกลไกเดียว — field_edits: [{field_id, operation, new_
    value, find_text}] operation: "set" (ตั้งค่าใหม่ทั้งหมด) / "replace_text" (แทนที่บางส่วนของค่าปัจจุบัน —
    ต้องมี find_text ตรงกับค่าปัจจุบันจริง) / "clear" (ล้างเป็นค่าว่าง) เขียนทุกช่องที่ถูกต้องผ่าน create_new_
    version ครั้งเดียว (ย้อนกลับได้ด้วย undo เดิม) ช่องที่ field_id ไม่มีจริง/editable=False/find_text ไม่เจอ ->
    ข้าม ไม่แก้ พร้อมบรรทัด "[[FAIL]] ..." อธิบาย (marker ภายใน — ดู main._render_receipt_markers) ไม่มีช่องไหน
    แก้ได้จริงเลย -> matched=False ไม่สร้างเวอร์ชันเปล่าๆ"""
    doc = get_editable_document(document_id, user_id)
    if doc is None:
        return {"matched": False, "message": "ไม่พบเอกสารนี้ อาจหมดอายุหรือถูกลบไปแล้ว"}
    if not field_edits:
        return {"matched": False, "message": "ไม่ได้ระบุช่องที่จะแก้"}

    fields_by_id = {f["field_id"]: f for f in build_editable_fields(doc)}

    to_apply = []
    receipt_lines = []
    for edit in field_edits:
        field_id = edit.get("field_id")
        field = fields_by_id.get(field_id)
        if field is None:
            receipt_lines.append("[[FAIL]] ไม่พบช่องนี้ในเอกสารแล้ว (อาจถูกแก้โครงสร้างไปก่อนหน้า)")
            continue
        if not field["editable"]:
            receipt_lines.append(f"[[FAIL]] '{field['label']}' เป็นช่องคำนวณอัตโนมัติจาก{field['reason']} แก้ตรงไม่ได้")
            continue

        current_text = "" if field["current_value"] is None else str(field["current_value"])
        operation = edit.get("operation")

        if operation == "clear":
            new_text = ""
        elif operation == "replace_text":
            find_text = edit.get("find_text")
            if not find_text or find_text not in current_text:
                receipt_lines.append(
                    f"[[FAIL]] ไม่พบข้อความ \"{find_text}\" ในค่าปัจจุบันของ '{field['label']}' ({current_text}) แก้ไม่ได้"
                )
                continue
            new_value = edit.get("new_value")
            new_text = current_text.replace(find_text, "" if new_value is None else str(new_value))
        elif operation == "set":
            new_value = edit.get("new_value")
            new_text = "" if new_value is None else str(new_value)
        else:
            receipt_lines.append(f"[[FAIL]] ไม่รู้จักวิธีแก้ '{operation}' สำหรับ '{field['label']}'")
            continue

        if new_text == current_text:
            receipt_lines.append(f"- '{field['label']}' เป็นค่านี้อยู่แล้วครับ")
            continue

        to_apply.append({"field": field, "old_text": current_text, "new_text": new_text})

    if not to_apply:
        message = "\n".join(receipt_lines) if receipt_lines else "ไม่มีช่องไหนที่แก้ได้จริง"
        return {"matched": False, "message": message}

    def _mutate(wb):
        for item in to_apply:
            field = item["field"]
            ws = wb[field["sheet"]]
            ok, value, _err = _coerce_value_for_cell(item["new_text"], field.get("number_format"))
            cell = _merged_cell_top_left(ws, ws.cell(row=field["row"], column=field["col"]))
            cell.value = value if ok else item["new_text"]

    version_result = create_new_version(
        document_id, _mutate, change_summary="แก้ช่องเอกสาร: " + ", ".join(i["field"]["label"] for i in to_apply),
    )
    if doc.get("chat_id"):
        set_chat_document_focus(doc["chat_id"], document_id)

    success_lines = []
    exclude_labels = set()
    for item in to_apply:
        field = item["field"]
        # แสดงวันที่แบบ วัน/เดือน/ปี แทน ISO string ดิบ (เช่น "2569-10-20T00:00:00" -> "20/10/2569")
        old_display = _format_display_value(item["old_text"], field.get("number_format"))
        new_display = _format_display_value(item["new_text"], field.get("number_format"))
        success_lines.append(f"[[OK]] แก้ '{field['label']}' จาก \"{old_display}\" เป็น \"{new_display}\"")
        exclude_labels.add(field["label"])
        if field["kind"] == "label":
            edit_log_id = log_document_edit(document_id, field["label"], item["old_text"], item["new_text"])
            link_edit_log_to_version(edit_log_id, version_result["version_id"])

    message_lines = success_lines + receipt_lines
    message_lines += _document_totals_lines(document_id, exclude_labels=exclude_labels)
    message_lines += ["", _UNDO_HINT_LINE]
    return {"matched": True, "message": "\n".join(message_lines)}
