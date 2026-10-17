"""คำสั่งแก้ไฟล์เดียวที่มีหลายคำสั่งย่อยผสมกันต้องได้ใบยืนยันครบทุกส่วน ไม่หายเงียบ และ undo ครั้งเดียวต้องย้อน
ทั้งคำสั่งถ้าสร้างหลายเวอร์ชัน — ไฟล์แยกจาก main.py เพราะต้องใช้ชื่อที่นิยามอยู่ "หลัง" จุด import นี้ในไฟล์นั้นเอง"""
from main import (
    apply_change_rate, _resolve_direct_structural_dispatch, _match_and_apply_excel_edit, rag_answer,
    get_chat_messages, _document_totals_lines, _UNDO_HINT_LINE, get_latest_document_version,
    update_document_version_change_summary, _STRUCTURAL_CLAUSE_SPLIT_RE, _STRUCTURAL_DIRECT_RATE_RE,
    _render_receipt_markers,
)
from document_fields import apply_field_edits

_TOTALS_MARKER = "\n\nยอดล่าสุด:"


def _strip_totals_and_undo_tail(message: str) -> str:
    """ตัดส่วนท้าย "ยอดล่าสุด:"/undo hint ออก — ใช้รวมหลายใบยืนยันเป็นข้อความเดียวแล้วค่อยต่อท้ายครั้งเดียวรวม"""
    idx = message.find(_TOTALS_MARKER)
    if idx != -1:
        return message[:idx]
    undo_idx = message.find("\n\n" + _UNDO_HINT_LINE)
    if undo_idx != -1:
        return message[:undo_idx]
    return message


def _strip_rate_clauses(edit_instruction: str) -> str:
    """ตัดประโยคย่อยที่ตรง _STRUCTURAL_DIRECT_RATE_RE (คำสั่งเปลี่ยนอัตรา) ออก — กันเปลี่ยนอัตราซ้ำสองรอบ (rate_
    changes apply ไปแล้วรอบหนึ่ง edit_instruction ไม่ควรพูดเรื่องเดียวกันซ้ำอีก)"""
    clauses = _STRUCTURAL_CLAUSE_SPLIT_RE.split(edit_instruction)
    remaining = [c for c in clauses if c.strip() and not _STRUCTURAL_DIRECT_RATE_RE.search(c)]
    return " และ ".join(remaining)


def _version_no(document_id: int) -> int:
    latest = get_latest_document_version(document_id)
    return latest["version_no"] if latest else 0


def apply_compound_edit(route: dict, target_doc: dict, query: str, user_id: int, chat_id: int) -> dict:
    """apply field_edits -> rate_changes -> table edit_instruction ตามลำดับ รวมใบยืนยันเป็นข้อความเดียว sub_
    commands ที่ไม่มีอะไรทำ -> "ยังไม่ได้ทำ: ..." สร้าง >1 เวอร์ชัน -> ฝัง marker ให้ undo ย้อนทีเดียวทั้งกลุ่ม"""
    document_id = target_doc["id"]
    version_before = _version_no(document_id)
    receipt_sections: list[str] = []
    accounted = 0

    field_edits = route.get("field_edits")
    if isinstance(field_edits, list) and field_edits:
        result = apply_field_edits(document_id, user_id, field_edits)
        receipt_sections.append(_strip_totals_and_undo_tail(result["message"]))
        accounted += len(field_edits)

    rate_changes = route.get("rate_changes")
    if isinstance(rate_changes, list) and rate_changes:
        for item in rate_changes:
            if not isinstance(item, dict):
                continue
            label = item.get("label")
            new_percent = item.get("new_percent")
            if not label or new_percent is None:
                continue
            result = apply_change_rate(document_id, user_id, label, new_percent)
            message = result["message"]
            if not message.startswith(("[[OK]]", "[[FAIL]]")):
                message = ("[[OK]] " if result.get("matched") else "[[FAIL]] ") + message
            receipt_sections.append(_strip_totals_and_undo_tail(message))
        accounted += len(rate_changes)

    edit_instruction = route.get("edit_instruction")
    if isinstance(edit_instruction, str) and edit_instruction.strip():
        if rate_changes:
            edit_instruction = _strip_rate_clauses(edit_instruction)
    if isinstance(edit_instruction, str) and edit_instruction.strip():
        direct = _resolve_direct_structural_dispatch(edit_instruction, [target_doc], chat_id, user_id)
        if direct is not None:
            table_answer = direct["answer"]
        else:
            matcher_result = _match_and_apply_excel_edit(document_id, user_id, edit_instruction)
            if matcher_result.get("matched"):
                table_answer = matcher_result["message"]
            else:
                history = get_chat_messages(chat_id) if chat_id else []
                table_answer, _sources = rag_answer(
                    edit_instruction, history=history, image_data=None, user_id=user_id, chat_id=chat_id,
                    structural_route=True, skip_legal_context=False,
                )
        receipt_sections.append(_strip_totals_and_undo_tail(table_answer))
        accounted += 1

    if not receipt_sections:
        # router ตัดสิน edit_document แต่ไม่ได้ส่ง field_edits/rate_changes/edit_instruction มาเลยสักอัน -> ทางเดิม
        history = get_chat_messages(chat_id) if chat_id else []
        answer, _sources = rag_answer(
            query, history=history, image_data=None, user_id=user_id, chat_id=chat_id,
            structural_route=True, skip_legal_context=False,
        )
        _mark_compound_group(document_id, version_before)
        return {"document_id": document_id, "answer": answer}

    sub_commands = route.get("sub_commands")
    if isinstance(sub_commands, list):
        for sub_command in sub_commands[accounted:]:
            if isinstance(sub_command, str) and sub_command.strip():
                receipt_sections.append(f"[[FAIL]] ยังไม่ได้ทำ: {sub_command.strip()}")

    _mark_compound_group(document_id, version_before)
    message_lines = ["\n\n".join(receipt_sections)] + _document_totals_lines(document_id) + ["", _UNDO_HINT_LINE]
    # แปลง marker เป็นข้อความธรรมดาก่อนคืนผู้ใช้ — จุดเดียวสุดท้ายของเส้นทาง compound edit นี้
    return {"document_id": document_id, "answer": _render_receipt_markers("\n".join(message_lines))}


def _mark_compound_group(document_id: int, version_before: int) -> None:
    """เทิร์นนี้สร้างมากกว่า 1 เวอร์ชัน -> ฝัง "__compound_group__:N" ใน change_summary ของเวอร์ชันสุดท้าย (ไม่
    แสดงผลที่ไหนเลย — ดู create_new_version) ให้ undo_last_edit() อ่านแล้วย้อนทีเดียวทั้ง N เวอร์ชัน"""
    latest = get_latest_document_version(document_id)
    if latest is None:
        return
    versions_created = latest["version_no"] - version_before
    if versions_created > 1:
        update_document_version_change_summary(latest["id"], f"__compound_group__:{versions_created}")
