"""คำสั่งแก้ไฟล์เดียวที่มีหลายคำสั่งย่อยผสมกัน (field_edits/rate_changes/edit_instruction) ต้องได้ใบยืนยัน ✓/✗
ครบทุกส่วน ไม่หายเงียบ — ไฟล์แยกจาก main.py เพราะต้องใช้ apply_change_rate/_resolve_direct_structural_dispatch/
_match_and_apply_excel_edit/rag_answer ที่นิยามอยู่ "หลัง" จุด import นี้ในไฟล์นั้นเอง (ดูคอมเมนต์ตรงจุด import)"""
from main import (
    apply_change_rate, _resolve_direct_structural_dispatch, _match_and_apply_excel_edit, rag_answer,
    get_chat_messages, _document_totals_lines, _UNDO_HINT_LINE,
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


def apply_compound_edit(route: dict, target_doc: dict, query: str, user_id: int, chat_id: int) -> dict:
    """apply field_edits -> rate_changes (ผ่าน apply_change_rate เดิม) -> table edits อื่น (edit_instruction) ตาม
    ลำดับ รวมใบยืนยันทุกส่วนเป็นข้อความเดียว ต่อท้ายด้วยยอดล่าสุด+undo hint ครั้งเดียวรวมทั้งหมด จำนวน sub_commands
    ที่ router แยกมามากกว่าจำนวนที่ apply/ปฏิเสธไปแล้ว -> ต่อท้าย "✗ ยังไม่ได้ทำ: ..." กันหายเงียบ"""
    document_id = target_doc["id"]
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
            if not message.startswith(("✓", "✗")):
                message = ("✓ " if result.get("matched") else "✗ ") + message
            receipt_sections.append(_strip_totals_and_undo_tail(message))
        accounted += len(rate_changes)

    edit_instruction = route.get("edit_instruction")
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
        return {"document_id": document_id, "answer": answer}

    sub_commands = route.get("sub_commands")
    if isinstance(sub_commands, list):
        for sub_command in sub_commands[accounted:]:
            if isinstance(sub_command, str) and sub_command.strip():
                receipt_sections.append(f"✗ ยังไม่ได้ทำ: {sub_command.strip()}")

    message_lines = ["\n\n".join(receipt_sections)] + _document_totals_lines(document_id) + ["", _UNDO_HINT_LINE]
    return {"document_id": document_id, "answer": "\n".join(message_lines)}
