"""ตรวจคำตอบที่ Claude เขียนเอง (ไม่ผ่าน tool จริง) ด้วย version ของเอกสารก่อน/หลังเทิร์น แทนการเดาจากคำพูด —
เรียกจาก main.py: rag_answer/advisor_answer หลังได้ raw_answer จาก run_agentic_tool_loop แล้วเสมอ

import จาก main.py แบบ deferred (หลังนิยาม _get_document_version_marker/_document_totals_lines/
_STRUCTURAL_NO_TOOL_FALLBACK_MESSAGE และ build_editable_fields) เหมือน document_fields.py ทุกประการ"""
import re

from main import (
    _get_document_version_marker, build_editable_fields, _document_totals_lines,
    _STRUCTURAL_NO_TOOL_FALLBACK_MESSAGE,
)

_MONEY_LINE_KEYWORDS = ("รวมเงิน", "รวมทั้งสิ้น", "ยอดรวม", "vat", "ภาษีมูลค่าเพิ่ม", "จำนวนเงิน")
_MONEY_TOKEN_RE = re.compile(r"\d[\d,]*(?:\.\d+)?\s*(?:บาท|%)")


def snapshot_document_versions(active_docs: list) -> dict:
    """คืน {document_id: version marker ปัจจุบัน} ของทุกไฟล์ใน active_docs — เรียก "ก่อน" run_agentic_tool_loop"""
    return {doc["id"]: _get_document_version_marker(doc) for doc in active_docs}


def _any_version_changed(active_docs: list, version_before: dict) -> bool:
    return any(_get_document_version_marker(doc) != version_before.get(doc["id"]) for doc in active_docs)


def _current_value_hint(active_docs: list, query: str) -> str:
    """เดาว่า query เอ่ยถึง field ไหน (ด้วย build_editable_fields) แล้วคืนค่าปัจจุบันจริงต่อท้าย คืน "" ถ้าเดาไม่ได้"""
    if len(active_docs) != 1:
        return ""
    try:
        fields = build_editable_fields(active_docs[0])
    except Exception:
        return ""
    normalized_query = query.replace(" ", "")
    for field in fields:
        label = str(field.get("label") or "")
        if len(label) >= 2 and label.replace(" ", "") in normalized_query:
            return f"\n\nค่าปัจจุบันของ '{label}': {field.get('current_value')}"
    return ""


def _strip_unverified_money_lines(raw_answer: str, active_docs: list) -> str:
    """ตัดบรรทัดที่อ้างยอดเงิน/VAT (มีคำใน _MONEY_LINE_KEYWORDS + ตัวเลข) ที่ไม่ตรงค่าจริงออก แล้วต่อยอดจริงแทน"""
    if len(active_docs) != 1:
        return raw_answer
    real_lines = _document_totals_lines(active_docs[0]["id"])
    real_values = [ln.split(":", 1)[-1].strip() for ln in real_lines if ln.startswith("-")]
    if not real_values:
        return raw_answer

    kept, dropped_any = [], False
    for line in raw_answer.splitlines():
        lowered = line.lower()
        has_keyword = any(kw in lowered for kw in _MONEY_LINE_KEYWORDS)
        has_money_token = bool(_MONEY_TOKEN_RE.search(line))
        if has_keyword and has_money_token and not any(v in line for v in real_values):
            dropped_any = True
            continue
        kept.append(line)
    if not dropped_any:
        return raw_answer
    return "\n".join(kept).rstrip() + "\n" + "\n".join(real_lines)


def truth_guard(
    raw_answer: str, *, query: str, active_docs: list, version_before: dict, edit_intent: bool, chat_id=None,
    strip_money_lines: bool = True,
) -> str:
    """edit_intent + มีไฟล์เปิด + ไม่มี version ไหนเปลี่ยนเลย -> แทน raw_answer ทั้งก้อนด้วยข้อความจากโค้ด มิฉะนั้น
    แค่ตัดตัวเลขยอดเงินที่ไม่ตรงค่าจริงออก (ถ้า strip_money_lines — ปิดได้สำหรับ advisor ที่พูดถึงยอดประมาณการ)"""
    if not active_docs:
        return raw_answer
    if edit_intent and not _any_version_changed(active_docs, version_before):
        print(f"[Guard] truth_check_replaced chat={chat_id}")
        return _STRUCTURAL_NO_TOOL_FALLBACK_MESSAGE + _current_value_hint(active_docs, query)
    if strip_money_lines:
        return _strip_unverified_money_lines(raw_answer, active_docs)
    return raw_answer
