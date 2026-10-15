"""ตรวจคำตอบที่ Claude เขียนเอง (ไม่ผ่าน tool จริง) ด้วยผลจริง (version ของเอกสารก่อน/หลังเทิร์น) แทนการเดาจาก
คำพูด (blacklist คำอย่าง _looks_like_fabricated_edit_claim เดิม) — เรียกจาก main.py: rag_answer/advisor_answer
เท่านั้น หลังได้ raw_answer จาก run_agentic_tool_loop แล้วเสมอ (ก่อนต่อท้าย action_receipts)

หลักการ: ผู้ใช้สั่งแก้ไฟล์ (edit_intent) ขณะมีไฟล์เปิดอยู่ แต่ version ของทุกไฟล์ที่เปิดอยู่ไม่เปลี่ยนเลยหลังเทิร์น
นี้ (ไม่มี tool ไหนแก้ไฟล์จริงสำเร็จ) -> ไม่เชื่อข้อความที่ Claude เขียนเองเลยไม่ว่าจะฟังดูน่าเชื่อแค่ไหน แทนด้วย
ข้อความจากโค้ดเสมอ — ครอบคลุมทุกเส้นทางที่จบที่ raw_answer เป็นข้อความอิสระของ Claude (router edit_document
fallback, router advisor/legal/library ที่ไม่มี field_edits ตรง, legacy cascade, single/multi-doc classifier
"unrelated" fallthrough ฯลฯ) เพราะ rag_answer/advisor_answer เป็นจุดที่ raw_answer ทุกเส้นทางไหลผ่านจริง

import จาก main.py แบบ deferred (หลังนิยาม _get_document_version_marker/_document_totals_lines และหลัง import
build_editable_fields จาก document_fields.py แล้ว) เหมือน document_fields.py ทุกประการ — ดู main.py ตรงจุดที่
import ไฟล์นี้กลับ (หลัง "from document_fields import build_editable_fields, apply_field_edits")"""
import re

from main import (
    _get_document_version_marker, build_editable_fields, _document_totals_lines,
    _STRUCTURAL_NO_TOOL_FALLBACK_MESSAGE,
)

_MONEY_LINE_KEYWORDS = ("รวมเงิน", "รวมทั้งสิ้น", "ยอดรวม", "vat", "ภาษีมูลค่าเพิ่ม", "จำนวนเงิน")
_MONEY_TOKEN_RE = re.compile(r"\d[\d,]*(?:\.\d+)?\s*(?:บาท|%)")


def snapshot_document_versions(active_docs: list) -> dict:
    """คืน {document_id: version marker ปัจจุบัน} ของทุกไฟล์ใน active_docs — เรียก "ก่อน" run_agentic_tool_loop
    เสมอ (ดู main.py: rag_answer/advisor_answer) ใช้เทียบกับสถานะหลังเทิร์นใน truth_guard ด้านล่าง"""
    return {doc["id"]: _get_document_version_marker(doc) for doc in active_docs}


def _any_version_changed(active_docs: list, version_before: dict) -> bool:
    return any(_get_document_version_marker(doc) != version_before.get(doc["id"]) for doc in active_docs)


def _current_value_hint(active_docs: list, query: str) -> str:
    """เดาว่า query เอ่ยถึง field ไหนในไฟล์ที่เปิดอยู่ (ด้วย build_editable_fields — ดู document_fields.py) แล้วคืน
    ค่าปัจจุบันจริงของ field นั้นต่อท้ายข้อความ guard ถ้าเดาได้ คืน "" ถ้าเดาไม่ได้/มีมากกว่า 1 ไฟล์เปิดอยู่"""
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
    """ตัดบรรทัดที่อ้างยอดเงิน/VAT/รวมทั้งสิ้น (มีคำใน _MONEY_LINE_KEYWORDS + ตัวเลขแบบเงิน) ที่ไม่ตรงกับค่าจริง
    ปัจจุบันของไฟล์เลยสักตัวเลข (ดู _document_totals_lines) ออก แล้วต่อท้ายด้วยยอดจริงแทน — ใช้เฉพาะมีไฟล์เดียว
    เปิดอยู่ชัดเจนเท่านั้น (หลายไฟล์ -> ไม่รู้จะเทียบกับไฟล์ไหน ปล่อยผ่าน)"""
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


def truth_guard(raw_answer: str, *, query: str, active_docs: list, version_before: dict, edit_intent: bool, chat_id=None) -> str:
    """เรียกหลังได้ raw_answer จาก run_agentic_tool_loop แล้วเสมอ (เฉพาะตอนไม่มี action_receipts ในเทิร์นนี้ —
    มี receipt แปลว่ามี tool แก้ไฟล์จริงสำเร็จแล้ว ไม่ต้องตรวจซ้ำ) edit_intent=True (ผู้ใช้สั่งแก้ไฟล์จริงๆ — ดู
    main.py: _query_has_edit_verb/structural_route) + มีไฟล์เปิดอยู่ + ไม่มี version ไหนเปลี่ยนเลย -> ไม่เชื่อ
    raw_answer เลยทั้งก้อน แทนด้วยข้อความจากโค้ด ไม่ว่าในกรณีอื่น (ไม่ใช่ edit_intent หรือ version เปลี่ยนจริง) ->
    แค่ตัดตัวเลขยอดเงินที่ไม่ตรงค่าจริงออก (ดู _strip_unverified_money_lines) เผื่อ Claude แต่งตัวเลขปนในคำตอบปกติ"""
    if not active_docs:
        return raw_answer
    if edit_intent and not _any_version_changed(active_docs, version_before):
        print(f"[Guard] truth_check_replaced chat={chat_id}")
        return _STRUCTURAL_NO_TOOL_FALLBACK_MESSAGE + _current_value_hint(active_docs, query)
    return _strip_unverified_money_lines(raw_answer, active_docs)
