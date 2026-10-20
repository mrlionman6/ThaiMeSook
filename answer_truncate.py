"""ตัดคำตอบที่ถูก Claude ตัดกลางคันเพราะชน max_tokens ให้เหลือถึงย่อหน้าสุดท้ายที่ดูจบครบ — เรียกจาก main.py:
rag_answer/rag_answer_stream เมื่อ stop_reason == "max_tokens" เท่านั้น"""
import re

_COMPLETE_ENDING_RE = re.compile(r"(ครับ|ค่ะ|คะ|[.?!ฯ])\s*$")


def truncate_to_last_complete_paragraph(text: str, source: str = "") -> str:
    """ตัดย่อหน้าสุดท้ายทิ้งถ้าดูค้างกลางคัน (ไม่จบด้วยเครื่องหมาย/คำลงท้ายปกติ) คืนค่าเดิมถ้าตัดแล้วว่างเปล่า"""
    stripped = text.rstrip()
    paragraphs = stripped.split("\n\n")
    if len(paragraphs) <= 1 or _COMPLETE_ENDING_RE.search(paragraphs[-1]):
        return text
    truncated = "\n\n".join(paragraphs[:-1]).rstrip()
    if not truncated:
        return text
    print(f"[AnswerTruncate] source={source} before_len={len(stripped)} after_len={len(truncated)}")
    return truncated
