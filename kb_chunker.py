"""หั่นเอกสารกฎหมายยาว (structure-aware) ให้เป็น chunk พร้อมเข้า Knowledge Base — แยกจาก legal_retrieval/legal_docs
ทั้งหมด ไม่แตะระบบค้น/ร่างเอกสารเลย

ลำดับ: (1) หา "มาตรา N" ทั้งหมด พร้อมเก็บบริบท "หมวด N"/"ส่วนที่ N" ที่อยู่ก่อนหน้า 1 มาตรา = 1 chunk (2) ไม่มี
มาตราเลย -> แยกย่อหน้า รวม/ตัดด้วย sentence tokenize ของ pythainlp ให้ไม่เกิน MAX_TOKENS (overlap 1 ประโยค)
(3) มาตราที่ยาวเกิน MAX_TOKENS ตัดย่อยตามประโยคแบบเดียวกัน ต่อท้าย section_ref ด้วย "(ต่อ N)"

import embed_model จาก main.py ตรงๆ (ไม่เคยถูก reassign หลัง startup — ต่างจาก bm25/knowledge_base_ids)"""
import re

from pythainlp.tokenize import sent_tokenize

from main import embed_model

MAX_TOKENS = 350

_THAI_ARABIC_DIGITS = "๐-๙0-9"
_SECTION_KEYWORD_RE = re.compile(rf"^[ \t]*((?:หมวด|ส่วนที่)\s*[{_THAI_ARABIC_DIGITS}]+)[ \t]*(.*)$")
_ARTICLE_LINE_RE = re.compile(rf"^[ \t]*มาตรา[ \t]*([{_THAI_ARABIC_DIGITS}]+(?:/[{_THAI_ARABIC_DIGITS}]+)?)[ \t]*(.*)$")


def _count_tokens(text: str) -> int:
    """นับ token ด้วย tokenizer จริงของ embed_model (ไม่ใช่นับตัวอักษร) — ใช้ชี้วัดว่า chunk เกิน MAX_TOKENS หรือไม่"""
    return len(embed_model.tokenizer.tokenize(text)) if text else 0


def _sentences(text: str) -> list[str]:
    return [s.strip() for s in sent_tokenize(text, engine="crfcut") if s.strip()]


def _pack_sentences(sentences: list[str], max_tokens: int, overlap_sentences: int = 1) -> list[str]:
    """รวมประโยคเป็น chunk ทีละตัวแบบ greedy ไม่ให้เกิน max_tokens เมื่อล้นให้ตัดเป็น chunk ใหม่ โดยพก
    overlap_sentences ประโยคสุดท้ายของ chunk ก่อนหน้ามาต่อเนื่องบริบท"""
    if not sentences:
        return []
    chunks = []
    current: list[str] = []
    for sent in sentences:
        if current and _count_tokens(" ".join(current + [sent])) > max_tokens:
            chunks.append(" ".join(current))
            current = current[-overlap_sentences:] if overlap_sentences > 0 else []
        current.append(sent)
    if current:
        chunks.append(" ".join(current))
    return chunks


def _parse_sections(lines: list[str]) -> tuple[dict[int, str], set[int]]:
    """เดินทีละบรรทัดหาหัว "หมวด N"/"ส่วนที่ N" — ชื่อหมวดอยู่ท้ายบรรทัดเดียวกันหรือบรรทัดถัดไปก็ได้ (รูปแบบ
    เอกสารกฎหมายไทยทั่วไปที่เจอบ่อยสุดคือขึ้นบรรทัดใหม่ เช่น "หมวด ๑" แล้วตามด้วย "บททั่วไป" คนละบรรทัด)
    คืน (section_at_line: บรรทัด -> ชื่อหมวดที่มีผล ณ บรรทัดนั้น, consumed: เซ็ตบรรทัดที่เป็นส่วนหัวล้วนๆ ไม่ใช่เนื้อหา)"""
    section_at_line: dict[int, str] = {}
    consumed: set[int] = set()
    current_section = None
    idx = 0
    while idx < len(lines):
        m = _SECTION_KEYWORD_RE.match(lines[idx])
        if not m:
            section_at_line[idx] = current_section
            idx += 1
            continue
        title = m.group(1).strip()
        same_line_title = m.group(2).strip()
        consumed.add(idx)
        j = idx + 1
        if same_line_title:
            title = f"{title} {same_line_title}"
        else:
            while j < len(lines) and lines[j].strip() == "":
                consumed.add(j)
                j += 1
            if j < len(lines):
                next_line = lines[j].strip()
                if next_line and not next_line.startswith("มาตรา") and not _SECTION_KEYWORD_RE.match(lines[j]):
                    title = f"{title} {next_line}"
                    consumed.add(j)
                    j += 1
        current_section = title
        idx = j
    return section_at_line, consumed


def _parse_articles(text: str) -> list[dict]:
    """คืน [{article_no, section, body}] ตามลำดับที่เจอในเอกสาร — ว่างถ้าไม่มี 'มาตรา' เลยสักตัว"""
    lines = text.split("\n")
    section_at_line, consumed_section_lines = _parse_sections(lines)

    article_lines = []
    for idx, line in enumerate(lines):
        m = _ARTICLE_LINE_RE.match(line)
        if m:
            article_lines.append((idx, m.group(1), m.group(2)))

    articles = []
    for k, (idx, article_no, rest_of_line) in enumerate(article_lines):
        end_idx = article_lines[k + 1][0] if k + 1 < len(article_lines) else len(lines)
        body_lines = [rest_of_line] + [
            lines[i] for i in range(idx + 1, end_idx) if i not in consumed_section_lines
        ]
        body = re.sub(r"\n[ \t]*\n+", "\n", "\n".join(body_lines)).strip()
        articles.append({"article_no": article_no, "section": section_at_line.get(idx), "body": body})
    return articles


def _build_header(source_title: str, section: str | None, section_ref: str | None) -> str:
    parts = [source_title] + [p for p in (section, section_ref) if p]
    return "[" + " > ".join(parts) + "]"


def _build_chunk(source_title: str, section: str | None, section_ref: str | None, body: str) -> dict:
    """token_count นับเฉพาะเนื้อ body (ไม่รวม header) — ตัวเลขเดียวกับที่ใช้ตัดสินใจหั่น/รวม chunk ตอนก่อนหน้านี้"""
    header = _build_header(source_title, section, section_ref)
    content = f"{header}\n{body}".strip()
    return {"content": content, "header": header, "section_ref": section_ref, "token_count": _count_tokens(body)}


def _chunk_fallback_paragraphs(text: str, source_title: str) -> list[dict]:
    paragraphs = [p.strip() for p in re.split(r"\n[ \t]*\n", text) if p.strip()]
    all_sentences = []
    for para in paragraphs:
        all_sentences.extend(_sentences(para))
    groups = _pack_sentences(all_sentences, MAX_TOKENS, overlap_sentences=1)
    return [_build_chunk(source_title, None, None, g) for g in groups]


def chunk_document(text: str, source_title: str) -> list[dict]:
    """คืน [{content, header, section_ref, token_count}] — content มี header ขึ้นต้นเสมอ (ใช้บันทึกตรงๆ ผ่าน
    add_knowledge_chunk ได้เลย)"""
    articles = _parse_articles(text)
    if not articles:
        return _chunk_fallback_paragraphs(text, source_title)

    chunks = []
    for art in articles:
        section_ref = f"มาตรา {art['article_no']}"
        if _count_tokens(art["body"]) <= MAX_TOKENS:
            chunks.append(_build_chunk(source_title, art["section"], section_ref, art["body"]))
            continue
        parts = _pack_sentences(_sentences(art["body"]), MAX_TOKENS, overlap_sentences=1)
        for i, part in enumerate(parts):
            ref = section_ref if i == 0 else f"{section_ref} (ต่อ {i})"
            chunks.append(_build_chunk(source_title, art["section"], ref, part))
    return chunks
