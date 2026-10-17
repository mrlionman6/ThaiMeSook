"""หั่นเอกสารกฎหมายยาว (structure-aware) ให้เป็น chunk พร้อมเข้า Knowledge Base — แยกจาก legal_retrieval/legal_docs
ทั้งหมด ไม่แตะระบบค้น/ร่างเอกสารเลย

ลำดับ: (0) ข้อความก่อน "มาตรา" แรก (ชื่อกฎหมาย/คำปรารภ/เหตุผล) ถ้ามี -> chunk แรก section_ref="บทนำ" (1) หา
"มาตรา N" ทั้งหมด พร้อมเก็บบริบท "หมวด N"/"ส่วนที่ N" ที่อยู่ก่อนหน้า 1 มาตรา = 1 chunk (2) ไม่มีมาตราเลย -> แยก
ย่อหน้า รวม/ตัดด้วย sentence tokenize ของ pythainlp ให้ไม่เกิน MAX_TOKENS (overlap 1 ประโยค) (3) บทนำ/มาตราที่ยาว
เกิน MAX_TOKENS ตัดย่อยตามประโยคแบบเดียวกัน ต่อท้าย section_ref ด้วย "(ต่อ N)" — ประโยคเดี่ยวที่ยังยาวเกิน budget
เอง (sent_tokenize ตัดไม่ได้อีก) ตัดย่อยด้วย word_tokenize ต่อที่ขอบคำ

token_count ของ chunk ที่คืนออกไปนับจาก content ทั้งก้อน (header+body) เพราะนั่นคือสิ่งที่ embed จริง budget ของ
body = MAX_TOKENS - token ของ header ของ chunk นั้น

import embed_model จาก main.py ตรงๆ (ไม่เคยถูก reassign หลัง startup — ต่างจาก bm25/knowledge_base_ids)"""
import re

from pythainlp.tokenize import sent_tokenize, word_tokenize

from main import embed_model

MAX_TOKENS = 350
_PREAMBLE_SECTION_REF = "บทนำ"
_CONSERVATIVE_CONT_SUFFIX = " (ต่อ 99)"  # ประมาณ budget แบบระวังไว้ก่อน (เผื่อเลขต่อท้ายเป็นสองหลัก) ก่อนรู้ว่าหั่นได้กี่ชิ้นจริง

_THAI_ARABIC_DIGITS = "๐-๙0-9"
_SECTION_KEYWORD_RE = re.compile(rf"^[ \t]*((?:หมวด|ส่วนที่)\s*[{_THAI_ARABIC_DIGITS}]+)[ \t]*(.*)$")
_ARTICLE_LINE_RE = re.compile(rf"^[ \t]*มาตรา[ \t]*([{_THAI_ARABIC_DIGITS}]+(?:/[{_THAI_ARABIC_DIGITS}]+)?)[ \t]*(.*)$")


def _count_tokens(text: str) -> int:
    """นับ token ด้วย tokenizer จริงของ embed_model (ไม่ใช่นับตัวอักษร)"""
    return len(embed_model.tokenizer.tokenize(text)) if text else 0


def _sentences(text: str) -> list[str]:
    return [s.strip() for s in sent_tokenize(text, engine="crfcut") if s.strip()]


def _split_long_sentence_by_words(sentence: str, max_tokens: int) -> list[str]:
    """ตัดประโยคเดียวที่ยาวเกิน budget เอง (sent_tokenize ตัดต่อไม่ได้แล้ว เช่น ข้อความไทยยาวไม่มีช่องว่าง/วรรคตอน)
    ด้วย word_tokenize เป็นช่วงๆ ไม่เกิน budget ตัดที่ขอบคำเสมอ ไม่ตัดกลางคำ"""
    words = word_tokenize(sentence, engine="newmm")
    parts = []
    current: list[str] = []
    for w in words:
        if current and _count_tokens("".join(current + [w])) > max_tokens:
            parts.append("".join(current))
            current = []
        current.append(w)
    if current:
        parts.append("".join(current))
    return parts


def _pack_sentences(sentences: list[str], max_tokens: int, overlap_sentences: int = 1) -> list[str]:
    """รวมประโยคเป็น chunk ทีละตัวแบบ greedy ไม่ให้เกิน max_tokens — ประโยคเดี่ยวที่ยาวเกิน budget เอง ถูกตัดย่อย
    ด้วย _split_long_sentence_by_words ก่อนเสมอ ไม่พก overlap ถ้าพกแล้วเกิน budget (กันเคสชิ้นที่ตัดมาใหญ่ใกล้ budget)"""
    if not sentences:
        return []
    expanded: list[str] = []
    for sent in sentences:
        if _count_tokens(sent) > max_tokens:
            expanded.extend(_split_long_sentence_by_words(sent, max_tokens))
        else:
            expanded.append(sent)

    chunks = []
    current: list[str] = []
    for sent in expanded:
        if current and _count_tokens(" ".join(current + [sent])) > max_tokens:
            chunks.append(" ".join(current))
            tail = current[-overlap_sentences:] if overlap_sentences > 0 else []
            if tail and _count_tokens(" ".join(tail + [sent])) > max_tokens:
                tail = []
            current = tail
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


def _parse_structure(text: str):
    """คืน (lines, section_at_line, consumed_section_lines, article_lines) ดิบๆ ให้ _parse_articles/_extract_preamble ใช้ร่วมกัน"""
    lines = text.split("\n")
    section_at_line, consumed_section_lines = _parse_sections(lines)
    article_lines = [
        (idx, m.group(1), m.group(2))
        for idx, line in enumerate(lines)
        for m in (_ARTICLE_LINE_RE.match(line),) if m
    ]
    return lines, section_at_line, consumed_section_lines, article_lines


def _parse_articles(text: str) -> list[dict]:
    """คืน [{article_no, section, body}] ตามลำดับที่เจอในเอกสาร — ว่างถ้าไม่มี 'มาตรา' เลยสักตัว"""
    lines, section_at_line, consumed_section_lines, article_lines = _parse_structure(text)
    articles = []
    for k, (idx, article_no, rest_of_line) in enumerate(article_lines):
        end_idx = article_lines[k + 1][0] if k + 1 < len(article_lines) else len(lines)
        body_lines = [rest_of_line] + [
            lines[i] for i in range(idx + 1, end_idx) if i not in consumed_section_lines
        ]
        body = re.sub(r"\n[ \t]*\n+", "\n", "\n".join(body_lines)).strip()
        articles.append({"article_no": article_no, "section": section_at_line.get(idx), "body": body})
    return articles


def _extract_preamble(text: str) -> str:
    """ข้อความก่อน 'มาตรา' แรก (ชื่อกฎหมาย/คำปรารภ/เหตุผล) — ไม่รวมบรรทัดหัว 'หมวด/ส่วนที่' ที่ถูกย้ายไปเป็น
    บริบทของ chunk ถัดไปแล้ว (กันซ้ำ) ว่างถ้าไม่มีมาตราเลย หรือไม่มีข้อความก่อนมาตราแรกจริงๆ"""
    lines, _section_at_line, consumed_section_lines, article_lines = _parse_structure(text)
    if not article_lines:
        return ""
    first_idx = article_lines[0][0]
    preamble_lines = [lines[i] for i in range(first_idx) if i not in consumed_section_lines]
    return re.sub(r"\n[ \t]*\n+", "\n", "\n".join(preamble_lines)).strip()


def _build_header(source_title: str, section: str | None, section_ref: str | None) -> str:
    parts = [source_title] + [p for p in (section, section_ref) if p]
    return "[" + " > ".join(parts) + "]"


_HEADER_BOUNDARY_SAFETY_MARGIN = 5  # tokenizer จริงนับ header+body ต่อกันได้มากกว่านับแยกแล้วบวกกันเล็กน้อยได้ (boundary effect) กันไว้ก่อน


def _body_budget(source_title: str, section: str | None, section_ref: str | None) -> int:
    """budget ของเนื้อ body = MAX_TOKENS - token ของ header ของ chunk นั้น (header+body คือก้อนที่ embed จริง) กัน
    ส่วนต่างไว้เผื่อ tokenizer นับ header+body ต่อกันได้ไม่เท่ากับนับแยกแล้วบวกกันเป๊ะๆ"""
    return MAX_TOKENS - _count_tokens(_build_header(source_title, section, section_ref)) - _HEADER_BOUNDARY_SAFETY_MARGIN


def _build_chunk(source_title: str, section: str | None, section_ref: str | None, body: str) -> dict:
    """token_count นับจาก content ทั้งก้อน (header+body) — ตัวเลขเดียวกับที่ embed_model จะ encode จริง"""
    header = _build_header(source_title, section, section_ref)
    content = f"{header}\n{body}".strip()
    return {"content": content, "header": header, "section_ref": section_ref, "token_count": _count_tokens(content)}


def _emit_section_chunks(source_title: str, section: str | None, section_ref: str, body: str) -> list[dict]:
    """สร้าง chunk จาก (section_ref, body) เดียว — พอดี budget ได้ 1 chunk ยาวเกินหั่นด้วย _pack_sentences แล้ว
    ต่อท้าย section_ref ด้วย "(ต่อ N)" ใช้ budget แบบระวังไว้ก่อนตอนหั่น (เผื่อ header ยาวขึ้นจากเลขต่อท้ายจริง)"""
    if _count_tokens(body) <= _body_budget(source_title, section, section_ref):
        return [_build_chunk(source_title, section, section_ref, body)]

    conservative_budget = _body_budget(source_title, section, section_ref + _CONSERVATIVE_CONT_SUFFIX)
    parts = _pack_sentences(_sentences(body), conservative_budget, overlap_sentences=1)
    chunks = []
    for i, part in enumerate(parts):
        ref = section_ref if i == 0 else f"{section_ref} (ต่อ {i})"
        chunks.append(_build_chunk(source_title, section, ref, part))
    return chunks


def _chunk_fallback_paragraphs(text: str, source_title: str) -> list[dict]:
    paragraphs = [p.strip() for p in re.split(r"\n[ \t]*\n", text) if p.strip()]
    all_sentences = []
    for para in paragraphs:
        all_sentences.extend(_sentences(para))
    budget = _body_budget(source_title, None, None)
    groups = _pack_sentences(all_sentences, budget, overlap_sentences=1)
    return [_build_chunk(source_title, None, None, g) for g in groups]


def chunk_document(text: str, source_title: str) -> list[dict]:
    """คืน [{content, header, section_ref, token_count}] — content มี header ขึ้นต้นเสมอ (ใช้บันทึกตรงๆ ผ่าน
    add_knowledge_chunk ได้เลย)"""
    articles = _parse_articles(text)
    if not articles:
        return _chunk_fallback_paragraphs(text, source_title)

    chunks = []
    preamble = _extract_preamble(text)
    if preamble:
        chunks.extend(_emit_section_chunks(source_title, None, _PREAMBLE_SECTION_REF, preamble))

    for art in articles:
        section_ref = f"มาตรา {art['article_no']}"
        chunks.extend(_emit_section_chunks(source_title, art["section"], section_ref, art["body"]))
    return chunks
