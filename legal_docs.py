"""ร่างบันทึกความเห็นทางกฎหมายเบื้องต้น (.docx) จากสถานการณ์ที่ผู้ใช้เล่าในแชท — แยกจากระบบแก้ไฟล์ Excel ทั้งหมด
(ไม่แตะ EditableDocument/DocumentVersion/focus เลย) รองรับกฎหมายไทยทุกหมวด ไม่จำกัดแค่ธุรกิจต่างด้าว/ภาษี

ขั้นตอน: (1) Claude ดึง slots จากข้อความ+ประวัติ ไม่เดาข้อเท็จจริงที่ไม่ได้พูดถึง (2) ค้น legal_retrieval.legal_search
หา chunk ที่เกี่ยวข้อง (3) Claude เรียบเรียงเนื้อหาบันทึกเป็น JSON มีโครงสร้าง อ้าง chunk_ids ได้ (4) โค้ดตรวจ
chunk_ids ให้อยู่ในชุดที่ค้นเจอจริงเท่านั้น แล้วส่งให้ legal_memo_format.build_memo_docx() สร้างไฟล์ .docx ทางการ

import จาก main.py แบบ deferred (หลังนิยาม client/ROUTER_MODEL/ADVISOR_MODEL/save_legal_memo) เหมือน
legal_retrieval.py — ดู main.py บรรทัดที่ import โมดูลนี้"""
import legal_retrieval
from legal_memo_format import build_memo_docx
from llm_tools import call_with_required_tool
from main import client, get_chat_messages, save_legal_memo, ROUTER_MODEL, ADVISOR_MODEL

MEMO_SLOTS_TOOL = {
    "name": "extract_memo_slots",
    "description": (
        "ดึงข้อมูลสถานการณ์ของผู้ใช้จากข้อความล่าสุด+ประวัติสนทนา เพื่อร่างบันทึกความเห็นทางกฎหมาย ห้ามเดา/แต่ง"
        "ข้อมูลที่ผู้ใช้ไม่ได้พูดถึงเด็ดขาด ไม่แน่ใจให้ใส่ null เสมอ ต้องเรียกเครื่องมือนี้เสมอ"
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "client_name": {"type": ["string", "null"], "description": "ชื่อลูกค้า/ผู้ขอคำปรึกษา ไม่ได้บอกใส่ null"},
            "legal_area": {
                "type": ["string", "null"],
                "description": (
                    "หมวดกฎหมายที่เกี่ยวข้อง รองรับกฎหมายไทยทุกหมวด เช่น อาญา, จราจร (เมาแล้วขับ), ทรัพย์สิน (บุกรุก), "
                    "หลักทรัพย์ (ปั่นหุ้น/ฉ้อโกง), แพ่ง, ครอบครัว, แรงงาน, ภาษี, ธุรกิจ/ธุรกิจต่างด้าว, ที่ดิน, สัญญา ฯลฯ "
                    "ไม่จำกัดเฉพาะหมวดใดหมวดหนึ่ง ไม่ชัดเจนใส่ null"
                ),
            },
            "facts_summary": {
                "type": ["string", "null"],
                "description": "สรุปสถานการณ์/ข้อเท็จจริงทั้งหมดที่ผู้ใช้เล่ามาเป็นย่อหน้าเดียว ไม่มีเลยใส่ null",
            },
            "key_facts": {
                "type": ["array", "null"],
                "description": "ข้อเท็จจริงสำคัญแยกเป็นข้อสั้นๆ (สัดส่วนหุ้น/ประเภทธุรกิจ/จำนวนเงิน/วันที่ ฯลฯ) ไม่มีใส่ null",
                "items": {"type": "string"},
            },
            "user_questions": {
                "type": ["array", "null"],
                "description": "คำถามที่ผู้ใช้ถามไว้ตรงๆ (ถ้ามี) ไม่มีใส่ null",
                "items": {"type": "string"},
            },
        },
        "required": ["client_name", "legal_area", "facts_summary", "key_facts", "user_questions"],
    },
}

_MEMO_PARAGRAPH_ITEM_SCHEMA = {
    "type": "object",
    "properties": {
        "text": {"type": "string"},
        "chunk_ids": {
            "type": "array", "items": {"type": "integer"},
            "description": "id ของ chunk จากฐานความรู้ที่ใช้เขียนย่อหน้านี้จริง (จากรายการที่ให้มา) ไม่มีใส่ []",
        },
    },
    "required": ["text", "chunk_ids"],
}

MEMO_CONTENT_TOOL = {
    "name": "draft_memo_content",
    "description": (
        "เรียบเรียงเนื้อหาบันทึกความเห็นทางกฎหมายเบื้องต้นจากสถานการณ์และ chunk ฐานความรู้ที่ให้มา รองรับกฎหมายไทย"
        "ทุกหมวด ไม่จำกัดเฉพาะหมวดใดหมวดหนึ่ง ห้ามเขียนทำนอง \"นอกเหนือความเชี่ยวชาญ\" ต้องเรียกเครื่องมือนี้เสมอ"
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "issues": {"type": "array", "items": _MEMO_PARAGRAPH_ITEM_SCHEMA},
            "analysis": {"type": "array", "items": _MEMO_PARAGRAPH_ITEM_SCHEMA},
            "our_side_arguments": {
                "type": "array", "items": _MEMO_PARAGRAPH_ITEM_SCHEMA,
                "description": "แนวทางข้อต่อสู้ฝ่ายผู้ขอความเห็น — คิดเชิงกลยุทธ์ ไม่ใช่แค่สรุปตัวบทซ้ำ",
            },
            "other_side_arguments": {
                "type": "array", "items": _MEMO_PARAGRAPH_ITEM_SCHEMA,
                "description": "ข้อโต้แย้งที่คู่กรณี/เจ้าหน้าที่น่าจะยกขึ้น และจุดอ่อนของฝ่ายเรา",
            },
            "conclusion": _MEMO_PARAGRAPH_ITEM_SCHEMA,
            "next_steps": {"type": "array", "items": {"type": "string"}},
        },
        "required": ["issues", "analysis", "our_side_arguments", "other_side_arguments", "conclusion", "next_steps"],
    },
}

MEMO_CONTENT_SYSTEM_PROMPT = (
    "คุณกำลังเรียบเรียงบันทึกความเห็นทางกฎหมายเบื้องต้นให้ทนายความใช้ประกอบการให้คำปรึกษา รองรับกฎหมายไทยทุกหมวด "
    "เช่น อาญา, จราจร (เมาแล้วขับ), ทรัพย์สิน (บุกรุก), หลักทรัพย์ (ปั่นหุ้น/ฉ้อโกง), แพ่ง, ครอบครัว, แรงงาน, ภาษี, "
    "ธุรกิจ/ธุรกิจต่างด้าว ฯลฯ ไม่จำกัดเฉพาะหมวดใดหมวดหนึ่ง ห้ามเขียนว่าเรื่องนี้ \"นอกเหนือความเชี่ยวชาญ\" เด็ดขาด "
    "ไม่ว่าจะเป็นกฎหมายหมวดใด\n"
    "- ได้รับ chunk จากฐานความรู้มาพร้อม id — ถ้า chunk ตรงประเด็น ให้ใช้เป็นหลักและใส่ id นั้นใน chunk_ids ของ"
    "ย่อหน้าที่ใช้ข้อมูลนั้นจริง\n"
    "- chunk ที่ให้มาไม่ครอบคลุมบางประเด็น ใช้ความรู้ทั่วไปของคุณเขียนเสริมได้ (ใส่ chunk_ids เป็น [] สำหรับย่อหน้า"
    "นั้น) ระบุชื่อกฎหมาย/หน่วยงานให้ถูกต้องเท่าที่มั่นใจ\n"
    "- ไม่แน่ใจตัวเลข/อัตรา/มาตราที่แน่นอน ให้เขียนเป็นภาพรวมแทนการระบุตัวเลขที่ไม่แน่ใจ ห้ามแต่งตัวเลข/มาตราขึ้นมา\n"
    "- ห้ามเดา/แต่งข้อเท็จจริงของผู้ใช้ที่ไม่ได้อยู่ในสถานการณ์ที่ให้มาเด็ดขาด\n"
    "- issues: ประเด็นที่ต้องพิจารณาแต่ละข้อ, analysis: การวิเคราะห์ทางกฎหมายแต่ละประเด็น\n"
    "- our_side_arguments: ข้อต่อสู้ ช่องทาง และการตีความที่เป็นประโยชน์ต่อฝ่ายผู้ขอความเห็นให้มากที่สุดเท่าที่กฎหมาย"
    "เปิดช่อง รวมถึงทางเลือกทางกฎหมายที่คนทั่วไปมักไม่รู้\n"
    "- other_side_arguments: ข้อโต้แย้งที่คู่กรณีหรือเจ้าหน้าที่น่าจะยกขึ้น และจุดอ่อนของฝ่ายเรา — ให้คิดเชิงกลยุทธ์ "
    "ไม่ใช่แค่สรุปตัวบทซ้ำ\n"
    "- conclusion: ความเห็นเบื้องต้นสรุปรวม 1 รายการ, next_steps: ขั้นตอนที่แนะนำให้ทำต่อ (ข้อความล้วน ไม่ต้องมี chunk_ids)"
)


def _extract_memo_slots(query: str, history: list) -> dict:
    """เรียก Claude (ROUTER_MODEL — งานดึงค่าโครงสร้าง ไม่ต้องเขียนประโยคยาว) ดึง slots จากข้อความ+ประวัติ 6
    ข้อความล่าสุด คืน dict ตาม MEMO_SLOTS_TOOL เสมอ (ทุกช่อง None ถ้า parse/เรียกไม่สำเร็จ)"""
    history_lines = [f"{m['role']}: {(m.get('content') or '')[:500]}" for m in (history or [])[-6:]]
    history_text = ("ประวัติสนทนาล่าสุด:\n" + "\n".join(history_lines) + "\n\n") if history_lines else ""
    prompt = f"{history_text}ข้อความล่าสุดของผู้ใช้: {query}"

    response, _reason, _diag = call_with_required_tool(
        client, ROUTER_MODEL, 800, "", [{"role": "user", "content": prompt}],
        [MEMO_SLOTS_TOOL], "extract_memo_slots",
    )
    if response is None:
        return {"client_name": None, "legal_area": None, "facts_summary": None, "key_facts": None, "user_questions": None}
    tool_use_block = next(b for b in response.content if b.type == "tool_use" and b.name == "extract_memo_slots")
    return tool_use_block.input


def _format_chunks_for_prompt(chunks: list[dict]) -> str:
    if not chunks:
        return "(ไม่มี chunk ที่เกี่ยวข้องพอ — ใช้ความรู้ทั่วไปทั้งหมด)"
    lines = []
    for c in chunks:
        tags_text = "/".join(c.get("tags") or []) or "-"
        lines.append(f"[id={c['id']}] (tag: {tags_text}) {c['content']}")
    return "\n\n".join(lines)


_EMPTY_DIAG = {"stop_reason": None, "blocks": "-", "out_tokens": None}


def _log_memo_draft_failed(reason: str, diag: dict) -> None:
    print(
        f"[LegalMemo] draft_failed reason={reason} stop_reason={diag.get('stop_reason')} "
        f"blocks={diag.get('blocks')} out_tokens={diag.get('out_tokens')}"
    )


def _draft_memo_content(slots: dict, grounded_chunks: list[dict]):
    """เรียก Claude (ADVISOR_MODEL — ต้องการภาษาไทยที่แม่นยำในคำตอบยาว เหมือน advisor_answer) เรียบเรียงเนื้อหา
    บันทึก คืน None ถ้าเรียกไม่สำเร็จ (log [LegalMemo] draft_failed รายละเอียดเสมอ ไม่ปล่อย exception หลุดเป็น 500)"""
    facts_lines = "\n".join(f"- {f}" for f in (slots.get("key_facts") or []))
    questions_lines = "\n".join(f"- {q}" for q in (slots.get("user_questions") or []))
    prompt = (
        f"หมวดกฎหมาย: {slots.get('legal_area') or 'ไม่ระบุ'}\n\n"
        f"สถานการณ์: {slots.get('facts_summary')}\n\n"
        f"ข้อเท็จจริงสำคัญ:\n{facts_lines or '(ไม่มี)'}\n\n"
        f"คำถามของผู้ใช้:\n{questions_lines or '(ไม่มี)'}\n\n"
        f"chunk จากฐานความรู้:\n{_format_chunks_for_prompt(grounded_chunks)}"
    )
    try:
        response, reason, diag = call_with_required_tool(
            client, ADVISOR_MODEL, 6000, MEMO_CONTENT_SYSTEM_PROMPT, [{"role": "user", "content": prompt}],
            [MEMO_CONTENT_TOOL], "draft_memo_content",
        )
    except Exception as e:
        _log_memo_draft_failed("api_error", {**_EMPTY_DIAG, "stop_reason": type(e).__name__})
        return None
    if response is None:
        _log_memo_draft_failed(reason or "no_tool_use", diag or _EMPTY_DIAG)
        return None
    tool_use_block = next(b for b in response.content if b.type == "tool_use" and b.name == "draft_memo_content")
    return tool_use_block.input


def _render_paragraph(item: dict, grounded_ids: set[int]) -> dict:
    """กรอง chunk_ids ให้เหลือเฉพาะที่อยู่ในชุด grounded จริง (ตัดที่ LLM อ้างมั่วทิ้ง) — ไม่มี label/อีโมจิ
    (legal_memo_format.build_memo_docx ใส่เลขอ้างอิง [n] เองจาก chunk_ids)"""
    valid_ids = [cid for cid in (item.get("chunk_ids") or []) if cid in grounded_ids]
    return {"text": item.get("text") or "", "chunk_ids": valid_ids}


def _missing_slot_labels(slots: dict) -> list[str]:
    labels = []
    if not slots.get("client_name"):
        labels.append("ชื่อลูกค้า")
    if not slots.get("legal_area"):
        labels.append("หมวดกฎหมาย")
    if not slots.get("key_facts"):
        labels.append("ข้อเท็จจริงสำคัญแบบละเอียด")
    if not slots.get("user_questions"):
        labels.append("คำถามเฉพาะเจาะจงของผู้ใช้")
    return labels


def dispatch_draft_legal_document(query: str, user_id: int, chat_id: int) -> dict:
    """จุดเดียวที่ main.py เรียกสำหรับ intent "draft_legal_document" — ไม่ผ่าน truth_guard/run_agentic_tool_loop
    เลย (ไม่ใช่การแก้ไฟล์ Excel) คืน {"document_id": None, "answer": ...} เสมอ"""
    history = get_chat_messages(chat_id) if chat_id else []
    slots = _extract_memo_slots(query, history)

    if not slots.get("facts_summary"):
        return {"document_id": None, "answer": "เล่าสถานการณ์ให้ฟังหน่อยครับ จะได้ร่างบันทึกความเห็นให้ถูกต้อง"}

    search_query = slots["facts_summary"] + (
        "\n" + "\n".join(slots.get("user_questions") or []) if slots.get("user_questions") else ""
    )
    results = legal_retrieval.legal_search(search_query, k=5, rerank_top_k=5)
    grounded_chunks = [r for r in results if r["score"] >= legal_retrieval.OUT_OF_KB_SCORE_THRESHOLD]
    grounded_ids = {c["id"] for c in grounded_chunks}
    grounded_by_id = {c["id"]: c for c in grounded_chunks}

    memo_content = _draft_memo_content(slots, grounded_chunks)
    if memo_content is None:
        return {"document_id": None, "answer": "ร่างบันทึกไม่สำเร็จ ลองพิมพ์ขอใหม่อีกครั้งครับ"}

    rendered = {
        "issues": [_render_paragraph(item, grounded_ids) for item in (memo_content.get("issues") or [])],
        "analysis": [_render_paragraph(item, grounded_ids) for item in (memo_content.get("analysis") or [])],
        "arguments": {
            "our_side": [_render_paragraph(item, grounded_ids) for item in (memo_content.get("our_side_arguments") or [])],
            "other_side": [_render_paragraph(item, grounded_ids) for item in (memo_content.get("other_side_arguments") or [])],
        },
        "conclusion": _render_paragraph(memo_content.get("conclusion") or {}, grounded_ids),
        "next_steps": [s for s in (memo_content.get("next_steps") or []) if s],
    }
    if not rendered["issues"] and not rendered["analysis"] and not rendered["conclusion"]["text"]:
        print("[LegalMemo] draft_failed reason=empty_content")
        return {"document_id": None, "answer": "ร่างบันทึกไม่สำเร็จ ลองพิมพ์ขอใหม่อีกครั้งครับ"}

    # ลำดับการอ้างต้องตรงกับลำดับที่ build_memo_docx แสดงผล (issues -> analysis -> arguments -> conclusion)
    # เพราะเลข [n] ในเอกสารนับตามลำดับที่เจอครั้งแรกใน used_chunks ไม่ใช่ sorted id
    all_items = (
        rendered["issues"] + rendered["analysis"]
        + rendered["arguments"]["our_side"] + rendered["arguments"]["other_side"]
        + [rendered["conclusion"]]
    )
    used_ids = list(dict.fromkeys(cid for item in all_items for cid in item["chunk_ids"]))
    used_chunks = [grounded_by_id[cid] for cid in used_ids if cid in grounded_by_id]

    docx_bytes = build_memo_docx(slots, rendered, used_chunks)
    title = f"บันทึกความเห็นทางกฎหมายเบื้องต้น - {slots.get('client_name') or 'ไม่ระบุชื่อลูกค้า'}"
    memo_id = save_legal_memo(user_id, chat_id, title, slots, docx_bytes)

    missing = _missing_slot_labels(slots)

    download_url = f"/api/legal-memo/{memo_id}/download"
    answer_lines = [f"ร่างบันทึกความเห็นทางกฎหมายเรียบร้อยครับ ดาวน์โหลด: [ลิงก์]({download_url})"]
    if missing:
        answer_lines.append("ยังไม่ได้ระบุ: " + ", ".join(missing))

    return {"document_id": None, "answer": "\n".join(answer_lines)}
