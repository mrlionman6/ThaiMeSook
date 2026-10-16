"""ค้น KB กฎหมายด้วย hybrid (BM25+vector) บวกคะแนนเสริมจาก tag ที่เกี่ยวข้องกับคำถาม คืน {id,content,tags,score} ให้
โค้ดแนบ metadata อ้างอิงเองได้ — เรียกจาก main.py: rebuild_index()/_prepare_rag_context() (ดู legal_search)

import จาก main.py แบบ deferred (หลังนิยาม rerank_with_scores) เหมือน document_fields.py — bm25/knowledge_base_ids/
knowledge_base_texts/embed_model ถูก reassign ใหม่ทุกครั้งที่ rebuild_index() ทำงาน จึงอ้างผ่าน `main.<ชื่อ>` สดทุก
ครั้งที่ใช้ แทนที่จะ `from main import` ตรงๆ"""
import json
import time
from typing import Optional

import numpy as np
from pythainlp.tokenize import word_tokenize

import main
from main import get_all_tags, get_all_chunk_tags_map, get_vector_scores_for_all, rerank_with_scores

# tag ที่ cosine(query, tag) >= ค่านี้ถือว่า "เกี่ยวข้อง" กับคำถาม — ปรับได้ (ดู relevant_tag_scores)
DEFAULT_TAG_RELEVANCE_THRESHOLD = 0.80
# น้ำหนักคะแนนเสริมจาก tag ในสูตรรวม (ดู legal_search) — beta=0 ปิดส่วนนี้ ใช้เป็น baseline เทียบในหน้า admin
DEFAULT_BETA = 0.15
# คะแนน rerank (CrossEncoder, raw logit ไม่ใช่ 0-1) อันดับ 1 ต่ำกว่านี้ -> สลับ system prompt เป็นโหมดความรู้ทั่วไป
# (ไม่ใช่ปฏิเสธ) เป็นจุดเริ่มต้นคร่าวๆ ปรับจริงด้วยหน้า admin "วัดผลระบบค้นกฎหมาย"
OUT_OF_KB_SCORE_THRESHOLD = -3.0

_tag_embeddings: dict[str, np.ndarray] = {}  # {tag_name: embedding} — รีเฟรชทุกครั้งที่ main.rebuild_index() ทำงาน


def refresh_tag_embeddings() -> None:
    """embed ชื่อ tag ทุกตัวในระบบใหม่ทั้งหมด ด้วย e5 instruction prefix "passage: " (ชื่อ tag ถือเป็น "เอกสาร"
    ที่จะถูกค้นหาด้วย query) เรียกจาก main.rebuild_index() ทุกครั้ง (ดู docstring ของไฟล์นี้)"""
    global _tag_embeddings
    tags = get_all_tags()
    if not tags:
        _tag_embeddings = {}
        return
    names = [t["name"] for t in tags]
    vectors = main.embed_model.encode([f"passage: {name}" for name in names])
    _tag_embeddings = {name: np.array(vec) for name, vec in zip(names, vectors)}


def _cosine(a: np.ndarray, b: np.ndarray) -> float:
    denom = np.linalg.norm(a) * np.linalg.norm(b)
    return float(np.dot(a, b) / denom) if denom else 0.0


def tag_scores(query: str) -> dict[str, float]:
    """cosine(query, tag) ของทุก tag ที่มีอยู่ตอนนี้ คืน {tag_name: cosine} ดิบทั้งหมด (ยังไม่กรอง threshold) —
    query ใส่ e5 instruction prefix "query: " ให้ถูกคู่กับ "passage: " ที่ใช้ตอน embed ชื่อ tag"""
    if not _tag_embeddings:
        return {}
    query_vec = np.array(main.embed_model.encode(f"query: {query}"))
    return {name: _cosine(query_vec, vec) for name, vec in _tag_embeddings.items()}


def relevant_tag_scores(query: str, threshold: float = DEFAULT_TAG_RELEVANCE_THRESHOLD) -> dict[str, float]:
    """เหมือน tag_scores แต่กรองเหลือเฉพาะ tag ที่คะแนนเกิน threshold (tag ที่ "เกี่ยวข้องจริง" กับคำถามนี้)"""
    return {name: score for name, score in tag_scores(query).items() if score >= threshold}


def _normalize(scores: np.ndarray) -> np.ndarray:
    if scores.max() == scores.min():
        return np.zeros_like(scores)
    return (scores - scores.min()) / (scores.max() - scores.min())


def legal_search(
    query: str, k: int = 5, alpha: float = 0.5, beta: float = DEFAULT_BETA,
    tag_threshold: float = DEFAULT_TAG_RELEVANCE_THRESHOLD, rerank_top_k: int = 3,
) -> list[dict]:
    """ค้น KB กฎหมายด้วย BM25+vector เหมือน hybrid_search เดิม (beta=0 ต้องได้ผลเท่ากันเป๊ะ) บวกคะแนนเสริม
    beta*max(ความเกี่ยวข้องของ tag ของ chunk) ก่อน rerank — ไม่เคยกรอง chunk ทิ้งด้วย tag เด็ดขาด"""
    start = time.monotonic()
    ids = main.knowledge_base_ids
    texts = main.knowledge_base_texts
    if not ids:
        print("[LegalRetrieval] top_ids=[] tag_hits=[] latency_ms=0")
        return []

    tokenized_query = word_tokenize(query, engine="newmm")
    bm25_scores = np.array(main.bm25.get_scores(tokenized_query))

    query_embedding = main.embed_model.encode(query).tolist()
    vector_scores_map = get_vector_scores_for_all(query_embedding)
    vector_scores = np.array([vector_scores_map.get(cid, 0.0) for cid in ids])

    base_scores = alpha * _normalize(vector_scores) + (1 - alpha) * _normalize(bm25_scores)

    chunk_tags_map = get_all_chunk_tags_map()
    tag_hits = relevant_tag_scores(query, threshold=tag_threshold) if beta else {}
    tag_boost = np.zeros(len(ids))
    if beta and tag_hits:
        for i, cid in enumerate(ids):
            relevant = [tag_hits[name] for name in chunk_tags_map.get(cid, []) if name in tag_hits]
            if relevant:
                tag_boost[i] = max(relevant)

    final_scores = base_scores + beta * tag_boost
    top_idx = final_scores.argsort()[::-1][:k]

    candidates = [
        {"id": ids[i], "content": texts[i], "tags": chunk_tags_map.get(ids[i], []), "score": float(final_scores[i])}
        for i in top_idx
    ]

    reranked_texts, rerank_scores = rerank_with_scores(query, [c["content"] for c in candidates], top_k=rerank_top_k)

    # rerank_with_scores คืนแค่ (content, score) เรียงใหม่ — จับคู่กลับเข้า dict เดิมด้วย content (pop ตัวแรกที่ยัง
    # ไม่ถูกใช้ กันจับคู่ผิดซ้ำถ้า 2 chunk มีเนื้อหาเดียวกันเป๊ะ)
    by_content: dict[str, list[dict]] = {}
    for c in candidates:
        by_content.setdefault(c["content"], []).append(c)
    results = []
    for text, score in zip(reranked_texts, rerank_scores):
        matched = by_content[text].pop(0)
        results.append({**matched, "score": float(score)})

    latency_ms = int((time.monotonic() - start) * 1000)
    print(f"[LegalRetrieval] top_ids={[r['id'] for r in results]} tag_hits={list(tag_hits.keys())} latency_ms={latency_ms}")
    return results


# ---------- วัดผล retrieval ด้วย eval/legal_eval_v1.json (ไม่เรียก LLM เลย — ดู run_eval) ----------
DEFAULT_EVAL_PATH = "eval/legal_eval_v1.json"


def _case_match_rank(case: dict, results: list[dict], limit: int = 3) -> Optional[int]:
    """คืนอันดับ (1-indexed) ของผลลัพธ์แรกที่ content มี chunk_keys ของ case นี้ปนอยู่ (substring ไม่สนตัวพิมพ์เล็ก-
    ใหญ่) คืน None ถ้าไม่เจอ"""
    keys = [k.lower() for k in (case.get("chunk_keys") or [])]
    if not keys:
        return None
    for rank, r in enumerate(results[:limit], start=1):
        content_lower = r["content"].lower()
        if any(key in content_lower for key in keys):
            return rank
    return None


def _score_case(case: dict, results: list[dict]) -> dict:
    """ให้คะแนน 1 คำถาม 1 รอบ (baseline หรือ tag+vector) — type ปกติ: hit@1/hit@3/mrr จาก _case_match_rank type=
    "out_of_kb": out_of_kb_pass แทน (ผ่านเมื่อคะแนน rerank อันดับ 1 ต่ำกว่า OUT_OF_KB_SCORE_THRESHOLD)"""
    top_score = results[0]["score"] if results else float("-inf")
    found_ids = [r["id"] for r in results[:3]]
    if case.get("type") == "out_of_kb":
        return {
            "rank": None, "hit_at_1": False, "hit_at_3": False, "mrr": 0.0,
            "out_of_kb_pass": top_score < OUT_OF_KB_SCORE_THRESHOLD,
            "top_score": top_score, "found_ids": found_ids,
        }
    rank = _case_match_rank(case, results, limit=3)
    return {
        "rank": rank, "hit_at_1": rank == 1, "hit_at_3": rank is not None,
        "mrr": (1.0 / rank) if rank else 0.0, "out_of_kb_pass": None,
        "top_score": top_score, "found_ids": found_ids,
    }


def _summarize(results: list[dict]) -> dict:
    n = len(results)
    if n == 0:
        return {"n": 0, "hit_at_1": 0.0, "hit_at_3": 0.0, "mrr": 0.0}
    return {
        "n": n,
        "hit_at_1": sum(r["hit_at_1"] for r in results) / n,
        "hit_at_3": sum(r["hit_at_3"] for r in results) / n,
        "mrr": sum(r["mrr"] for r in results) / n,
    }


def _aggregate(cases: list[dict], scores: list[dict]) -> dict:
    """สรุปผลแยกตาม type + รวม (ไม่รวม out_of_kb เข้า hit@1/hit@3/mrr ปกติ — วัดด้วย pass_rate แยกต่างหาก)"""
    by_type: dict[str, list[dict]] = {}
    out_of_kb_scores = []
    for case, score in zip(cases, scores):
        if case.get("type") == "out_of_kb":
            out_of_kb_scores.append(score)
        else:
            by_type.setdefault(case.get("type") or "ไม่ระบุ", []).append(score)

    summary_by_type = {t: _summarize(group) for t, group in by_type.items()}
    overall = _summarize([score for group in by_type.values() for score in group])
    out_of_kb_summary = {
        "n": len(out_of_kb_scores),
        "pass_rate": (
            sum(1 for s in out_of_kb_scores if s["out_of_kb_pass"]) / len(out_of_kb_scores)
            if out_of_kb_scores else 0.0
        ),
    }
    return {"by_type": summary_by_type, "overall": overall, "out_of_kb": out_of_kb_summary}


def run_eval(eval_path: str = DEFAULT_EVAL_PATH) -> dict:
    """รันทุกคำถามใน eval_path 2 แบบ (baseline: beta=0, tag_vector: beta=DEFAULT_BETA) ไม่เรียก LLM เลย คืน
    {"baseline":{...}, "tag_vector":{...}, "cases":[...]} ให้หน้า admin render ตาราง"""
    with open(eval_path, encoding="utf-8") as f:
        cases = json.load(f)["cases"]

    baseline_scores, tag_vector_scores, rows = [], [], []
    for case in cases:
        query = case["query"]
        baseline_results = legal_search(query, beta=0)
        tag_vector_results = legal_search(query, beta=DEFAULT_BETA)

        baseline_score = _score_case(case, baseline_results)
        tag_vector_score = _score_case(case, tag_vector_results)
        baseline_scores.append(baseline_score)
        tag_vector_scores.append(tag_vector_score)

        rows.append({
            "id": case["id"], "type": case.get("type"), "query": query,
            "chunk_keys": case.get("chunk_keys") or [],
            "baseline": baseline_score, "tag_vector": tag_vector_score,
        })

    return {
        "baseline": _aggregate(cases, baseline_scores),
        "tag_vector": _aggregate(cases, tag_vector_scores),
        "cases": rows,
        "threshold": OUT_OF_KB_SCORE_THRESHOLD,  # ค่าจริงที่ production ใช้ตัดสิน 📚/💡 — ให้หน้า admin ใช้เป็นค่าเริ่มต้น
    }
