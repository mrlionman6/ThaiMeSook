"""
tools/eval_retrieval.py (v2 — strict-recall round)

Offline retrieval-quality evaluation for ThaiMeSook's hybrid search pipeline.
Reimplements the pipeline from main.py from scratch (does NOT import main.py or
db.py, does NOT connect to any database) against a static knowledge_base_export.json
export, to compare candidate reranking strategies (A/C/F/D/E) without touching
production data or requiring a live DB.

Usage:
    python tools/eval_retrieval.py <path-to-knowledge_base_export.json>

Terminal output is English/ASCII only (numbers, PASS/FAIL, variant labels) per
project convention (this shell mangles Thai text). Full detail is written to
eval_out/eval_results_v2.txt.

Round 2 change from v1: strict metric. A question only "hits" (hit_strict) when
EVERY chunk in its necessary set is present in the variant's final set — not just
any one of them (that looser criterion is kept alongside as hit_lenient, for
comparison). All variants in this round use NO e5 prefix on embeddings (matches
current production behavior exactly — v1's B variant with e5 prefixes is dropped).

Pipeline reproduced from main.py (read for reference only, not imported):
  - BM25Okapi over word_tokenize(doc, engine="newmm") from pythainlp.
  - SentenceTransformer('intfloat/multilingual-e5-large') cosine similarity,
    computed with plain numpy (query vs every chunk vector), no e5 prefix.
  - hybrid: final = 0.5*minmax(vector_scores) + 0.5*minmax(bm25_scores), top 5.
  - CrossEncoder('cross-encoder/mmarco-mMiniLMv2-L12-H384-v1') for reranking,
    using RAW (unprefixed) query/chunk text.

Variants (all share the identical hybrid top-5, since none use e5 prefixes):
  A = hybrid top5 -> rerank -> top3 (current production behavior).
  C = hybrid top5 -> rerank -> keep all 5 (sorted by rerank score).
  F = hybrid top5, kept in original hybrid order, reranker never reorders/truncates.
  D = hybrid top3 directly, no reranker call at all.
  E = hybrid top5 -> rerank with CrossEncoder('BAAI/bge-reranker-v2-m3') -> top3.
      Optional: if this model fails to load/download, E is skipped entirely and
      is NOT treated as a stop condition (only A/C/F/D are mandatory).
"""

import sys
import os
import json

import numpy as np
from rank_bm25 import BM25Okapi
from pythainlp.tokenize import word_tokenize
from sentence_transformers import SentenceTransformer, CrossEncoder

# (question_number, question_text, necessary_chunk_numbers)
# q12 is an unrelated control question (necessary=[]) excluded from recall.
QUESTIONS = [
    (1, "ต่างชาติถือหุ้นในบริษัทไทยที่ทำธุรกิจบริการได้ไม่เกินกี่เปอร์เซ็นต์", [1, 4]),
    (2, "ต่างชาติทำนาหรือทำสวนได้ไหม", [2]),
    (3, "ธุรกิจที่ต้องผ่านการอนุมัติจาก ครม. คืออะไร", [3]),
    (4, "ทุนจดทะเบียนขั้นต่ำเท่าไรถึงจะขอ FBL ได้", [5]),
    (5, "ขั้นตอนการขอใบอนุญาตประกอบธุรกิจคนต่างด้าวทำอย่างไร", [6]),
    (6, "ให้คนไทยถือหุ้นแทนต่างชาติผิดกฎหมายไหม", [7]),
    (7, "ต่างชาติซื้อที่ดินได้ไหมถ้าได้ BOI", [12]),
    (8, "ได้ BOI แล้วยกเว้นภาษีนิติบุคคลกี่ปี", [10]),
    (9, "BOI ช่วยเรื่อง work permit ของต่างชาติไหม", [13]),
    (10, "BOI กับ FBA ต่างกันอย่างไร", [16]),
    (11, "ยื่นขอ BOI ต้องเตรียมเอกสารอะไรบ้าง", [15]),
    (12, "กล้อง CCTV ยี่ห้อไหนนิยม", []),
    (13, "ธุรกิจในบัญชีหนึ่ง บัญชีสอง บัญชีสามต่างกันอย่างไร", [2, 3, 4]),
    (14, "ไม่ได้ BOI แล้วอยากทำธุรกิจบริการในบัญชีสาม ต้องทำอย่างไร", [4, 6]),
    (15, "บริษัทไทยที่ต่างชาติถือหุ้น 60% นับเป็นคนต่างด้าวไหม ต้องขอ FBL ไหม", [1, 4]),
    (16, "ธุรกิจนายหน้าต่างชาติทำได้ไหม", [4]),
]

CONTROL_QUESTION_NUM = 12


def minmax_normalize(scores):
    scores = np.array(scores, dtype=float)
    if scores.max() == scores.min():
        return np.zeros_like(scores)
    return (scores - scores.min()) / (scores.max() - scores.min())


def cosine_sim(query_vec, doc_vecs):
    query_vec = np.array(query_vec, dtype=float)
    doc_vecs = np.array(doc_vecs, dtype=float)
    qn = query_vec / np.linalg.norm(query_vec)
    dn = doc_vecs / np.linalg.norm(doc_vecs, axis=1, keepdims=True)
    return dn @ qn


def main():
    if len(sys.argv) < 2:
        print("STOP: no knowledge_base_export.json path given as argv[1]")
        sys.exit(1)

    kb_path = sys.argv[1]
    if not os.path.exists(kb_path):
        print("STOP: file not found")
        sys.exit(1)

    with open(kb_path, encoding="utf-8") as f:
        data = json.load(f)

    chunks = [item["content"] if isinstance(item, dict) else item for item in data]
    n = len(chunks)
    print("num_chunks:", n)

    try:
        embed_model = SentenceTransformer("intfloat/multilingual-e5-large")
        reranker = CrossEncoder("cross-encoder/mmarco-mMiniLMv2-L12-H384-v1")
    except Exception:
        print("STOP: mandatory model load failed")
        sys.exit(1)

    bge_reranker = None
    try:
        bge_reranker = CrossEncoder("BAAI/bge-reranker-v2-m3")
        print("variant_E=available")
    except Exception:
        print("variant_E=skipped")

    os.makedirs("eval_out", exist_ok=True)

    tokenized_kb = [word_tokenize(doc, engine="newmm") for doc in chunks]
    bm25 = BM25Okapi(tokenized_kb)
    emb_chunks = embed_model.encode(chunks)  # no e5 prefix — matches production

    variant_names = ["A", "C", "F", "D"] + (["E"] if bge_reranker is not None else [])
    out_lines = []
    recall_strict = {v: 0 for v in variant_names}
    hit_strict_by_q = {v: {} for v in variant_names}

    scored_questions = [q for q in QUESTIONS if q[0] != CONTROL_QUESTION_NUM]
    num_scored = len(scored_questions)

    for q_num, q_text, necessary in QUESTIONS:
        tokenized_query = word_tokenize(q_text, engine="newmm")
        bm25_scores = np.array(bm25.get_scores(tokenized_query))
        bm25_norm = minmax_normalize(bm25_scores)

        q_emb = embed_model.encode(q_text)  # no e5 prefix
        vec_scores = cosine_sim(q_emb, emb_chunks)
        final_scores = 0.5 * minmax_normalize(vec_scores) + 0.5 * bm25_norm
        top5_idx = list(final_scores.argsort()[::-1][:5])
        candidates = [chunks[i] for i in top5_idx]
        pairs = [[q_text, c] for c in candidates]
        rerank_scores = reranker.predict(pairs)
        ranked = sorted(zip(top5_idx, rerank_scores), key=lambda x: x[1], reverse=True)

        variant_final_idx = {
            "A": [i for i, s in ranked[:3]],
            "C": [i for i, s in ranked],
            "F": list(top5_idx),
            "D": list(top5_idx[:3]),
        }
        if bge_reranker is not None:
            e_scores = bge_reranker.predict(pairs)
            ranked_e = sorted(zip(top5_idx, e_scores), key=lambda x: x[1], reverse=True)
            variant_final_idx["E"] = [i for i, s in ranked_e[:3]]

        necessary_set = set(necessary)
        for v in variant_names:
            final_chunks = [i + 1 for i in variant_final_idx[v]]
            final_set = set(final_chunks)
            hit_strict = 1 if necessary_set.issubset(final_set) else 0
            hit_lenient = 1 if (necessary_set & final_set) else 0

            out_lines.append(
                f"variant={v} q={q_num} "
                f"final={','.join(str(c) for c in final_chunks)} "
                f"hit_strict={hit_strict} hit_lenient={hit_lenient}"
            )

            if q_num != CONTROL_QUESTION_NUM:
                recall_strict[v] += hit_strict
            hit_strict_by_q[v][q_num] = hit_strict

    with open("eval_out/eval_results_v2.txt", "w", encoding="utf-8") as f:
        for line in out_lines:
            f.write(line + "\n")
        f.write("\n")
        for v in variant_names:
            f.write(f"recall_strict_{v}={recall_strict[v]}/{num_scored}\n")

    for v in variant_names:
        print(f"recall_strict_{v}={recall_strict[v]}/{num_scored}")
    for v in variant_names:
        print(f"q1_hit_strict_{v}={hit_strict_by_q[v][1]}")
    for v in variant_names:
        print(f"q13_hit_strict_{v}={hit_strict_by_q[v][13]}")


if __name__ == "__main__":
    main()
