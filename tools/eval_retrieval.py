"""
tools/eval_retrieval.py

Offline retrieval-quality evaluation for ThaiMeSook's hybrid search pipeline.
Reimplements the pipeline from main.py from scratch (does NOT import main.py or
db.py, does NOT connect to any database) against a static knowledge_base_export.json
export, to compare 4 candidate embedding/reranking strategies (A-D) without touching
production data or requiring a live DB.

Usage:
    python tools/eval_retrieval.py <path-to-knowledge_base_export.json>

Terminal output is English/ASCII only (numbers, PASS/FAIL, variant labels) per
project convention (this shell mangles Thai text). Full detail (including Thai
question/chunk text is never printed) is written to eval_out/eval_results.txt.

Pipeline reproduced from main.py (read for reference only, not imported):
  - BM25Okapi over word_tokenize(doc, engine="newmm") from pythainlp — identical
    for all 4 variants (BM25 is lexical, unrelated to e5 embedding prefixes).
  - SentenceTransformer('intfloat/multilingual-e5-large') cosine similarity,
    computed with plain numpy (query vs every chunk vector).
  - hybrid: final = 0.5*minmax(vector_scores) + 0.5*minmax(bm25_scores), top 5.
  - CrossEncoder('cross-encoder/mmarco-mMiniLMv2-L12-H384-v1') reranks the hybrid
    top-5 using RAW (unprefixed) query/chunk text — the cross-encoder has no e5
    prefix convention, so all 4 variants rerank identically.

Variants:
  A = current behavior: no e5 prefix on chunk/query embeddings. hybrid top5 -> rerank -> top3.
  B = "passage: "+chunk / "query: "+query embeddings. hybrid top5 -> rerank -> top3.
  C = same embeddings as B, rerank all 5, keep all 5 (no truncation to top3).
  D = same embeddings as B, no reranker at all — final = hybrid top3 directly.

Each result line has one "scores" field, defined as the score of the ranking
method that produced the "final" list: CrossEncoder rerank score for A/B/C,
hybrid combined score for D (D has no rerank step).
"""

import sys
import os
import json

import numpy as np
from rank_bm25 import BM25Okapi
from pythainlp.tokenize import word_tokenize
from sentence_transformers import SentenceTransformer, CrossEncoder

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
]


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
        print("STOP_POINT_B: no knowledge_base_export.json path given as argv[1]")
        sys.exit(1)

    kb_path = sys.argv[1]
    if not os.path.exists(kb_path):
        print("STOP_POINT_B: file not found")
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
        print("STOP_POINT_C: model load failed")
        sys.exit(1)

    os.makedirs("eval_out", exist_ok=True)

    # BM25 index is identical across all 4 variants (lexical, no e5 prefix concept)
    tokenized_kb = [word_tokenize(doc, engine="newmm") for doc in chunks]
    bm25 = BM25Okapi(tokenized_kb)

    # Precompute chunk embeddings once per prefix scheme
    emb_no_prefix = embed_model.encode(chunks)
    emb_passage_prefix = embed_model.encode(["passage: " + c for c in chunks])

    out_lines = []
    recall_hits = {v: 0 for v in "ABCD"}
    q1_hit = {v: 0 for v in "ABCD"}

    for q_num, q_text, correct_chunks in QUESTIONS:
        tokenized_query = word_tokenize(q_text, engine="newmm")
        bm25_scores = np.array(bm25.get_scores(tokenized_query))
        bm25_norm = minmax_normalize(bm25_scores)

        # --- Variant A: no prefix ---
        q_emb_a = embed_model.encode(q_text)
        vec_scores_a = cosine_sim(q_emb_a, emb_no_prefix)
        final_a = 0.5 * minmax_normalize(vec_scores_a) + 0.5 * bm25_norm
        top5_idx_a = final_a.argsort()[::-1][:5]
        hybrid_top5_a = [i + 1 for i in top5_idx_a]
        candidates_a = [chunks[i] for i in top5_idx_a]
        pairs_a = [[q_text, c] for c in candidates_a]
        rerank_scores_a = reranker.predict(pairs_a)
        ranked_a = sorted(zip(top5_idx_a, rerank_scores_a), key=lambda x: x[1], reverse=True)
        final_idx_a = [i for i, s in ranked_a[:3]]
        final_scores_a = [float(s) for i, s in ranked_a[:3]]
        final_chunks_a = [i + 1 for i in final_idx_a]

        # --- Shared hybrid for B/C/D: passage/query prefix ---
        q_emb_prefixed = embed_model.encode("query: " + q_text)
        vec_scores_bcd = cosine_sim(q_emb_prefixed, emb_passage_prefix)
        final_bcd = 0.5 * minmax_normalize(vec_scores_bcd) + 0.5 * bm25_norm
        top5_idx_bcd = final_bcd.argsort()[::-1][:5]
        hybrid_top5_bcd = [i + 1 for i in top5_idx_bcd]
        candidates_bcd = [chunks[i] for i in top5_idx_bcd]
        pairs_bcd = [[q_text, c] for c in candidates_bcd]
        rerank_scores_bcd = reranker.predict(pairs_bcd)
        ranked_bcd = sorted(zip(top5_idx_bcd, rerank_scores_bcd), key=lambda x: x[1], reverse=True)

        # --- Variant B: rerank -> top3 ---
        final_idx_b = [i for i, s in ranked_bcd[:3]]
        final_scores_b = [float(s) for i, s in ranked_bcd[:3]]
        final_chunks_b = [i + 1 for i in final_idx_b]

        # --- Variant C: rerank -> keep all 5 ---
        final_idx_c = [i for i, s in ranked_bcd]
        final_scores_c = [float(s) for i, s in ranked_bcd]
        final_chunks_c = [i + 1 for i in final_idx_c]

        # --- Variant D: no reranker, hybrid top3 directly ---
        top3_idx_d = top5_idx_bcd[:3]
        final_chunks_d = [i + 1 for i in top3_idx_d]
        final_scores_d = [float(final_bcd[i]) for i in top3_idx_d]

        variant_results = {
            "A": (hybrid_top5_a, final_chunks_a, final_scores_a),
            "B": (hybrid_top5_bcd, final_chunks_b, final_scores_b),
            "C": (hybrid_top5_bcd, final_chunks_c, final_scores_c),
            "D": (hybrid_top5_bcd, final_chunks_d, final_scores_d),
        }

        for v in "ABCD":
            hybrid_top5, final_chunks, final_scores = variant_results[v]
            hit = 1 if (set(final_chunks) & set(correct_chunks)) else 0
            line = (
                f"variant={v} q={q_num} "
                f"hybrid_top5={','.join(str(c) for c in hybrid_top5)} "
                f"final={','.join(str(c) for c in final_chunks)} "
                f"scores={','.join(f'{s:.4f}' for s in final_scores)} "
                f"hit={hit}"
            )
            out_lines.append(line)
            if q_num != 12:  # control question excluded from recall
                recall_hits[v] += hit
            if q_num == 1:
                q1_hit[v] = hit

    num_scored_questions = sum(1 for q_num, _, _ in QUESTIONS if q_num != 12)

    with open("eval_out/eval_results.txt", "w", encoding="utf-8") as f:
        for line in out_lines:
            f.write(line + "\n")
        f.write("\n")
        for v in "ABCD":
            f.write(f"recall_final_{v}={recall_hits[v]}/{num_scored_questions}\n")

    print("recall_A=%d/%d" % (recall_hits["A"], num_scored_questions))
    print("recall_B=%d/%d" % (recall_hits["B"], num_scored_questions))
    print("recall_C=%d/%d" % (recall_hits["C"], num_scored_questions))
    print("recall_D=%d/%d" % (recall_hits["D"], num_scored_questions))
    print("q1_hit_A=%d" % q1_hit["A"])
    print("q1_hit_B=%d" % q1_hit["B"])
    print("q1_hit_C=%d" % q1_hit["C"])
    print("q1_hit_D=%d" % q1_hit["D"])


if __name__ == "__main__":
    main()
