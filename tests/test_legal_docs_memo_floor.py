"""ทดสอบ _filter_grounded_chunks/MEMO_CHUNK_SCORE_FLOOR (แทนที่ OUT_OF_KB_SCORE_THRESHOLD เดิมที่เข้มเกินไปจน
memo ไม่มี [n] อ้างอิงเลย — ดู round2-p0 ข้อ 3) legal_docs.py/legal_retrieval.py import จาก main.py ตรงๆ ตอนบน
ไฟล์ ต้องมี fake sys.modules["main"] ก่อน import เสมอ (ไม่ต่อ DB/โมเดลจริงเลย)"""
import sys
import types
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

fake_main = types.ModuleType("main")
fake_main.client = object()
fake_main.get_chat_messages = lambda chat_id: []
fake_main.save_legal_memo = lambda *a, **k: 1
fake_main.ROUTER_MODEL = "claude-haiku-4-5-20251001"
fake_main.ADVISOR_MODEL = "claude-sonnet-5-5"
fake_main.get_all_tags = lambda: []
fake_main.get_all_chunk_tags_map = lambda: {}
fake_main.get_vector_scores_for_all = lambda query_embedding: {}
fake_main.rerank_with_scores = lambda query, candidates, top_k=3: ([], [])
sys.modules["main"] = fake_main

fake_db = types.ModuleType("db")
fake_db.get_legal_memo = lambda memo_id, user_id: None
fake_db.list_legal_memos_by_chat = lambda chat_id, user_id: []
fake_db.close_legal_memo = lambda memo_id, user_id: False
sys.modules["db"] = fake_db

import legal_docs  # noqa: E402


def test_floor_default_matches_env_default():
    assert legal_docs.MEMO_CHUNK_SCORE_FLOOR == -8.0


def test_filter_drops_scores_below_floor():
    results = [{"id": 176, "score": -8.01}, {"id": 175, "score": -7.99}, {"id": 174, "score": -4.2}]
    kept = legal_docs._filter_grounded_chunks(results, floor=-8.0)
    assert [c["id"] for c in kept] == [175, 174]


def test_filter_keeps_all_when_every_score_is_above_floor():
    results = [{"id": 176, "score": -4.2}, {"id": 175, "score": -3.9}]
    assert legal_docs._filter_grounded_chunks(results, floor=-8.0) == results


def test_filter_drops_all_when_every_score_is_below_floor():
    results = [{"id": 1, "score": -9.0}, {"id": 2, "score": -9.5}]
    assert legal_docs._filter_grounded_chunks(results, floor=-8.0) == []
