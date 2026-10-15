"""Harness ที่ดึงฟังก์ชันจริงจาก main.py/db.py มาทดสอบ โดยไม่ import ทั้งสองไฟล์ตรงๆ (ต้องมี pgvector/
sentence_transformers/DATABASE_URL จริงถึงจะ import main.py ได้ทั้งไฟล์ ซึ่งไม่มีในสภาพแวดล้อมทดสอบ) ใช้ ast
ดึงเฉพาะฟังก์ชัน/ค่าคงที่ที่จำเป็นออกมา exec() ลงใน namespace ที่แยกต่างหาก พร้อม mock ฟังก์ชันที่คุยกับ DB จริง
ทั้งหมดด้วย dict ในหน่วยความจำ (ไม่ต่อ DB จริงเลยตามกติกาของงานนี้)

หมายเหตุ: นี่คือการทดสอบฟังก์ชันจริงจาก source code ปัจจุบันของ main.py/db.py เสมอ (อ่านไฟล์สดทุกครั้งที่รัน
pytest) ไม่ใช่การ mock ตรรกะเอง — ถ้า main.py เปลี่ยนฟังก์ชันที่ทดสอบอยู่ การทดสอบนี้จะพังทันทีเพื่อสะท้อนความจริง
"""
import ast
import io
import os
import time
import re
import html
import json
import zipfile
import difflib
import datetime
import tempfile
import concurrent.futures
import urllib.parse
from pathlib import Path
from zoneinfo import ZoneInfo
from typing import Optional
from xml.sax.saxutils import escape as _xml_escape

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.styles.numbers import is_date_format
from openpyxl.workbook.properties import CalcProperties
from pycel import ExcelCompiler
from pythainlp.tokenize import word_tokenize
from rank_bm25 import BM25Okapi

MAIN_PY = Path(__file__).parent.parent / "main.py"
DB_PY = Path(__file__).parent.parent / "db.py"
DOCUMENT_FIELDS_PY = Path(__file__).parent.parent / "document_fields.py"
LLM_TOOLS_PY = Path(__file__).parent.parent / "llm_tools.py"
ANSWER_GUARD_PY = Path(__file__).parent.parent / "answer_guard.py"
COMPOUND_EDITS_PY = Path(__file__).parent.parent / "compound_edits.py"

WANTED_MAIN_FUNCS = {
    "_json_safe_cell_value", "_parse_iso_datetime_like",
    "_normalize_table_header_text", "_merged_cell_top_left", "_is_item_table_row_blank", "_detect_item_table",
    "_is_formula_cell", "_clean_multi_pair_label_text", "_match_multi_pair_row", "_extract_excel_labels",
    "_coerce_value_for_cell", "_load_workbook_preserving_macros",
    "_cell_values_equal", "_filter_labels_actually_edited", "_escape_table_cell", "_cell_ref",
    "_recalculate_formula_values", "_build_document_version", "get_working_copy", "create_new_version",
    "_compute_table_amount_values", "read_document_values", "_format_document_values_text", "_is_label_cell_formula",
    "_suggest_similar_labels", "_normalize_label",
    "_format_rate_number", "_rewrite_percent_text", "_rewrite_rate_formula", "_resolve_rate_formula",
    "_format_display_value",
    "_split_rate_label_base_and_percent", "_format_money_value",
    "_build_rate_change_mutate_fn", "_label_side_effects", "_format_side_effect_lines",
    "_get_single_item_table",
    "_build_delete_table_item_mutate_fn", "_describe_table_item",
    "_build_add_table_item_mutate_fn",
    "_table_has_non_formula_amount_value",
    "_resolve_and_build_mutate_fn", "_replay_pending_queue",
    "apply_change_rate", "apply_update_table_item", "apply_delete_table_item", "apply_add_table_item",
    "_apply_single_queue_item", "_document_totals_lines", "_apply_structural_commands",
    "undo_last_edit", "_is_undo_command",
    "_is_document_reading_question", "_has_legal_keyword", "_answer_document_reading_question",
    "_is_reading_type_question", "_resolve_document_reading_dispatch",
    "_query_mentions_table_item_or_label",
    "_is_structural_edit_command",
    "_find_single_rate_label",
    "_match_structural_clause", "_resolve_table_target", "_normalize_structural_target_text",
    "_combine_description_replacement", "_parse_structural_number", "_parse_direct_structural_commands",
    "_resolve_direct_structural_dispatch",
    "_has_explicit_finalize_word", "_classify_pending_queue_command",
    "_normalize_pending_queue_command_text", "_mentions_pending_queue_intent",
    "_get_document_version_marker",
    "_label_position", "_find_label_by_position", "_current_label_display_name",
    "_get_active_chat_document", "execute_open_library_file", "_normalize_filename_for_matching",
    "_match_and_apply_excel_edit", "_parse_json_response",
    "rag_answer", "contains_unexpected_script", "_normalize_polite_particles",
    "advisor_answer", "_advisor_system_prompt", "_advisor_offer_line",
    "_advisor_target_doc", "_advisor_document_context",
    "run_agentic_tool_loop", "_build_final_answer_with_citations", "_safe_execute_tool", "_build_action_receipt",
    "execute_present_recommendations", "_render_recommendation_item_line", "_render_recommendations_answer",
    "_extract_domain_from_url", "_safe_markdown_url",
    "_resolve_recommendation_command_dispatch", "_resolve_recommendation_replace_command",
    "_resolve_recommendation_add_command", "_parse_recommendation_price_clause",
    "_apply_recommendation_replace", "_apply_recommendation_add",
    "_resolve_recommendation_intent_with_llm", "_resolve_recommendation_intent_price",
    "_validate_and_apply_recommendation_action",
    "_normalize_recommendation_intent_text", "_query_has_recommendation_intent_verb",
    "_sanitize_recommendation_text",
    "_classify_question_route", "_classify_legal_or_advisor_route",
    "_query_matches_existing_customer", "_looks_like_library_query", "_resolve_library_route_access",
    "_query_matches_document_code_context", "_normalize_customer_name", "_normalize_library_search_text",
    "_categorize_anthropic_api_error", "_ask_question_api_error_response",
    "_get_enriched_library_files", "_library_file_has_macros", "_guess_library_document_type",
    "build_library_search_card", "save_library_search_card", "backfill_library_search_cards",
    "_library_search_exact_match_fields", "_library_search_matches_filters", "_library_search_matched_on_text",
    "_library_search_matches_contains_item",
    "_score_library_search_candidates", "_library_search_build_candidates", "search_library", "execute_search_library",
    "_resolve_library_exact_match_dispatch",
    "_order_library_documents_by_folder", "_order_library_documents_by_key", "_order_library_documents_by_customer",
    "_resolve_flat_document_labels", "_format_document_line_flat",
    "_format_library_document_list", "_record_chat_document_listing", "_escape_markdown_text",
    "_strip_file_extension", "_format_customer_list", "_record_chat_customer_listing",
    "_looks_like_full_library_listing_query", "_resolve_library_full_listing_dispatch",
    "_read_values_from_bytes", "_tokenize_library_query",
    "_is_chat_export_command", "_build_chat_export_answer",
    "_format_chat_export_timestamp", "_build_chat_export_text", "_strip_markdown_for_export",
    "_thai_group_to_text", "_thai_integer_to_text", "_bahttext", "_try_evaluate_bahttext",
    "_format_cell_display_text", "_cell_style_info", "_evaluate_preview_formula_cells",
    "_list_preview_sheet_summaries", "_render_sheet_window", "_merge_preview_highlight",
    "_resolve_preview_mode_bytes", "_preview_cache_get", "_preview_cache_set",
    "_build_router_document_context", "_format_router_editable_fields_text",
    "_build_router_recommendations_context", "_build_router_history_context",
    "_resolve_router_target_document", "route_intent_with_llm",
    "_dispatch_router_edit_document", "_dispatch_router_recommendation_action", "_dispatch_router_advisor",
    "_dispatch_router_legal", "_dispatch_router_library_search", "_dispatch_router_clarification",
    "_is_router_clarification_fresh",
    "_resolve_llm_router_dispatch",
    "_normalize_thai_typos", "_query_has_edit_verb", "_force_structural_edit_dispatch",
    "_query_references_open_file", "_resolve_fallback_edit_intent", "_query_has_edit_command_form",
}

WANTED_MAIN_CONSTS = {
    "EXCEL_LABEL_SECTION_HEADING_LOOKBACK", "_ITEM_TABLE_HEADER_WORDS", "_ITEM_TABLE_COLUMN_MAP",
    "_ITEM_TABLE_STOP_PREFIXES", "_ITEM_TABLE_MAX_ROWS", "_ITEM_TABLE_EMPTY_STREAK_LIMIT",
    "FORMULA_RECALC_TIMEOUT_SECONDS", "_RATE_FORMULA_PATTERNS", "_PERCENT_IN_TEXT_RE", "_TABLE_INPUT_FIELDS",
    "_DOCUMENT_REFERENCE_KEYWORDS", "_DOCUMENT_EDIT_COMMAND_WORDS", "_STRUCTURAL_ITEM_NO_RE",
    "_MIN_TABLE_OR_LABEL_MATCH_LENGTH",
    "_LEGAL_KNOWLEDGE_KEYWORDS", "_FINALIZE_EXPLICIT_WORDS",
    "_CONFIRM_ALL_TEXTS", "_CONFIRM_LATEST_TEXTS", "_CONFIRM_SINGLE_TEXTS",
    "_CANCEL_LATEST_TEXTS", "_CANCEL_ALL_TEXTS",
    "_PENDING_QUEUE_COMMAND_TRAILING_WORDS", "_PENDING_QUEUE_NEGATION_RE",
    "_OPEN_COMMAND_QUESTION_WORDS",
    "_TABLE_FIELD_THAI_NAMES", "_UNDO_HINT_LINE", "_UNDO_COMMAND_TEXTS", "_UNDO_COMMAND_MAX_TOKENS",
    "_STRUCTURAL_DIRECT_DELETE_RE", "_STRUCTURAL_DIRECT_RATE_RE",
    "_STRUCTURAL_CLAUSE_SPLIT_RE", "_STRUCTURAL_ITEM_NO_TARGET_RE",
    "_STRUCTURAL_DIRECT_FIELD_RE", "_STRUCTURAL_FIELD_NAME_MAP", "_STRUCTURAL_DIRECT_DESCRIPTION_RE",
    "_LABEL_TRAILING_PUNCTUATION", "_TABLE_NON_FORMULA_AMOUNT_ERROR",
    "EXCEL_EDITOR_SYSTEM_PROMPT", "FORMULA_EDIT_WARNING_NOTICE",
    "_STRUCTURAL_NO_TOOL_FALLBACK_MESSAGE",
    "_FABRICATED_EDIT_CLAIM_RE", "_FABRICATED_EDIT_CLAIM_STRUCTURAL_ONLY_PATTERNS",
    "_FABRICATED_EDIT_CLAIM_GENERAL_PATTERNS", "_RECEIPT_LINE_RE",
    "UNEXPECTED_SCRIPT_PATTERN", "MAX_ANSWER_RETRIES", "MAX_HISTORY_MESSAGES", "MAX_TOOL_ITERATIONS",
    "_MAX_ANSWER_CITATIONS",
    "_ADVISOR_REPLACEMENT_OFFER_LINE", "ADVISOR_TOOLS", "GENERAL_WEB_SEARCH_TOOL",
    "PRESENT_RECOMMENDATIONS_TOOL", "ADVISOR_MODEL", "_NO_VERIFIED_RECOMMENDATIONS_MESSAGE",
    "_SEARCH_TOOL_BLOCK_TYPES", "ADVISOR_MAX_TOOL_ITERATIONS", "MAX_PAUSE_TURNS", "_URL_SAFE_CHARACTERS",
    "ADVISOR_MAX_TOKENS",
    "_RECOMMENDATION_REPLACE_RE", "_RECOMMENDATION_ADD_RE", "_RECOMMENDATION_QTY_CLAUSE_RE",
    "_PRICE_CLAUSE_EXPLICIT_RE", "_PRICE_CLAUSE_ORIGINAL_RE", "_PRICE_CLAUSE_MARKUP_RE",
    "QUESTION_DOMAIN_CLASSIFIER_SYSTEM_PROMPT",
    "_CHAT_EXPORT_ROLE_LABELS", "_CHAT_EXPORT_COMMAND_TEXTS",
    "_MARKDOWN_LINK_RE", "_MARKDOWN_BOLD_RE", "_MARKDOWN_HEADING_RE",
    "_POLITE_PARTICLE_COMPOUND_RE", "_POLITE_PARTICLE_TRAILING_RE",
    "_THAI_DIGIT_WORDS", "_THAI_PLACE_WORDS",
    "_BAHTTEXT_FORMULA_RE", "_SIMPLE_CELL_REFERENCE_RE", "_PREVIEW_NOT_COMPUTABLE_TITLE",
    "_PREVIEW_MAX_COLUMNS", "_PREVIEW_MAX_ROWS_PER_PAGE", "_PREVIEW_DEFAULT_MAX_ROWS",
    "_PREVIEW_CACHE", "_PREVIEW_CACHE_MAX_ENTRIES",
    "RESOLVE_RECOMMENDATION_ACTION_TOOL", "_RECOMMENDATION_INTENT_SYSTEM_PROMPT", "_RECOMMENDATION_INTENT_VERBS",
    "_CITE_TAG_RE", "_TRAILING_TEXT_ARTIFACT_RE", "_REPEATED_WHITESPACE_RE",
    "_LIBRARY_ROUTE_KEYWORDS", "_LIBRARY_DOCUMENT_CODE_RE", "_LIBRARY_LOGIN_REQUIRED_MESSAGE",
    "_LIBRARY_DOCUMENT_CODE_TRAILING_WORDS", "_LIBRARY_DOCUMENT_REFERENCE_WORDS",
    "_CUSTOMER_NAME_NOISE_WORDS", "_THAI_TONE_MARKS_RE",
    "SEARCH_LIBRARY_TOOL", "LIBRARY_ROUTE_SYSTEM_PROMPT", "LIBRARY_TOOLS", "LIBRARY_LOGGED_IN_RULES",
    "_LIBRARY_DOCUMENT_TYPE_WORDS", "_LIBRARY_CARD_DATE_LABEL_WORDS", "_LIBRARY_CARD_TOTAL_LABEL_WORDS",
    "_LIBRARY_SEARCH_MAX_RESULTS", "_LIBRARY_SEARCH_SCORE_THRESHOLD", "_LIBRARY_SEARCH_EXACT_MATCH_BONUS",
    "_LIBRARY_SEARCH_RELATIVE_CUTOFF",
    "_LIBRARY_QUERY_STOPWORDS", "_LIBRARY_QUERY_STOPWORDS_NORMALIZED",
    "_LIBRARY_FULL_LISTING_KEYWORDS", "_LIBRARY_FULL_LISTING_SIZE_CUTOFF",
    "ROUTER_MODE", "ROUTER_MODEL", "ROUTE_INTENT_TOOL", "ROUTER_SYSTEM_PROMPT",
    "_ROUTER_FILE_INTENTS", "_ROUTER_VALID_INTENTS", "_ROUTER_FILE_LOGIN_REQUIRED_MESSAGE",
    "_ROUTER_CLARIFICATION_THIRD_PERSON_FALLBACK",
    "_ROUTER_EDITABLE_FIELDS_MAX", "_ROUTER_EDITABLE_FIELD_VALUE_MAX_CHARS",
    "_EDIT_INTENT_VERB_WORDS", "_EDIT_INTENT_FILE_REFERENCE_WORDS", "_EDIT_INTENT_QUESTION_ENDINGS",
}

# _looks_like_fabricated_edit_claim ไม่ได้อยู่ใน WANTED_MAIN_FUNCS ด้านบนโดยตั้งใจ — ไม่ต้องใช้ใน test ชุดนี้
# โดยตรง (ทดสอบผ่าน rag_answer ซึ่งหนักเกินจะ extract ทั้งฟังก์ชัน) แต่ยังปล่อยให้ดึงได้ถ้าจำเป็นภายหลัง
WANTED_MAIN_FUNCS.add("_looks_like_fabricated_edit_claim")
WANTED_MAIN_FUNCS.add("_strip_fabricated_receipt_lines")

WANTED_DB_FUNCS = {"_is_macro_workbook"}

# feat/edit-any-field: document_fields.py ดึงทั้งไฟล์อัตโนมัติด้วย _extract_all แทนไล่ชื่อทีละตัว (ดูเหตุผลใน
# docstring ของ _extract_all) exec ลง ns เดียวกับ main.py เพราะฟังก์ชันในไฟล์นั้น bare-reference ชื่อจาก main.py
# (get_working_copy/create_new_version/ฯลฯ) ตรงๆ โดยไม่มี "document_fields."/"main." นำหน้า


class _FakeDb:
    """เก็บสถานะจำลองของ EditableDocument/DocumentVersion/ChatPendingAction/DocumentEditLog ในหน่วยความจำล้วนๆ
    (ไม่ต่อ DB จริงเลย) — เลียนแบบ semantics ของ db.py เท่าที่ฟังก์ชันที่ทดสอบต้องใช้จริงเท่านั้น"""

    def __init__(self):
        self.documents: dict[int, dict] = {}
        self.versions: dict[int, list[dict]] = {}
        self.edit_logs: dict[int, dict] = {}
        self.pending_actions: dict[int, list[dict]] = {}  # chat_id -> [{"id","action","preview_text","created_at"}, ...] เรียงเก่า->ใหม่
        self.focus: dict[int, int] = {}
        self.last_recommendations: dict[int, list] = {}
        self.closed_ids: set[int] = set()
        self.library_files: dict[int, dict] = {}
        self.library_file_meta: dict[int, dict] = {}
        self.file_categories: dict[int, str] = {}
        self.library_search_cards: dict[int, dict] = {}
        self.last_listing: dict[int, dict] = {}  # chat_id -> {"kind":, "items":} จาก set_chat_last_listing
        self.router_clarifications: dict[int, dict] = {}  # chat_id -> {"message_count_at_creation": int}
        self.chat_messages: dict[int, list[dict]] = {}  # chat_id -> [{"role","content"}, ...] เรียงเก่า->ใหม่
        self._next_document_id = 1
        self._next_version_id = 1
        self._next_edit_log_id = 1
        self._next_pending_id = 1

    def add_document(self, document_id: int, filename: str, original_bytes: bytes, label_map: dict,
                     chat_id: int = 1, user_id: int = 1) -> None:
        self.documents[document_id] = {
            "id": document_id, "filename": filename, "original_bytes": original_bytes,
            "label_map": label_map, "chat_id": chat_id, "user_id": user_id,
        }

    def get_editable_document_raw(self, document_id):
        return self.documents.get(document_id)

    def get_editable_document(self, document_id, user_id):
        doc = self.documents.get(document_id)
        if doc and doc["user_id"] == user_id:
            return doc
        return None

    def list_active_editable_documents_by_chat(self, chat_id, user_id):
        return [
            doc for doc in self.documents.values()
            if doc["chat_id"] == chat_id and doc["user_id"] == user_id and doc["id"] not in self.closed_ids
        ]

    def close_document(self, document_id: int) -> None:
        self.closed_ids.add(document_id)

    def create_editable_document(self, user_id, filename, original_bytes, label_map, chat_id):
        document_id = self._next_document_id
        self._next_document_id += 1
        self.add_document(document_id, filename, original_bytes, label_map, chat_id=chat_id, user_id=user_id)
        return document_id

    def get_latest_document_version(self, document_id):
        versions = self.versions.get(document_id, [])
        if not versions:
            return None
        v = versions[-1]
        return {"id": v["id"], "file_bytes": v["file_bytes"], "label_map": v["label_map"]}

    def get_second_latest_document_version(self, document_id):
        versions = self.versions.get(document_id, [])
        if len(versions) < 2:
            return None
        v = versions[-2]
        return {"id": v["id"], "file_bytes": v["file_bytes"], "label_map": v["label_map"]}

    def create_document_version(self, document_id, file_bytes, label_map, change_summary=None):
        vid = self._next_version_id
        self._next_version_id += 1
        self.versions.setdefault(document_id, []).append({
            "id": vid, "file_bytes": file_bytes, "label_map": label_map, "change_summary": change_summary,
        })
        return vid

    def delete_latest_document_version(self, document_id):
        versions = self.versions.get(document_id, [])
        if not versions:
            return None
        deleted = versions.pop()
        return {"version_no": len(versions) + 1, "label_map": deleted["label_map"]}

    def link_edit_log_to_version(self, edit_log_id, version_id):
        if edit_log_id in self.edit_logs:
            self.edit_logs[edit_log_id]["version_id"] = version_id

    def log_document_edit(self, document_id, label, old_value, new_value, was_formula=False):
        eid = self._next_edit_log_id
        self._next_edit_log_id += 1
        self.edit_logs[eid] = {
            "id": eid, "document_id": document_id, "label": label,
            "old_value": old_value, "new_value": new_value, "was_formula": was_formula,
        }
        return eid

    def set_chat_document_focus(self, chat_id, document_id):
        self.focus[chat_id] = document_id

    def set_chat_last_recommendations(self, chat_id, items):
        self.last_recommendations[chat_id] = list(items)

    def get_chat_last_recommendations(self, chat_id):
        return self.last_recommendations.get(chat_id)

    def set_chat_router_clarification_flag(self, chat_id):
        count = len(self.chat_messages.get(chat_id, []))
        self.router_clarifications[chat_id] = {"message_count_at_creation": count}

    def get_chat_router_clarification_flag(self, chat_id):
        return self.router_clarifications.get(chat_id)

    def get_chat_document_focus(self, chat_id, user_id):
        document_id = self.focus.get(chat_id)
        if document_id is None:
            return None
        doc = self.documents.get(document_id)
        if doc and doc["user_id"] == user_id:
            return doc
        return None

    def get_chat_messages(self, chat_id):
        """ตัวแทน db.get_chat_messages — เทสต์ seed ประวัติแชทโดยเขียนตรงเข้า self.chat_messages[chat_id] เอง
        (ดู _resolve_recommendation_intent_with_llm/_answer_document_reading_question ที่เรียกใช้ฟังก์ชันนี้)"""
        return [dict(m) for m in self.chat_messages.get(chat_id, [])]

    def set_chat_last_listing(self, chat_id, kind, items):
        self.last_listing[chat_id] = {"kind": kind, "items": items}

    def get_chat_last_listing(self, chat_id):
        listing = self.last_listing.get(chat_id)
        return dict(listing) if listing is not None else None

    def add_library_file(
        self, file_id: int, filename: str, storage_key: str, label_map: Optional[dict] = None,
        category_id: Optional[int] = None, summary: Optional[str] = None,
    ) -> None:
        """ตัวแทน library_files 1 แถว — key เป็น "id" (ไม่ใช่ "file_id") ให้ตรงกับรูปแบบจริงที่ db.py คืน (ดู
        _library_file_to_dict) เพราะ _get_enriched_library_files/build_library_search_card อ่าน f["id"] ตรงๆ"""
        self.library_files[file_id] = {
            "id": file_id, "filename": filename, "category_id": category_id, "storage_key": storage_key,
            "label_map": label_map or {}, "summary": summary,
            "uploaded_at": datetime.datetime.utcnow(), "updated_at": datetime.datetime.utcnow(),
        }

    def get_library_file(self, file_id):
        return self.library_files.get(file_id)

    def get_library_files(self, category_id=None, only_uncategorized=False):
        return [dict(f) for f in self.library_files.values()]

    def add_file_category(self, category_id: int, name: str) -> None:
        self.file_categories[category_id] = name

    def get_file_categories_with_counts(self):
        return [{"id": cid, "name": name} for cid, name in self.file_categories.items()]

    def add_library_file_meta(
        self, file_id: int, customer_code: Optional[str] = None, customer_name: Optional[str] = None,
        project_code: Optional[str] = None, project_name: Optional[str] = None,
        document_no: Optional[str] = None, is_template: bool = False,
    ) -> None:
        """ตัวแทน library_file_meta 1 แถว — ให้เทสต์ตั้งค่าข้อมูลระบุตัวตนของไฟล์ตรงๆ (ดู get_library_file_meta/
        get_all_library_file_meta ที่ _get_enriched_library_files ใช้)"""
        self.library_file_meta[file_id] = {
            "customer_code": customer_code, "customer_name": customer_name, "project_code": project_code,
            "project_name": project_name, "document_no": document_no, "is_template": is_template,
        }

    def get_library_file_meta(self, file_id):
        meta = self.library_file_meta.get(file_id)
        return dict(meta) if meta is not None else None

    def get_all_library_file_meta(self):
        return {fid: dict(meta) for fid, meta in self.library_file_meta.items()}

    def add_library_search_card(
        self, file_id: int, card_text: str, total_amount: Optional[float] = None,
        issue_date: Optional[str] = None, embedding: Optional[list] = None,
    ) -> None:
        """ตัวแทน library_search_index 1 แถว — ให้เทสต์ seed "บัตรค้นหา" ตรงๆ โดยไม่ต้องผ่าน
        build_library_search_card จริง (ซึ่งต้องมีไฟล์ Excel จริงให้อ่าน) ใช้ทดสอบ search_library/
        _score_library_search_candidates แยกจากการทดสอบ build_library_search_card เอง"""
        self.upsert_library_search_card(file_id, card_text, embedding, total_amount, issue_date)

    def upsert_library_search_card(self, file_id, card_text, embedding, total_amount=None, issue_date=None):
        if file_id not in self.library_files:
            return False
        self.library_search_cards[file_id] = {
            "card_text": card_text, "total_amount": total_amount, "issue_date": issue_date,
            "has_embedding": embedding is not None,
        }
        return True

    def get_all_library_search_cards(self):
        return {fid: dict(card) for fid, card in self.library_search_cards.items()}

    def get_library_file_ids_without_search_card(self):
        return sorted(set(self.library_files) - set(self.library_search_cards))

    def add_chat_pending_action(self, chat_id, action, preview_text):
        pending_id = self._next_pending_id
        self._next_pending_id += 1
        row = {
            "id": pending_id, "action": action, "preview_text": preview_text,
            "created_at": datetime.datetime.utcnow(),
        }
        self.pending_actions.setdefault(chat_id, []).append(row)
        return pending_id

    def get_chat_pending_actions(self, chat_id):
        return [dict(row) for row in self.pending_actions.get(chat_id, [])]

    def delete_chat_pending_action(self, pending_id):
        for chat_id, rows in self.pending_actions.items():
            self.pending_actions[chat_id] = [r for r in rows if r["id"] != pending_id]

    def clear_chat_pending_actions(self, chat_id):
        self.pending_actions.pop(chat_id, None)

    def update_chat_pending_action_version(self, pending_id, version_marker):
        for rows in self.pending_actions.values():
            for row in rows:
                if row["id"] == pending_id:
                    row["action"] = dict(row["action"])
                    row["action"]["version_id_at_propose"] = version_marker

    def add_storage_blob(self, storage_key: str, value) -> None:
        """value: bytes (คืนตรงๆ) หรือ Exception instance (raise เมื่อถูกเรียก) — จำลอง storage_get จริง"""
        self.storage_blobs = getattr(self, "storage_blobs", {})
        self.storage_blobs[storage_key] = value

    def storage_get(self, storage_key):
        value = getattr(self, "storage_blobs", {}).get(storage_key)
        if isinstance(value, Exception):
            raise value
        return value


# ตัวแทน module "anthropic" จริง (ไม่ได้ติดตั้งในสภาพแวดล้อมทดสอบนี้ — ดู requirements.txt) เท่าที่
# _categorize_anthropic_api_error/_ask_question_api_error_response ใน main.py ต้องใช้ (isinstance check +
# except anthropic.APIError) จำลองลำดับชั้นจริงของ SDK เท่าที่จำเป็น: APIError เป็นฐานของทุกตัว, APIStatusError
# (มี .status_code) เป็นฐานของ BadRequestError/RateLimitError/InternalServerError
class _FakeAnthropicAPIError(Exception):
    pass


class _FakeAnthropicAPIStatusError(_FakeAnthropicAPIError):
    def __init__(self, message: str, status_code: Optional[int] = None):
        super().__init__(message)
        self.status_code = status_code


class _FakeAnthropicBadRequestError(_FakeAnthropicAPIStatusError):
    pass


class _FakeAnthropicRateLimitError(_FakeAnthropicAPIStatusError):
    pass


class _FakeAnthropicInternalServerError(_FakeAnthropicAPIStatusError):
    pass


class _FakeAnthropicModule:
    APIError = _FakeAnthropicAPIError
    APIStatusError = _FakeAnthropicAPIStatusError
    BadRequestError = _FakeAnthropicBadRequestError
    RateLimitError = _FakeAnthropicRateLimitError
    InternalServerError = _FakeAnthropicInternalServerError


class _FakeAnthropicContentBlock:
    def __init__(self, text: str):
        self.text = text


class _FakeAnthropicResponse:
    def __init__(self, text: str):
        self.content = [_FakeAnthropicContentBlock(text)]


class _FakeAnthropicMessages:
    def __init__(self, client: "_FakeAnthropicClient"):
        self._client = client

    def create(self, **kwargs):
        self._client.last_call_kwargs = kwargs
        self._client.call_count += 1
        if self._client.response_queue:
            # response_queue (ถ้ามีของเหลืออยู่) มาก่อน raise_on_create เสมอ — ให้เทสต์คุม control flow ของ
            # run_agentic_tool_loop ตัวจริงได้ (เช่น รอบปกติสำเร็จก่อน แล้วค่อยให้การเรียกพิเศษ "หลัง" loop ปกติ
            # ล้มเหลว โดยตั้ง raise_on_create ไว้ด้วยพร้อมกัน — คิวว่างเมื่อไหร่ค่อยเริ่มใช้ raise_on_create)
            return self._client.response_queue.pop(0)
        if self._client.raise_on_create is not None:
            raise self._client.raise_on_create
        return _FakeAnthropicResponse(self._client.response_text)


class _FakeAnthropicClient:
    """ตัวแทน anthropic.Client — คืนข้อความคงที่ (ตั้งค่าได้) เสมอ ไม่เรียก API จริงเลย เก็บ kwargs ของการเรียก
    ล่าสุดไว้ด้วย (last_call_kwargs) ให้เทสต์ตรวจสอบได้ว่าส่ง system/messages อะไรเข้าไป ตั้ง raise_on_create
    เป็น Exception instance เพื่อจำลอง API ล้ม/timeout ได้ response_queue: ถ้าใส่ไว้ (ดู _FakeAgenticResponse —
    ใช้เฉพาะเทสต์ run_agentic_tool_loop ตัวจริงที่ต้องควบคุม stop_reason/content หลายรอบ) ให้ pop ออกมาคืนทีละ
    รายการตามลำดับที่เรียก create() แทนที่จะคืน response_text คงที่ซ้ำทุกครั้ง — ว่างเปล่า (ค่าเริ่มต้น) ไม่กระทบ
    เทสต์เดิมที่ใช้ response_text เลยสักจุด"""

    def __init__(self, response_text: str = "คำตอบจำลองจาก LLM"):
        self.response_text = response_text
        self.last_call_kwargs = None
        self.raise_on_create: Exception | None = None
        self.messages = _FakeAnthropicMessages(self)
        self.response_queue: list = []
        self.call_count = 0


class _FakeCitation:
    """ตัวแทน citation ที่ Anthropic แนบมากับ text block ตอนอ้างอิงผลค้นจาก web_search (url/title) — ใช้เฉพาะ
    เทสต์ _build_final_answer_with_citations/run_agentic_tool_loop ตัวจริงเท่านั้น"""

    def __init__(self, url: str, title: str):
        self.url = url
        self.title = title

    def model_dump(self):
        return {"url": self.url, "title": self.title}


class _FakeAgenticTextBlock:
    """ตัวแทน content block ชนิด "text" ของ Claude พร้อม .citations ที่ควบคุมได้ (ดู _FakeCitation) — ต่างจาก
    _FakeAnthropicContentBlock ธรรมดาด้านบน (ใช้กับ response.content[0].text.strip() แบบ classifier อื่นๆ
    เท่านั้น ไม่มี .type/.model_dump()/.citations) ใช้เฉพาะเทสต์ run_agentic_tool_loop ตัวจริง"""

    def __init__(self, text: str, citations: Optional[list] = None):
        self.type = "text"
        self.text = text
        self.citations = citations or []

    def model_dump(self):
        return {"type": "text", "text": self.text, "citations": [c.model_dump() for c in self.citations]}


class _FakeAgenticResponse:
    """ตัวแทน response ของ client.messages.create() ที่ควบคุม content blocks + stop_reason ได้ตรงๆ จากเทสต์ —
    ใช้เฉพาะเทสต์ run_agentic_tool_loop ตัวจริง (ดู response_queue ของ _FakeAnthropicClient)"""

    def __init__(self, content_blocks: list, stop_reason: str = "end_turn"):
        self.content = content_blocks
        self.stop_reason = stop_reason


class _FakeAgenticOtherBlock:
    """ตัวแทน content block ชนิดอื่นที่ไม่ใช่ text/tool_use (เช่น server_tool_use/web_search_tool_result) — ใช้
    เฉพาะเทสต์ _build_final_answer_with_citations/run_agentic_tool_loop ตัวจริง ไม่ต้องมี field ครบเหมือนจริง
    content: list ของผลลัพธ์ web_search (แต่ละชิ้นมี "url"/"title") ใช้เฉพาะตอน block_type ==
    "web_search_tool_result" เท่านั้น"""

    def __init__(self, block_type: str, content: Optional[list] = None):
        self.type = block_type
        self.content = content

    def model_dump(self):
        return {"type": self.type, "content": self.content}


class _FakeAgenticToolUseBlock:
    """ตัวแทน content block ชนิด "tool_use" — ใช้เฉพาะเทสต์ run_agentic_tool_loop ตัวจริงที่ต้องจำลอง Claude
    เรียก custom tool (เช่น present_recommendations)"""

    def __init__(self, tool_use_id: str, name: str, tool_input: dict):
        self.type = "tool_use"
        self.id = tool_use_id
        self.name = name
        self.input = tool_input

    def model_dump(self):
        return {"type": "tool_use", "id": self.id, "name": self.name, "input": self.input}


class StorageNotConfiguredError(Exception):
    pass


def _extract(tree: ast.Module, wanted_funcs: set, wanted_consts: set, ns: dict, label: str) -> None:
    found_funcs, found_consts = set(), set()
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name in wanted_funcs:
            exec(compile(ast.Module(body=[node], type_ignores=[]), f"<{label}>", "exec"), ns)
            found_funcs.add(node.name)
        elif isinstance(node, ast.Assign) and len(node.targets) == 1 and isinstance(node.targets[0], ast.Name):
            if node.targets[0].id in wanted_consts:
                exec(compile(ast.Module(body=[node], type_ignores=[]), f"<{label}>", "exec"), ns)
                found_consts.add(node.targets[0].id)
        elif isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name) and node.value is not None:
            # ค่าคงที่ module-level ที่มี type annotation กำกับ (เช่น "_PREVIEW_CACHE: dict[tuple, dict] = {}")
            # — ast.Assign ธรรมดาไม่ครอบคลุม ต้องดึง ast.AnnAssign แยกต่างหาก (exec แค่ชื่อ = ค่า ทิ้ง annotation
            # ไปเลย ไม่จำเป็นต้องมีอยู่ใน namespace ของเทสต์)
            if node.target.id in wanted_consts:
                plain_assign = ast.Assign(targets=[node.target], value=node.value)
                ast.copy_location(plain_assign, node)
                ast.fix_missing_locations(plain_assign)
                exec(compile(ast.Module(body=[plain_assign], type_ignores=[]), f"<{label}>", "exec"), ns)
                found_consts.add(node.target.id)

    missing = (wanted_funcs - found_funcs) | (wanted_consts - found_consts)
    if missing:
        raise RuntimeError(f"ดึงฟังก์ชัน/ค่าคงที่จาก {label} ไม่ครบ: {sorted(missing)}")


def _extract_all(tree: ast.Module, ns: dict, label: str) -> None:
    """เหมือน _extract แต่ดึงฟังก์ชัน/ค่าคงที่ระดับบนสุด "ทุกตัว" ของไฟล์อัตโนมัติ แทนการไล่ระบุชื่อทีละตัว — ใช้กับ
    ไฟล์เล็กที่ทุกอย่างในนั้นเกี่ยวข้องกันหมด (เช่น document_fields.py) กันบั๊กที่เคยเกิด: ดึงแค่ฟังก์ชัน public
    (build_editable_fields/apply_field_edits) แต่ลืมฟังก์ชันช่วยภายใน (_is_rate_formula_text ฯลฯ) ที่สองตัวนั้น
    เรียกใช้ ทำให้ NameError ตอนเทสต์เรียกจริง — เขียนแยกจาก _extract เอง (ไม่เรียกต่อ) เพราะ _extract เช็ค "missing"
    แยกชุดฟังก์ชัน/ค่าคงที่ ถ้าส่งชื่อเดียวกันเข้าไปเป็นทั้งสองชุดจะเข้าใจผิดว่าฟังก์ชันหายไปจากชุดค่าคงที่เสมอ"""
    for node in tree.body:
        if isinstance(node, ast.FunctionDef):
            exec(compile(ast.Module(body=[node], type_ignores=[]), f"<{label}>", "exec"), ns)
        elif isinstance(node, ast.Assign) and len(node.targets) == 1 and isinstance(node.targets[0], ast.Name):
            exec(compile(ast.Module(body=[node], type_ignores=[]), f"<{label}>", "exec"), ns)
        elif isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name) and node.value is not None:
            plain_assign = ast.Assign(targets=[node.target], value=node.value)
            ast.copy_location(plain_assign, node)
            ast.fix_missing_locations(plain_assign)
            exec(compile(ast.Module(body=[plain_assign], type_ignores=[]), f"<{label}>", "exec"), ns)


class _FakeAgenticLoop:
    """ตัวแทน run_agentic_tool_loop ที่ควบคุมผลลัพธ์ได้ตรงๆ จากเทสต์ (ตั้งค่าผ่าน set_response ก่อนเรียก
    rag_answer) — ใช้ทดสอบ guard logic ของ rag_answer เอง (structural_no_tool/fabricated_edit_claim_blocked)
    แบบแยกส่วน โดยไม่ต้องเรียก Claude จริงหรือประกอบ tools list เต็มๆ (หนักเกินไปสำหรับเทสต์นี้)"""

    def __init__(self):
        self.response = ("คำตอบปกติ", [], False)
        self.last_call_args = None

    def set_response(self, text: str, receipts: list, structural_tool_called: bool) -> None:
        self.response = (text, receipts, structural_tool_called)

    def __call__(self, system_prompt, initial_messages, user_id=None, chat_id=None, tools_override=None, model=None,
                 max_iterations=None, max_tokens=None):
        self.last_call_args = {
            "system_prompt": system_prompt, "messages": initial_messages, "user_id": user_id, "chat_id": chat_id,
            "tools_override": tools_override, "model": model, "max_iterations": max_iterations,
            "max_tokens": max_tokens,
        }
        return self.response


class _FakeRagContext:
    """ตัวแทน _prepare_rag_context ที่ตั้งค่า top_chunks ได้จากเทสต์ตรงๆ (ค่าเริ่มต้นว่างเปล่า) — ใช้ยืนยันว่า
    guard fabricated_edit_claim_blocked ของ rag_answer ไม่พึ่ง ctx["top_chunks"] เป็นเงื่อนไขเลย (ใน production
    rerank_with_scores คืน top_k=5 เสมอไม่มี threshold กรองทิ้ง จึงแทบไม่เคยว่างเปล่าจริงๆ — ไม่ได้แปลว่าคำตอบมา
    จากฐานความรู้จริงเสมอไป) ค่าเริ่มต้นว่างเปล่าเหมือนเดิมสำหรับเทสต์อื่นที่ไม่เกี่ยวข้องกับเงื่อนไขนี้"""

    def __init__(self):
        self.top_chunks: list = []
        self.last_call_kwargs: Optional[dict] = None

    def __call__(self, query, history, image_data, user_id=None, chat_id=None, skip_legal_context=False):
        self.last_call_kwargs = {
            "query": query, "history": history, "image_data": image_data, "user_id": user_id,
            "chat_id": chat_id, "skip_legal_context": skip_legal_context,
        }
        return {
            "early_exit": False, "system_prompt": "ระบบทดสอบ", "messages": [{"role": "user", "content": query}],
            "top_chunks": self.top_chunks, "scores": [0.9] * len(self.top_chunks),
        }


def _fake_log_if_low_confidence(query, answer, top_chunks, scores):
    pass


def _fake_find_customers(query: str) -> list:
    """Stub เปล่าๆ แทน find_customers จริง (ดู main.py) — การทดสอบ _looks_like_library_query/
    _classify_question_route ในชุดนี้ไม่ครอบคลุมเงื่อนไข "ข้อความตรงกับชื่อลูกค้าที่มีอยู่จริง" เพราะ
    find_customers ยังพึ่ง get_all_customers ที่ยังไม่มีใน fake DB นี้ (ส่วน _get_enriched_library_files/
    get_file_categories_with_counts/get_all_library_file_meta/get_library_files มีแล้ว — ใช้โดย search_library/
    build_library_search_card โดยตรง) ไม่คุ้มที่จะสร้าง fake ส่วนลูกค้าเพิ่มสำหรับเงื่อนไขเดียวนี้ (ดู docstring
    ของ _looks_like_library_query ใน main.py และรายละเอียดใน PR description ของ fix/library-route)"""
    return []


class _FakeEncodedVector:
    """ตัวแทนผลลัพธ์ของ embed_model.encode(...) จริง (numpy array) — มีแค่ .tolist() ที่โค้ดจริงเรียกใช้
    (embed_model.encode(text).tolist()) ใช้เฉพาะตอน build_library_search_card/search_library เรียก ไม่มีความหมาย
    ทางคณิตศาสตร์จริง (ไม่โหลดโมเดลจริงเลยตามกติกาของงานนี้ — ดู _FakeEmbedModel)"""

    def __init__(self, values: list[float]):
        self._values = values

    def tolist(self):
        return list(self._values)


class _FakeEmbedModel:
    """ตัวแทน embed_model (SentenceTransformer จริง) — คืนเวกเตอร์คงที่เสมอไม่ว่า text จะเป็นอะไร (ไม่โหลดโมเดล
    จริงเลยตามกติกาของงานนี้ — sentence_transformers ไม่มีในสภาพแวดล้อมทดสอบนี้) เพียงพอสำหรับทดสอบ
    search_library/build_library_search_card เพราะ get_library_search_vector_scores_for_all (ด้านล่าง) คืนค่า
    ว่างเปล่าเสมอด้วยเหตุผลเดียวกัน — คะแนนรวมใน _score_library_search_candidates จึงมาจาก BM25 (คำนวณจริงด้วย
    rank_bm25 — มีในสภาพแวดล้อมนี้) และคะแนนตรงเป๊ะล้วนๆ ไม่ได้ทดสอบน้ำหนักฝั่ง embedding จริงจัง (เทสต์ของ
    _score_library_search_candidates เอง ส่ง embedding_scores มาตรงๆ ได้อยู่แล้วถ้าต้องการ)"""

    def encode(self, text):
        return _FakeEncodedVector([0.0, 0.0])


def _fake_get_library_search_vector_scores_for_all(query_embedding) -> dict:
    """Stub คืนค่าว่างเปล่าเสมอ แทน db.get_library_search_vector_scores_for_all จริง (pgvector cosine distance
    ผ่าน SQL จริง — ไม่มี DB จริงให้ต่อในสภาพแวดล้อมทดสอบนี้) ดู docstring ของ _FakeEmbedModel"""
    return {}


def build_namespace() -> dict:
    """คืน namespace (dict) ที่มีฟังก์ชันจริงจาก main.py/db.py ครบตาม WANTED_* ข้างบน ผูกกับ fake db ตัวใหม่
    (แยกกันทุกครั้งที่เรียก กันเทสต์หนึ่งกระทบอีกเทสต์หนึ่ง)"""
    fake_db = _FakeDb()
    fake_client = _FakeAnthropicClient()
    fake_agentic_loop = _FakeAgenticLoop()
    fake_rag_context = _FakeRagContext()

    ns = {
        "run_agentic_tool_loop": fake_agentic_loop,
        "_prepare_rag_context": fake_rag_context,
        "_log_if_low_confidence": _fake_log_if_low_confidence,
        "io": io, "os": os, "re": re, "html": html, "json": json, "zipfile": zipfile, "difflib": difflib,
        "datetime": datetime, "ZoneInfo": ZoneInfo, "urllib": urllib, "time": time,
        "tempfile": tempfile, "concurrent": concurrent.futures and __import__("concurrent"),
        "openpyxl": openpyxl, "MergedCell": MergedCell, "is_date_format": is_date_format,
        "CalcProperties": CalcProperties, "ExcelCompiler": ExcelCompiler, "Optional": Optional,
        "word_tokenize": word_tokenize, "BM25Okapi": BM25Okapi,
        "client": fake_client,
        "anthropic": _FakeAnthropicModule,
        "embed_model": _FakeEmbedModel(),
        "StorageNotConfiguredError": StorageNotConfiguredError,
        "get_editable_document_raw": fake_db.get_editable_document_raw,
        "get_editable_document": fake_db.get_editable_document,
        "get_latest_document_version": fake_db.get_latest_document_version,
        "get_second_latest_document_version": fake_db.get_second_latest_document_version,
        "create_document_version": fake_db.create_document_version,
        "delete_latest_document_version": fake_db.delete_latest_document_version,
        "link_edit_log_to_version": fake_db.link_edit_log_to_version,
        "log_document_edit": fake_db.log_document_edit,
        "set_chat_document_focus": fake_db.set_chat_document_focus,
        "set_chat_last_recommendations": fake_db.set_chat_last_recommendations,
        "get_chat_last_recommendations": fake_db.get_chat_last_recommendations,
        "set_chat_router_clarification_flag": fake_db.set_chat_router_clarification_flag,
        "get_chat_router_clarification_flag": fake_db.get_chat_router_clarification_flag,
        "get_chat_document_focus": fake_db.get_chat_document_focus,
        "add_chat_pending_action": fake_db.add_chat_pending_action,
        "get_chat_pending_actions": fake_db.get_chat_pending_actions,
        "delete_chat_pending_action": fake_db.delete_chat_pending_action,
        "clear_chat_pending_actions": fake_db.clear_chat_pending_actions,
        "update_chat_pending_action_version": fake_db.update_chat_pending_action_version,
        "list_active_editable_documents_by_chat": fake_db.list_active_editable_documents_by_chat,
        "create_editable_document": fake_db.create_editable_document,
        "get_library_file": fake_db.get_library_file,
        "get_library_files": fake_db.get_library_files,
        "get_file_categories_with_counts": fake_db.get_file_categories_with_counts,
        "get_library_file_meta": fake_db.get_library_file_meta,
        "get_all_library_file_meta": fake_db.get_all_library_file_meta,
        "upsert_library_search_card": fake_db.upsert_library_search_card,
        "get_all_library_search_cards": fake_db.get_all_library_search_cards,
        "get_library_file_ids_without_search_card": fake_db.get_library_file_ids_without_search_card,
        "get_library_search_vector_scores_for_all": _fake_get_library_search_vector_scores_for_all,
        "storage_get": fake_db.storage_get,
        "get_chat_messages": fake_db.get_chat_messages,
        "find_customers": _fake_find_customers,
        "set_chat_last_listing": fake_db.set_chat_last_listing,
        "get_chat_last_listing": fake_db.get_chat_last_listing,
    }

    main_tree = ast.parse(MAIN_PY.read_text(encoding="utf-8"))
    db_tree = ast.parse(DB_PY.read_text(encoding="utf-8"))
    document_fields_tree = ast.parse(DOCUMENT_FIELDS_PY.read_text(encoding="utf-8"))
    llm_tools_tree = ast.parse(LLM_TOOLS_PY.read_text(encoding="utf-8"))
    answer_guard_tree = ast.parse(ANSWER_GUARD_PY.read_text(encoding="utf-8"))
    compound_edits_tree = ast.parse(COMPOUND_EDITS_PY.read_text(encoding="utf-8"))
    _extract(main_tree, WANTED_MAIN_FUNCS, WANTED_MAIN_CONSTS, ns, "main.py")
    _extract(db_tree, WANTED_DB_FUNCS, set(), ns, "db.py")
    _extract_all(document_fields_tree, ns, "document_fields.py")
    # llm_tools.py/answer_guard.py/compound_edits.py เหมือน document_fields.py — ดึงทั้งไฟล์ด้วย _extract_all
    # "หลัง" main.py เสมอ
    _extract_all(llm_tools_tree, ns, "llm_tools.py")
    _extract_all(answer_guard_tree, ns, "answer_guard.py")
    _extract_all(compound_edits_tree, ns, "compound_edits.py")

    # run_agentic_tool_loop อยู่ใน WANTED_MAIN_FUNCS (ดึงของจริงมาเพื่อทดสอบ pause_turn/citation joining โดยตรง
    # — ดู _FakeAgenticResponse/_FakeAnthropicClient.response_queue) แต่ rag_answer/advisor_answer (ดึงของจริง
    # เหมือนกัน) ต้องเรียก "ตัวปลอม" ที่ควบคุมผลลัพธ์ได้ง่ายกว่าแทน (ดู _FakeAgenticLoop) — เก็บของจริงไว้ใต้ชื่อ
    # แยกต่างหากก่อน แล้วค่อยทับคืนด้วยตัวปลอมในชื่อเดิม (ลำดับสำคัญ: ต้องทับ "หลัง" _extract เสมอ)
    ns["_real_run_agentic_tool_loop"] = ns["run_agentic_tool_loop"]
    ns["run_agentic_tool_loop"] = fake_agentic_loop

    ns["_fake_db"] = fake_db
    ns["_fake_client"] = fake_client
    ns["_fake_agentic_loop"] = fake_agentic_loop
    ns["_fake_rag_context"] = fake_rag_context
    return ns


def inject_formula_cache(xlsx_bytes: bytes, cache_values: dict, sheet_file: str = "xl/worksheets/sheet1.xml") -> bytes:
    """ใส่ <v>ค่า cache</v> ให้เซลล์สูตรที่ระบุ (จำลองว่า Excel/LibreOffice เคยคำนวณไว้แล้ว) — ไฟล์ synthetic ที่
    สร้างด้วย openpyxl ล้วนๆ ไม่มีค่า cache เลย ทำให้ _extract_excel_labels (data_only=True) อ่านได้ None เสมอ
    สำหรับเซลล์สูตร ใช้เฉพาะตอนสร้าง fixture ทดสอบเท่านั้น ไม่เกี่ยวกับโค้ดจริงใน main.py/db.py เลย

    cache_values: {"G16": 1000, "G26": 8020.0, "B29": "some text"} ค่า str ใส่ตรงๆ ส่วนตัวเลขแปลงเป็น string
    ธรรมดา (xlsx เก็บตัวเลขเป็นข้อความใน XML อยู่แล้ว)"""
    with zipfile.ZipFile(io.BytesIO(xlsx_bytes)) as zin:
        names = zin.namelist()
        contents = {name: zin.read(name) for name in names}

    sheet_xml = contents[sheet_file].decode("utf-8")

    for cell_ref, value in cache_values.items():
        # openpyxl เขียน <v /> ว่างเปล่าต่อท้าย <f>...</f> ไว้แล้วเสมอ (placeholder) ต้องแทนที่ตัวนั้น ไม่ใช่
        # แทรกใหม่เฉยๆ (ไม่งั้นได้ <v/><v>...</v> ซ้อนกัน XML ผิด) ค่า cache เป็น string (เช่นผล BAHTTEXT)
        # ต้องเติม t="str" ในแท็ก <c> ด้วย ไม่งั้น openpyxl พยายามแปลงเป็นตัวเลขแล้ว error
        is_string_value = isinstance(value, str)
        pattern = re.compile(rf'<c r="{cell_ref}"([^>]*)>(.*?)</f>(?:<v\s*/>|<v>.*?</v>)?</c>')

        def _replace(m: "re.Match", _value=value, _is_string=is_string_value) -> str:
            attrs, inner = m.group(1), m.group(2)
            if _is_string:
                attrs = re.sub(r'\st="[^"]*"', "", attrs) + ' t="str"'
            value_text = _xml_escape(str(_value)) if _is_string else str(_value)
            return f'<c r="{cell_ref}"{attrs}>{inner}</f><v>{value_text}</v></c>'

        sheet_xml, count = pattern.subn(_replace, sheet_xml, count=1)
        if count == 0:
            raise ValueError(f"ไม่พบเซลล์สูตร {cell_ref} ใน {sheet_file} ให้แทรก cache")

    contents[sheet_file] = sheet_xml.encode("utf-8")

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            zout.writestr(name, contents[name])
    return buf.getvalue()


# ---------- fixture workbook สังเคราะห์ (ไม่ใช่ไฟล์ลูกค้าจริง) ----------
# ชีตเดียว "Quotation": หัวตารางแถว 15, แถวรายการ 16-25 (8 รายการจริง + 2 แถวว่าง) คอลัมน์ A=ลำดับ
# B(merge B:C)=รายการ D=จำนวน E=หน่วย F=ราคา/หน่วย G=จำนวนเงิน(สูตร) แถว 26=รวมเงิน 27=ภาษีมูลค่าเพิ่ม 7%
# 28=รวมราคาทั้งสิ้น 29=จำนวนเงินเป็นตัวอักษร(BAHTTEXT)
FIXTURE_ITEMS = [
    ("Item A", 10, "ชิ้น", 100),
    ("Item B", 5, "ชิ้น", 200),
    ("Item C", 3, "ชิ้น", 50),
    ("Item D", 2, "ชุด", 300),
    ("Item E", 1, "เครื่อง", 5000),
    ("Item F", 4, "ชิ้น", 25),
    ("Item G", 8, "ชิ้น", 10),
    ("Item H", 6, "ชิ้น", 15),
]
FIXTURE_HEADER_ROW = 15
FIXTURE_FIRST_ITEM_ROW = 16
FIXTURE_LAST_ITEM_ROW = 25  # ความจุ 10 แถว (16-25) แต่มีแค่ 8 รายการจริง (16-23) แถว 24,25 ว่าง


def build_fixture_workbook() -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Quotation"

    ws["A15"] = "ลำดับ"
    ws["B15"] = "รายการ"
    ws.merge_cells("B15:C15")
    ws["D15"] = "จำนวน"
    ws["E15"] = "หน่วย"
    ws["F15"] = "ราคา/หน่วย"
    ws["G15"] = "จำนวนเงิน"

    item_rows = range(FIXTURE_FIRST_ITEM_ROW, FIXTURE_FIRST_ITEM_ROW + len(FIXTURE_ITEMS))
    cache_values = {}
    for row, (desc, qty, unit, price) in zip(item_rows, FIXTURE_ITEMS):
        ws.cell(row=row, column=1, value=row - FIXTURE_FIRST_ITEM_ROW + 1)  # ลำดับ (ค่าคงที่)
        ws.cell(row=row, column=2, value=desc)
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
        ws.cell(row=row, column=4, value=qty)
        ws.cell(row=row, column=5, value=unit)
        ws.cell(row=row, column=6, value=price)
        ws.cell(row=row, column=7, value=f"=D{row}*F{row}")
        cache_values[f"G{row}"] = qty * price

    # แถว 24, 25 ว่างเปล่า (ไม่มี description/qty — ยังอยู่ในช่วงตารางเพื่อให้ capacity = 10) แต่คอลัมน์จำนวน
    # เงินยังมีสูตรรออยู่เหมือนไฟล์จริง (เช่น real ป่าไม้เขียวก็มี "=SUM(G24*I24)" แม้แถวว่าง) ประเมินได้ 0
    for row in range(FIXTURE_FIRST_ITEM_ROW + len(FIXTURE_ITEMS), FIXTURE_LAST_ITEM_ROW + 1):
        ws.cell(row=row, column=7, value=f"=D{row}*F{row}")
        cache_values[f"G{row}"] = 0

    total = sum(qty * price for _, qty, _, price in FIXTURE_ITEMS)
    ws["A26"] = "รวมเงิน"
    ws["G26"] = f"=SUM(G{FIXTURE_FIRST_ITEM_ROW}:G{FIXTURE_LAST_ITEM_ROW})"
    cache_values["G26"] = total

    vat = round(total * 0.07, 2)
    ws["A27"] = "ภาษีมูลค่าเพิ่ม 7%"
    ws["G27"] = "=SUM(G26)*7/100"
    cache_values["G27"] = vat

    grand_total = total + vat
    ws["A28"] = "รวมราคาทั้งสิ้น"
    ws["G28"] = "=SUM(G26:G27)"
    cache_values["G28"] = grand_total

    ws["A29"] = "จำนวนเงินเป็นตัวอักษร"
    ws["B29"] = "=BAHTTEXT(G28)"
    cache_values["B29"] = "จำลองข้อความ (BAHTTEXT คำนวณไม่ได้ด้วย pycel)"

    buf = io.BytesIO()
    wb.save(buf)
    raw = buf.getvalue()
    return inject_formula_cache(raw, cache_values)


FIXTURE_TOTAL = sum(qty * price for _, qty, _, price in FIXTURE_ITEMS)
FIXTURE_VAT = round(FIXTURE_TOTAL * 0.07, 2)
FIXTURE_GRAND_TOTAL = FIXTURE_TOTAL + FIXTURE_VAT
