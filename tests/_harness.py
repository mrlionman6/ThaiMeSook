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
import re
import json
import zipfile
import difflib
import datetime
import tempfile
import concurrent.futures
from pathlib import Path
from typing import Optional
from xml.sax.saxutils import escape as _xml_escape

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.styles.numbers import is_date_format
from openpyxl.workbook.properties import CalcProperties
from pycel import ExcelCompiler
from pythainlp.tokenize import word_tokenize

MAIN_PY = Path(__file__).parent.parent / "main.py"
DB_PY = Path(__file__).parent.parent / "db.py"

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
    "_build_rate_change_mutate_fn", "_label_side_effects", "_format_side_effect_lines",
    "_get_single_item_table",
    "_build_delete_table_item_mutate_fn", "_describe_table_item",
    "_build_add_table_item_mutate_fn",
    "_table_has_non_formula_amount_value",
    "_resolve_and_build_mutate_fn", "_replay_pending_queue",
    "propose_change_rate", "propose_update_table_item", "propose_delete_table_item", "propose_add_table_item",
    "_apply_single_queue_item", "_apply_pending_queue", "_pending_queue_baseline_valid",
    "_handle_pending_queue_command",
    "undo_last_edit",
    "_is_document_reading_question", "_has_legal_keyword", "_answer_document_reading_question",
    "_is_structural_edit_command",
    "_has_explicit_finalize_word", "_classify_pending_queue_command", "_is_pending_queue_fresh",
    "_describe_pending_queue", "_get_document_version_marker",
    "_label_position", "_find_label_by_position",
    "_get_active_chat_document", "execute_open_library_file", "_normalize_filename_for_matching",
    "_match_and_apply_excel_edit", "_parse_json_response",
    "rag_answer", "contains_unexpected_script", "_normalize_polite_particles",
}

WANTED_MAIN_CONSTS = {
    "EXCEL_LABEL_SECTION_HEADING_LOOKBACK", "_ITEM_TABLE_HEADER_WORDS", "_ITEM_TABLE_COLUMN_MAP",
    "_ITEM_TABLE_STOP_PREFIXES", "_ITEM_TABLE_MAX_ROWS", "_ITEM_TABLE_EMPTY_STREAK_LIMIT",
    "FORMULA_RECALC_TIMEOUT_SECONDS", "_RATE_FORMULA_PATTERNS", "_PERCENT_IN_TEXT_RE", "_TABLE_INPUT_FIELDS",
    "_DOCUMENT_READING_QUESTION_WORDS", "_DOCUMENT_EDIT_COMMAND_WORDS", "_STRUCTURAL_ITEM_NO_RE",
    "_LEGAL_KNOWLEDGE_KEYWORDS", "_FINALIZE_EXPLICIT_WORDS",
    "_CONFIRM_ALL_TEXTS", "_CONFIRM_LATEST_TEXTS", "_CONFIRM_SINGLE_TEXTS",
    "_CANCEL_LATEST_TEXTS", "_CANCEL_ALL_TEXTS", "PENDING_ACTION_TTL_SECONDS",
    "_LABEL_TRAILING_PUNCTUATION", "_TABLE_NON_FORMULA_AMOUNT_ERROR",
    "EXCEL_EDITOR_SYSTEM_PROMPT", "FORMULA_EDIT_WARNING_NOTICE",
    "_STRUCTURAL_NO_TOOL_FALLBACK_MESSAGE",
    "_FABRICATED_EDIT_CLAIM_RE", "_FABRICATED_EDIT_CLAIM_STRUCTURAL_ONLY_PATTERNS",
    "UNEXPECTED_SCRIPT_PATTERN", "MAX_ANSWER_RETRIES",
    "_POLITE_PARTICLE_COMPOUND_RE", "_POLITE_PARTICLE_TRAILING_RE",
}

# _looks_like_fabricated_edit_claim ไม่ได้อยู่ใน WANTED_MAIN_FUNCS ด้านบนโดยตั้งใจ — ไม่ต้องใช้ใน test ชุดนี้
# โดยตรง (ทดสอบผ่าน rag_answer ซึ่งหนักเกินจะ extract ทั้งฟังก์ชัน) แต่ยังปล่อยให้ดึงได้ถ้าจำเป็นภายหลัง
WANTED_MAIN_FUNCS.add("_looks_like_fabricated_edit_claim")

WANTED_DB_FUNCS = {"_is_macro_workbook"}


class _FakeDb:
    """เก็บสถานะจำลองของ EditableDocument/DocumentVersion/ChatPendingAction/DocumentEditLog ในหน่วยความจำล้วนๆ
    (ไม่ต่อ DB จริงเลย) — เลียนแบบ semantics ของ db.py เท่าที่ฟังก์ชันที่ทดสอบต้องใช้จริงเท่านั้น"""

    def __init__(self):
        self.documents: dict[int, dict] = {}
        self.versions: dict[int, list[dict]] = {}
        self.edit_logs: dict[int, dict] = {}
        self.pending_actions: dict[int, list[dict]] = {}  # chat_id -> [{"id","action","preview_text","created_at"}, ...] เรียงเก่า->ใหม่
        self.focus: dict[int, int] = {}
        self.closed_ids: set[int] = set()
        self.library_files: dict[int, dict] = {}
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

    def get_chat_document_focus(self, chat_id, user_id):
        document_id = self.focus.get(chat_id)
        if document_id is None:
            return None
        doc = self.documents.get(document_id)
        if doc and doc["user_id"] == user_id:
            return doc
        return None

    def add_library_file(self, file_id: int, filename: str, storage_key: str) -> None:
        self.library_files[file_id] = {"file_id": file_id, "filename": filename, "storage_key": storage_key}

    def get_library_file(self, file_id):
        return self.library_files.get(file_id)

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
        if self._client.raise_on_create is not None:
            raise self._client.raise_on_create
        return _FakeAnthropicResponse(self._client.response_text)


class _FakeAnthropicClient:
    """ตัวแทน anthropic.Client — คืนข้อความคงที่ (ตั้งค่าได้) เสมอ ไม่เรียก API จริงเลย เก็บ kwargs ของการเรียก
    ล่าสุดไว้ด้วย (last_call_kwargs) ให้เทสต์ตรวจสอบได้ว่าส่ง system/messages อะไรเข้าไป ตั้ง raise_on_create
    เป็น Exception instance เพื่อจำลอง API ล้ม/timeout ได้"""

    def __init__(self, response_text: str = "คำตอบจำลองจาก LLM"):
        self.response_text = response_text
        self.last_call_kwargs = None
        self.raise_on_create: Exception | None = None
        self.messages = _FakeAnthropicMessages(self)


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

    missing = (wanted_funcs - found_funcs) | (wanted_consts - found_consts)
    if missing:
        raise RuntimeError(f"ดึงฟังก์ชัน/ค่าคงที่จาก {label} ไม่ครบ: {sorted(missing)}")


class _FakeAgenticLoop:
    """ตัวแทน run_agentic_tool_loop ที่ควบคุมผลลัพธ์ได้ตรงๆ จากเทสต์ (ตั้งค่าผ่าน set_response ก่อนเรียก
    rag_answer) — ใช้ทดสอบ guard logic ของ rag_answer เอง (structural_no_tool/fabricated_edit_claim_blocked)
    แบบแยกส่วน โดยไม่ต้องเรียก Claude จริงหรือประกอบ tools list เต็มๆ (หนักเกินไปสำหรับเทสต์นี้)"""

    def __init__(self):
        self.response = ("คำตอบปกติ", [], False)
        self.last_call_args = None

    def set_response(self, text: str, receipts: list, structural_tool_called: bool) -> None:
        self.response = (text, receipts, structural_tool_called)

    def __call__(self, system_prompt, initial_messages, user_id=None, chat_id=None):
        self.last_call_args = {"system_prompt": system_prompt, "user_id": user_id, "chat_id": chat_id}
        return self.response


class _FakeRagContext:
    """ตัวแทน _prepare_rag_context ที่ตั้งค่า top_chunks ได้จากเทสต์ตรงๆ (ค่าเริ่มต้นว่างเปล่า) — ใช้ยืนยันว่า
    guard fabricated_edit_claim_blocked ของ rag_answer ไม่พึ่ง ctx["top_chunks"] เป็นเงื่อนไขเลย (ใน production
    rerank_with_scores คืน top_k=5 เสมอไม่มี threshold กรองทิ้ง จึงแทบไม่เคยว่างเปล่าจริงๆ — ไม่ได้แปลว่าคำตอบมา
    จากฐานความรู้จริงเสมอไป) ค่าเริ่มต้นว่างเปล่าเหมือนเดิมสำหรับเทสต์อื่นที่ไม่เกี่ยวข้องกับเงื่อนไขนี้"""

    def __init__(self):
        self.top_chunks: list = []

    def __call__(self, query, history, image_data, user_id=None, chat_id=None):
        return {
            "early_exit": False, "system_prompt": "ระบบทดสอบ", "messages": [{"role": "user", "content": query}],
            "top_chunks": self.top_chunks, "scores": [0.9] * len(self.top_chunks),
        }


def _fake_log_if_low_confidence(query, answer, top_chunks, scores):
    pass


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
        "io": io, "os": os, "re": re, "json": json, "zipfile": zipfile, "difflib": difflib, "datetime": datetime,
        "tempfile": tempfile, "concurrent": concurrent.futures and __import__("concurrent"),
        "openpyxl": openpyxl, "MergedCell": MergedCell, "is_date_format": is_date_format,
        "CalcProperties": CalcProperties, "ExcelCompiler": ExcelCompiler, "Optional": Optional,
        "word_tokenize": word_tokenize,
        "client": fake_client,
        "StorageNotConfiguredError": StorageNotConfiguredError,
        "get_editable_document_raw": fake_db.get_editable_document_raw,
        "get_editable_document": fake_db.get_editable_document,
        "get_latest_document_version": fake_db.get_latest_document_version,
        "create_document_version": fake_db.create_document_version,
        "delete_latest_document_version": fake_db.delete_latest_document_version,
        "link_edit_log_to_version": fake_db.link_edit_log_to_version,
        "log_document_edit": fake_db.log_document_edit,
        "set_chat_document_focus": fake_db.set_chat_document_focus,
        "get_chat_document_focus": fake_db.get_chat_document_focus,
        "add_chat_pending_action": fake_db.add_chat_pending_action,
        "get_chat_pending_actions": fake_db.get_chat_pending_actions,
        "delete_chat_pending_action": fake_db.delete_chat_pending_action,
        "clear_chat_pending_actions": fake_db.clear_chat_pending_actions,
        "update_chat_pending_action_version": fake_db.update_chat_pending_action_version,
        "list_active_editable_documents_by_chat": fake_db.list_active_editable_documents_by_chat,
        "create_editable_document": fake_db.create_editable_document,
        "get_library_file": fake_db.get_library_file,
        "storage_get": fake_db.storage_get,
    }

    main_tree = ast.parse(MAIN_PY.read_text(encoding="utf-8"))
    db_tree = ast.parse(DB_PY.read_text(encoding="utf-8"))
    _extract(main_tree, WANTED_MAIN_FUNCS, WANTED_MAIN_CONSTS, ns, "main.py")
    _extract(db_tree, WANTED_DB_FUNCS, set(), ns, "db.py")

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
