"""เปรียบเทียบเอกสาร 2 ไฟล์ที่เปิดอยู่ในแชทเดียวกัน — router (AI) แค่เลือกว่าอยากเทียบไฟล์ไหนจาก document_id ที่
เห็น โค้ดเป็นคนตรวจว่า id มีจริงในแชทนี้ เลือกไฟล์ให้เองถ้า router ระบุมาไม่ครบ/ไม่มี แล้วคืนตารางผลต่างที่สร้าง
ด้วยโค้ดล้วนๆ อยู่แล้ว (ดู execute_compare_chat_documents/_compare_editable_documents) ไม่ผ่าน LLM เลยสักขั้น

import จาก main.py แบบ deferred (หลังนิยาม execute_compare_chat_documents) เหมือน document_fields.py ทุกประการ"""
from main import execute_compare_chat_documents


def _resolve_compare_target_ids(route: dict, active_docs: list[dict]) -> list[int]:
    """กรอง compare_document_ids จาก router เหลือเฉพาะ id ที่เป็นไฟล์เปิดอยู่จริงในแชทนี้เท่านั้น"""
    active_ids = {doc["id"] for doc in active_docs}
    raw_ids = route.get("compare_document_ids") or []
    return [doc_id for doc_id in raw_ids if isinstance(doc_id, int) and doc_id in active_ids]


def dispatch_compare_documents(route: dict, active_docs: list[dict], chat_id: int, user_id: int) -> dict:
    """เลือก 2 ไฟล์ที่จะเทียบด้วยโค้ดเสมอ ไม่เชื่อ router เพียวๆ: compare_document_ids ที่ตรวจผ่านครบ 2 ตัวใช้ตาม
    นั้น ไม่ครบ/ไม่มี -> เปิดอยู่พอดี 2 ไฟล์ใช้ 2 ไฟล์นั้นเลย เปิดอยู่มากกว่า 2 ให้เลือกก่อน เปิดอยู่น้อยกว่า 2 บอก
    ตามจริงว่าเปิดอะไรอยู่บ้าง (active_docs มาจาก list_active_editable_documents_by_chat ของ caller อยู่แล้ว)"""
    valid_ids = _resolve_compare_target_ids(route, active_docs)
    if len(valid_ids) < 2:
        if len(active_docs) == 2:
            valid_ids = [doc["id"] for doc in active_docs]
        elif len(active_docs) > 2:
            names = ", ".join(f"'{d['filename']}'" for d in active_docs)
            return {
                "document_id": None,
                "answer": f"ตอนนี้เปิดอยู่ {len(active_docs)} ไฟล์ในแชทนี้ ({names}) ต้องการเปรียบเทียบไฟล์ไหนกับไฟล์ไหนครับ",
            }
        elif len(active_docs) == 1:
            return {
                "document_id": None,
                "answer": f"ตอนนี้เปิดอยู่แค่ไฟล์เดียวครับ ('{active_docs[0]['filename']}') เปิดอีกไฟล์ก่อนถึงจะเปรียบเทียบได้ครับ",
            }
        else:
            return {
                "document_id": None,
                "answer": "ตอนนี้ยังไม่มีไฟล์เปิดอยู่ในแชทนี้เลยครับ เปิด 2 ไฟล์ก่อนถึงจะเปรียบเทียบได้",
            }

    doc_a_id, doc_b_id = valid_ids[0], valid_ids[1]
    result = execute_compare_chat_documents(
        {"document_id_a": doc_a_id, "document_id_b": doc_b_id}, user_id, chat_id,
    )
    answer = result.get("result") or result.get("error") or "เปรียบเทียบไม่สำเร็จ ลองใหม่อีกครั้ง"
    return {"document_id": None, "answer": answer}
