"""eval_answers.py — ส่งคำถามจาก answer_eval_set.json เข้า /ask ทีละข้อ แบบไม่ล็อกอิน ไม่ส่ง chat_id
(ยืนยันรูปแบบคำขอจากโค้ดจริงของ /ask ใน main.py ก่อนเขียนสคริปต์นี้แล้ว: รับ multipart/form-data หรือ
application/x-www-form-urlencoded ก็ได้ผ่าน FastAPI Form()/File() — ฟิลด์ query (str), chat_id (int,
ไม่บังคับ), image (ไม่บังคับ) สคริปต์นี้ส่งแค่ query ฟิลด์เดียว ไม่ส่ง chat_id/image เลย จึงไม่มี session/cookie
ใดๆ เข้ามาเกี่ยวข้อง = ไม่ล็อกอินเสมอ ผลตอบกลับเป็น JSON {"answer":, "sources":, "chat_id":}) เครื่องมือนี้ไม่
ประเมินความถูกต้องของคำตอบเองเลย (ต้องให้มนุษย์อ่าน eval_out/answers_<label>.md เทียบกับความรู้จริง) แค่เก็บ
คำถาม+คำตอบเต็มไว้ให้ครบทุกข้อ

ใช้: python tools/eval_answers.py --label <ชื่อรอบทดสอบ> [--base-url http://127.0.0.1:8000]

terminal พิมพ์แค่ "[PASS] <id>" / "[FAIL] <id> exception_type=..." ต่อข้อแบบ ASCII ล้วน (ห้ามพิมพ์คำตอบ
ภาษาไทยยาวๆ ออก terminal เด็ดขาด) ผลเต็ม (คำถาม+คำตอบ) บันทึกไว้ใน eval_out/answers_<label>.md (อ่านง่าย) และ
eval_out/answers_<label>.json (ไว้ประมวลผลต่อด้วยโปรแกรมอื่นถ้าต้องการ) — eval_out/ อยู่ใน .gitignore แล้ว"""
import argparse
import json
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

REQUEST_TIMEOUT_SECONDS = 120
DELAY_BETWEEN_QUESTIONS_SECONDS = 3
EVAL_SET_PATH = Path(__file__).parent / "answer_eval_set.json"
OUTPUT_DIR = Path(__file__).parent.parent / "eval_out"


def ask(base_url: str, query: str) -> dict:
    """ส่งคำถามเดียวไปที่ POST /ask แบบไม่ล็อกอิน (ไม่แนบ cookie/session ใดๆ) และไม่ส่ง chat_id เลย (ปล่อยให้
    เป็นค่าว่าง — เพราะไม่ล็อกอิน ask_question() จะไม่สร้าง/บันทึกแชทจริงอยู่แล้วไม่ว่าจะส่ง chat_id มาหรือไม่)
    ใช้ application/x-www-form-urlencoded แทน multipart/form-data เพราะ FastAPI Form() รับได้ทั้งคู่ และ
    คำขอนี้มีแค่ฟิลด์เดียว (query) ไม่ต้องสร้าง multipart boundary เอง คืน dict ที่ parse จาก JSON response
    เสมอ (raise ออกไปให้ caller จัดการถ้า HTTP error หรือ parse ไม่ได้)"""
    data = urllib.parse.urlencode({"query": query}).encode("utf-8")
    req = urllib.request.Request(
        base_url.rstrip("/") + "/ask",
        data=data,
        method="POST",
        headers={"Content-Type": "application/x-www-form-urlencoded"},
    )
    with urllib.request.urlopen(req, timeout=REQUEST_TIMEOUT_SECONDS) as resp:
        return json.loads(resp.read().decode("utf-8"))


def main():
    parser = argparse.ArgumentParser(description="Run the legal-answer eval set against a running /ask endpoint")
    parser.add_argument("--label", required=True, help="Label for this eval run, used in the output filenames")
    parser.add_argument("--base-url", default="http://127.0.0.1:8000", help="Base URL of the running server")
    args = parser.parse_args()

    with open(EVAL_SET_PATH, encoding="utf-8") as f:
        eval_set = json.load(f)

    OUTPUT_DIR.mkdir(exist_ok=True)
    results = []

    for i, item in enumerate(eval_set):
        qid = item["id"]
        query = item["query"]
        try:
            response = ask(args.base_url, query)
            results.append({
                "id": qid, "query": query, "ok": True,
                "answer": response.get("answer", ""), "error": None,
            })
            print(f"[PASS] {qid}")
        except urllib.error.HTTPError as e:
            try:
                error_body = e.read().decode("utf-8", errors="replace")
            except Exception:
                error_body = ""
            results.append({
                "id": qid, "query": query, "ok": False, "answer": None,
                "error": f"HTTPError {e.code}: {error_body}",
            })
            print(f"[FAIL] {qid} exception_type=HTTPError status={e.code}")
        except Exception as e:
            results.append({
                "id": qid, "query": query, "ok": False, "answer": None,
                "error": f"{type(e).__name__}: {e}",
            })
            print(f"[FAIL] {qid} exception_type={type(e).__name__}")

        if i < len(eval_set) - 1:
            time.sleep(DELAY_BETWEEN_QUESTIONS_SECONDS)

    json_path = OUTPUT_DIR / f"answers_{args.label}.json"
    md_path = OUTPUT_DIR / f"answers_{args.label}.md"

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    with open(md_path, "w", encoding="utf-8") as f:
        f.write(f"# Answer eval results: {args.label}\n\n")
        for r in results:
            f.write(f"## {r['id']}\n\n")
            f.write(f"**Question:** {r['query']}\n\n")
            if r["ok"]:
                f.write(f"**Answer:**\n\n{r['answer']}\n\n")
            else:
                f.write(f"**Error:** {r['error']}\n\n")
            f.write("---\n\n")

    succeeded = sum(1 for r in results if r["ok"])
    print(f"[DONE] {succeeded}/{len(results)} succeeded, saved to {json_path} and {md_path}")


if __name__ == "__main__":
    main()
