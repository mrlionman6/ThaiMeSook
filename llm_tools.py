"""เรียก Claude แบบบังคับให้ใช้ tool เสมอ (tool_choice={"type":"tool",...}) ทนทานต่อรุ่นที่ไม่รองรับ tool_choice
type "tool"/"any" เลย (เช่น claude-sonnet-5-5 — คืน 400 "tool_choice: type ... are not supported for this
model" ทุกครั้ง ทำให้ route_intent_with_llm/_resolve_recommendation_intent_with_llm/ฯลฯ ล้มเหลวทุกข้อความ) —
ใช้แทนการเรียก client.messages.create(..., tool_choice={"type":"tool", "name": ...}) ตรงๆ ทุกจุดในโปรเจกต์นี้"""
import anthropic

# model ที่เคยเจอ 400 เพราะ tool_choice type "tool"/"any" แล้ว — จำไว้ต่อ process (ไม่ใช่ถาวรข้าม restart) กัน
# เสีย 1 call เปล่าๆ ทุกครั้งถัดไปสำหรับ model เดียวกัน
_DOWNGRADED_MODELS: set = set()


def _is_tool_choice_unsupported_error(exc: Exception) -> bool:
    return (
        isinstance(exc, anthropic.BadRequestError)
        and getattr(exc, "status_code", None) == 400
        and "tool_choice" in str(exc)
    )


def _find_tool_use(response, tool_name: str):
    return next((b for b in response.content if b.type == "tool_use" and b.name == tool_name), None)


def call_with_required_tool(client, model: str, max_tokens: int, system: str, messages: list, tools: list, tool_name: str):
    """เรียก client.messages.create บังคับให้เรียก tool ชื่อ tool_name เสมอ — ลอง tool_choice={"type":"tool",...}
    ก่อนเสมอ (เหมือนเดิมทุกจุด) ยกเว้น model นี้เคย 400 เพราะ tool_choice มาแล้วในรอบก่อนหน้า (ข้ามไปใช้ auto
    ทันที) เจอ 400 ที่ข้อความมี "tool_choice" (ดู _is_tool_choice_unsupported_error) -> จำ model นี้ไว้ แล้ว retry
    อีกครั้งเดียวด้วย tool_choice={"type":"auto"} พร้อมเติมประโยคบังคับต่อท้าย system เท่านั้น — exception อื่นที่
    ไม่ใช่กรณีนี้ (rate limit/auth/network ฯลฯ) ปล่อยให้หลุดขึ้นไปให้ caller จัดการเหมือนเดิมทุกประการ ไม่ retry เอง

    คืน response ของ Anthropic SDK ถ้ามี tool_use ของ tool_name จริง (ไม่ว่าจะเรียกแบบ forced หรือ auto) หรือ
    None ถ้าเรียกสำเร็จแต่ไม่มี tool_use ของ tool นี้เลย (ให้ caller fallback ไปทางอื่นเหมือนเดิม)"""
    if model not in _DOWNGRADED_MODELS:
        try:
            response = client.messages.create(
                model=model, max_tokens=max_tokens, system=system, tools=tools,
                tool_choice={"type": "tool", "name": tool_name}, messages=messages,
            )
            return response if _find_tool_use(response, tool_name) is not None else None
        except Exception as e:
            if not _is_tool_choice_unsupported_error(e):
                raise
            _DOWNGRADED_MODELS.add(model)
            print(f"[LLM] tool_choice_downgrade model={model} tool={tool_name}")

    auto_system = system + f"\n\nสำคัญ: ต้องเรียกใช้เครื่องมือ {tool_name} เสมอทุกครั้ง ห้ามตอบเป็นข้อความธรรมดาเด็ดขาด"
    response = client.messages.create(
        model=model, max_tokens=max_tokens, system=auto_system, tools=tools,
        tool_choice={"type": "auto"}, messages=messages,
    )
    return response if _find_tool_use(response, tool_name) is not None else None
