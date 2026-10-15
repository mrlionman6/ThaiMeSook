"""เรียก Claude แบบบังคับให้ใช้ tool เสมอ ทนทานต่อรุ่นที่ไม่รองรับ tool_choice type "tool"/"any" (คืน 400) —
ใช้แทนการเรียก client.messages.create(..., tool_choice={"type":"tool", "name": ...}) ตรงๆ ทุกจุด"""
import anthropic

# model ที่เคยเจอ 400 เพราะ tool_choice แล้ว — จำไว้ต่อ process กันเสีย 1 call เปล่าๆ ทุกครั้งถัดไป
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
    """ลอง tool_choice={"type":"tool",...} ก่อนเสมอ เจอ 400 เพราะ tool_choice -> retry ด้วย "auto" + ประโยคบังคับ
    ใน system (จำ model นี้ไว้ ครั้งถัดไปข้ามไป auto ทันที) คืน response ถ้ามี tool_use จริง หรือ None ถ้าไม่มี"""
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
