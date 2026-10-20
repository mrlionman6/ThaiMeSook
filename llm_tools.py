"""เรียก Claude แบบบังคับให้ใช้ tool เสมอ ทนทานต่อรุ่นที่ไม่รองรับ tool_choice type "tool"/"any" (คืน 400) —
ใช้แทนการเรียก client.messages.create(..., tool_choice={"type":"tool", "name": ...}) ตรงๆ ทุกจุด

ลำดับ: tool_choice={"type":"tool",name} -> {"type":"any"} -> {"type":"auto"}+ย้ำใน system (auto ไม่เรียก tool
เลย ลองอ่าน JSON จาก text ที่ตรง required key ของ schema ก่อน ไม่เจอค่อย retry auto อีก 1 ครั้ง) จำ (model,
tool_choice type ที่เจอ 400) ไว้ข้าม tier นั้นทันทีครั้งถัดไป — ไม่ผูกกับ tool_name เพราะ 400 นี้มาจากตัวโมเดลไม่
รองรับ tool_choice ชนิดนั้นเอง (เช่น claude-sonnet-5-5 ไม่รองรับ "tool" เลยตามที่เจอจริงบน production)

call_with_required_tool คืน (response, reason, diag) เสมอ: reason=None ตอนสำเร็จ (response มี tool_use จริง
หรือ JSON ที่ parse ได้ครบ schema ห่อเป็น response ปลอมหน้าตาเหมือนเดิมทุกจุดที่ caller ใช้ — .content/
.stop_reason/.usage, diag=None) ไม่สำเร็จ response เป็น None เสมอ, reason อธิบายสาเหตุ ("no_tool_use" /
"max_tokens" / "json_invalid") และ diag เป็น {"stop_reason","blocks","out_tokens"} ให้ caller log ละเอียดได้
error ที่ไม่เกี่ยวกับ tool_choice (เช่น rate limit) ยัง raise ขึ้นไปเหมือนเดิมทุกประการ ไม่ถูกแปลงเป็น reason

thinking (optional): ส่งต่อให้ client.messages.create ทุกจุด (ดู create_with_optional_thinking) — โมเดลไม่รองรับ
รูปแบบที่ส่งไปก็ถอยไปเรียกแบบไม่ส่ง thinking ให้เองอัตโนมัติ ไม่ raise"""
import json
import re
import types
from typing import Optional

import anthropic

# {(model, tool_choice_type)} — ชนิด tool_choice ที่เจอ 400 ไปแล้วสำหรับโมเดลนี้ ข้าม tier นั้นทันทีครั้งถัดไป
_UNSUPPORTED_TOOL_CHOICE: set[tuple[str, str]] = set()

_JSON_OBJECT_RE = re.compile(r"\{.*\}", re.DOTALL)


def is_thinking_unsupported_error(exc: Exception) -> bool:
    """True ถ้า 400 นี้มาจากพารามิเตอร์ thinking ที่โมเดลนี้ไม่รองรับรูปแบบที่ส่งไป (เช่น rollback ผ่าน env ไปเป็น
    โมเดลรุ่นเก่าที่ไม่รู้จัก {"type":"between_tools"}) — ใช้กับ create_with_optional_thinking เท่านั้น"""
    return (
        isinstance(exc, anthropic.BadRequestError)
        and getattr(exc, "status_code", None) == 400
        and "thinking" in str(exc).lower()
    )


def create_with_optional_thinking(client, thinking: Optional[dict] = None, **kwargs):
    """เรียก client.messages.create ปกติ — ถ้าส่ง thinking มาด้วยแล้วโมเดลนี้ตอบ 400 ว่าไม่รองรับรูปแบบนั้น
    (ดู is_thinking_unsupported_error) เรียกซ้ำทันทีแบบไม่ส่ง thinking เลย (กัน rollback โมเดลด้วย env แล้วพัง)"""
    if thinking is None:
        return client.messages.create(**kwargs)
    try:
        return client.messages.create(thinking=thinking, **kwargs)
    except Exception as e:
        if not is_thinking_unsupported_error(e):
            raise
        print(f"[LLM] thinking_unsupported_fallback model={kwargs.get('model')}")
        return client.messages.create(**kwargs)


def _is_tool_choice_unsupported_error(exc: Exception) -> bool:
    return (
        isinstance(exc, anthropic.BadRequestError)
        and getattr(exc, "status_code", None) == 400
        and "tool_choice" in str(exc)
    )


def _find_tool_use(response, tool_name: str):
    return next((b for b in response.content if b.type == "tool_use" and b.name == tool_name), None)


def _required_keys_for_tool(tools: list, tool_name: str) -> list[str]:
    for t in tools:
        if t.get("name") == tool_name:
            return t.get("input_schema", {}).get("required") or []
    return []


def _response_diag(response) -> dict:
    """สรุป response สำหรับ log วินิจฉัย ไม่มีเนื้อหาข้อความจริงเลย แค่ชนิด block/stop_reason/จำนวน token"""
    blocks = ",".join(getattr(b, "type", "?") for b in (getattr(response, "content", None) or []))
    usage = getattr(response, "usage", None)
    out_tokens = getattr(usage, "output_tokens", None) if usage is not None else None
    return {"stop_reason": getattr(response, "stop_reason", None), "blocks": blocks or "-", "out_tokens": out_tokens}


def _extract_json_tool_input(text: str, required_keys: list[str]) -> Optional[dict]:
    """auto mode ไม่เรียก tool แต่ตอบเป็นข้อความ — ลองอ่าน JSON object จาก text ที่มี key ครบ required_keys (ทั้ง
    ก้อนเป๊ะๆ ก่อน แล้วค่อยลองช่วง {...} ที่ยาวที่สุด) คืน None ถ้าไม่เจอ/parse ไม่ได้/key ไม่ครบ"""
    if not text:
        return None
    candidates = [text.strip()]
    match = _JSON_OBJECT_RE.search(text)
    if match:
        candidates.append(match.group(0))
    for candidate in candidates:
        try:
            parsed = json.loads(candidate)
        except (json.JSONDecodeError, ValueError):
            continue
        if isinstance(parsed, dict) and all(k in parsed for k in required_keys):
            return parsed
    return None


def _synthetic_tool_use_response(original_response, tool_name: str, parsed_input: dict):
    """ห่อ JSON ที่ดึงจาก text ให้มีหน้าตาเหมือน response ที่มี tool_use block จริง (.content/.stop_reason/.usage)
    — caller ทุกจุดอ่านฟิลด์พวกนี้เหมือนกันหมด ไม่ต้องแก้โค้ด caller ตามการ fallback นี้เลย"""
    block = types.SimpleNamespace(type="tool_use", name=tool_name, input=parsed_input)
    return types.SimpleNamespace(
        content=[block], stop_reason=original_response.stop_reason, usage=getattr(original_response, "usage", None),
    )


def _call_auto_with_fallback(client, model, max_tokens, system, messages, tools, tool_name, required_keys, thinking=None):
    auto_system = system + f"\n\nสำคัญ: ต้องเรียกใช้เครื่องมือ {tool_name} เสมอทุกครั้ง ห้ามตอบเป็นข้อความธรรมดาเด็ดขาด"
    text, response = "", None
    for attempt in range(2):  # ลองครั้งแรก + retry อีก 1 ครั้งพร้อมข้อความย้ำ
        response = create_with_optional_thinking(
            client, thinking, model=model, max_tokens=max_tokens, system=auto_system, tools=tools,
            tool_choice={"type": "auto"}, messages=messages,
        )
        if _find_tool_use(response, tool_name) is not None:
            return response, None, None
        if response.stop_reason == "max_tokens":
            return None, "max_tokens", _response_diag(response)
        text = "".join(getattr(b, "text", "") for b in response.content if getattr(b, "type", None) == "text")
        parsed = _extract_json_tool_input(text, required_keys)
        if parsed is not None:
            return _synthetic_tool_use_response(response, tool_name, parsed), None, None
        if attempt == 0:
            auto_system += f"\n\nครั้งก่อนคุณไม่ได้เรียกเครื่องมือ {tool_name} เลย ครั้งนี้ต้องเรียกเครื่องมือนี้เท่านั้น"
    return None, ("json_invalid" if text.strip() else "no_tool_use"), _response_diag(response)


def call_with_required_tool(
    client, model: str, max_tokens: int, system: str, messages: list, tools: list, tool_name: str,
    thinking: Optional[dict] = None,
):
    """thinking (ถ้าส่งมา): ส่งต่อให้ client.messages.create ทุกจุดในนี้ ผ่าน create_with_optional_thinking —
    โมเดลไม่รองรับรูปแบบที่ส่งไปก็เรียกซ้ำแบบไม่ส่ง thinking ให้เองอัตโนมัติ"""
    for choice_type, choice_kwargs in (("tool", {"name": tool_name}), ("any", {})):
        if (model, choice_type) in _UNSUPPORTED_TOOL_CHOICE:
            continue
        try:
            response = create_with_optional_thinking(
                client, thinking, model=model, max_tokens=max_tokens, system=system, tools=tools,
                tool_choice={"type": choice_type, **choice_kwargs}, messages=messages,
            )
        except Exception as e:
            if not _is_tool_choice_unsupported_error(e):
                raise
            _UNSUPPORTED_TOOL_CHOICE.add((model, choice_type))
            print(f"[LLM] tool_choice_downgrade model={model} tool={tool_name} from={choice_type}")
            continue
        if _find_tool_use(response, tool_name) is not None:
            return response, None, None
        if response.stop_reason == "max_tokens":
            return None, "max_tokens", _response_diag(response)

    required_keys = _required_keys_for_tool(tools, tool_name)
    return _call_auto_with_fallback(client, model, max_tokens, system, messages, tools, tool_name, required_keys, thinking)
