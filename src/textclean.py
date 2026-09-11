"""DOCX 에 넣기 전에 문자열을 XML 안전하게 정리하는 유틸.

Word/Excel/웹에서 붙여넣은 텍스트에는 수직탭(0x0B, Word 의 Shift+Enter 줄바꿈)
같은 제어문자가 섞여 들어온다. python-docx 는 이런 문자를 만나면
"All strings must be XML compatible" ValueError 로 렌더링을 중단하므로,
렌더링 직전에 한 번 걸러 준다.
"""

from __future__ import annotations

from pydantic import BaseModel

NEWLINE = chr(10)
TAB = chr(9)

# 줄바꿈으로 바꿔줄 문자들 (수직탭·폼피드·NEL·유니코드 줄/문단 구분자)
_TO_NEWLINE = {chr(0x0B), chr(0x0C), chr(0x85), chr(0x2028), chr(0x2029)}
# 유지할 제어문자 (줄바꿈·탭)
_KEEP = {NEWLINE, TAB}


def clean_text(value: str) -> str:
    """제어문자를 제거/치환해 XML 에 안전한 문자열로 만든다."""
    value = value.replace(chr(13) + NEWLINE, NEWLINE).replace(chr(13), NEWLINE)
    out = []
    for ch in value:
        if ch in _TO_NEWLINE:
            out.append(NEWLINE)
        elif ch in _KEEP:
            out.append(ch)
        elif ord(ch) < 0x20 or ord(ch) == 0x7F or 0x80 <= ord(ch) <= 0x9F:
            continue
        else:
            out.append(ch)
    return "".join(out)


def clean_value(value):
    """문자열/리스트/딕셔너리/pydantic 모델을 재귀적으로 정리한다."""
    if isinstance(value, str):
        return clean_text(value)
    if isinstance(value, BaseModel):
        clean_model(value)
        return value
    if isinstance(value, list):
        return [clean_value(v) for v in value]
    if isinstance(value, tuple):
        return tuple(clean_value(v) for v in value)
    if isinstance(value, dict):
        return {k: clean_value(v) for k, v in value.items()}
    return value


def clean_model(model: BaseModel) -> BaseModel:
    """pydantic 모델의 모든 문자열 필드를 제자리에서 정리한다."""
    for field_name in type(model).model_fields:
        current = getattr(model, field_name, None)
        cleaned = clean_value(current)
        if cleaned is not current:
            object.__setattr__(model, field_name, cleaned)
    return model
