from __future__ import annotations

from dataclasses import dataclass


@dataclass
class QueryInfo:
    name: str
    url: str


@dataclass
class QueryField:
    label: str
    name: str


@dataclass
class QueryAttempt:
    success: bool
    message: str
    data: dict[str, str] | None = None
    captcha_required: bool = False
    used_ocr: bool = False


@dataclass
class QueryFailure:
    time: str
    index: int
    data: str
    reason: str
