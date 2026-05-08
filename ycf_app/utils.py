from __future__ import annotations

from datetime import datetime
from typing import Any

import pandas as pd


def now_text() -> str:
    return datetime.now().strftime("%H:%M:%S")


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    if pd.isna(value):
        return ""
    return str(value).strip()


def unique_key(existing: list[str], key: str) -> str:
    if key not in existing:
        return key
    index = 2
    while f"{key}_{index}" in existing:
        index += 1
    return f"{key}_{index}"
