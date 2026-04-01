from __future__ import annotations

import re


SPACE_RE = re.compile(r"\s+")


def normalize_text(value: str | None) -> str | None:
    if value is None:
        return None
    normalized = SPACE_RE.sub(" ", value.replace("\n", " ")).strip()
    return normalized or None


def normalize_for_match(value: str | None) -> str:
    text = normalize_text(value) or ""
    return text.lower().replace("ё", "е")


def safe_int(value: str | None) -> int | None:
    if not value:
        return None
    cleaned = re.sub(r"[^\d]", "", value)
    if not cleaned:
        return None
    return int(cleaned)
