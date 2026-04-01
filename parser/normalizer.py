from __future__ import annotations

import re


SPACE_RE = re.compile(r"\s+")
DASH_BLOCK_RE = re.compile(r"^(?:[-–—]\s*){2,}$")


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


def collapse_repeated_tokens(text: str | None) -> str | None:
    norm = normalize_text(text)
    if not norm:
        return None
    tokens = norm.split()
    collapsed: list[str] = []
    prev = None
    for tok in tokens:
        if tok == prev:
            continue
        collapsed.append(tok)
        prev = tok
    return normalize_text(" ".join(collapsed))


def collapse_repeated_phrases(text: str | None) -> str | None:
    norm = normalize_text(text)
    if not norm:
        return None
    words = norm.split()
    for size in range(1, max(2, len(words) // 2 + 1)):
        phrase = words[:size]
        repeats = len(words) // size
        if repeats >= 2 and phrase * repeats == words[: size * repeats] and not words[size * repeats :]:
            return normalize_text(" ".join(phrase))
    return norm


def normalize_dash_noise(text: str | None) -> str | None:
    norm = normalize_text(text)
    if not norm:
        return None
    if DASH_BLOCK_RE.match(norm):
        return None
    return norm
