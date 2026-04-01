from __future__ import annotations

from dataclasses import dataclass

from openpyxl.worksheet.worksheet import Worksheet

from parser.merged_cells import iter_sheet_rows
from parser.normalizer import normalize_for_match


BASE_MARKERS = ("извещение", "изм.", "содержание изменения")
FULL_HEADER_MARKERS = (
    "причина",
    "дата выпуска",
    "указание о заделе",
    "указания о внедрении",
)
CONTINUATION_HINT = "лист"


@dataclass
class SheetClassification:
    is_candidate: bool
    kind: str
    marker_hits: dict[str, bool]


def _collect_text(sheet: Worksheet, max_scan_rows: int = 120) -> str:
    chunks: list[str] = []
    for row_idx, row in iter_sheet_rows(sheet):
        if row_idx > max_scan_rows:
            break
        for value in row:
            normalized = normalize_for_match(value)
            if normalized:
                chunks.append(normalized)
    return " | ".join(chunks)


def classify_sheet(sheet: Worksheet) -> SheetClassification:
    blob = _collect_text(sheet)
    marker_hits = {m: m in blob for m in BASE_MARKERS}
    has_base = all(marker_hits.values())

    full_hits = [m for m in FULL_HEADER_MARKERS if m in blob]
    has_full_header = len(full_hits) >= 2
    has_continuation_hint = CONTINUATION_HINT in blob

    if not has_base:
        kind = "unknown"
        candidate = False
    elif has_full_header:
        kind = "full"
        candidate = True
    elif has_continuation_hint:
        kind = "continuation"
        candidate = True
    else:
        kind = "unknown"
        candidate = True

    return SheetClassification(is_candidate=candidate, kind=kind, marker_hits=marker_hits)
