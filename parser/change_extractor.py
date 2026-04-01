from __future__ import annotations

import re

from openpyxl.worksheet.worksheet import Worksheet

from parser.merged_cells import iter_sheet_rows
from parser.models import ChangeBlock, ZoneRef
from parser.normalizer import normalize_for_match, normalize_text

CHANGE_INDEX_RE = re.compile(r"\bизм\.?\s*([\w\-/]+)", re.IGNORECASE)
DOC_CODE_RE = re.compile(r"\b[А-ЯA-Z0-9]{2,}[-/][А-ЯA-Z0-9\-/]{2,}\b")
HEADER_HINTS = (
    "содержание изменения",
    "извещение",
    "лист",
    "документ",
    "формат",
)


def _flatten_row(values: list[str | None]) -> str:
    return normalize_text(" ".join(v for v in values if v)) or ""


def _is_header_like(text: str) -> bool:
    lowered = normalize_for_match(text)
    return any(hint in lowered for hint in HEADER_HINTS)


def _extract_meta_signals(text: str) -> tuple[str | None, str | None, bool]:
    idx_match = CHANGE_INDEX_RE.search(text)
    doc_match = DOC_CODE_RE.search(text)

    change_index = idx_match.group(1) if idx_match else None
    doc_code = doc_match.group(0) if doc_match else None

    has_meta_keyword = "изм" in normalize_for_match(text)
    is_meta = has_meta_keyword and (change_index is not None or doc_code is not None)
    return change_index, doc_code, is_meta


def _collect_rows(sheet: Worksheet) -> list[tuple[int, str]]:
    rows: list[tuple[int, str]] = []
    for row_idx, row_values in iter_sheet_rows(sheet):
        text = _flatten_row(row_values)
        if text:
            rows.append((row_idx, text))
    return rows


def _find_change_starts(rows: list[tuple[int, str]]) -> list[int]:
    starts: list[int] = []
    for row_idx, text in rows:
        _, _, is_meta = _extract_meta_signals(text)
        if not is_meta:
            continue
        if _is_header_like(text):
            continue
        starts.append(row_idx)
    return starts


def extract_changes(sheet: Worksheet, sheet_index: int, start_global_seq: int) -> tuple[list[ChangeBlock], int]:
    rows = _collect_rows(sheet)
    starts = _find_change_starts(rows)
    if not starts:
        return [], start_global_seq

    row_map = {idx: text for idx, text in rows}
    changes: list[ChangeBlock] = []

    for seq_on_sheet, start_row in enumerate(starts, start=1):
        end_row = starts[seq_on_sheet] - 1 if seq_on_sheet < len(starts) else sheet.max_row
        meta_text = row_map.get(start_row, "")

        change_idx, doc_code, _ = _extract_meta_signals(meta_text)

        body_lines: list[str] = []
        for row_idx, line in rows:
            if row_idx <= start_row:
                continue
            if row_idx > end_row:
                break
            if _is_header_like(line):
                continue
            # skip accidental nested meta rows inside body window
            _, _, nested_meta = _extract_meta_signals(line)
            if nested_meta:
                continue
            body_lines.append(line)

        body_text = normalize_text("\n".join(body_lines))

        changes.append(
            ChangeBlock(
                sheet_index=sheet_index,
                change_seq_global=start_global_seq,
                change_seq_on_sheet=seq_on_sheet,
                change_index=change_idx,
                doc_code=doc_code,
                change_text=body_text,
                raw_meta_text=normalize_text(meta_text),
                zone_ref=ZoneRef(
                    meta_row_start=start_row,
                    meta_row_end=start_row,
                    body_row_start=start_row + 1 if start_row + 1 <= end_row else None,
                    body_row_end=end_row if end_row > start_row else None,
                ),
            )
        )
        start_global_seq += 1

    return changes, start_global_seq


__all__ = ["extract_changes", "_extract_meta_signals", "_find_change_starts"]
