from __future__ import annotations

import re

from openpyxl.worksheet.worksheet import Worksheet

from parser.change_extractor import extract_changes
from parser.merged_cells import iter_sheet_rows
from parser.models import SheetLocalHeader, SheetResult
from parser.normalizer import normalize_for_match, normalize_text, safe_int
from parser.sheet_classifier import classify_sheet

NOTICE_NUMBER_RE = re.compile(r"извещение\s*(?:об\s*изменении)?\s*№\s*([A-ZА-Я0-9][A-ZА-Я0-9.\-/]{2,})", re.IGNORECASE)
NOTICE_LINE_RE = re.compile(r"\b№\s*([A-ZА-Я0-9][A-ZА-Я0-9.\-/]{2,})", re.IGNORECASE)
SHEET_RE = re.compile(r"лист\s*№?\s*(\d+)", re.IGNORECASE)
INVALID_NOTICE_VALUES = {"извещение", "об изменении", "изменение"}


def _scan_header_rows(sheet: Worksheet, max_rows: int = 50) -> list[tuple[int, str]]:
    rows: list[tuple[int, str]] = []
    for row_idx, row in iter_sheet_rows(sheet):
        if row_idx > max_rows:
            break
        text = normalize_text(" ".join(v for v in row if v))
        if text:
            rows.append((row_idx, text))
    return rows


def _is_valid_notice_candidate(value: str | None) -> bool:
    if not value:
        return False
    norm = normalize_for_match(value)
    if norm in INVALID_NOTICE_VALUES:
        return False
    # Must contain at least one digit to avoid labels like "Извещение"
    return any(ch.isdigit() for ch in value)


def detect_notice_number(sheet: Worksheet) -> str | None:
    rows = _scan_header_rows(sheet)

    # Strict primary pattern with explicit "Извещение ... №"
    for _, text in rows:
        match = NOTICE_NUMBER_RE.search(text)
        if match and _is_valid_notice_candidate(match.group(1)):
            return match.group(1)

    # Fallback: find rows containing "извещение" and parse nearest number marker.
    for _, text in rows:
        if "извещение" not in normalize_for_match(text):
            continue
        match = NOTICE_LINE_RE.search(text)
        if match and _is_valid_notice_candidate(match.group(1)):
            return match.group(1)

    return None


def detect_sheet_no(sheet: Worksheet) -> int | None:
    for _, header in _scan_header_rows(sheet):
        match = SHEET_RE.search(header)
        if match:
            return safe_int(match.group(1))

    normalized_title = normalize_for_match(sheet.title)
    return safe_int(normalized_title)


def parse_sheet(
    sheet: Worksheet, sheet_index: int, start_global_seq: int
) -> tuple[SheetResult, int, bool, dict[str, object]]:
    classification = classify_sheet(sheet)
    notice_number = detect_notice_number(sheet)
    sheet_no = detect_sheet_no(sheet)

    changes, next_seq, extraction_debug = extract_changes(
        sheet,
        sheet_index=sheet_index,
        start_global_seq=start_global_seq,
    )

    result = SheetResult(
        sheet_index=sheet_index,
        sheet_name=sheet.title,
        sheet_kind=classification.kind,
        sheet_no_detected=sheet_no,
        sheet_local_header=SheetLocalHeader(notice_number=notice_number),
        changes=changes,
    )
    return result, next_seq, classification.is_candidate, extraction_debug
