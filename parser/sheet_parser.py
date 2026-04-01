from __future__ import annotations

import re

from openpyxl.worksheet.worksheet import Worksheet

from parser.change_extractor import extract_changes
from parser.merged_cells import iter_sheet_rows
from parser.models import SheetLocalHeader, SheetResult
from parser.normalizer import normalize_for_match, normalize_text, safe_int
from parser.sheet_classifier import classify_sheet

NOTICE_RE = re.compile(r"извещение\s*№?\s*([\w\-/]+)", re.IGNORECASE)
SHEET_RE = re.compile(r"лист\s*№?\s*(\d+)", re.IGNORECASE)


def _scan_header_text(sheet: Worksheet, max_rows: int = 40) -> str:
    parts: list[str] = []
    for row_idx, row in iter_sheet_rows(sheet):
        if row_idx > max_rows:
            break
        text = normalize_text(" ".join(v for v in row if v))
        if text:
            parts.append(text)
    return "\n".join(parts)


def detect_notice_number(sheet: Worksheet) -> str | None:
    header = _scan_header_text(sheet)
    match = NOTICE_RE.search(header)
    return match.group(1) if match else None


def detect_sheet_no(sheet: Worksheet) -> int | None:
    header = _scan_header_text(sheet)
    match = SHEET_RE.search(header)
    if match:
        return safe_int(match.group(1))

    normalized_title = normalize_for_match(sheet.title)
    return safe_int(normalized_title)


def parse_sheet(sheet: Worksheet, sheet_index: int, start_global_seq: int) -> tuple[SheetResult, int, bool]:
    classification = classify_sheet(sheet)
    notice_number = detect_notice_number(sheet)
    sheet_no = detect_sheet_no(sheet)

    changes, next_seq = extract_changes(sheet, sheet_index=sheet_index, start_global_seq=start_global_seq)

    result = SheetResult(
        sheet_index=sheet_index,
        sheet_name=sheet.title,
        sheet_kind=classification.kind,
        sheet_no_detected=sheet_no,
        sheet_local_header=SheetLocalHeader(notice_number=notice_number),
        changes=changes,
    )
    return result, next_seq, classification.is_candidate
