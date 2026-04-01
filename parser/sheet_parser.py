from __future__ import annotations

import re

from openpyxl.worksheet.worksheet import Worksheet

from parser.change_extractor import DOC_CODE_RE, extract_changes
from parser.merged_cells import iter_sheet_rows
from parser.models import SheetLocalHeader, SheetResult
from parser.normalizer import normalize_for_match, normalize_text, safe_int
from parser.sheet_classifier import classify_sheet

NOTICE_NUMBER_RE = re.compile(r"извещение\s*(?:об\s*изменении)?\s*№\s*([A-ZА-Я0-9][A-ZА-Я0-9\-/]{2,})", re.IGNORECASE)
NOTICE_LINE_RE = re.compile(r"\b№\s*([A-ZА-Я0-9][A-ZА-Я0-9\-/]{2,})", re.IGNORECASE)
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


def _is_valid_notice_candidate(value: str | None) -> tuple[bool, str | None]:
    if not value:
        return False, "empty"
    norm = normalize_for_match(value)
    if norm in INVALID_NOTICE_VALUES:
        return False, "header_label"
    if DOC_CODE_RE.fullmatch(value):
        return False, "looks_like_doc_code"
    if not any(ch.isdigit() for ch in value):
        return False, "no_digits"
    return True, None


def detect_notice_number(sheet: Worksheet, sheet_kind: str) -> tuple[str | None, dict[str, object]]:
    rows = _scan_header_rows(sheet)
    candidates: list[str] = []
    rejects: list[dict[str, str]] = []

    # Per requirement: attempt extraction from full-sheet header only.
    if sheet_kind != "full":
        return None, {
            "notice_candidates": [],
            "rejected_notice_candidates": [{"value": "<skipped>", "reason": "not_full_sheet"}],
        }

    for _, text in rows:
        for pattern in (NOTICE_NUMBER_RE, NOTICE_LINE_RE):
            match = pattern.search(text)
            if not match:
                continue
            candidate = match.group(1)
            candidates.append(candidate)
            valid, reason = _is_valid_notice_candidate(candidate)
            if valid:
                return candidate, {
                    "notice_candidates": candidates,
                    "rejected_notice_candidates": rejects,
                }
            rejects.append({"value": candidate, "reason": reason or "unknown"})

    return None, {
        "notice_candidates": candidates,
        "rejected_notice_candidates": rejects,
    }


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
    notice_number, notice_debug = detect_notice_number(sheet, classification.kind)
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

    debug_payload = {
        **extraction_debug,
        "notice_candidates": notice_debug.get("notice_candidates", []),
        "rejected_notice_candidates": notice_debug.get("rejected_notice_candidates", []),
    }
    return result, next_seq, classification.is_candidate, debug_payload
