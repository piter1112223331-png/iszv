from __future__ import annotations

import re
from dataclasses import asdict

from openpyxl.worksheet.worksheet import Worksheet

from parser.merged_cells import iter_sheet_rows
from parser.models import Approvals, DocumentHeader
from parser.normalizer import (
    collapse_repeated_phrases,
    collapse_repeated_tokens,
    normalize_dash_noise,
    normalize_for_match,
    normalize_text,
    safe_int,
)

HEADER_ANCHORS = {
    "developer": ("предприятие", "организация"),
    "notice_number": ("извещение",),
    "reason": ("причина",),
    "code": ("код",),
    "sheet_no_declared": ("лист",),
    "release_center": ("центр", "выпуска"),
    "release_date": ("дата выпуска",),
    "stock_instruction": ("указание о заделе",),
    "implementation_instruction": ("указания о внедрении", "указание о внедрении"),
    "applicability": ("применяемость",),
    "distribution": ("разослать",),
}
FIELD_WIDTH = {
    "developer": 4,
    "notice_number": 2,
    "reason": 3,
    "code": 1,
    "sheet_no_declared": 1,
    "release_center": 2,
    "release_date": 1,
    "stock_instruction": 4,
    "implementation_instruction": 4,
    "applicability": 3,
    "distribution": 3,
}
HEADER_NOISE_MARKERS = (
    "причина",
    "код",
    "лист",
    "листов",
    "дата выпуска",
    "срок изм",
    "срок действия пи",
    "указание о заделе",
    "указания о внедрении",
    "применяемость",
    "разослать",
    "извещение",
    "составил",
    "проверил",
    "н. контроль",
    "утвердил",
    "т. контроль",
    "предст. заказ",
)
ORG_PREFIXES = ("АО", "ООО", "ПАО")
APPROVAL_ANCHORS = {
    "author": ("составил",),
    "reviewer": ("проверил",),
    "norm_control": ("н. контроль",),
    "approver": ("утвердил",),
}
DOC_CODE_LIKE_RE = re.compile(r"\b[А-ЯA-Z]{2,}\.[0-9]{4}\.[0-9]{3,4}\.[0-9]{4,5}\b")
DATE_RE = re.compile(r"\b\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b")
NOTICE_RE = re.compile(r"\b[А-ЯA-Z]{1,3}\.[0-9]{4}\.[0-9]{4}\b")


def _is_label_like(text: str | None) -> bool:
    norm = normalize_for_match(text)
    if not norm:
        return False
    return any(marker in norm for marker in HEADER_NOISE_MARKERS)


def _extract_right_zone(row_values: list[str | None], anchor_col: int, width: int) -> str | None:
    cells = row_values[anchor_col : anchor_col + width]
    parts = [normalize_text(v) for v in cells if normalize_text(v) and not _is_label_like(v)]
    return normalize_text(" ".join(parts)) if parts else None


def _local_window_values(
    rows: list[tuple[int, list[str | None]]],
    row_idx: int,
    anchor_col_idx: int,
    offsets: tuple[int, ...] = (1, 2),
    include_next_row: bool = True,
) -> list[str]:
    values: list[str] = []
    row_jumps = (0, 1) if include_next_row else (0,)
    for row_jump in row_jumps:
        target_idx = row_idx + row_jump
        if target_idx >= len(rows):
            continue
        row = rows[target_idx][1]
        for off in offsets:
            col_idx = anchor_col_idx + off
            if col_idx < 0 or col_idx >= len(row):
                continue
            val = normalize_text(row[col_idx])
            if val:
                values.append(val)
    return values


def _post_process_field(field: str, value: str | None, row_norm: list[str]) -> tuple[str | None, bool]:
    original = value
    value = normalize_dash_noise(value)
    value = collapse_repeated_tokens(value)
    if field in {"reason", "stock_instruction", "implementation_instruction", "applicability", "distribution"}:
        value = collapse_repeated_phrases(value)

    if field == "code" and value:
        if re.fullmatch(r"\d+\s+\d+", value):
            value = None
    if field == "release_date" and value and not DATE_RE.search(value):
        value = None
    if field == "notice_number":
        if not value:
            return None, value != original
        if DOC_CODE_LIKE_RE.search(value):
            value = None
        elif NOTICE_RE.search(value):
            value = NOTICE_RE.search(value).group(0)
        else:
            value = None

    return value, value != original


def _sanitize_candidate(field: str, raw: str | None, row_norm: list[str]) -> tuple[str | None, str | None, bool]:
    cleaned = normalize_text(raw)
    if not cleaned:
        return None, "empty", False

    if field in {"implementation_instruction", "applicability", "distribution"} and cleaned in {"-", "—", "–"}:
        return "-", None, False

    norm = normalize_for_match(cleaned)
    if _is_label_like(cleaned) and len(cleaned.split()) <= 8:
        return None, "header_like_noise", False

    meaningful = [w for w in norm.split() if all(marker not in w for marker in HEADER_NOISE_MARKERS)]
    if not meaningful and not any(ch.isdigit() for ch in cleaned):
        return None, "no_meaningful_tokens", False

    processed, collapsed = _post_process_field(field, cleaned, row_norm)
    if processed is None:
        return None, "postprocess_rejected", collapsed
    return processed, None, collapsed


def _find_header_anchor(
    rows: list[tuple[int, list[str | None]]],
    anchors: tuple[str, ...],
    *,
    top_rows: int = 80,
    all_words: bool = True,
) -> tuple[int, int, str] | None:
    for ridx, (excel_row, row) in enumerate(rows[:top_rows]):
        norm_cells = [normalize_for_match(v) for v in row]
        for cidx, cell_norm in enumerate(norm_cells):
            if not cell_norm:
                continue
            matched = all(a in cell_norm for a in anchors) if all_words else any(a in cell_norm for a in anchors)
            if matched:
                return ridx, cidx, f"R{excel_row}C{cidx + 1}"
    return None


def _extract_header_field_local(
    rows: list[tuple[int, list[str | None]]],
    field: str,
    anchors: tuple[str, ...],
    *,
    all_words: bool = True,
    offsets: tuple[int, ...] = (1, 2),
    include_next_row: bool = True,
) -> tuple[str | None, dict[str, object]]:
    debug_info: dict[str, object] = {"anchor_cell": None, "value_window": [], "raw_candidate": None}
    anchor = _find_header_anchor(rows, anchors, top_rows=80, all_words=all_words)
    if not anchor:
        return None, debug_info
    ridx, cidx, anchor_cell = anchor
    debug_info["anchor_cell"] = anchor_cell
    window = _local_window_values(rows, ridx, cidx, offsets=offsets, include_next_row=include_next_row)
    debug_info["value_window"] = window
    for candidate in window:
        if _is_label_like(candidate):
            continue
        if field == "code":
            # exclude sheet counters caught near "Лист/Листов"
            if not re.fullmatch(r"\d+", candidate):
                continue
            left_ctx = normalize_for_match(rows[ridx][1][max(0, cidx - 1)] if cidx > 0 else None)
            if "лист" in left_ctx:
                continue
        if field == "developer" and not any(candidate.startswith(prefix) for prefix in ORG_PREFIXES):
            continue
        debug_info["raw_candidate"] = candidate
        return candidate, debug_info
    return None, debug_info


def _extract_sheet_number_local(rows: list[tuple[int, list[str | None]]], target: str) -> tuple[int | None, dict[str, object]]:
    debug_info: dict[str, object] = {"anchor_cell": None, "value_window": [], "raw_candidate": None}
    anchor = _find_header_anchor(rows, (target,), top_rows=80, all_words=False)
    if not anchor:
        return None, debug_info
    ridx, cidx, anchor_cell = anchor
    debug_info["anchor_cell"] = anchor_cell
    window = _local_window_values(rows, ridx, cidx, offsets=(1, 2), include_next_row=True)
    debug_info["value_window"] = window
    for candidate in window:
        if " " in candidate:
            continue
        val = safe_int(candidate)
        if val is None:
            continue
        debug_info["raw_candidate"] = candidate
        return val, debug_info
    return None, debug_info


def _extract_sheet_numbers(rows: list[tuple[int, list[str | None]]]) -> tuple[int | None, int | None, dict[str, object]]:
    dbg = {"sheet_no_declared_candidate": None, "sheet_total_declared_candidate": None, "sheet_no_declared_search_window": []}
    sheet_no = None
    sheet_total = None

    for _, row in rows[:80]:
        row_text = normalize_for_match(" ".join(v for v in row if v))
        if row_text:
            dbg["sheet_no_declared_search_window"].append(row_text)

        pair = re.search(r"лист\s*(\d+)\s*листов\s*(\d+)", row_text)
        if pair:
            sheet_no = int(pair.group(1))
            sheet_total = int(pair.group(2))
            dbg["sheet_no_declared_candidate"] = sheet_no
            dbg["sheet_total_declared_candidate"] = sheet_total
            break

        norm_cells = [normalize_for_match(v) for v in row]
        for col_idx, c in enumerate(norm_cells):
            if c == "лист" or c.startswith("лист "):
                right = safe_int(normalize_text(row[col_idx + 1] if col_idx + 1 < len(row) else None))
                if right is not None:
                    sheet_no = right
                    dbg["sheet_no_declared_candidate"] = right
                    break
        for col_idx, c in enumerate(norm_cells):
            if "листов" in c:
                right = safe_int(normalize_text(row[col_idx + 1] if col_idx + 1 < len(row) else None))
                if right is not None:
                    sheet_total = right
                    dbg["sheet_total_declared_candidate"] = right

        if sheet_no is not None and sheet_total is not None:
            break

    if sheet_no is None and sheet_total == 1:
        sheet_no = 1
        dbg["sheet_no_declared_candidate"] = 1

    return sheet_no, sheet_total, dbg


def _extract_sheet_total_declared(rows: list[tuple[int, list[str | None]]]) -> tuple[int | None, dict[str, object]]:
    debug = {
        "raw_candidate": None,
        "cleaned_candidate": None,
        "rejected_reason": None,
        "final_value": None,
        "collapsed_repeats_applied": False,
    }

    for i, (_, row_values) in enumerate(rows):
        for col_idx, cell in enumerate(row_values, start=1):
            norm = normalize_for_match(cell)
            if "листов" not in norm:
                continue

            right = normalize_text(row_values[col_idx] if col_idx < len(row_values) else None)
            below = None
            if i + 1 < len(rows):
                below = normalize_text(rows[i + 1][1][col_idx - 1] if col_idx - 1 < len(rows[i + 1][1]) else None)

            for candidate in (right, below):
                if not candidate:
                    continue
                debug["raw_candidate"] = candidate
                value = safe_int(candidate)
                debug["cleaned_candidate"] = candidate
                if value is None or value > 999:
                    debug["rejected_reason"] = "not_short_numeric"
                    continue
                debug["final_value"] = value
                return value, debug

    debug["rejected_reason"] = "not_found"
    return None, debug


def _extract_approval_value(rows: list[tuple[int, list[str | None]]], row_idx: int, anchor_col: int) -> tuple[str | None, str | None, list[str]]:
    values = _local_window_values(rows, row_idx, anchor_col - 1, offsets=(1, 2, 3), include_next_row=True)
    for cand in values:
        cand_norm = normalize_for_match(cand)
        if _is_label_like(cand) or "контроль" in cand_norm or "предст" in cand_norm:
            continue
        return cand, None, values
    return None, ("label_like" if values else "empty"), values


def _extract_approvals(rows: list[tuple[int, list[str | None]]]) -> tuple[Approvals, dict[str, object]]:
    approvals = Approvals()
    debug = {
        "found": [],
        "missing": [],
        "approvals_candidates": {},
        "approvals_rejected_reasons": {},
        "approval_block_detected": False,
        "approval_anchor_cells": {},
        "approvals_raw_candidates": {},
    }
    tail = rows[max(0, len(rows) - 80) :]

    anchor_hits = 0
    for ridx, (_, row) in enumerate(tail):
        norm_row = [normalize_for_match(v) for v in row]
        for field, anchors in APPROVAL_ANCHORS.items():
            if getattr(approvals, field) is not None:
                continue
            for col_idx, text in enumerate(norm_row, start=1):
                if text and all(a in text for a in anchors):
                    debug["approval_anchor_cells"][field] = f"R{tail[ridx][0]}C{col_idx}"
                    anchor_hits += 1
                    val, reason, window = _extract_approval_value(tail, ridx, col_idx)
                    debug["approvals_candidates"][field] = normalize_text(row[col_idx] if col_idx < len(row) else None)
                    debug.setdefault("approvals_value_windows", {})[field] = window
                    debug["approvals_raw_candidates"][field] = val
                    if val:
                        setattr(approvals, field, val)
                    else:
                        debug["approvals_rejected_reasons"][field] = reason
                    break
    debug["approval_block_detected"] = anchor_hits >= 2

    for k, v in asdict(approvals).items():
        (debug["found"] if v else debug["missing"]).append(k)
    return approvals, debug


def extract_document_header(sheet: Worksheet | None) -> tuple[DocumentHeader, Approvals, dict[str, object]]:
    header = DocumentHeader()
    approvals = Approvals()
    debug: dict[str, object] = {
        "found_fields": [],
        "missing_fields": [],
        "fields": {},
        "normalized_header_fields": {},
        "collapsed_repeats_applied": [],
        "approvals_found": [],
        "approvals_missing": [],
        "approvals_candidates": {},
        "approvals_rejected_reasons": {},
        "approvals_value_windows": {},
        "header_anchor_cells": {},
    }
    if sheet is None:
        debug["missing_fields"] = list(asdict(header).keys())
        return header, approvals, debug

    rows = list(iter_sheet_rows(sheet))[:180]

    sheet_no, sheet_total_hint, sheet_dbg = _extract_sheet_numbers(rows)
    if sheet_no is not None:
        header.sheet_no_declared = sheet_no
    sheet_total, sheet_total_debug = _extract_sheet_total_declared(rows)
    header.sheet_total_declared = sheet_total if sheet_total is not None else sheet_total_hint
    debug["fields"]["sheet_total_declared"] = sheet_total_debug
    debug["fields"]["sheet_no_declared"] = {
        "raw_candidate": sheet_dbg.get("sheet_no_declared_candidate"),
        "cleaned_candidate": sheet_dbg.get("sheet_no_declared_candidate"),
        "rejected_reason": None if sheet_no is not None else "not_found",
        "final_value": header.sheet_no_declared,
        "collapsed_repeats_applied": False,
    }

    # strict local-zone extraction for unstable business fields
    for field, anchors, all_words in (
        ("developer", HEADER_ANCHORS["developer"], False),
        ("code", ("код",), True),
        ("implementation_instruction", ("указания о внедрении", "указание о внедрении"), False),
        ("applicability", ("применяемость",), True),
        ("distribution", ("разослать",), True),
    ):
        candidate, local_dbg = _extract_header_field_local(rows, field, anchors, all_words=all_words, offsets=(1, 2), include_next_row=True)
        debug["header_anchor_cells"][field] = local_dbg["anchor_cell"]
        cleaned_candidate, reject_reason, collapsed = _sanitize_candidate(field, candidate, [])
        if cleaned_candidate is not None and reject_reason is None:
            setattr(header, field, cleaned_candidate)
            debug["normalized_header_fields"][field] = cleaned_candidate
        debug["fields"][field] = {
            "anchor_cell": local_dbg["anchor_cell"],
            "value_window": local_dbg["value_window"],
            "raw_candidate": local_dbg["raw_candidate"],
            "cleaned_candidate": cleaned_candidate,
            "rejected_reason": reject_reason if local_dbg["anchor_cell"] else "anchor_not_found",
            "final_value": getattr(header, field),
            "collapsed_repeats_applied": collapsed,
        }

    sheet_no_local, sheet_no_dbg = _extract_sheet_number_local(rows, "лист")
    debug["header_anchor_cells"]["sheet_no_declared"] = sheet_no_dbg["anchor_cell"]
    if sheet_no_local is not None:
        header.sheet_no_declared = sheet_no_local
    debug["fields"]["sheet_no_declared"] = {
        "anchor_cell": sheet_no_dbg["anchor_cell"],
        "value_window": sheet_no_dbg["value_window"],
        "raw_candidate": sheet_no_dbg["raw_candidate"],
        "cleaned_candidate": sheet_no_dbg["raw_candidate"],
        "rejected_reason": None if header.sheet_no_declared is not None else "not_found",
        "final_value": header.sheet_no_declared,
        "collapsed_repeats_applied": False,
    }

    sheet_total_local, sheet_total_local_dbg = _extract_sheet_number_local(rows, "листов")
    debug["header_anchor_cells"]["sheet_total_declared"] = sheet_total_local_dbg["anchor_cell"]
    if sheet_total_local is not None:
        header.sheet_total_declared = sheet_total_local
        debug["fields"]["sheet_total_declared"]["final_value"] = sheet_total_local
        debug["fields"]["sheet_total_declared"]["raw_candidate"] = sheet_total_local_dbg["raw_candidate"]
        debug["fields"]["sheet_total_declared"]["cleaned_candidate"] = sheet_total_local_dbg["raw_candidate"]
        debug["fields"]["sheet_total_declared"]["rejected_reason"] = None
    debug["fields"]["sheet_total_declared"]["anchor_cell"] = sheet_total_local_dbg["anchor_cell"]
    debug["fields"]["sheet_total_declared"]["value_window"] = sheet_total_local_dbg["value_window"]

    for i, (_, row_values) in enumerate(rows[:80]):
        row_norm = [normalize_for_match(v) for v in row_values]
        row_txt = normalize_text(" ".join(v for v in row_values if v))
        for field, anchors in HEADER_ANCHORS.items():
            if field in {"developer", "sheet_no_declared", "sheet_total_declared", "code", "implementation_instruction", "applicability", "distribution"} or getattr(header, field) is not None:
                continue
            for col_idx, norm in enumerate(row_norm, start=1):
                if not norm:
                    continue
                if field in {"developer", "implementation_instruction"}:
                    if not any(a in norm for a in anchors):
                        continue
                else:
                    if not all(anchor in norm for anchor in anchors):
                        continue

                width = FIELD_WIDTH.get(field, 3)
                raw_candidate = _extract_right_zone(row_values, col_idx, width)
                if not raw_candidate and i + 1 < len(rows):
                    raw_candidate = _extract_right_zone(rows[i + 1][1], col_idx, width)

                cleaned_candidate, reject_reason, collapsed = _sanitize_candidate(field, raw_candidate, row_norm)
                if reject_reason is None and cleaned_candidate:
                    setattr(header, field, cleaned_candidate)
                    debug["normalized_header_fields"][field] = getattr(header, field)

                debug["fields"][field] = {
                    "raw_candidate": raw_candidate,
                    "cleaned_candidate": cleaned_candidate,
                    "rejected_reason": reject_reason,
                    "final_value": getattr(header, field),
                    "collapsed_repeats_applied": collapsed,
                }
                if collapsed:
                    debug["collapsed_repeats_applied"].append(field)
                if field == "code" and getattr(header, field) is None:
                    continue
                break

    approvals, appr_debug = _extract_approvals(rows)
    debug["approvals_found"] = appr_debug["found"]
    debug["approvals_missing"] = appr_debug["missing"]
    debug["approvals_candidates"] = appr_debug["approvals_candidates"]
    debug["approvals_rejected_reasons"] = appr_debug["approvals_rejected_reasons"]
    debug["approval_block_detected"] = appr_debug.get("approval_block_detected", False)
    debug["approval_anchor_cells"] = appr_debug.get("approval_anchor_cells", {})
    debug["approvals_value_windows"] = appr_debug.get("approvals_value_windows", {})
    debug["approvals_raw_candidates"] = appr_debug.get("approvals_raw_candidates", {})
    debug["approvals_final_candidates"] = asdict(approvals)

    fields = asdict(header)
    debug["found_fields"] = [k for k, v in fields.items() if v is not None]
    debug["missing_fields"] = [k for k, v in fields.items() if v is None]

    for k in fields:
        if k not in debug["fields"]:
            debug["fields"][k] = {
                "raw_candidate": None,
                "cleaned_candidate": None,
                "rejected_reason": "anchor_not_found",
                "final_value": fields[k],
                "collapsed_repeats_applied": False,
            }

    return header, approvals, debug
