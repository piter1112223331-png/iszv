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


def _extract_code_dedicated(rows: list[tuple[int, list[str | None]]]) -> tuple[str | None, list[str]]:
    candidates: list[tuple[int, str]] = []
    windows: list[str] = []
    for _, row in rows[:80]:
        norm_cells = [normalize_for_match(v) for v in row]
        row_has_sheet = any("лист" in c for c in norm_cells)
        for col_idx, c in enumerate(norm_cells):
            if c != "код":
                continue
            val = normalize_text(row[col_idx + 1] if col_idx + 1 < len(row) else None)
            if not val:
                continue
            windows.append(val)
            if re.fullmatch(r"\d+", val):
                score = 1 if row_has_sheet else 2
                candidates.append((score, val))
            elif not _is_label_like(val):
                candidates.append((1, val))
    if not candidates:
        return None, windows
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1], windows


def _extract_dash_field(rows: list[tuple[int, list[str | None]]], label: str) -> tuple[str | None, list[str]]:
    windows: list[str] = []
    for _, row in rows[:120]:
        norm = [normalize_for_match(v) for v in row]
        for col_idx, c in enumerate(norm):
            if label not in c:
                continue
            for off in (1, 2):
                if col_idx + off < len(row):
                    val = normalize_text(row[col_idx + off])
                    if val:
                        windows.append(val)
                        if val in {"-", "—", "–"}:
                            return "-", windows
    return None, windows


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
    # strict right-side window + one-row-below fallback in approval block
    values: list[str] = []
    for delta_row in (0, 1):
        if row_idx + delta_row >= len(rows):
            continue
        row = rows[row_idx + delta_row][1]
        window = [normalize_text(row[i]) for i in range(anchor_col, min(len(row), anchor_col + 3))]
        values.extend([c for c in window if c])
        for cand in window:
            if not cand:
                continue
            cand_norm = normalize_for_match(cand)
            if _is_label_like(cand) or "контроль" in cand_norm or "предст" in cand_norm:
                continue
            return cand, None, values
    return None, ("label_like" if values else "empty"), values


def _extract_approvals(rows: list[tuple[int, list[str | None]]]) -> tuple[Approvals, dict[str, object]]:
    approvals = Approvals()
    debug = {"found": [], "missing": [], "approvals_candidates": {}, "approvals_rejected_reasons": {}}
    tail = rows[max(0, len(rows) - 80) :]

    for ridx, (_, row) in enumerate(tail):
        norm_row = [normalize_for_match(v) for v in row]
        for field, anchors in APPROVAL_ANCHORS.items():
            if getattr(approvals, field) is not None:
                continue
            for col_idx, text in enumerate(norm_row, start=1):
                if text and all(a in text for a in anchors):
                    val, reason, window = _extract_approval_value(tail, ridx, col_idx)
                    debug["approvals_candidates"][field] = normalize_text(row[col_idx] if col_idx < len(row) else None)
                    debug.setdefault("approvals_value_windows", {})[field] = window
                    if val:
                        setattr(approvals, field, val)
                    else:
                        debug["approvals_rejected_reasons"][field] = reason
                    break

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
        "developer_search_window": [],
        "code_search_window": [],
        "sheet_no_declared_search_window": [],
        "sheet_no_value_window": [],
        "code_value_window": [],
        "dash_field_value_windows": {},
        "approvals_value_windows_expanded": {},
        "approvals_final_candidates": {},
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
    debug["sheet_no_declared_search_window"] = sheet_dbg.get("sheet_no_declared_search_window", [])
    debug["sheet_no_value_window"] = [sheet_dbg.get("sheet_no_declared_candidate")] if sheet_dbg.get("sheet_no_declared_candidate") is not None else []
    debug["fields"]["sheet_no_declared"] = {
        "raw_candidate": sheet_dbg.get("sheet_no_declared_candidate"),
        "cleaned_candidate": sheet_dbg.get("sheet_no_declared_candidate"),
        "rejected_reason": None if sheet_no is not None else "not_found",
        "final_value": header.sheet_no_declared,
        "collapsed_repeats_applied": False,
    }

    code_val, code_windows = _extract_code_dedicated(rows)
    if code_val is not None:
        header.code = code_val
        debug["fields"]["code"] = {"raw_candidate": code_val, "cleaned_candidate": code_val, "rejected_reason": None, "final_value": code_val, "collapsed_repeats_applied": False}
        debug["normalized_header_fields"]["code"] = code_val
    debug["code_value_window"] = code_windows

    for fld, lbl in (("implementation_instruction", "указания о внедрении"), ("applicability", "применяемость"), ("distribution", "разослать")):
        dash_val, dash_windows = _extract_dash_field(rows, lbl)
        debug["dash_field_value_windows"][fld] = dash_windows
        if dash_val is not None:
            setattr(header, fld, dash_val)
            debug["normalized_header_fields"][fld] = dash_val
            debug["fields"][fld] = {"raw_candidate": dash_val, "cleaned_candidate": dash_val, "rejected_reason": None, "final_value": dash_val, "collapsed_repeats_applied": False}

    for i, (_, row_values) in enumerate(rows[:80]):
        row_norm = [normalize_for_match(v) for v in row_values]
        row_txt = normalize_text(" ".join(v for v in row_values if v))
        if row_txt:
            debug["developer_search_window"].append(row_txt)
            debug["code_search_window"].append(row_txt)

        for field, anchors in HEADER_ANCHORS.items():
            if field in {"sheet_no_declared", "code", "implementation_instruction", "applicability", "distribution"} or getattr(header, field) is not None:
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

    if header.developer is None:
        for _, row in rows[:30]:
            for cell in row:
                text = normalize_text(cell)
                if text and (('АО "' in text) or ('ООО' in text) or ('ПАО' in text)):
                    header.developer = text
                    debug["fields"]["developer"] = {"raw_candidate": text, "cleaned_candidate": text, "rejected_reason": None, "final_value": text, "collapsed_repeats_applied": False}
                    debug["normalized_header_fields"]["developer"] = text
                    break
            if header.developer:
                break

    approvals, appr_debug = _extract_approvals(rows)
    debug["approvals_found"] = appr_debug["found"]
    debug["approvals_missing"] = appr_debug["missing"]
    debug["approvals_candidates"] = appr_debug["approvals_candidates"]
    debug["approvals_rejected_reasons"] = appr_debug["approvals_rejected_reasons"]
    debug["approvals_value_windows_expanded"] = appr_debug.get("approvals_value_windows", {})
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
