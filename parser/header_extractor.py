from __future__ import annotations

import re
from dataclasses import asdict

from openpyxl.worksheet.worksheet import Worksheet

from parser.merged_cells import iter_sheet_rows
from parser.models import DocumentHeader
from parser.normalizer import normalize_for_match, normalize_text, safe_int

HEADER_ANCHORS = {
    "sender": ("предприятие", "организация"),
    "reason": ("причина",),
    "code": ("код",),
    "release_center": ("центр", "выпуска"),
    "release_date": ("дата выпуска",),
    "stock_instruction": ("указание о заделе",),
    "implementation_instruction": ("указания о внедрении",),
    "applicability": ("применяемость",),
    "distribution": ("разослать",),
}
SHEET_TOTAL_RE = re.compile(r"(?:лист|листов)\D{0,8}(\d+)", re.IGNORECASE)


def _extract_value_after_anchor(row_values: list[str | None], anchor_col: int) -> str | None:
    right_values = [normalize_text(v) for v in row_values[anchor_col:]]
    right_values = [v for v in right_values if v]
    if right_values:
        return normalize_text(" ".join(right_values))
    return None


def extract_document_header(sheet: Worksheet | None) -> tuple[DocumentHeader, dict[str, object]]:
    header = DocumentHeader()
    debug: dict[str, object] = {"found_fields": [], "missing_fields": []}
    if sheet is None:
        debug["missing_fields"] = list(asdict(header).keys())
        return header, debug

    scanned_rows = list(iter_sheet_rows(sheet))[:140]

    for i, (row_idx, row_values) in enumerate(scanned_rows):
        row_norm = [normalize_for_match(v) for v in row_values]
        row_text = normalize_text(" ".join(v for v in row_values if v)) or ""

        # sheet total declared
        if header.sheet_total_declared is None:
            m = SHEET_TOTAL_RE.search(row_text)
            if m:
                header.sheet_total_declared = safe_int(m.group(1))

        for field, anchors in HEADER_ANCHORS.items():
            if getattr(header, field) is not None:
                continue
            for col_idx, norm in enumerate(row_norm, start=1):
                if not norm:
                    continue
                if all(anchor in norm for anchor in anchors):
                    value = _extract_value_after_anchor(row_values, col_idx)
                    if not value and i + 1 < len(scanned_rows):
                        # fallback: next row same zone
                        _, next_row = scanned_rows[i + 1]
                        value = _extract_value_after_anchor(next_row, col_idx)
                    if value:
                        setattr(header, field, value)
                    break

    fields = asdict(header)
    debug["found_fields"] = [k for k, v in fields.items() if v is not None]
    debug["missing_fields"] = [k for k, v in fields.items() if v is None]
    return header, debug
