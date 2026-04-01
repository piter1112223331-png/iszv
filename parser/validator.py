from __future__ import annotations

from dataclasses import asdict

from parser.models import ParsedDocument, ValidationResult


IMPORTANT_HEADER_FIELDS = (
    "reason",
    "release_date",
    "stock_instruction",
    "implementation_instruction",
    "applicability",
    "distribution",
)


def validate_document(doc: ParsedDocument) -> ValidationResult:
    errors: list[str] = []
    warnings: list[str] = []

    if doc.sheet_count_detected == 0:
        errors.append("no_candidate_sheets")

    if not doc.all_changes:
        errors.append("no_change_blocks")

    if not doc.notice_number:
        warnings.append("notice_number_missing")

    header_dict = asdict(doc.document_header)
    for field in IMPORTANT_HEADER_FIELDS:
        if header_dict.get(field) is None:
            warnings.append(f"header_field_missing:{field}")

    if doc.document_header.sheet_total_declared is None:
        warnings.append("sheet_total_declared_missing")
    elif doc.document_header.sheet_total_declared != doc.sheet_count_detected:
        warnings.append("sheet_count_mismatch")

    for change in doc.all_changes:
        label = f"sheet={change.sheet_index}:seq={change.change_seq_on_sheet}"
        if not change.doc_code:
            errors.append(f"empty_doc_code:{label}")
        if not change.change_index:
            errors.append(f"empty_change_index:{label}")
        if not change.change_text:
            warnings.append(f"empty_change_text:{label}")

        z = change.zone_ref
        if z.body_row_start is None or z.body_row_end is None or (z.body_row_end < z.body_row_start):
            warnings.append(f"suspicious_block_boundary:{label}")

    return ValidationResult(template_detected=doc.sheet_count_detected > 0, errors=errors, warnings=warnings)
