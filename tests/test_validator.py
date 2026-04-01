from parser.models import (
    ChangeBlock,
    DocumentHeader,
    ParsedDocument,
    SheetLocalHeader,
    SheetResult,
    ValidationResult,
    ZoneRef,
)
from parser.validator import validate_document


def _doc_with_changes(changes):
    return ParsedDocument(
        document_type="change_notice",
        template_version="notice_multi_sheet_v1",
        source_file="x.xlsx",
        notice_number=None,
        sheet_count_detected=1,
        document_header=DocumentHeader(sheet_total_declared=2),
        sheets=[
            SheetResult(
                sheet_index=1,
                sheet_name="1",
                sheet_kind="full",
                sheet_no_detected=1,
                sheet_local_header=SheetLocalHeader(notice_number=None),
                changes=changes,
            )
        ],
        all_changes=changes,
        validation=ValidationResult(),
    )


def test_no_change_blocks_error():
    doc = _doc_with_changes([])
    result = validate_document(doc)
    assert "no_change_blocks" in result.errors


def test_empty_change_text_warning_and_empty_fields_errors():
    change = ChangeBlock(
        sheet_index=1,
        change_seq_global=1,
        change_seq_on_sheet=1,
        change_index=None,
        doc_code=None,
        change_text=None,
        raw_meta_text=None,
        zone_ref=ZoneRef(),
    )
    result = validate_document(_doc_with_changes([change]))
    assert any(e.startswith("empty_doc_code") for e in result.errors)
    assert any(e.startswith("empty_change_index") for e in result.errors)
    assert any(w.startswith("empty_change_text") for w in result.warnings)


def test_notice_missing_and_sheet_count_mismatch_warnings():
    change = ChangeBlock(
        sheet_index=1,
        change_seq_global=1,
        change_seq_on_sheet=1,
        change_index="1",
        doc_code="ЕСРТ.0016.0000.0001",
        change_text="ok",
        raw_meta_text="ok",
        zone_ref=ZoneRef(body_row_start=1, body_row_end=2),
    )
    result = validate_document(_doc_with_changes([change]))
    assert "notice_number_missing" in result.warnings
    assert "sheet_count_mismatch" in result.warnings
    assert "header_field_missing:sender" in result.warnings
    assert "header_field_missing:release_center" in result.warnings
