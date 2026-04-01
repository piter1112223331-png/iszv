from parser.models import (
    Approvals,
    ChangeBlock,
    DocumentHeader,
    ParsedDocument,
    SheetLocalHeader,
    SheetResult,
    ValidationResult,
    ZoneRef,
)
from parser.payloads import enrich_document


def _doc(errors=None, warnings=None):
    ch = ChangeBlock(
        sheet_index=1,
        change_seq_global=1,
        change_seq_on_sheet=1,
        change_index="1",
        doc_code="ЕСРТ.0016.0000.0001",
        change_text="Текст",
        raw_meta_text="raw",
        zone_ref=ZoneRef(body_row_start=1, body_row_end=2),
    )
    return ParsedDocument(
        document_type="change_notice",
        template_version="notice_multi_sheet_v1",
        source_file="f.xlsx",
        notice_number="ИИ.0000.0001",
        sheet_count_detected=1,
        document_header=DocumentHeader(
            developer='АО "ИЦ КТ"',
            notice_number="ИИ.0000.0001",
            reason="Причина",
            code="7",
            sheet_no_declared=1,
            sheet_total_declared=1,
            stock_instruction="Нет",
            implementation_instruction="-",
            applicability="-",
            distribution="-",
        ),
        sheets=[
            SheetResult(
                sheet_index=1,
                sheet_name="1",
                sheet_kind="full",
                sheet_no_detected=1,
                sheet_local_header=SheetLocalHeader(),
                changes=[ch],
            )
        ],
        all_changes=[ch],
        validation=ValidationResult(template_detected=True, errors=errors or [], warnings=warnings or []),
        approvals=Approvals(author="Петров", reviewer="Иванов", norm_control="Свердлов", approver="Алербов"),
    )


def test_status_computation():
    d = enrich_document(_doc(errors=["e1"]))
    assert d.status == "error"
    d = enrich_document(_doc(warnings=["w1"]))
    assert d.status == "warning"
    d = enrich_document(_doc())
    assert d.status == "ok"


def test_summary_counters():
    d = enrich_document(_doc(warnings=["w1"], errors=["e1"]))
    assert d.summary["change_blocks_count"] == 1
    assert d.summary["warnings_count"] == 1
    assert d.summary["errors_count"] == 1


def test_llm_payload_generation_and_no_duplicate_changes():
    d = enrich_document(_doc())
    payload = d.llm_payload
    assert len(payload["changes"]) == 1
    assert "raw_meta_text" not in payload["changes"][0]
    assert "zone_ref" not in payload["changes"][0]
    assert payload["document_header_compact"]["developer"] == 'АО "ИЦ КТ"'
    assert payload["approvals"]["author"] == "Петров"


def test_flowis_payload_generation_with_approvals():
    d = enrich_document(_doc(warnings=["w1"]))
    f = d.flowis_payload
    assert f["status"] == "warning"
    assert f["changes_count"] == 1
    assert f["developer"] == 'АО "ИЦ КТ"'
    assert f["approvals_author"] == "Петров"


def test_validation_details_generation():
    d = enrich_document(_doc(errors=["empty_doc_code:sheet=1:seq=1"], warnings=["header_field_missing:developer"]))
    assert d.validation_details["errors"][0]["code"] == "empty_doc_code"
    assert d.validation_details["warnings"][0]["scope"] == "header"


def test_per_change_flags():
    d = enrich_document(_doc())
    ch = d.all_changes[0]
    assert ch.has_doc_code is True
    assert ch.has_change_text is True
    assert ch.text_length == len("Текст")
