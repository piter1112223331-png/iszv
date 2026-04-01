from __future__ import annotations

from dataclasses import asdict

from parser.models import ParsedDocument


def _tech_flags(doc: ParsedDocument) -> None:
    for change in doc.all_changes:
        change.has_doc_code = bool(change.doc_code)
        change.has_change_text = bool(change.change_text)
        change.text_length = len(change.change_text or "")


def compute_status(doc: ParsedDocument) -> str:
    if doc.validation.errors:
        return "error"
    if doc.validation.warnings:
        return "warning"
    return "ok"


def build_summary(doc: ParsedDocument, header_debug: dict[str, object] | None = None) -> dict:
    header_debug = header_debug or {}
    full = sum(1 for s in doc.sheets if s.sheet_kind == "full")
    cont = sum(1 for s in doc.sheets if s.sheet_kind == "continuation")
    return {
        "candidate_sheets_count": len(doc.sheets),
        "full_sheets_count": full,
        "continuation_sheets_count": cont,
        "change_blocks_count": len(doc.all_changes),
        "warnings_count": len(doc.validation.warnings),
        "errors_count": len(doc.validation.errors),
        "header_fields_found_count": len(header_debug.get("found_fields", [])),
        "header_fields_missing_count": len(header_debug.get("missing_fields", [])),
    }


def build_llm_payload(doc: ParsedDocument) -> dict:
    header_compact = {k: v for k, v in asdict(doc.document_header).items() if v is not None}

    compact_changes = []
    for ch in doc.all_changes:
        item = {
            "sheet_no": ch.sheet_index,
            "change_index": ch.change_index,
            "doc_code": ch.doc_code,
            "change_text": ch.change_text,
            "change_seq_global": ch.change_seq_global,
            "has_doc_code": ch.has_doc_code,
            "has_change_text": ch.has_change_text,
            "text_length": ch.text_length,
        }
        item = {k: v for k, v in item.items() if v is not None}
        compact_changes.append(item)

    approvals = {k: v for k, v in asdict(doc.approvals).items() if v is not None}

    return {
        "source_file": doc.source_file,
        "notice_number": doc.notice_number,
        "sheet_count_detected": doc.sheet_count_detected,
        "document_header_compact": header_compact,
        "changes": compact_changes,
        "warnings": doc.validation.warnings,
        "errors": doc.validation.errors,
        **({"approvals": approvals} if approvals else {}),
        "summary": {
            "changes_count": len(compact_changes),
            "warnings_count": len(doc.validation.warnings),
            "errors_count": len(doc.validation.errors),
        },
    }


def build_flowis_payload(doc: ParsedDocument) -> dict:
    approvals = {k: v for k, v in asdict(doc.approvals).items() if v is not None}

    return {
        "source_file": doc.source_file,
        "status": doc.status,
        "notice_number": doc.notice_number,
        "developer": doc.document_header.developer,
        "code": doc.document_header.code,
        "sheet_no_declared": doc.document_header.sheet_no_declared,
        "sheet_count_detected": doc.sheet_count_detected,
        "sheet_total_declared": doc.document_header.sheet_total_declared,
        "reason": doc.document_header.reason,
        "stock_instruction": doc.document_header.stock_instruction,
        "implementation_instruction": doc.document_header.implementation_instruction,
        "applicability": doc.document_header.applicability,
        "distribution": doc.document_header.distribution,
        "approvals_author": doc.approvals.author,
        "approvals_reviewer": doc.approvals.reviewer,
        "approvals_norm_control": doc.approvals.norm_control,
        "approvals_approver": doc.approvals.approver,
        "changes_count": len(doc.all_changes),
        "warnings": doc.validation.warnings,
        "errors": doc.validation.errors,
    }


def _to_detail(item: str, kind: str) -> dict:
    code = item
    ref = None
    scope = "document"

    if ":" in item:
        code, tail = item.split(":", 1)
        ref = tail

    if code.startswith("header_field_missing"):
        scope = "header"
    elif code.startswith("empty_doc_code") or code.startswith("empty_change_index") or code.startswith("empty_change_text"):
        scope = "change"
    elif code.startswith("sheet_count"):
        scope = "sheet"

    return {"code": code, "scope": scope, "ref": ref}


def build_validation_details(doc: ParsedDocument) -> dict:
    return {
        "errors": [_to_detail(e, "error") for e in doc.validation.errors],
        "warnings": [_to_detail(w, "warning") for w in doc.validation.warnings],
    }


def enrich_document(doc: ParsedDocument, header_debug: dict[str, object] | None = None) -> ParsedDocument:
    _tech_flags(doc)
    doc.status = compute_status(doc)
    doc.summary = build_summary(doc, header_debug=header_debug)
    doc.llm_payload = build_llm_payload(doc)
    doc.flowis_payload = build_flowis_payload(doc)
    doc.validation_details = build_validation_details(doc)
    return doc
