from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Literal


SheetKind = Literal["full", "continuation", "unknown"]


@dataclass
class ZoneRef:
    meta_row_start: int | None = None
    meta_row_end: int | None = None
    body_row_start: int | None = None
    body_row_end: int | None = None


@dataclass
class ChangeBlock:
    sheet_index: int
    change_seq_global: int
    change_seq_on_sheet: int
    change_index: str | None
    doc_code: str | None
    change_text: str | None
    raw_meta_text: str | None
    zone_ref: ZoneRef = field(default_factory=ZoneRef)


@dataclass
class SheetLocalHeader:
    notice_number: str | None = None


@dataclass
class SheetResult:
    sheet_index: int
    sheet_name: str
    sheet_kind: SheetKind
    sheet_no_detected: int | None
    sheet_local_header: SheetLocalHeader = field(default_factory=SheetLocalHeader)
    changes: list[ChangeBlock] = field(default_factory=list)


@dataclass
class DocumentHeader:
    sender: str | None = None
    reason: str | None = None
    code: str | None = None
    sheet_total_declared: int | None = None
    release_center: str | None = None
    release_date: str | None = None
    stock_instruction: str | None = None
    implementation_instruction: str | None = None
    applicability: str | None = None
    distribution: str | None = None


@dataclass
class ValidationResult:
    template_detected: bool = True
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


@dataclass
class ParsedDocument:
    document_type: str
    template_version: str
    source_file: str
    notice_number: str | None
    sheet_count_detected: int
    document_header: DocumentHeader
    sheets: list[SheetResult]
    all_changes: list[ChangeBlock]
    validation: ValidationResult

    def to_dict(self) -> dict:
        return asdict(self)
