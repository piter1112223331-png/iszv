from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from parser.models import DocumentHeader, ParsedDocument, ValidationResult
from parser.sheet_parser import parse_sheet
from parser.workbook_loader import WorkbookLoadError, load_xlsx


def parse_notice(path: str, debug: bool = False) -> ParsedDocument:
    workbook = load_xlsx(path)

    parsed_sheets = []
    all_changes = []
    next_seq = 1
    candidate_sheets = 0

    for idx, sheet in enumerate(workbook.worksheets, start=1):
        parsed, next_seq, is_candidate, extraction_debug = parse_sheet(
            sheet,
            sheet_index=idx,
            start_global_seq=next_seq,
        )
        if is_candidate:
            candidate_sheets += 1
            parsed_sheets.append(parsed)
            all_changes.extend(parsed.changes)

        if debug:
            print(
                (
                    f"[debug] sheet={sheet.title!r} "
                    f"kind={parsed.sheet_kind} "
                    f"notice_number={parsed.sheet_local_header.notice_number!r} "
                    f"sheet_no={parsed.sheet_no_detected!r} "
                    f"changes={len(parsed.changes)} "
                    f"candidate={is_candidate} "
                    f"table_header_row_start={extraction_debug.get('table_header_row_start')} "
                    f"potential_meta_rows={extraction_debug.get('potential_meta_rows')} "
                    f"rejected_meta_rows={extraction_debug.get('rejected_meta_rows')} "
                    f"reject_reasons={extraction_debug.get('reject_reasons')}"
                ),
                file=sys.stderr,
            )

    full_sheets = [s for s in parsed_sheets if s.sheet_kind == "full"]
    notice_number = None
    for collection in (full_sheets, parsed_sheets):
        for sheet in collection:
            if sheet.sheet_local_header.notice_number:
                notice_number = sheet.sheet_local_header.notice_number
                break
        if notice_number:
            break

    validation = ValidationResult(template_detected=candidate_sheets > 0)
    if candidate_sheets == 0:
        validation.errors.append("No notice-like sheets were detected using MVP-1 markers")

    if debug:
        print(f"[debug] candidate_sheets={candidate_sheets} all_changes={len(all_changes)}", file=sys.stderr)

    return ParsedDocument(
        document_type="change_notice",
        template_version="notice_multi_sheet_v1",
        source_file=Path(path).name,
        notice_number=notice_number,
        sheet_count_detected=len(parsed_sheets),
        document_header=DocumentHeader(),
        sheets=parsed_sheets,
        all_changes=all_changes,
        validation=validation,
    )


def build_cli() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="MVP-1 parser for XLSX change notice")
    parser.add_argument("xlsx_path", help="Path to source XLSX file")
    parser.add_argument("-o", "--output", help="Output JSON path; if omitted print to stdout")
    parser.add_argument("--indent", type=int, default=2, help="JSON indent")
    parser.add_argument("--pretty", action="store_true", help="Pretty JSON output with indent=2")
    parser.add_argument("--debug", action="store_true", help="Print diagnostic info to stderr")
    return parser


def main() -> int:
    parser = build_cli()
    args = parser.parse_args()

    try:
        result = parse_notice(args.xlsx_path, debug=args.debug)
    except WorkbookLoadError as exc:
        parser.error(str(exc))
        return 2

    indent = 2 if args.pretty else args.indent
    payload = json.dumps(result.to_dict(), ensure_ascii=False, indent=indent)
    if args.output:
        Path(args.output).write_text(payload + "\n", encoding="utf-8")
    else:
        print(payload)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
