import pytest

openpyxl = pytest.importorskip("openpyxl")
from openpyxl import Workbook

from parser.header_extractor import extract_document_header


def test_header_extraction_by_anchors():
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1, "Причина")
    ws.cell(1, 2, "Замена")
    ws.cell(2, 1, "Дата выпуска")
    ws.cell(2, 2, "2026-03-10")
    ws.cell(3, 1, "Разослать")
    ws.cell(3, 2, "ОТК")
    ws.cell(4, 1, "Листов 3")

    header, dbg = extract_document_header(ws)

    assert header.reason == "Замена"
    assert header.release_date == "2026-03-10"
    assert header.distribution == "ОТК"
    assert header.sheet_total_declared == 3
    assert "reason" in dbg["found_fields"]
