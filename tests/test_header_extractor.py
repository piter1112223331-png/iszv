import pytest

openpyxl = pytest.importorskip("openpyxl")
from openpyxl import Workbook

from parser.header_extractor import extract_document_header


def test_developer_extraction_from_org_cell():
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1, "Организация")
    ws.cell(1, 2, 'АО "ИЦ КТ"')

    header, _, _ = extract_document_header(ws)
    assert header.developer == 'АО "ИЦ КТ"'


def test_code_extraction_from_dedicated_cell_not_sheet_counter():
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1, "Код")
    ws.cell(1, 2, "7")
    ws.cell(1, 3, "Лист")
    ws.cell(1, 4, "1")
    ws.cell(1, 5, "Листов")
    ws.cell(1, 6, "1")

    header, _, _ = extract_document_header(ws)
    assert header.code == "7"


def test_sheet_no_declared_extraction_from_list_cell():
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1, "Лист")
    ws.cell(1, 2, "1")
    ws.cell(1, 3, "Листов")
    ws.cell(1, 4, "3")

    header, _, _ = extract_document_header(ws)
    assert header.sheet_no_declared == 1
    assert header.sheet_total_declared == 3


def test_dash_preserving_fields_return_dash():
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1, "Указания о внедрении")
    ws.cell(1, 2, "-")
    ws.cell(2, 1, "Применяемость")
    ws.cell(2, 2, "-")
    ws.cell(3, 1, "Разослать")
    ws.cell(3, 2, "-")

    header, _, _ = extract_document_header(ws)
    assert header.implementation_instruction == "-"
    assert header.applicability == "-"
    assert header.distribution == "-"


def test_approvals_values_extracted_to_right_of_labels():
    wb = Workbook()
    ws = wb.active
    ws.cell(20, 1, "Составил")
    ws.cell(20, 2, "Составил")
    ws.cell(20, 3, "Петров")
    ws.cell(21, 1, "Проверил")
    ws.cell(21, 2, "Проверил")
    ws.cell(21, 3, "Иванов")
    ws.cell(22, 1, "Н. контроль")
    ws.cell(22, 2, "Н. контроль")
    ws.cell(22, 3, "Свердлов")
    ws.cell(23, 1, "Утвердил")
    ws.cell(23, 2, "Утвердил")
    ws.cell(23, 3, "Алербов")

    _, approvals, debug = extract_document_header(ws)
    assert approvals.author == "Петров"
    assert approvals.reviewer == "Иванов"
    assert approvals.norm_control == "Свердлов"
    assert approvals.approver == "Алербов"
    assert debug["approvals_rejected_reasons"] == {}
