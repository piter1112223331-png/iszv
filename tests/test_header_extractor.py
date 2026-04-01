import pytest

openpyxl = pytest.importorskip("openpyxl")
from openpyxl import Workbook

from parser.header_extractor import extract_document_header


def test_developer_extraction():
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1, "Предприятие организация")
    ws.cell(1, 2, 'АО "ИЦ КТ"')

    header, _, _ = extract_document_header(ws)
    assert header.developer == 'АО "ИЦ КТ"'


def test_notice_number_extraction_not_doc_code():
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1, "Извещение")
    ws.cell(1, 2, "ИИ.0000.0001")

    header, _, _ = extract_document_header(ws)
    assert header.notice_number == "ИИ.0000.0001"


def test_code_extraction():
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1, "Код")
    ws.cell(1, 2, "7")

    header, _, _ = extract_document_header(ws)
    assert header.code == "7"


def test_sheet_no_declared_vs_sheet_total_declared():
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1, "Лист")
    ws.cell(1, 2, "1")
    ws.cell(1, 3, "Листов")
    ws.cell(1, 4, "3")

    header, _, _ = extract_document_header(ws)
    assert header.sheet_no_declared == 1
    assert header.sheet_total_declared == 3


def test_approvals_extraction():
    wb = Workbook()
    ws = wb.active
    ws.cell(20, 1, "Составил")
    ws.cell(20, 2, "Петров")
    ws.cell(21, 1, "Проверил")
    ws.cell(21, 2, "Иванов")
    ws.cell(22, 1, "Н. контроль")
    ws.cell(22, 2, "Свердлов")
    ws.cell(23, 1, "Утвердил")
    ws.cell(23, 2, "Алербов")

    _, approvals, _ = extract_document_header(ws)
    assert approvals.author == "Петров"
    assert approvals.reviewer == "Иванов"
    assert approvals.norm_control == "Свердлов"
    assert approvals.approver == "Алербов"
