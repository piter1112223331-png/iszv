import pytest

openpyxl = pytest.importorskip("openpyxl")
from openpyxl import Workbook

from parser.header_extractor import extract_document_header


def test_label_text_should_not_become_value():
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1, "Причина")
    ws.cell(1, 2, "Причина")

    header, dbg = extract_document_header(ws)

    assert header.reason is None
    assert dbg["fields"]["reason"]["rejected_reason"] in {"header_like_noise", "no_meaningful_tokens", "empty"}


def test_repeated_header_labels_are_rejected():
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1, "Разослать")
    ws.cell(1, 2, "Разослать Разослать")

    header, _ = extract_document_header(ws)
    assert header.distribution is None


def test_value_zone_right_of_anchor_extracted_without_anchor():
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1, "Причина")
    ws.cell(1, 2, "Замена узла")

    header, _ = extract_document_header(ws)
    assert header.reason == "Замена узла"


def test_sheet_total_declared_short_numeric_near_listov():
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1, "Листов")
    ws.cell(1, 2, "3")

    header, dbg = extract_document_header(ws)
    assert header.sheet_total_declared == 3
    assert dbg["fields"]["sheet_total_declared"]["final_value"] == 3


def test_null_returned_instead_of_header_noise():
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1, "Дата выпуска")
    ws.cell(1, 2, "Дата выпуска Срок изм. Срок действия ПИ")

    header, _ = extract_document_header(ws)
    assert header.release_date is None
