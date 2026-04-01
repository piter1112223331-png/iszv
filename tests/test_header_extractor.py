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
    assert dbg["fields"]["reason"]["rejected_reason"] in {"header_like_noise", "no_meaningful_tokens", "empty", "postprocess_rejected"}


def test_repeated_phrase_collapse():
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1, "Причина")
    ws.cell(1, 2, "Устранение ошибок Устранение ошибок Устранение ошибок Устранение ошибок")

    header, dbg = extract_document_header(ws)
    assert header.reason == "Устранение ошибок"
    assert "reason" in dbg["collapsed_repeats_applied"]


def test_dash_noise_normalized_to_null():
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1, "Разослать")
    ws.cell(1, 2, "- - -")

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


def test_code_zone_does_not_absorb_listov_values():
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1, "Код")
    ws.cell(1, 2, "1 3")

    header, _ = extract_document_header(ws)
    assert header.code is None


def test_release_date_zone_does_not_absorb_stock_instruction_text():
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1, "Дата выпуска")
    ws.cell(1, 2, "Задел использовать")

    header, _ = extract_document_header(ws)
    assert header.release_date is None
