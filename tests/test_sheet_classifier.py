import pytest
openpyxl = pytest.importorskip("openpyxl")
from openpyxl import Workbook

from parser.sheet_classifier import classify_sheet


def _sheet_with(*lines: str):
    wb = Workbook()
    ws = wb.active
    for idx, text in enumerate(lines, start=1):
        ws.cell(row=idx, column=1, value=text)
    return ws


def test_classifies_full_sheet():
    ws = _sheet_with(
        "Извещение № 123",
        "Изм. 1",
        "Содержание изменения",
        "Причина",
        "Дата выпуска",
    )
    result = classify_sheet(ws)
    assert result.is_candidate is True
    assert result.kind == "full"


def test_classifies_continuation_sheet():
    ws = _sheet_with(
        "Извещение № 123",
        "Изм. 2",
        "Содержание изменения",
        "Лист 2",
    )
    result = classify_sheet(ws)
    assert result.is_candidate is True
    assert result.kind == "continuation"
