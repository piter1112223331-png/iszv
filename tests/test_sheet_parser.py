import pytest

openpyxl = pytest.importorskip("openpyxl")
from openpyxl import Workbook

from parser.sheet_parser import detect_notice_number


def test_notice_number_is_not_label_word():
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1, "Извещение")
    ws.cell(1, 2, "об изменении")
    ws.cell(2, 1, "Извещение № ЕСРТ.0016.716.04121")

    assert detect_notice_number(ws) == "ЕСРТ.0016.716.04121"
