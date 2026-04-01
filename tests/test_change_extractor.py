import pytest
openpyxl = pytest.importorskip("openpyxl")
from openpyxl import Workbook

from parser.change_extractor import _extract_meta_signals, extract_changes


def test_meta_signal_requires_more_than_keyword():
    idx, doc, is_meta = _extract_meta_signals("Изм.")
    assert idx is None
    assert doc is None
    assert is_meta is False


def test_extract_changes_skips_headers_in_body():
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1, "Извещение")
    ws.cell(2, 1, "Содержание изменения")
    ws.cell(3, 1, "Изм. 1 ABC-123")
    ws.cell(4, 1, "Содержание изменения")
    ws.cell(5, 1, "Изменить материал детали")
    ws.cell(6, 1, "Изм. 2 DEF-555")
    ws.cell(7, 1, "Лист")
    ws.cell(8, 1, "Добавить покрытие")

    changes, _ = extract_changes(ws, sheet_index=1, start_global_seq=1)

    assert len(changes) == 2
    assert changes[0].change_index == "1"
    assert changes[0].doc_code == "ABC-123"
    assert changes[0].change_text == "Изменить материал детали"
    assert changes[1].change_text == "Добавить покрытие"
