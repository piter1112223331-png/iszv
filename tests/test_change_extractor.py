import pytest

openpyxl = pytest.importorskip("openpyxl")
from openpyxl import Workbook

from parser.change_extractor import extract_changes


def test_service_header_rows_are_not_meta_rows():
    wb = Workbook()
    ws = wb.active
    ws.cell(1, 1, "Извещение № ЕСРТ.0016.716.04121")
    ws.cell(2, 1, "Указание о заделе")
    ws.cell(3, 1, "Изм.")
    ws.cell(3, 2, "Содержание изменения")
    ws.cell(4, 1, "Срок")
    ws.cell(4, 2, "Прочее")

    changes, _, dbg = extract_changes(ws, sheet_index=1, start_global_seq=1)

    assert len(changes) == 0
    assert dbg["table_header_row_start"] == 3


def test_meta_row_below_table_with_index_and_doc_code_is_detected():
    wb = Workbook()
    ws = wb.active
    ws.cell(3, 1, "Изм.")
    ws.cell(3, 3, "Содержание изменения")
    ws.cell(5, 1, "1")
    ws.cell(5, 3, "ЕСРТ.0016.716.04121")
    ws.cell(6, 3, "Заменить покрытие")

    changes, _, dbg = extract_changes(ws, sheet_index=1, start_global_seq=1)

    assert dbg["potential_meta_rows"] >= 1
    assert len(changes) == 1
    assert changes[0].change_index == "1"
    assert changes[0].doc_code == "ЕСРТ.0016.716.04121"


def test_merged_duplicates_are_collapsed_in_change_text():
    wb = Workbook()
    ws = wb.active
    ws.cell(2, 1, "Изм.")
    ws.cell(2, 3, "Содержание изменения")
    ws.cell(4, 1, "2")
    ws.cell(4, 3, "ЕСРТ.0016.0000.0001")
    ws.cell(5, 3, "Повторяющийся текст")
    ws.merge_cells("C5:E5")
    ws.cell(6, 3, "Повторяющийся текст")

    changes, _, _ = extract_changes(ws, sheet_index=1, start_global_seq=1)

    assert len(changes) == 1
    assert changes[0].change_text == "Повторяющийся текст"
