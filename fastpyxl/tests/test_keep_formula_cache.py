# Copyright (c) 2010-2024 fastpyxl

"""Tests for keep_formula_cache dual-load of formulas and cached values."""

from __future__ import annotations

from pathlib import Path

import pytest

from fastpyxl import Workbook, load_workbook
from fastpyxl.worksheet.copier import WorksheetCopy


GENUINE_SAMPLE = Path(__file__).resolve().parent / "data/genuine/sample.xlsx"
FORMULA_SHEET = "Sheet3 - Formulas"
FORMULA_COORD = "D2"
FORMULA_TEXT = "='Sheet2 - Numbers'!D5"
CACHED_RESULT = 5


def test_keep_formula_cache_rejects_non_bool():
    with pytest.raises(ValueError, match="keep_formula_cache"):
        load_workbook(GENUINE_SAMPLE, keep_formula_cache="yes")


def test_keep_formula_cache_rejects_combo_with_data_only():
    with pytest.raises(ValueError, match="mutually exclusive|data_only"):
        load_workbook(GENUINE_SAMPLE, data_only=True, keep_formula_cache=True)


def test_default_mode_cached_value_is_none():
    wb = load_workbook(GENUINE_SAMPLE)
    try:
        assert wb.data_only is False
        assert wb.keep_formula_cache is False
        cell = wb[FORMULA_SHEET][FORMULA_COORD]
        assert cell.value == FORMULA_TEXT
        assert cell.data_type == "f"
        assert cell.cached_value is None
    finally:
        wb.close()


def test_data_only_mode_cached_value_is_none():
    wb = load_workbook(GENUINE_SAMPLE, data_only=True)
    try:
        assert wb.data_only is True
        assert wb.keep_formula_cache is False
        cell = wb[FORMULA_SHEET][FORMULA_COORD]
        assert cell.value == CACHED_RESULT
        assert cell.cached_value is None
    finally:
        wb.close()


def test_dual_load_exposes_formula_and_cache():
    wb = load_workbook(GENUINE_SAMPLE, keep_formula_cache=True)
    try:
        assert wb.data_only is False
        assert wb.keep_formula_cache is True
        cell = wb[FORMULA_SHEET][FORMULA_COORD]
        assert cell.value == FORMULA_TEXT
        assert cell.data_type == "f"
        assert cell.cached_value == CACHED_RESULT
        assert isinstance(cell.value, str)
    finally:
        wb.close()


def test_dual_load_read_only():
    wb = load_workbook(GENUINE_SAMPLE, read_only=True, keep_formula_cache=True)
    try:
        assert wb.data_only is False
        assert wb.keep_formula_cache is True
        cell = wb[FORMULA_SHEET][FORMULA_COORD]
        assert cell.value == FORMULA_TEXT
        assert cell.data_type == "f"
        assert cell.cached_value == CACHED_RESULT
    finally:
        wb.close()


def test_values_only_iteration_yields_primary_value():
    wb = load_workbook(GENUINE_SAMPLE, keep_formula_cache=True)
    try:
        ws = wb[FORMULA_SHEET]
        rows = list(ws.iter_rows(min_row=2, max_row=2, min_col=4, max_col=4, values_only=True))
        assert rows == [(FORMULA_TEXT,)]
    finally:
        wb.close()


def test_assigning_value_clears_cached_value():
    wb = load_workbook(GENUINE_SAMPLE, keep_formula_cache=True)
    try:
        cell = wb[FORMULA_SHEET][FORMULA_COORD]
        assert cell.cached_value == CACHED_RESULT
        cell.value = "=1+1"
        assert cell.value == "=1+1"
        assert cell.cached_value is None
    finally:
        wb.close()


def test_cached_value_setter():
    wb = Workbook()
    ws = wb.active
    cell = ws["A1"]
    cell.value = "=SUM(1,1)"
    assert cell.cached_value is None
    cell.cached_value = 2
    assert cell.cached_value == 2
    cell.value = 3
    assert cell.cached_value is None


def test_non_formula_cell_cached_value_is_none():
    wb = load_workbook(GENUINE_SAMPLE, keep_formula_cache=True)
    try:
        cell = wb["Sheet2 - Numbers"]["D5"]
        assert cell.data_type != "f"
        assert cell.cached_value is None
    finally:
        wb.close()


def test_worksheet_copy_preserves_cached_value():
    wb = load_workbook(GENUINE_SAMPLE, keep_formula_cache=True)
    try:
        source = wb[FORMULA_SHEET]
        target = wb.create_sheet("copy")
        WorksheetCopy(source, target).copy_worksheet()
        assert target[FORMULA_COORD].value == FORMULA_TEXT
        assert target[FORMULA_COORD].cached_value == CACHED_RESULT
    finally:
        wb.close()


def test_internal_value_remains_formula():
    wb = load_workbook(GENUINE_SAMPLE, keep_formula_cache=True)
    try:
        cell = wb[FORMULA_SHEET][FORMULA_COORD]
        assert cell.internal_value == FORMULA_TEXT
        assert cell.internal_value != cell.cached_value
    finally:
        wb.close()


def test_cell_has_no_cached_value_slot():
    from fastpyxl.cell.cell import Cell

    assert "_cached_value" not in Cell.__slots__


def test_default_mode_formula_cache_map_stays_empty():
    wb = load_workbook(GENUINE_SAMPLE)
    try:
        for ws in wb.worksheets:
            assert ws._formula_caches == {}
            for cell in ws._cells.values():
                assert not hasattr(cell, "_cached_value")
    finally:
        wb.close()


def test_dual_load_stores_cache_in_worksheet_side_map():
    wb = load_workbook(GENUINE_SAMPLE, keep_formula_cache=True)
    try:
        ws = wb[FORMULA_SHEET]
        cell = ws[FORMULA_COORD]
        assert cell.cached_value == CACHED_RESULT
        assert not hasattr(cell, "_cached_value")
        assert ws._formula_caches[(cell.row, cell.column)] == CACHED_RESULT
    finally:
        wb.close()


def test_deleting_cell_drops_side_map_entry():
    wb = load_workbook(GENUINE_SAMPLE, keep_formula_cache=True)
    try:
        ws = wb[FORMULA_SHEET]
        cell = ws[FORMULA_COORD]
        key = (cell.row, cell.column)
        assert key in ws._formula_caches
        del ws[FORMULA_COORD]
        assert key not in ws._formula_caches
    finally:
        wb.close()


def test_move_range_remaps_formula_cache():
    wb = Workbook()
    ws = wb.active
    cell = ws["A1"]
    cell.value = "=1+1"
    cell.cached_value = 2
    assert ws._formula_caches[(1, 1)] == 2
    ws.move_range("A1", rows=1, cols=1)
    assert (1, 1) not in ws._formula_caches
    assert ws._formula_caches[(2, 2)] == 2
    assert ws["B2"].cached_value == 2


def test_cached_value_zero_is_preserved_in_side_map():
    wb = Workbook()
    ws = wb.active
    cell = ws["A1"]
    cell.value = "=0"
    cell.cached_value = 0
    assert cell.cached_value == 0
    assert ws._formula_caches[(1, 1)] == 0
