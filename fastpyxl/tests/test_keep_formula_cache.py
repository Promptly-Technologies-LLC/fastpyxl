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
