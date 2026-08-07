# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

import datetime
from pathlib import Path

import pytest

from .conftest import GENUINE_SAMPLE_FIXTURE
from .helpers import assert_workbooks_match, save_and_reload_cross


pytestmark = pytest.mark.openpyxl_compat


def test_empty_workbook_defaults_match(fastpyxl, openpyxl):
    fast_wb = fastpyxl.Workbook()
    openpyxl_wb = openpyxl.Workbook()
    assert_workbooks_match(fast_wb, openpyxl_wb)


def test_append_rows_and_cell_values_match(fastpyxl, openpyxl):
    fast_wb = fastpyxl.Workbook()
    openpyxl_wb = openpyxl.Workbook()
    for wb in (fast_wb, openpyxl_wb):
        ws = wb.active
        ws.title = "Data"
        ws["A1"] = 42
        ws["B1"] = "hello"
        ws.append([1, 2, 3])
        ws["A3"] = datetime.datetime(2024, 1, 15, 10, 30)
    assert_workbooks_match(fast_wb, openpyxl_wb)


def test_formula_assignment_matches(fastpyxl, openpyxl):
    fast_wb = fastpyxl.Workbook()
    openpyxl_wb = openpyxl.Workbook()
    for wb in (fast_wb, openpyxl_wb):
        ws = wb.active
        ws["A1"] = 10
        ws["A2"] = 20
        ws["A3"] = "=A1+A2"
        ws["A4"] = "=SUM(A1:A3)"
    assert_workbooks_match(fast_wb, openpyxl_wb)


def test_sheet_management_matches(fastpyxl, openpyxl):
    fast_wb = fastpyxl.Workbook()
    openpyxl_wb = openpyxl.Workbook()
    for wb in (fast_wb, openpyxl_wb):
        wb.active.title = "First"
        second = wb.create_sheet("Second", 1)
        second["A1"] = "moved"
        wb.create_sheet("Third")
        wb.move_sheet("Third", offset=-1)
    assert_workbooks_match(fast_wb, openpyxl_wb)
    assert fast_wb.sheetnames == openpyxl_wb.sheetnames


def test_fastpyxl_save_openpyxl_read(tmp_path: Path, fastpyxl, openpyxl):
    wb = fastpyxl.Workbook()
    ws = wb.active
    ws.title = "Export"
    ws.append(["name", "count"])
    ws.append(["widgets", 12])
    ws["C2"] = "=B2*2"

    path = tmp_path / "fastpyxl_written.xlsx"
    reloaded = save_and_reload_cross(wb, openpyxl.load_workbook, path)
    reference = fastpyxl.load_workbook(path)
    assert_workbooks_match(reference, reloaded)


def test_openpyxl_save_fastpyxl_read(tmp_path: Path, fastpyxl, openpyxl):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Import"
    ws.append(["name", "count"])
    ws.append(["gadgets", 7])
    ws["C2"] = "=B2+1"

    path = tmp_path / "openpyxl_written.xlsx"
    reloaded = save_and_reload_cross(wb, fastpyxl.load_workbook, path)
    reference = openpyxl.load_workbook(path)
    assert_workbooks_match(reloaded, reference)


@pytest.mark.parametrize("extension", ["xlsx", "xlsm", "xltx", "xltm"])
@pytest.mark.parametrize("writer", ["fastpyxl", "openpyxl"])
def test_save_extension_round_trip(
    extension: str,
    writer: str,
    tmp_path: Path,
    fastpyxl,
    openpyxl,
):
    writer_lib = fastpyxl if writer == "fastpyxl" else openpyxl
    wb = writer_lib.Workbook()
    wb.active["A1"] = f"saved-as-{extension}-by-{writer}"
    path = tmp_path / f"workbook-{writer}.{extension}"

    wb.save(path)
    fast_reloaded = fastpyxl.load_workbook(path)
    openpyxl_reloaded = openpyxl.load_workbook(path)
    assert_workbooks_match(fast_reloaded, openpyxl_reloaded)


def test_read_only_iteration_values_match(fastpyxl, openpyxl):
    fixture = GENUINE_SAMPLE_FIXTURE
    fast_wb = fastpyxl.load_workbook(fixture, read_only=True)
    openpyxl_wb = openpyxl.load_workbook(fixture, read_only=True)
    try:
        assert_workbooks_match(fast_wb, openpyxl_wb)
    finally:
        fast_wb.close()
        openpyxl_wb.close()


def _populate_write_only_workbook(workbook):
    """Append mixed-type rows to a write-only workbook and return it."""
    data = workbook.create_sheet("Data")
    data.append(["name", "count", "when", "day", "formula", "flag"])
    data.append(
        [
            "widgets",
            12,
            datetime.datetime(2024, 1, 15, 10, 30),
            datetime.date(2024, 2, 1),
            "=B2*2",
            True,
        ]
    )
    data.append(
        [
            "gadgets",
            7.5,
            None,
            datetime.date(2024, 3, 1),
            "=SUM(B2:B3)",
            False,
        ]
    )
    more = workbook.create_sheet("More")
    more.append([1, "a", 3.14])
    return workbook


def test_write_only_round_trip_matches(tmp_path: Path, fastpyxl, openpyxl):
    fast_path = tmp_path / "fastpyxl_write_only.xlsx"
    openpyxl_path = tmp_path / "openpyxl_write_only.xlsx"

    _populate_write_only_workbook(fastpyxl.Workbook(write_only=True)).save(fast_path)
    _populate_write_only_workbook(openpyxl.Workbook(write_only=True)).save(
        openpyxl_path
    )

    fast_reloaded = fastpyxl.load_workbook(fast_path)
    openpyxl_reloaded = openpyxl.load_workbook(openpyxl_path)
    try:
        assert_workbooks_match(fast_reloaded, openpyxl_reloaded)
    finally:
        fast_reloaded.close()
        openpyxl_reloaded.close()


def test_write_only_fastpyxl_save_openpyxl_read(tmp_path: Path, fastpyxl, openpyxl):
    path = tmp_path / "fastpyxl_write_only_cross.xlsx"
    _populate_write_only_workbook(fastpyxl.Workbook(write_only=True)).save(path)

    reference = fastpyxl.load_workbook(path)
    reloaded = openpyxl.load_workbook(path)
    try:
        assert_workbooks_match(reference, reloaded)
    finally:
        reference.close()
        reloaded.close()


def test_write_only_openpyxl_save_fastpyxl_read(tmp_path: Path, fastpyxl, openpyxl):
    path = tmp_path / "openpyxl_write_only_cross.xlsx"
    _populate_write_only_workbook(openpyxl.Workbook(write_only=True)).save(path)

    reloaded = fastpyxl.load_workbook(path)
    reference = openpyxl.load_workbook(path)
    try:
        assert_workbooks_match(reloaded, reference)
    finally:
        reloaded.close()
        reference.close()
