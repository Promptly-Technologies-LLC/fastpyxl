# Copyright (c) 2010-2024 fastpyxl

"""Tests for access='indexed' on-demand workbook reads (issue #75)."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pytest

from fastpyxl import Workbook, load_workbook
from fastpyxl.cell.read_only import EMPTY_CELL, ReadOnlyCell
from fastpyxl.worksheet._indexed import IndexedWorksheet
from fastpyxl.worksheet._read_only import ReadOnlyWorksheet


GENUINE_SAMPLE = Path(__file__).resolve().parent / "data/genuine/sample.xlsx"


def _save_bytes(wb: Workbook) -> BytesIO:
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def _sparse_workbook_bytes(rows: int = 200, hit_rows=(1, 50, 100, 150, 200)) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sparse"
    for r in hit_rows:
        ws.cell(r, 1, value=f"r{r}")
        ws.cell(r, 3, value=r * 10)
    ws.cell(1, 2, value="shared")
    return _save_bytes(wb)


def test_access_rejects_unknown_value():
    with pytest.raises(ValueError, match="access"):
        load_workbook(GENUINE_SAMPLE, access="lazy")


def test_indexed_implies_read_only():
    wb = load_workbook(GENUINE_SAMPLE, access="indexed")
    try:
        assert wb.read_only is True
        assert isinstance(wb.active, IndexedWorksheet)
        assert isinstance(wb.active, ReadOnlyWorksheet)
        with pytest.raises(TypeError):
            wb.save(BytesIO())
    finally:
        wb.close()


def test_indexed_with_explicit_read_only_ok():
    wb = load_workbook(GENUINE_SAMPLE, read_only=True, access="indexed")
    try:
        assert wb.read_only is True
        assert isinstance(wb.active, IndexedWorksheet)
    finally:
        wb.close()


def test_indexed_cell_lookup_matches_read_only():
    src = _sparse_workbook_bytes()
    data = src.getvalue()

    ro = load_workbook(BytesIO(data), read_only=True)
    ix = load_workbook(BytesIO(data), access="indexed")
    try:
        for coord in ("A1", "C50", "A100", "C200", "B1", "Z9"):
            a = ro.active[coord]
            b = ix.active[coord]
            if a is EMPTY_CELL:
                assert b is EMPTY_CELL
            else:
                assert isinstance(b, ReadOnlyCell)
                assert a.value == b.value
                assert a.data_type == b.data_type
                assert a.coordinate == b.coordinate
    finally:
        ro.close()
        ix.close()


def test_indexed_missing_cell_is_empty_singleton():
    wb = load_workbook(_sparse_workbook_bytes(), access="indexed")
    try:
        assert wb.active["B99"] is EMPTY_CELL
        assert wb.active.cell(row=99, column=2) is EMPTY_CELL
    finally:
        wb.close()


def test_indexed_does_not_densify_on_miss():
    wb = load_workbook(_sparse_workbook_bytes(), access="indexed")
    try:
        ws = wb.active
        before = ws.calculate_dimension()
        assert ws["ZZ999"] is EMPTY_CELL
        assert ws.calculate_dimension() == before
    finally:
        wb.close()


def test_indexed_shared_strings_resolved_on_demand():
    from fastpyxl.reader.indexed_strings import IndexedSharedStrings

    # genuine/sample.xlsx stores values in xl/sharedStrings.xml (not inlineStr).
    ix = load_workbook(GENUINE_SAMPLE, access="indexed")
    try:
        assert isinstance(ix._indexed_strings, IndexedSharedStrings)
        assert len(ix._indexed_strings) >= 1
        cell = ix["Sheet1 - Text"]["A1"]
        assert cell.value == ix._indexed_strings[0]
        # A second lookup should hit the LRU without changing semantics.
        assert ix["Sheet1 - Text"]["A1"].value == cell.value
    finally:
        ix.close()


def test_indexed_iter_rows_matches_read_only():
    data = _sparse_workbook_bytes().getvalue()
    ro = load_workbook(BytesIO(data), read_only=True)
    ix = load_workbook(BytesIO(data), access="indexed")
    try:
        ro_vals = list(ro.active.iter_rows(min_row=1, max_row=200, max_col=3, values_only=True))
        ix_vals = list(ix.active.iter_rows(min_row=1, max_row=200, max_col=3, values_only=True))
        assert ix_vals == ro_vals
    finally:
        ro.close()
        ix.close()


def test_indexed_close_releases_resources():
    wb = load_workbook(_sparse_workbook_bytes(), access="indexed")
    ws = wb.active
    assert ws["A1"].value == "r1"
    wb.close()
    with pytest.raises(Exception):
        _ = ws["A1"].value


def test_indexed_shared_formula_translation():
    """Shared formula dependents must resolve via masters captured at index time."""
    from zipfile import ZIP_DEFLATED, ZipFile

    wb = Workbook()
    ws = wb.active
    ws["A1"] = 1
    ws["A2"] = 2
    ws["B1"] = "=A1*10"
    zf_data = _save_bytes(wb).getvalue()

    sheet_xml = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b'<dimension ref="A1:B2"/><sheetData>'
        b'<row r="1" spans="1:2">'
        b'<c r="A1"><v>1</v></c>'
        b'<c r="B1" t="str"><f t="shared" ref="B1:B2" si="0">A1*10</f></c>'
        b'</row>'
        b'<row r="2" spans="1:2">'
        b'<c r="A2"><v>2</v></c>'
        b'<c r="B2"><f t="shared" si="0"/></c>'
        b'</row>'
        b'</sheetData></worksheet>'
    )
    out = BytesIO()
    with ZipFile(BytesIO(zf_data), "r") as zin, ZipFile(out, "w", ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename.endswith("sheet1.xml"):
                data = sheet_xml
            zout.writestr(info.filename, data)

    out.seek(0)
    ix = load_workbook(out, access="indexed")
    try:
        assert ix.active["B1"].value == "=A1*10"
        assert ix.active["B2"].value == "=A2*10"
    finally:
        ix.close()


def test_default_and_read_only_unchanged():
    wb = load_workbook(GENUINE_SAMPLE)
    try:
        assert wb.read_only is False
        assert not isinstance(wb.active, IndexedWorksheet)
    finally:
        wb.close()

    wb = load_workbook(GENUINE_SAMPLE, read_only=True)
    try:
        assert wb.read_only is True
        assert isinstance(wb.active, ReadOnlyWorksheet)
        assert not isinstance(wb.active, IndexedWorksheet)
    finally:
        wb.close()
