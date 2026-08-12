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
    ix = load_workbook(_workbook_with_sheet_xml(sheet_xml), access="indexed")
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


def _workbook_with_sheet_xml(sheet_xml: bytes) -> BytesIO:
    from zipfile import ZIP_DEFLATED, ZipFile

    zf_data = _save_bytes(Workbook()).getvalue()
    out = BytesIO()
    with ZipFile(BytesIO(zf_data), "r") as zin, ZipFile(out, "w", ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename.endswith("sheet1.xml"):
                data = sheet_xml
            zout.writestr(info.filename, data)
    out.seek(0)
    return out


def test_indexed_rich_text_getitem_matches_iter_rows():
    from fastpyxl.cell.rich_text import CellRichText

    sheet_xml = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b'<dimension ref="A1"/><sheetData><row r="1" spans="1:1">'
        b'<c r="A1" t="inlineStr"><is><r><rPr><b/></rPr><t>Hello</t></r>'
        b'<r><t xml:space="preserve"> world</t></r></is></c>'
        b'</row></sheetData></worksheet>'
    )
    src = _workbook_with_sheet_xml(sheet_xml)
    wb = load_workbook(src, access="indexed", rich_text=True)
    try:
        via_getitem = wb.active["A1"].value
        via_iter = next(wb.active.iter_rows(min_row=1, max_row=1, max_col=1))[0].value
        assert isinstance(via_getitem, CellRichText)
        assert isinstance(via_iter, CellRichText)
        assert str(via_getitem) == str(via_iter) == "Hello world"
    finally:
        wb.close()


def test_indexed_shared_formula_accepts_single_quoted_t_attr():
    sheet_xml = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b'<dimension ref="A1:B2"/><sheetData>'
        b'<row r="1" spans="1:2">'
        b'<c r="A1"><v>1</v></c>'
        b"<c r=\"B1\" t=\"str\"><f t='shared' ref=\"B1:B2\" si=\"0\">A1*10</f></c>"
        b'</row>'
        b'<row r="2" spans="1:2">'
        b'<c r="A2"><v>2</v></c>'
        b"<c r=\"B2\"><f t='shared' si=\"0\"/></c>"
        b'</row></sheetData></worksheet>'
    )
    wb = load_workbook(_workbook_with_sheet_xml(sheet_xml), access="indexed")
    try:
        assert wb.active["B1"].value == "=A1*10"
        assert wb.active["B2"].value == "=A2*10"
    finally:
        wb.close()


def test_indexed_keep_formula_cache():
    wb = load_workbook(GENUINE_SAMPLE, access="indexed", keep_formula_cache=True)
    try:
        cell = wb["Sheet3 - Formulas"]["D2"]
        assert cell.value == "='Sheet2 - Numbers'!D5"
        assert cell.cached_value == 5
    finally:
        wb.close()


def test_decompressed_part_spools_large_members_without_second_bytes_copy():
    from fastpyxl.reader.part_cache import DecompressedPart
    from fastpyxl.reader.xml_index import build_row_index

    # Force spooling with a tiny threshold; payload includes a real <row>.
    payload = (
        b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        b"<sheetData>" + (b"<row r=\"1\"/>" + b"x" * 200) + b"</sheetData></worksheet>"
    )
    part = DecompressedPart(payload, spool_max=64)
    try:
        assert part.spooled is True
        with part.bytes_view() as view:
            # File-backed view must support indexing without requiring .read() into bytes.
            index = build_row_index(view)
            assert 1 in index
        assert part.read_at(*index[1]).startswith(b"<row")
    finally:
        part.close()


def test_decompressed_part_from_archive_streams_when_large(tmp_path):
    from zipfile import ZIP_DEFLATED, ZipFile

    from fastpyxl.reader.part_cache import DecompressedPart

    raw = b"a" * 10_000
    zpath = tmp_path / "part.zip"
    with ZipFile(zpath, "w", ZIP_DEFLATED) as zf:
        zf.writestr("xl/worksheets/sheet1.xml", raw)
    with ZipFile(zpath, "r") as zf:
        part = DecompressedPart.from_archive(zf, "xl/worksheets/sheet1.xml", spool_max=100)
        try:
            assert part.spooled is True
            assert part.size == len(raw)
            assert part.read_at(0, 4) == b"aaaa"
        finally:
            part.close()
