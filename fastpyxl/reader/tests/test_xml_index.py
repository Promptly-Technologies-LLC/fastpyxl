# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.reader.xml_index import build_row_index, build_si_index


def test_build_row_index_skips_gaps_and_self_closing():
    xml = b"""<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
      <row r="2" spans="1:1"><c r="A2"><v>1</v></c></row>
      <row r="4"/>
      <row r="5" spans="1:1"><c r="A5"><v>2</v></c></row>
    </sheetData>
    </worksheet>"""
    index = build_row_index(xml)
    assert set(index) == {2, 4, 5}
    for start, end in index.values():
        assert xml[start:start + 4] == b"<row"
        assert xml[end - 2 : end] in (b"/>", b"w>")


def test_build_row_index_assigns_sequential_when_r_missing():
    xml = b"""<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
      <row spans="1:1"><c><v>1</v></c></row>
      <row spans="1:1"><c><v>2</v></c></row>
    </sheetData>
    </worksheet>"""
    index = build_row_index(xml)
    assert list(index) == [1, 2]


def test_build_si_index_preserves_order_and_duplicates():
    xml = b"""<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="2">
    <si><t>a</t></si><si><t></t></si><si><t>a</t></si>
    </sst>"""
    spans = build_si_index(xml)
    assert len(spans) == 3
    assert xml[spans[0][0] : spans[0][1]] == b"<si><t>a</t></si>"
    assert xml[spans[1][0] : spans[1][1]] == b"<si><t></t></si>"
