# Copyright (c) 2010-2024 fastpyxl

import datetime
from io import BytesIO

# package imports
from fastpyxl.workbook.defined_name import DefinedName
from fastpyxl.utils.exceptions import ReadOnlyWorkbookException
from fastpyxl.worksheet.worksheet import Worksheet

from fastpyxl.xml.constants import (
    XLSM,
    XLSX,
    XLTM,
    XLTX
)

import pytest

@pytest.fixture
def Workbook():
    """Workbook Class"""
    from fastpyxl import Workbook
    return Workbook


@pytest.fixture
def Table():
    """Table Class"""
    from fastpyxl.worksheet.table import Table
    return Table


class TestWorkbook:

    @pytest.mark.parametrize("has_vba, as_template, content_type",
                             [
                                 (None, False, XLSX),
                                 (None, True, XLTX),
                                 (True, False, XLSM),
                                 (True, True, XLTM)
                             ]
                             )
    def test_template(self, has_vba, as_template, content_type, Workbook):
        wb = Workbook()
        wb.vba_archive = has_vba
        wb.template = as_template
        assert wb.mime_type == content_type


    def test_named_styles(self, Workbook):
        wb = Workbook()
        assert wb.named_styles == ['Normal']


    def test_immutable_builtins(self, Workbook):
        wb1 = Workbook()
        wb2 = Workbook()
        normal = wb1._named_styles['Normal']
        normal.font.color = "FF0000"
        assert wb2._named_styles['Normal'].font.color.index == 1


    def test_duplicate_table_name(self, Workbook, Table):
        wb = Workbook()
        ws = wb.create_sheet()
        ws.add_table(Table(displayName="Table1", ref="A1:D10"))
        assert wb._duplicate_name("Table1") is True
        assert wb._duplicate_name("TABLE1") is True


    def test_duplicate_defined_name(self, Workbook):
        wb1 = Workbook()
        wb1.defined_names["dfn1"] = DefinedName("dfn1")
        assert wb1._duplicate_name("dfn1") is True
        assert wb1._duplicate_name("DFN1") is True

def test_get_active_sheet(Workbook):
    wb = Workbook()
    assert wb.active == wb.worksheets[0]


def test_set_active_by_sheet(Workbook):
    wb = Workbook()
    names = ['Sheet', 'Sheet1', 'Sheet2',]
    for n in names:
        wb.create_sheet(n)

    for n in names:
        sheet = wb[n]
        wb.active = sheet
        assert wb.active == wb[n]


def test_set_active_by_index(Workbook):
    wb = Workbook()
    names = ['Sheet', 'Sheet1', 'Sheet2',]
    for n in names:
        wb.create_sheet(n)

    for idx, name in enumerate(names):
        wb.active = idx
        assert wb.active == wb.worksheets[idx]


def test_set_invalid_active_index(Workbook):
    wb = Workbook()
    with pytest.raises(ValueError):
        wb.active = 1


def test_set_invalid_sheet_by_name(Workbook):
    wb = Workbook()
    with pytest.raises(TypeError):
        wb.active = "Sheet"


def test_set_invalid_child_as_active(Workbook):
    wb1 = Workbook()
    wb2 = Workbook()
    ws2 = wb2['Sheet']
    with pytest.raises(ValueError):
        wb1.active = ws2


def test_set_hidden_sheet_as_active(Workbook):
    wb = Workbook()
    ws = wb.create_sheet()
    ws.sheet_state = 'hidden'
    with pytest.raises(ValueError):
        wb.active = ws


def test_no_active(Workbook):
    wb = Workbook(write_only=True)
    assert wb.active is None


def test_create_sheet(Workbook):
    wb = Workbook()
    new_sheet = wb.create_sheet()
    assert new_sheet == wb.worksheets[-1]

def test_create_sheet_with_name(Workbook):
    wb = Workbook()
    new_sheet = wb.create_sheet(title='LikeThisName')
    assert new_sheet == wb.worksheets[-1]

def test_add_correct_sheet(Workbook):
    wb = Workbook()
    new_sheet = wb.create_sheet()
    wb._add_sheet(new_sheet)
    assert new_sheet == wb.worksheets[2]

def test_add_sheetname(Workbook):
    wb = Workbook()
    with pytest.raises(TypeError):
        wb._add_sheet("Test")


def test_add_sheet_from_other_workbook(Workbook):
    wb1 = Workbook()
    wb2 = Workbook()
    ws = wb1.active
    with pytest.raises(ValueError):
        wb2._add_sheet(ws)


def test_create_sheet_readonly(Workbook):
    wb = Workbook()
    wb._read_only = True
    with pytest.raises(ReadOnlyWorkbookException):
        wb.create_sheet()


def test_remove_sheet(Workbook):
    wb = Workbook()
    new_sheet = wb.create_sheet(0)
    wb.remove(new_sheet)
    assert new_sheet not in wb.worksheets


def test_move_sheet(Workbook):
    wb = Workbook()
    for i in range(9):
        wb.create_sheet()
    assert wb.sheetnames == ['Sheet', 'Sheet1', 'Sheet2', 'Sheet3', 'Sheet4',
                            'Sheet5', 'Sheet6', 'Sheet7', 'Sheet8', 'Sheet9']
    ws = wb['Sheet9']
    wb.move_sheet(ws, -5)
    assert wb.sheetnames == ['Sheet', 'Sheet1', 'Sheet2', 'Sheet3', 'Sheet9',
                            'Sheet4', 'Sheet5', 'Sheet6', 'Sheet7', 'Sheet8']


def test_move_sheet_by_name(Workbook):
    wb = Workbook()
    for i in range(9):
        wb.create_sheet()
    assert wb.sheetnames == ['Sheet', 'Sheet1', 'Sheet2', 'Sheet3', 'Sheet4',
                            'Sheet5', 'Sheet6', 'Sheet7', 'Sheet8', 'Sheet9']
    wb.move_sheet("Sheet9", -5)
    assert wb.sheetnames == ['Sheet', 'Sheet1', 'Sheet2', 'Sheet3', 'Sheet9',
                            'Sheet4', 'Sheet5', 'Sheet6', 'Sheet7', 'Sheet8']

def test_getitem(Workbook):
    wb = Workbook()
    ws = wb['Sheet']
    assert isinstance(ws, Worksheet)
    with pytest.raises(KeyError):
        wb['NotThere']


def test_get_chartsheet(Workbook):
    wb = Workbook()
    cs = wb.create_chartsheet()
    assert wb[cs.title] is cs


def test_del_worksheet(Workbook):
    wb = Workbook()
    del wb['Sheet']
    assert wb.worksheets == []


def test_del_chartsheet(Workbook):
    wb = Workbook()
    cs = wb.create_chartsheet()
    del wb[cs.title]
    assert wb.chartsheets == []


def test_contains(Workbook):
    wb = Workbook()
    assert "Sheet" in wb
    assert "NotThere" not in wb

def test_iter(Workbook):
    wb = Workbook()
    for ws in wb:
        pass
    assert ws.title == "Sheet"

def test_index(Workbook):
    wb = Workbook()
    new_sheet = wb.create_sheet()
    sheet_index = wb.index(new_sheet)
    assert sheet_index == 1


def test_get_sheet_names(Workbook):
    wb = Workbook()
    names = ['Sheet', 'Sheet1', 'Sheet2', 'Sheet3', 'Sheet4', 'Sheet5']
    for count in range(5):
        wb.create_sheet(0)
    assert wb.sheetnames == names


def test_add_invalid_worksheet_class_instance(Workbook):

    class AlternativeWorksheet:
        def __init__(self, parent_workbook, title=None):
            self.parent_workbook = parent_workbook
            if not title:
                title = 'AlternativeSheet'
            self.title = title

    wb = Workbook
    ws = AlternativeWorksheet(parent_workbook=wb)
    with pytest.raises(TypeError):
        wb._add_sheet(worksheet=ws)


class TestCopy:


    def test_worksheet_copy(self, Workbook):
        wb = Workbook()
        ws1 = wb.active
        ws2 = wb.copy_worksheet(ws1)
        assert ws2 is not None


    @pytest.mark.parametrize("title, copy",
                             [
                                 ("TestSheet", "TestSheet Copy"),
                                 (u"D\xfcsseldorf", u"D\xfcsseldorf Copy")
                                 ]
                             )
    def test_worksheet_copy_name(self, title, copy, Workbook):
        wb = Workbook()
        ws1 = wb.active
        ws1.title = title
        ws2 = wb.copy_worksheet(ws1)
        assert ws2.title == copy


    def test_cannot_copy_readonly(self, Workbook):
        wb = Workbook()
        ws = wb.active
        wb._read_only = True
        with pytest.raises(ValueError):
            wb.copy_worksheet(ws)


    def test_cannot_copy_writeonly(self, Workbook):
        wb = Workbook(write_only=True)
        ws = wb.create_sheet()
        with pytest.raises(ValueError):
            wb.copy_worksheet(ws)


    def test_default_epoch(self, Workbook):
        wb = Workbook()
        assert wb.epoch == datetime.datetime(1899, 12, 30)


    def test_assign_epoch(self, Workbook):
        wb = Workbook()
        wb.epoch = datetime.datetime(1904, 1, 1)


    def test_invalid_epoch(self, Workbook):
        wb = Workbook()
        with pytest.raises(ValueError):
            wb.epoch = datetime.datetime(1970, 1, 1)


class TestSheetnameCaching:
    """Regression tests for sheetnames caching (issue #36)."""

    def test_cache_invalidated_on_create(self, Workbook):
        wb = Workbook()
        _ = wb.sheetnames  # prime cache
        wb.create_sheet("New")
        assert "New" in wb.sheetnames

    def test_cache_invalidated_on_remove(self, Workbook):
        wb = Workbook()
        ws = wb.create_sheet("Temp")
        _ = wb.sheetnames  # prime cache
        wb.remove(ws)
        assert "Temp" not in wb.sheetnames

    def test_cache_invalidated_on_move(self, Workbook):
        wb = Workbook()
        wb.create_sheet("B")
        _ = wb.sheetnames  # prime cache
        wb.move_sheet("B", -1)
        assert wb.sheetnames[0] == "B"

    def test_cache_invalidated_on_rename(self, Workbook):
        wb = Workbook()
        _ = wb.sheetnames  # prime cache
        wb.active.title = "Renamed"
        assert wb.sheetnames == ["Renamed"]

    def test_sheetnames_returns_copy(self, Workbook):
        """Mutating the returned list must not corrupt the cache."""
        wb = Workbook()
        names = wb.sheetnames
        names.append("Phantom")
        assert "Phantom" not in wb.sheetnames


class TestSheetLookup:
    """Tests for O(1) sheet lookup by title (issue #40)."""

    def test_getitem_returns_correct_sheet(self, Workbook):
        wb = Workbook()
        ws1 = wb.active
        ws2 = wb.create_sheet("Second")
        ws3 = wb.create_sheet("Third")
        assert wb["Sheet"] is ws1
        assert wb["Second"] is ws2
        assert wb["Third"] is ws3

    def test_getitem_miss_raises_keyerror(self, Workbook):
        wb = Workbook()
        with pytest.raises(KeyError):
            wb["NoSuchSheet"]

    def test_contains_finds_existing_sheet(self, Workbook):
        wb = Workbook()
        wb.create_sheet("Alpha")
        assert "Alpha" in wb
        assert "Sheet" in wb

    def test_contains_rejects_missing_sheet(self, Workbook):
        wb = Workbook()
        assert "Missing" not in wb

    def test_getitem_after_add(self, Workbook):
        wb = Workbook()
        _ = wb["Sheet"]  # prime any cache
        ws = wb.create_sheet("Added")
        assert wb["Added"] is ws

    def test_getitem_after_remove(self, Workbook):
        wb = Workbook()
        wb.create_sheet("Temp")
        _ = wb["Temp"]  # prime any cache
        del wb["Temp"]
        with pytest.raises(KeyError):
            wb["Temp"]

    def test_getitem_after_rename(self, Workbook):
        wb = Workbook()
        ws = wb.active
        _ = wb["Sheet"]  # prime any cache
        ws.title = "Renamed"
        assert wb["Renamed"] is ws
        with pytest.raises(KeyError):
            wb["Sheet"]

    def test_contains_after_rename(self, Workbook):
        wb = Workbook()
        _ = "Sheet" in wb  # prime any cache
        wb.active.title = "Renamed"
        assert "Renamed" in wb
        assert "Sheet" not in wb

    def test_getitem_after_move(self, Workbook):
        wb = Workbook()
        ws_b = wb.create_sheet("B")
        wb.move_sheet("B", -1)
        assert wb["B"] is ws_b

    def test_getitem_uses_cached_map(self, Workbook):
        """After first lookup, subsequent lookups should use a cached map
        rather than scanning _sheets linearly."""
        wb = Workbook()
        for i in range(10):
            wb.create_sheet(f"S{i}")
        # Prime the cache
        _ = wb["S5"]
        # The workbook should now have a title map cache
        assert wb._sheet_title_map is not None

    def test_contains_uses_cached_map(self, Workbook):
        wb = Workbook()
        wb.create_sheet("Test")
        _ = "Test" in wb
        assert wb._sheet_title_map is not None


class TestStyleMaterialization:

    def test_materialize_pending_registers_shared_tables(self, Workbook):
        from fastpyxl.styles import Font

        wb = Workbook()
        cell = wb.active["A1"]
        nfonts = len(wb._fonts)
        cell.font = Font(bold=True)
        assert len(wb._fonts) == nfonts

        wb.materialize_pending_style_components(cell)
        assert len(wb._fonts) == nfonts + 1

    def test_save_prepares_styles_before_sheet_xml(self, Workbook):
        from fastpyxl import load_workbook
        from fastpyxl.styles import Font

        wb = Workbook()
        ws = wb.active
        ws["A1"].font = Font(name="Courier", sz=11)
        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        wb2 = load_workbook(buf)
        assert wb2.active["A1"].font.name == "Courier"
        assert wb2.active["A1"].font.size == 11

    def test_named_style_registration_deferred_until_save(self, Workbook):
        from fastpyxl.styles import NamedStyle

        wb = Workbook()
        cell = wb.active["A1"]
        before = len(wb._named_styles)
        cell.style = NamedStyle(name="SaveDeferred")
        assert len(wb._named_styles) == before

        buf = BytesIO()
        wb.save(buf)
        assert any(ns.name == "SaveDeferred" for ns in wb._named_styles)

