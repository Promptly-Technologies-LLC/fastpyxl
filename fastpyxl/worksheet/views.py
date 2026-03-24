# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.descriptors.excel import ExtensionList
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import Field


def _enum_converter(value, allowed, field_name):
    if value is None:
        return None
    if value not in allowed:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return value


class Pane(Serialisable):
    xSplit: float | None = Field.attribute(expected_type=float, allow_none=True)
    ySplit: float | None = Field.attribute(expected_type=float, allow_none=True)
    topLeftCell: str | None = Field.attribute(expected_type=str, allow_none=True)
    activePane: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("bottomRight", "topRight", "bottomLeft", "topLeft"), "activePane"),
    )
    state: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("split", "frozen", "frozenSplit"), "state"),
    )

    def __init__(self, xSplit=None, ySplit=None, topLeftCell=None, activePane="topLeft", state="split"):
        self.xSplit = xSplit
        self.ySplit = ySplit
        self.topLeftCell = topLeftCell
        self.activePane = activePane
        self.state = state


class Selection(Serialisable):
    pane: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("bottomRight", "topRight", "bottomLeft", "topLeft"), "pane"),
    )
    activeCell: str | None = Field.attribute(expected_type=str, allow_none=True)
    activeCellId: int | None = Field.attribute(expected_type=int, allow_none=True)
    sqref: str | None = Field.attribute(expected_type=str, allow_none=True)

    def __init__(self, pane=None, activeCell="A1", activeCellId=None, sqref="A1"):
        self.pane = pane
        self.activeCell = activeCell
        self.activeCellId = activeCellId
        self.sqref = sqref


class SheetView(Serialisable):
    """Information about the visible portions of this sheet."""

    tagname = "sheetView"

    windowProtection: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showFormulas: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showGridLines: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showRowColHeaders: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showZeros: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    rightToLeft: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    tabSelected: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showRuler: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showOutlineSymbols: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    defaultGridColor: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showWhiteSpace: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    view: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("normal", "pageBreakPreview", "pageLayout"), "view"),
    )
    topLeftCell: str | None = Field.attribute(expected_type=str, allow_none=True)
    colorId: int | None = Field.attribute(expected_type=int, allow_none=True)
    zoomScale: int | None = Field.attribute(expected_type=int, allow_none=True)
    zoomScaleNormal: int | None = Field.attribute(expected_type=int, allow_none=True)
    zoomScaleSheetLayoutView: int | None = Field.attribute(expected_type=int, allow_none=True)
    zoomScalePageLayoutView: int | None = Field.attribute(expected_type=int, allow_none=True)
    zoomToFit: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    workbookViewId: int | None = Field.attribute(expected_type=int, allow_none=True)
    selection: list[Selection] = Field.sequence(expected_type=Selection, default=list)
    pane: Pane | None = Field.element(expected_type=Pane, allow_none=True)

    xml_order = ("pane", "selection")

    def __init__(self, windowProtection=None, showFormulas=None, showGridLines=None, showRowColHeaders=None,
                 showZeros=None, rightToLeft=None, tabSelected=None, showRuler=None, showOutlineSymbols=None,
                 defaultGridColor=None, showWhiteSpace=None, view=None, topLeftCell=None, colorId=None,
                 zoomScale=None, zoomScaleNormal=None, zoomScaleSheetLayoutView=None, zoomScalePageLayoutView=None,
                 zoomToFit=None, workbookViewId=0, selection=None, pane=None):
        self.windowProtection = windowProtection
        self.showFormulas = showFormulas
        self.showGridLines = showGridLines
        self.showRowColHeaders = showRowColHeaders
        self.showZeros = showZeros
        self.rightToLeft = rightToLeft
        self.tabSelected = tabSelected
        self.showRuler = showRuler
        self.showOutlineSymbols = showOutlineSymbols
        self.defaultGridColor = defaultGridColor
        self.showWhiteSpace = showWhiteSpace
        self.view = view
        self.topLeftCell = topLeftCell
        self.colorId = colorId
        self.zoomScale = zoomScale
        self.zoomScaleNormal = zoomScaleNormal
        self.zoomScaleSheetLayoutView = zoomScaleSheetLayoutView
        self.zoomScalePageLayoutView = zoomScalePageLayoutView
        self.zoomToFit = zoomToFit
        self.workbookViewId = workbookViewId
        self.pane = pane
        if selection is None:
            selection = (Selection(),)
        self.selection = list(selection)


class SheetViewList(Serialisable):

    tagname = "sheetViews"

    sheetView: list[SheetView] = Field.sequence(expected_type=SheetView, default=list)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True, serialize=False)

    xml_order = ("sheetView",)

    def __init__(self, sheetView=None, extLst=None):
        if sheetView is None:
            sheetView = [SheetView()]
        self.sheetView = sheetView
        self.extLst = extLst

    @property
    def active(self):
        """
        Returns the first sheet view which is assumed to be active
        """
        return self.sheetView[0]
