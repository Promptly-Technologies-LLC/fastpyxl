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
    xSplit: float | None = Field.attribute(expected_type=float, allow_none=True, default=None)
    ySplit: float | None = Field.attribute(expected_type=float, allow_none=True, default=None)
    topLeftCell: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    activePane: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("bottomRight", "topRight", "bottomLeft", "topLeft"), "activePane"), default=None,
    )
    state: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("split", "frozen", "frozenSplit"), "state"), default=None,
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
        converter=lambda v: _enum_converter(v, ("bottomRight", "topRight", "bottomLeft", "topLeft"), "pane"), default=None,
    )
    activeCell: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    activeCellId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    sqref: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)

    def __init__(self, pane=None, activeCell="A1", activeCellId=None, sqref="A1"):
        self.pane = pane
        self.activeCell = activeCell
        self.activeCellId = activeCellId
        self.sqref = sqref


class SheetView(Serialisable):
    """Information about the visible portions of this sheet."""

    tagname = "sheetView"

    windowProtection: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showFormulas: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showGridLines: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showRowColHeaders: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showZeros: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    rightToLeft: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    tabSelected: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showRuler: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showOutlineSymbols: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    defaultGridColor: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showWhiteSpace: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    view: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("normal", "pageBreakPreview", "pageLayout"), "view"), default=None,
    )
    topLeftCell: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    colorId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    zoomScale: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    zoomScaleNormal: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    zoomScaleSheetLayoutView: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    zoomScalePageLayoutView: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    zoomToFit: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    workbookViewId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    selection: list[Selection] = Field.sequence(expected_type=Selection, default=list)
    pane: Pane | None = Field.element(expected_type=Pane, allow_none=True, default=None)

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
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True, serialize=False, default=None)

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
