# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.descriptors.excel import ExtensionList
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import Field


class ChartsheetView(Serialisable):
    tagname = "sheetView"

    tabSelected: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    zoomScale: int | None = Field.attribute(expected_type=int, allow_none=True)
    workbookViewId: int | None = Field.attribute(expected_type=int, allow_none=True, default=0)
    zoomToFit: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True)
    xml_order = ()

    def __init__(self,
                 tabSelected=None,
                 zoomScale=None,
                 workbookViewId=0,
                 zoomToFit=True,
                 extLst=None,
                 ):
        self.tabSelected = tabSelected
        self.zoomScale = zoomScale
        self.workbookViewId = workbookViewId
        self.zoomToFit = zoomToFit
        self.extLst = extLst


class ChartsheetViewList(Serialisable):
    tagname = "sheetViews"

    sheetView: list[ChartsheetView] = Field.sequence(expected_type=ChartsheetView)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True)
    xml_order = ("sheetView", "extLst")

    def __init__(self,
                 sheetView=None,
                 extLst=None,
                 ):
        if sheetView is None:
            sheetView = [ChartsheetView()]
        self.sheetView = sheetView
        self.extLst = extLst
