# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.descriptors.excel import (
    ExtensionList,
)
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import Field


class BookView(Serialisable):

    tagname = "workbookView"

    visibility: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("visible", "hidden", "veryHidden"), "visibility"), default=None,
    )
    minimized: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showHorizontalScroll: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showVerticalScroll: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showSheetTabs: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    xWindow: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    yWindow: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    windowWidth: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    windowHeight: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    tabRatio: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    firstSheet: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    activeTab: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    autoFilterDateGrouping: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True, default=None)

    xml_order = ()

    def __init__(self,
                 visibility="visible",
                 minimized=False,
                 showHorizontalScroll=True,
                 showVerticalScroll=True,
                 showSheetTabs=True,
                 xWindow=None,
                 yWindow=None,
                 windowWidth=None,
                 windowHeight=None,
                 tabRatio=600,
                 firstSheet=0,
                 activeTab=0,
                 autoFilterDateGrouping=True,
                 extLst=None,
                ):
        self.visibility = visibility
        self.minimized = minimized
        self.showHorizontalScroll = showHorizontalScroll
        self.showVerticalScroll = showVerticalScroll
        self.showSheetTabs = showSheetTabs
        self.xWindow = xWindow
        self.yWindow = yWindow
        self.windowWidth = windowWidth
        self.windowHeight = windowHeight
        self.tabRatio = tabRatio
        self.firstSheet = firstSheet
        self.activeTab = activeTab
        self.autoFilterDateGrouping = autoFilterDateGrouping
        self.extLst = extLst


class CustomWorkbookView(Serialisable):

    tagname = "customWorkbookView"

    name: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    guid: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    autoUpdate: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    mergeInterval: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    changesSavedWin: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    onlySync: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    personalView: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    includePrintSettings: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    includeHiddenRowCol: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    maximized: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    minimized: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showHorizontalScroll: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showVerticalScroll: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showSheetTabs: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    xWindow: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    yWindow: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    windowWidth: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    windowHeight: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    tabRatio: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    activeSheetId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    showFormulaBar: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showStatusbar: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showComments: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("commNone", "commIndicator", "commIndAndComment"), "showComments"), default=None,
    )
    showObjects: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("all", "placeholders"), "showObjects"), default=None,
    )
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True, default=None)

    xml_order = ()

    def __init__(self,
                 name=None,
                 guid=None,
                 autoUpdate=None,
                 mergeInterval=None,
                 changesSavedWin=None,
                 onlySync=None,
                 personalView=None,
                 includePrintSettings=None,
                 includeHiddenRowCol=None,
                 maximized=None,
                 minimized=None,
                 showHorizontalScroll=None,
                 showVerticalScroll=None,
                 showSheetTabs=None,
                 xWindow=None,
                 yWindow=None,
                 windowWidth=None,
                 windowHeight=None,
                 tabRatio=None,
                 activeSheetId=None,
                 showFormulaBar=None,
                 showStatusbar=None,
                 showComments="commIndicator",
                 showObjects="all",
                 extLst=None,
                ):
        self.name = name
        self.guid = guid
        self.autoUpdate = autoUpdate
        self.mergeInterval = mergeInterval
        self.changesSavedWin = changesSavedWin
        self.onlySync = onlySync
        self.personalView = personalView
        self.includePrintSettings = includePrintSettings
        self.includeHiddenRowCol = includeHiddenRowCol
        self.maximized = maximized
        self.minimized = minimized
        self.showHorizontalScroll = showHorizontalScroll
        self.showVerticalScroll = showVerticalScroll
        self.showSheetTabs = showSheetTabs
        self.xWindow = xWindow
        self.yWindow = yWindow
        self.windowWidth = windowWidth
        self.windowHeight = windowHeight
        self.tabRatio = tabRatio
        self.activeSheetId = activeSheetId
        self.showFormulaBar = showFormulaBar
        self.showStatusbar = showStatusbar
        self.showComments = showComments
        self.showObjects = showObjects
        self.extLst = extLst


def _enum_converter(value, allowed_values, field_name: str):
    if value is None:
        return None
    if value not in allowed_values:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return value
