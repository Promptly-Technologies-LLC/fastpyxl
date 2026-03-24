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
        converter=lambda v: _enum_converter(v, ("visible", "hidden", "veryHidden"), "visibility"),
    )
    minimized: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showHorizontalScroll: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showVerticalScroll: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showSheetTabs: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    xWindow: int | None = Field.attribute(expected_type=int, allow_none=True)
    yWindow: int | None = Field.attribute(expected_type=int, allow_none=True)
    windowWidth: int | None = Field.attribute(expected_type=int, allow_none=True)
    windowHeight: int | None = Field.attribute(expected_type=int, allow_none=True)
    tabRatio: int | None = Field.attribute(expected_type=int, allow_none=True)
    firstSheet: int | None = Field.attribute(expected_type=int, allow_none=True)
    activeTab: int | None = Field.attribute(expected_type=int, allow_none=True)
    autoFilterDateGrouping: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True)

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

    name: str | None = Field.attribute(expected_type=str, allow_none=True)
    guid: str | None = Field.attribute(expected_type=str, allow_none=True)
    autoUpdate: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    mergeInterval: int | None = Field.attribute(expected_type=int, allow_none=True)
    changesSavedWin: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    onlySync: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    personalView: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    includePrintSettings: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    includeHiddenRowCol: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    maximized: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    minimized: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showHorizontalScroll: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showVerticalScroll: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showSheetTabs: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    xWindow: int | None = Field.attribute(expected_type=int, allow_none=True)
    yWindow: int | None = Field.attribute(expected_type=int, allow_none=True)
    windowWidth: int | None = Field.attribute(expected_type=int, allow_none=True)
    windowHeight: int | None = Field.attribute(expected_type=int, allow_none=True)
    tabRatio: int | None = Field.attribute(expected_type=int, allow_none=True)
    activeSheetId: int | None = Field.attribute(expected_type=int, allow_none=True)
    showFormulaBar: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showStatusbar: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showComments: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("commNone", "commIndicator", "commIndAndComment"), "showComments"),
    )
    showObjects: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("all", "placeholders"), "showObjects"),
    )
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True)

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
