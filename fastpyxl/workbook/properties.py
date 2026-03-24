# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import Field


class WorkbookProperties(Serialisable):

    tagname = "workbookPr"

    date1904: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    dateCompatibility: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showObjects: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("all", "placeholders"), "showObjects"), default=None,
    )
    showBorderUnselectedTables: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    filterPrivacy: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    promptedSolutions: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showInkAnnotation: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    backupFile: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    saveExternalLinkValues: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    updateLinks: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("userSet", "never", "always"), "updateLinks"), default=None,
    )
    codeName: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    hidePivotFieldList: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showPivotChartFilter: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    allowRefreshQuery: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    publishItems: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    checkCompatibility: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    autoCompressPictures: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    refreshAllConnections: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    defaultThemeVersion: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

    def __init__(self,
                 date1904=None,
                 dateCompatibility=None,
                 showObjects=None,
                 showBorderUnselectedTables=None,
                 filterPrivacy=None,
                 promptedSolutions=None,
                 showInkAnnotation=None,
                 backupFile=None,
                 saveExternalLinkValues=None,
                 updateLinks=None,
                 codeName=None,
                 hidePivotFieldList=None,
                 showPivotChartFilter=None,
                 allowRefreshQuery=None,
                 publishItems=None,
                 checkCompatibility=None,
                 autoCompressPictures=None,
                 refreshAllConnections=None,
                 defaultThemeVersion=None,
                ):
        self.date1904 = date1904
        self.dateCompatibility = dateCompatibility
        self.showObjects = showObjects
        self.showBorderUnselectedTables = showBorderUnselectedTables
        self.filterPrivacy = filterPrivacy
        self.promptedSolutions = promptedSolutions
        self.showInkAnnotation = showInkAnnotation
        self.backupFile = backupFile
        self.saveExternalLinkValues = saveExternalLinkValues
        self.updateLinks = updateLinks
        self.codeName = codeName
        self.hidePivotFieldList = hidePivotFieldList
        self.showPivotChartFilter = showPivotChartFilter
        self.allowRefreshQuery = allowRefreshQuery
        self.publishItems = publishItems
        self.checkCompatibility = checkCompatibility
        self.autoCompressPictures = autoCompressPictures
        self.refreshAllConnections = refreshAllConnections
        self.defaultThemeVersion = defaultThemeVersion


class CalcProperties(Serialisable):

    tagname = "calcPr"

    calcId: int | None = Field.attribute(expected_type=int, allow_none=True, default=124519)
    calcMode: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("manual", "auto", "autoNoTable"), "calcMode"), default=None,
    )
    fullCalcOnLoad: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    refMode: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("A1", "R1C1"), "refMode"), default=None,
    )
    iterate: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    iterateCount: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    iterateDelta: float | None = Field.attribute(expected_type=float, allow_none=True, default=None)
    fullPrecision: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    calcCompleted: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    calcOnSave: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    concurrentCalc: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    concurrentManualCount: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    forceFullCalc: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)

    def __init__(self,
                 calcId=124519,
                 calcMode=None,
                 fullCalcOnLoad=True,
                 refMode=None,
                 iterate=None,
                 iterateCount=None,
                 iterateDelta=None,
                 fullPrecision=None,
                 calcCompleted=None,
                 calcOnSave=None,
                 concurrentCalc=None,
                 concurrentManualCount=None,
                 forceFullCalc=None,
                ):
        self.calcId = calcId
        self.calcMode = calcMode
        self.fullCalcOnLoad = fullCalcOnLoad
        self.refMode = refMode
        self.iterate = iterate
        self.iterateCount = iterateCount
        self.iterateDelta = iterateDelta
        self.fullPrecision = fullPrecision
        self.calcCompleted = calcCompleted
        self.calcOnSave = calcOnSave
        self.concurrentCalc = concurrentCalc
        self.concurrentManualCount = concurrentManualCount
        self.forceFullCalc = forceFullCalc


class FileVersion(Serialisable):

    tagname = "fileVersion"

    appName: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    lastEdited: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    lowestEdited: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    rupBuild: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    codeName: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)

    def __init__(self,
                 appName=None,
                 lastEdited=None,
                 lowestEdited=None,
                 rupBuild=None,
                 codeName=None,
                ):
        self.appName = appName
        self.lastEdited = lastEdited
        self.lowestEdited = lowestEdited
        self.rupBuild = rupBuild
        self.codeName = codeName


def _enum_converter(value, allowed_values, field_name: str):
    if value is None:
        return None
    if value not in allowed_values:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return value
