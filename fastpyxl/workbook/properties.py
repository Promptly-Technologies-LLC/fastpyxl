# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import Field


class WorkbookProperties(Serialisable):

    tagname = "workbookPr"

    date1904: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    dateCompatibility: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showObjects: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("all", "placeholders"), "showObjects"),
    )
    showBorderUnselectedTables: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    filterPrivacy: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    promptedSolutions: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showInkAnnotation: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    backupFile: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    saveExternalLinkValues: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    updateLinks: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("userSet", "never", "always"), "updateLinks"),
    )
    codeName: str | None = Field.attribute(expected_type=str, allow_none=True)
    hidePivotFieldList: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showPivotChartFilter: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    allowRefreshQuery: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    publishItems: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    checkCompatibility: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    autoCompressPictures: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    refreshAllConnections: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    defaultThemeVersion: int | None = Field.attribute(expected_type=int, allow_none=True)

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
        converter=lambda v: _enum_converter(v, ("manual", "auto", "autoNoTable"), "calcMode"),
    )
    fullCalcOnLoad: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    refMode: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("A1", "R1C1"), "refMode"),
    )
    iterate: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    iterateCount: int | None = Field.attribute(expected_type=int, allow_none=True)
    iterateDelta: float | None = Field.attribute(expected_type=float, allow_none=True)
    fullPrecision: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    calcCompleted: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    calcOnSave: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    concurrentCalc: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    concurrentManualCount: int | None = Field.attribute(expected_type=int, allow_none=True)
    forceFullCalc: bool | None = Field.attribute(expected_type=bool, allow_none=True)

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

    appName: str | None = Field.attribute(expected_type=str, allow_none=True)
    lastEdited: str | None = Field.attribute(expected_type=str, allow_none=True)
    lowestEdited: str | None = Field.attribute(expected_type=str, allow_none=True)
    rupBuild: str | None = Field.attribute(expected_type=str, allow_none=True)
    codeName: str | None = Field.attribute(expected_type=str, allow_none=True)

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
