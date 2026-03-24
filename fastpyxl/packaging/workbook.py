# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.descriptors.excel import ExtensionList

from fastpyxl.xml.constants import REL_NS, SHEET_MAIN_NS

from fastpyxl.workbook.defined_name import DefinedNameList
from fastpyxl.workbook.external_reference import ExternalReference
from fastpyxl.workbook.function_group import FunctionGroupList
from fastpyxl.workbook.properties import WorkbookProperties, CalcProperties, FileVersion
from fastpyxl.workbook.protection import WorkbookProtection, FileSharing
from fastpyxl.workbook.smart_tags import SmartTagList, SmartTagProperties
from fastpyxl.workbook.views import CustomWorkbookView, BookView
from fastpyxl.workbook.web import WebPublishing, WebPublishObjectList
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field


class FileRecoveryProperties(Serialisable):

    tagname = "fileRecoveryPr"

    autoRecover: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    crashSave: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    dataExtractLoad: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    repairLoad: bool | None = Field.attribute(expected_type=bool, allow_none=True)

    def __init__(self,
                 autoRecover=None,
                 crashSave=None,
                 dataExtractLoad=None,
                 repairLoad=None,
                ):
        self.autoRecover = autoRecover
        self.crashSave = crashSave
        self.dataExtractLoad = dataExtractLoad
        self.repairLoad = repairLoad


class ChildSheet(Serialisable):
    """
    Represents a reference to a worksheet or chartsheet in workbook.xml

    It contains the title, order and state but only an indirect reference to
    the objects themselves.
    """

    tagname = "sheet"

    name: str | None = Field.attribute(expected_type=str, allow_none=True)
    sheetId: int | None = Field.attribute(expected_type=int, allow_none=True)
    state: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("visible", "hidden", "veryHidden"), "state"),
    )
    id: str | None = Field.attribute(expected_type=str, allow_none=True, namespace=REL_NS)

    def __init__(self,
                 name=None,
                 sheetId=None,
                 state="visible",
                 id=None,
                ):
        self.name = name
        self.sheetId = sheetId
        self.state = state
        self.id = id


class PivotCache(Serialisable):

    tagname = "pivotCache"

    cacheId: int | None = Field.attribute(expected_type=int, allow_none=True)
    id: str | None = Field.attribute(expected_type=str, allow_none=True, namespace=REL_NS)

    def __init__(self,
                 cacheId=None,
                 id=None
                ):
        self.cacheId = cacheId
        self.id = id


class WorkbookPackage(Serialisable):
    """
    Represent the workbook file in the archive
    """

    tagname = "workbook"

    conformance: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("strict", "transitional"), "conformance"),
    )
    fileVersion: FileVersion | None = Field.element(expected_type=FileVersion, allow_none=True)
    fileSharing: FileSharing | None = Field.element(expected_type=FileSharing, allow_none=True)
    workbookPr: WorkbookProperties | None = Field.element(expected_type=WorkbookProperties, allow_none=True)
    properties = AliasField("workbookPr")
    workbookProtection: WorkbookProtection | None = Field.element(
        expected_type=WorkbookProtection, allow_none=True
    )
    bookViews: list[BookView] = Field.nested_sequence(expected_type=BookView, default=list)
    sheets: list[ChildSheet] = Field.nested_sequence(expected_type=ChildSheet, default=list)
    functionGroups: FunctionGroupList | None = Field.element(expected_type=FunctionGroupList, allow_none=True)
    externalReferences: list[ExternalReference] = Field.nested_sequence(
        expected_type=ExternalReference, default=list
    )
    definedNames: DefinedNameList | None = Field.element(expected_type=DefinedNameList, allow_none=True)
    calcPr: CalcProperties | None = Field.element(expected_type=CalcProperties, allow_none=True)
    oleSize: str | None = Field.nested_value(expected_type=str, allow_none=True, value_attribute="ref")
    customWorkbookViews: list[CustomWorkbookView] = Field.nested_sequence(
        expected_type=CustomWorkbookView, default=list
    )
    pivotCaches: list[PivotCache] = Field.nested_sequence(expected_type=PivotCache, allow_none=True, default=list)
    smartTagPr: SmartTagProperties | None = Field.element(expected_type=SmartTagProperties, allow_none=True)
    smartTagTypes: SmartTagList | None = Field.element(expected_type=SmartTagList, allow_none=True)
    webPublishing: WebPublishing | None = Field.element(expected_type=WebPublishing, allow_none=True)
    fileRecoveryPr: FileRecoveryProperties | None = Field.element(
        expected_type=FileRecoveryProperties, allow_none=True
    )
    webPublishObjects: WebPublishObjectList | None = Field.element(
        expected_type=WebPublishObjectList, allow_none=True
    )
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True, serialize=False)
    Ignorable: str | None = Field.nested_value(
        expected_type=str,
        namespace="http://schemas.openxmlformats.org/markup-compatibility/2006",
        allow_none=True,
        serialize=False,
    )

    xml_order = (
        "fileVersion",
        "fileSharing",
        "workbookPr",
        "workbookProtection",
        "bookViews",
        "sheets",
        "functionGroups",
        "externalReferences",
        "definedNames",
        "calcPr",
        "oleSize",
        "customWorkbookViews",
        "pivotCaches",
        "smartTagPr",
        "smartTagTypes",
        "webPublishing",
        "fileRecoveryPr",
        "webPublishObjects",
    )

    def __init__(self,
                 conformance=None,
                 fileVersion=None,
                 fileSharing=None,
                 workbookPr=None,
                 workbookProtection=None,
                 bookViews=(),
                 sheets=(),
                 functionGroups=None,
                 externalReferences=(),
                 definedNames=None,
                 calcPr=None,
                 oleSize=None,
                 customWorkbookViews=(),
                 pivotCaches=(),
                 smartTagPr=None,
                 smartTagTypes=None,
                 webPublishing=None,
                 fileRecoveryPr=None,
                 webPublishObjects=None,
                 extLst=None,
                 Ignorable=None,
                ):
        self.conformance = conformance
        self.fileVersion = fileVersion
        self.fileSharing = fileSharing
        if workbookPr is None:
            workbookPr = WorkbookProperties()
        self.workbookPr = workbookPr
        self.workbookProtection = workbookProtection
        self.bookViews = list(bookViews)
        self.sheets = list(sheets)
        self.functionGroups = functionGroups
        self.externalReferences = list(externalReferences)
        self.definedNames = definedNames
        self.calcPr = calcPr
        self.oleSize = oleSize
        self.customWorkbookViews = list(customWorkbookViews)
        self.pivotCaches = list(pivotCaches)
        self.smartTagPr = smartTagPr
        self.smartTagTypes = smartTagTypes
        self.webPublishing = webPublishing
        self.fileRecoveryPr = fileRecoveryPr
        self.webPublishObjects = webPublishObjects
        self.extLst = extLst
        self.Ignorable = Ignorable

    def to_tree(self, tagname=None, idx=None, namespace=None):
        tree = super().to_tree(tagname, idx, namespace)
        tree.set("xmlns", SHEET_MAIN_NS)
        return tree

    @property
    def active(self):
        for view in self.bookViews:
            if view.activeTab is not None:
                return view.activeTab
        return 0


def _enum_converter(value, allowed_values, field_name: str):
    if value is None:
        return None
    if value not in allowed_values:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return value
