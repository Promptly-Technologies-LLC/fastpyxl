# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import Field

from fastpyxl.packaging.relationship import (
    Relationship,
    get_rels_path,
    get_dependents
    )
from fastpyxl.xml.constants import REL_NS, SHEET_MAIN_NS
from fastpyxl.xml.functions import fromstring


"""Manage links to external Workbooks"""


class ExternalCell(Serialisable):

    r: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    t: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("b", "d", "n", "e", "s", "str", "inlineStr"), "t"), default=None,
    )
    vm: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    v: str | None = Field.nested_text(expected_type=str, allow_none=True, default=None)

    def __init__(self,
                 r=None,
                 t=None,
                 vm=None,
                 v=None,
                ):
        self.r = r
        self.t = t
        self.vm = vm
        self.v = v


class ExternalRow(Serialisable):

    r: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    cell: list[ExternalCell] = Field.sequence(expected_type=ExternalCell, default=list)
    xml_order = ('cell',)

    def __init__(self,
                 r=None,
                 cell=None,
                ):
        self.r = r
        self.cell = list(cell or ())


class ExternalSheetData(Serialisable):

    sheetId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    refreshError: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    row: list[ExternalRow] = Field.sequence(expected_type=ExternalRow, default=list)
    xml_order = ('row',)

    def __init__(self,
                 sheetId=None,
                 refreshError=None,
                 row=(),
                ):
        self.sheetId = sheetId
        self.refreshError = refreshError
        self.row = list(row)


class ExternalSheetDataSet(Serialisable):

    sheetData: list[ExternalSheetData] | None = Field.sequence(expected_type=ExternalSheetData, allow_none=True, default=list)
    xml_order = ('sheetData',)

    def __init__(self,
                 sheetData=None,
                ):
        self.sheetData = sheetData


class ExternalSheetNames(Serialisable):

    sheetName: list[str] = Field.sequence(expected_type=str, primitive_attribute="val", default=list)
    xml_order = ('sheetName',)

    def __init__(self,
                 sheetName=(),
                ):
        self.sheetName = list(sheetName)


class ExternalDefinedName(Serialisable):

    tagname = "definedName"

    name: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    refersTo: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    sheetId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

    def __init__(self,
                 name=None,
                 refersTo=None,
                 sheetId=None,
                ):
        self.name = name
        self.refersTo = refersTo
        self.sheetId = sheetId


class ExternalBook(Serialisable):

    tagname = "externalBook"

    sheetNames: ExternalSheetNames | None = Field.element(expected_type=ExternalSheetNames, allow_none=True, default=None)
    definedNames: list[ExternalDefinedName] = Field.nested_sequence(expected_type=ExternalDefinedName, default=list)
    sheetDataSet: ExternalSheetDataSet | None = Field.element(expected_type=ExternalSheetDataSet, allow_none=True, default=None)
    id: str | None = Field.attribute(expected_type=str, allow_none=True, namespace=REL_NS, default=None)

    xml_order = ('sheetNames', 'definedNames', 'sheetDataSet')

    def __init__(self,
                 sheetNames=None,
                 definedNames=(),
                 sheetDataSet=None,
                 id=None,
                ):
        self.sheetNames = sheetNames
        self.definedNames = list(definedNames)
        self.sheetDataSet = sheetDataSet
        self.id = id


class ExternalLink(Serialisable):

    tagname = "externalLink"

    _id = None
    _path = "/xl/externalLinks/externalLink{0}.xml"
    _rel_type = "externalLink"
    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"

    externalBook: ExternalBook | None = Field.element(expected_type=ExternalBook, allow_none=True, default=None)
    file_link: Relationship | None = Field.element(
        expected_type=Relationship,
        allow_none=True,
        serialize=False, default=None,
    )

    xml_order = ('externalBook',)

    def __init__(self,
                 externalBook=None,
                 ddeLink=None,
                 oleLink=None,
                 extLst=None,
                ):
        self.externalBook = externalBook
        self.file_link = None
        # ignore other items for the moment.


    def to_tree(self):
        node = super().to_tree()
        node.set("xmlns", SHEET_MAIN_NS)
        return node


    @property
    def path(self):
        return self._path.format(self._id)


def read_external_link(archive, book_path):
    src = archive.read(book_path)
    node = fromstring(src)
    book = ExternalLink.from_tree(node)

    link_path = get_rels_path(book_path)
    deps = get_dependents(archive, link_path)
    book.file_link = deps[0]

    return book


def _enum_converter(value, allowed_values, field_name: str):
    if value is None:
        return None
    if value not in allowed_values:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return value
