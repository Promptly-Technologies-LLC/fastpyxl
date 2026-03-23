# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import Field

from fastpyxl.packaging.relationship import (
    Relationship,
    get_rels_path,
    get_dependents
    )
from fastpyxl.xml.constants import SHEET_MAIN_NS
from fastpyxl.xml.functions import fromstring


"""Manage links to external Workbooks"""


class ExternalCell(Serialisable):

    r: str | None = Field.attribute(expected_type=str, allow_none=True)
    t: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("b", "d", "n", "e", "s", "str", "inlineStr"), "t"),
    )
    vm: int | None = Field.attribute(expected_type=int, allow_none=True)
    v: str | None = Field.nested_text(expected_type=str, allow_none=True)

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

    r: int | None = Field.attribute(expected_type=int, allow_none=True)
    cell: list[ExternalCell] = Field.sequence(expected_type=ExternalCell)
    xml_order = ('cell',)

    def __init__(self,
                 r=(),
                 cell=None,
                ):
        self.r = r
        self.cell = cell


class ExternalSheetData(Serialisable):

    sheetId: int | None = Field.attribute(expected_type=int, allow_none=True)
    refreshError: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    row: list[ExternalRow] = Field.sequence(expected_type=ExternalRow)
    xml_order = ('row',)

    def __init__(self,
                 sheetId=None,
                 refreshError=None,
                 row=(),
                ):
        self.sheetId = sheetId
        self.refreshError = refreshError
        self.row = row


class ExternalSheetDataSet(Serialisable):

    sheetData: list[ExternalSheetData] | None = Field.sequence(expected_type=ExternalSheetData, allow_none=True)
    xml_order = ('sheetData',)

    def __init__(self,
                 sheetData=None,
                ):
        self.sheetData = sheetData


class ExternalSheetNames(Serialisable):

    sheetName: list[str] = Field.sequence(expected_type=str)
    xml_order = ('sheetName',)

    def __init__(self,
                 sheetName=(),
                ):
        self.sheetName = sheetName


class ExternalDefinedName(Serialisable):

    tagname = "definedName"

    name: str | None = Field.attribute(expected_type=str, allow_none=True)
    refersTo: str | None = Field.attribute(expected_type=str, allow_none=True)
    sheetId: int | None = Field.attribute(expected_type=int, allow_none=True)

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

    sheetNames: ExternalSheetNames | None = Field.element(expected_type=ExternalSheetNames, allow_none=True)
    definedNames: list[ExternalDefinedName] = Field.nested_sequence(expected_type=ExternalDefinedName)
    sheetDataSet: ExternalSheetDataSet | None = Field.element(expected_type=ExternalSheetDataSet, allow_none=True)
    id: str | None = Field.attribute(expected_type=str, allow_none=True)

    xml_order = ('sheetNames', 'definedNames', 'sheetDataSet')

    def __init__(self,
                 sheetNames=None,
                 definedNames=(),
                 sheetDataSet=None,
                 id=None,
                ):
        self.sheetNames = sheetNames
        self.definedNames = definedNames
        self.sheetDataSet = sheetDataSet
        self.id = id


class ExternalLink(Serialisable):

    tagname = "externalLink"

    _id = None
    _path = "/xl/externalLinks/externalLink{0}.xml"
    _rel_type = "externalLink"
    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"

    externalBook: ExternalBook | None = Field.element(expected_type=ExternalBook, allow_none=True)
    file_link: Relationship | None = Field.element(expected_type=Relationship, allow_none=True)  # link to external file

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
