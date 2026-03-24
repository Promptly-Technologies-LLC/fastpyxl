# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.descriptors.excel import ExtensionList

from fastpyxl.xml.constants import SHEET_MAIN_NS
from fastpyxl.xml.functions import tostring
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import Field

from .fields import (
    Boolean,
    Error,
    Missing,
    Number,
    Text,
    DateTimeField,
    Index,
    TupleList,  # noqa: F401 -- re-exported
)


class Record(Serialisable):

    tagname = "r"

    _fields: list[Missing | Number | Boolean | Error | Text | DateTimeField | Index] = Field.multi_sequence(
        parts={
            "m": Missing,
            "n": Number,
            "b": Boolean,
            "e": Error,
            "s": Text,
            "d": DateTimeField,
            "x": Index,
        }
    )


    def __init__(self,
                 _fields=(),
                 m=None,
                 n=None,
                 b=None,
                 e=None,
                 s=None,
                 d=None,
                 x=None,
                ):
        self._fields = list(_fields)


class RecordList(Serialisable):

    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"
    rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords"
    _id = 1
    _path = "/xl/pivotCache/pivotCacheRecords{0}.xml"

    tagname ="pivotCacheRecords"

    r: list[Record] | None = Field.sequence(expected_type=Record, allow_none=True)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True)

    xml_order = ("r",)

    def __init__(self,
                 count=None,
                 r=(),
                 extLst=None,
                ):
        self.r = list(r)
        self.extLst = extLst


    @property
    def count(self):
        r = self.r
        return len(r) if r is not None else 0

    def __iter__(self):
        yield from super().__iter__()
        yield "count", str(self.count)


    def to_tree(self):
        tree = super().to_tree()
        tree.set("xmlns", SHEET_MAIN_NS)
        return tree


    @property
    def path(self):
        return self._path.format(self._id)


    def _write(self, archive, manifest):
        """
        Write to zipfile and update manifest
        """
        xml = tostring(self.to_tree())
        archive.writestr(self.path[1:], xml)
        manifest.append(self)


    def _write_rels(self, archive, manifest):
        pass
