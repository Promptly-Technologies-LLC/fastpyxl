# Copyright (c) 2010-2024 fastpyxl

from datetime import datetime

from fastpyxl.compat import safe_string
from fastpyxl.descriptors.excel import HexBinary
from fastpyxl.packaging.core import _datetime_converter
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import Field

class Index(Serialisable):

    tagname = "x"

    v: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

    def __init__(self,
                 v=0,
                ):
        self.v = v


class Tuple(Serialisable):

    tagname = "tpl"

    fld: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    hier: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    item: int | None = Field.attribute(expected_type=int, allow_none=False, default=None)

    def __init__(self,
                 fld=None,
                 hier=None,
                 item=None,
                ):
        self.fld = fld
        self.hier = hier
        self.item = item


class TupleList(Serialisable):

    tagname = "tpls"

    c: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    tpl: Tuple | None = Field.element(expected_type=Tuple, allow_none=True, default=None)

    xml_order = ("tpl",)

    def __init__(self,
                 c=None,
                 tpl=None,
                ):
        self.c = c
        self.tpl = tpl


class Missing(Serialisable):

    tagname = "m"

    tpls: list[TupleList] = Field.sequence(expected_type=TupleList, default=list)
    x: list[Index] = Field.sequence(expected_type=Index, default=list)
    u: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    f: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    c: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    cp: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    _in: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    bc: str | None = Field.attribute(expected_type=HexBinary, allow_none=True, default=None)
    fc: str | None = Field.attribute(expected_type=HexBinary, allow_none=True, default=None)
    i: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    un: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    st: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    b: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)

    xml_order = ("tpls", "x")

    def __init__(self,
                 tpls=(),
                 x=(),
                 u=None,
                 f=None,
                 c=None,
                 cp=None,
                 _in=None,
                 bc=None,
                 fc=None,
                 i=None,
                 un=None,
                 st=None,
                 b=None,
                ):
        self.tpls = list(tpls)
        self.x = list(x)
        self.u = u
        self.f = f
        self.c = c
        self.cp = cp
        self._in = _in
        self.bc = bc
        self.fc = fc
        self.i = i
        self.un = un
        self.st = st
        self.b = b


class Number(Serialisable):

    tagname = "n"

    tpls: list[TupleList] = Field.sequence(expected_type=TupleList, default=list)
    x: list[Index] = Field.sequence(expected_type=Index, default=list)
    v: float | None = Field.attribute(expected_type=float, allow_none=False, default=None)
    u: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    f: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    c: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    cp: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    _in: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    bc: str | None = Field.attribute(expected_type=HexBinary, allow_none=True, default=None)
    fc: str | None = Field.attribute(expected_type=HexBinary, allow_none=True, default=None)
    i: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    un: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    st: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    b: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)

    xml_order = ("tpls", "x")

    def __init__(self,
                 tpls=(),
                 x=(),
                 v=None,
                 u=None,
                 f=None,
                 c=None,
                 cp=None,
                 _in=None,
                 bc=None,
                 fc=None,
                 i=None,
                 un=None,
                 st=None,
                 b=None,
                ):
        self.tpls = list(tpls)
        self.x = list(x)
        self.v = v
        self.u = u
        self.f = f
        self.c = c
        self.cp = cp
        self._in = _in
        self.bc = bc
        self.fc = fc
        self.i = i
        self.un = un
        self.st = st
        self.b = b


class Error(Serialisable):

    tagname = "e"

    tpls: TupleList | None = Field.element(expected_type=TupleList, allow_none=True, default=None)
    x: list[Index] = Field.sequence(expected_type=Index, default=list)
    v: str | None = Field.attribute(expected_type=str, allow_none=False, default=None)
    u: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    f: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    c: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    cp: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    _in: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    bc: str | None = Field.attribute(expected_type=HexBinary, allow_none=True, default=None)
    fc: str | None = Field.attribute(expected_type=HexBinary, allow_none=True, default=None)
    i: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    un: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    st: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    b: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)

    xml_order = ("tpls", "x")

    def __init__(self,
                 tpls=None,
                 x=(),
                 v=None,
                 u=None,
                 f=None,
                 c=None,
                 cp=None,
                 _in=None,
                 bc=None,
                 fc=None,
                 i=None,
                 un=None,
                 st=None,
                 b=None,
                ):
        self.tpls = tpls
        self.x = list(x)
        self.v = v
        self.u = u
        self.f = f
        self.c = c
        self.cp = cp
        self._in = _in
        self.bc = bc
        self.fc = fc
        self.i = i
        self.un = un
        self.st = st
        self.b = b


class Boolean(Serialisable):

    tagname = "b"

    x: list[Index] = Field.sequence(expected_type=Index, default=list)
    v: bool = Field.attribute(expected_type=bool, allow_none=False, default=False)
    u: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    f: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    c: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    cp: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

    xml_order = ("x",)

    def __init__(self,
                 x=(),
                 v=False,
                 u=None,
                 f=None,
                 c=None,
                 cp=None,
                ):
        self.x = list(x)
        self.v = v
        self.u = u
        self.f = f
        self.c = c
        self.cp = cp

    def __iter__(self):
        for attr in self.__attrs__:
            field = self.__fields__[attr]
            value = getattr(self, attr)
            xml_attr = field.tag
            if xml_attr.startswith("_"):
                xml_attr = xml_attr[1:]
            if field.hyphenated:
                xml_attr = xml_attr.replace("_", "-")
            if attr == "v":
                yield xml_attr, "1" if value else "0"
                continue
            if value is None:
                continue
            yield xml_attr, safe_string(value)


class Text(Serialisable):

    tagname = "s"

    tpls: list[TupleList] = Field.sequence(expected_type=TupleList, default=list)
    x: list[Index] = Field.sequence(expected_type=Index, default=list)
    v: str | None = Field.attribute(expected_type=str, allow_none=False, default=None)
    u: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    f: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    c: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    cp: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    _in: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    bc: str | None = Field.attribute(expected_type=HexBinary, allow_none=True, default=None)
    fc: str | None = Field.attribute(expected_type=HexBinary, allow_none=True, default=None)
    i: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    un: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    st: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    b: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)

    xml_order = ("tpls", "x")

    def __init__(self,
                 tpls=(),
                 x=(),
                 v=None,
                 u=None,
                 f=None,
                 c=None,
                 cp=None,
                 _in=None,
                 bc=None,
                 fc=None,
                 i=None,
                 un=None,
                 st=None,
                 b=None,
                 ):
        self.tpls = list(tpls)
        self.x = list(x)
        self.v = v
        self.u = u
        self.f = f
        self.c = c
        self.cp = cp
        self._in = _in
        self.bc = bc
        self.fc = fc
        self.i = i
        self.un = un
        self.st = st
        self.b = b


class DateTimeField(Serialisable):

    tagname = "d"

    x: list[Index] = Field.sequence(expected_type=Index, default=list)
    v: datetime | None = Field.attribute(
        expected_type=datetime,
        allow_none=False,
        converter=_datetime_converter, default=None,
    )
    u: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    f: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    c: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    cp: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

    xml_order = ("x",)

    def __init__(self,
                 x=(),
                 v=None,
                 u=None,
                 f=None,
                 c=None,
                 cp=None,
                 ):
        self.x = list(x)
        self.v = v
        self.u = u
        self.f = f
        self.c = c
        self.cp = cp
