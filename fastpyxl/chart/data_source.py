"""
Collection of utility primitives for charts.
"""

from __future__ import annotations

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.descriptors.excel import ExtensionList


def _convert_number_value(value):
    if value is None:
        return None
    if value == "#N/A":
        return "#N/A"
    return float(value)


class NumFmt(Serialisable):
    tagname = "numFmt"

    formatCode: str | None = Field.attribute(expected_type=str, allow_none=True)
    sourceLinked: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=False)

    def __init__(self, formatCode=None, sourceLinked=False):
        self.formatCode = formatCode
        self.sourceLinked = sourceLinked


class NumVal(Serialisable):
    tagname = "pt"

    idx: int | None = Field.attribute(expected_type=int, allow_none=True)
    formatCode: str | None = Field.nested_text(expected_type=str, allow_none=True)
    v: float | str | None = Field.nested_text(
        expected_type=object,
        allow_none=True,
        converter=_convert_number_value,
    )

    def __init__(
        self,
        idx=None,
        formatCode=None,
        v=None,
    ):
        self.idx = idx
        self.formatCode = formatCode
        self.v = v


class NumData(Serialisable):
    tagname = "numCache"

    formatCode: str | None = Field.nested_text(expected_type=str, allow_none=True)
    ptCount: int | None = Field.nested_value(expected_type=int, allow_none=True)
    pt: list[NumVal] | None = Field.sequence(expected_type=NumVal, allow_none=True)
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = ("formatCode", "ptCount", "pt")

    def __init__(
        self,
        formatCode=None,
        ptCount=None,
        pt=(),
        extLst=None,
    ):
        self.formatCode = formatCode
        self.ptCount = ptCount
        self.pt = list(pt) if pt is not None else []
        self.extLst = extLst


class NumRef(Serialisable):
    tagname = "numRef"

    f: str | None = Field.nested_text(expected_type=str, allow_none=True)
    ref = AliasField("f")
    numCache: NumData | None = Field.element(expected_type=NumData, allow_none=True)
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = ("f", "numCache")

    def __init__(
        self,
        f=None,
        numCache=None,
        extLst=None,
    ):
        self.f = f
        self.numCache = numCache
        self.extLst = extLst


class StrVal(Serialisable):
    tagname = "strVal"

    idx: int | None = Field.attribute(expected_type=int, allow_none=True, default=0)
    v: str | None = Field.nested_text(expected_type=str, allow_none=True)

    def __init__(self, idx=0, v=None):
        self.idx = idx
        self.v = v


class StrData(Serialisable):
    tagname = "strData"

    ptCount: int | None = Field.nested_value(expected_type=int, allow_none=True)
    pt: list[StrVal] | None = Field.sequence(expected_type=StrVal, allow_none=True)
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = ("ptCount", "pt")

    def __init__(
        self,
        ptCount=None,
        pt=(),
        extLst=None,
    ):
        self.ptCount = ptCount
        self.pt = list(pt) if pt is not None else []
        self.extLst = extLst


class StrRef(Serialisable):
    tagname = "strRef"

    f: str | None = Field.nested_text(expected_type=str, allow_none=True)
    strCache: StrData | None = Field.element(expected_type=StrData, allow_none=True)
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = ("f", "strCache")

    def __init__(
        self,
        f=None,
        strCache=None,
        extLst=None,
    ):
        self.f = f
        self.strCache = strCache
        self.extLst = extLst


class NumDataSource(Serialisable):
    tagname = "val"

    numRef: NumRef | None = Field.element(expected_type=NumRef, allow_none=True)
    numLit: NumData | None = Field.element(expected_type=NumData, allow_none=True)

    xml_order = ("numRef", "numLit")

    def __init__(
        self,
        numRef=None,
        numLit=None,
    ):
        self.numRef = numRef
        self.numLit = numLit


class Level(Serialisable):
    tagname = "lvl"

    pt: list[StrVal] | None = Field.sequence(expected_type=StrVal, allow_none=True)

    def __init__(self, pt=()):
        self.pt = list(pt) if pt is not None else []


class MultiLevelStrData(Serialisable):
    tagname = "multiLvlStrData"

    ptCount: int | None = Field.attribute(
        expected_type=int, allow_none=True, xml_name="ptCount"
    )
    lvl: list[Level] | None = Field.sequence(expected_type=Level, allow_none=True)
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = ("ptCount", "lvl")

    def __init__(
        self,
        ptCount=None,
        lvl=(),
        extLst=None,
    ):
        self.ptCount = ptCount
        self.lvl = list(lvl) if lvl is not None else []
        self.extLst = extLst


class MultiLevelStrRef(Serialisable):
    tagname = "multiLvlStrRef"

    f: str | None = Field.nested_text(expected_type=str, allow_none=True)
    multiLvlStrCache: MultiLevelStrData | None = Field.element(
        expected_type=MultiLevelStrData, allow_none=True
    )
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = ("multiLvlStrCache", "f")

    def __init__(
        self,
        f=None,
        multiLvlStrCache=None,
        extLst=None,
    ):
        self.f = f
        self.multiLvlStrCache = multiLvlStrCache
        self.extLst = extLst


class AxDataSource(Serialisable):
    tagname = "cat"

    numRef: NumRef | None = Field.element(expected_type=NumRef, allow_none=True)
    numLit: NumData | None = Field.element(expected_type=NumData, allow_none=True)
    strRef: StrRef | None = Field.element(expected_type=StrRef, allow_none=True)
    strLit: StrData | None = Field.element(expected_type=StrData, allow_none=True)
    multiLvlStrRef: MultiLevelStrRef | None = Field.element(
        expected_type=MultiLevelStrRef, allow_none=True
    )

    xml_order = ("multiLvlStrRef", "numLit", "numRef", "strLit", "strRef")

    def __init__(
        self,
        numRef=None,
        numLit=None,
        strRef=None,
        strLit=None,
        multiLvlStrRef=None,
    ):
        if not any([numLit, numRef, strRef, strLit, multiLvlStrRef]):
            raise TypeError("A data source must be provided")
        self.numRef = numRef
        self.numLit = numLit
        self.strRef = strRef
        self.strLit = strLit
        self.multiLvlStrRef = multiLvlStrRef
