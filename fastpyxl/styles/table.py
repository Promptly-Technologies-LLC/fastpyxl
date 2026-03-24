# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.compat import safe_string
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import Field



class TableStyleElement(Serialisable):

    tagname = "tableStyleElement"

    type: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, _TABLE_STYLE_ELEMENT_TYPES, "type"), default=None,
    )
    size: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    dxfId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

    def __init__(self,
                 type=None,
                 size=None,
                 dxfId=None,
                ):
        self.type = type
        self.size = size
        self.dxfId = dxfId


class TableStyle(Serialisable):

    tagname = "tableStyle"

    name: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    pivot: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    table: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    count: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    tableStyleElement: list[TableStyleElement] | None = Field.sequence(
        expected_type=TableStyleElement,
        allow_none=True, default=list,
    )

    def __init__(self,
                 name=None,
                 pivot=None,
                 table=None,
                 count=None,
                 tableStyleElement=(),
                ):
        self.name = name
        self.pivot = pivot
        self.table = table
        self.count = count
        self.tableStyleElement = list(tableStyleElement)


class TableStyleList(Serialisable):

    tagname = "tableStyles"

    defaultTableStyle: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    defaultPivotStyle: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    tableStyle: list[TableStyle] | None = Field.sequence(expected_type=TableStyle, allow_none=True, default=list)

    def __init__(self,
                 count=None,
                 defaultTableStyle="TableStyleMedium9",
                 defaultPivotStyle="PivotStyleLight16",
                 tableStyle=(),
                ):
        self.defaultTableStyle = defaultTableStyle
        self.defaultPivotStyle = defaultPivotStyle
        self.tableStyle = list(tableStyle)


    @property
    def count(self):
        ts = self.tableStyle
        return len(ts) if ts is not None else 0


    def __iter__(self):
        yield "count", safe_string(self.count)
        if self.defaultTableStyle is not None:
            yield "defaultTableStyle", safe_string(self.defaultTableStyle)
        if self.defaultPivotStyle is not None:
            yield "defaultPivotStyle", safe_string(self.defaultPivotStyle)


_TABLE_STYLE_ELEMENT_TYPES = (
    "wholeTable",
    "headerRow",
    "totalRow",
    "firstColumn",
    "lastColumn",
    "firstRowStripe",
    "secondRowStripe",
    "firstColumnStripe",
    "secondColumnStripe",
    "firstHeaderCell",
    "lastHeaderCell",
    "firstTotalCell",
    "lastTotalCell",
    "firstSubtotalColumn",
    "secondSubtotalColumn",
    "thirdSubtotalColumn",
    "firstSubtotalRow",
    "secondSubtotalRow",
    "thirdSubtotalRow",
    "blankRow",
    "firstColumnSubheading",
    "secondColumnSubheading",
    "thirdColumnSubheading",
    "firstRowSubheading",
    "secondRowSubheading",
    "thirdRowSubheading",
    "pageFieldLabels",
    "pageFieldValues",
)


def _enum_converter(value, allowed_values, field_name: str):
    if value is None:
        return None
    if value not in allowed_values:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return value
