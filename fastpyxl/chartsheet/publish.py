# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import Field


class WebPublishItem(Serialisable):
    tagname = "webPublishItem"

    id: int | None = Field.attribute(expected_type=int, allow_none=True)
    divId: str | None = Field.attribute(expected_type=str, allow_none=True)
    sourceType: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(
            v,
            ("sheet", "printArea", "autoFilter", "range", "chart", "pivotTable", "query", "label"),
            "sourceType",
        ),
    )
    sourceRef: str | None = Field.attribute(expected_type=str, allow_none=True)
    sourceObject: str | None = Field.attribute(expected_type=str, allow_none=True)
    destinationFile: str | None = Field.attribute(expected_type=str, allow_none=True)
    title: str | None = Field.attribute(expected_type=str, allow_none=True)
    autoRepublish: bool | None = Field.attribute(expected_type=bool, allow_none=True)

    def __init__(self,
                 id=None,
                 divId=None,
                 sourceType=None,
                 sourceRef=None,
                 sourceObject=None,
                 destinationFile=None,
                 title=None,
                 autoRepublish=None,
                 ):
        self.id = id
        self.divId = divId
        self.sourceType = sourceType
        self.sourceRef = sourceRef
        self.sourceObject = sourceObject
        self.destinationFile = destinationFile
        self.title = title
        self.autoRepublish = autoRepublish


class WebPublishItems(Serialisable):
    tagname = "WebPublishItems"

    webPublishItem: list[WebPublishItem] = Field.sequence(expected_type=WebPublishItem)
    xml_order = ("webPublishItem",)

    def __init__(self,
                 count=None,
                 webPublishItem=None,
                 ):
        del count
        self.webPublishItem = webPublishItem


    @property
    def count(self):
        return len(self.webPublishItem)


    def __iter__(self):
        yield "count", str(self.count)


def _enum_converter(value, allowed_values, field_name: str):
    if value is None:
        return None
    if value not in allowed_values:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return value
