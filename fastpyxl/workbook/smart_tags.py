# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import Field

class SmartTag(Serialisable):

    tagname = "smartTagType"

    namespaceUri: str | None = Field.attribute(expected_type=str, allow_none=True)
    name: str | None = Field.attribute(expected_type=str, allow_none=True)
    url: str | None = Field.attribute(expected_type=str, allow_none=True)

    def __init__(self,
                 namespaceUri=None,
                 name=None,
                 url=None,
                ):
        self.namespaceUri = namespaceUri
        self.name = name
        self.url = url


class SmartTagList(Serialisable):

    tagname = "smartTagTypes"

    smartTagType: list[SmartTag] | None = Field.sequence(expected_type=SmartTag, allow_none=True)
    xml_order = ('smartTagType',)

    def __init__(self,
                 smartTagType=(),
                ):
        self.smartTagType = smartTagType


class SmartTagProperties(Serialisable):

    tagname = "smartTagPr"

    embed: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    show: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("all", "noIndicator"), "show"),
    )

    def __init__(self,
                 embed=None,
                 show=None,
                ):
        self.embed = embed
        self.show = show


def _enum_converter(value, allowed_values, field_name: str):
    if value is None:
        return None
    if value not in allowed_values:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return value
