# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import Field


def _none_or_enum(allowed: frozenset, field_name: str):
    def _c(value):
        if value is None:
            return None
        if value == "none":
            return None
        if value not in allowed:
            raise FieldValidationError(f"{field_name} rejected value {value!r}")
        return value

    return _c


class PictureOptions(Serialisable):
    tagname = "pictureOptions"

    applyToFront: bool | None = Field.nested_bool(allow_none=True, default=None)
    applyToSides: bool | None = Field.nested_bool(allow_none=True, default=None)
    applyToEnd: bool | None = Field.nested_bool(allow_none=True, default=None)
    pictureFormat: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_none_or_enum(
            frozenset({"stretch", "stack", "stackScale"}),
            "pictureFormat",
        ), default=None,
    )
    pictureStackUnit: float | None = Field.nested_value(
        expected_type=float, allow_none=True, default=None
    )

    xml_order = (
        "applyToFront",
        "applyToSides",
        "applyToEnd",
        "pictureFormat",
        "pictureStackUnit",
    )

    def __init__(
        self,
        applyToFront=None,
        applyToSides=None,
        applyToEnd=None,
        pictureFormat=None,
        pictureStackUnit=None,
    ):
        self.applyToFront = applyToFront
        self.applyToSides = applyToSides
        self.applyToEnd = applyToEnd
        self.pictureFormat = pictureFormat
        self.pictureStackUnit = pictureStackUnit
