# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.compat import safe_string
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field


horizontal_alignments = (
    "general", "left", "center", "right", "fill", "justify", "centerContinuous",
    "distributed", )
vertical_aligments = (
    "top", "center", "bottom", "justify", "distributed",
)


def _enum_converter(value, allowed_values, field_name: str):
    if value is None:
        return None
    if value not in allowed_values:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return value


def _text_rotation_converter(value):
    if value is None:
        return None
    try:
        numeric = int(value)
    except Exception as exc:  # pragma: no cover
        raise FieldValidationError(f"textRotation rejected value {value!r}") from exc
    if numeric not in range(181) and numeric != 255:
        raise FieldValidationError(f"textRotation rejected value {value!r}")
    return numeric


def _range_converter(value, *, field_name: str, min_value: float, max_value: float | None):
    if value is None:
        return None
    try:
        numeric = float(value)
    except Exception as exc:  # pragma: no cover
        raise FieldValidationError(f"{field_name} rejected value {value!r}") from exc
    if numeric < min_value:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    if max_value is not None and numeric > max_value:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return numeric


class Alignment(Serialisable):
    """Alignment options for use in styles."""

    tagname = "alignment"

    horizontal: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, horizontal_alignments, "horizontal"),
    )
    vertical: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, vertical_aligments, "vertical"),
    )
    textRotation: int | None = Field.attribute(
        expected_type=int,
        allow_none=True,
        converter=_text_rotation_converter,
    )
    text_rotation: int | None = AliasField("textRotation")
    wrapText: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    wrap_text: bool | None = AliasField("wrapText")
    shrinkToFit: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    shrink_to_fit: bool | None = AliasField("shrinkToFit")
    indent: float | int | None = Field.attribute(
        expected_type=float,
        allow_none=True,
        converter=lambda v: _range_converter(v, field_name="indent", min_value=0, max_value=255),
    )
    relativeIndent: float | int | None = Field.attribute(
        expected_type=float,
        allow_none=True,
        converter=lambda v: _range_converter(v, field_name="relativeIndent", min_value=-255, max_value=255),
    )
    justifyLastLine: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    readingOrder: float | int | None = Field.attribute(
        expected_type=float,
        allow_none=True,
        converter=lambda v: _range_converter(v, field_name="readingOrder", min_value=0, max_value=None),
    )

    def __init__(self, horizontal=None, vertical=None,
                 textRotation=0, wrapText=None, shrinkToFit=None, indent=0, relativeIndent=0,
                 justifyLastLine=None, readingOrder=0, text_rotation=None,
                 wrap_text=None, shrink_to_fit=None, mergeCell=None):
        self.horizontal = horizontal
        self.vertical = vertical
        self.indent = indent
        self.relativeIndent = relativeIndent
        self.justifyLastLine = justifyLastLine
        self.readingOrder = readingOrder
        if text_rotation is not None:
            textRotation = text_rotation
        if textRotation is not None:
            self.textRotation = int(textRotation)
        if wrap_text is not None:
            wrapText = wrap_text
        self.wrapText = wrapText
        if shrink_to_fit is not None:
            shrinkToFit = shrink_to_fit
        self.shrinkToFit = shrinkToFit
        # mergeCell is vestigial


    def __iter__(self):
        explicit_bool_attrs = ("wrapText", "shrinkToFit", "justifyLastLine")
        for attr in self.__attrs__:
            value = getattr(self, attr)
            if value is None:
                continue
            if attr in explicit_bool_attrs:
                yield attr, safe_string(value)
            elif value != 0:
                yield attr, safe_string(value)
