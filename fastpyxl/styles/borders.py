# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.compat import safe_string
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field

from .colors import Color


BORDER_NONE = None
BORDER_DASHDOT = 'dashDot'
BORDER_DASHDOTDOT = 'dashDotDot'
BORDER_DASHED = 'dashed'
BORDER_DOTTED = 'dotted'
BORDER_DOUBLE = 'double'
BORDER_HAIR = 'hair'
BORDER_MEDIUM = 'medium'
BORDER_MEDIUMDASHDOT = 'mediumDashDot'
BORDER_MEDIUMDASHDOTDOT = 'mediumDashDotDot'
BORDER_MEDIUMDASHED = 'mediumDashed'
BORDER_SLANTDASHDOT = 'slantDashDot'
BORDER_THICK = 'thick'
BORDER_THIN = 'thin'


class Side(Serialisable):

    """Border options for use in styles.
    Caution: if you do not specify a border_style, other attributes will
    have no effect !"""


    color: Color | None = Field.element(
        expected_type=Color,
        allow_none=True,
        converter=lambda v: _color_converter(v, "color"), default=None,
    )
    style: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(
            v,
            (
                "dashDot",
                "dashDotDot",
                "dashed",
                "dotted",
                "double",
                "hair",
                "medium",
                "mediumDashDot",
                "mediumDashDotDot",
                "mediumDashed",
                "slantDashDot",
                "thick",
                "thin",
            ),
            "style",
        ), default=None,
    )
    border_style: str | None = AliasField("style", default=None)

    def __init__(self, style=None, color=None, border_style=None):
        if border_style is not None:
            style = border_style
        self.style = style
        self.color = color


class Border(Serialisable):
    """Border positioning for use in styles."""

    tagname = "border"

    xml_order = (
        "start",
        "end",
        "left",
        "right",
        "top",
        "bottom",
        "diagonal",
        "vertical",
        "horizontal",
    )

    # child elements
    start: Side | None = Field.element(expected_type=Side, allow_none=True, default=None)
    end: Side | None = Field.element(expected_type=Side, allow_none=True, default=None)
    left: Side | None = Field.element(expected_type=Side, allow_none=True, default=None)
    right: Side | None = Field.element(expected_type=Side, allow_none=True, default=None)
    top: Side | None = Field.element(expected_type=Side, allow_none=True, default=None)
    bottom: Side | None = Field.element(expected_type=Side, allow_none=True, default=None)
    diagonal: Side | None = Field.element(expected_type=Side, allow_none=True, default=None)
    vertical: Side | None = Field.element(expected_type=Side, allow_none=True, default=None)
    horizontal: Side | None = Field.element(expected_type=Side, allow_none=True, default=None)
    # attributes
    outline: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    diagonalUp: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    diagonalDown: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)

    def __init__(self, left=None, right=None, top=None,
                 bottom=None, diagonal=None, diagonal_direction=None,
                 vertical=None, horizontal=None, diagonalUp=False, diagonalDown=False,
                 outline=True, start=None, end=None):
        self.left = left
        self.right = right
        self.top = top
        self.bottom = bottom
        self.diagonal = diagonal
        self.vertical = vertical
        self.horizontal = horizontal
        self.diagonal_direction = diagonal_direction
        self.diagonalUp = diagonalUp
        self.diagonalDown = diagonalDown
        self.outline = outline
        self.start = start
        self.end = end

    def __iter__(self):
        for attr in self.__attrs__:
            value = getattr(self, attr)
            if value and attr != "outline":
                yield attr, safe_string(value)
            elif attr == "outline" and not value:
                yield attr, safe_string(value)

DEFAULT_BORDER = Border(left=Side(), right=Side(), top=Side(), bottom=Side(), diagonal=Side())


def _color_converter(value, field_name: str):
    if value is None:
        return None
    if isinstance(value, Color):
        return value
    if isinstance(value, str):
        return Color(rgb=value)
    raise FieldValidationError(f"{field_name} rejected value {value!r}")


def _enum_converter(value, allowed_values, field_name: str):
    if value is None:
        return None
    if value not in allowed_values:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return value
