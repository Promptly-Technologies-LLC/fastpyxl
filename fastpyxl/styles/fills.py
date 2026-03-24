
# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.compat import safe_string
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field

from .colors import Color

from fastpyxl.xml.functions import Element, localname


FILL_NONE = 'none'
FILL_SOLID = 'solid'
FILL_PATTERN_DARKDOWN = 'darkDown'
FILL_PATTERN_DARKGRAY = 'darkGray'
FILL_PATTERN_DARKGRID = 'darkGrid'
FILL_PATTERN_DARKHORIZONTAL = 'darkHorizontal'
FILL_PATTERN_DARKTRELLIS = 'darkTrellis'
FILL_PATTERN_DARKUP = 'darkUp'
FILL_PATTERN_DARKVERTICAL = 'darkVertical'
FILL_PATTERN_GRAY0625 = 'gray0625'
FILL_PATTERN_GRAY125 = 'gray125'
FILL_PATTERN_LIGHTDOWN = 'lightDown'
FILL_PATTERN_LIGHTGRAY = 'lightGray'
FILL_PATTERN_LIGHTGRID = 'lightGrid'
FILL_PATTERN_LIGHTHORIZONTAL = 'lightHorizontal'
FILL_PATTERN_LIGHTTRELLIS = 'lightTrellis'
FILL_PATTERN_LIGHTUP = 'lightUp'
FILL_PATTERN_LIGHTVERTICAL = 'lightVertical'
FILL_PATTERN_MEDIUMGRAY = 'mediumGray'

fills = (
    FILL_NONE,
    FILL_SOLID,
    FILL_PATTERN_DARKDOWN,
    FILL_PATTERN_DARKGRAY,
    FILL_PATTERN_DARKGRID,
    FILL_PATTERN_DARKHORIZONTAL,
    FILL_PATTERN_DARKTRELLIS,
    FILL_PATTERN_DARKUP,
    FILL_PATTERN_DARKVERTICAL,
    FILL_PATTERN_GRAY0625,
    FILL_PATTERN_GRAY125,
    FILL_PATTERN_LIGHTDOWN,
    FILL_PATTERN_LIGHTGRAY,
    FILL_PATTERN_LIGHTGRID,
    FILL_PATTERN_LIGHTHORIZONTAL,
    FILL_PATTERN_LIGHTTRELLIS,
    FILL_PATTERN_LIGHTUP,
    FILL_PATTERN_LIGHTVERTICAL,
    FILL_PATTERN_MEDIUMGRAY,
)


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


def _range_converter(value, *, field_name: str, min_value: float, max_value: float):
    if value is None:
        return None
    try:
        numeric = float(value)
    except Exception as exc:  # pragma: no cover
        raise FieldValidationError(f"{field_name} rejected value {value!r}") from exc
    if numeric < min_value or numeric > max_value:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return numeric


class Fill(Serialisable):

    """Base class"""

    tagname = "fill"

    @classmethod
    def from_tree(cls, el):
        children = [c for c in el]
        if not children:
            return
        child = children[0]
        if "patternFill" in child.tag:
            return PatternFill._from_tree(child)
        return super(Fill, GradientFill).from_tree(child)


class PatternFill(Fill):
    """Area fill patterns for use in styles.
    Caution: if you do not specify a fill_type, other attributes will have
    no effect !"""

    tagname = "patternFill"

    patternType: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: None if v == "none" else _enum_converter(v, fills, "patternType"), default=None,
    )
    fill_type: str | None = AliasField("patternType", default=None)
    fgColor: Color | None = Field.element(
        expected_type=Color,
        allow_none=True,
        converter=lambda v: _color_converter(v, "fgColor"), default=None,
    )
    start_color: Color | None = AliasField("fgColor", default=None)
    bgColor: Color | None = Field.element(
        expected_type=Color,
        allow_none=True,
        converter=lambda v: _color_converter(v, "bgColor"), default=None,
    )
    end_color: Color | None = AliasField("bgColor", default=None)

    def __init__(self, patternType=None, fgColor=Color(), bgColor=Color(),
                 fill_type=None, start_color=None, end_color=None):
        if fill_type is not None:
            patternType = fill_type
        self.patternType = patternType
        if start_color is not None:
            fgColor = start_color
        self.fgColor = fgColor
        if end_color is not None:
            bgColor = end_color
        self.bgColor = bgColor

    @classmethod
    def _from_tree(cls, el):
        attrib = dict(el.attrib)
        for child in el:
            desc = localname(child)
            attrib[desc] = Color.from_tree(child)
        return cls(**attrib)


    def to_tree(self, tagname=None, idx=None):
        parent = Element("fill")
        el = Element(self.tagname)
        if self.patternType is not None:
            el.set('patternType', self.patternType)
        for c in self.__elements__:
            value = getattr(self, c)
            if value != Color():
                el.append(value.to_tree(c))
        parent.append(el)
        return parent


DEFAULT_EMPTY_FILL = PatternFill()
DEFAULT_GRAY_FILL = PatternFill(patternType='gray125')


class Stop(Serialisable):

    tagname = "stop"

    position: float | None = Field.attribute(
        expected_type=float,
        allow_none=True,
        converter=lambda v: _range_converter(v, field_name="position", min_value=0, max_value=1), default=None,
    )
    color: Color | None = Field.element(
        expected_type=Color,
        allow_none=True,
        converter=lambda v: _color_converter(v, "color"), default=None,
    )

    def __init__(self, color, position):
        if position is None:
            raise TypeError("position cannot be None")
        self.position = position
        self.color = color


def _assign_position(values):
    """
    Automatically assign positions if a list of colours is provided.

    It is not permitted to mix colours and stops
    """
    n_values = len(values)
    n_stops = sum(isinstance(value, Stop) for value in values)

    if n_stops == 0:
        interval = 1
        if n_values > 2:
            interval = 1 / (n_values - 1)
        values = [Stop(value, i * interval)
                  for i, value in enumerate(values)]

    elif n_stops < n_values:
        raise ValueError('Cannot interpret mix of Stops and Colors in GradientFill')

    pos = set()
    for stop in values:
        if stop.position in pos:
            raise ValueError("Duplicate position {0}".format(stop.position))
        pos.add(stop.position)

    return values


class GradientFill(Fill):
    """Fill areas with gradient

    Two types of gradient fill are supported:

        - A type='linear' gradient interpolates colours between
          a set of specified Stops, across the length of an area.
          The gradient is left-to-right by default, but this
          orientation can be modified with the degree
          attribute.  A list of Colors can be provided instead
          and they will be positioned with equal distance between them.

        - A type='path' gradient applies a linear gradient from each
          edge of the area. Attributes top, right, bottom, left specify
          the extent of fill from the respective borders. Thus top="0.2"
          will fill the top 20% of the cell.

    """

    tagname = "gradientFill"

    type: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("linear", "path"), "type"), default=None,
    )
    fill_type: str | None = AliasField("type", default=None)
    degree: float | None = Field.attribute(expected_type=float, allow_none=True, default=None)
    left: float | None = Field.attribute(expected_type=float, allow_none=True, default=None)
    right: float | None = Field.attribute(expected_type=float, allow_none=True, default=None)
    top: float | None = Field.attribute(expected_type=float, allow_none=True, default=None)
    bottom: float | None = Field.attribute(expected_type=float, allow_none=True, default=None)
    stop: list[Stop] = Field.sequence(expected_type=Stop, default=list)


    def __init__(self, type="linear", degree=0, left=0, right=0, top=0,
                 bottom=0, stop=()):
        self.degree = degree
        self.left = left
        self.right = right
        self.top = top
        self.bottom = bottom
        self.stop = list(stop)
        self.type = type


    def __setattr__(self, name, value):
        if name == "stop" and value is not None:
            value = _assign_position(value)
        super().__setattr__(name, value)


    def __iter__(self):
        for attr in self.__attrs__:
            value = getattr(self, attr)
            if value:
                yield attr, safe_string(value)


    def to_tree(self, tagname=None, namespace=None, idx=None):
        parent = Element("fill")
        el = super().to_tree()
        parent.append(el)
        return parent


