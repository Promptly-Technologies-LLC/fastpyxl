# Copyright (c) 2010-2024 fastpyxl


from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field
from .colors import Color

from fastpyxl.compat import safe_string
from fastpyxl.xml.functions import Element
from fastpyxl.xml.constants import SHEET_MAIN_NS


def _no_value(tagname, value, namespace=None):
    if value:
        return Element(tagname, val=safe_string(value))


def _validate_range(value, *, field_name: str, min_value: float, max_value: float):
    if value is None:
        return None
    try:
        numeric = float(value)
    except Exception as exc:  # pragma: no cover
        raise FieldValidationError(
            f"{field_name} rejected value {value!r}"
        ) from exc
    if numeric < min_value or numeric > max_value:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return numeric


def _color_converter(value):
    if value is None:
        return None
    if isinstance(value, Color):
        return value
    if isinstance(value, str):
        return Color(rgb=value)
    raise FieldValidationError(f"color rejected value {value!r}")


def _underline_converter(value):
    if value is None or value == "none":
        return None
    allowed = {
        "single",
        "double",
        "singleAccounting",
        "doubleAccounting",
    }
    if value not in allowed:
        raise FieldValidationError(f"u rejected value {value!r}")
    return value


def _vert_align_converter(value):
    if value is None or value == "none":
        return None
    allowed = {"superscript", "subscript", "baseline"}
    if value not in allowed:
        raise FieldValidationError(f"vertAlign rejected value {value!r}")
    return value


def _scheme_converter(value):
    if value is None or value == "none":
        return None
    allowed = {"major", "minor"}
    if value not in allowed:
        raise FieldValidationError(f"scheme rejected value {value!r}")
    return value


class Font(Serialisable):
    """Font options used in styles."""

    UNDERLINE_DOUBLE = 'double'
    UNDERLINE_DOUBLE_ACCOUNTING = 'doubleAccounting'
    UNDERLINE_SINGLE = 'single'
    UNDERLINE_SINGLE_ACCOUNTING = 'singleAccounting'

    tagname = "font"

    xml_order = (
        "name",
        "charset",
        "family",
        "b",
        "i",
        "strike",
        "outline",
        "shadow",
        "condense",
        "color",
        "extend",
        "sz",
        "u",
        "vertAlign",
        "scheme",
    )

    name: str | None = Field.nested_value(expected_type=str, allow_none=True, default=None)
    charset: int | None = Field.nested_value(expected_type=int, allow_none=True, default=None)
    family: float | None = Field.nested_value(
        expected_type=float,
        allow_none=True,
        converter=lambda v: _validate_range(v, field_name="family", min_value=0, max_value=14), default=None,
    )
    sz: float | None = Field.nested_value(expected_type=float, allow_none=True, default=None)
    size: float | None = AliasField("sz", default=None)

    b: bool | None = Field.nested_bool(renderer=_no_value, default=None)
    bold: bool | None = AliasField("b", default=None)
    i: bool | None = Field.nested_bool(renderer=_no_value, default=None)
    italic: bool | None = AliasField("i", default=None)

    strike: bool | None = Field.nested_bool(allow_none=True, default=None)
    strikethrough: bool | None = AliasField("strike", default=None)
    outline: bool | None = Field.nested_bool(allow_none=True, default=None)
    shadow: bool | None = Field.nested_bool(allow_none=True, default=None)
    condense: bool | None = Field.nested_bool(allow_none=True, default=None)

    color: Color | None = Field.element(
        expected_type=Color,
        allow_none=True,
        converter=_color_converter, default=None,
    )
    extend: bool | None = Field.nested_bool(allow_none=True, default=None)

    u: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_underline_converter, default=None,
    )
    underline: str | None = AliasField("u", default=None)

    vertAlign: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_vert_align_converter, default=None,
    )
    scheme: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_scheme_converter, default=None,
    )


    def __init__(self, name=None, sz=None, b=None, i=None, charset=None,
                 u=None, strike=None, color=None, scheme=None, family=None, size=None,
                 bold=None, italic=None, strikethrough=None, underline=None,
                 vertAlign=None, outline=None, shadow=None, condense=None,
                 extend=None):
        self.name = name
        self.family = family
        if size is not None:
            sz = size
        self.sz = sz
        if bold is not None:
            b = bold
        self.b = b
        if italic is not None:
            i = italic
        self.i = i
        if underline is not None:
            u = underline
        self.u = u
        if strikethrough is not None:
            strike = strikethrough
        self.strike = strike
        self.color = color
        self.vertAlign = vertAlign
        self.charset = charset
        self.outline = outline
        self.shadow = shadow
        self.condense = condense
        self.extend = extend
        self.scheme = scheme


    @classmethod
    def from_tree(cls, node):
        """
        Set default value for underline if child element is present
        """
        underline = node.find("{%s}u" % SHEET_MAIN_NS)
        if underline is not None and underline.get('val') is None:
            underline.set("val", "single")
        return super().from_tree(node)

    def __copy__(self):
        cp = super().__copy__()
        # `b` and `i` use a custom renderer that omits the element when falsey.
        # During copy via XML round-trip that means `False` becomes `None` unless
        # we explicitly preserve the semantic value here.
        if self.b is False and cp.b is None:
            cp.b = False
        if self.i is False and cp.i is None:
            cp.i = False
        return cp


DEFAULT_FONT = Font(name="Calibri", sz=11, family=2, b=False, i=False,
                    color=Color(theme=1), scheme="minor")
