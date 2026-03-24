# Copyright (c) 2010-2024 fastpyxl

"""
Richtext definition
"""

from fastpyxl.xml.constants import SHEET_MAIN_NS

from fastpyxl.styles.colors import Color
from fastpyxl.styles.fonts import _no_value
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field


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


def _family_converter(value):
    if value is None:
        return None
    try:
        numeric = float(value)
    except Exception as exc:  # pragma: no cover
        raise FieldValidationError(f"family rejected value {value!r}") from exc
    if numeric < 0 or numeric > 14:
        raise FieldValidationError(f"family rejected value {value!r}")
    return numeric


_PHONETIC_TYPES = frozenset(
    {
        "halfwidthKatakana",
        "fullwidthKatakana",
        "Hiragana",
        "noConversion",
    }
)


def _phonetic_type_converter(value):
    if value is None:
        return None
    if value not in _PHONETIC_TYPES:
        raise FieldValidationError(f"type rejected value {value!r}")
    return value


_PHONETIC_ALIGN = frozenset(
    {"noControl", "left", "center", "distributed"}
)


def _phonetic_alignment_converter(value):
    if value is None:
        return None
    if value not in _PHONETIC_ALIGN:
        raise FieldValidationError(f"alignment rejected value {value!r}")
    return value


class PhoneticProperties(Serialisable):

    tagname = "phoneticPr"

    fontId: int | None = Field.attribute(expected_type=int, allow_none=True)
    type: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=_phonetic_type_converter,
    )
    alignment: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=_phonetic_alignment_converter,
    )

    def __init__(
        self,
        fontId=None,
        type=None,
        alignment=None,
    ):
        self.fontId = fontId
        self.type = type
        self.alignment = alignment


class PhoneticText(Serialisable):

    tagname = "rPh"

    sb: int | None = Field.attribute(expected_type=int, allow_none=True)
    eb: int | None = Field.attribute(expected_type=int, allow_none=True)
    t: str | None = Field.nested_text(expected_type=str, allow_none=True)
    text: str | None = AliasField("t")

    def __init__(
        self,
        sb=None,
        eb=None,
        t=None,
    ):
        self.sb = sb
        self.eb = eb
        self.t = t


class InlineFont(Serialisable):
    """
    Font for inline text because, yes what you need are different objects with the same elements but different constraints.
    """

    tagname = "RPrElt"

    xml_order = (
        "rFont",
        "charset",
        "family",
        "b",
        "i",
        "strike",
        "outline",
        "shadow",
        "condense",
        "extend",
        "color",
        "sz",
        "u",
        "vertAlign",
        "scheme",
    )

    rFont: str | None = Field.nested_value(expected_type=str, allow_none=True)
    charset: int | None = Field.nested_value(expected_type=int, allow_none=True)
    family: float | None = Field.nested_value(
        expected_type=float,
        allow_none=True,
        converter=_family_converter,
    )
    b: bool | None = Field.nested_bool(renderer=_no_value)
    i: bool | None = Field.nested_bool(renderer=_no_value)
    strike: bool | None = Field.nested_bool(allow_none=True)
    outline: bool | None = Field.nested_bool(allow_none=True)
    shadow: bool | None = Field.nested_bool(allow_none=True)
    condense: bool | None = Field.nested_bool(allow_none=True)
    extend: bool | None = Field.nested_bool(allow_none=True)
    color: Color | None = Field.element(
        expected_type=Color,
        allow_none=True,
        converter=_color_converter,
    )
    sz: float | None = Field.nested_value(expected_type=float, allow_none=True)
    u: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_underline_converter,
    )
    vertAlign: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_vert_align_converter,
    )
    scheme: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_scheme_converter,
    )

    def __init__(
        self,
        rFont=None,
        charset=None,
        family=None,
        b=None,
        i=None,
        strike=None,
        outline=None,
        shadow=None,
        condense=None,
        extend=None,
        color=None,
        sz=None,
        u=None,
        vertAlign=None,
        scheme=None,
    ):
        self.rFont = rFont
        self.charset = charset
        self.family = family
        self.b = b
        self.i = i
        self.strike = strike
        self.outline = outline
        self.shadow = shadow
        self.condense = condense
        self.extend = extend
        self.color = color
        self.sz = sz
        self.u = u
        self.vertAlign = vertAlign
        self.scheme = scheme

    @classmethod
    def from_tree(cls, node):
        underline = node.find("{%s}u" % SHEET_MAIN_NS)
        if underline is not None and underline.get("val") is None:
            underline.set("val", "single")
        return super().from_tree(node)


class RichText(Serialisable):

    tagname = "RElt"

    rPr: InlineFont | None = Field.element(expected_type=InlineFont, allow_none=True)
    font: InlineFont | None = AliasField("rPr")
    t: str | None = Field.nested_text(expected_type=str, allow_none=True)
    text: str | None = AliasField("t")

    xml_order = ("rPr", "t")

    def __init__(
        self,
        rPr=None,
        t=None,
    ):
        self.rPr = rPr
        self.t = t


class Text(Serialisable):

    tagname = "text"

    xml_order = ("t", "r", "rPh", "phoneticPr")

    t: str | None = Field.nested_text(expected_type=str, allow_none=True)
    plain: str | None = AliasField("t")
    r: list[RichText] = Field.sequence(expected_type=RichText, default=list)
    formatted = AliasField("r")
    rPh: list[PhoneticText] = Field.sequence(expected_type=PhoneticText, default=list)
    phonetic = AliasField("rPh")
    phoneticPr: PhoneticProperties | None = Field.element(
        expected_type=PhoneticProperties, allow_none=True
    )
    PhoneticProperties = AliasField("phoneticPr")

    def __init__(
        self,
        t=None,
        r=(),
        rPh=(),
        phoneticPr=None,
    ):
        self.t = t
        self.r = list(r)
        self.rPh = list(rPh)
        self.phoneticPr = phoneticPr

    @classmethod
    def from_tree(cls, node):
        underline = node.find("{%s}u" % SHEET_MAIN_NS)
        if underline is not None and underline.get("val") is None:
            underline.set("val", "single")
        return super().from_tree(node)

    @property
    def content(self):
        """
        Text stripped of all formatting
        """
        snippets = []
        if self.plain is not None:
            snippets.append(self.plain)
        for block in self.formatted:
            if block.t is not None:
                snippets.append(block.t)
        return "".join(snippets)

