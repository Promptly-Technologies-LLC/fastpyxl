# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.xml.constants import DRAWING_NS
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field

from .colors import ColorChoice
from .fill import GradientFillProperties, PatternFillProperties
from fastpyxl.descriptors.excel import ExtensionList as OfficeArtExtensionList

"""
Line elements from drawing main schema
"""


class LineEndProperties(Serialisable):

    tagname = "end"
    namespace = DRAWING_NS

    type: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("none", "triangle", "stealth", "diamond", "oval", "arrow"), "type"),
    )
    w: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("sm", "med", "lg"), "w"),
    )
    len: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("sm", "med", "lg"), "len"),
    )

    def __init__(self,
                 type=None,
                 w=None,
                 len=None,
                ):
        self.type = type
        self.w = w
        self.len = len


class DashStop(Serialisable):

    tagname = "ds"
    namespace = DRAWING_NS

    d: int | None = Field.attribute(expected_type=int, allow_none=True)
    length = AliasField('d')
    sp: int | None = Field.attribute(expected_type=int, allow_none=True)
    space = AliasField('sp')

    def __init__(self,
                 d=0,
                 sp=0,
                ):
        self.d = d
        self.sp = sp


class DashStopList(Serialisable):

    ds: list[DashStop] | None = Field.sequence(expected_type=DashStop, allow_none=True)

    def __init__(self,
                 ds=None,
                ):
        self.ds = ds


class _NoFill(Serialisable):
    tagname = "noFill"
    namespace = DRAWING_NS


class _Round(Serialisable):
    tagname = "round"
    namespace = DRAWING_NS


class _Bevel(Serialisable):
    tagname = "bevel"
    namespace = DRAWING_NS


class _Miter(Serialisable):
    tagname = "miter"
    namespace = DRAWING_NS
    lim: int | None = Field.attribute(expected_type=int, allow_none=True)

    def __init__(self, lim=None):
        self.lim = lim


def _line_no_fill_coerce(value):
    if value is True:
        return _NoFill()
    if value in (False, None):
        return None
    return value


def _solid_fill_coerce(value):
    if value is None:
        return None
    if isinstance(value, ColorChoice):
        return value
    if isinstance(value, str):
        return ColorChoice(srgbClr=value)
    return value


def _enum_converter(value, allowed_values, field_name: str):
    if value is None:
        return None
    if value not in allowed_values:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return value


class LineProperties(Serialisable):

    tagname = "ln"
    namespace = DRAWING_NS

    w: int | None = Field.attribute(
        expected_type=int,
        allow_none=True,
        converter=lambda v: _range_converter(v, 0, 20116800, "w"),
    )  # EMU
    width = AliasField('w')
    cap: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("rnd", "sq", "flat"), "cap"),
    )
    cmpd: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("sng", "dbl", "thickThin", "thinThick", "tri"), "cmpd"),
    )
    algn: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("ctr", "in"), "algn"),
    )

    noFill: _NoFill | None = Field.element(
        expected_type=_NoFill,
        allow_none=True,
        converter=_line_no_fill_coerce,
    )
    solidFill: ColorChoice | None = Field.element(
        expected_type=ColorChoice, allow_none=True, converter=_solid_fill_coerce
    )
    gradFill: GradientFillProperties | None = Field.element(expected_type=GradientFillProperties, allow_none=True)
    pattFill: PatternFillProperties | None = Field.element(expected_type=PatternFillProperties, allow_none=True)

    prstDash: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(
            v,
            ("solid", "dot", "dash", "lgDash", "dashDot", "lgDashDot", "lgDashDotDot", "sysDash", "sysDot", "sysDashDot", "sysDashDotDot"),
            "prstDash",
        ),
    )
    dashStyle = AliasField('prstDash')

    custDash: DashStopList | None = Field.element(expected_type=DashStopList, allow_none=True)

    round: _Round | None = Field.element(expected_type=_Round, allow_none=True)
    bevel: _Bevel | None = Field.element(expected_type=_Bevel, allow_none=True)
    miter: _Miter | None = Field.element(expected_type=_Miter, allow_none=True)

    headEnd: LineEndProperties | None = Field.element(expected_type=LineEndProperties, allow_none=True)
    tailEnd: LineEndProperties | None = Field.element(expected_type=LineEndProperties, allow_none=True)
    extLst: OfficeArtExtensionList | None = Field.element(expected_type=OfficeArtExtensionList, allow_none=True)

    xml_order = ('noFill', 'solidFill', 'gradFill', 'pattFill', 'prstDash', 'custDash', 'round', 'bevel', 'miter', 'headEnd', 'tailEnd')

    def __init__(self,
                 w=None,
                 cap=None,
                 cmpd=None,
                 algn=None,
                 noFill=None,
                 solidFill=None,
                 gradFill=None,
                 pattFill=None,
                 prstDash=None,
                 custDash=None,
                 round=None,
                 bevel=None,
                 miter=None,
                 headEnd=None,
                 tailEnd=None,
                 extLst=None,
                ):
        self.w = w
        self.cap = cap
        self.cmpd = cmpd
        self.algn = algn
        self.noFill = _NoFill() if noFill else None
        self.solidFill = solidFill
        self.gradFill = gradFill
        self.pattFill = pattFill
        if prstDash is None:
            prstDash = "solid"
        self.prstDash = prstDash
        self.custDash = DashStopList(ds=[custDash]) if isinstance(custDash, DashStop) else custDash
        self.round = _Round() if round else None
        self.bevel = _Bevel() if bevel else None
        self.miter = _Miter(lim=miter) if isinstance(miter, int) else miter
        self.headEnd = headEnd
        self.tailEnd = tailEnd
        self.extLst = extLst


def _range_converter(value, minimum: int, maximum: int, field_name: str):
    if value is None:
        return None
    try:
        numeric = int(value)
    except (TypeError, ValueError) as exc:
        raise FieldValidationError(f"{field_name} rejected value {value!r}") from exc
    if not minimum <= numeric <= maximum:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return numeric
