# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.descriptors.serialisable import Serialisable
from fastpyxl.descriptors import (
    Typed,
    Set,
)
from fastpyxl.styles.colors import aRGB_REGEX
from fastpyxl.xml.constants import DRAWING_NS
from fastpyxl.xml.functions import Element
from fastpyxl.typed_serialisable.render import namespaced_tag
from fastpyxl.typed_serialisable.base import Serialisable as TypedSerialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.descriptors.excel import ExtensionList as OfficeArtExtensionList

PRESET_COLORS = [
        'aliceBlue', 'antiqueWhite', 'aqua', 'aquamarine',
        'azure', 'beige', 'bisque', 'black', 'blanchedAlmond', 'blue',
        'blueViolet', 'brown', 'burlyWood', 'cadetBlue', 'chartreuse',
        'chocolate', 'coral', 'cornflowerBlue', 'cornsilk', 'crimson', 'cyan',
        'darkBlue', 'darkCyan', 'darkGoldenrod', 'darkGray', 'darkGrey',
        'darkGreen', 'darkKhaki', 'darkMagenta', 'darkOliveGreen', 'darkOrange',
        'darkOrchid', 'darkRed', 'darkSalmon', 'darkSeaGreen', 'darkSlateBlue',
        'darkSlateGray', 'darkSlateGrey', 'darkTurquoise', 'darkViolet',
        'dkBlue', 'dkCyan', 'dkGoldenrod', 'dkGray', 'dkGrey', 'dkGreen',
        'dkKhaki', 'dkMagenta', 'dkOliveGreen', 'dkOrange', 'dkOrchid', 'dkRed',
        'dkSalmon', 'dkSeaGreen', 'dkSlateBlue', 'dkSlateGray', 'dkSlateGrey',
        'dkTurquoise', 'dkViolet', 'deepPink', 'deepSkyBlue', 'dimGray',
        'dimGrey', 'dodgerBlue', 'firebrick', 'floralWhite', 'forestGreen',
        'fuchsia', 'gainsboro', 'ghostWhite', 'gold', 'goldenrod', 'gray',
        'grey', 'green', 'greenYellow', 'honeydew', 'hotPink', 'indianRed',
        'indigo', 'ivory', 'khaki', 'lavender', 'lavenderBlush', 'lawnGreen',
        'lemonChiffon', 'lightBlue', 'lightCoral', 'lightCyan',
        'lightGoldenrodYellow', 'lightGray', 'lightGrey', 'lightGreen',
        'lightPink', 'lightSalmon', 'lightSeaGreen', 'lightSkyBlue',
        'lightSlateGray', 'lightSlateGrey', 'lightSteelBlue', 'lightYellow',
        'ltBlue', 'ltCoral', 'ltCyan', 'ltGoldenrodYellow', 'ltGray', 'ltGrey',
        'ltGreen', 'ltPink', 'ltSalmon', 'ltSeaGreen', 'ltSkyBlue',
        'ltSlateGray', 'ltSlateGrey', 'ltSteelBlue', 'ltYellow', 'lime',
        'limeGreen', 'linen', 'magenta', 'maroon', 'medAquamarine', 'medBlue',
        'medOrchid', 'medPurple', 'medSeaGreen', 'medSlateBlue',
        'medSpringGreen', 'medTurquoise', 'medVioletRed', 'mediumAquamarine',
        'mediumBlue', 'mediumOrchid', 'mediumPurple', 'mediumSeaGreen',
        'mediumSlateBlue', 'mediumSpringGreen', 'mediumTurquoise',
        'mediumVioletRed', 'midnightBlue', 'mintCream', 'mistyRose', 'moccasin',
        'navajoWhite', 'navy', 'oldLace', 'olive', 'oliveDrab', 'orange',
        'orangeRed', 'orchid', 'paleGoldenrod', 'paleGreen', 'paleTurquoise',
        'paleVioletRed', 'papayaWhip', 'peachPuff', 'peru', 'pink', 'plum',
        'powderBlue', 'purple', 'red', 'rosyBrown', 'royalBlue', 'saddleBrown',
        'salmon', 'sandyBrown', 'seaGreen', 'seaShell', 'sienna', 'silver',
        'skyBlue', 'slateBlue', 'slateGray', 'slateGrey', 'snow', 'springGreen',
        'steelBlue', 'tan', 'teal', 'thistle', 'tomato', 'turquoise', 'violet',
        'wheat', 'white', 'whiteSmoke', 'yellow', 'yellowGreen'
    ]


SCHEME_COLORS= ['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2', 'accent3',
                'accent4', 'accent5', 'accent6', 'hlink', 'folHlink', 'phClr', 'dk1', 'lt1',
                'dk2', 'lt2'
                ]

SYSTEM_COLOR_VALS = frozenset(
    (
        "scrollBar",
        "background",
        "activeCaption",
        "inactiveCaption",
        "menu",
        "window",
        "windowFrame",
        "menuText",
        "windowText",
        "captionText",
        "activeBorder",
        "inactiveBorder",
        "appWorkspace",
        "highlight",
        "highlightText",
        "btnFace",
        "btnShadow",
        "grayText",
        "btnText",
        "inactiveCaptionText",
        "btnHighlight",
        "3dDkShadow",
        "3dLight",
        "infoText",
        "infoBk",
        "hotLight",
        "gradientActiveCaption",
        "gradientInactiveCaption",
        "menuHighlight",
        "menuBar",
    )
)

SCHEME_COLOR_VALS = frozenset(SCHEME_COLORS)


def _system_color_val(v):
    if v is None:
        return None
    if v not in SYSTEM_COLOR_VALS:
        raise FieldValidationError(f"val rejected value {v!r}")
    return v


def _scheme_color_val(v):
    if v is None:
        return None
    if v not in SCHEME_COLOR_VALS:
        raise FieldValidationError(f"val rejected value {v!r}")
    return v


def _system_last_clr(v):
    if v is None:
        return None
    if not isinstance(v, str) or aRGB_REGEX.match(v) is None:
        raise FieldValidationError(f"lastClr rejected value {v!r}")
    if len(v) == 6:
        return "00" + v
    return v


def _drawml_empty_element(tag, value, ns=None):
    if not value:
        return None
    return Element(namespaced_tag(tag, ns or DRAWING_NS))


class Transform(TypedSerialisable):

    tagname = "comp"

    def __init__(self):
        pass


class SystemColor(TypedSerialisable):

    tagname = "sysClr"
    namespace = DRAWING_NS

    tint: int | None = Field.nested_value(expected_type=int, allow_none=True)
    shade: int | None = Field.nested_value(expected_type=int, allow_none=True)
    comp: Transform | None = Field.element(expected_type=Transform, allow_none=True)
    inv: Transform | None = Field.element(expected_type=Transform, allow_none=True)
    gray: Transform | None = Field.element(expected_type=Transform, allow_none=True)
    alpha: int | None = Field.nested_value(expected_type=int, allow_none=True)
    alphaOff: int | None = Field.nested_value(expected_type=int, allow_none=True)
    alphaMod: int | None = Field.nested_value(expected_type=int, allow_none=True)
    hue: int | None = Field.nested_value(expected_type=int, allow_none=True)
    hueOff: int | None = Field.nested_value(expected_type=int, allow_none=True)
    hueMod: int | None = Field.nested_value(expected_type=int, allow_none=True)
    sat: int | None = Field.nested_value(expected_type=int, allow_none=True)
    satOff: int | None = Field.nested_value(expected_type=int, allow_none=True)
    satMod: int | None = Field.nested_value(expected_type=int, allow_none=True)
    lum: int | None = Field.nested_value(expected_type=int, allow_none=True)
    lumOff: int | None = Field.nested_value(expected_type=int, allow_none=True)
    lumMod: int | None = Field.nested_value(expected_type=int, allow_none=True)
    red: int | None = Field.nested_value(expected_type=int, allow_none=True)
    redOff: int | None = Field.nested_value(expected_type=int, allow_none=True)
    redMod: int | None = Field.nested_value(expected_type=int, allow_none=True)
    green: int | None = Field.nested_value(expected_type=int, allow_none=True)
    greenOff: int | None = Field.nested_value(expected_type=int, allow_none=True)
    greenMod: int | None = Field.nested_value(expected_type=int, allow_none=True)
    blue: int | None = Field.nested_value(expected_type=int, allow_none=True)
    blueOff: int | None = Field.nested_value(expected_type=int, allow_none=True)
    blueMod: int | None = Field.nested_value(expected_type=int, allow_none=True)
    gamma: Transform | None = Field.element(expected_type=Transform, allow_none=True)
    invGamma: Transform | None = Field.element(expected_type=Transform, allow_none=True)

    val: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        default="windowText",
        converter=_system_color_val,
    )
    lastClr: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=_system_last_clr,
    )

    xml_order = (
        "tint",
        "shade",
        "comp",
        "inv",
        "gray",
        "alpha",
        "alphaOff",
        "alphaMod",
        "hue",
        "hueOff",
        "hueMod",
        "sat",
        "satOff",
        "satMod",
        "lum",
        "lumOff",
        "lumMod",
        "red",
        "redOff",
        "redMod",
        "green",
        "greenOff",
        "greenMod",
        "blue",
        "blueOff",
        "blueMod",
        "gamma",
        "invGamma",
    )

    def __init__(self,
                 val="windowText",
                 lastClr=None,
                 tint=None,
                 shade=None,
                 comp=None,
                 inv=None,
                 gray=None,
                 alpha=None,
                 alphaOff=None,
                 alphaMod=None,
                 hue=None,
                 hueOff=None,
                 hueMod=None,
                 sat=None,
                 satOff=None,
                 satMod=None,
                 lum=None,
                 lumOff=None,
                 lumMod=None,
                 red=None,
                 redOff=None,
                 redMod=None,
                 green=None,
                 greenOff=None,
                 greenMod=None,
                 blue=None,
                 blueOff=None,
                 blueMod=None,
                 gamma=None,
                 invGamma=None
                ):
        self.val = val
        self.lastClr = lastClr
        self.tint = tint
        self.shade = shade
        self.comp = comp
        self.inv = inv
        self.gray = gray
        self.alpha = alpha
        self.alphaOff = alphaOff
        self.alphaMod = alphaMod
        self.hue = hue
        self.hueOff = hueOff
        self.hueMod = hueMod
        self.sat = sat
        self.satOff = satOff
        self.satMod = satMod
        self.lum = lum
        self.lumOff = lumOff
        self.lumMod = lumMod
        self.red = red
        self.redOff = redOff
        self.redMod = redMod
        self.green = green
        self.greenOff = greenOff
        self.greenMod = greenMod
        self.blue = blue
        self.blueOff = blueOff
        self.blueMod = blueMod
        self.gamma = gamma
        self.invGamma = invGamma


class HSLColor(TypedSerialisable):

    tagname = "hslClr"

    hue: int | None = Field.attribute(expected_type=int, allow_none=True)
    sat: float | int | None = Field.attribute(
        expected_type=float,
        allow_none=True,
        converter=lambda v: _range_converter(v, field_name="sat", min_value=0, max_value=100),
    )
    lum: float | int | None = Field.attribute(
        expected_type=float,
        allow_none=True,
        converter=lambda v: _range_converter(v, field_name="lum", min_value=0, max_value=100),
    )

    #TODO add color transform options

    def __init__(self,
                 hue=None,
                 sat=None,
                 lum=None,
                ):
        self.hue = hue
        self.sat = sat
        self.lum = lum



class RGBPercent(TypedSerialisable):

    tagname = "rgbClr"

    r: float | int | None = Field.attribute(
        expected_type=float,
        allow_none=True,
        converter=lambda v: _range_converter(v, field_name="r", min_value=0, max_value=100),
    )
    g: float | int | None = Field.attribute(
        expected_type=float,
        allow_none=True,
        converter=lambda v: _range_converter(v, field_name="g", min_value=0, max_value=100),
    )
    b: float | int | None = Field.attribute(
        expected_type=float,
        allow_none=True,
        converter=lambda v: _range_converter(v, field_name="b", min_value=0, max_value=100),
    )

    #TODO add color transform options

    def __init__(self,
                 r=None,
                 g=None,
                 b=None,
                ):
        self.r = r
        self.g = g
        self.b = b


class SchemeColor(TypedSerialisable):

    tagname = "schemeClr"
    namespace = DRAWING_NS

    tint: int | None = Field.nested_value(expected_type=int, allow_none=True)
    shade: int | None = Field.nested_value(expected_type=int, allow_none=True)
    comp: bool | None = Field.nested_bool(
        allow_none=True,
        renderer=_drawml_empty_element,
    )
    inv: int | None = Field.nested_value(expected_type=int, allow_none=True)
    gray: int | None = Field.nested_value(expected_type=int, allow_none=True)
    alpha: int | None = Field.nested_value(expected_type=int, allow_none=True)
    alphaOff: int | None = Field.nested_value(expected_type=int, allow_none=True)
    alphaMod: int | None = Field.nested_value(expected_type=int, allow_none=True)
    hue: int | None = Field.nested_value(expected_type=int, allow_none=True)
    hueOff: int | None = Field.nested_value(expected_type=int, allow_none=True)
    hueMod: int | None = Field.nested_value(expected_type=int, allow_none=True)
    sat: int | None = Field.nested_value(expected_type=int, allow_none=True)
    satOff: int | None = Field.nested_value(expected_type=int, allow_none=True)
    satMod: int | None = Field.nested_value(expected_type=int, allow_none=True)
    lum: int | None = Field.nested_value(expected_type=int, allow_none=True)
    lumOff: int | None = Field.nested_value(expected_type=int, allow_none=True)
    lumMod: int | None = Field.nested_value(expected_type=int, allow_none=True)
    red: int | None = Field.nested_value(expected_type=int, allow_none=True)
    redOff: int | None = Field.nested_value(expected_type=int, allow_none=True)
    redMod: int | None = Field.nested_value(expected_type=int, allow_none=True)
    green: int | None = Field.nested_value(expected_type=int, allow_none=True)
    greenOff: int | None = Field.nested_value(expected_type=int, allow_none=True)
    greenMod: int | None = Field.nested_value(expected_type=int, allow_none=True)
    blue: int | None = Field.nested_value(expected_type=int, allow_none=True)
    blueOff: int | None = Field.nested_value(expected_type=int, allow_none=True)
    blueMod: int | None = Field.nested_value(expected_type=int, allow_none=True)
    gamma: bool | None = Field.nested_bool(
        allow_none=True,
        renderer=_drawml_empty_element,
    )
    invGamma: bool | None = Field.nested_bool(
        allow_none=True,
        renderer=_drawml_empty_element,
    )
    val: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=_scheme_color_val,
    )

    xml_order = (
        "tint",
        "shade",
        "comp",
        "inv",
        "gray",
        "alpha",
        "alphaOff",
        "alphaMod",
        "hue",
        "hueOff",
        "hueMod",
        "sat",
        "satOff",
        "satMod",
        "lum",
        "lumMod",
        "lumOff",
        "red",
        "redOff",
        "redMod",
        "green",
        "greenOff",
        "greenMod",
        "blue",
        "blueOff",
        "blueMod",
        "gamma",
        "invGamma",
    )

    def __init__(self,
                 tint=None,
                 shade=None,
                 comp=None,
                 inv=None,
                 gray=None,
                 alpha=None,
                 alphaOff=None,
                 alphaMod=None,
                 hue=None,
                 hueOff=None,
                 hueMod=None,
                 sat=None,
                 satOff=None,
                 satMod=None,
                 lum=None,
                 lumOff=None,
                 lumMod=None,
                 red=None,
                 redOff=None,
                 redMod=None,
                 green=None,
                 greenOff=None,
                 greenMod=None,
                 blue=None,
                 blueOff=None,
                 blueMod=None,
                 gamma=None,
                 invGamma=None,
                 val=None,
                ):
        self.tint = tint
        self.shade = shade
        self.comp = comp
        self.inv = inv
        self.gray = gray
        self.alpha = alpha
        self.alphaOff = alphaOff
        self.alphaMod = alphaMod
        self.hue = hue
        self.hueOff = hueOff
        self.hueMod = hueMod
        self.sat = sat
        self.satOff = satOff
        self.satMod = satMod
        self.lum = lum
        self.lumOff = lumOff
        self.lumMod = lumMod
        self.red = red
        self.redOff = redOff
        self.redMod = redMod
        self.green = green
        self.greenOff = greenOff
        self.greenMod = greenMod
        self.blue = blue
        self.blueOff = blueOff
        self.blueMod = blueMod
        self.gamma = gamma
        self.invGamma = invGamma
        self.val = val

class ColorChoice(TypedSerialisable):

    tagname = "colorChoice"
    namespace = DRAWING_NS

    scrgbClr: RGBPercent | None = Field.element(expected_type=RGBPercent, allow_none=True)
    RGBPercent = AliasField("scrgbClr")
    srgbClr: str | None = Field.nested_value(expected_type=str, allow_none=True)
    RGB = AliasField("srgbClr")
    hslClr: HSLColor | None = Field.element(expected_type=HSLColor, allow_none=True)
    sysClr: SystemColor | None = Field.element(expected_type=SystemColor, allow_none=True)
    schemeClr: SchemeColor | None = Field.element(expected_type=SchemeColor, allow_none=True)
    prstClr: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, PRESET_COLORS, "prstClr"),
    )

    xml_order = ('scrgbClr', 'srgbClr', 'hslClr', 'sysClr', 'schemeClr', 'prstClr')

    def __init__(self,
                 scrgbClr=None,
                 srgbClr=None,
                 hslClr=None,
                 sysClr=None,
                 schemeClr=None,
                 prstClr=None,
                ):
        self.scrgbClr = scrgbClr
        self.srgbClr = srgbClr
        self.hslClr = hslClr
        self.sysClr = sysClr
        self.schemeClr = schemeClr
        self.prstClr = prstClr

_COLOR_SET = ('dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3',
               'accent4', 'accent5', 'accent6', 'hlink', 'folHlink')


class ColorMapping(Serialisable):

    tagname = "clrMapOvr"

    bg1 = Set(values=_COLOR_SET)
    tx1 = Set(values=_COLOR_SET)
    bg2 = Set(values=_COLOR_SET)
    tx2 = Set(values=_COLOR_SET)
    accent1 = Set(values=_COLOR_SET)
    accent2 = Set(values=_COLOR_SET)
    accent3 = Set(values=_COLOR_SET)
    accent4 = Set(values=_COLOR_SET)
    accent5 = Set(values=_COLOR_SET)
    accent6 = Set(values=_COLOR_SET)
    hlink = Set(values=_COLOR_SET)
    folHlink = Set(values=_COLOR_SET)
    extLst = Typed(expected_type=OfficeArtExtensionList, allow_none=True)

    def __init__(self,
                 bg1="lt1",
                 tx1="dk1",
                 bg2="lt2",
                 tx2="dk2",
                 accent1="accent1",
                 accent2="accent2",
                 accent3="accent3",
                 accent4="accent4",
                 accent5="accent5",
                 accent6="accent6",
                 hlink="hlink",
                 folHlink="folHlink",
                 extLst=None,
                ):
        self.bg1 = bg1
        self.tx1 = tx1
        self.bg2 = bg2
        self.tx2 = tx2
        self.accent1 = accent1
        self.accent2 = accent2
        self.accent3 = accent3
        self.accent4 = accent4
        self.accent5 = accent5
        self.accent6 = accent6
        self.hlink = hlink
        self.folHlink = folHlink
        self.extLst = extLst


class ColorChoiceDescriptor(Typed):
    """
    Objects can choose from 7 different kinds of color system.
    Assume RGBHex if a string is passed in.
    """

    expected_type = ColorChoice
    allow_none = True

    def __set__(self, instance, value):
        if isinstance(value, str):
            value = ColorChoice(srgbClr=value)
        else:
            if hasattr(self, "namespace") and value is not None:
                value.namespace = self.namespace
        super().__set__(instance, value)


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


def _enum_converter(value, allowed_values, field_name: str):
    if value is None:
        return None
    if value not in allowed_values:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return value
