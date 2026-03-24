# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.descriptors.excel import ExtensionList as OfficeArtExtensionList
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field
from fastpyxl.xml.constants import DRAWING_NS, REL_NS

from .colors import (
    PRESET_COLORS,
    ColorChoice,
    HSLColor,
    RGBPercent,
    SchemeColor,
    SystemColor,
)
from .effect import (
    AlphaBiLevelEffect,
    AlphaCeilingEffect,
    AlphaFloorEffect,
    AlphaInverseEffect,
    AlphaModulateEffect,
    AlphaModulateFixedEffect,
    AlphaReplaceEffect,
    BiLevelEffect,
    BlurEffect,
    ColorChangeEffect,
    ColorReplaceEffect,
    DuotoneEffect,
    FillOverlayEffect,
    GrayscaleEffect,
    HSLEffect,
    LuminanceEffect,
    TintEffect,
)

"""
Fill elements from drawing main schema
"""


def _enum_converter(value, allowed_values, field_name: str):
    if value is None:
        return None
    if value not in allowed_values:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return value


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


class PatternFillProperties(Serialisable):
    tagname = "pattFill"
    namespace = DRAWING_NS

    prst: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(
            v,
            (
                "pct5", "pct10", "pct20", "pct25", "pct30", "pct40", "pct50", "pct60", "pct70",
                "pct75", "pct80", "pct90", "horz", "vert", "ltHorz", "ltVert", "dkHorz", "dkVert",
                "narHorz", "narVert", "dashHorz", "dashVert", "cross", "dnDiag", "upDiag", "ltDnDiag",
                "ltUpDiag", "dkDnDiag", "dkUpDiag", "wdDnDiag", "wdUpDiag", "dashDnDiag", "dashUpDiag",
                "diagCross", "smCheck", "lgCheck", "smGrid", "lgGrid", "dotGrid", "smConfetti",
                "lgConfetti", "horzBrick", "diagBrick", "solidDmnd", "openDmnd", "dotDmnd", "plaid",
                "sphere", "weave", "divot", "shingle", "wave", "trellis", "zigZag",
            ),
            "prst",
        ),
    )
    preset = AliasField("prst")
    fgClr: ColorChoice | None = Field.element(expected_type=ColorChoice, allow_none=True)
    foreground = AliasField("fgClr")
    bgClr: ColorChoice | None = Field.element(expected_type=ColorChoice, allow_none=True)
    background = AliasField("bgClr")

    xml_order = ("fgClr", "bgClr")

    def __init__(self, prst=None, fgClr=None, bgClr=None):
        self.prst = prst
        self.fgClr = fgClr
        self.bgClr = bgClr


class RelativeRect(Serialisable):
    tagname = "rect"
    namespace = DRAWING_NS

    l: int | None = Field.attribute(expected_type=int, allow_none=True)  # noqa: E741
    left = AliasField("l")
    t: int | None = Field.attribute(expected_type=int, allow_none=True)
    top = AliasField("t")
    r: int | None = Field.attribute(expected_type=int, allow_none=True)
    right = AliasField("r")
    b: int | None = Field.attribute(expected_type=int, allow_none=True)
    bottom = AliasField("b")

    def __init__(self, l=None, t=None, r=None, b=None):  # noqa: E741
        self.l = l  # noqa: E741
        self.t = t
        self.r = r
        self.b = b


class StretchInfoProperties(Serialisable):
    tagname = "stretch"
    namespace = DRAWING_NS

    fillRect: RelativeRect | None = Field.element(expected_type=RelativeRect, allow_none=True)

    def __init__(self, fillRect=RelativeRect()):
        self.fillRect = fillRect


class GradientStop(Serialisable):
    tagname = "gs"
    namespace = DRAWING_NS

    pos: int | None = Field.attribute(
        expected_type=int,
        allow_none=True,
        converter=lambda v: _range_converter(v, 0, 100000, "pos"),
    )
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

    xml_order = ("scrgbClr", "srgbClr", "hslClr", "sysClr", "schemeClr", "prstClr")

    def __init__(self, pos=None, scrgbClr=None, srgbClr=None, hslClr=None, sysClr=None, schemeClr=None, prstClr=None):
        if pos is None:
            pos = 0
        self.pos = pos
        self.scrgbClr = scrgbClr
        self.srgbClr = srgbClr
        self.hslClr = hslClr
        self.sysClr = sysClr
        self.schemeClr = schemeClr
        self.prstClr = prstClr


class LinearShadeProperties(Serialisable):
    tagname = "lin"
    namespace = DRAWING_NS

    ang: int | None = Field.attribute(expected_type=int, allow_none=True)
    scaled: bool | None = Field.attribute(expected_type=bool, allow_none=True)

    def __init__(self, ang=None, scaled=None):
        self.ang = ang
        self.scaled = scaled


class PathShadeProperties(Serialisable):
    tagname = "path"
    namespace = DRAWING_NS

    path: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("shape", "circle", "rect"), "path"),
    )
    fillToRect: RelativeRect | None = Field.element(expected_type=RelativeRect, allow_none=True)

    def __init__(self, path=None, fillToRect=None):
        self.path = path
        self.fillToRect = fillToRect


class GradientFillProperties(Serialisable):
    tagname = "gradFill"
    namespace = DRAWING_NS

    flip: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("x", "y", "xy"), "flip"),
    )
    rotWithShape: bool | None = Field.attribute(expected_type=bool, allow_none=True)

    gsLst: list[GradientStop] = Field.nested_sequence(expected_type=GradientStop, count=False)
    stop_list = AliasField("gsLst")

    lin: LinearShadeProperties | None = Field.element(expected_type=LinearShadeProperties, allow_none=True)
    linear = AliasField("lin")
    path: PathShadeProperties | None = Field.element(expected_type=PathShadeProperties, allow_none=True)
    tileRect: RelativeRect | None = Field.element(expected_type=RelativeRect, allow_none=True)

    xml_order = ("gsLst", "lin", "path", "tileRect")

    def __init__(self, flip=None, rotWithShape=None, gsLst=(), lin=None, path=None, tileRect=None):
        self.flip = flip
        self.rotWithShape = rotWithShape
        self.gsLst = list(gsLst)
        self.lin = lin
        self.path = path
        self.tileRect = tileRect


class SolidColorFillProperties(Serialisable):
    tagname = "solidFill"

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

    xml_order = ("scrgbClr", "srgbClr", "hslClr", "sysClr", "schemeClr", "prstClr")

    def __init__(self, scrgbClr=None, srgbClr=None, hslClr=None, sysClr=None, schemeClr=None, prstClr=None):
        self.scrgbClr = scrgbClr
        self.srgbClr = srgbClr
        self.hslClr = hslClr
        self.sysClr = sysClr
        self.schemeClr = schemeClr
        self.prstClr = prstClr


class Blip(Serialisable):
    tagname = "blip"
    namespace = DRAWING_NS

    cstate: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("email", "screen", "print", "hqprint"), "cstate"),
    )
    embed: str | None = Field.attribute(expected_type=str, allow_none=True, namespace=REL_NS)
    link: str | None = Field.attribute(expected_type=str, allow_none=True, namespace=REL_NS)
    noGrp: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    noSelect: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    noRot: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    noChangeAspect: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    noMove: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    noResize: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    noEditPoints: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    noAdjustHandles: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    noChangeArrowheads: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    noChangeShapeType: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    extLst: OfficeArtExtensionList | None = Field.element(expected_type=OfficeArtExtensionList, allow_none=True)
    alphaBiLevel: AlphaBiLevelEffect | None = Field.element(expected_type=AlphaBiLevelEffect, allow_none=True)
    alphaCeiling: AlphaCeilingEffect | None = Field.element(expected_type=AlphaCeilingEffect, allow_none=True)
    alphaFloor: AlphaFloorEffect | None = Field.element(expected_type=AlphaFloorEffect, allow_none=True)
    alphaInv: AlphaInverseEffect | None = Field.element(expected_type=AlphaInverseEffect, allow_none=True)
    alphaMod: AlphaModulateEffect | None = Field.element(expected_type=AlphaModulateEffect, allow_none=True)
    alphaModFix: AlphaModulateFixedEffect | None = Field.element(expected_type=AlphaModulateFixedEffect, allow_none=True)
    alphaRepl: AlphaReplaceEffect | None = Field.element(expected_type=AlphaReplaceEffect, allow_none=True)
    biLevel: BiLevelEffect | None = Field.element(expected_type=BiLevelEffect, allow_none=True)
    blur: BlurEffect | None = Field.element(expected_type=BlurEffect, allow_none=True)
    clrChange: ColorChangeEffect | None = Field.element(expected_type=ColorChangeEffect, allow_none=True)
    clrRepl: ColorReplaceEffect | None = Field.element(expected_type=ColorReplaceEffect, allow_none=True)
    duotone: DuotoneEffect | None = Field.element(expected_type=DuotoneEffect, allow_none=True)
    fillOverlay: FillOverlayEffect | None = Field.element(expected_type=FillOverlayEffect, allow_none=True)
    grayscl: GrayscaleEffect | None = Field.element(expected_type=GrayscaleEffect, allow_none=True)
    hsl: HSLEffect | None = Field.element(expected_type=HSLEffect, allow_none=True)
    lum: LuminanceEffect | None = Field.element(expected_type=LuminanceEffect, allow_none=True)
    tint: TintEffect | None = Field.element(expected_type=TintEffect, allow_none=True)

    xml_order = (
        "alphaBiLevel", "alphaCeiling", "alphaFloor", "alphaInv", "alphaMod", "alphaModFix", "alphaRepl",
        "biLevel", "blur", "clrChange", "clrRepl", "duotone", "fillOverlay", "grayscl", "hsl", "lum", "tint",
    )

    def __init__(
        self,
        cstate=None,
        embed=None,
        link=None,
        noGrp=None,
        noSelect=None,
        noRot=None,
        noChangeAspect=None,
        noMove=None,
        noResize=None,
        noEditPoints=None,
        noAdjustHandles=None,
        noChangeArrowheads=None,
        noChangeShapeType=None,
        extLst=None,
        alphaBiLevel=None,
        alphaCeiling=None,
        alphaFloor=None,
        alphaInv=None,
        alphaMod=None,
        alphaModFix=None,
        alphaRepl=None,
        biLevel=None,
        blur=None,
        clrChange=None,
        clrRepl=None,
        duotone=None,
        fillOverlay=None,
        grayscl=None,
        hsl=None,
        lum=None,
        tint=None,
    ):
        self.cstate = cstate
        self.embed = embed
        self.link = link
        self.noGrp = noGrp
        self.noSelect = noSelect
        self.noRot = noRot
        self.noChangeAspect = noChangeAspect
        self.noMove = noMove
        self.noResize = noResize
        self.noEditPoints = noEditPoints
        self.noAdjustHandles = noAdjustHandles
        self.noChangeArrowheads = noChangeArrowheads
        self.noChangeShapeType = noChangeShapeType
        self.extLst = extLst
        self.alphaBiLevel = alphaBiLevel
        self.alphaCeiling = alphaCeiling
        self.alphaFloor = alphaFloor
        self.alphaInv = alphaInv
        self.alphaMod = alphaMod
        self.alphaModFix = alphaModFix
        self.alphaRepl = alphaRepl
        self.biLevel = biLevel
        self.blur = blur
        self.clrChange = clrChange
        self.clrRepl = clrRepl
        self.duotone = duotone
        self.fillOverlay = fillOverlay
        self.grayscl = grayscl
        self.hsl = hsl
        self.lum = lum
        self.tint = tint


class TileInfoProperties(Serialisable):
    tx: int | None = Field.attribute(expected_type=int, allow_none=True)
    ty: int | None = Field.attribute(expected_type=int, allow_none=True)
    sx: int | None = Field.attribute(expected_type=int, allow_none=True)
    sy: int | None = Field.attribute(expected_type=int, allow_none=True)
    flip: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("x", "y", "xy"), "flip"),
    )
    algn: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("tl", "t", "tr", "l", "ctr", "r", "bl", "b", "br"), "algn"),
    )

    def __init__(self, tx=None, ty=None, sx=None, sy=None, flip=None, algn=None):
        self.tx = tx
        self.ty = ty
        self.sx = sx
        self.sy = sy
        self.flip = flip
        self.algn = algn


class BlipFillProperties(Serialisable):
    tagname = "blipFill"

    dpi: int | None = Field.attribute(expected_type=int, allow_none=True)
    rotWithShape: bool | None = Field.attribute(expected_type=bool, allow_none=True)

    blip: Blip | None = Field.element(expected_type=Blip, allow_none=True)
    srcRect: RelativeRect | None = Field.element(expected_type=RelativeRect, allow_none=True)
    tile: TileInfoProperties | None = Field.element(expected_type=TileInfoProperties, allow_none=True)
    stretch: StretchInfoProperties | None = Field.element(expected_type=StretchInfoProperties, allow_none=True)

    xml_order = ("blip", "srcRect", "tile", "stretch")

    def __init__(self, dpi=None, rotWithShape=None, blip=None, tile=None, stretch=StretchInfoProperties(), srcRect=None):
        self.dpi = dpi
        self.rotWithShape = rotWithShape
        self.blip = blip
        self.tile = tile
        self.stretch = stretch
        self.srcRect = srcRect
