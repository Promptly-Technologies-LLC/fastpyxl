# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import Field

from .colors import ColorChoice


class TintEffect(Serialisable):

    tagname = "tint"

    hue: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    amt: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

    def __init__(self,
                 hue=0,
                 amt=0,
                ):
        self.hue = hue
        self.amt = amt


class LuminanceEffect(Serialisable):

    tagname = "lum"

    bright: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)  # Pct?
    contrast: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)  # Pct#

    def __init__(self,
                 bright=0,
                 contrast=0,
                ):
        self.bright = bright
        self.contrast = contrast


class HSLEffect(Serialisable):

    tagname = "hsl"

    hue: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    sat: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    lum: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

    def __init__(self,
                 hue=None,
                 sat=None,
                 lum=None,
                ):
        self.hue = hue
        self.sat = sat
        self.lum = lum


class GrayscaleEffect(Serialisable):

    tagname = "grayscl"


class FillOverlayEffect(Serialisable):

    tagname = "fillOverlay"

    blend: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ('over', 'mult', 'screen', 'darken', 'lighten'), "blend"),
        default=None,
    )

    def __init__(self,
                 blend=None,
                ):
        self.blend = blend


class DuotoneEffect(Serialisable):

    tagname = "duotone"


class ColorReplaceEffect(Serialisable):

    tagname = "clrRepl"


class Color(Serialisable):

    tagname = "clr"


class ColorChangeEffect(Serialisable):

    useA: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    clrFrom: Color | None = Field.element(expected_type=Color, allow_none=True, default=None)
    clrTo: Color | None = Field.element(expected_type=Color, allow_none=True, default=None)
    xml_order = ("clrFrom", "clrTo")

    def __init__(self,
                 useA=None,
                 clrFrom=None,
                 clrTo=None,
                ):
        self.useA = useA
        self.clrFrom = clrFrom
        self.clrTo = clrTo


class BlurEffect(Serialisable):

    rad: float | None = Field.attribute(expected_type=float, allow_none=True, default=None)
    grow: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)

    def __init__(self,
                 rad=None,
                 grow=None,
                ):
        self.rad = rad
        self.grow = grow


class BiLevelEffect(Serialisable):

    thresh: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

    def __init__(self,
                 thresh=None,
                ):
        self.thresh = thresh


class AlphaReplaceEffect(Serialisable):

    a: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

    def __init__(self,
                 a=None,
                ):
        self.a = a


class AlphaModulateFixedEffect(Serialisable):

    amt: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

    def __init__(self,
                 amt=None,
                ):
        self.amt = amt


class EffectContainer(Serialisable):

    type: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("sib", "tree"), "type"), default=None,
    )
    name: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)

    def __init__(self,
                 type=None,
                 name=None,
                ):
        self.type = type
        self.name = name


class AlphaModulateEffect(Serialisable):

    cont: EffectContainer | None = Field.element(expected_type=EffectContainer, allow_none=True, default=None)

    def __init__(self,
                 cont=None,
                ):
        self.cont = cont


class AlphaInverseEffect(Serialisable):

    tagname = "alphaInv"


class AlphaFloorEffect(Serialisable):

    tagname = "alphaFloor"


class AlphaCeilingEffect(Serialisable):

    tagname = "alphaCeiling"


class AlphaBiLevelEffect(Serialisable):

    thresh: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

    def __init__(self,
                 thresh=None,
                ):
        self.thresh = thresh


class GlowEffect(ColorChoice):

    rad: float | None = Field.attribute(expected_type=float, allow_none=True, default=None)
    def __init__(self,
                 rad=None,
                 **kw
                ):
        self.rad = rad
        super().__init__(**kw)


class InnerShadowEffect(ColorChoice):

    blurRad: float | None = Field.attribute(expected_type=float, allow_none=True, default=None)
    dist: float | None = Field.attribute(expected_type=float, allow_none=True, default=None)
    dir: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    def __init__(self,
                 blurRad=None,
                 dist=None,
                 dir=None,
                 **kw
                 ):
        self.blurRad = blurRad
        self.dist = dist
        self.dir = dir
        super().__init__(**kw)


class OuterShadow(ColorChoice):

    tagname = "outerShdw"

    blurRad: float | None = Field.attribute(expected_type=float, allow_none=True, default=None)
    dist: float | None = Field.attribute(expected_type=float, allow_none=True, default=None)
    dir: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    sx: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    sy: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    kx: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    ky: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    algn: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("tl", "t", "tr", "l", "ctr", "r", "bl", "b", "br"), "algn"), default=None,
    )
    rotWithShape: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    def __init__(self,
                 blurRad=None,
                 dist=None,
                 dir=None,
                 sx=None,
                 sy=None,
                 kx=None,
                 ky=None,
                 algn=None,
                 rotWithShape=None,
                 **kw
                ):
        self.blurRad = blurRad
        self.dist = dist
        self.dir = dir
        self.sx = sx
        self.sy = sy
        self.kx = kx
        self.ky = ky
        self.algn = algn
        self.rotWithShape = rotWithShape
        super().__init__(**kw)


class PresetShadowEffect(ColorChoice):

    prst: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, tuple(f"shdw{i}" for i in range(1, 21)), "prst"), default=None,
    )
    dist: float | None = Field.attribute(expected_type=float, allow_none=True, default=None)
    dir: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    def __init__(self,
                 prst=None,
                 dist=None,
                 dir=None,
                 **kw
                ):
        self.prst = prst
        self.dist = dist
        self.dir = dir
        super().__init__(**kw)


class ReflectionEffect(Serialisable):

    blurRad: float | None = Field.attribute(expected_type=float, allow_none=True, default=None)
    stA: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    stPos: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    endA: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    endPos: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    dist: float | None = Field.attribute(expected_type=float, allow_none=True, default=None)
    dir: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    fadeDir: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    sx: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    sy: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    kx: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    ky: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    algn: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("tl", "t", "tr", "l", "ctr", "r", "bl", "b", "br"), "algn"), default=None,
    )
    rotWithShape: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)

    def __init__(self,
                 blurRad=None,
                 stA=None,
                 stPos=None,
                 endA=None,
                 endPos=None,
                 dist=None,
                 dir=None,
                 fadeDir=None,
                 sx=None,
                 sy=None,
                 kx=None,
                 ky=None,
                 algn=None,
                 rotWithShape=None,
                ):
        self.blurRad = blurRad
        self.stA = stA
        self.stPos = stPos
        self.endA = endA
        self.endPos = endPos
        self.dist = dist
        self.dir = dir
        self.fadeDir = fadeDir
        self.sx = sx
        self.sy = sy
        self.kx = kx
        self.ky = ky
        self.algn = algn
        self.rotWithShape = rotWithShape


class SoftEdgesEffect(Serialisable):

    rad: float | None = Field.attribute(expected_type=float, allow_none=True, default=None)

    def __init__(self,
                 rad=None,
                ):
        self.rad = rad


class EffectList(Serialisable):

    blur: BlurEffect | None = Field.element(expected_type=BlurEffect, allow_none=True, default=None)
    fillOverlay: FillOverlayEffect | None = Field.element(expected_type=FillOverlayEffect, allow_none=True, default=None)
    glow: GlowEffect | None = Field.element(expected_type=GlowEffect, allow_none=True, default=None)
    innerShdw: InnerShadowEffect | None = Field.element(expected_type=InnerShadowEffect, allow_none=True, default=None)
    outerShdw: OuterShadow | None = Field.element(expected_type=OuterShadow, allow_none=True, default=None)
    prstShdw: PresetShadowEffect | None = Field.element(expected_type=PresetShadowEffect, allow_none=True, default=None)
    reflection: ReflectionEffect | None = Field.element(expected_type=ReflectionEffect, allow_none=True, default=None)
    softEdge: SoftEdgesEffect | None = Field.element(expected_type=SoftEdgesEffect, allow_none=True, default=None)

    xml_order = ('blur', 'fillOverlay', 'glow', 'innerShdw', 'outerShdw',
                 'prstShdw', 'reflection', 'softEdge')

    def __init__(self,
                 blur=None,
                 fillOverlay=None,
                 glow=None,
                 innerShdw=None,
                 outerShdw=None,
                 prstShdw=None,
                 reflection=None,
                 softEdge=None,
                ):
        self.blur = blur
        self.fillOverlay = fillOverlay
        self.glow = glow
        self.innerShdw = innerShdw
        self.outerShdw = outerShdw
        self.prstShdw = prstShdw
        self.reflection = reflection
        self.softEdge = softEdge


def _enum_converter(value, allowed_values, field_name: str):
    if value is None:
        return None
    if value not in allowed_values:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return value
