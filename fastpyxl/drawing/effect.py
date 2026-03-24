# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.descriptors.serialisable import Serialisable
from fastpyxl.descriptors import (
    Set,
    Integer,
)
from fastpyxl.typed_serialisable.base import Serialisable as TypedSerialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import Field

from .colors import ColorChoice


class TintEffect(TypedSerialisable):

    tagname = "tint"

    hue: int | None = Field.attribute(expected_type=int, allow_none=True)
    amt: int | None = Field.attribute(expected_type=int, allow_none=True)

    def __init__(self,
                 hue=0,
                 amt=0,
                ):
        self.hue = hue
        self.amt = amt


class LuminanceEffect(TypedSerialisable):

    tagname = "lum"

    bright: int | None = Field.attribute(expected_type=int, allow_none=True)  # Pct?
    contrast: int | None = Field.attribute(expected_type=int, allow_none=True)  # Pct#

    def __init__(self,
                 bright=0,
                 contrast=0,
                ):
        self.bright = bright
        self.contrast = contrast


class HSLEffect(Serialisable):

    hue = Integer()
    sat = Integer()
    lum = Integer()

    def __init__(self,
                 hue=None,
                 sat=None,
                 lum=None,
                ):
        self.hue = hue
        self.sat = sat
        self.lum = lum


class GrayscaleEffect(TypedSerialisable):

    tagname = "grayscl"


class FillOverlayEffect(Serialisable):

    blend = Set(values=(['over', 'mult', 'screen', 'darken', 'lighten']))

    def __init__(self,
                 blend=None,
                ):
        self.blend = blend


class DuotoneEffect(Serialisable):

    pass

class ColorReplaceEffect(Serialisable):

    pass

class Color(Serialisable):

    pass

class ColorChangeEffect(TypedSerialisable):

    useA: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    clrFrom: Color | None = Field.element(expected_type=Color, allow_none=True)
    clrTo: Color | None = Field.element(expected_type=Color, allow_none=True)
    xml_order = ("clrFrom", "clrTo")

    def __init__(self,
                 useA=None,
                 clrFrom=None,
                 clrTo=None,
                ):
        self.useA = useA
        self.clrFrom = clrFrom
        self.clrTo = clrTo


class BlurEffect(TypedSerialisable):

    rad: float | None = Field.attribute(expected_type=float, allow_none=True)
    grow: bool | None = Field.attribute(expected_type=bool, allow_none=True)

    def __init__(self,
                 rad=None,
                 grow=None,
                ):
        self.rad = rad
        self.grow = grow


class BiLevelEffect(TypedSerialisable):

    thresh: int | None = Field.attribute(expected_type=int, allow_none=True)

    def __init__(self,
                 thresh=None,
                ):
        self.thresh = thresh


class AlphaReplaceEffect(TypedSerialisable):

    a: int | None = Field.attribute(expected_type=int, allow_none=True)

    def __init__(self,
                 a=None,
                ):
        self.a = a


class AlphaModulateFixedEffect(TypedSerialisable):

    amt: int | None = Field.attribute(expected_type=int, allow_none=True)

    def __init__(self,
                 amt=None,
                ):
        self.amt = amt


class EffectContainer(TypedSerialisable):

    type: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("sib", "tree"), "type"),
    )
    name: str | None = Field.attribute(expected_type=str, allow_none=True)

    def __init__(self,
                 type=None,
                 name=None,
                ):
        self.type = type
        self.name = name


class AlphaModulateEffect(TypedSerialisable):

    cont: EffectContainer | None = Field.element(expected_type=EffectContainer, allow_none=True)

    def __init__(self,
                 cont=None,
                ):
        self.cont = cont


class AlphaInverseEffect(Serialisable):

    pass

class AlphaFloorEffect(Serialisable):

    pass

class AlphaCeilingEffect(Serialisable):

    pass

class AlphaBiLevelEffect(TypedSerialisable):

    thresh: int | None = Field.attribute(expected_type=int, allow_none=True)

    def __init__(self,
                 thresh=None,
                ):
        self.thresh = thresh


class GlowEffect(ColorChoice):

    rad: float | None = Field.attribute(expected_type=float, allow_none=True)
    def __init__(self,
                 rad=None,
                 **kw
                ):
        self.rad = rad
        super().__init__(**kw)


class InnerShadowEffect(ColorChoice):

    blurRad: float | None = Field.attribute(expected_type=float, allow_none=True)
    dist: float | None = Field.attribute(expected_type=float, allow_none=True)
    dir: int | None = Field.attribute(expected_type=int, allow_none=True)
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

    blurRad: float | None = Field.attribute(expected_type=float, allow_none=True)
    dist: float | None = Field.attribute(expected_type=float, allow_none=True)
    dir: int | None = Field.attribute(expected_type=int, allow_none=True)
    sx: int | None = Field.attribute(expected_type=int, allow_none=True)
    sy: int | None = Field.attribute(expected_type=int, allow_none=True)
    kx: int | None = Field.attribute(expected_type=int, allow_none=True)
    ky: int | None = Field.attribute(expected_type=int, allow_none=True)
    algn: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("tl", "t", "tr", "l", "ctr", "r", "bl", "b", "br"), "algn"),
    )
    rotWithShape: bool | None = Field.attribute(expected_type=bool, allow_none=True)
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
        converter=lambda v: _enum_converter(v, tuple(f"shdw{i}" for i in range(1, 21)), "prst"),
    )
    dist: float | None = Field.attribute(expected_type=float, allow_none=True)
    dir: int | None = Field.attribute(expected_type=int, allow_none=True)
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


class ReflectionEffect(TypedSerialisable):

    blurRad: float | None = Field.attribute(expected_type=float, allow_none=True)
    stA: int | None = Field.attribute(expected_type=int, allow_none=True)
    stPos: int | None = Field.attribute(expected_type=int, allow_none=True)
    endA: int | None = Field.attribute(expected_type=int, allow_none=True)
    endPos: int | None = Field.attribute(expected_type=int, allow_none=True)
    dist: float | None = Field.attribute(expected_type=float, allow_none=True)
    dir: int | None = Field.attribute(expected_type=int, allow_none=True)
    fadeDir: int | None = Field.attribute(expected_type=int, allow_none=True)
    sx: int | None = Field.attribute(expected_type=int, allow_none=True)
    sy: int | None = Field.attribute(expected_type=int, allow_none=True)
    kx: int | None = Field.attribute(expected_type=int, allow_none=True)
    ky: int | None = Field.attribute(expected_type=int, allow_none=True)
    algn: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("tl", "t", "tr", "l", "ctr", "r", "bl", "b", "br"), "algn"),
    )
    rotWithShape: bool | None = Field.attribute(expected_type=bool, allow_none=True)

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


class SoftEdgesEffect(TypedSerialisable):

    rad: float | None = Field.attribute(expected_type=float, allow_none=True)

    def __init__(self,
                 rad=None,
                ):
        self.rad = rad


class EffectList(TypedSerialisable):

    blur: BlurEffect | None = Field.element(expected_type=BlurEffect, allow_none=True)
    fillOverlay: FillOverlayEffect | None = Field.element(expected_type=FillOverlayEffect, allow_none=True)
    glow: GlowEffect | None = Field.element(expected_type=GlowEffect, allow_none=True)
    innerShdw: InnerShadowEffect | None = Field.element(expected_type=InnerShadowEffect, allow_none=True)
    outerShdw: OuterShadow | None = Field.element(expected_type=OuterShadow, allow_none=True)
    prstShdw: PresetShadowEffect | None = Field.element(expected_type=PresetShadowEffect, allow_none=True)
    reflection: ReflectionEffect | None = Field.element(expected_type=ReflectionEffect, allow_none=True)
    softEdge: SoftEdgesEffect | None = Field.element(expected_type=SoftEdgesEffect, allow_none=True)

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
