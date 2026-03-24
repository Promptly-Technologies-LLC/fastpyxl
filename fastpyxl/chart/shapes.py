# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from fastpyxl.xml.constants import DRAWING_NS
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.drawing.colors import ColorChoice
from fastpyxl.descriptors.excel import ExtensionList as OfficeArtExtensionList
from fastpyxl.drawing.fill import GradientFillProperties, PatternFillProperties
from fastpyxl.drawing.line import LineProperties
from fastpyxl.drawing.geometry import (
    Shape3D,
    Scene3D,
    Transform2D,
    CustomGeometry2D,
    PresetGeometry2D,
)


_BW_VALUES = frozenset(
    {
        "clr",
        "auto",
        "gray",
        "ltGray",
        "invGray",
        "grayWhite",
        "blackGray",
        "blackWhite",
        "black",
        "white",
        "hidden",
    }
)


def _bw_mode_converter(value):
    if value is None:
        return None
    if value not in _BW_VALUES:
        raise FieldValidationError(f"bwMode rejected value {value!r}")
    return value


def _solid_fill_converter(value):
    if value is None:
        return None
    if isinstance(value, str):
        return ColorChoice(srgbClr=value)
    return value


class _NoFill(Serialisable):
    tagname = "noFill"
    namespace = DRAWING_NS


def _no_fill_converter(value):
    if value is True:
        return _NoFill()
    if value in (False, None):
        return None
    return value


class GraphicalProperties(Serialisable):
    """
    Somewhat vaguely 21.2.2.197 says this:

    This element specifies the formatting for the parent chart element. The
    custGeom, prstGeom, scene3d, and xfrm elements are not supported. The
    bwMode attribute is not supported.

    This doesn't leave much. And the element is used in different places.
    """

    tagname = "spPr"

    bwMode: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=_bw_mode_converter,
    )
    xfrm: Transform2D | None = Field.element(expected_type=Transform2D, allow_none=True)
    transform = AliasField("xfrm")
    custGeom: CustomGeometry2D | None = Field.element(
        expected_type=CustomGeometry2D, allow_none=True, serialize=False
    )
    prstGeom: PresetGeometry2D | None = Field.element(
        expected_type=PresetGeometry2D, allow_none=True
    )

    noFill: _NoFill | None = Field.element(
        expected_type=_NoFill,
        allow_none=True,
        converter=_no_fill_converter,
    )
    solidFill: ColorChoice | None = Field.element(
        expected_type=ColorChoice,
        allow_none=True,
        converter=_solid_fill_converter,
    )
    gradFill: GradientFillProperties | None = Field.element(
        expected_type=GradientFillProperties, allow_none=True
    )
    pattFill: PatternFillProperties | None = Field.element(
        expected_type=PatternFillProperties, allow_none=True
    )

    ln: LineProperties | None = Field.element(expected_type=LineProperties, allow_none=True)
    line = AliasField("ln")
    scene3d: Scene3D | None = Field.element(expected_type=Scene3D, allow_none=True)
    sp3d: Shape3D | None = Field.element(expected_type=Shape3D, allow_none=True)
    shape3D = AliasField("sp3d")
    extLst: OfficeArtExtensionList | None = Field.element(
        expected_type=OfficeArtExtensionList, allow_none=True, serialize=False
    )

    xml_order = (
        "xfrm",
        "prstGeom",
        "noFill",
        "solidFill",
        "gradFill",
        "pattFill",
        "ln",
        "scene3d",
        "sp3d",
    )

    def __init__(
        self,
        bwMode=None,
        xfrm=None,
        noFill=None,
        solidFill=None,
        gradFill=None,
        pattFill=None,
        ln=None,
        scene3d=None,
        custGeom=None,
        prstGeom=None,
        sp3d=None,
        extLst=None,
    ):
        self.bwMode = bwMode
        self.xfrm = xfrm
        self.noFill = noFill
        self.solidFill = solidFill
        self.gradFill = gradFill
        self.pattFill = pattFill
        if ln is None:
            ln = LineProperties()
        self.ln = ln
        self.custGeom = custGeom
        self.prstGeom = prstGeom
        self.scene3d = scene3d
        self.sp3d = sp3d
        self.extLst = extLst
