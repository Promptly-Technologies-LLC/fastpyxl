# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.xml.constants import DRAWING_NS
from fastpyxl.descriptors.excel import ExtensionList as OfficeArtExtensionList
from fastpyxl.typed_serialisable.base import Serialisable as TypedSerialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import Field

from .geometry import GroupTransform2D, Scene3D
from .text import Hyperlink


class GroupShapeProperties(TypedSerialisable):

    tagname = "grpSpPr"

    bwMode: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(
            v,
            ("clr", "auto", "gray", "ltGray", "invGray", "grayWhite", "blackGray", "blackWhite", "black", "white", "hidden"),
            "bwMode",
        ), default=None,
    )
    xfrm: GroupTransform2D | None = Field.element(expected_type=GroupTransform2D, allow_none=True, default=None)
    scene3d: Scene3D | None = Field.element(expected_type=Scene3D, allow_none=True, default=None)
    extLst: OfficeArtExtensionList | None = Field.element(expected_type=OfficeArtExtensionList, allow_none=True, default=None)
    xml_order = ("xfrm", "scene3d", "extLst")

    def __init__(self,
                 bwMode=None,
                 xfrm=None,
                 scene3d=None,
                 extLst=None,
                ):
        self.bwMode = bwMode
        self.xfrm = xfrm
        self.scene3d = scene3d
        self.extLst = extLst


class GroupLocking(TypedSerialisable):

    tagname = "grpSpLocks"
    namespace = DRAWING_NS

    noGrp: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    noUngrp: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    noSelect: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    noRot: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    noChangeAspect: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    noMove: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    noResize: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    noChangeArrowheads: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    noEditPoints: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    noAdjustHandles: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    noChangeShapeType: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    extLst: OfficeArtExtensionList | None = Field.element(expected_type=OfficeArtExtensionList, allow_none=True, default=None)
    xml_order = ()

    def __init__(self,
                 noGrp=None,
                 noUngrp=None,
                 noSelect=None,
                 noRot=None,
                 noChangeAspect=None,
                 noChangeArrowheads=None,
                 noMove=None,
                 noResize=None,
                 noEditPoints=None,
                 noAdjustHandles=None,
                 noChangeShapeType=None,
                 extLst=None,
                ):
        self.noGrp = noGrp
        self.noUngrp = noUngrp
        self.noSelect = noSelect
        self.noRot = noRot
        self.noChangeAspect = noChangeAspect
        self.noChangeArrowheads = noChangeArrowheads
        self.noMove = noMove
        self.noResize = noResize
        self.noEditPoints = noEditPoints
        self.noAdjustHandles = noAdjustHandles
        self.noChangeShapeType = noChangeShapeType
        self.extLst = extLst


class NonVisualGroupDrawingShapeProps(TypedSerialisable):

    tagname = "cNvGrpSpPr"

    grpSpLocks: GroupLocking | None = Field.element(expected_type=GroupLocking, allow_none=True, default=None)
    extLst: OfficeArtExtensionList | None = Field.element(expected_type=OfficeArtExtensionList, allow_none=True, default=None)
    xml_order = ("grpSpLocks", "extLst")

    def __init__(self,
                 grpSpLocks=None,
                 extLst=None,
                ):
        self.grpSpLocks = grpSpLocks
        self.extLst = extLst


class NonVisualDrawingShapeProps(TypedSerialisable):

    tagname = "cNvSpPr"

    spLocks: GroupLocking | None = Field.element(expected_type=GroupLocking, allow_none=True, default=None)
    txBax: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    extLst: OfficeArtExtensionList | None = Field.element(expected_type=OfficeArtExtensionList, allow_none=True, default=None)

    xml_order = ("spLocks",)

    def __init__(self,
                 spLocks=None,
                 txBox=None,
                 extLst=None,
                ):
        self.spLocks = spLocks
        self.txBax = txBox


class NonVisualDrawingProps(TypedSerialisable):

    tagname = "cNvPr"

    id: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    name: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    descr: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    hidden: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    title: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    hlinkClick: Hyperlink | None = Field.element(expected_type=Hyperlink, allow_none=True, default=None)
    hlinkHover: Hyperlink | None = Field.element(expected_type=Hyperlink, allow_none=True, default=None)
    extLst: OfficeArtExtensionList | None = Field.element(expected_type=OfficeArtExtensionList, allow_none=True, default=None)

    xml_order = ("hlinkClick", "hlinkHover", "extLst")

    def __init__(self,
                 id=None,
                 name=None,
                 descr=None,
                 hidden=None,
                 title=None,
                 hlinkClick=None,
                 hlinkHover=None,
                 extLst=None,
                ):
        self.id = id
        self.name = name
        self.descr = descr
        self.hidden = hidden
        self.title = title
        self.hlinkClick = hlinkClick
        self.hlinkHover = hlinkHover
        self.extLst = extLst

class NonVisualGroupShape(TypedSerialisable):

    tagname = "nvGrpSpPr"

    cNvPr: NonVisualDrawingProps | None = Field.element(expected_type=NonVisualDrawingProps, allow_none=True, default=None)
    cNvGrpSpPr: NonVisualGroupDrawingShapeProps | None = Field.element(
        expected_type=NonVisualGroupDrawingShapeProps,
        allow_none=True, default=None,
    )

    xml_order = ("cNvPr", "cNvGrpSpPr")

    def __init__(self,
                 cNvPr=None,
                 cNvGrpSpPr=None,
                ):
        self.cNvPr = cNvPr
        self.cNvGrpSpPr = cNvGrpSpPr


def _enum_converter(value, allowed_values, field_name: str):
    if value is None:
        return None
    if value not in allowed_values:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return value

