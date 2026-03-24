# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.xml.constants import DRAWING_NS

from fastpyxl.descriptors.excel import ExtensionList as OfficeArtExtensionList
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.chart.shapes import GraphicalProperties

from .fill import BlipFillProperties
from .properties import NonVisualDrawingProps
from .geometry import ShapeStyle


class PictureLocking(Serialisable):

    tagname = "picLocks"
    namespace = DRAWING_NS

    # Using attribute group AG_Locking
    noCrop: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    noGrp: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    noSelect: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    noRot: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    noChangeAspect: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    noMove: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    noResize: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    noEditPoints: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    noAdjustHandles: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    noChangeArrowheads: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    noChangeShapeType: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    extLst: OfficeArtExtensionList | None = Field.element(expected_type=OfficeArtExtensionList, allow_none=True, default=None)

    def __init__(self,
                 noCrop=None,
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
                ):
        self.noCrop = noCrop
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


class NonVisualPictureProperties(Serialisable):

    tagname = "cNvPicPr"

    preferRelativeResize: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    picLocks: PictureLocking | None = Field.element(expected_type=PictureLocking, allow_none=True, default=None)
    extLst: OfficeArtExtensionList | None = Field.element(expected_type=OfficeArtExtensionList, allow_none=True, default=None)

    xml_order = ("picLocks",)

    def __init__(self,
                 preferRelativeResize=None,
                 picLocks=None,
                 extLst=None,
                ):
        self.preferRelativeResize = preferRelativeResize
        self.picLocks = picLocks
        self.extLst = extLst


class PictureNonVisual(Serialisable):

    tagname = "nvPicPr"

    cNvPr: NonVisualDrawingProps | None = Field.element(expected_type=NonVisualDrawingProps, allow_none=True, default=None)
    cNvPicPr: NonVisualPictureProperties | None = Field.element(expected_type=NonVisualPictureProperties, allow_none=True, default=None)

    xml_order = ("cNvPr", "cNvPicPr")

    def __init__(self,
                 cNvPr=None,
                 cNvPicPr=None,
                ):
        if cNvPr is None:
            cNvPr = NonVisualDrawingProps(id=0, name="Image 1", descr="Name of file")
        self.cNvPr = cNvPr
        if cNvPicPr is None:
            cNvPicPr = NonVisualPictureProperties()
        self.cNvPicPr = cNvPicPr




class PictureFrame(Serialisable):

    tagname = "pic"

    macro: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    fPublished: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    nvPicPr: PictureNonVisual | None = Field.element(expected_type=PictureNonVisual, allow_none=True, default=None)
    blipFill: BlipFillProperties | None = Field.element(expected_type=BlipFillProperties, allow_none=True, default=None)
    spPr: GraphicalProperties | None = Field.element(expected_type=GraphicalProperties, allow_none=True, default=None)
    graphicalProperties = AliasField('spPr', default=None)
    style: ShapeStyle | None = Field.element(expected_type=ShapeStyle, allow_none=True, default=None)

    xml_order = ("nvPicPr", "blipFill", "spPr", "style")

    def __init__(self,
                 macro=None,
                 fPublished=None,
                 nvPicPr=None,
                 blipFill=None,
                 spPr=None,
                 style=None,
                ):
        self.macro = macro
        self.fPublished = fPublished
        if nvPicPr is None:
            nvPicPr = PictureNonVisual()
        self.nvPicPr = nvPicPr
        if blipFill is None:
            blipFill = BlipFillProperties()
        self.blipFill = blipFill
        if spPr is None:
            spPr = GraphicalProperties()
        self.spPr = spPr
        self.style = style
