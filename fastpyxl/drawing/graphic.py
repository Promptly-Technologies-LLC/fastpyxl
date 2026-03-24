# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.xml.constants import CHART_NS, DRAWING_NS, SHEET_DRAWING_NS
from fastpyxl.descriptors.excel import ExtensionList as OfficeArtExtensionList
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import AliasField, Field

from .picture import PictureFrame
from .properties import (
    NonVisualDrawingProps,
    NonVisualGroupShape,
    GroupShapeProperties,
)
from .relation import ChartRelation
from .xdr import XDRTransform2D


class GraphicFrameLocking(Serialisable):

    noGrp: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    noDrilldown: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    noSelect: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    noChangeAspect: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    noMove: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    noResize: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    extLst: OfficeArtExtensionList | None = Field.element(expected_type=OfficeArtExtensionList, allow_none=True, default=None)

    def __init__(self,
                 noGrp=None,
                 noDrilldown=None,
                 noSelect=None,
                 noChangeAspect=None,
                 noMove=None,
                 noResize=None,
                 extLst=None,
                ):
        self.noGrp = noGrp
        self.noDrilldown = noDrilldown
        self.noSelect = noSelect
        self.noChangeAspect = noChangeAspect
        self.noMove = noMove
        self.noResize = noResize
        self.extLst = extLst


class NonVisualGraphicFrameProperties(Serialisable):

    tagname = "cNvGraphicFramePr"

    graphicFrameLocks: GraphicFrameLocking | None = Field.element(expected_type=GraphicFrameLocking, allow_none=True, default=None)
    extLst: OfficeArtExtensionList | None = Field.element(expected_type=OfficeArtExtensionList, allow_none=True, default=None)

    def __init__(self,
                 graphicFrameLocks=None,
                 extLst=None,
                ):
        self.graphicFrameLocks = graphicFrameLocks
        self.extLst = extLst


class NonVisualGraphicFrame(Serialisable):

    tagname = "nvGraphicFramePr"

    cNvPr: NonVisualDrawingProps | None = Field.element(expected_type=NonVisualDrawingProps, allow_none=True, default=None)
    cNvGraphicFramePr: NonVisualGraphicFrameProperties | None = Field.element(
        expected_type=NonVisualGraphicFrameProperties,
        allow_none=True, default=None,
    )

    xml_order = ('cNvPr', 'cNvGraphicFramePr')

    def __init__(self,
                 cNvPr=None,
                 cNvGraphicFramePr=None,
                ):
        if cNvPr is None:
            cNvPr = NonVisualDrawingProps(id=0, name="Chart 0")
        self.cNvPr = cNvPr
        if cNvGraphicFramePr is None:
            cNvGraphicFramePr = NonVisualGraphicFrameProperties()
        self.cNvGraphicFramePr = cNvGraphicFramePr


class GraphicData(Serialisable):

    tagname = "graphicData"
    namespace = DRAWING_NS

    uri: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    chart: ChartRelation | None = Field.element(expected_type=ChartRelation, allow_none=True, default=None)


    def __init__(self,
                 uri=CHART_NS,
                 chart=None,
                ):
        self.uri = uri
        self.chart = chart


class GraphicObject(Serialisable):

    tagname = "graphic"
    namespace = DRAWING_NS

    graphicData: GraphicData | None = Field.element(expected_type=GraphicData, allow_none=True, default=None)

    def __init__(self,
                 graphicData=None,
                ):
        if graphicData is None:
            graphicData = GraphicData()
        self.graphicData = graphicData


class GraphicFrame(Serialisable):

    tagname = "graphicFrame"

    nvGraphicFramePr: NonVisualGraphicFrame | None = Field.element(expected_type=NonVisualGraphicFrame, allow_none=True, default=None)
    xfrm: XDRTransform2D | None = Field.element(expected_type=XDRTransform2D, allow_none=True, default=None)
    graphic: GraphicObject | None = Field.element(expected_type=GraphicObject, allow_none=True, default=None)
    macro: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    fPublished: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)

    xml_order = ('nvGraphicFramePr', 'xfrm', 'graphic')

    def __init__(self,
                 nvGraphicFramePr=None,
                 xfrm=None,
                 graphic=None,
                 macro=None,
                 fPublished=None,
                 ):
        if nvGraphicFramePr is None:
            nvGraphicFramePr = NonVisualGraphicFrame()
        self.nvGraphicFramePr = nvGraphicFramePr
        if xfrm is None:
            xfrm = XDRTransform2D()
        self.xfrm = xfrm
        if graphic is None:
            graphic = GraphicObject()
        self.graphic = graphic
        self.macro = macro
        self.fPublished = fPublished


class GroupShape(Serialisable):

    tagname = "grpSp"
    namespace = SHEET_DRAWING_NS

    nvGrpSpPr: NonVisualGroupShape | None = Field.element(expected_type=NonVisualGroupShape, allow_none=True, default=None)
    nonVisualProperties = AliasField("nvGrpSpPr", default=None)
    grpSpPr: GroupShapeProperties | None = Field.element(expected_type=GroupShapeProperties, allow_none=True, default=None)
    visualProperties = AliasField("grpSpPr", default=None)
    pic: PictureFrame | None = Field.element(expected_type=PictureFrame, allow_none=True, default=None)

    xml_order = ("nvGrpSpPr", "grpSpPr", "pic")

    def __init__(self,
                 nvGrpSpPr=None,
                 grpSpPr=None,
                 pic=None,
                ):
        self.nvGrpSpPr = nvGrpSpPr
        self.grpSpPr = grpSpPr
        self.pic = pic
