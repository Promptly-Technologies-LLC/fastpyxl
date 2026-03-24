# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.descriptors.excel import ExtensionList as OfficeArtExtensionList
from fastpyxl.chart.shapes import GraphicalProperties
from fastpyxl.chart.text import RichText
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import AliasField, Field

from .properties import (
    NonVisualDrawingProps,
    NonVisualDrawingShapeProps,
)
from .geometry import ShapeStyle

class Connection(Serialisable):

    id: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    idx: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

    def __init__(self,
                 id=None,
                 idx=None,
                ):
        self.id = id
        self.idx = idx


class ConnectorLocking(Serialisable):

    extLst: OfficeArtExtensionList | None = Field.element(expected_type=OfficeArtExtensionList, allow_none=True, default=None)

    def __init__(self,
                 extLst=None,
                ):
        self.extLst = extLst


class NonVisualConnectorProperties(Serialisable):

    cxnSpLocks: ConnectorLocking | None = Field.element(expected_type=ConnectorLocking, allow_none=True, default=None)
    stCxn: Connection | None = Field.element(expected_type=Connection, allow_none=True, default=None)
    endCxn: Connection | None = Field.element(expected_type=Connection, allow_none=True, default=None)
    extLst: OfficeArtExtensionList | None = Field.element(expected_type=OfficeArtExtensionList, allow_none=True, default=None)

    def __init__(self,
                 cxnSpLocks=None,
                 stCxn=None,
                 endCxn=None,
                 extLst=None,
                ):
        self.cxnSpLocks = cxnSpLocks
        self.stCxn = stCxn
        self.endCxn = endCxn
        self.extLst = extLst


class ConnectorNonVisual(Serialisable):

    cNvPr: NonVisualDrawingProps | None = Field.element(expected_type=NonVisualDrawingProps, allow_none=True, default=None)
    cNvCxnSpPr: NonVisualConnectorProperties | None = Field.element(
        expected_type=NonVisualConnectorProperties,
        allow_none=True, default=None,
    )

    xml_order = ("cNvPr", "cNvCxnSpPr")

    def __init__(self,
                 cNvPr=None,
                 cNvCxnSpPr=None,
                ):
        self.cNvPr = cNvPr
        self.cNvCxnSpPr = cNvCxnSpPr


class ConnectorShape(Serialisable):

    tagname = "cxnSp"

    nvCxnSpPr: ConnectorNonVisual | None = Field.element(expected_type=ConnectorNonVisual, allow_none=True, default=None)
    spPr: GraphicalProperties | None = Field.element(expected_type=GraphicalProperties, allow_none=True, default=None)
    style: ShapeStyle | None = Field.element(expected_type=ShapeStyle, allow_none=True, default=None)
    macro: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    fPublished: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)

    xml_order = ("nvCxnSpPr", "spPr", "style")

    def __init__(self,
                 nvCxnSpPr=None,
                 spPr=None,
                 style=None,
                 macro=None,
                 fPublished=None,
                 ):
        self.nvCxnSpPr = nvCxnSpPr
        self.spPr = spPr
        self.style = style
        self.macro = macro
        self.fPublished = fPublished


class ShapeMeta(Serialisable):

    tagname = "nvSpPr"

    cNvPr: NonVisualDrawingProps | None = Field.element(expected_type=NonVisualDrawingProps, allow_none=True, default=None)
    cNvSpPr: NonVisualDrawingShapeProps | None = Field.element(expected_type=NonVisualDrawingShapeProps, allow_none=True, default=None)

    xml_order = ("cNvPr", "cNvSpPr")

    def __init__(self, cNvPr=None, cNvSpPr=None):
        self.cNvPr = cNvPr
        self.cNvSpPr = cNvSpPr


class Shape(Serialisable):

    macro: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    textlink: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    fPublished: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    fLocksText: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    nvSpPr: ShapeMeta | None = Field.element(expected_type=ShapeMeta, allow_none=True, default=None)
    meta = AliasField("nvSpPr", default=None)
    spPr: GraphicalProperties | None = Field.element(expected_type=GraphicalProperties, allow_none=True, default=None)
    graphicalProperties = AliasField("spPr", default=None)
    style: ShapeStyle | None = Field.element(expected_type=ShapeStyle, allow_none=True, default=None)
    txBody: RichText | None = Field.element(expected_type=RichText, allow_none=True, default=None)

    xml_order = ("nvSpPr", "spPr", "style", "txBody")

    def __init__(self,
                 macro=None,
                 textlink=None,
                 fPublished=None,
                 fLocksText=None,
                 nvSpPr=None,
                 spPr=None,
                 style=None,
                 txBody=None,
                ):
        self.macro = macro
        self.textlink = textlink
        self.fPublished = fPublished
        self.fLocksText = fLocksText
        self.nvSpPr = nvSpPr
        self.spPr = spPr
        self.style = style
        self.txBody = txBody
