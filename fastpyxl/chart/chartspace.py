
# Copyright (c) 2010-2024 fastpyxl

"""
Enclosing chart object. The various chart types are actually child objects.
Will probably need to call this indirectly
"""

from __future__ import annotations

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.descriptors.excel import ExtensionList
from fastpyxl.xml.constants import CHART_NS, REL_NS

from fastpyxl.drawing.colors import ColorMapping
from .text import RichText
from .shapes import GraphicalProperties
from .legend import Legend
from ._3d import (
    FIELD_BACK_WALL,
    FIELD_FLOOR,
    FIELD_SIDE_WALL,
    FIELD_VIEW3D,
)
from .plotarea import PlotArea
from .title import Title, title_from_value
from .pivot import (
    PivotFormat,
    PivotSource,
)
from .print_settings import PrintSettings


def _disp_blanks(v):
    if v is None or v == "none":
        return None
    allowed = frozenset({"span", "gap", "zero"})
    if v not in allowed:
        raise FieldValidationError(f"dispBlanksAs rejected value {v!r}")
    return v


def _chart_style(v):
    if v is None:
        return None
    try:
        n = int(v)
    except (TypeError, ValueError) as exc:
        raise FieldValidationError(f"style rejected value {v!r}") from exc
    if n < 1 or n > 48:
        raise FieldValidationError(f"style rejected value {v!r}")
    return n


class ChartContainer(Serialisable):
    tagname = "chart"

    title: Title | None = Field.element(
        expected_type=Title,
        allow_none=True,
        converter=title_from_value,
    )
    autoTitleDeleted: bool | None = Field.nested_bool(allow_none=True)
    pivotFmts: list[PivotFormat] | None = Field.nested_sequence(
        expected_type=PivotFormat,
        allow_none=True,
    )
    view3D = FIELD_VIEW3D
    floor = FIELD_FLOOR
    sideWall = FIELD_SIDE_WALL
    backWall = FIELD_BACK_WALL
    plotArea: PlotArea | None = Field.element(expected_type=PlotArea, allow_none=True)
    legend: Legend | None = Field.element(expected_type=Legend, allow_none=True)
    plotVisOnly: bool | None = Field.nested_bool(allow_none=True, default=True)
    dispBlanksAs: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_disp_blanks,
    )
    showDLblsOverMax: bool | None = Field.nested_bool(allow_none=True)
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = (
        "title",
        "autoTitleDeleted",
        "pivotFmts",
        "view3D",
        "floor",
        "sideWall",
        "backWall",
        "plotArea",
        "legend",
        "plotVisOnly",
        "dispBlanksAs",
        "showDLblsOverMax",
    )

    def __init__(
        self,
        title=None,
        autoTitleDeleted=None,
        pivotFmts=(),
        view3D=None,
        floor=None,
        sideWall=None,
        backWall=None,
        plotArea=None,
        legend=None,
        plotVisOnly=True,
        dispBlanksAs="gap",
        showDLblsOverMax=None,
        extLst=None,
    ):
        self.title = title_from_value(title) if title is not None else None
        self.autoTitleDeleted = autoTitleDeleted
        self.pivotFmts = list(pivotFmts) if pivotFmts is not None else []
        self.view3D = view3D
        self.floor = floor
        self.sideWall = sideWall
        self.backWall = backWall
        if plotArea is None:
            plotArea = PlotArea()
        self.plotArea = plotArea
        self.legend = legend
        self.plotVisOnly = plotVisOnly
        self.dispBlanksAs = dispBlanksAs
        self.showDLblsOverMax = showDLblsOverMax
        self.extLst = extLst


class Protection(Serialisable):
    tagname = "protection"

    chartObject: bool | None = Field.nested_bool(allow_none=True)
    data: bool | None = Field.nested_bool(allow_none=True)
    formatting: bool | None = Field.nested_bool(allow_none=True)
    selection: bool | None = Field.nested_bool(allow_none=True)
    userInterface: bool | None = Field.nested_bool(allow_none=True)

    xml_order = (
        "chartObject",
        "data",
        "formatting",
        "selection",
        "userInterface",
    )

    def __init__(
        self,
        chartObject=None,
        data=None,
        formatting=None,
        selection=None,
        userInterface=None,
    ):
        self.chartObject = chartObject
        self.data = data
        self.formatting = formatting
        self.selection = selection
        self.userInterface = userInterface


class ExternalData(Serialisable):
    tagname = "externalData"

    autoUpdate: bool | None = Field.nested_bool(allow_none=True)
    id: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        namespace=REL_NS,
    )

    def __init__(self, autoUpdate=None, id=None):
        self.autoUpdate = autoUpdate
        self.id = id


class ChartSpace(Serialisable):
    tagname = "chartSpace"

    date1904: bool | None = Field.nested_bool(allow_none=True)
    lang: str | None = Field.nested_text(expected_type=str, allow_none=True)
    roundedCorners: bool | None = Field.nested_bool(allow_none=True)
    style: int | None = Field.nested_value(
        expected_type=int,
        allow_none=True,
        converter=_chart_style,
    )
    clrMapOvr: ColorMapping | None = Field.element(
        expected_type=ColorMapping, allow_none=True
    )
    pivotSource: PivotSource | None = Field.element(
        expected_type=PivotSource, allow_none=True
    )
    protection: Protection | None = Field.element(
        expected_type=Protection, allow_none=True
    )
    chart: ChartContainer | None = Field.element(expected_type=ChartContainer)
    spPr: GraphicalProperties | None = Field.element(
        expected_type=GraphicalProperties, allow_none=True
    )
    graphical_properties = AliasField("spPr")
    txPr: RichText | None = Field.element(expected_type=RichText, allow_none=True)
    textProperties = AliasField("txPr")
    externalData: ExternalData | None = Field.element(
        expected_type=ExternalData, allow_none=True
    )
    printSettings: PrintSettings | None = Field.element(
        expected_type=PrintSettings, allow_none=True
    )
    userShapes: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        namespace=REL_NS,
    )
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = (
        "date1904",
        "lang",
        "roundedCorners",
        "style",
        "clrMapOvr",
        "pivotSource",
        "protection",
        "chart",
        "spPr",
        "txPr",
        "externalData",
        "printSettings",
        "userShapes",
    )

    def __init__(
        self,
        date1904=None,
        lang=None,
        roundedCorners=None,
        style=None,
        clrMapOvr=None,
        pivotSource=None,
        protection=None,
        chart=None,
        spPr=None,
        txPr=None,
        externalData=None,
        printSettings=None,
        userShapes=None,
        extLst=None,
    ):
        self.date1904 = date1904
        self.lang = lang
        self.roundedCorners = roundedCorners
        self.style = style
        self.clrMapOvr = clrMapOvr
        self.pivotSource = pivotSource
        self.protection = protection
        self.chart = chart
        self.spPr = spPr
        self.txPr = txPr
        self.externalData = externalData
        self.printSettings = printSettings
        self.userShapes = userShapes
        self.extLst = extLst

    def to_tree(self, tagname=None, idx=None, namespace=None):
        del tagname, idx, namespace
        tree = super().to_tree()
        tree.set("xmlns", CHART_NS)
        return tree
