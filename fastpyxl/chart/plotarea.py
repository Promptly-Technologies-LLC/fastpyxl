# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.descriptors.excel import ExtensionList

from fastpyxl.typed_serialisable.errors import FieldValidationError

from ._3d import _3DBase
from .area_chart import AreaChart, AreaChart3D
from .bar_chart import BarChart, BarChart3D
from .bubble_chart import BubbleChart
from .line_chart import LineChart, LineChart3D
from .pie_chart import PieChart, PieChart3D, ProjectedPieChart, DoughnutChart
from .radar_chart import RadarChart
from .scatter_chart import ScatterChart
from .stock_chart import StockChart
from .surface_chart import SurfaceChart, SurfaceChart3D
from .layout import Layout
from .shapes import GraphicalProperties
from .text import RichText

from .axis import (
    NumericAxis,
    TextAxis,
    SeriesAxis,
    DateAxis,
)


class DataTable(Serialisable):
    tagname = "dTable"

    showHorzBorder: bool | None = Field.nested_bool(allow_none=True)
    showVertBorder: bool | None = Field.nested_bool(allow_none=True)
    showOutline: bool | None = Field.nested_bool(allow_none=True)
    showKeys: bool | None = Field.nested_bool(allow_none=True)
    spPr: GraphicalProperties | None = Field.element(
        expected_type=GraphicalProperties, allow_none=True
    )
    graphicalProperties = AliasField("spPr")
    txPr: RichText | None = Field.element(expected_type=RichText, allow_none=True)
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = ("showHorzBorder", "showVertBorder", "showOutline", "showKeys", "spPr", "txPr")

    def __init__(
        self,
        showHorzBorder=None,
        showVertBorder=None,
        showOutline=None,
        showKeys=None,
        spPr=None,
        txPr=None,
        extLst=None,
    ):
        self.showHorzBorder = showHorzBorder
        self.showVertBorder = showVertBorder
        self.showOutline = showOutline
        self.showKeys = showKeys
        self.spPr = spPr
        self.txPr = txPr
        self.extLst = extLst


_CHART_PARTS = {
    "areaChart": AreaChart,
    "area3DChart": AreaChart3D,
    "lineChart": LineChart,
    "line3DChart": LineChart3D,
    "stockChart": StockChart,
    "radarChart": RadarChart,
    "scatterChart": ScatterChart,
    "pieChart": PieChart,
    "pie3DChart": PieChart3D,
    "doughnutChart": DoughnutChart,
    "barChart": BarChart,
    "bar3DChart": BarChart3D,
    "ofPieChart": ProjectedPieChart,
    "surfaceChart": SurfaceChart,
    "surface3DChart": SurfaceChart3D,
    "bubbleChart": BubbleChart,
}

_AX_PARTS = {
    "valAx": NumericAxis,
    "catAx": TextAxis,
    "dateAx": DateAxis,
    "serAx": SeriesAxis,
}

_CHART_ASSIGN = frozenset(_CHART_PARTS)
_AX_ASSIGN = frozenset(_AX_PARTS)


class PlotArea(Serialisable):
    tagname = "plotArea"

    layout: Layout | None = Field.element(expected_type=Layout, allow_none=True)
    dTable: DataTable | None = Field.element(expected_type=DataTable, allow_none=True)
    spPr: GraphicalProperties | None = Field.element(
        expected_type=GraphicalProperties, allow_none=True
    )
    graphicalProperties = AliasField("spPr")
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    _charts: list | None = Field.multi_sequence(
        parts=_CHART_PARTS,
        allow_none=True,
        default=list,
    )
    _axes: list | None = Field.multi_sequence(
        parts=_AX_PARTS,
        allow_none=True,
        default=list,
    )

    xml_order = ("layout", "_charts", "_axes", "dTable", "spPr")

    def __init__(
        self,
        layout=None,
        dTable=None,
        spPr=None,
        _charts=(),
        _axes=(),
        extLst=None,
    ):
        self.layout = layout
        self.dTable = dTable
        self.spPr = spPr
        self._charts = list(_charts) if _charts is not None else []
        self._axes = list(_axes) if _axes is not None else []
        self.extLst = extLst

    def __setattr__(self, name, value):
        if name in ("_charts", "_axes"):
            object.__setattr__(self, name, value)
            return
        if name in _CHART_ASSIGN:
            if value is None:
                return
            if not isinstance(value, _CHART_PARTS[name]):
                raise FieldValidationError(f"{name} rejected value {value!r}")
            cur = list(object.__getattribute__(self, "_charts"))
            cur.append(value)
            object.__setattr__(self, "_charts", cur)
            return
        if name in _AX_ASSIGN:
            if value is None:
                return
            if not isinstance(value, _AX_PARTS[name]):
                raise FieldValidationError(f"{name} rejected value {value!r}")
            cur = list(object.__getattribute__(self, "_axes"))
            cur.append(value)
            object.__setattr__(self, "_axes", cur)
            return
        super().__setattr__(name, value)

    def to_tree(self, tagname=None, idx=None, namespace=None):
        del tagname, idx, namespace
        ax_ids = {ax.axId for ax in self._axes}
        for chart in self._charts:
            for _id, axis in chart._axes.items():
                if _id not in ax_ids:
                    setattr(self, axis.tagname, axis)
                    ax_ids.add(_id)
        return super().to_tree()

    @classmethod
    def from_tree(cls, node):
        self = super().from_tree(node)
        axes = dict((axis.axId, axis) for axis in self._axes)
        for chart in self._charts:
            if isinstance(chart, (ScatterChart, BubbleChart)):
                x, y = (axes[ax_id] for ax_id in chart.axId)
                chart.x_axis = x
                chart.y_axis = y
                continue

            for ax_id in chart.axId:
                axis = axes.get(ax_id)
                if axis is None and isinstance(chart, _3DBase):
                    chart.z_axis = None
                    continue
                if axis.tagname in ("catAx", "dateAx"):
                    chart.x_axis = axis
                elif axis.tagname == "valAx":
                    chart.y_axis = axis
                elif axis.tagname == "serAx":
                    chart.z_axis = axis

        return self
