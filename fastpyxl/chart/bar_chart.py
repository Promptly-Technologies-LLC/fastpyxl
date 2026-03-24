# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.descriptors.excel import ExtensionList

from .descriptors import NestedGapAmount, NestedOverlap
from ._chart import ChartBase
from ._3d import (
    FIELD_BACK_WALL_ON_CHART,
    FIELD_FLOOR_ON_CHART,
    FIELD_SIDE_WALL_ON_CHART,
    FIELD_VIEW3D_ON_CHART,
    _3DBase,
)
from .axis import TextAxis, NumericAxis, SeriesAxis, ChartLines
from .series import Series, _shape_converter
from .legend import Legend
from .label import DataLabelList


def _bar_dir(v):
    if v is None:
        return None
    if v not in ("bar", "col"):
        raise FieldValidationError(f"barDir rejected value {v!r}")
    return v


def _bar_grouping(v):
    if v is None:
        return None
    allowed = frozenset({"percentStacked", "clustered", "standard", "stacked"})
    if v not in allowed:
        raise FieldValidationError(f"grouping rejected value {v!r}")
    return v


class _BarChartBase(ChartBase):
    barDir: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_bar_dir,
    )
    type = AliasField("barDir")
    grouping: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_bar_grouping,
    )
    varyColors: bool | None = Field.nested_bool(allow_none=True)
    ser: list[Series] | None = Field.sequence(expected_type=Series, allow_none=True)
    dLbls: DataLabelList | None = Field.element(
        expected_type=DataLabelList, allow_none=True
    )
    dataLabels = AliasField("dLbls")

    xml_order = ("barDir", "grouping", "varyColors", "ser", "dLbls")

    _series_type = "bar"

    def __init__(
        self,
        barDir="col",
        grouping="clustered",
        varyColors=None,
        ser=(),
        dLbls=None,
        **kw,
    ):
        self.barDir = barDir
        self.grouping = grouping
        self.varyColors = varyColors
        self.ser = list(ser) if ser is not None else []
        self.dLbls = dLbls
        super().__init__(**kw)


class BarChart(_BarChartBase):
    tagname = "barChart"

    gapWidth = NestedGapAmount
    overlap = NestedOverlap
    serLines: ChartLines | None = Field.element(expected_type=ChartLines, allow_none=True)
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = _BarChartBase.xml_order + ("gapWidth", "overlap", "serLines", "axId")

    def __init__(
        self,
        gapWidth=150,
        overlap=None,
        serLines=None,
        extLst=None,
        **kw,
    ):
        self.gapWidth = gapWidth
        self.overlap = overlap
        self.serLines = serLines
        self.extLst = extLst
        self.x_axis = TextAxis()
        self.y_axis = NumericAxis()
        self.legend = Legend()
        super().__init__(**kw)
        if not self.axId:
            self.axId = list(self._axes.keys())


class BarChart3D(_BarChartBase, _3DBase):
    tagname = "bar3DChart"

    view3D = FIELD_VIEW3D_ON_CHART
    floor = FIELD_FLOOR_ON_CHART
    sideWall = FIELD_SIDE_WALL_ON_CHART
    backWall = FIELD_BACK_WALL_ON_CHART

    gapWidth = NestedGapAmount
    gapDepth = NestedGapAmount
    shape: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_shape_converter,
    )
    serLines: ChartLines | None = Field.element(expected_type=ChartLines, allow_none=True)
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = _BarChartBase.xml_order + (
        "gapWidth",
        "gapDepth",
        "shape",
        "serLines",
        "axId",
    )

    def __init__(
        self,
        gapWidth=150,
        gapDepth=150,
        shape=None,
        serLines=None,
        extLst=None,
        **kw,
    ):
        self.gapWidth = gapWidth
        self.gapDepth = gapDepth
        self.shape = shape
        self.serLines = serLines
        self.extLst = extLst
        self.x_axis = TextAxis()
        self.y_axis = NumericAxis()
        self.z_axis = SeriesAxis()
        _BarChartBase.__init__(self, **kw)
        _3DBase.__init__(self)
