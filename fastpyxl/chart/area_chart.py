# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.descriptors.excel import ExtensionList

from ._chart import ChartBase
from .descriptors import NestedGapAmount
from .axis import TextAxis, NumericAxis, SeriesAxis, ChartLines
from .label import DataLabelList
from .series import Series


def _area_grouping(v):
    if v is None:
        return None
    allowed = frozenset({"percentStacked", "standard", "stacked"})
    if v not in allowed:
        raise FieldValidationError(f"grouping rejected value {v!r}")
    return v


class _AreaChartBase(ChartBase):
    grouping: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_area_grouping,
    )
    varyColors: bool | None = Field.nested_bool(allow_none=True)
    ser: list[Series] | None = Field.sequence(expected_type=Series, allow_none=True)
    dLbls: DataLabelList | None = Field.element(
        expected_type=DataLabelList, allow_none=True
    )
    dataLabels = AliasField("dLbls")
    dropLines: ChartLines | None = Field.element(expected_type=ChartLines, allow_none=True)

    _series_type = "area"

    xml_order = ("grouping", "varyColors", "ser", "dLbls", "dropLines")

    def __init__(
        self,
        grouping="standard",
        varyColors=None,
        ser=(),
        dLbls=None,
        dropLines=None,
        **kw,
    ):
        self.grouping = grouping
        self.varyColors = varyColors
        self.ser = list(ser) if ser is not None else []
        self.dLbls = dLbls
        self.dropLines = dropLines
        super().__init__(**kw)


class AreaChart(_AreaChartBase):
    tagname = "areaChart"

    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = _AreaChartBase.xml_order + ("axId",)

    def __init__(self, axId=None, extLst=None, **kw):
        del axId
        self.extLst = extLst
        self.x_axis = TextAxis()
        self.y_axis = NumericAxis()
        super().__init__(**kw)
        if not self.axId:
            self.axId = list(self._axes.keys())


class AreaChart3D(_AreaChartBase):
    tagname = "area3DChart"

    gapDepth = NestedGapAmount

    xml_order = _AreaChartBase.xml_order + ("gapDepth", "axId")

    def __init__(self, gapDepth=None, **kw):
        self.gapDepth = gapDepth
        super().__init__(**kw)
        self.x_axis = TextAxis()
        self.y_axis = NumericAxis()
        self.z_axis = SeriesAxis()
        if not self.axId:
            self.axId = list(self._axes.keys())
