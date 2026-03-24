# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.descriptors.excel import ExtensionList

from ._chart import ChartBase
from .axis import TextAxis, NumericAxis
from .series import Series
from .label import DataLabelList


def _radar_style(v):
    if v is None:
        return None
    if v not in ("standard", "marker", "filled"):
        raise FieldValidationError(f"radarStyle rejected value {v!r}")
    return v


class RadarChart(ChartBase):
    tagname = "radarChart"

    radarStyle: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_radar_style, default=None,
    )
    type = AliasField("radarStyle", default=None)
    varyColors: bool | None = Field.nested_bool(allow_none=True, default=None)
    ser: list[Series] | None = Field.sequence(expected_type=Series, allow_none=True, default=list)
    dLbls: DataLabelList | None = Field.element(
        expected_type=DataLabelList, allow_none=True, default=None
    )
    dataLabels = AliasField("dLbls", default=None)
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False, default=None
    )

    _series_type = "radar"

    xml_order = ("radarStyle", "varyColors", "ser", "dLbls", "axId")

    def __init__(
        self,
        radarStyle="standard",
        varyColors=None,
        ser=(),
        dLbls=None,
        extLst=None,
        **kw,
    ):
        self.radarStyle = radarStyle
        self.varyColors = varyColors
        self.ser = list(ser) if ser is not None else []
        self.dLbls = dLbls
        self.extLst = extLst
        self.x_axis = TextAxis()
        self.y_axis = NumericAxis()
        super().__init__(**kw)
