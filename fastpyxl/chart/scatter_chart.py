# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.descriptors.excel import ExtensionList

from ._chart import ChartBase
from .axis import NumericAxis
from .series import XYSeries
from .label import DataLabelList


def _scatter_style(v):
    if v is None or v == "none":
        return None
    allowed = frozenset({"line", "lineMarker", "marker", "smooth", "smoothMarker"})
    if v not in allowed:
        raise FieldValidationError(f"scatterStyle rejected value {v!r}")
    return v


class ScatterChart(ChartBase):
    tagname = "scatterChart"

    scatterStyle: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_scatter_style, default=None,
    )
    varyColors: bool | None = Field.nested_bool(allow_none=True, default=None)
    ser: list[XYSeries] | None = Field.sequence(expected_type=XYSeries, allow_none=True, default=list)
    dLbls: DataLabelList | None = Field.element(
        expected_type=DataLabelList, allow_none=True, default=None
    )
    dataLabels = AliasField("dLbls", default=None)
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False, default=None
    )

    _series_type = "scatter"

    xml_order = ("scatterStyle", "varyColors", "ser", "dLbls", "axId")

    def __init__(
        self,
        scatterStyle=None,
        varyColors=None,
        ser=(),
        dLbls=None,
        extLst=None,
        **kw,
    ):
        self.scatterStyle = scatterStyle
        self.varyColors = varyColors
        self.ser = list(ser) if ser is not None else []
        self.dLbls = dLbls
        self.extLst = extLst
        self.x_axis = NumericAxis(axId=10, crossAx=20)
        self.y_axis = NumericAxis(axId=20, crossAx=10)
        super().__init__(**kw)
