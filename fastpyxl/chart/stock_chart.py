# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.descriptors.excel import ExtensionList

from ._chart import ChartBase
from .axis import TextAxis, NumericAxis, ChartLines
from .updown_bars import UpDownBars
from .label import DataLabelList
from .series import Series


class StockChart(ChartBase):
    tagname = "stockChart"

    ser: list[Series] | None = Field.sequence(expected_type=Series, allow_none=True)
    dLbls: DataLabelList | None = Field.element(
        expected_type=DataLabelList, allow_none=True
    )
    dataLabels = AliasField("dLbls")
    dropLines: ChartLines | None = Field.element(expected_type=ChartLines, allow_none=True)
    hiLowLines: ChartLines | None = Field.element(expected_type=ChartLines, allow_none=True)
    upDownBars: UpDownBars | None = Field.element(expected_type=UpDownBars, allow_none=True)
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    _series_type = "line"

    xml_order = ("ser", "dLbls", "dropLines", "hiLowLines", "upDownBars", "axId")

    def __init__(
        self,
        ser=(),
        dLbls=None,
        dropLines=None,
        hiLowLines=None,
        upDownBars=None,
        extLst=None,
        **kw,
    ):
        self.ser = list(ser) if ser is not None else []
        self.dLbls = dLbls
        self.dropLines = dropLines
        self.hiLowLines = hiLowLines
        self.upDownBars = upDownBars
        self.extLst = extLst
        self.x_axis = TextAxis()
        self.y_axis = NumericAxis()
        super().__init__(**kw)
