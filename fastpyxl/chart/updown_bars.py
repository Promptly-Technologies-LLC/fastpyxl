# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import Field

from fastpyxl.descriptors.excel import ExtensionList

from .axis import ChartLines
from .descriptors import NestedGapAmount


class UpDownBars(Serialisable):
    tagname = "upbars"

    gapWidth = NestedGapAmount
    upBars: ChartLines | None = Field.element(expected_type=ChartLines, allow_none=True, default=None)
    downBars: ChartLines | None = Field.element(expected_type=ChartLines, allow_none=True, default=None)
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False, default=None
    )

    xml_order = ("gapWidth", "upBars", "downBars")

    def __init__(
        self,
        gapWidth=150,
        upBars=None,
        downBars=None,
        extLst=None,
    ):
        self.gapWidth = gapWidth
        self.upBars = upBars
        self.downBars = downBars
        self.extLst = extLst
