# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.descriptors.excel import ExtensionList

from .data_source import NumFmt
from .descriptors import num_fmt_from_value
from .shapes import GraphicalProperties
from .text import RichText, Text
from .layout import Layout


def _trendline_type(v):
    if v is None:
        return None
    allowed = frozenset({"exp", "linear", "log", "movingAvg", "poly", "power"})
    if v not in allowed:
        raise FieldValidationError(f"trendlineType rejected value {v!r}")
    return v


class TrendlineLabel(Serialisable):
    tagname = "trendlineLbl"

    layout: Layout | None = Field.element(expected_type=Layout, allow_none=True)
    tx: Text | None = Field.element(expected_type=Text, allow_none=True)
    numFmt: NumFmt | None = Field.element(
        expected_type=NumFmt,
        allow_none=True,
        converter=num_fmt_from_value,
    )
    spPr: GraphicalProperties | None = Field.element(
        expected_type=GraphicalProperties, allow_none=True
    )
    graphicalProperties = AliasField("spPr")
    txPr: RichText | None = Field.element(expected_type=RichText, allow_none=True)
    textProperties = AliasField("txPr")
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = ("layout", "tx", "numFmt", "spPr", "txPr")

    def __init__(
        self,
        layout=None,
        tx=None,
        numFmt=None,
        spPr=None,
        txPr=None,
        extLst=None,
    ):
        self.layout = layout
        self.tx = tx
        self.numFmt = num_fmt_from_value(numFmt) if numFmt is not None else None
        self.spPr = spPr
        self.txPr = txPr
        self.extLst = extLst


class Trendline(Serialisable):
    tagname = "trendline"

    name: str | None = Field.attribute(expected_type=str, allow_none=True)
    spPr: GraphicalProperties | None = Field.element(
        expected_type=GraphicalProperties, allow_none=True
    )
    graphicalProperties = AliasField("spPr")
    trendlineType: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_trendline_type,
    )
    order: int | None = Field.nested_value(expected_type=int, allow_none=True)
    period: int | None = Field.nested_value(expected_type=int, allow_none=True)
    forward: float | None = Field.nested_value(expected_type=float, allow_none=True)
    backward: float | None = Field.nested_value(expected_type=float, allow_none=True)
    intercept: float | None = Field.nested_value(expected_type=float, allow_none=True)
    dispRSqr: bool | None = Field.nested_bool(allow_none=True)
    dispEq: bool | None = Field.nested_bool(allow_none=True)
    trendlineLbl: TrendlineLabel | None = Field.element(
        expected_type=TrendlineLabel, allow_none=True
    )
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = (
        "spPr",
        "trendlineType",
        "order",
        "period",
        "forward",
        "backward",
        "intercept",
        "dispRSqr",
        "dispEq",
        "trendlineLbl",
    )

    def __init__(
        self,
        name=None,
        spPr=None,
        trendlineType="linear",
        order=None,
        period=None,
        forward=None,
        backward=None,
        intercept=None,
        dispRSqr=None,
        dispEq=None,
        trendlineLbl=None,
        extLst=None,
    ):
        self.name = name
        self.spPr = spPr
        self.trendlineType = trendlineType
        self.order = order
        self.period = period
        self.forward = forward
        self.backward = backward
        self.intercept = intercept
        self.dispRSqr = dispRSqr
        self.dispEq = dispEq
        self.trendlineLbl = trendlineLbl
        self.extLst = extLst
