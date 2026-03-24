# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.descriptors.excel import ExtensionList, _explicit_none
from fastpyxl.xml.constants import CHART_NS

from .data_source import NumFmt
from .descriptors import num_fmt_from_value
from .layout import Layout
from .text import Text, RichText
from .shapes import GraphicalProperties
from .title import Title, title_from_value


def _coerce_tick_mark(v):
    if v is None or v == "none":
        return None
    allowed = frozenset({"cross", "in", "out"})
    if v not in allowed:
        raise FieldValidationError(f"tick mark rejected value {v!r}")
    return v


def _nested_set_only(allowed: frozenset, field_name: str):
    def _c(v):
        if v is None:
            return None
        if v not in allowed:
            raise FieldValidationError(f"{field_name} rejected value {v!r}")
        return v

    return _c


def _none_set(allowed: frozenset, field_name: str):
    def _c(v):
        if v is None or v == "none":
            return None
        if v not in allowed:
            raise FieldValidationError(f"{field_name} rejected value {v!r}")
        return v

    return _c


def _chart_lbl_offset_minmax(value, *, field_name: str, min_v: int, max_v: int):
    if value is None:
        return None
    try:
        n = int(value)
    except (TypeError, ValueError) as exc:
        raise FieldValidationError(f"{field_name} rejected value {value!r}") from exc
    if n < min_v or n > max_v:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return n


class ChartLines(Serialisable):
    tagname = "chartLines"

    spPr: GraphicalProperties | None = Field.element(
        expected_type=GraphicalProperties, allow_none=True
    )
    graphicalProperties = AliasField("spPr")

    def __init__(self, spPr=None):
        self.spPr = spPr


class Scaling(Serialisable):
    tagname = "scaling"

    logBase: float | None = Field.nested_value(expected_type=float, allow_none=True)
    orientation: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_nested_set_only(frozenset({"maxMin", "minMax"}), "orientation"),
    )
    max: float | None = Field.nested_value(expected_type=float, allow_none=True)
    min: float | None = Field.nested_value(expected_type=float, allow_none=True)
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = ("logBase", "orientation", "max", "min")

    def __init__(
        self,
        logBase=None,
        orientation="minMax",
        max=None,
        min=None,
        extLst=None,
    ):
        self.logBase = logBase
        self.orientation = orientation
        self.max = max
        self.min = min
        self.extLst = extLst


class _BaseAxis(Serialisable):
    axId: int | None = Field.nested_value(expected_type=int, allow_none=True)
    scaling: Scaling | None = Field.element(expected_type=Scaling, allow_none=True)
    delete: bool | None = Field.nested_bool(allow_none=True)
    axPos: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_nested_set_only(frozenset({"b", "l", "r", "t"}), "axPos"),
    )
    majorGridlines: ChartLines | None = Field.element(
        expected_type=ChartLines, allow_none=True
    )
    minorGridlines: ChartLines | None = Field.element(
        expected_type=ChartLines, allow_none=True
    )
    title: Title | None = Field.element(
        expected_type=Title,
        allow_none=True,
        converter=title_from_value,
    )
    numFmt: NumFmt | None = Field.element(
        expected_type=NumFmt,
        allow_none=True,
        converter=num_fmt_from_value,
    )
    number_format = AliasField("numFmt")
    majorTickMark: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_coerce_tick_mark,
        renderer=_explicit_none,
    )
    minorTickMark: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_coerce_tick_mark,
        renderer=_explicit_none,
    )
    tickLblPos: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_none_set(frozenset({"high", "low", "nextTo"}), "tickLblPos"),
    )
    spPr: GraphicalProperties | None = Field.element(
        expected_type=GraphicalProperties, allow_none=True
    )
    graphicalProperties = AliasField("spPr")
    txPr: RichText | None = Field.element(expected_type=RichText, allow_none=True)
    textProperties = AliasField("txPr")
    crossAx: int | None = Field.nested_value(expected_type=int, allow_none=True)
    crosses: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_none_set(
            frozenset({"autoZero", "max", "min"}),
            "crosses",
        ),
    )
    crossesAt: float | None = Field.nested_value(expected_type=float, allow_none=True)

    xml_order = (
        "axId",
        "scaling",
        "delete",
        "axPos",
        "majorGridlines",
        "minorGridlines",
        "title",
        "numFmt",
        "majorTickMark",
        "minorTickMark",
        "tickLblPos",
        "spPr",
        "txPr",
        "crossAx",
        "crosses",
        "crossesAt",
    )

    def __init__(
        self,
        axId=None,
        scaling=None,
        delete=None,
        axPos="l",
        majorGridlines=None,
        minorGridlines=None,
        title=None,
        numFmt=None,
        majorTickMark=None,
        minorTickMark=None,
        tickLblPos=None,
        spPr=None,
        txPr=None,
        crossAx=None,
        crosses=None,
        crossesAt=None,
    ):
        self.axId = axId
        if scaling is None:
            scaling = Scaling()
        self.scaling = scaling
        self.delete = delete
        self.axPos = axPos
        self.majorGridlines = majorGridlines
        self.minorGridlines = minorGridlines
        self.title = title_from_value(title) if title is not None else None
        self.numFmt = num_fmt_from_value(numFmt) if numFmt is not None else None
        self.majorTickMark = majorTickMark
        self.minorTickMark = minorTickMark
        self.tickLblPos = tickLblPos
        self.spPr = spPr
        self.txPr = txPr
        self.crossAx = crossAx
        self.crosses = crosses
        self.crossesAt = crossesAt


class DisplayUnitsLabel(Serialisable):
    tagname = "dispUnitsLbl"

    layout: Layout | None = Field.element(expected_type=Layout, allow_none=True)
    tx: Text | None = Field.element(expected_type=Text, allow_none=True)
    text = AliasField("tx")
    spPr: GraphicalProperties | None = Field.element(
        expected_type=GraphicalProperties, allow_none=True
    )
    graphicalProperties = AliasField("spPr")
    txPr: RichText | None = Field.element(expected_type=RichText, allow_none=True)
    textPropertes = AliasField("txPr")

    xml_order = ("layout", "tx", "spPr", "txPr")

    def __init__(
        self,
        layout=None,
        tx=None,
        spPr=None,
        txPr=None,
    ):
        self.layout = layout
        self.tx = tx
        self.spPr = spPr
        self.txPr = txPr


class DisplayUnitsLabelList(Serialisable):
    tagname = "dispUnits"

    custUnit: float | None = Field.nested_value(expected_type=float, allow_none=True)
    builtInUnit: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_none_set(
            frozenset(
                {
                    "hundreds",
                    "thousands",
                    "tenThousands",
                    "hundredThousands",
                    "millions",
                    "tenMillions",
                    "hundredMillions",
                    "billions",
                    "trillions",
                }
            ),
            "builtInUnit",
        ),
    )
    dispUnitsLbl: DisplayUnitsLabel | None = Field.element(
        expected_type=DisplayUnitsLabel, allow_none=True
    )
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = ("custUnit", "builtInUnit", "dispUnitsLbl")

    def __init__(
        self,
        custUnit=None,
        builtInUnit=None,
        dispUnitsLbl=None,
        extLst=None,
    ):
        self.custUnit = custUnit
        self.builtInUnit = builtInUnit
        self.dispUnitsLbl = dispUnitsLbl
        self.extLst = extLst


class NumericAxis(_BaseAxis):
    tagname = "valAx"

    crossBetween: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_none_set(frozenset({"between", "midCat"}), "crossBetween"),
    )
    majorUnit: float | None = Field.nested_value(expected_type=float, allow_none=True)
    minorUnit: float | None = Field.nested_value(expected_type=float, allow_none=True)
    dispUnits: DisplayUnitsLabelList | None = Field.element(
        expected_type=DisplayUnitsLabelList, allow_none=True
    )
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = _BaseAxis.xml_order + (
        "crossBetween",
        "majorUnit",
        "minorUnit",
        "dispUnits",
    )

    def __init__(
        self,
        crossBetween=None,
        majorUnit=None,
        minorUnit=None,
        dispUnits=None,
        extLst=None,
        **kw,
    ):
        self.crossBetween = crossBetween
        self.majorUnit = majorUnit
        self.minorUnit = minorUnit
        self.dispUnits = dispUnits
        self.extLst = extLst
        kw.setdefault("majorGridlines", ChartLines())
        kw.setdefault("axId", 100)
        kw.setdefault("crossAx", 10)
        super().__init__(**kw)

    @classmethod
    def from_tree(cls, node):
        self = super().from_tree(node)
        gridlines = node.find("{%s}majorGridlines" % CHART_NS)
        if gridlines is None:
            self.majorGridlines = None
        return self


class TextAxis(_BaseAxis):
    tagname = "catAx"

    auto: bool | None = Field.nested_bool(allow_none=True)
    lblAlgn: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_none_set(frozenset({"ctr", "l", "r"}), "lblAlgn"),
    )
    lblOffset: int | None = Field.nested_value(
        expected_type=int,
        allow_none=True,
        converter=lambda v: _chart_lbl_offset_minmax(
            v, field_name="lblOffset", min_v=0, max_v=1000
        ),
    )
    tickLblSkip: int | None = Field.nested_value(expected_type=int, allow_none=True)
    tickMarkSkip: int | None = Field.nested_value(expected_type=int, allow_none=True)
    noMultiLvlLbl: bool | None = Field.nested_bool(allow_none=True)
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = _BaseAxis.xml_order + (
        "auto",
        "lblAlgn",
        "lblOffset",
        "tickLblSkip",
        "tickMarkSkip",
        "noMultiLvlLbl",
    )

    def __init__(
        self,
        auto=None,
        lblAlgn=None,
        lblOffset=100,
        tickLblSkip=None,
        tickMarkSkip=None,
        noMultiLvlLbl=None,
        extLst=None,
        **kw,
    ):
        self.auto = auto
        self.lblAlgn = lblAlgn
        self.lblOffset = lblOffset
        self.tickLblSkip = tickLblSkip
        self.tickMarkSkip = tickMarkSkip
        self.noMultiLvlLbl = noMultiLvlLbl
        self.extLst = extLst
        kw.setdefault("axId", 10)
        kw.setdefault("crossAx", 100)
        super().__init__(**kw)


class DateAxis(TextAxis):
    tagname = "dateAx"

    lblOffset: int | None = Field.nested_value(expected_type=int, allow_none=True)

    baseTimeUnit: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_none_set(frozenset({"days", "months", "years"}), "baseTimeUnit"),
    )
    majorUnit: float | None = Field.nested_value(expected_type=float, allow_none=True)
    majorTimeUnit: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_none_set(frozenset({"days", "months", "years"}), "majorTimeUnit"),
    )
    minorUnit: float | None = Field.nested_value(expected_type=float, allow_none=True)
    minorTimeUnit: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_none_set(frozenset({"days", "months", "years"}), "minorTimeUnit"),
    )

    xml_order = (
        "axId",
        "scaling",
        "delete",
        "axPos",
        "majorGridlines",
        "minorGridlines",
        "title",
        "numFmt",
        "majorTickMark",
        "minorTickMark",
        "tickLblPos",
        "spPr",
        "txPr",
        "crossAx",
        "crosses",
        "crossesAt",
        "auto",
        "lblOffset",
        "baseTimeUnit",
        "majorUnit",
        "majorTimeUnit",
        "minorUnit",
        "minorTimeUnit",
    )

    def __init__(
        self,
        auto=None,
        lblOffset=None,
        baseTimeUnit=None,
        majorUnit=None,
        majorTimeUnit=None,
        minorUnit=None,
        minorTimeUnit=None,
        extLst=None,
        **kw,
    ):
        self.baseTimeUnit = baseTimeUnit
        self.majorUnit = majorUnit
        self.majorTimeUnit = majorTimeUnit
        self.minorUnit = minorUnit
        self.minorTimeUnit = minorTimeUnit
        kw.setdefault("axId", 500)
        kw.setdefault("lblOffset", lblOffset)
        super().__init__(auto=auto, extLst=extLst, **kw)


class SeriesAxis(_BaseAxis):
    tagname = "serAx"

    tickLblSkip: int | None = Field.nested_value(expected_type=int, allow_none=True)
    tickMarkSkip: int | None = Field.nested_value(expected_type=int, allow_none=True)
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = _BaseAxis.xml_order + ("tickLblSkip", "tickMarkSkip")

    def __init__(
        self,
        tickLblSkip=None,
        tickMarkSkip=None,
        extLst=None,
        **kw,
    ):
        self.tickLblSkip = tickLblSkip
        self.tickMarkSkip = tickMarkSkip
        self.extLst = extLst
        kw.setdefault("axId", 1000)
        kw.setdefault("crossAx", 10)
        super().__init__(**kw)
