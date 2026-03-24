# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.descriptors.excel import ExtensionList

from .layout import Layout
from .shapes import GraphicalProperties
from .text import RichText


class LegendEntry(Serialisable):
    tagname = "legendEntry"

    idx: int | None = Field.nested_value(expected_type=int, allow_none=True)
    delete: bool | None = Field.nested_bool()
    txPr: RichText | None = Field.element(expected_type=RichText, allow_none=True)
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = ("idx", "delete", "txPr")

    def __init__(self, idx=0, delete=False, txPr=None, extLst=None):
        self.idx = idx
        self.delete = delete
        self.txPr = txPr
        self.extLst = extLst


class Legend(Serialisable):
    tagname = "legend"

    legendPos: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _legend_pos(v),
    )
    position = AliasField("legendPos")
    legendEntry: list[LegendEntry] | None = Field.sequence(
        expected_type=LegendEntry, allow_none=True
    )
    layout: Layout | None = Field.element(expected_type=Layout, allow_none=True)
    overlay: bool | None = Field.nested_bool(allow_none=True)
    spPr: GraphicalProperties | None = Field.element(
        expected_type=GraphicalProperties, allow_none=True
    )
    graphicalProperties = AliasField("spPr")
    txPr: RichText | None = Field.element(expected_type=RichText, allow_none=True)
    textProperties = AliasField("txPr")
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = ("legendPos", "legendEntry", "layout", "overlay", "spPr", "txPr")

    def __init__(
        self,
        legendPos="r",
        legendEntry=(),
        layout=None,
        overlay=None,
        spPr=None,
        txPr=None,
        extLst=None,
    ):
        self.legendPos = legendPos
        self.legendEntry = list(legendEntry) if legendEntry is not None else []
        self.layout = layout
        self.overlay = overlay
        self.spPr = spPr
        self.txPr = txPr
        self.extLst = extLst


def _legend_pos(v):
    from fastpyxl.typed_serialisable.errors import FieldValidationError

    if v is None:
        return None
    allowed = frozenset({"b", "tr", "l", "r", "t"})
    if v not in allowed:
        raise FieldValidationError(f"legendPos rejected value {v!r}")
    return v
