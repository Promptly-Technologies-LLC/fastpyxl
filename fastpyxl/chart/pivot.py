
# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.descriptors.excel import ExtensionList

from .label import DataLabel
from .marker import Marker
from .shapes import GraphicalProperties
from .text import RichText


class PivotSource(Serialisable):
    tagname = "pivotSource"

    name: str | None = Field.nested_text(expected_type=str, allow_none=True)
    fmtId: int | None = Field.nested_value(expected_type=int, allow_none=True)
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = ("name", "fmtId")

    def __init__(self, name=None, fmtId=None, extLst=None):
        self.name = name
        self.fmtId = fmtId
        self.extLst = extLst


class PivotFormat(Serialisable):
    tagname = "pivotFmt"

    idx: int | None = Field.nested_value(expected_type=int, allow_none=True)
    spPr: GraphicalProperties | None = Field.element(
        expected_type=GraphicalProperties, allow_none=True
    )
    graphicalProperties = AliasField("spPr")
    txPr: RichText | None = Field.element(expected_type=RichText, allow_none=True)
    TextBody = AliasField("txPr")
    marker: Marker | None = Field.element(expected_type=Marker, allow_none=True)
    dLbl: DataLabel | None = Field.element(expected_type=DataLabel, allow_none=True)
    DataLabel = AliasField("dLbl")
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = ("idx", "spPr", "txPr", "marker", "dLbl")

    def __init__(
        self,
        idx=0,
        spPr=None,
        txPr=None,
        marker=None,
        dLbl=None,
        extLst=None,
    ):
        self.idx = idx
        self.spPr = spPr
        self.txPr = txPr
        self.marker = marker
        self.dLbl = dLbl
        self.extLst = extLst
