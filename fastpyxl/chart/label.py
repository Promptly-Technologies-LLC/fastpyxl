# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.descriptors.excel import ExtensionList

from .shapes import GraphicalProperties
from .text import RichText


def _none_set(allowed: frozenset, field_name: str):
    def _c(v):
        if v is None or v == "none":
            return None
        if v not in allowed:
            raise FieldValidationError(f"{field_name} rejected value {v!r}")
        return v

    return _c


class _DataLabelBase(Serialisable):
    numFmt: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        value_attribute="formatCode",
    )
    spPr: GraphicalProperties | None = Field.element(
        expected_type=GraphicalProperties, allow_none=True
    )
    graphicalProperties = AliasField("spPr")
    txPr: RichText | None = Field.element(expected_type=RichText, allow_none=True)
    textProperties = AliasField("txPr")
    dLblPos: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_none_set(
            frozenset(
                {
                    "bestFit",
                    "b",
                    "ctr",
                    "inBase",
                    "inEnd",
                    "l",
                    "outEnd",
                    "r",
                    "t",
                }
            ),
            "dLblPos",
        ),
    )
    position = AliasField("dLblPos")
    showLegendKey: bool | None = Field.nested_bool(allow_none=True)
    showVal: bool | None = Field.nested_bool(allow_none=True)
    showCatName: bool | None = Field.nested_bool(allow_none=True)
    showSerName: bool | None = Field.nested_bool(allow_none=True)
    showPercent: bool | None = Field.nested_bool(allow_none=True)
    showBubbleSize: bool | None = Field.nested_bool(allow_none=True)
    showLeaderLines: bool | None = Field.nested_bool(allow_none=True)
    separator: str | None = Field.nested_text(expected_type=str, allow_none=True)
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    _base_xml_order = (
        "numFmt",
        "spPr",
        "txPr",
        "dLblPos",
        "showLegendKey",
        "showVal",
        "showCatName",
        "showSerName",
        "showPercent",
        "showBubbleSize",
        "showLeaderLines",
        "separator",
    )

    def __init__(
        self,
        numFmt=None,
        spPr=None,
        txPr=None,
        dLblPos=None,
        showLegendKey=None,
        showVal=None,
        showCatName=None,
        showSerName=None,
        showPercent=None,
        showBubbleSize=None,
        showLeaderLines=None,
        separator=None,
        extLst=None,
    ):
        self.numFmt = numFmt
        self.spPr = spPr
        self.txPr = txPr
        self.dLblPos = dLblPos
        self.showLegendKey = showLegendKey
        self.showVal = showVal
        self.showCatName = showCatName
        self.showSerName = showSerName
        self.showPercent = showPercent
        self.showBubbleSize = showBubbleSize
        self.showLeaderLines = showLeaderLines
        self.separator = separator
        self.extLst = extLst


class DataLabel(_DataLabelBase):
    tagname = "dLbl"

    idx: int | None = Field.nested_value(expected_type=int, allow_none=True)

    xml_order = ("idx",) + _DataLabelBase._base_xml_order

    def __init__(self, idx=0, **kw):
        self.idx = idx
        super().__init__(**kw)


class DataLabelList(_DataLabelBase):
    tagname = "dLbls"

    dLbl: list[DataLabel] | None = Field.sequence(expected_type=DataLabel, allow_none=True)
    delete: bool | None = Field.nested_bool(allow_none=True)

    xml_order = ("delete", "dLbl") + _DataLabelBase._base_xml_order

    def __init__(self, dLbl=(), delete=None, **kw):
        self.dLbl = list(dLbl) if dLbl is not None else []
        self.delete = delete
        super().__init__(**kw)
