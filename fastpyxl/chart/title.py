# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.descriptors.excel import ExtensionList

from .text import Text, RichText
from .layout import Layout
from .shapes import GraphicalProperties

from fastpyxl.drawing.text import (
    Paragraph,
    RegularTextRun,
    ParagraphProperties,
    CharacterProperties,
)


class Title(Serialisable):
    tagname = "title"

    tx: Text | None = Field.element(expected_type=Text, allow_none=True)
    text = AliasField("tx")
    layout: Layout | None = Field.element(expected_type=Layout, allow_none=True)
    overlay: bool | None = Field.nested_bool(allow_none=True)
    spPr: GraphicalProperties | None = Field.element(
        expected_type=GraphicalProperties, allow_none=True
    )
    graphicalProperties = AliasField("spPr")
    txPr: RichText | None = Field.element(expected_type=RichText, allow_none=True)
    body = AliasField("txPr")
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = ("tx", "layout", "overlay", "spPr", "txPr")

    def __init__(
        self,
        tx=None,
        layout=None,
        overlay=None,
        spPr=None,
        txPr=None,
        extLst=None,
    ):
        if tx is None:
            tx = Text()
        self.tx = tx
        self.layout = layout
        self.overlay = overlay
        self.spPr = spPr
        self.txPr = txPr
        self.extLst = extLst


def title_maker(text):
    title = Title()
    paraprops = ParagraphProperties()
    paraprops.defRPr = CharacterProperties()
    paras = [Paragraph(r=[RegularTextRun(t=s)], pPr=paraprops) for s in text.split("\n")]

    title.tx.rich.paragraphs = paras
    return title


def title_from_value(value):
    if value is None:
        return None
    if isinstance(value, str):
        return title_maker(value)
    return value
