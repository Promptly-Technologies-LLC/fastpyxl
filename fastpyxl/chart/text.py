# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.drawing.text import (
    RichTextProperties,
    ListStyle,
    Paragraph,
)

from .data_source import StrRef


class RichText(Serialisable):
    """
    From the specification: 21.2.2.216

    This element specifies text formatting. The lstStyle element is not supported.
    """

    tagname = "rich"

    bodyPr: RichTextProperties | None = Field.element(
        expected_type=RichTextProperties, allow_none=True
    )
    properties = AliasField("bodyPr")
    lstStyle: ListStyle | None = Field.element(expected_type=ListStyle, allow_none=True)
    p: list[Paragraph] | None = Field.sequence(expected_type=Paragraph, allow_none=True)
    paragraphs = AliasField("p")

    xml_order = ("bodyPr", "lstStyle", "p")

    def __init__(self, bodyPr=None, lstStyle=None, p=None):
        if bodyPr is None:
            bodyPr = RichTextProperties()
        self.bodyPr = bodyPr
        self.lstStyle = lstStyle
        if p is None:
            p = [Paragraph()]
        self.p = list(p)


class Text(Serialisable):
    """
    The value can be either a cell reference or a text element
    If both are present then the reference will be used.
    """

    tagname = "tx"

    strRef: StrRef | None = Field.element(expected_type=StrRef, allow_none=True)
    rich: RichText | None = Field.element(expected_type=RichText, allow_none=True)

    xml_order = ("strRef", "rich")

    def __init__(self, strRef=None, rich=None):
        self.strRef = strRef
        if rich is None:
            rich = RichText()
        self.rich = rich

    def to_tree(self, tagname=None, idx=None, namespace=None):
        del idx, namespace
        if self.strRef and self.rich:
            self.rich = None
        return super().to_tree(tagname)
