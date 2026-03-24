# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.worksheet.page import PrintPageSetup
from fastpyxl.worksheet.header_footer import HeaderFooter


class PageMargins(Serialisable):
    """
    Identical to fastpyxl.worksheet.page.Pagemargins but element names are different :-/
    """

    tagname = "pageMargins"

    l: float | None = Field.attribute(expected_type=float, default=0.75)  # noqa: E741
    left = AliasField("l", default=None)
    r: float | None = Field.attribute(expected_type=float, default=0.75)
    right = AliasField("r", default=None)
    t: float | None = Field.attribute(expected_type=float, default=1)
    top = AliasField("t", default=None)
    b: float | None = Field.attribute(expected_type=float, default=1)
    bottom = AliasField("b", default=None)
    header: float | None = Field.attribute(expected_type=float, default=0.5)
    footer: float | None = Field.attribute(expected_type=float, default=0.5)

    def __init__(self, l=0.75, r=0.75, t=1, b=1, header=0.5, footer=0.5):  # noqa: E741
        self.l = l  # noqa: E741
        self.r = r
        self.t = t
        self.b = b
        self.header = header
        self.footer = footer


class PrintSettings(Serialisable):
    tagname = "printSettings"

    headerFooter: HeaderFooter | None = Field.element(
        expected_type=HeaderFooter, allow_none=True, default=None
    )
    pageMargins: PageMargins | None = Field.element(
        expected_type=PageMargins, allow_none=True, default=None
    )
    pageSetup: PrintPageSetup | None = Field.element(
        expected_type=PrintPageSetup, allow_none=True, default=None
    )

    xml_order = ("headerFooter", "pageMargins", "pageSetup")

    def __init__(self, headerFooter=None, pageMargins=None, pageSetup=None):
        self.headerFooter = headerFooter
        self.pageMargins = pageMargins
        self.pageSetup = pageSetup
