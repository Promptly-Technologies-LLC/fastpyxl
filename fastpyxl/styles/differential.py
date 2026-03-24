# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.compat import safe_string
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import AliasField, Field
from fastpyxl.styles import (
    Font,
    Fill,
    Border,
    Alignment,
    Protection,
    )
from .numbers import NumberFormat


class DifferentialStyle(Serialisable):

    tagname = "dxf"

    xml_order = ("font", "numFmt", "fill", "alignment", "border", "protection")

    font: Font | None = Field.element(expected_type=Font, allow_none=True)
    numFmt: NumberFormat | None = Field.element(expected_type=NumberFormat, allow_none=True)
    fill: Fill | None = Field.element(expected_type=Fill, allow_none=True)
    alignment: Alignment | None = Field.element(expected_type=Alignment, allow_none=True)
    border: Border | None = Field.element(expected_type=Border, allow_none=True)
    protection: Protection | None = Field.element(expected_type=Protection, allow_none=True)

    def __init__(self,
                 font=None,
                 numFmt=None,
                 fill=None,
                 alignment=None,
                 border=None,
                 protection=None,
                 extLst=None,
                ):
        self.font = font
        self.numFmt = numFmt
        self.fill = fill
        self.alignment = alignment
        self.border = border
        self.protection = protection
        self.extLst = extLst


class DifferentialStyleList(Serialisable):
    """
    Dedupable container for differential styles.
    """

    tagname = "dxfs"

    dxf: list[DifferentialStyle] = Field.sequence(expected_type=DifferentialStyle)
    styles: list[DifferentialStyle] = AliasField("dxf")


    def __init__(self, dxf=(), count=None):
        self.dxf = list(dxf)


    def append(self, dxf):
        """
        Check to see whether style already exists and append it if does not.
        """
        if not isinstance(dxf, DifferentialStyle):
            raise TypeError('expected ' + str(DifferentialStyle))
        if dxf in self.styles:
            return
        self.styles.append(dxf)


    def add(self, dxf):
        """
        Add a differential style and return its index
        """
        self.append(dxf)
        return self.styles.index(dxf)


    def __bool__(self):
        return bool(self.styles)


    def __getitem__(self, idx):
        return self.styles[idx]


    @property
    def count(self):
        return len(self.dxf)


    def __iter__(self):
        yield "count", safe_string(self.count)
