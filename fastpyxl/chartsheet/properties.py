# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.styles import Color
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import Field


class ChartsheetProperties(Serialisable):
    tagname = "sheetPr"

    published: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    codeName: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    tabColor: Color | None = Field.element(expected_type=Color, allow_none=True, default=None)
    xml_order = ("tabColor",)

    def __init__(self,
                 published=None,
                 codeName=None,
                 tabColor=None,
                 ):
        self.published = published
        self.codeName = codeName
        self.tabColor = tabColor
