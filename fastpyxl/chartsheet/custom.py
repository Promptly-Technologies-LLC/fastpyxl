# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.worksheet.header_footer import HeaderFooter

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import Field
from fastpyxl.worksheet.page import (
    PageMargins,
    PrintPageSetup
)


class CustomChartsheetView(Serialisable):
    tagname = "customSheetView"

    guid: str | None = Field.attribute(expected_type=str, allow_none=True)
    scale: int | None = Field.attribute(expected_type=int, allow_none=True)
    state: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("visible", "hidden", "veryHidden"), "state"),
    )
    zoomToFit: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    pageMargins: PageMargins | None = Field.element(expected_type=PageMargins, allow_none=True)
    pageSetup: PrintPageSetup | None = Field.element(expected_type=PrintPageSetup, allow_none=True)
    headerFooter: HeaderFooter | None = Field.element(expected_type=HeaderFooter, allow_none=True)

    xml_order = ('pageMargins', 'pageSetup', 'headerFooter')

    def __init__(self,
                 guid=None,
                 scale=None,
                 state='visible',
                 zoomToFit=None,
                 pageMargins=None,
                 pageSetup=None,
                 headerFooter=None,
                 ):
        self.guid = guid
        self.scale = scale
        self.state = state
        self.zoomToFit = zoomToFit
        self.pageMargins = pageMargins
        self.pageSetup = pageSetup
        self.headerFooter = headerFooter


class CustomChartsheetViews(Serialisable):
    tagname = "customSheetViews"

    customSheetView: list[CustomChartsheetView] | None = Field.sequence(
        expected_type=CustomChartsheetView,
        allow_none=True,
    )
    xml_order = ('customSheetView',)

    def __init__(self,
                 customSheetView=None,
                 ):
        self.customSheetView = customSheetView


def _enum_converter(value, allowed_values, field_name: str):
    if value is None:
        return None
    if value not in allowed_values:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return value
