# Copyright (c) 2010-2024 fastpyxl

"""Worksheet Properties"""

from fastpyxl.styles.colors import Color
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import Field


def _color_converter(value):
    if value is None:
        return None
    if isinstance(value, Color):
        return value
    if isinstance(value, str):
        return Color(rgb=value)
    raise FieldValidationError(f"tabColor rejected value {value!r}")


class Outline(Serialisable):

    tagname = "outlinePr"

    applyStyles: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    summaryBelow: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    summaryRight: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showOutlineSymbols: bool | None = Field.attribute(expected_type=bool, allow_none=True)

    def __init__(self,
                 applyStyles=None,
                 summaryBelow=None,
                 summaryRight=None,
                 showOutlineSymbols=None
                 ):
        self.applyStyles = applyStyles
        self.summaryBelow = summaryBelow
        self.summaryRight = summaryRight
        self.showOutlineSymbols = showOutlineSymbols


class PageSetupProperties(Serialisable):

    tagname = "pageSetUpPr"

    autoPageBreaks: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    fitToPage: bool | None = Field.attribute(expected_type=bool, allow_none=True)

    def __init__(self, autoPageBreaks=None, fitToPage=None):
        self.autoPageBreaks = autoPageBreaks
        self.fitToPage = fitToPage


class WorksheetProperties(Serialisable):

    tagname = "sheetPr"

    codeName: str | None = Field.attribute(expected_type=str, allow_none=True)
    enableFormatConditionsCalculation: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    filterMode: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    published: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    syncHorizontal: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    syncRef: str | None = Field.attribute(expected_type=str, allow_none=True)
    syncVertical: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    transitionEvaluation: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    transitionEntry: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    tabColor: Color | None = Field.element(expected_type=Color, allow_none=True, converter=_color_converter)
    outlinePr: Outline | None = Field.element(expected_type=Outline, allow_none=True)
    pageSetUpPr: PageSetupProperties | None = Field.element(expected_type=PageSetupProperties, allow_none=True)

    xml_order = ("tabColor", "outlinePr", "pageSetUpPr")

    def __init__(self,
                 codeName=None,
                 enableFormatConditionsCalculation=None,
                 filterMode=None,
                 published=None,
                 syncHorizontal=None,
                 syncRef=None,
                 syncVertical=None,
                 transitionEvaluation=None,
                 transitionEntry=None,
                 tabColor=None,
                 outlinePr=None,
                 pageSetUpPr=None
                 ):
        self.codeName = codeName
        self.enableFormatConditionsCalculation = enableFormatConditionsCalculation
        self.filterMode = filterMode
        self.published = published
        self.syncHorizontal = syncHorizontal
        self.syncRef = syncRef
        self.syncVertical = syncVertical
        self.transitionEvaluation = transitionEvaluation
        self.transitionEntry = transitionEntry
        self.tabColor = tabColor
        if outlinePr is None:
            self.outlinePr = Outline(summaryBelow=True, summaryRight=True)
        else:
            self.outlinePr = outlinePr

        if pageSetUpPr is None:
            pageSetUpPr = PageSetupProperties()
        self.pageSetUpPr = pageSetUpPr
