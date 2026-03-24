# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import Field
from fastpyxl.xml.constants import REL_NS


class PrintPageSetup(Serialisable):
    """ Worksheet print page setup """

    tagname = "pageSetup"

    orientation: str | None = Field.attribute(expected_type=str, allow_none=True, converter=lambda v: _enum(v, ("default", "portrait", "landscape"), "orientation"))
    paperSize: int | None = Field.attribute(expected_type=int, allow_none=True)
    scale: int | None = Field.attribute(expected_type=int, allow_none=True)
    fitToHeight: int | None = Field.attribute(expected_type=int, allow_none=True)
    fitToWidth: int | None = Field.attribute(expected_type=int, allow_none=True)
    firstPageNumber: int | None = Field.attribute(expected_type=int, allow_none=True)
    useFirstPageNumber: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    paperHeight: str | None = Field.attribute(expected_type=str, allow_none=True)
    paperWidth: str | None = Field.attribute(expected_type=str, allow_none=True)
    pageOrder: str | None = Field.attribute(expected_type=str, allow_none=True, converter=lambda v: _enum(v, ("downThenOver", "overThenDown"), "pageOrder"))
    usePrinterDefaults: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    blackAndWhite: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    draft: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    cellComments: str | None = Field.attribute(expected_type=str, allow_none=True, converter=lambda v: _enum(v, ("asDisplayed", "atEnd", "none"), "cellComments"))
    errors: str | None = Field.attribute(expected_type=str, allow_none=True, converter=lambda v: _enum(v, ("displayed", "blank", "dash", "NA"), "errors"))
    horizontalDpi: int | None = Field.attribute(expected_type=int, allow_none=True)
    verticalDpi: int | None = Field.attribute(expected_type=int, allow_none=True)
    copies: int | None = Field.attribute(expected_type=int, allow_none=True)
    id: str | None = Field.attribute(expected_type=str, allow_none=True, namespace=REL_NS)

    def __init__(self,
                 worksheet=None,
                 orientation=None,
                 paperSize=None,
                 scale=None,
                 fitToHeight=None,
                 fitToWidth=None,
                 firstPageNumber=None,
                 useFirstPageNumber=None,
                 paperHeight=None,
                 paperWidth=None,
                 pageOrder=None,
                 usePrinterDefaults=None,
                 blackAndWhite=None,
                 draft=None,
                 cellComments=None,
                 errors=None,
                 horizontalDpi=None,
                 verticalDpi=None,
                 copies=None,
                 id=None):
        self._parent = worksheet
        self.orientation = orientation
        self.paperSize = paperSize
        self.scale = scale
        self.fitToHeight = fitToHeight
        self.fitToWidth = fitToWidth
        self.firstPageNumber = firstPageNumber
        self.useFirstPageNumber = useFirstPageNumber
        self.paperHeight = paperHeight
        self.paperWidth = paperWidth
        self.pageOrder = pageOrder
        self.usePrinterDefaults = usePrinterDefaults
        self.blackAndWhite = blackAndWhite
        self.draft = draft
        self.cellComments = cellComments
        self.errors = errors
        self.horizontalDpi = horizontalDpi
        self.verticalDpi = verticalDpi
        self.copies = copies
        self.id = id

    def __bool__(self):
        return bool(dict(self))

    @property
    def sheet_properties(self):
        """
        Proxy property
        """
        return self._parent.sheet_properties.pageSetUpPr

    @property
    def fitToPage(self):
        return self.sheet_properties.fitToPage

    @fitToPage.setter
    def fitToPage(self, value):
        self.sheet_properties.fitToPage = value

    @property
    def autoPageBreaks(self):
        return self.sheet_properties.autoPageBreaks

    @autoPageBreaks.setter
    def autoPageBreaks(self, value):
        self.sheet_properties.autoPageBreaks = value

    @classmethod
    def from_tree(cls, node):
        self = super().from_tree(node)
        self.id = None # strip link to binary settings
        return self


class PrintOptions(Serialisable):
    """ Worksheet print options """

    tagname = "printOptions"
    horizontalCentered: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    verticalCentered: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    headings: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    gridLines: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    gridLinesSet: bool | None = Field.attribute(expected_type=bool, allow_none=True)

    def __init__(self, horizontalCentered=None,
                 verticalCentered=None,
                 headings=None,
                 gridLines=None,
                 gridLinesSet=None,
                 ):
        self.horizontalCentered = horizontalCentered
        self.verticalCentered = verticalCentered
        self.headings = headings
        self.gridLines = gridLines
        self.gridLinesSet = gridLinesSet


    def __bool__(self):
        return bool(dict(self))


class PageMargins(Serialisable):
    """
    Information about page margins for view/print layouts.
    Standard values (in inches)
    left, right = 0.75
    top, bottom = 1
    header, footer = 0.5
    """
    tagname = "pageMargins"

    left: float | None = Field.attribute(expected_type=float, allow_none=True, default=0.75)
    right: float | None = Field.attribute(expected_type=float, allow_none=True, default=0.75)
    top: float | None = Field.attribute(expected_type=float, allow_none=True, default=1)
    bottom: float | None = Field.attribute(expected_type=float, allow_none=True, default=1)
    header: float | None = Field.attribute(expected_type=float, allow_none=True, default=0.5)
    footer: float | None = Field.attribute(expected_type=float, allow_none=True, default=0.5)

    def __init__(self, left=0.75, right=0.75, top=1, bottom=1, header=0.5,
                 footer=0.5):
        self.left = left
        self.right = right
        self.top = top
        self.bottom = bottom
        self.header = header
        self.footer = footer


def _enum(value, allowed, field_name):
    if value is None:
        return None
    if value not in allowed:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return value
