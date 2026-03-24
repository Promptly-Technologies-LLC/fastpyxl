# Copyright (c) 2010-2024 fastpyxl


from fastpyxl.descriptors.excel import ExtensionList
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field
from fastpyxl.drawing.spreadsheet_drawing import (
    AbsoluteAnchor,
    SpreadsheetDrawing,
)
from fastpyxl.worksheet.page import (
    PageMargins,
    PrintPageSetup
)
from fastpyxl.worksheet.drawing import Drawing
from fastpyxl.worksheet.header_footer import HeaderFooter
from fastpyxl.workbook.child import _WorkbookChild
from fastpyxl.xml.constants import SHEET_MAIN_NS

from .relation import DrawingHF, SheetBackgroundPicture
from .properties import ChartsheetProperties
from .protection import ChartsheetProtection
from .views import ChartsheetViewList
from .custom import CustomChartsheetViews
from .publish import WebPublishItems


class Chartsheet(_WorkbookChild, Serialisable):

    tagname = "chartsheet"
    _default_title = "Chart"
    _rel_type = "chartsheet"
    _path = "/xl/chartsheets/sheet{0}.xml"
    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"

    sheetPr: ChartsheetProperties | None = Field.element(expected_type=ChartsheetProperties, allow_none=True, default=None)
    sheetViews: ChartsheetViewList | None = Field.element(expected_type=ChartsheetViewList, allow_none=True, default=None)
    sheetProtection: ChartsheetProtection | None = Field.element(expected_type=ChartsheetProtection, allow_none=True, default=None)
    customSheetViews: CustomChartsheetViews | None = Field.element(expected_type=CustomChartsheetViews, allow_none=True, default=None)
    pageMargins: PageMargins | None = Field.element(expected_type=PageMargins, allow_none=True, default=None)
    pageSetup: PrintPageSetup | None = Field.element(expected_type=PrintPageSetup, allow_none=True, default=None)
    drawing: Drawing | None = Field.element(expected_type=Drawing, allow_none=True, default=None)
    drawingHF: DrawingHF | None = Field.element(expected_type=DrawingHF, allow_none=True, default=None)
    picture: SheetBackgroundPicture | None = Field.element(expected_type=SheetBackgroundPicture, allow_none=True, default=None)
    webPublishItems: WebPublishItems | None = Field.element(expected_type=WebPublishItems, allow_none=True, default=None)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True, default=None)
    headerFooter: HeaderFooter | None = Field.element(expected_type=HeaderFooter, allow_none=True, default=None)
    HeaderFooter = AliasField('headerFooter', default=None)

    xml_order = (
        'sheetPr', 'sheetViews', 'sheetProtection', 'customSheetViews',
        'pageMargins', 'pageSetup', 'headerFooter', 'drawing', 'drawingHF',
        'picture', 'webPublishItems', 'extLst')

    def __init__(self,
                 sheetPr=None,
                 sheetViews=None,
                 sheetProtection=None,
                 customSheetViews=None,
                 pageMargins=None,
                 pageSetup=None,
                 headerFooter=None,
                 drawing=None,
                 drawingHF=None,
                 picture=None,
                 webPublishItems=None,
                 extLst=None,
                 parent=None,
                 title="",
                 sheet_state='visible',
                 ):
        super().__init__(parent, title)
        self._charts = []
        self.sheetPr = sheetPr
        if sheetViews is None:
            sheetViews = ChartsheetViewList()
        self.sheetViews = sheetViews
        self.sheetProtection = sheetProtection
        self.customSheetViews = customSheetViews
        self.pageMargins = pageMargins
        self.pageSetup = pageSetup
        if headerFooter is not None:
            self.headerFooter = headerFooter
        self.drawing = Drawing("rId1")
        self.drawingHF = drawingHF
        self.picture = picture
        self.webPublishItems = webPublishItems
        self.extLst = extLst
        if sheet_state not in ('visible', 'hidden', 'veryHidden'):
            raise FieldValidationError(f"sheet_state rejected value {sheet_state!r}")
        self.sheet_state = sheet_state


    def add_chart(self, chart):
        chart.anchor = AbsoluteAnchor()
        self._charts.append(chart)


    def to_tree(self):
        self._drawing = SpreadsheetDrawing()
        self._drawing.charts = self._charts
        tree = super().to_tree()
        if not self.headerFooter:
            el = tree.find('headerFooter')
            tree.remove(el)
        tree.set("xmlns", SHEET_MAIN_NS)
        return tree
