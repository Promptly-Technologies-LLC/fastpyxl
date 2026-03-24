# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.descriptors.excel import ExtensionList
from fastpyxl.xml.constants import SHEET_MAIN_NS, REL_NS
from fastpyxl.xml.functions import tostring
from fastpyxl.utils import range_boundaries
from fastpyxl.utils.escape import escape, unescape
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import Field

from .related import Related

from .filters import (
    AutoFilter,
    SortState,
)

TABLESTYLES = tuple(
    ["TableStyleMedium{0}".format(i) for i in range(1, 29)]
    + ["TableStyleLight{0}".format(i) for i in range(1, 22)]
    + ["TableStyleDark{0}".format(i) for i in range(1, 12)]
)

PIVOTSTYLES = tuple(
    ["PivotStyleMedium{0}".format(i) for i in range(1, 29)]
    + ["PivotStyleLight{0}".format(i) for i in range(1, 29)]
    + ["PivotStyleDark{0}".format(i) for i in range(1, 29)]
)


def _enum(value, allowed, field_name):
    if value is None:
        return None
    if value not in allowed:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return value


def _table_name_converter(value):
    if value is not None and " " in value:
        raise FieldValidationError("displayName rejected value with spaces")
    return value


class TableStyleInfo(Serialisable):

    tagname = "tableStyleInfo"

    name: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    showFirstColumn: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showLastColumn: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showRowStripes: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showColumnStripes: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)

    def __init__(self, name=None, showFirstColumn=None, showLastColumn=None, showRowStripes=None, showColumnStripes=None):
        self.name = name
        self.showFirstColumn = showFirstColumn
        self.showLastColumn = showLastColumn
        self.showRowStripes = showRowStripes
        self.showColumnStripes = showColumnStripes


class XMLColumnProps(Serialisable):

    tagname = "xmlColumnPr"

    mapId: int | None = Field.attribute(expected_type=int, default=None)
    xpath: str | None = Field.attribute(expected_type=str, default=None)
    denormalized: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    xmlDataType: str | None = Field.attribute(expected_type=str, default=None)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True, serialize=False, default=None)

    def __init__(self, mapId=None, xpath=None, denormalized=None, xmlDataType=None, extLst=None):
        self.mapId = mapId
        self.xpath = xpath
        self.denormalized = denormalized
        self.xmlDataType = xmlDataType
        self.extLst = extLst


class TableFormula(Serialisable):

    tagname = "tableFormula"

    array: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)

    def __init__(self, array=None, attr_text=None):
        self.array = array
        self.attr_text = attr_text

    @property
    def text(self):
        return self.attr_text

    @text.setter
    def text(self, value):
        self.attr_text = value

    def to_tree(self, tagname=None, idx=None, namespace=None):
        tree = super().to_tree(tagname, idx, namespace)
        tree.text = self.attr_text
        return tree

    @classmethod
    def from_tree(cls, node):
        obj = super().from_tree(node)
        obj.attr_text = node.text
        return obj


class TableColumn(Serialisable):

    tagname = "tableColumn"

    id: int | None = Field.attribute(expected_type=int, default=None)
    uniqueName: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    name: str | None = Field.attribute(expected_type=str, default=None)
    totalsRowFunction: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum(v, ('sum', 'min', 'max', 'average', 'count', 'countNums', 'stdDev', 'var', 'custom'), 'totalsRowFunction'), default=None
    )
    totalsRowLabel: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    queryTableFieldId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    headerRowDxfId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    dataDxfId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    totalsRowDxfId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    headerRowCellStyle: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    dataCellStyle: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    totalsRowCellStyle: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    calculatedColumnFormula: TableFormula | None = Field.element(expected_type=TableFormula, allow_none=True, default=None)
    totalsRowFormula: TableFormula | None = Field.element(expected_type=TableFormula, allow_none=True, default=None)
    xmlColumnPr: XMLColumnProps | None = Field.element(expected_type=XMLColumnProps, allow_none=True, default=None)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True, serialize=False, default=None)

    xml_order = ('calculatedColumnFormula', 'totalsRowFormula', 'xmlColumnPr', 'extLst')

    def __init__(self, id=None, uniqueName=None, name=None, totalsRowFunction=None, totalsRowLabel=None, queryTableFieldId=None,
                 headerRowDxfId=None, dataDxfId=None, totalsRowDxfId=None, headerRowCellStyle=None, dataCellStyle=None,
                 totalsRowCellStyle=None, calculatedColumnFormula=None, totalsRowFormula=None, xmlColumnPr=None, extLst=None):
        self.id = id
        self.uniqueName = uniqueName
        self.name = name
        self.totalsRowFunction = totalsRowFunction
        self.totalsRowLabel = totalsRowLabel
        self.queryTableFieldId = queryTableFieldId
        self.headerRowDxfId = headerRowDxfId
        self.dataDxfId = dataDxfId
        self.totalsRowDxfId = totalsRowDxfId
        self.headerRowCellStyle = headerRowCellStyle
        self.dataCellStyle = dataCellStyle
        self.totalsRowCellStyle = totalsRowCellStyle
        self.calculatedColumnFormula = calculatedColumnFormula
        self.totalsRowFormula = totalsRowFormula
        self.xmlColumnPr = xmlColumnPr
        self.extLst = extLst

    def __iter__(self):
        for k, v in super().__iter__():
            if k == 'name':
                v = escape(v)
            yield k, v

    @classmethod
    def from_tree(cls, node):
        self = super().from_tree(node)
        self.name = unescape(self.name)
        return self


class _TableColumnList(Serialisable):
    tagname = "tableColumns"

    tableColumn: list[TableColumn] = Field.sequence(expected_type=TableColumn, default=list)

    xml_order = ('tableColumn',)

    def __init__(self, tableColumn=()):
        self.tableColumn = list(tableColumn)

    def __iter__(self):
        yield 'count', str(len(self.tableColumn))


class Table(Serialisable):

    _path = "/tables/table{0}.xml"
    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
    _rel_type = REL_NS + "/table"
    _rel_id = None

    tagname = "table"

    id: int | None = Field.attribute(expected_type=int, default=None)
    name: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    displayName: str | None = Field.attribute(expected_type=str, converter=_table_name_converter, default=None)
    comment: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    ref: str | None = Field.attribute(expected_type=str, default=None)
    tableType: str | None = Field.attribute(expected_type=str, allow_none=True, converter=lambda v: _enum(v, ('worksheet', 'xml', 'queryTable'), 'tableType'), default=None)
    headerRowCount: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    insertRow: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    insertRowShift: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    totalsRowCount: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    totalsRowShown: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    published: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    headerRowDxfId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    dataDxfId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    totalsRowDxfId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    headerRowBorderDxfId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    tableBorderDxfId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    totalsRowBorderDxfId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    headerRowCellStyle: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    dataCellStyle: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    totalsRowCellStyle: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    connectionId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    autoFilter: AutoFilter | None = Field.element(expected_type=AutoFilter, allow_none=True, default=None)
    sortState: SortState | None = Field.element(expected_type=SortState, allow_none=True, default=None)
    tableColumns: list[TableColumn] = Field.nested_sequence(expected_type=TableColumn, count=True, default=list)
    tableStyleInfo: TableStyleInfo | None = Field.element(expected_type=TableStyleInfo, allow_none=True, default=None)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True, serialize=False, default=None)

    xml_order = ('autoFilter', 'sortState', 'tableColumns', 'tableStyleInfo')

    def __init__(self,
                 id=1,
                 displayName=None,
                 ref=None,
                 name=None,
                 comment=None,
                 tableType=None,
                 headerRowCount=1,
                 insertRow=None,
                 insertRowShift=None,
                 totalsRowCount=None,
                 totalsRowShown=None,
                 published=None,
                 headerRowDxfId=None,
                 dataDxfId=None,
                 totalsRowDxfId=None,
                 headerRowBorderDxfId=None,
                 tableBorderDxfId=None,
                 totalsRowBorderDxfId=None,
                 headerRowCellStyle=None,
                 dataCellStyle=None,
                 totalsRowCellStyle=None,
                 connectionId=None,
                 autoFilter=None,
                 sortState=None,
                 tableColumns=(),
                 tableStyleInfo=None,
                 extLst=None,
                ):
        self.id = id
        self.displayName = displayName
        if name is None:
            name = displayName
        self.name = name
        self.comment = comment
        self.ref = ref
        self.tableType = tableType
        self.headerRowCount = headerRowCount
        self.insertRow = insertRow
        self.insertRowShift = insertRowShift
        self.totalsRowCount = totalsRowCount
        self.totalsRowShown = totalsRowShown
        self.published = published
        self.headerRowDxfId = headerRowDxfId
        self.dataDxfId = dataDxfId
        self.totalsRowDxfId = totalsRowDxfId
        self.headerRowBorderDxfId = headerRowBorderDxfId
        self.tableBorderDxfId = tableBorderDxfId
        self.totalsRowBorderDxfId = totalsRowBorderDxfId
        self.headerRowCellStyle = headerRowCellStyle
        self.dataCellStyle = dataCellStyle
        self.totalsRowCellStyle = totalsRowCellStyle
        self.connectionId = connectionId
        self.autoFilter = autoFilter
        self.sortState = sortState
        self.tableColumns = list(tableColumns)
        self.tableStyleInfo = tableStyleInfo
        self.extLst = extLst

    def to_tree(self, tagname=None, idx=None, namespace=None):
        tree = super().to_tree(tagname, idx, namespace)
        tree.set("xmlns", SHEET_MAIN_NS)
        return tree

    @property
    def path(self):
        """
        Return path within the archive
        """
        return "/xl" + self._path.format(self.id)

    def _write(self, archive):
        """
        Serialise to XML and write to archive
        """
        xml = self.to_tree()
        archive.writestr(self.path[1:], tostring(xml))

    def _initialise_columns(self):
        """
        Create a list of table columns from a cell range
        Always set a ref if we have headers (the default)
        Column headings must be strings and must match cells in the worksheet.
        """

        min_col, min_row, max_col, max_row = range_boundaries(self.ref)
        del min_row, max_row
        for idx in range(min_col, max_col+1):
            col = TableColumn(id=idx, name="Column{0}".format(idx))
            self.tableColumns.append(col)
        if self.headerRowCount and not self.autoFilter:
            self.autoFilter = AutoFilter(ref=self.ref)

    @property
    def column_names(self):
        return [column.name for column in self.tableColumns]


class TablePartList(Serialisable):

    tagname = "tableParts"

    tablePart: list[Related] = Field.sequence(expected_type=Related, default=list)

    xml_order = ('tablePart',)

    def __init__(self, count=None, tablePart=()):
        del count
        self.tablePart = list(tablePart)

    def append(self, part):
        self.tablePart.append(part)

    @property
    def count(self):
        return len(self.tablePart)

    def __iter__(self):
        yield 'count', str(self.count)

    def __bool__(self):
        return bool(self.tablePart)


class TableList(dict):

    def add(self, table):
        if not isinstance(table, Table):
            raise TypeError("You can only add tables")
        self[table.name] = table

    def get(self, name=None, table_range=None):
        if name is not None:
            return super().get(name)
        for table in self.values():
            if table_range == table.ref:
                return table

    def items(self):
        return [(name, table.ref) for name, table in super().items()]
