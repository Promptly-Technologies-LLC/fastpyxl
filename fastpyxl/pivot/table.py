# Copyright (c) 2010-2024 fastpyxl


from collections import defaultdict
from fastpyxl.descriptors.serialisable import Serialisable
from fastpyxl.descriptors import (
    Typed,
    Integer,
    NoneSet,
    Set,
    String,
    Bool,
    Sequence,
)

from fastpyxl.descriptors.excel import ExtensionList
from fastpyxl.descriptors.sequence import NestedSequence
from fastpyxl.typed_serialisable.base import Serialisable as TypedSerialisable
from fastpyxl.typed_serialisable.fields import Field
from fastpyxl.xml.constants import REL_NS, SHEET_MAIN_NS
from fastpyxl.xml.functions import tostring
from fastpyxl.packaging.relationship import (
    RelationshipList,
    Relationship,
    get_rels_path
)
from .fields import Index

from fastpyxl.worksheet.filters import (
    AutoFilter,
)


class HierarchyUsage(TypedSerialisable):

    tagname = "hierarchyUsage"

    hierarchyUsage: int | None = Field.attribute(expected_type=int, allow_none=False)

    def __init__(self,
                 hierarchyUsage=None,
                ):
        self.hierarchyUsage = hierarchyUsage


class ColHierarchiesUsage(TypedSerialisable):

    tagname = "colHierarchiesUsage"

    colHierarchyUsage: list[HierarchyUsage] = Field.sequence(expected_type=HierarchyUsage)

    xml_order = ("colHierarchyUsage",)

    def __init__(self,
                 count=None,
                 colHierarchyUsage=(),
                ):
        self.colHierarchyUsage = list(colHierarchyUsage)


    @property
    def count(self):
        return len(self.colHierarchyUsage)

    def __iter__(self):
        yield from super().__iter__()
        yield "count", str(self.count)


class RowHierarchiesUsage(TypedSerialisable):

    tagname = "rowHierarchiesUsage"

    rowHierarchyUsage: list[HierarchyUsage] = Field.sequence(expected_type=HierarchyUsage)

    xml_order = ("rowHierarchyUsage",)

    def __init__(self,
                 count=None,
                 rowHierarchyUsage=(),
                ):
        self.rowHierarchyUsage = list(rowHierarchyUsage)

    @property
    def count(self):
        return len(self.rowHierarchyUsage)

    def __iter__(self):
        yield from super().__iter__()
        yield "count", str(self.count)


class PivotFilter(Serialisable):

    tagname = "filter"

    fld = Integer()
    mpFld = Integer(allow_none=True)
    type = Set(values=(['unknown', 'count', 'percent', 'sum', 'captionEqual',
                        'captionNotEqual', 'captionBeginsWith', 'captionNotBeginsWith',
                        'captionEndsWith', 'captionNotEndsWith', 'captionContains',
                        'captionNotContains', 'captionGreaterThan', 'captionGreaterThanOrEqual',
                        'captionLessThan', 'captionLessThanOrEqual', 'captionBetween',
                        'captionNotBetween', 'valueEqual', 'valueNotEqual', 'valueGreaterThan',
                        'valueGreaterThanOrEqual', 'valueLessThan', 'valueLessThanOrEqual',
                        'valueBetween', 'valueNotBetween', 'dateEqual', 'dateNotEqual',
                        'dateOlderThan', 'dateOlderThanOrEqual', 'dateNewerThan',
                        'dateNewerThanOrEqual', 'dateBetween', 'dateNotBetween', 'tomorrow',
                        'today', 'yesterday', 'nextWeek', 'thisWeek', 'lastWeek', 'nextMonth',
                        'thisMonth', 'lastMonth', 'nextQuarter', 'thisQuarter', 'lastQuarter',
                        'nextYear', 'thisYear', 'lastYear', 'yearToDate', 'Q1', 'Q2', 'Q3', 'Q4',
                        'M1', 'M2', 'M3', 'M4', 'M5', 'M6', 'M7', 'M8', 'M9', 'M10', 'M11',
                        'M12']))
    evalOrder = Integer(allow_none=True)
    id = Integer()
    iMeasureHier = Integer(allow_none=True)
    iMeasureFld = Integer(allow_none=True)
    name = String(allow_none=True)
    description = String(allow_none=True)
    stringValue1 = String(allow_none=True)
    stringValue2 = String(allow_none=True)
    autoFilter = Typed(expected_type=AutoFilter, )
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('autoFilter',)

    def __init__(self,
                 fld=None,
                 mpFld=None,
                 type=None,
                 evalOrder=None,
                 id=None,
                 iMeasureHier=None,
                 iMeasureFld=None,
                 name=None,
                 description=None,
                 stringValue1=None,
                 stringValue2=None,
                 autoFilter=None,
                 extLst=None,
                ):
        self.fld = fld
        self.mpFld = mpFld
        self.type = type
        self.evalOrder = evalOrder
        self.id = id
        self.iMeasureHier = iMeasureHier
        self.iMeasureFld = iMeasureFld
        self.name = name
        self.description = description
        self.stringValue1 = stringValue1
        self.stringValue2 = stringValue2
        self.autoFilter = autoFilter


class PivotFilters(Serialisable):

    count = Integer()
    filter = Typed(expected_type=PivotFilter, allow_none=True)

    __elements__ = ('filter',)

    def __init__(self,
                 count=None,
                 filter=None,
                ):
        self.filter = filter


class PivotTableStyle(TypedSerialisable):

    tagname = "pivotTableStyleInfo"

    name: str | None = Field.attribute(expected_type=str, allow_none=True)
    showRowHeaders: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    showColHeaders: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    showRowStripes: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    showColStripes: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    showLastColumn: bool | None = Field.attribute(expected_type=bool, allow_none=False)

    def __init__(self,
                 name=None,
                 showRowHeaders=False,
                 showColHeaders=False,
                 showRowStripes=False,
                 showColStripes=False,
                 showLastColumn=False,
                ):
        self.name = name
        self.showRowHeaders = showRowHeaders
        self.showColHeaders = showColHeaders
        self.showRowStripes = showRowStripes
        self.showColStripes = showColStripes
        self.showLastColumn = showLastColumn


class MemberList(TypedSerialisable):

    tagname = "members"

    level: int | None = Field.attribute(expected_type=int, allow_none=True)
    member: list[str] = Field.sequence(
        expected_type=str,
        primitive_attribute="name",
    )

    xml_order = ("member",)

    def __init__(self,
                 count=None,
                 level=None,
                 member=(),
                ):
        self.level = level
        self.member = list(member)

    @property
    def count(self):
        return len(self.member)

    def __iter__(self):
        yield from super().__iter__()
        yield "count", str(self.count)


class MemberProperty(TypedSerialisable):

    tagname = "mps"

    name: str | None = Field.attribute(expected_type=str, allow_none=True)
    showCell: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showTip: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showAsCaption: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    nameLen: int | None = Field.attribute(expected_type=int, allow_none=True)
    pPos: int | None = Field.attribute(expected_type=int, allow_none=True)
    pLen: int | None = Field.attribute(expected_type=int, allow_none=True)
    level: int | None = Field.attribute(expected_type=int, allow_none=True)
    field: int | None = Field.attribute(expected_type=int, allow_none=False)

    def __init__(self,
                 name=None,
                 showCell=None,
                 showTip=None,
                 showAsCaption=None,
                 nameLen=None,
                 pPos=None,
                 pLen=None,
                 level=None,
                 field=None,
                ):
        self.name = name
        self.showCell = showCell
        self.showTip = showTip
        self.showAsCaption = showAsCaption
        self.nameLen = nameLen
        self.pPos = pPos
        self.pLen = pLen
        self.level = level
        self.field = field


class PivotHierarchy(Serialisable):

    tagname = "pivotHierarchy"

    outline = Bool()
    multipleItemSelectionAllowed = Bool()
    subtotalTop = Bool()
    showInFieldList = Bool()
    dragToRow = Bool()
    dragToCol = Bool()
    dragToPage = Bool()
    dragToData = Bool()
    dragOff = Bool()
    includeNewItemsInFilter = Bool()
    caption = String(allow_none=True)
    mps = NestedSequence(expected_type=MemberProperty, count=True)
    members = Typed(expected_type=MemberList, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('mps', 'members',)

    def __init__(self,
                 outline=None,
                 multipleItemSelectionAllowed=None,
                 subtotalTop=None,
                 showInFieldList=None,
                 dragToRow=None,
                 dragToCol=None,
                 dragToPage=None,
                 dragToData=None,
                 dragOff=None,
                 includeNewItemsInFilter=None,
                 caption=None,
                 mps=(),
                 members=None,
                 extLst=None,
                ):
        self.outline = outline
        self.multipleItemSelectionAllowed = multipleItemSelectionAllowed
        self.subtotalTop = subtotalTop
        self.showInFieldList = showInFieldList
        self.dragToRow = dragToRow
        self.dragToCol = dragToCol
        self.dragToPage = dragToPage
        self.dragToData = dragToData
        self.dragOff = dragOff
        self.includeNewItemsInFilter = includeNewItemsInFilter
        self.caption = caption
        self.mps = mps
        self.members = members
        self.extLst = extLst


class Reference(TypedSerialisable):

    tagname = "reference"

    field: int | None = Field.attribute(expected_type=int, allow_none=True)
    selected: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    byPosition: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    relative: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    defaultSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    sumSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    countASubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    avgSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    maxSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    minSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    productSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    countSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    stdDevSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    stdDevPSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    varSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    varPSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    x: list[int] = Field.sequence(expected_type=int, primitive_attribute="v")
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True)

    xml_order = ("x",)

    def __init__(self,
                 field=None,
                 count=None,
                 selected=None,
                 byPosition=None,
                 relative=None,
                 defaultSubtotal=None,
                 sumSubtotal=None,
                 countASubtotal=None,
                 avgSubtotal=None,
                 maxSubtotal=None,
                 minSubtotal=None,
                 productSubtotal=None,
                 countSubtotal=None,
                 stdDevSubtotal=None,
                 stdDevPSubtotal=None,
                 varSubtotal=None,
                 varPSubtotal=None,
                 x=(),
                 extLst=None,
                ):
        self.field = field
        self.selected = selected
        self.byPosition = byPosition
        self.relative = relative
        self.defaultSubtotal = defaultSubtotal
        self.sumSubtotal = sumSubtotal
        self.countASubtotal = countASubtotal
        self.avgSubtotal = avgSubtotal
        self.maxSubtotal = maxSubtotal
        self.minSubtotal = minSubtotal
        self.productSubtotal = productSubtotal
        self.countSubtotal = countSubtotal
        self.stdDevSubtotal = stdDevSubtotal
        self.stdDevPSubtotal = stdDevPSubtotal
        self.varSubtotal = varSubtotal
        self.varPSubtotal = varPSubtotal
        self.x = list(x)
        self.extLst = extLst


    @property
    def count(self):
        return len(self.field)


class PivotArea(TypedSerialisable):

    tagname = "pivotArea"

    references: list[Reference] = Field.nested_sequence(expected_type=Reference, count=True)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True)
    field: int | None = Field.attribute(expected_type=int, allow_none=True)
    type: str | None = Field.attribute(expected_type=str, allow_none=False)
    dataOnly: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    labelOnly: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    grandRow: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    grandCol: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    cacheIndex: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    outline: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    offset: str | None = Field.attribute(expected_type=str, allow_none=True)
    collapsedLevelsAreSubtotals: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    axis: str | None = Field.attribute(expected_type=str, allow_none=True)
    fieldPosition: int | None = Field.attribute(expected_type=int, allow_none=True)

    xml_order = ("references",)

    def __init__(self,
                 references=(),
                 extLst=None,
                 field=None,
                 type="normal",
                 dataOnly=True,
                 labelOnly=None,
                 grandRow=None,
                 grandCol=None,
                 cacheIndex=None,
                 outline=True,
                 offset=None,
                 collapsedLevelsAreSubtotals=None,
                 axis=None,
                 fieldPosition=None,
                ):
        self.references = list(references)
        self.extLst = extLst
        self.field = field
        self.type = type
        self.dataOnly = dataOnly
        self.labelOnly = labelOnly
        self.grandRow = grandRow
        self.grandCol = grandCol
        self.cacheIndex = cacheIndex
        self.outline = outline
        self.offset = offset
        self.collapsedLevelsAreSubtotals = collapsedLevelsAreSubtotals
        self.axis = axis
        self.fieldPosition = fieldPosition


class ChartFormat(TypedSerialisable):

    tagname = "chartFormat"

    chart: int | None = Field.attribute(expected_type=int, allow_none=False)
    format: int | None = Field.attribute(expected_type=int, allow_none=False)
    series: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    pivotArea: PivotArea | None = Field.element(expected_type=PivotArea, allow_none=False)

    xml_order = ("pivotArea",)

    def __init__(self,
                 chart=None,
                 format=None,
                 series=None,
                 pivotArea=None,
                ):
        self.chart = chart
        self.format = format
        self.series = series
        self.pivotArea = pivotArea


class ConditionalFormat(TypedSerialisable):

    tagname = "conditionalFormat"

    scope: str | None = Field.attribute(expected_type=str, allow_none=False)
    type: str | None = Field.attribute(expected_type=str, allow_none=True)
    priority: int | None = Field.attribute(expected_type=int, allow_none=False)
    pivotAreas: list[PivotArea] = Field.nested_sequence(expected_type=PivotArea)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True)

    xml_order = ("pivotAreas",)

    def __init__(self,
                 scope="selection",
                 type=None,
                 priority=None,
                 pivotAreas=(),
                 extLst=None,
                ):
        self.scope = scope
        self.type = type
        self.priority = priority
        self.pivotAreas = list(pivotAreas)
        self.extLst = extLst


class ConditionalFormatList(TypedSerialisable):

    tagname = "conditionalFormats"

    conditionalFormat: list[ConditionalFormat] = Field.sequence(expected_type=ConditionalFormat)

    def __init__(self, conditionalFormat=(), count=None):
        self.conditionalFormat = list(conditionalFormat)


    def by_priority(self):
        """
        Return a dictionary of format objects keyed by (field id and format property).
        This can be used to map the formats to field but also to dedupe to match
        worksheet definitions which are grouped by cell range
        """

        fmts = {}
        for fmt in self.conditionalFormat:
            for area in fmt.pivotAreas:
                for ref in area.references:
                    for field in ref.x:
                        value = field.v if hasattr(field, "v") else field
                        key = (value, fmt.priority)
                        fmts[key] = fmt

        return fmts


    def _dedupe(self):
        """
        Group formats by field index and priority.
        Sorted to match sorting and grouping for corresponding worksheet formats

        The implemtenters notes contain significant deviance from the OOXML
        specification, in particular how conditional formats in tables relate to
        those defined in corresponding worksheets and how to determine which
        format applies to which fields.

        There are some magical interdependencies:

        * Every pivot table fmt must have a worksheet cxf with the same priority.

        * In the reference part the field 4294967294 refers to a data field, the
        spec says -2

        * Data fields are referenced by the 0-index reference.x.v value

        Things are made more complicated by the fact that field items behave
        diffently if the parent is a reference or shared item: "In Office if the
        parent is the reference element, then restrictions of this value are
        defined by reference@field. If the parent is the tables element, then
        this value specifies the index into the table tag position in @url."
        Yeah, right!
        """
        fmts = self.by_priority()
        # sort by priority in order, keeping the highest numerical priority, least when
        # actually applied
        # this is not documented but it's what Excel is happy with
        fmts = {field:fmt for (field, priority), fmt in sorted(fmts.items(), reverse=True)}
        #fmts = {field:fmt for (field, priority), fmt in fmts.items()}
        if fmts:
            self.conditionalFormat = list(fmts.values())


    @property
    def count(self):
        return len(self.conditionalFormat)

    def __iter__(self):
        yield from super().__iter__()
        yield "count", str(self.count)


    def to_tree(self, tagname=None):
        self._dedupe()
        return super().to_tree(tagname)


class Format(TypedSerialisable):

    tagname = "format"

    action: str | None = Field.attribute(expected_type=str, allow_none=False)
    dxfId: int | None = Field.attribute(expected_type=int, allow_none=True)
    pivotArea: PivotArea | None = Field.element(expected_type=PivotArea, allow_none=False)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True)

    xml_order = ("pivotArea",)

    def __init__(self,
                 action="formatting",
                 dxfId=None,
                 pivotArea=None,
                 extLst=None,
                ):
        self.action = action
        self.dxfId = dxfId
        self.pivotArea = pivotArea
        self.extLst = extLst


class DataField(TypedSerialisable):

    tagname = "dataField"

    name: str | None = Field.attribute(expected_type=str, allow_none=True)
    fld: int | None = Field.attribute(expected_type=int, allow_none=False)
    subtotal: str | None = Field.attribute(expected_type=str, allow_none=False)
    showDataAs: str | None = Field.attribute(expected_type=str, allow_none=False)
    baseField: int | None = Field.attribute(expected_type=int, allow_none=False)
    baseItem: int | None = Field.attribute(expected_type=int, allow_none=False)
    numFmtId: int | None = Field.attribute(expected_type=int, allow_none=True)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True)

    xml_order = ()


    def __init__(self,
                 name=None,
                 fld=None,
                 subtotal="sum",
                 showDataAs="normal",
                 baseField=-1,
                 baseItem=1048832,
                 numFmtId=None,
                 extLst=None,
                ):
        self.name = name
        self.fld = fld
        self.subtotal = subtotal
        self.showDataAs = showDataAs
        self.baseField = baseField
        self.baseItem = baseItem
        self.numFmtId = numFmtId
        self.extLst = extLst


class PageField(TypedSerialisable):

    tagname = "pageField"

    fld: int | None = Field.attribute(expected_type=int, allow_none=False)
    item: int | None = Field.attribute(expected_type=int, allow_none=True)
    hier: int | None = Field.attribute(expected_type=int, allow_none=True)
    name: str | None = Field.attribute(expected_type=str, allow_none=True)
    cap: str | None = Field.attribute(expected_type=str, allow_none=True)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True)

    xml_order = ()

    def __init__(self,
                 fld=None,
                 item=None,
                 hier=None,
                 name=None,
                 cap=None,
                 extLst=None,
                ):
        self.fld = fld
        self.item = item
        self.hier = hier
        self.name = name
        self.cap = cap
        self.extLst = extLst


class RowColItem(TypedSerialisable):

    tagname = "i"

    t: str | None = Field.attribute(expected_type=str, allow_none=False)
    r: int | None = Field.attribute(expected_type=int, allow_none=False)
    i: int | None = Field.attribute(expected_type=int, allow_none=False)
    x: list[int] = Field.sequence(expected_type=int, primitive_attribute="v")

    xml_order = ("x",)

    def __init__(self,
                 t="data",
                 r=0,
                 i=0,
                 x=(),
                ):
        self.t = t
        self.r = r
        self.i = i
        self.x = list(x)


class RowColField(TypedSerialisable):

    tagname = "field"

    x: int | None = Field.attribute(expected_type=int, allow_none=False)

    def __init__(self,
                 x=None,
                ):
        self.x = x


class AutoSortScope(TypedSerialisable):

    pivotArea: PivotArea | None = Field.element(expected_type=PivotArea, allow_none=False)

    xml_order = ("pivotArea",)

    def __init__(self,
                 pivotArea=None,
                ):
        self.pivotArea = pivotArea


class FieldItem(TypedSerialisable):

    tagname = "item"

    n: str | None = Field.attribute(expected_type=str, allow_none=True)
    t: str | None = Field.attribute(expected_type=str, allow_none=False)
    h: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    s: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    sd: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    f: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    m: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    c: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    x: int | None = Field.attribute(expected_type=int, allow_none=True)
    d: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    e: bool | None = Field.attribute(expected_type=bool, allow_none=True)

    def __init__(self,
                 n=None,
                 t="data",
                 h=None,
                 s=None,
                 sd=True,
                 f=None,
                 m=None,
                 c=None,
                 x=None,
                 d=None,
                 e=None,
                ):
        self.n = n
        self.t = t
        self.h = h
        self.s = s
        self.sd = sd
        self.f = f
        self.m = m
        self.c = c
        self.x = x
        self.d = d
        self.e = e


class PivotField(TypedSerialisable):

    tagname = "pivotField"

    items: list[FieldItem] = Field.nested_sequence(expected_type=FieldItem, count=True)
    autoSortScope: AutoSortScope | None = Field.element(expected_type=AutoSortScope, allow_none=True)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True)
    name: str | None = Field.attribute(expected_type=str, allow_none=True)
    axis: str | None = Field.attribute(expected_type=str, allow_none=True)
    dataField: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    subtotalCaption: str | None = Field.attribute(expected_type=str, allow_none=True)
    showDropDowns: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    hiddenLevel: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    uniqueMemberProperty: str | None = Field.attribute(expected_type=str, allow_none=True)
    compact: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    allDrilled: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    numFmtId: int | None = Field.attribute(expected_type=int, allow_none=True)
    outline: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    subtotalTop: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    dragToRow: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    dragToCol: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    multipleItemSelectionAllowed: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    dragToPage: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    dragToData: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    dragOff: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showAll: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    insertBlankRow: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    serverField: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    insertPageBreak: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    autoShow: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    topAutoShow: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    hideNewItems: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    measureFilter: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    includeNewItemsInFilter: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    itemPageCount: int | None = Field.attribute(expected_type=int, allow_none=True)
    sortType: str | None = Field.attribute(expected_type=str, allow_none=False)
    dataSourceSort: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    nonAutoSortDefault: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    rankBy: int | None = Field.attribute(expected_type=int, allow_none=True)
    defaultSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    sumSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    countASubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    avgSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    maxSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    minSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    productSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    countSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    stdDevSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    stdDevPSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    varSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    varPSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showPropCell: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showPropTip: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showPropAsCaption: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    defaultAttributeDrillState: bool | None = Field.attribute(expected_type=bool, allow_none=True)

    xml_order = ("items", "autoSortScope")

    def __init__(self,
                 items=(),
                 autoSortScope=None,
                 name=None,
                 axis=None,
                 dataField=None,
                 subtotalCaption=None,
                 showDropDowns=True,
                 hiddenLevel=None,
                 uniqueMemberProperty=None,
                 compact=True,
                 allDrilled=None,
                 numFmtId=None,
                 outline=True,
                 subtotalTop=True,
                 dragToRow=True,
                 dragToCol=True,
                 multipleItemSelectionAllowed=None,
                 dragToPage=True,
                 dragToData=True,
                 dragOff=True,
                 showAll=True,
                 insertBlankRow=None,
                 serverField=None,
                 insertPageBreak=None,
                 autoShow=None,
                 topAutoShow=True,
                 hideNewItems=None,
                 measureFilter=None,
                 includeNewItemsInFilter=None,
                 itemPageCount=10,
                 sortType="manual",
                 dataSourceSort=None,
                 nonAutoSortDefault=None,
                 rankBy=None,
                 defaultSubtotal=True,
                 sumSubtotal=None,
                 countASubtotal=None,
                 avgSubtotal=None,
                 maxSubtotal=None,
                 minSubtotal=None,
                 productSubtotal=None,
                 countSubtotal=None,
                 stdDevSubtotal=None,
                 stdDevPSubtotal=None,
                 varSubtotal=None,
                 varPSubtotal=None,
                 showPropCell=None,
                 showPropTip=None,
                 showPropAsCaption=None,
                 defaultAttributeDrillState=None,
                 extLst=None,
                ):
        self.items = list(items)
        self.autoSortScope = autoSortScope
        self.name = name
        self.axis = axis
        self.dataField = dataField
        self.subtotalCaption = subtotalCaption
        self.showDropDowns = showDropDowns
        self.hiddenLevel = hiddenLevel
        self.uniqueMemberProperty = uniqueMemberProperty
        self.compact = compact
        self.allDrilled = allDrilled
        self.numFmtId = numFmtId
        self.outline = outline
        self.subtotalTop = subtotalTop
        self.dragToRow = dragToRow
        self.dragToCol = dragToCol
        self.multipleItemSelectionAllowed = multipleItemSelectionAllowed
        self.dragToPage = dragToPage
        self.dragToData = dragToData
        self.dragOff = dragOff
        self.showAll = showAll
        self.insertBlankRow = insertBlankRow
        self.serverField = serverField
        self.insertPageBreak = insertPageBreak
        self.autoShow = autoShow
        self.topAutoShow = topAutoShow
        self.hideNewItems = hideNewItems
        self.measureFilter = measureFilter
        self.includeNewItemsInFilter = includeNewItemsInFilter
        self.itemPageCount = itemPageCount
        self.sortType = sortType
        self.dataSourceSort = dataSourceSort
        self.nonAutoSortDefault = nonAutoSortDefault
        self.rankBy = rankBy
        self.defaultSubtotal = defaultSubtotal
        self.sumSubtotal = sumSubtotal
        self.countASubtotal = countASubtotal
        self.avgSubtotal = avgSubtotal
        self.maxSubtotal = maxSubtotal
        self.minSubtotal = minSubtotal
        self.productSubtotal = productSubtotal
        self.countSubtotal = countSubtotal
        self.stdDevSubtotal = stdDevSubtotal
        self.stdDevPSubtotal = stdDevPSubtotal
        self.varSubtotal = varSubtotal
        self.varPSubtotal = varPSubtotal
        self.showPropCell = showPropCell
        self.showPropTip = showPropTip
        self.showPropAsCaption = showPropAsCaption
        self.defaultAttributeDrillState = defaultAttributeDrillState
        self.extLst = extLst


class Location(TypedSerialisable):

    tagname = "location"

    ref: str | None = Field.attribute(expected_type=str, allow_none=False)
    firstHeaderRow: int | None = Field.attribute(expected_type=int, allow_none=False)
    firstDataRow: int | None = Field.attribute(expected_type=int, allow_none=False)
    firstDataCol: int | None = Field.attribute(expected_type=int, allow_none=False)
    rowPageCount: int | None = Field.attribute(expected_type=int, allow_none=True)
    colPageCount: int | None = Field.attribute(expected_type=int, allow_none=True)

    def __init__(self,
                 ref=None,
                 firstHeaderRow=None,
                 firstDataRow=None,
                 firstDataCol=None,
                 rowPageCount=None,
                 colPageCount=None,
                ):
        self.ref = ref
        self.firstHeaderRow = firstHeaderRow
        self.firstDataRow = firstDataRow
        self.firstDataCol = firstDataCol
        self.rowPageCount = rowPageCount
        self.colPageCount = colPageCount


class TableDefinition(TypedSerialisable):

    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"
    rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"
    _id = 1
    _path = "/xl/pivotTables/pivotTable{0}.xml"

    tagname = "pivotTableDefinition"
    cache = None

    name: str | None = Field.attribute(expected_type=str, allow_none=False)
    cacheId: int | None = Field.attribute(expected_type=int, allow_none=False)
    dataOnRows: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    dataPosition: int | None = Field.attribute(expected_type=int, allow_none=True)
    dataCaption: str | None = Field.attribute(expected_type=str, allow_none=False)
    grandTotalCaption: str | None = Field.attribute(expected_type=str, allow_none=True)
    errorCaption: str | None = Field.attribute(expected_type=str, allow_none=True)
    showError: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    missingCaption: str | None = Field.attribute(expected_type=str, allow_none=True)
    showMissing: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    pageStyle: str | None = Field.attribute(expected_type=str, allow_none=True)
    pivotTableStyle: str | None = Field.attribute(expected_type=str, allow_none=True)
    vacatedStyle: str | None = Field.attribute(expected_type=str, allow_none=True)
    tag: str | None = Field.attribute(expected_type=str, allow_none=True)
    updatedVersion: int | None = Field.attribute(expected_type=int, allow_none=False)
    minRefreshableVersion: int | None = Field.attribute(expected_type=int, allow_none=False)
    asteriskTotals: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    showItems: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    editData: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    disableFieldList: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    showCalcMbrs: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    visualTotals: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    showMultipleLabel: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    showDataDropDown: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    showDrill: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    printDrill: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    showMemberPropertyTips: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    showDataTips: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    enableWizard: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    enableDrill: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    enableFieldProperties: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    preserveFormatting: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    useAutoFormatting: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    pageWrap: int | None = Field.attribute(expected_type=int, allow_none=False)
    pageOverThenDown: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    subtotalHiddenItems: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    rowGrandTotals: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    colGrandTotals: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    fieldPrintTitles: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    itemPrintTitles: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    mergeItem: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    showDropZones: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    createdVersion: int | None = Field.attribute(expected_type=int, allow_none=False)
    indent: int | None = Field.attribute(expected_type=int, allow_none=False)
    showEmptyRow: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    showEmptyCol: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    showHeaders: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    compact: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    outline: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    outlineData: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    compactData: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    published: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    gridDropZones: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    immersive: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    multipleFieldFilters: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    chartFormat: int | None = Field.attribute(expected_type=int, allow_none=False)
    rowHeaderCaption: str | None = Field.attribute(expected_type=str, allow_none=True)
    colHeaderCaption: str | None = Field.attribute(expected_type=str, allow_none=True)
    fieldListSortAscending: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    mdxSubqueries: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    customListSort: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    autoFormatId: int | None = Field.attribute(expected_type=int, allow_none=True)
    applyNumberFormats: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    applyBorderFormats: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    applyFontFormats: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    applyPatternFormats: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    applyAlignmentFormats: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    applyWidthHeightFormats: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    location: Location | None = Field.element(expected_type=Location, allow_none=False)
    pivotFields: list[PivotField] = Field.nested_sequence(expected_type=PivotField, count=True)
    rowFields: list[RowColField] = Field.nested_sequence(expected_type=RowColField, count=True)
    rowItems: list[RowColItem] = Field.nested_sequence(expected_type=RowColItem, count=True)
    colFields: list[RowColField] = Field.nested_sequence(expected_type=RowColField, count=True)
    colItems: list[RowColItem] = Field.nested_sequence(expected_type=RowColItem, count=True)
    pageFields: list[PageField] = Field.nested_sequence(expected_type=PageField, count=True)
    dataFields: list[DataField] = Field.nested_sequence(expected_type=DataField, count=True)
    formats: list[Format] = Field.nested_sequence(expected_type=Format, count=True)
    conditionalFormats: ConditionalFormatList | None = Field.element(expected_type=ConditionalFormatList, allow_none=True)
    chartFormats: list[ChartFormat] = Field.nested_sequence(expected_type=ChartFormat, count=True)
    pivotHierarchies: list[PivotHierarchy] = Field.nested_sequence(expected_type=PivotHierarchy, count=True)
    pivotTableStyleInfo: PivotTableStyle | None = Field.element(expected_type=PivotTableStyle, allow_none=True)
    filters: list[PivotFilter] = Field.nested_sequence(expected_type=PivotFilter, count=True)
    rowHierarchiesUsage: RowHierarchiesUsage | None = Field.element(expected_type=RowHierarchiesUsage, allow_none=True)
    colHierarchiesUsage: ColHierarchiesUsage | None = Field.element(expected_type=ColHierarchiesUsage, allow_none=True)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True)
    id: str | None = Field.attribute(expected_type=str, allow_none=True, namespace=REL_NS)

    xml_order = (
        "location",
        "pivotFields",
        "rowFields",
        "rowItems",
        "colFields",
        "colItems",
        "pageFields",
        "dataFields",
        "formats",
        "conditionalFormats",
        "chartFormats",
        "pivotHierarchies",
        "pivotTableStyleInfo",
        "filters",
        "rowHierarchiesUsage",
        "colHierarchiesUsage",
    )

    def __init__(self,
                 name=None,
                 cacheId=None,
                 dataOnRows=False,
                 dataPosition=None,
                 dataCaption=None,
                 grandTotalCaption=None,
                 errorCaption=None,
                 showError=False,
                 missingCaption=None,
                 showMissing=True,
                 pageStyle=None,
                 pivotTableStyle=None,
                 vacatedStyle=None,
                 tag=None,
                 updatedVersion=0,
                 minRefreshableVersion=0,
                 asteriskTotals=False,
                 showItems=True,
                 editData=False,
                 disableFieldList=False,
                 showCalcMbrs=True,
                 visualTotals=True,
                 showMultipleLabel=True,
                 showDataDropDown=True,
                 showDrill=True,
                 printDrill=False,
                 showMemberPropertyTips=True,
                 showDataTips=True,
                 enableWizard=True,
                 enableDrill=True,
                 enableFieldProperties=True,
                 preserveFormatting=True,
                 useAutoFormatting=False,
                 pageWrap=0,
                 pageOverThenDown=False,
                 subtotalHiddenItems=False,
                 rowGrandTotals=True,
                 colGrandTotals=True,
                 fieldPrintTitles=False,
                 itemPrintTitles=False,
                 mergeItem=False,
                 showDropZones=True,
                 createdVersion=0,
                 indent=1,
                 showEmptyRow=False,
                 showEmptyCol=False,
                 showHeaders=True,
                 compact=True,
                 outline=False,
                 outlineData=False,
                 compactData=True,
                 published=False,
                 gridDropZones=False,
                 immersive=True,
                 multipleFieldFilters=False,
                 chartFormat=0,
                 rowHeaderCaption=None,
                 colHeaderCaption=None,
                 fieldListSortAscending=False,
                 mdxSubqueries=False,
                 customListSort=None,
                 autoFormatId=None,
                 applyNumberFormats=False,
                 applyBorderFormats=False,
                 applyFontFormats=False,
                 applyPatternFormats=False,
                 applyAlignmentFormats=False,
                 applyWidthHeightFormats=False,
                 location=None,
                 pivotFields=(),
                 rowFields=(),
                 rowItems=(),
                 colFields=(),
                 colItems=(),
                 pageFields=(),
                 dataFields=(),
                 formats=(),
                 conditionalFormats=None,
                 chartFormats=(),
                 pivotHierarchies=(),
                 pivotTableStyleInfo=None,
                 filters=(),
                 rowHierarchiesUsage=None,
                 colHierarchiesUsage=None,
                 extLst=None,
                 id=None,
                ):
        self.name = name
        self.cacheId = cacheId
        self.dataOnRows = dataOnRows
        self.dataPosition = dataPosition
        self.dataCaption = dataCaption
        self.grandTotalCaption = grandTotalCaption
        self.errorCaption = errorCaption
        self.showError = showError
        self.missingCaption = missingCaption
        self.showMissing = showMissing
        self.pageStyle = pageStyle
        self.pivotTableStyle = pivotTableStyle
        self.vacatedStyle = vacatedStyle
        self.tag = tag
        self.updatedVersion = updatedVersion
        self.minRefreshableVersion = minRefreshableVersion
        self.asteriskTotals = asteriskTotals
        self.showItems = showItems
        self.editData = editData
        self.disableFieldList = disableFieldList
        self.showCalcMbrs = showCalcMbrs
        self.visualTotals = visualTotals
        self.showMultipleLabel = showMultipleLabel
        self.showDataDropDown = showDataDropDown
        self.showDrill = showDrill
        self.printDrill = printDrill
        self.showMemberPropertyTips = showMemberPropertyTips
        self.showDataTips = showDataTips
        self.enableWizard = enableWizard
        self.enableDrill = enableDrill
        self.enableFieldProperties = enableFieldProperties
        self.preserveFormatting = preserveFormatting
        self.useAutoFormatting = useAutoFormatting
        self.pageWrap = pageWrap
        self.pageOverThenDown = pageOverThenDown
        self.subtotalHiddenItems = subtotalHiddenItems
        self.rowGrandTotals = rowGrandTotals
        self.colGrandTotals = colGrandTotals
        self.fieldPrintTitles = fieldPrintTitles
        self.itemPrintTitles = itemPrintTitles
        self.mergeItem = mergeItem
        self.showDropZones = showDropZones
        self.createdVersion = createdVersion
        self.indent = indent
        self.showEmptyRow = showEmptyRow
        self.showEmptyCol = showEmptyCol
        self.showHeaders = showHeaders
        self.compact = compact
        self.outline = outline
        self.outlineData = outlineData
        self.compactData = compactData
        self.published = published
        self.gridDropZones = gridDropZones
        self.immersive = immersive
        self.multipleFieldFilters = multipleFieldFilters
        self.chartFormat = chartFormat
        self.rowHeaderCaption = rowHeaderCaption
        self.colHeaderCaption = colHeaderCaption
        self.fieldListSortAscending = fieldListSortAscending
        self.mdxSubqueries = mdxSubqueries
        self.customListSort = customListSort
        self.autoFormatId = autoFormatId
        self.applyNumberFormats = applyNumberFormats
        self.applyBorderFormats = applyBorderFormats
        self.applyFontFormats = applyFontFormats
        self.applyPatternFormats = applyPatternFormats
        self.applyAlignmentFormats = applyAlignmentFormats
        self.applyWidthHeightFormats = applyWidthHeightFormats
        self.location = location
        self.pivotFields = list(pivotFields)
        self.rowFields = list(rowFields)
        self.rowItems = list(rowItems)
        self.colFields = list(colFields)
        self.colItems = list(colItems)
        self.pageFields = list(pageFields)
        self.dataFields = list(dataFields)
        self.formats = list(formats)
        self.conditionalFormats = conditionalFormats
        self.conditionalFormats = None
        self.chartFormats = list(chartFormats)
        self.pivotHierarchies = list(pivotHierarchies)
        self.pivotTableStyleInfo = pivotTableStyleInfo
        self.filters = list(filters)
        self.rowHierarchiesUsage = rowHierarchiesUsage
        self.colHierarchiesUsage = colHierarchiesUsage
        self.extLst = extLst
        self.id = id


    def to_tree(self):
        tree = super().to_tree()
        tree.set("xmlns", SHEET_MAIN_NS)
        return tree


    @property
    def path(self):
        return self._path.format(self._id)


    def _write(self, archive, manifest):
        """
        Add to zipfile and update manifest
        """
        self._write_rels(archive, manifest)
        xml = tostring(self.to_tree())
        archive.writestr(self.path[1:], xml)
        manifest.append(self)


    def _write_rels(self, archive, manifest):
        """
        Write the relevant child objects and add links
        """
        if self.cache is None:
            return

        rels = RelationshipList()
        r = Relationship(Type=self.cache.rel_type, Target=self.cache.path)
        rels.append(r)
        self.id = r.id
        if self.cache.path[1:] not in archive.namelist():
            self.cache._write(archive, manifest)

        path = get_rels_path(self.path)
        xml = tostring(rels.to_tree())
        archive.writestr(path[1:], xml)


    def formatted_fields(self):
        """Map fields to associated conditional formats by priority"""
        if not self.conditionalFormats:
            return {}
        fields = defaultdict(list)
        for idx, prio in self.conditionalFormats.by_priority():
            name = self.dataFields[idx].name
            fields[name].append(prio)
        return fields


    @property
    def summary(self):
        """
        Provide a simplified summary of the table
        """

        return f"{self.name} {dict(self.location)}"
