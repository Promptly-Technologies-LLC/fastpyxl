# Copyright (c) 2010-2024 fastpyxl


from collections import defaultdict

from fastpyxl.descriptors.excel import ExtensionList
from fastpyxl.typed_serialisable.base import Serialisable as TypedSerialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import Field
from fastpyxl.xml.constants import REL_NS, SHEET_MAIN_NS
from fastpyxl.xml.functions import tostring
from fastpyxl.packaging.relationship import (
    RelationshipList,
    Relationship,
    get_rels_path
)

from fastpyxl.worksheet.filters import (
    AutoFilter,
)


class HierarchyUsage(TypedSerialisable):

    tagname = "hierarchyUsage"

    hierarchyUsage: int | None = Field.attribute(expected_type=int, allow_none=False, default=None)

    def __init__(self,
                 hierarchyUsage=None,
                ):
        self.hierarchyUsage = hierarchyUsage


class ColHierarchiesUsage(TypedSerialisable):

    tagname = "colHierarchiesUsage"

    colHierarchyUsage: list[HierarchyUsage] = Field.sequence(expected_type=HierarchyUsage, default=list)

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

    rowHierarchyUsage: list[HierarchyUsage] = Field.sequence(expected_type=HierarchyUsage, default=list)

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


PIVOT_FILTER_TYPES = frozenset(
    (
        "unknown",
        "count",
        "percent",
        "sum",
        "captionEqual",
        "captionNotEqual",
        "captionBeginsWith",
        "captionNotBeginsWith",
        "captionEndsWith",
        "captionNotEndsWith",
        "captionContains",
        "captionNotContains",
        "captionGreaterThan",
        "captionGreaterThanOrEqual",
        "captionLessThan",
        "captionLessThanOrEqual",
        "captionBetween",
        "captionNotBetween",
        "valueEqual",
        "valueNotEqual",
        "valueGreaterThan",
        "valueGreaterThanOrEqual",
        "valueLessThan",
        "valueLessThanOrEqual",
        "valueBetween",
        "valueNotBetween",
        "dateEqual",
        "dateNotEqual",
        "dateOlderThan",
        "dateOlderThanOrEqual",
        "dateNewerThan",
        "dateNewerThanOrEqual",
        "dateBetween",
        "dateNotBetween",
        "tomorrow",
        "today",
        "yesterday",
        "nextWeek",
        "thisWeek",
        "lastWeek",
        "nextMonth",
        "thisMonth",
        "lastMonth",
        "nextQuarter",
        "thisQuarter",
        "lastQuarter",
        "nextYear",
        "thisYear",
        "lastYear",
        "yearToDate",
        "Q1",
        "Q2",
        "Q3",
        "Q4",
        "M1",
        "M2",
        "M3",
        "M4",
        "M5",
        "M6",
        "M7",
        "M8",
        "M9",
        "M10",
        "M11",
        "M12",
    )
)


def _pivot_filter_type(v):
    if v is None:
        return None
    if v not in PIVOT_FILTER_TYPES:
        raise FieldValidationError(f"type rejected value {v!r}")
    return v


class PivotFilter(TypedSerialisable):

    tagname = "filter"

    fld: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    mpFld: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    type: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=_pivot_filter_type, default=None,
    )
    evalOrder: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    id: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    iMeasureHier: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    iMeasureFld: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    name: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    description: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    stringValue1: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    stringValue2: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    autoFilter: AutoFilter | None = Field.element(
        expected_type=AutoFilter,
        allow_none=True, default=None,
    )
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList,
        allow_none=True,
        serialize=False, default=None,
    )

    xml_order = ("autoFilter",)

    def __init__(
        self,
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
        self.extLst = extLst


class PivotFilters(TypedSerialisable):

    tagname = "filters"

    filter: PivotFilter | None = Field.element(expected_type=PivotFilter, allow_none=True, default=None)

    xml_order = ("filter",)

    def __init__(self, count=None, filter=None):
        del count
        self.filter = filter

    @property
    def count(self):
        return 1 if self.filter is not None else 0

    def __iter__(self):
        yield from super().__iter__()
        yield "count", str(self.count)


class PivotTableStyle(TypedSerialisable):

    tagname = "pivotTableStyleInfo"

    name: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    showRowHeaders: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    showColHeaders: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    showRowStripes: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    showColStripes: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    showLastColumn: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)

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

    level: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    member: list[str] = Field.sequence(
        expected_type=str,
        primitive_attribute="name", default=list,
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

    name: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    showCell: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showTip: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showAsCaption: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    nameLen: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    pPos: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    pLen: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    level: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    field: int | None = Field.attribute(expected_type=int, allow_none=False, default=None)

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


class PivotHierarchy(TypedSerialisable):

    tagname = "pivotHierarchy"

    outline: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    multipleItemSelectionAllowed: bool | None = Field.attribute(
        expected_type=bool,
        allow_none=True, default=None,
    )
    subtotalTop: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showInFieldList: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    dragToRow: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    dragToCol: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    dragToPage: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    dragToData: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    dragOff: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    includeNewItemsInFilter: bool | None = Field.attribute(
        expected_type=bool,
        allow_none=True, default=None,
    )
    caption: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    mps: list[MemberProperty] = Field.nested_sequence(
        expected_type=MemberProperty,
        count=True,
        default=list,
    )
    members: MemberList | None = Field.element(expected_type=MemberList, allow_none=True, default=None)
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList,
        allow_none=True,
        serialize=False, default=None,
    )

    xml_order = ("mps", "members")

    def __init__(
        self,
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
        self.mps = list(mps) if mps is not None else []
        self.members = members
        self.extLst = extLst


class Reference(TypedSerialisable):

    tagname = "reference"

    field: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    selected: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    byPosition: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    relative: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    defaultSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    sumSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    countASubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    avgSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    maxSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    minSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    productSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    countSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    stdDevSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    stdDevPSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    varSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    varPSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    x: list[int] = Field.sequence(expected_type=int, primitive_attribute="v", default=list)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True, default=None)

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
        return len(self.x)


class PivotArea(TypedSerialisable):

    tagname = "pivotArea"

    references: list[Reference] = Field.nested_sequence(expected_type=Reference, count=True, default=list)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True, default=None)
    field: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    type: str | None = Field.attribute(expected_type=str, allow_none=False, default=None)
    dataOnly: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    labelOnly: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    grandRow: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    grandCol: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    cacheIndex: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    outline: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    offset: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    collapsedLevelsAreSubtotals: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    axis: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    fieldPosition: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

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

    chart: int | None = Field.attribute(expected_type=int, allow_none=False, default=None)
    format: int | None = Field.attribute(expected_type=int, allow_none=False, default=None)
    series: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    pivotArea: PivotArea | None = Field.element(expected_type=PivotArea, allow_none=False, default=None)

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

    scope: str | None = Field.attribute(expected_type=str, allow_none=False, default=None)
    type: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    priority: int | None = Field.attribute(expected_type=int, allow_none=False, default=None)
    pivotAreas: list[PivotArea] = Field.nested_sequence(expected_type=PivotArea, default=list)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True, default=None)

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

    conditionalFormat: list[ConditionalFormat] = Field.sequence(expected_type=ConditionalFormat, default=list)

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
        best: dict[int, tuple[int, ConditionalFormat]] = {}
        for fmt in self.conditionalFormat:
            prio = fmt.priority
            if prio is None:
                continue
            for area in fmt.pivotAreas:
                for ref in area.references:
                    for v in ref.x:
                        if v not in best or prio > best[v][0]:
                            best[v] = (prio, fmt)
        if best:
            ordered = sorted(best.values(), key=lambda t: -t[0])
            self.conditionalFormat = [fmt for _, fmt in ordered]


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

    action: str | None = Field.attribute(expected_type=str, allow_none=False, default=None)
    dxfId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    pivotArea: PivotArea | None = Field.element(expected_type=PivotArea, allow_none=False, default=None)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True, default=None)

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

    name: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    fld: int | None = Field.attribute(expected_type=int, allow_none=False, default=None)
    subtotal: str | None = Field.attribute(expected_type=str, allow_none=False, default=None)
    showDataAs: str | None = Field.attribute(expected_type=str, allow_none=False, default=None)
    baseField: int | None = Field.attribute(expected_type=int, allow_none=False, default=None)
    baseItem: int | None = Field.attribute(expected_type=int, allow_none=False, default=None)
    numFmtId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True, default=None)

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

    fld: int | None = Field.attribute(expected_type=int, allow_none=False, default=None)
    item: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    hier: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    name: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    cap: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True, default=None)

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

    t: str | None = Field.attribute(expected_type=str, allow_none=False, default=None)
    r: int | None = Field.attribute(expected_type=int, allow_none=False, default=None)
    i: int | None = Field.attribute(expected_type=int, allow_none=False, default=None)
    x: list[int] = Field.sequence(expected_type=int, primitive_attribute="v", default=list)

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

    x: int | None = Field.attribute(expected_type=int, allow_none=False, default=None)

    def __init__(self,
                 x=None,
                ):
        self.x = x


class AutoSortScope(TypedSerialisable):

    pivotArea: PivotArea | None = Field.element(expected_type=PivotArea, allow_none=False, default=None)

    xml_order = ("pivotArea",)

    def __init__(self,
                 pivotArea=None,
                ):
        self.pivotArea = pivotArea


class FieldItem(TypedSerialisable):

    tagname = "item"

    n: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    t: str | None = Field.attribute(expected_type=str, allow_none=False, default=None)
    h: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    s: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    sd: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    f: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    m: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    c: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    x: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    d: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    e: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)

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

    items: list[FieldItem] = Field.nested_sequence(expected_type=FieldItem, count=True, default=list)
    autoSortScope: AutoSortScope | None = Field.element(expected_type=AutoSortScope, allow_none=True, default=None)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True, default=None)
    name: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    axis: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    dataField: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    subtotalCaption: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    showDropDowns: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    hiddenLevel: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    uniqueMemberProperty: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    compact: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    allDrilled: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    numFmtId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    outline: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    subtotalTop: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    dragToRow: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    dragToCol: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    multipleItemSelectionAllowed: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    dragToPage: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    dragToData: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    dragOff: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showAll: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    insertBlankRow: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    serverField: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    insertPageBreak: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    autoShow: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    topAutoShow: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    hideNewItems: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    measureFilter: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    includeNewItemsInFilter: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    itemPageCount: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    sortType: str | None = Field.attribute(expected_type=str, allow_none=False, default=None)
    dataSourceSort: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    nonAutoSortDefault: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    rankBy: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    defaultSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    sumSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    countASubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    avgSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    maxSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    minSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    productSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    countSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    stdDevSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    stdDevPSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    varSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    varPSubtotal: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showPropCell: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showPropTip: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    showPropAsCaption: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    defaultAttributeDrillState: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)

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

    ref: str | None = Field.attribute(expected_type=str, allow_none=False, default=None)
    firstHeaderRow: int | None = Field.attribute(expected_type=int, allow_none=False, default=None)
    firstDataRow: int | None = Field.attribute(expected_type=int, allow_none=False, default=None)
    firstDataCol: int | None = Field.attribute(expected_type=int, allow_none=False, default=None)
    rowPageCount: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    colPageCount: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

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

    name: str | None = Field.attribute(expected_type=str, allow_none=False, default=None)
    cacheId: int | None = Field.attribute(expected_type=int, allow_none=False, default=None)
    dataOnRows: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    dataPosition: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    dataCaption: str | None = Field.attribute(expected_type=str, allow_none=False, default=None)
    grandTotalCaption: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    errorCaption: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    showError: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    missingCaption: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    showMissing: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    pageStyle: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    pivotTableStyle: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    vacatedStyle: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    tag: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    updatedVersion: int | None = Field.attribute(expected_type=int, allow_none=False, default=None)
    minRefreshableVersion: int | None = Field.attribute(expected_type=int, allow_none=False, default=None)
    asteriskTotals: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    showItems: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    editData: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    disableFieldList: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    showCalcMbrs: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    visualTotals: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    showMultipleLabel: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    showDataDropDown: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    showDrill: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    printDrill: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    showMemberPropertyTips: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    showDataTips: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    enableWizard: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    enableDrill: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    enableFieldProperties: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    preserveFormatting: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    useAutoFormatting: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    pageWrap: int | None = Field.attribute(expected_type=int, allow_none=False, default=None)
    pageOverThenDown: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    subtotalHiddenItems: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    rowGrandTotals: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    colGrandTotals: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    fieldPrintTitles: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    itemPrintTitles: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    mergeItem: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    showDropZones: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    createdVersion: int | None = Field.attribute(expected_type=int, allow_none=False, default=None)
    indent: int | None = Field.attribute(expected_type=int, allow_none=False, default=None)
    showEmptyRow: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    showEmptyCol: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    showHeaders: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    compact: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    outline: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    outlineData: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    compactData: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    published: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    gridDropZones: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    immersive: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    multipleFieldFilters: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    chartFormat: int | None = Field.attribute(expected_type=int, allow_none=False, default=None)
    rowHeaderCaption: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    colHeaderCaption: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    fieldListSortAscending: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    mdxSubqueries: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    customListSort: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    autoFormatId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    applyNumberFormats: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    applyBorderFormats: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    applyFontFormats: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    applyPatternFormats: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    applyAlignmentFormats: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    applyWidthHeightFormats: bool | None = Field.attribute(expected_type=bool, allow_none=False, default=None)
    location: Location | None = Field.element(expected_type=Location, allow_none=False, default=None)
    pivotFields: list[PivotField] = Field.nested_sequence(expected_type=PivotField, count=True, default=list)
    rowFields: list[RowColField] = Field.nested_sequence(expected_type=RowColField, count=True, default=list)
    rowItems: list[RowColItem] = Field.nested_sequence(expected_type=RowColItem, count=True, default=list)
    colFields: list[RowColField] = Field.nested_sequence(expected_type=RowColField, count=True, default=list)
    colItems: list[RowColItem] = Field.nested_sequence(expected_type=RowColItem, count=True, default=list)
    pageFields: list[PageField] = Field.nested_sequence(expected_type=PageField, count=True, default=list)
    dataFields: list[DataField] = Field.nested_sequence(expected_type=DataField, count=True, default=list)
    formats: list[Format] = Field.nested_sequence(expected_type=Format, count=True, default=list)
    conditionalFormats: ConditionalFormatList | None = Field.element(expected_type=ConditionalFormatList, allow_none=True, default=None)
    chartFormats: list[ChartFormat] = Field.nested_sequence(expected_type=ChartFormat, count=True, default=list)
    pivotHierarchies: list[PivotHierarchy] = Field.nested_sequence(expected_type=PivotHierarchy, count=True, default=list)
    pivotTableStyleInfo: PivotTableStyle | None = Field.element(expected_type=PivotTableStyle, allow_none=True, default=None)
    filters: list[PivotFilter] = Field.nested_sequence(expected_type=PivotFilter, count=True, default=list)
    rowHierarchiesUsage: RowHierarchiesUsage | None = Field.element(expected_type=RowHierarchiesUsage, allow_none=True, default=None)
    colHierarchiesUsage: ColHierarchiesUsage | None = Field.element(expected_type=ColHierarchiesUsage, allow_none=True, default=None)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True, default=None)
    id: str | None = Field.attribute(expected_type=str, allow_none=True, namespace=REL_NS, default=None)

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

        loc = self.location
        return f"{self.name} {dict(iter(loc))}" if loc is not None else str(self.name)
