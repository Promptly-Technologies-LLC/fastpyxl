# Copyright (c) 2010-2024 fastpyxl

from datetime import datetime

from fastpyxl.descriptors.excel import (
    ExtensionList,
)
from fastpyxl.descriptors.nested import NestedInteger
from fastpyxl.xml.constants import REL_NS, SHEET_MAIN_NS
from fastpyxl.xml.functions import tostring
from fastpyxl.typed_serialisable.base import Serialisable as TypedSerialisable
from fastpyxl.typed_serialisable.fields import Field
from fastpyxl.packaging.relationship import (
    RelationshipList,
    Relationship,
    get_rels_path
)

from .table import (
    PivotArea,
)
from .fields import (
    Boolean,
    Error,
    Missing,
    Number,
    Text,
    TupleList,
    DateTimeField,
)


def _coerce_iso_datetime(value):
    if value is None or isinstance(value, datetime):
        return value
    return datetime.fromisoformat(value)


class MeasureDimensionMap(TypedSerialisable):

    tagname = "map"

    measureGroup: int | None = Field.attribute(expected_type=int, allow_none=True)
    dimension: int | None = Field.attribute(expected_type=int, allow_none=True)

    def __init__(self,
                 measureGroup=None,
                 dimension=None,
                ):
        self.measureGroup = measureGroup
        self.dimension = dimension


class MeasureGroup(TypedSerialisable):

    tagname = "measureGroup"

    name: str | None = Field.attribute(expected_type=str, allow_none=False)
    caption: str | None = Field.attribute(expected_type=str, allow_none=False)

    def __init__(self,
                 name=None,
                 caption=None,
                ):
        self.name = name
        self.caption = caption


class PivotDimension(TypedSerialisable):

    tagname = "dimension"

    measure: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    name: str | None = Field.attribute(expected_type=str, allow_none=False)
    uniqueName: str | None = Field.attribute(expected_type=str, allow_none=False)
    caption: str | None = Field.attribute(expected_type=str, allow_none=False)

    def __init__(self,
                 measure=None,
                 name=None,
                 uniqueName=None,
                 caption=None,
                ):
        self.measure = measure
        self.name = name
        self.uniqueName = uniqueName
        self.caption = caption


class CalculatedMember(TypedSerialisable):

    tagname = "calculatedMember"

    name: str | None = Field.attribute(expected_type=str, allow_none=False)
    mdx: str | None = Field.attribute(expected_type=str, allow_none=False)
    memberName: str | None = Field.attribute(expected_type=str, allow_none=True)
    hierarchy: str | None = Field.attribute(expected_type=str, allow_none=True)
    parent: str | None = Field.attribute(expected_type=str, allow_none=True)
    solveOrder: int | None = Field.attribute(expected_type=int, allow_none=True)
    set: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True)

    xml_order = ()

    def __init__(self,
                 name=None,
                 mdx=None,
                 memberName=None,
                 hierarchy=None,
                 parent=None,
                 solveOrder=None,
                 set=None,
                 extLst=None,
                ):
        self.name = name
        self.mdx = mdx
        self.memberName = memberName
        self.hierarchy = hierarchy
        self.parent = parent
        self.solveOrder = solveOrder
        self.set = set
        self.extLst = extLst


class CalculatedItem(TypedSerialisable):

    tagname = "calculatedItem"

    field: int | None = Field.attribute(expected_type=int, allow_none=True)
    formula: str | None = Field.attribute(expected_type=str, allow_none=False)
    pivotArea: PivotArea | None = Field.element(expected_type=PivotArea, allow_none=False)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True)

    xml_order = ("pivotArea", "extLst")

    def __init__(self,
                 field=None,
                 formula=None,
                 pivotArea=None,
                 extLst=None,
                ):
        self.field = field
        self.formula = formula
        self.pivotArea = pivotArea
        self.extLst = extLst


class ServerFormat(TypedSerialisable):

    tagname = "serverFormat"

    culture: str | None = Field.attribute(expected_type=str, allow_none=True)
    format: str | None = Field.attribute(expected_type=str, allow_none=True)

    def __init__(self,
                 culture=None,
                 format=None,
                ):
        self.culture = culture
        self.format = format


class Query(TypedSerialisable):

    tagname = "query"

    mdx: str | None = Field.attribute(expected_type=str, allow_none=False)
    tpls: TupleList | None = Field.element(expected_type=TupleList, allow_none=True)

    xml_order = ("tpls",)

    def __init__(self,
                 mdx=None,
                 tpls=None,
                ):
        self.mdx = mdx
        self.tpls = tpls


class OLAPSet(TypedSerialisable):

    tagname = "set"

    count: int | None = Field.attribute(expected_type=int, allow_none=False)
    maxRank: int | None = Field.attribute(expected_type=int, allow_none=False)
    setDefinition: str | None = Field.attribute(expected_type=str, allow_none=False)
    sortType: str | None = Field.attribute(expected_type=str, allow_none=True)
    queryFailed: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    tpls: TupleList | None = Field.element(expected_type=TupleList, allow_none=True)
    sortByTuple: TupleList | None = Field.element(expected_type=TupleList, allow_none=True)

    xml_order = ("tpls", "sortByTuple")

    def __init__(self,
                 count=None,
                 maxRank=None,
                 setDefinition=None,
                 sortType=None,
                 queryFailed=None,
                 tpls=None,
                 sortByTuple=None,
                ):
        self.count = count
        self.maxRank = maxRank
        self.setDefinition = setDefinition
        self.sortType = sortType
        self.queryFailed = queryFailed
        self.tpls = tpls
        self.sortByTuple = sortByTuple


class PCDSDTCEntries(TypedSerialisable):
    # Implements CT_PCDSDTCEntries

    tagname = "entries"

    count: int | None = Field.attribute(expected_type=int, allow_none=True)
    # elements are choice
    m: Missing | None = Field.element(expected_type=Missing, allow_none=True)
    n: Number | None = Field.element(expected_type=Number, allow_none=True)
    e: Error | None = Field.element(expected_type=Error, allow_none=True)
    s: Text | None = Field.element(expected_type=Text, allow_none=True)

    xml_order = ("m", "n", "e", "s")

    def __init__(self,
                 count=None,
                 m=None,
                 n=None,
                 e=None,
                 s=None,
                ):
        self.count = count
        self.m = m
        self.n = n
        self.e = e
        self.s = s


class TupleCache(TypedSerialisable):

    tagname = "tupleCache"

    entries: PCDSDTCEntries | None = Field.element(expected_type=PCDSDTCEntries, allow_none=True)
    sets: list[OLAPSet] = Field.nested_sequence(expected_type=OLAPSet, count=True)
    queryCache: list[Query] = Field.nested_sequence(expected_type=Query, count=True)
    serverFormats: list[ServerFormat] = Field.nested_sequence(expected_type=ServerFormat, count=True)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True)

    xml_order = ("entries", "sets", "queryCache", "serverFormats", "extLst")

    def __init__(self,
                 entries=None,
                 sets=(),
                 queryCache=(),
                 serverFormats=(),
                 extLst=None,
                ):
        self.entries = entries
        self.sets = sets
        self.queryCache = queryCache
        self.serverFormats = serverFormats
        self.extLst = extLst


class OLAPKPI(TypedSerialisable):

    tagname = "kpi"

    uniqueName: str | None = Field.attribute(expected_type=str, allow_none=False)
    caption: str | None = Field.attribute(expected_type=str, allow_none=True)
    displayFolder: str | None = Field.attribute(expected_type=str, allow_none=True)
    measureGroup: str | None = Field.attribute(expected_type=str, allow_none=True)
    parent: str | None = Field.attribute(expected_type=str, allow_none=True)
    value: str | None = Field.attribute(expected_type=str, allow_none=False)
    goal: str | None = Field.attribute(expected_type=str, allow_none=True)
    status: str | None = Field.attribute(expected_type=str, allow_none=True)
    trend: str | None = Field.attribute(expected_type=str, allow_none=True)
    weight: str | None = Field.attribute(expected_type=str, allow_none=True)
    time: str | None = Field.attribute(expected_type=str, allow_none=True)

    def __init__(self,
                 uniqueName=None,
                 caption=None,
                 displayFolder=None,
                 measureGroup=None,
                 parent=None,
                 value=None,
                 goal=None,
                 status=None,
                 trend=None,
                 weight=None,
                 time=None,
                ):
        self.uniqueName = uniqueName
        self.caption = caption
        self.displayFolder = displayFolder
        self.measureGroup = measureGroup
        self.parent = parent
        self.value = value
        self.goal = goal
        self.status = status
        self.trend = trend
        self.weight = weight
        self.time = time


class GroupMember(TypedSerialisable):

    tagname = "groupMember"

    uniqueName: str | None = Field.attribute(expected_type=str, allow_none=False)
    group: bool | None = Field.attribute(expected_type=bool, allow_none=False)

    def __init__(self,
                 uniqueName=None,
                 group=False,
                ):
        self.uniqueName = uniqueName
        self.group = group


class LevelGroup(TypedSerialisable):

    tagname = "group"

    name: str | None = Field.attribute(expected_type=str, allow_none=True)
    uniqueName: str | None = Field.attribute(expected_type=str, allow_none=True)
    caption: str | None = Field.attribute(expected_type=str, allow_none=True)
    uniqueParent: str | None = Field.attribute(expected_type=str, allow_none=True)
    id: int | None = Field.attribute(expected_type=int, allow_none=True)
    groupMembers: list[GroupMember] | None = Field.nested_sequence(
        expected_type=GroupMember,
        allow_none=True,
        count=True,
    )

    xml_order = ("groupMembers",)

    def __init__(self,
                 name=None,
                 uniqueName=None,
                 caption=None,
                 uniqueParent=None,
                 id=None,
                 groupMembers=(),
                ):
        self.name = name
        self.uniqueName = uniqueName
        self.caption = caption
        self.uniqueParent = uniqueParent
        self.id = id
        self.groupMembers = groupMembers


class GroupLevel(TypedSerialisable):

    tagname = "groupLevel"

    uniqueName: str | None = Field.attribute(expected_type=str, allow_none=False)
    caption: str | None = Field.attribute(expected_type=str, allow_none=False)
    user: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    customRollUp: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    groups: list[LevelGroup] = Field.nested_sequence(expected_type=LevelGroup, count=True)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True)

    xml_order = ("groups", "extLst")

    def __init__(self,
                 uniqueName=None,
                 caption=None,
                 user=None,
                 customRollUp=None,
                 groups=(),
                 extLst=None,
                ):
        self.uniqueName = uniqueName
        self.caption = caption
        self.user = user
        self.customRollUp = customRollUp
        self.groups = groups
        self.extLst = extLst


class FieldUsage(TypedSerialisable):

    tagname = "fieldUsage"

    x: int | None = Field.attribute(expected_type=int, allow_none=False)

    def __init__(self,
                 x=None,
                ):
        self.x = x


class CacheHierarchy(TypedSerialisable):

    tagname = "cacheHierarchy"

    uniqueName: str | None = Field.attribute(expected_type=str, allow_none=False)
    caption: str | None = Field.attribute(expected_type=str, allow_none=True)
    measure: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    set: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    parentSet: int | None = Field.attribute(expected_type=int, allow_none=True)
    iconSet: int | None = Field.attribute(expected_type=int, allow_none=False)
    attribute: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    time: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    keyAttribute: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    defaultMemberUniqueName: str | None = Field.attribute(expected_type=str, allow_none=True)
    allUniqueName: str | None = Field.attribute(expected_type=str, allow_none=True)
    allCaption: str | None = Field.attribute(expected_type=str, allow_none=True)
    dimensionUniqueName: str | None = Field.attribute(expected_type=str, allow_none=True)
    displayFolder: str | None = Field.attribute(expected_type=str, allow_none=True)
    measureGroup: str | None = Field.attribute(expected_type=str, allow_none=True)
    measures: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    count: int | None = Field.attribute(expected_type=int, allow_none=False)
    oneField: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    memberValueDatatype: int | None = Field.attribute(expected_type=int, allow_none=True)
    unbalanced: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    unbalancedGroup: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    hidden: bool | None = Field.attribute(expected_type=bool, allow_none=False)
    fieldsUsage: list[FieldUsage] = Field.nested_sequence(expected_type=FieldUsage, count=True)
    groupLevels: list[GroupLevel] = Field.nested_sequence(expected_type=GroupLevel, count=True)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True)

    xml_order = ("fieldsUsage", "groupLevels")

    def __init__(self,
                 uniqueName="",
                 caption=None,
                 measure=False,
                 set=False,
                 parentSet=None,
                 iconSet=0,
                 attribute=None,
                 time=None,
                 keyAttribute=False,
                 defaultMemberUniqueName=None,
                 allUniqueName=None,
                 allCaption=None,
                 dimensionUniqueName=None,
                 displayFolder=None,
                 measureGroup=None,
                 measures=False,
                 count=None,
                 oneField=False,
                 memberValueDatatype=None,
                 unbalanced=None,
                 unbalancedGroup=None,
                 hidden=False,
                 fieldsUsage=(),
                 groupLevels=(),
                 extLst=None,
                ):
        self.uniqueName = uniqueName
        self.caption = caption
        self.measure = measure
        self.set = set
        self.parentSet = parentSet
        self.iconSet = iconSet
        self.attribute = attribute
        self.time = time
        self.keyAttribute = keyAttribute
        self.defaultMemberUniqueName = defaultMemberUniqueName
        self.allUniqueName = allUniqueName
        self.allCaption = allCaption
        self.dimensionUniqueName = dimensionUniqueName
        self.displayFolder = displayFolder
        self.measureGroup = measureGroup
        self.measures = measures
        self.count = count
        self.oneField = oneField
        self.memberValueDatatype = memberValueDatatype
        self.unbalanced = unbalanced
        self.unbalancedGroup = unbalancedGroup
        self.hidden = hidden
        self.fieldsUsage = fieldsUsage
        self.groupLevels = groupLevels
        self.extLst = extLst


class GroupItems(TypedSerialisable):

    tagname = "groupItems"

    m: list[Missing] = Field.sequence(expected_type=Missing)
    n: list[Number] = Field.sequence(expected_type=Number)
    b: list[Boolean] = Field.sequence(expected_type=Boolean)
    e: list[Error] = Field.sequence(expected_type=Error)
    s: list[Text] = Field.sequence(expected_type=Text)
    d: list[DateTimeField] = Field.sequence(expected_type=DateTimeField)

    xml_order = ("m", "n", "b", "e", "s", "d")

    def __init__(self,
                 count=None,
                 m=(),
                 n=(),
                 b=(),
                 e=(),
                 s=(),
                 d=(),
                ):
        self.m = m
        self.n = n
        self.b = b
        self.e = e
        self.s = s
        self.d = d


    @property
    def count(self):
        return len(self.m + self.n + self.b + self.e + self.s + self.d)

    def __iter__(self):
        yield from super().__iter__()
        yield "count", str(self.count)


class RangePr(TypedSerialisable):

    tagname = "rangePr"

    autoStart: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    autoEnd: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    groupBy: str | None = Field.attribute(expected_type=str, allow_none=False)
    startNum: float | None = Field.attribute(expected_type=float, allow_none=True)
    endNum: float | None = Field.attribute(expected_type=float, allow_none=True)
    startDate: datetime | None = Field.attribute(expected_type=datetime, allow_none=True, converter=_coerce_iso_datetime)
    endDate: datetime | None = Field.attribute(expected_type=datetime, allow_none=True, converter=_coerce_iso_datetime)
    groupInterval: float | None = Field.attribute(expected_type=float, allow_none=True)

    def __init__(self,
                 autoStart=True,
                 autoEnd=True,
                 groupBy="range",
                 startNum=None,
                 endNum=None,
                 startDate=None,
                 endDate=None,
                 groupInterval=1,
                ):
        self.autoStart = autoStart
        self.autoEnd = autoEnd
        self.groupBy = groupBy
        self.startNum = startNum
        self.endNum = endNum
        self.startDate = startDate
        self.endDate = endDate
        self.groupInterval = groupInterval


class FieldGroup(TypedSerialisable):

    tagname = "fieldGroup"

    par: int | None = Field.attribute(expected_type=int, allow_none=True)
    base: int | None = Field.attribute(expected_type=int, allow_none=True)
    rangePr: RangePr | None = Field.element(expected_type=RangePr, allow_none=True)
    discretePr: list[NestedInteger] = Field.nested_sequence(expected_type=NestedInteger, count=True)
    groupItems: GroupItems | None = Field.element(expected_type=GroupItems, allow_none=True)

    xml_order = ("rangePr", "discretePr", "groupItems")

    def __init__(self,
                 par=None,
                 base=None,
                 rangePr=None,
                 discretePr=(),
                 groupItems=None,
                ):
        self.par = par
        self.base = base
        self.rangePr = rangePr
        self.discretePr = discretePr
        self.groupItems = groupItems


class SharedItems(TypedSerialisable):

    tagname = "sharedItems"

    _fields: list[Missing | Number | Boolean | Error | Text | DateTimeField] = Field.multi_sequence(
        parts={
            "m": Missing,
            "n": Number,
            "b": Boolean,
            "e": Error,
            "s": Text,
            "d": DateTimeField,
        }
    )
    # attributes are optional and must be derived from associated cache records
    containsSemiMixedTypes: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    containsNonDate: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    containsDate: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    containsString: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    containsBlank: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    containsMixedTypes: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    containsNumber: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    containsInteger: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    minValue: float | None = Field.attribute(expected_type=float, allow_none=True)
    maxValue: float | None = Field.attribute(expected_type=float, allow_none=True)
    minDate: datetime | None = Field.attribute(expected_type=datetime, allow_none=True)
    maxDate: datetime | None = Field.attribute(expected_type=datetime, allow_none=True)
    longText: bool | None = Field.attribute(expected_type=bool, allow_none=True)

    xml_order = ()

    def __init__(self,
                 _fields=(),
                 containsSemiMixedTypes=None,
                 containsNonDate=None,
                 containsDate=None,
                 containsString=None,
                 containsBlank=None,
                 containsMixedTypes=None,
                 containsNumber=None,
                 containsInteger=None,
                 minValue=None,
                 maxValue=None,
                 minDate=None,
                 maxDate=None,
                 count=None,
                 longText=None,
                ):
        self._fields = _fields
        self.containsBlank = containsBlank
        self.containsDate = containsDate
        self.containsNonDate = containsNonDate
        self.containsString = containsString
        self.containsMixedTypes = containsMixedTypes
        self.containsSemiMixedTypes = containsSemiMixedTypes
        self.containsNumber = containsNumber
        self.containsInteger = containsInteger
        self.minValue = minValue
        self.maxValue = maxValue
        self.minDate = minDate
        self.maxDate = maxDate
        self.longText = longText


    @property
    def count(self):
        return len(self._fields)

    def __iter__(self):
        yield from super().__iter__()
        yield "count", str(self.count)


class CacheField(TypedSerialisable):

    tagname = "cacheField"

    sharedItems: SharedItems | None = Field.element(expected_type=SharedItems, allow_none=True)
    fieldGroup: FieldGroup | None = Field.element(expected_type=FieldGroup, allow_none=True)
    mpMap: NestedInteger | None = Field.element(expected_type=NestedInteger, allow_none=True)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True)
    name: str | None = Field.attribute(expected_type=str, allow_none=False)
    caption: str | None = Field.attribute(expected_type=str, allow_none=True)
    propertyName: str | None = Field.attribute(expected_type=str, allow_none=True)
    serverField: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    uniqueList: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    numFmtId: int | None = Field.attribute(expected_type=int, allow_none=True)
    formula: str | None = Field.attribute(expected_type=str, allow_none=True)
    sqlType: int | None = Field.attribute(expected_type=int, allow_none=True)
    hierarchy: int | None = Field.attribute(expected_type=int, allow_none=True)
    level: int | None = Field.attribute(expected_type=int, allow_none=True)
    databaseField: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    mappingCount: int | None = Field.attribute(expected_type=int, allow_none=True)
    memberPropertyField: bool | None = Field.attribute(expected_type=bool, allow_none=True)

    xml_order = ("sharedItems", "fieldGroup", "mpMap")

    def __init__(self,
                 sharedItems=None,
                 fieldGroup=None,
                 mpMap=None,
                 extLst=None,
                 name=None,
                 caption=None,
                 propertyName=None,
                 serverField=None,
                 uniqueList=True,
                 numFmtId=None,
                 formula=None,
                 sqlType=0,
                 hierarchy=0,
                 level=0,
                 databaseField=True,
                 mappingCount=None,
                 memberPropertyField=None,
                ):
        self.sharedItems = sharedItems
        self.fieldGroup = fieldGroup
        self.mpMap = mpMap
        self.extLst = extLst
        self.name = name
        self.caption = caption
        self.propertyName = propertyName
        self.serverField = serverField
        self.uniqueList = uniqueList
        self.numFmtId = numFmtId
        self.formula = formula
        self.sqlType = sqlType
        self.hierarchy = hierarchy
        self.level = level
        self.databaseField = databaseField
        self.mappingCount = mappingCount
        self.memberPropertyField = memberPropertyField


class RangeSet(TypedSerialisable):

    tagname = "rangeSet"

    i1: int | None = Field.attribute(expected_type=int, allow_none=True)
    i2: int | None = Field.attribute(expected_type=int, allow_none=True)
    i3: int | None = Field.attribute(expected_type=int, allow_none=True)
    i4: int | None = Field.attribute(expected_type=int, allow_none=True)
    ref: str | None = Field.attribute(expected_type=str, allow_none=False)
    name: str | None = Field.attribute(expected_type=str, allow_none=True)
    sheet: str | None = Field.attribute(expected_type=str, allow_none=True)

    def __init__(self,
                 i1=None,
                 i2=None,
                 i3=None,
                 i4=None,
                 ref=None,
                 name=None,
                 sheet=None,
                ):
        self.i1 = i1
        self.i2 = i2
        self.i3 = i3
        self.i4 = i4
        self.ref = ref
        self.name = name
        self.sheet = sheet


class PageItem(TypedSerialisable):

    tagname = "pageItem"

    name: str | None = Field.attribute(expected_type=str, allow_none=False)

    def __init__(self,
                 name=None,
                ):
        self.name = name


class Consolidation(TypedSerialisable):

    tagname = "consolidation"

    autoPage: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    pages: list[PageItem] = Field.nested_sequence(expected_type=PageItem, count=True)
    rangeSets: list[RangeSet] = Field.nested_sequence(expected_type=RangeSet, count=True)

    xml_order = ("pages", "rangeSets")

    def __init__(self,
                 autoPage=None,
                 pages=(),
                 rangeSets=(),
                ):
        self.autoPage = autoPage
        self.pages = pages
        self.rangeSets = rangeSets


class WorksheetSource(TypedSerialisable):

    tagname = "worksheetSource"

    ref: str | None = Field.attribute(expected_type=str, allow_none=True)
    name: str | None = Field.attribute(expected_type=str, allow_none=True)
    sheet: str | None = Field.attribute(expected_type=str, allow_none=True)

    def __init__(self,
                 ref=None,
                 name=None,
                 sheet=None,
                ):
        self.ref = ref
        self.name = name
        self.sheet = sheet


class CacheSource(TypedSerialisable):

    tagname = "cacheSource"

    type: str | None = Field.attribute(expected_type=str, allow_none=False)
    connectionId: int | None = Field.attribute(expected_type=int, allow_none=True)
    # some elements are choice
    worksheetSource: WorksheetSource | None = Field.element(expected_type=WorksheetSource, allow_none=True)
    consolidation: Consolidation | None = Field.element(expected_type=Consolidation, allow_none=True)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True)

    xml_order = ("worksheetSource", "consolidation")

    def __init__(self,
                 type=None,
                 connectionId=None,
                 worksheetSource=None,
                 consolidation=None,
                 extLst=None,
                ):
        self.type = type
        self.connectionId = connectionId
        self.worksheetSource = worksheetSource
        self.consolidation = consolidation
        self.extLst = extLst


class CacheDefinition(TypedSerialisable):

    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"
    rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"
    _id = 1
    _path = "/xl/pivotCache/pivotCacheDefinition{0}.xml"
    records = None

    tagname = "pivotCacheDefinition"

    invalid: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    saveData: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    refreshOnLoad: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    optimizeMemory: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    enableRefresh: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    refreshedBy: str | None = Field.attribute(expected_type=str, allow_none=True)
    refreshedDate: float | None = Field.attribute(expected_type=float, allow_none=True)
    refreshedDateIso: datetime | None = Field.attribute(expected_type=datetime, allow_none=True)
    backgroundQuery: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    missingItemsLimit: int | None = Field.attribute(expected_type=int, allow_none=True)
    createdVersion: int | None = Field.attribute(expected_type=int, allow_none=True)
    refreshedVersion: int | None = Field.attribute(expected_type=int, allow_none=True)
    minRefreshableVersion: int | None = Field.attribute(expected_type=int, allow_none=True)
    recordCount: int | None = Field.attribute(expected_type=int, allow_none=True)
    upgradeOnRefresh: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    supportSubquery: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    supportAdvancedDrill: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    cacheSource: CacheSource | None = Field.element(expected_type=CacheSource, allow_none=False)
    cacheFields: list[CacheField] = Field.nested_sequence(expected_type=CacheField, count=True)
    cacheHierarchies: list[CacheHierarchy] = Field.nested_sequence(expected_type=CacheHierarchy, allow_none=True)
    kpis: list[OLAPKPI] = Field.nested_sequence(expected_type=OLAPKPI, count=True)
    tupleCache: TupleCache | None = Field.element(expected_type=TupleCache, allow_none=True)
    calculatedItems: list[CalculatedItem] = Field.nested_sequence(expected_type=CalculatedItem, count=True)
    calculatedMembers: list[CalculatedMember] = Field.nested_sequence(expected_type=CalculatedMember, count=True)
    dimensions: list[PivotDimension] = Field.nested_sequence(expected_type=PivotDimension, allow_none=True)
    measureGroups: list[MeasureGroup] = Field.nested_sequence(expected_type=MeasureGroup, count=True)
    maps: list[MeasureDimensionMap] = Field.nested_sequence(expected_type=MeasureDimensionMap, count=True)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True)
    id: str | None = Field.attribute(expected_type=str, allow_none=True, namespace=REL_NS)

    xml_order = (
        "cacheSource",
        "cacheFields",
        "cacheHierarchies",
        "kpis",
        "tupleCache",
        "calculatedItems",
        "calculatedMembers",
        "dimensions",
        "measureGroups",
        "maps",
    )

    def __init__(self,
                 invalid=None,
                 saveData=None,
                 refreshOnLoad=None,
                 optimizeMemory=None,
                 enableRefresh=None,
                 refreshedBy=None,
                 refreshedDate=None,
                 refreshedDateIso=None,
                 backgroundQuery=None,
                 missingItemsLimit=None,
                 createdVersion=None,
                 refreshedVersion=None,
                 minRefreshableVersion=None,
                 recordCount=None,
                 upgradeOnRefresh=None,
                 tupleCache=None,
                 supportSubquery=None,
                 supportAdvancedDrill=None,
                 cacheSource=None,
                 cacheFields=(),
                 cacheHierarchies=(),
                 kpis=(),
                 calculatedItems=(),
                 calculatedMembers=(),
                 dimensions=(),
                 measureGroups=(),
                 maps=(),
                 extLst=None,
                 id = None,
                ):
        self.invalid = invalid
        self.saveData = saveData
        self.refreshOnLoad = refreshOnLoad
        self.optimizeMemory = optimizeMemory
        self.enableRefresh = enableRefresh
        self.refreshedBy = refreshedBy
        self.refreshedDate = refreshedDate
        self.refreshedDateIso = refreshedDateIso
        self.backgroundQuery = backgroundQuery
        self.missingItemsLimit = missingItemsLimit
        self.createdVersion = createdVersion
        self.refreshedVersion = refreshedVersion
        self.minRefreshableVersion = minRefreshableVersion
        self.recordCount = recordCount
        self.upgradeOnRefresh = upgradeOnRefresh
        self.supportSubquery = supportSubquery
        self.supportAdvancedDrill = supportAdvancedDrill
        self.cacheSource = cacheSource
        self.cacheFields = cacheFields
        self.cacheHierarchies = cacheHierarchies
        self.kpis = kpis
        self.tupleCache = tupleCache
        self.calculatedItems = calculatedItems
        self.calculatedMembers = calculatedMembers
        self.dimensions = dimensions
        self.measureGroups = measureGroups
        self.maps = maps
        self.extLst = extLst
        self.id = id


    def to_tree(self):
        node = super().to_tree()
        node.set("xmlns", SHEET_MAIN_NS)
        return node


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
        if self.records is None:
            return

        rels = RelationshipList()
        r = Relationship(Type=self.records.rel_type, Target=self.records.path)
        rels.append(r)
        self.id = r.id
        self.records._id = self._id
        self.records._write(archive, manifest)

        path = get_rels_path(self.path)
        xml = tostring(rels.to_tree())
        archive.writestr(path[1:], xml)
