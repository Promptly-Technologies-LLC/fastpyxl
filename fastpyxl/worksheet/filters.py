# Copyright (c) 2010-2024 fastpyxl

import re

from fastpyxl.descriptors.excel import ExtensionList
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field
from fastpyxl.utils import absolute_coordinate


def _enum(value, allowed, field_name):
    if value is None:
        return None
    if value not in allowed:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return value


class SortCondition(Serialisable):

    tagname = "sortCondition"

    descending: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    sortBy: str | None = Field.attribute(expected_type=str, allow_none=True, converter=lambda v: _enum(v, ('value', 'cellColor', 'fontColor', 'icon'), 'sortBy'))
    ref: str | None = Field.attribute(expected_type=str, converter=lambda v: v.upper() if isinstance(v, str) else v)
    customList: str | None = Field.attribute(expected_type=str, allow_none=True)
    dxfId: int | None = Field.attribute(expected_type=int, allow_none=True)
    iconSet: str | None = Field.attribute(expected_type=str, allow_none=True, converter=lambda v: _enum(v, (
        '3Arrows', '3ArrowsGray', '3Flags',
        '3TrafficLights1', '3TrafficLights2', '3Signs', '3Symbols', '3Symbols2',
        '4Arrows', '4ArrowsGray', '4RedToBlack', '4Rating', '4TrafficLights',
        '5Arrows', '5ArrowsGray', '5Rating', '5Quarters'
    ), 'iconSet'))
    iconId: int | None = Field.attribute(expected_type=int, allow_none=True)

    def __init__(self, ref=None, descending=None, sortBy=None, customList=None, dxfId=None, iconSet=None, iconId=None):
        self.descending = descending
        self.sortBy = sortBy
        self.ref = ref
        self.customList = customList
        self.dxfId = dxfId
        self.iconSet = iconSet
        self.iconId = iconId


class SortState(Serialisable):

    tagname = "sortState"

    columnSort: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    caseSensitive: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    sortMethod: str | None = Field.attribute(expected_type=str, allow_none=True, converter=lambda v: _enum(v, ('stroke', 'pinYin'), 'sortMethod'))
    ref: str | None = Field.attribute(expected_type=str, converter=lambda v: v.upper() if isinstance(v, str) else v)
    sortCondition: list[SortCondition] = Field.sequence(expected_type=SortCondition, allow_none=True, default=list)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True, serialize=False)

    xml_order = ('sortCondition',)

    def __init__(self, columnSort=None, caseSensitive=None, sortMethod=None, ref=None, sortCondition=(), extLst=None):
        self.columnSort = columnSort
        self.caseSensitive = caseSensitive
        self.sortMethod = sortMethod
        self.ref = ref
        self.sortCondition = list(sortCondition)
        self.extLst = extLst

    def __bool__(self):
        return self.ref is not None


class IconFilter(Serialisable):

    tagname = "iconFilter"

    iconSet: str | None = Field.attribute(expected_type=str, converter=lambda v: _enum(v, (
        '3Arrows', '3ArrowsGray', '3Flags',
        '3TrafficLights1', '3TrafficLights2', '3Signs', '3Symbols', '3Symbols2',
        '4Arrows', '4ArrowsGray', '4RedToBlack', '4Rating', '4TrafficLights',
        '5Arrows', '5ArrowsGray', '5Rating', '5Quarters'
    ), 'iconSet'))
    iconId: int | None = Field.attribute(expected_type=int, allow_none=True)

    def __init__(self, iconSet=None, iconId=None):
        self.iconSet = iconSet
        self.iconId = iconId


class ColorFilter(Serialisable):

    tagname = "colorFilter"

    dxfId: int | None = Field.attribute(expected_type=int, allow_none=True)
    cellColor: bool | None = Field.attribute(expected_type=bool, allow_none=True)

    def __init__(self, dxfId=None, cellColor=None):
        self.dxfId = dxfId
        self.cellColor = cellColor


class DynamicFilter(Serialisable):

    tagname = "dynamicFilter"

    type: str | None = Field.attribute(expected_type=str, converter=lambda v: _enum(v, (
        'null', 'aboveAverage', 'belowAverage', 'tomorrow',
        'today', 'yesterday', 'nextWeek', 'thisWeek', 'lastWeek', 'nextMonth',
        'thisMonth', 'lastMonth', 'nextQuarter', 'thisQuarter', 'lastQuarter',
        'nextYear', 'thisYear', 'lastYear', 'yearToDate', 'Q1', 'Q2', 'Q3', 'Q4',
        'M1', 'M2', 'M3', 'M4', 'M5', 'M6', 'M7', 'M8', 'M9', 'M10', 'M11', 'M12'
    ), 'type'))
    val: float | None = Field.attribute(expected_type=float, allow_none=True)
    valIso: str | None = Field.attribute(expected_type=str, allow_none=True)
    maxVal: float | None = Field.attribute(expected_type=float, allow_none=True)
    maxValIso: str | None = Field.attribute(expected_type=str, allow_none=True)

    def __init__(self, type=None, val=None, valIso=None, maxVal=None, maxValIso=None):
        self.type = type
        self.val = val
        self.valIso = valIso
        self.maxVal = maxVal
        self.maxValIso = maxValIso


class CustomFilter(Serialisable):

    tagname = "customFilter"

    val: str | None = Field.attribute(expected_type=str)
    operator: str | None = Field.attribute(expected_type=str, converter=lambda v: _enum(v, ('equal', 'lessThan', 'lessThanOrEqual', 'notEqual', 'greaterThanOrEqual', 'greaterThan'), 'operator'))

    def __init__(self, operator="equal", val=None):
        self.operator = operator
        self.val = val

    def _get_subtype(self):
        if self.val == " ":
            subtype = BlankFilter
        else:
            try:
                v = self.val
                if v is None:
                    raise ValueError("val required")
                float(v)
                subtype = NumberFilter
            except ValueError:
                subtype = StringFilter
        return subtype

    def convert(self):
        """Convert to more specific filter"""
        typ = self._get_subtype()
        if typ in (BlankFilter, NumberFilter):
            return typ(**dict(self))

        operator, term = StringFilter._guess_operator(self.val)
        flt = StringFilter(operator, term)
        if self.operator == "notEqual":
            flt.exclude = True
        return flt


class BlankFilter(CustomFilter):
    """
    Exclude blanks
    """

    def __init__(self, **kw):
        pass

    @property
    def operator(self):
        return "notEqual"

    @property
    def val(self):
        return " "


class NumberFilter(CustomFilter):

    operator: str | None = Field.attribute(expected_type=str, converter=lambda v: _enum(v, ('equal', 'lessThan', 'lessThanOrEqual', 'notEqual', 'greaterThanOrEqual', 'greaterThan'), 'operator'))
    val: float | None = Field.attribute(expected_type=float)

    def __init__(self, operator="equal", val=None):
        self.operator = operator
        self.val = val


string_format_mapping = {
    "contains": "*{}*",
    "startswith": "{}*",
    "endswith": "*{}",
    "wildcard":  "{}",
}


class StringFilter(CustomFilter):

    operator: str | None = Field.attribute(expected_type=str, converter=lambda v: _enum(v, ('contains', 'startswith', 'endswith', 'wildcard'), 'operator'))
    val: str | None = Field.attribute(expected_type=str)
    exclude: bool | None = Field.attribute(expected_type=bool, default=False)

    def __init__(self, operator="contains", val=None, exclude=False):
        self.operator = operator
        self.val = val
        self.exclude = exclude

    def _escape(self):
        """Escape wildcards ~, * ? when serialising"""
        v = self.val
        if v is None:
            return ""
        if self.operator == "wildcard":
            return v
        return re.sub(r"~|\*|\?", r"~\g<0>", v)

    @staticmethod
    def _unescape(value):
        """
        Unescape value
        """
        return re.sub(r"~(?P<op>[~*?])", r"\g<op>", value)

    @staticmethod
    def _guess_operator(value):
        value = StringFilter._unescape(value)
        endswith = r"^(?P<endswith>\*)(?P<term>[^\*\?]*$)"
        startswith = r"^(?P<term>[^\*\?]*)(?P<startswith>\*)$"
        contains = r"^(?P<contains>\*)(?P<term>[^\*\?]*)\*$"
        d = {"wildcard": True, "term": value}
        for pat in [contains, startswith, endswith]:
            m = re.match(pat, value)
            if m:
                d = m.groupdict()

        term = d.pop("term")
        op = list(d)[0]
        return op, term

    def to_tree(self, tagname=None, idx=None, namespace=None):
        string_op = self.operator
        assert string_op is not None
        fmt = string_format_mapping[string_op]
        op = self.exclude and "notEqual" or "equal"
        value = fmt.format(self._escape())
        flt = CustomFilter(op, value)
        return flt.to_tree(tagname, idx, namespace)


class CustomFilters(Serialisable):

    tagname = "customFilters"

    _and: bool | None = Field.attribute(expected_type=bool, allow_none=True, xml_name='and')
    customFilter: list[CustomFilter] = Field.sequence(expected_type=CustomFilter, default=list) # min 1, max 2

    xml_order = ('customFilter',)

    def __init__(self, _and=None, customFilter=()):
        self._and = _and
        self.customFilter = list(customFilter)


class Top10(Serialisable):

    tagname = "top10"

    top: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    percent: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    val: float | None = Field.attribute(expected_type=float)
    filterVal: float | None = Field.attribute(expected_type=float, allow_none=True)

    def __init__(self, top=None, percent=None, val=None, filterVal=None):
        self.top = top
        self.percent = percent
        self.val = val
        self.filterVal = filterVal


class DateGroupItem(Serialisable):

    tagname = "dateGroupItem"

    year: int | None = Field.attribute(expected_type=int)
    month: int | None = Field.attribute(expected_type=int, allow_none=True)
    day: int | None = Field.attribute(expected_type=int, allow_none=True)
    hour: int | None = Field.attribute(expected_type=int, allow_none=True)
    minute: int | None = Field.attribute(expected_type=int, allow_none=True)
    second: int | None = Field.attribute(expected_type=int, allow_none=True)
    dateTimeGrouping: str | None = Field.attribute(expected_type=str, converter=lambda v: _enum(v, ('year', 'month', 'day', 'hour', 'minute', 'second'), 'dateTimeGrouping'))

    def __init__(self, year=None, month=None, day=None, hour=None, minute=None, second=None, dateTimeGrouping=None):
        self.year = year
        self.month = month
        self.day = day
        self.hour = hour
        self.minute = minute
        self.second = second
        self.dateTimeGrouping = dateTimeGrouping


class Filters(Serialisable):

    tagname = "filters"

    blank: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    calendarType: str | None = Field.attribute(expected_type=str, allow_none=True, converter=lambda v: _enum(v, (
        "gregorian","gregorianUs","gregorianMeFrench","gregorianArabic", "hijri","hebrew",
        "taiwan","japan", "thai","korea", "saka","gregorianXlitEnglish","gregorianXlitFrench"
    ), 'calendarType'))
    filter: list[str] = Field.sequence(expected_type=str, primitive_attribute='val', default=list)
    dateGroupItem: list[DateGroupItem] = Field.sequence(expected_type=DateGroupItem, allow_none=True, default=list)

    xml_order = ('filter', 'dateGroupItem')

    def __init__(self, blank=None, calendarType=None, filter=(), dateGroupItem=()):
        self.blank = blank
        self.calendarType = calendarType
        self.filter = list(filter)
        self.dateGroupItem = list(dateGroupItem)


class FilterColumn(Serialisable):

    tagname = "filterColumn"

    colId: int | None = Field.attribute(expected_type=int)
    col_id = AliasField('colId')
    hiddenButton: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    showButton: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    # some elements are choice
    filters: Filters | None = Field.element(expected_type=Filters, allow_none=True)
    top10: Top10 | None = Field.element(expected_type=Top10, allow_none=True)
    customFilters: CustomFilters | None = Field.element(expected_type=CustomFilters, allow_none=True)
    dynamicFilter: DynamicFilter | None = Field.element(expected_type=DynamicFilter, allow_none=True)
    colorFilter: ColorFilter | None = Field.element(expected_type=ColorFilter, allow_none=True)
    iconFilter: IconFilter | None = Field.element(expected_type=IconFilter, allow_none=True)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True, serialize=False)

    xml_order = ('filters', 'top10', 'customFilters', 'dynamicFilter',
                 'colorFilter', 'iconFilter')

    def __init__(self,
                 colId=None,
                 hiddenButton=False,
                 showButton=True,
                 filters=None,
                 top10=None,
                 customFilters=None,
                 dynamicFilter=None,
                 colorFilter=None,
                 iconFilter=None,
                 extLst=None,
                 blank=None,
                 vals=None,
                ):
        self.colId = colId
        self.hiddenButton = hiddenButton
        self.showButton = showButton
        self.filters = filters
        self.top10 = top10
        self.customFilters = customFilters
        self.dynamicFilter = dynamicFilter
        self.colorFilter = colorFilter
        self.iconFilter = iconFilter
        self.extLst = extLst
        if blank is not None and self.filters:
            self.filters.blank = blank
        if vals is not None and self.filters:
            self.filters.filter = vals


class AutoFilter(Serialisable):

    tagname = "autoFilter"

    ref: str | None = Field.attribute(expected_type=str, converter=lambda v: v.upper() if isinstance(v, str) else v)
    filterColumn: list[FilterColumn] = Field.sequence(expected_type=FilterColumn, allow_none=True, default=list)
    sortState: SortState | None = Field.element(expected_type=SortState, allow_none=True)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True, serialize=False)

    xml_order = ('filterColumn', 'sortState')

    def __init__(self, ref=None, filterColumn=(), sortState=None, extLst=None):
        self.ref = ref
        self.filterColumn = list(filterColumn)
        self.sortState = sortState
        self.extLst = extLst

    def __bool__(self):
        return self.ref is not None

    def __str__(self):
        return absolute_coordinate(self.ref)

    def add_filter_column(self, col_id, vals, blank=False):
        """
        Add row filter for specified column.

        :param col_id: Zero-origin column id. 0 means first column.
        :type  col_id: int
        :param vals: Value list to show.
        :type  vals: str[]
        :param blank: Show rows that have blank cell if True (default=``False``)
        :type  blank: bool
        """
        self.filterColumn.append(FilterColumn(colId=col_id, filters=Filters(blank=blank, filter=vals)))

    def add_sort_condition(self, ref, descending=False):
        """
        Add sort condition for cpecified range of cells.

        :param ref: range of the cells (e.g. 'A2:A150')
        :type  ref: string, is the same as that of the filter
        :param descending: Descending sort order (default=``False``)
        :type  descending: bool
        """
        cond = SortCondition(ref, descending)
        if self.sortState is None:
            self.sortState = SortState(ref=self.ref)
        self.sortState.sortCondition.append(cond)
