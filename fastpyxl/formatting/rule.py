# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.descriptors.excel import ExtensionList
from fastpyxl.styles.colors import Color
from fastpyxl.styles.differential import DifferentialStyle
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import Field

from fastpyxl.utils.cell import COORD_RE


def _enum_converter(value, allowed_values, field_name: str):
    if value is None:
        return None
    if value not in allowed_values:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return value


def _color_converter(value):
    if value is None:
        return None
    if isinstance(value, Color):
        return value
    if isinstance(value, str):
        return Color(rgb=value)
    raise FieldValidationError(f"color rejected value {value!r}")


class FormatObject(Serialisable):

    tagname = "cfvo"

    type: str | None = Field.attribute(
        expected_type=str,
        converter=lambda v: _enum_converter(v, ("num", "percent", "max", "min", "formula", "percentile"), "type"),
    )
    val: object | None = Field.attribute(expected_type=object, allow_none=True)
    gte: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True, serialize=False)

    def __init__(self, type, val=None, gte=None, extLst=None):
        self.type = type
        self.val = val
        self.gte = gte
        self.extLst = extLst

    def __setattr__(self, name, value):
        if name == "val" and value is not None:
            ref = isinstance(value, str) and COORD_RE.match(value)
            if getattr(self, "type", None) == "formula" or ref:
                value = str(value)
            else:
                try:
                    value = float(value)
                except Exception as exc:
                    raise TypeError(f"val rejected value {value!r}") from exc
        super().__setattr__(name, value)


class RuleType(Serialisable):

    cfvo: list[FormatObject] = Field.sequence(expected_type=FormatObject, default=list)


class IconSet(RuleType):

    tagname = "iconSet"

    iconSet: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, (
            '3Arrows', '3ArrowsGray', '3Flags',
            '3TrafficLights1', '3TrafficLights2', '3Signs', '3Symbols', '3Symbols2',
            '4Arrows', '4ArrowsGray', '4RedToBlack', '4Rating', '4TrafficLights',
            '5Arrows', '5ArrowsGray', '5Rating', '5Quarters'
        ), "iconSet"),
    )
    showValue: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    percent: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    reverse: bool | None = Field.attribute(expected_type=bool, allow_none=True)

    xml_order = ("cfvo",)

    def __init__(self, iconSet=None, showValue=None, percent=None, reverse=None, cfvo=None):
        self.iconSet = iconSet
        self.showValue = showValue
        self.percent = percent
        self.reverse = reverse
        self.cfvo = list(cfvo or ())


class DataBar(RuleType):

    tagname = "dataBar"

    minLength: int | None = Field.attribute(expected_type=int, allow_none=True)
    maxLength: int | None = Field.attribute(expected_type=int, allow_none=True)
    showValue: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    color: Color | None = Field.element(expected_type=Color, converter=_color_converter)

    xml_order = ('cfvo', 'color')

    def __init__(self, minLength=None, maxLength=None, showValue=None, cfvo=None, color=None):
        self.minLength = minLength
        self.maxLength = maxLength
        self.showValue = showValue
        self.cfvo = list(cfvo or ())
        self.color = color


class ColorScale(RuleType):

    tagname = "colorScale"

    color: list[Color] = Field.sequence(expected_type=Color, default=list)

    xml_order = ('cfvo', 'color')

    def __init__(self, cfvo=None, color=None):
        self.cfvo = list(cfvo or ())
        self.color = list(color or ())


class Rule(Serialisable):

    tagname = "cfRule"

    type: str | None = Field.attribute(
        expected_type=str,
        converter=lambda v: _enum_converter(v, (
            'expression', 'cellIs', 'colorScale', 'dataBar',
            'iconSet', 'top10', 'uniqueValues', 'duplicateValues', 'containsText',
            'notContainsText', 'beginsWith', 'endsWith', 'containsBlanks',
            'notContainsBlanks', 'containsErrors', 'notContainsErrors', 'timePeriod',
            'aboveAverage'
        ), 'type'),
    )
    dxfId: int | None = Field.attribute(expected_type=int, allow_none=True)
    priority: int | None = Field.attribute(expected_type=int, default=0)
    stopIfTrue: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    aboveAverage: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    percent: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    bottom: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    operator: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, (
            'lessThan', 'lessThanOrEqual', 'equal',
            'notEqual', 'greaterThanOrEqual', 'greaterThan', 'between', 'notBetween',
            'containsText', 'notContains', 'beginsWith', 'endsWith'
        ), 'operator'),
    )
    text: str | None = Field.attribute(expected_type=str, allow_none=True)
    timePeriod: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, (
            'today', 'yesterday', 'tomorrow', 'last7Days',
            'thisMonth', 'lastMonth', 'nextMonth', 'thisWeek', 'lastWeek', 'nextWeek'
        ), 'timePeriod'),
    )
    rank: int | None = Field.attribute(expected_type=int, allow_none=True)
    stdDev: int | None = Field.attribute(expected_type=int, allow_none=True)
    equalAverage: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    formula: list[str] = Field.sequence(expected_type=str, default=list)
    colorScale: ColorScale | None = Field.element(expected_type=ColorScale, allow_none=True)
    dataBar: DataBar | None = Field.element(expected_type=DataBar, allow_none=True)
    iconSet: IconSet | None = Field.element(expected_type=IconSet, allow_none=True)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True, serialize=False)
    dxf: DifferentialStyle | None = Field.element(expected_type=DifferentialStyle, allow_none=True, serialize=False)

    xml_order = ('colorScale', 'dataBar', 'iconSet', 'formula')

    def __init__(self,
                 type,
                 dxfId=None,
                 priority=0,
                 stopIfTrue=None,
                 aboveAverage=None,
                 percent=None,
                 bottom=None,
                 operator=None,
                 text=None,
                 timePeriod=None,
                 rank=None,
                 stdDev=None,
                 equalAverage=None,
                 formula=(),
                 colorScale=None,
                 dataBar=None,
                 iconSet=None,
                 extLst=None,
                 dxf=None,
                ):
        self.type = type
        self.dxfId = dxfId
        self.priority = priority
        self.stopIfTrue = stopIfTrue
        self.aboveAverage = aboveAverage
        self.percent = percent
        self.bottom = bottom
        self.operator = operator
        self.text = text
        self.timePeriod = timePeriod
        self.rank = rank
        self.stdDev = stdDev
        self.equalAverage = equalAverage
        self.formula = list(formula)
        self.colorScale = colorScale
        self.dataBar = dataBar
        self.iconSet = iconSet
        self.extLst = extLst
        self.dxf = dxf


def ColorScaleRule(start_type=None,
                 start_value=None,
                 start_color=None,
                 mid_type=None,
                 mid_value=None,
                 mid_color=None,
                 end_type=None,
                 end_value=None,
                 end_color=None):

    """Backwards compatibility"""
    formats = []
    if start_type is not None:
        formats.append(FormatObject(type=start_type, val=start_value))
    if mid_type is not None:
        formats.append(FormatObject(type=mid_type, val=mid_value))
    if end_type is not None:
        formats.append(FormatObject(type=end_type, val=end_value))
    colors = []
    for v in (start_color, mid_color, end_color):
        if v is not None:
            if not isinstance(v, Color):
                v = Color(v)
            colors.append(v)
    cs = ColorScale(cfvo=formats, color=colors)
    rule = Rule(type="colorScale", colorScale=cs)
    return rule


def FormulaRule(formula=None, stopIfTrue=None, font=None, border=None,
                fill=None):
    """
    Conditional formatting with custom differential style
    """
    rule = Rule(type="expression", formula=formula, stopIfTrue=stopIfTrue)
    rule.dxf =  DifferentialStyle(font=font, border=border, fill=fill)
    return rule


def CellIsRule(operator=None, formula=None, stopIfTrue=None, font=None, border=None, fill=None):
    """
    Conditional formatting rule based on cell contents.
    """
    # Excel doesn't use >, >=, etc, but allow for ease of python development
    expand = {">": "greaterThan", ">=": "greaterThanOrEqual", "<": "lessThan", "<=": "lessThanOrEqual",
              "=": "equal", "==": "equal", "!=": "notEqual"}

    if operator is not None:
        operator = expand.get(operator, operator)

    rule = Rule(type='cellIs', operator=operator, formula=formula, stopIfTrue=stopIfTrue)
    rule.dxf = DifferentialStyle(font=font, border=border, fill=fill)

    return rule


def IconSetRule(icon_style=None, type=None, values=None, showValue=None, percent=None, reverse=None):
    """
    Convenience function for creating icon set rules
    """
    cfvo = []
    for val in values or ():
        cfvo.append(FormatObject(type, val))
    icon_set = IconSet(iconSet=icon_style, cfvo=cfvo, showValue=showValue,
                       percent=percent, reverse=reverse)
    rule = Rule(type='iconSet', iconSet=icon_set)

    return rule


def DataBarRule(start_type=None, start_value=None, end_type=None,
                end_value=None, color=None, showValue=None, minLength=None, maxLength=None):
    start = FormatObject(start_type, start_value)
    end = FormatObject(end_type, end_value)
    data_bar = DataBar(cfvo=[start, end], color=color, showValue=showValue,
                       minLength=minLength, maxLength=maxLength)
    rule = Rule(type='dataBar', dataBar=data_bar)

    return rule
