# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.descriptors.excel import ExtensionList


def _none_set_converter(allowed: frozenset, field_name: str):
    def _c(value):
        if value is None:
            return None
        if value == "none":
            return None
        if value not in allowed:
            raise FieldValidationError(f"{field_name} rejected value {value!r}")
        return value

    return _c


def _range_converter(value, *, field_name: str, min_v: float, max_v: float):
    if value is None:
        return None
    try:
        numeric = float(value)
    except (TypeError, ValueError) as exc:
        raise FieldValidationError(f"{field_name} rejected value {value!r}") from exc
    if numeric < min_v or numeric > max_v:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return numeric


class ManualLayout(Serialisable):
    tagname = "manualLayout"

    layoutTarget: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_none_set_converter(frozenset({"inner", "outer"}), "layoutTarget"), default=None,
    )
    xMode: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_none_set_converter(frozenset({"edge", "factor"}), "xMode"), default=None,
    )
    yMode: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_none_set_converter(frozenset({"edge", "factor"}), "yMode"), default=None,
    )
    wMode: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_none_set_converter(frozenset({"edge", "factor"}), "wMode"), default=None,
    )
    hMode: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_none_set_converter(frozenset({"edge", "factor"}), "hMode"), default=None,
    )
    x: float | None = Field.nested_value(
        expected_type=float,
        allow_none=True,
        converter=lambda v: _range_converter(v, field_name="x", min_v=-1, max_v=1), default=None,
    )
    y: float | None = Field.nested_value(
        expected_type=float,
        allow_none=True,
        converter=lambda v: _range_converter(v, field_name="y", min_v=-1, max_v=1), default=None,
    )
    w: float | None = Field.nested_value(
        expected_type=float,
        allow_none=True,
        converter=lambda v: _range_converter(v, field_name="w", min_v=0, max_v=1), default=None,
    )
    width = AliasField("w", default=None)
    h: float | None = Field.nested_value(
        expected_type=float,
        allow_none=True,
        converter=lambda v: _range_converter(v, field_name="h", min_v=0, max_v=1), default=None,
    )
    height = AliasField("h", default=None)
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False, default=None
    )

    xml_order = ("layoutTarget", "xMode", "yMode", "wMode", "hMode", "x", "y", "w", "h")

    def __init__(
        self,
        layoutTarget=None,
        xMode=None,
        yMode=None,
        wMode="factor",
        hMode="factor",
        x=None,
        y=None,
        w=None,
        h=None,
        extLst=None,
    ):
        self.layoutTarget = layoutTarget
        self.xMode = xMode
        self.yMode = yMode
        self.wMode = wMode
        self.hMode = hMode
        self.x = x
        self.y = y
        self.w = w
        self.h = h
        self.extLst = extLst


class Layout(Serialisable):
    tagname = "layout"

    manualLayout: ManualLayout | None = Field.element(
        expected_type=ManualLayout, allow_none=True, default=None
    )
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False, default=None
    )

    xml_order = ("manualLayout",)

    def __init__(self, manualLayout=None, extLst=None):
        self.manualLayout = manualLayout
        self.extLst = extLst
