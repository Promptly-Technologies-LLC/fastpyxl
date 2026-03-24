# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.descriptors.excel import ExtensionList

from .data_source import NumDataSource
from .shapes import GraphicalProperties


def _none_set(allowed: frozenset, field_name: str):
    def _c(v):
        if v is None or v == "none":
            return None
        if v not in allowed:
            raise FieldValidationError(f"{field_name} rejected value {v!r}")
        return v

    return _c


def _req_set(allowed: frozenset, field_name: str):
    def _c(v):
        if v is None:
            return None
        if v not in allowed:
            raise FieldValidationError(f"{field_name} rejected value {v!r}")
        return v

    return _c


class ErrorBars(Serialisable):
    tagname = "errBars"

    errDir: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_none_set(frozenset({"x", "y"}), "errDir"), default=None,
    )
    direction = AliasField("errDir", default=None)
    errBarType: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_req_set(frozenset({"both", "minus", "plus"}), "errBarType"), default=None,
    )
    style = AliasField("errBarType", default=None)
    errValType: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_req_set(
            frozenset({"cust", "fixedVal", "percentage", "stdDev", "stdErr"}),
            "errValType",
        ), default=None,
    )
    size = AliasField("errValType", default=None)
    noEndCap: bool | None = Field.nested_bool(allow_none=True, default=None)
    plus: NumDataSource | None = Field.element(expected_type=NumDataSource, allow_none=True, default=None)
    minus: NumDataSource | None = Field.element(expected_type=NumDataSource, allow_none=True, default=None)
    val: float | None = Field.nested_value(expected_type=float, allow_none=True, default=None)
    spPr: GraphicalProperties | None = Field.element(
        expected_type=GraphicalProperties, allow_none=True, default=None
    )
    graphicalProperties = AliasField("spPr", default=None)
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False, default=None
    )

    xml_order = (
        "errDir",
        "errBarType",
        "errValType",
        "noEndCap",
        "minus",
        "plus",
        "val",
        "spPr",
    )

    def __init__(
        self,
        errDir=None,
        errBarType="both",
        errValType="fixedVal",
        noEndCap=None,
        plus=None,
        minus=None,
        val=None,
        spPr=None,
        extLst=None,
    ):
        self.errDir = errDir
        self.errBarType = errBarType
        self.errValType = errValType
        self.noEndCap = noEndCap
        self.plus = plus
        self.minus = minus
        self.val = val
        self.spPr = spPr
        self.extLst = extLst
