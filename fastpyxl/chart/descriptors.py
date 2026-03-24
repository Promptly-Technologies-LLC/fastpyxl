# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import Field

from .data_source import NumFmt


def _bounded_float(value, *, lo: float, hi: float, field_name: str):
    if value is None:
        return None
    try:
        numeric = float(value)
    except (TypeError, ValueError) as exc:
        raise FieldValidationError(f"{field_name} rejected value {value!r}") from exc
    if numeric < lo or numeric > hi:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return numeric


def _gap_amount(v):
    return _bounded_float(v, lo=0, hi=500, field_name="gapWidth")


def _overlap(v):
    return _bounded_float(v, lo=-100, hi=100, field_name="overlap")


NestedGapAmount = Field.nested_value(
    expected_type=float,
    allow_none=True,
    converter=_gap_amount,
)

NestedOverlap = Field.nested_value(
    expected_type=float,
    allow_none=True,
    converter=_overlap,
)


def num_fmt_from_value(value):
    if value is None:
        return None
    if isinstance(value, str):
        return NumFmt(formatCode=value)
    return value
