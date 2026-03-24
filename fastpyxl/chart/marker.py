# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.descriptors.excel import ExtensionList, _explicit_none

from .picture import PictureOptions
from .shapes import GraphicalProperties


def _coerce_marker_symbol(v):
    if v is None or v == "none":
        return None
    allowed = frozenset(
        {
            "circle",
            "dash",
            "diamond",
            "dot",
            "picture",
            "plus",
            "square",
            "star",
            "triangle",
            "x",
            "auto",
        }
    )
    if v not in allowed:
        raise FieldValidationError(f"symbol rejected value {v!r}")
    return v


def _marker_size(v):
    if v is None:
        return None
    try:
        n = int(v)
    except (TypeError, ValueError) as exc:
        raise FieldValidationError(f"size rejected value {v!r}") from exc
    if n < 2 or n > 72:
        raise FieldValidationError(f"size rejected value {v!r}")
    return n


class Marker(Serialisable):
    tagname = "marker"

    symbol: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_coerce_marker_symbol,
        renderer=_explicit_none,
    )
    size: int | None = Field.nested_value(
        expected_type=int,
        allow_none=True,
        converter=_marker_size,
    )
    spPr: GraphicalProperties | None = Field.element(
        expected_type=GraphicalProperties, allow_none=True
    )
    graphicalProperties = AliasField("spPr")
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = ("symbol", "size", "spPr")

    def __init__(
        self,
        symbol=None,
        size=None,
        spPr=None,
        extLst=None,
    ):
        self.symbol = symbol
        self.size = size
        if spPr is None:
            spPr = GraphicalProperties()
        self.spPr = spPr
        self.extLst = extLst


class DataPoint(Serialisable):
    tagname = "dPt"

    idx: int | None = Field.nested_value(expected_type=int, allow_none=True)
    invertIfNegative: bool | None = Field.nested_bool(allow_none=True)
    marker: Marker | None = Field.element(expected_type=Marker, allow_none=True)
    bubble3D: bool | None = Field.nested_bool(allow_none=True)
    explosion: int | None = Field.nested_value(expected_type=int, allow_none=True)
    spPr: GraphicalProperties | None = Field.element(
        expected_type=GraphicalProperties, allow_none=True
    )
    graphicalProperties = AliasField("spPr")
    pictureOptions: PictureOptions | None = Field.element(
        expected_type=PictureOptions, allow_none=True
    )
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = (
        "idx",
        "invertIfNegative",
        "marker",
        "bubble3D",
        "explosion",
        "spPr",
        "pictureOptions",
    )

    def __init__(
        self,
        idx=None,
        invertIfNegative=None,
        marker=None,
        bubble3D=None,
        explosion=None,
        spPr=None,
        pictureOptions=None,
        extLst=None,
    ):
        self.idx = idx
        self.invertIfNegative = invertIfNegative
        self.marker = marker
        self.bubble3D = bubble3D
        self.explosion = explosion
        if spPr is None:
            spPr = GraphicalProperties()
        self.spPr = spPr
        self.pictureOptions = pictureOptions
        self.extLst = extLst
