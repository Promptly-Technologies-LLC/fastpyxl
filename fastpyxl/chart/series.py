# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.descriptors.excel import ExtensionList

from .shapes import GraphicalProperties
from .data_source import (
    AxDataSource,
    NumDataSource,
    StrRef,
)
from .error_bar import ErrorBars
from .label import DataLabelList
from .marker import DataPoint, PictureOptions, Marker
from .trendline import Trendline


def _shape_converter(v):
    if v is None or v == "none":
        return None
    allowed = frozenset(
        {"cone", "coneToMax", "box", "cylinder", "pyramid", "pyramidToMax"}
    )
    if v not in allowed:
        raise FieldValidationError(f"shape rejected value {v!r}")
    return v


attribute_mapping = {
    "area": (
        "idx",
        "order",
        "tx",
        "spPr",
        "pictureOptions",
        "dPt",
        "dLbls",
        "errBars",
        "trendline",
        "cat",
        "val",
    ),
    "bar": (
        "idx",
        "order",
        "tx",
        "spPr",
        "invertIfNegative",
        "pictureOptions",
        "dPt",
        "dLbls",
        "trendline",
        "errBars",
        "cat",
        "val",
        "shape",
    ),
    "bubble": (
        "idx",
        "order",
        "tx",
        "spPr",
        "invertIfNegative",
        "dPt",
        "dLbls",
        "trendline",
        "errBars",
        "xVal",
        "yVal",
        "bubbleSize",
        "bubble3D",
    ),
    "line": (
        "idx",
        "order",
        "tx",
        "spPr",
        "marker",
        "dPt",
        "dLbls",
        "trendline",
        "errBars",
        "cat",
        "val",
        "smooth",
    ),
    "pie": (
        "idx",
        "order",
        "tx",
        "spPr",
        "explosion",
        "dPt",
        "dLbls",
        "cat",
        "val",
    ),
    "radar": (
        "idx",
        "order",
        "tx",
        "spPr",
        "marker",
        "dPt",
        "dLbls",
        "cat",
        "val",
    ),
    "scatter": (
        "idx",
        "order",
        "tx",
        "spPr",
        "marker",
        "dPt",
        "dLbls",
        "trendline",
        "errBars",
        "xVal",
        "yVal",
        "smooth",
    ),
    "surface": ("idx", "order", "tx", "spPr", "cat", "val"),
}


class SeriesLabel(Serialisable):
    tagname = "tx"

    strRef: StrRef | None = Field.element(expected_type=StrRef, allow_none=True, default=None)
    v: str | None = Field.nested_text(expected_type=str, allow_none=True, default=None)
    value = AliasField("v", default=None)

    xml_order = ("strRef", "v")

    def __init__(self, strRef=None, v=None):
        self.strRef = strRef
        self.v = v


class Series(Serialisable):
    """
    Generic series object. Should not be instantiated directly.
    User the chart.Series factory instead.
    """

    tagname = "ser"

    idx: int | None = Field.nested_value(expected_type=int, allow_none=True, default=None)
    order: int | None = Field.nested_value(expected_type=int, allow_none=True, default=None)
    tx: SeriesLabel | None = Field.element(expected_type=SeriesLabel, allow_none=True, default=None)
    title = AliasField("tx", default=None)
    spPr: GraphicalProperties | None = Field.element(
        expected_type=GraphicalProperties, allow_none=True, default=None
    )
    graphicalProperties = AliasField("spPr", default=None)

    pictureOptions: PictureOptions | None = Field.element(
        expected_type=PictureOptions, allow_none=True, default=None
    )
    dPt: list[DataPoint] | None = Field.sequence(expected_type=DataPoint, allow_none=True, default=list)
    data_points = AliasField("dPt", default=None)
    dLbls: DataLabelList | None = Field.element(
        expected_type=DataLabelList, allow_none=True, default=None
    )
    labels = AliasField("dLbls", default=None)
    trendline: Trendline | None = Field.element(expected_type=Trendline, allow_none=True, default=None)
    errBars: ErrorBars | None = Field.element(expected_type=ErrorBars, allow_none=True, default=None)
    cat: AxDataSource | None = Field.element(expected_type=AxDataSource, allow_none=True, default=None)
    identifiers = AliasField("cat", default=None)
    val: NumDataSource | None = Field.element(expected_type=NumDataSource, allow_none=True, default=None)
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False, default=None
    )

    invertIfNegative: bool | None = Field.nested_bool(allow_none=True, default=None)
    shape: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=_shape_converter, default=None,
    )

    xVal: AxDataSource | None = Field.element(expected_type=AxDataSource, allow_none=True, default=None)
    yVal: NumDataSource | None = Field.element(expected_type=NumDataSource, allow_none=True, default=None)
    bubbleSize: NumDataSource | None = Field.element(
        expected_type=NumDataSource, allow_none=True, default=None
    )
    zVal = AliasField("bubbleSize", default=None)
    bubble3D: bool | None = Field.nested_bool(allow_none=True, default=None)

    marker: Marker | None = Field.element(expected_type=Marker, allow_none=True, default=None)
    smooth: bool | None = Field.nested_bool(allow_none=True, default=None)

    explosion: int | None = Field.nested_value(expected_type=int, allow_none=True, default=None)

    _serialize_element_order: tuple[str, ...] | None = None

    def __init__(
        self,
        idx=0,
        order=0,
        tx=None,
        spPr=None,
        pictureOptions=None,
        dPt=(),
        dLbls=None,
        trendline=None,
        errBars=None,
        cat=None,
        val=None,
        invertIfNegative=None,
        shape=None,
        xVal=None,
        yVal=None,
        bubbleSize=None,
        bubble3D=None,
        marker=None,
        smooth=None,
        explosion=None,
        extLst=None,
    ):
        self.idx = idx
        self.order = order
        self.tx = tx
        if spPr is None:
            spPr = GraphicalProperties()
        self.spPr = spPr
        self.pictureOptions = pictureOptions
        self.dPt = list(dPt) if dPt is not None else []
        self.dLbls = dLbls
        self.trendline = trendline
        self.errBars = errBars
        self.cat = cat
        self.val = val
        self.invertIfNegative = invertIfNegative
        self.shape = shape
        self.xVal = xVal
        self.yVal = yVal
        self.bubbleSize = bubbleSize
        self.bubble3D = bubble3D
        if marker is None:
            marker = Marker()
        self.marker = marker
        self.smooth = smooth
        self.explosion = explosion
        self.extLst = extLst

    def _element_names_for_serialize(self):
        order = getattr(self, "_serialize_element_order", None)
        if order is not None:
            return tuple(order)
        return super()._element_names_for_serialize()

    def to_tree(self, tagname=None, idx=None, namespace=None):
        del namespace
        if idx is not None:
            if self.order == self.idx:
                self.order = idx
            self.idx = idx
        return super().to_tree(tagname, idx)


class XYSeries(Series):
    pass
