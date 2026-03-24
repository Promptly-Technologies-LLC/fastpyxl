# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.descriptors.excel import ExtensionList

from ._chart import ChartBase
from ._3d import (
    FIELD_BACK_WALL_ON_CHART,
    FIELD_FLOOR_ON_CHART,
    FIELD_SIDE_WALL_ON_CHART,
    FIELD_VIEW3D_ON_CHART,
    _3DBase,
)
from .axis import TextAxis, NumericAxis, SeriesAxis
from .shapes import GraphicalProperties
from .series import Series


class BandFormat(Serialisable):
    tagname = "bandFmt"

    idx: int | None = Field.nested_value(expected_type=int, allow_none=True)
    spPr: GraphicalProperties | None = Field.element(
        expected_type=GraphicalProperties, allow_none=True
    )
    graphicalProperties = AliasField("spPr")

    xml_order = ("idx", "spPr")

    def __init__(self, idx=0, spPr=None):
        self.idx = idx
        self.spPr = spPr


class BandFormatList(Serialisable):
    tagname = "bandFmts"

    bandFmt: list[BandFormat] | None = Field.sequence(
        expected_type=BandFormat, allow_none=True
    )

    xml_order = ("bandFmt",)

    def __init__(self, bandFmt=()):
        self.bandFmt = list(bandFmt) if bandFmt is not None else []


class _SurfaceChartBase(ChartBase):
    wireframe: bool | None = Field.nested_bool(allow_none=True)
    ser: list[Series] | None = Field.sequence(expected_type=Series, allow_none=True)
    bandFmts: BandFormatList | None = Field.element(
        expected_type=BandFormatList, allow_none=True
    )

    _series_type = "surface"

    xml_order = ("wireframe", "ser", "bandFmts")

    def __init__(self, wireframe=None, ser=(), bandFmts=None, **kw):
        self.wireframe = wireframe
        self.ser = list(ser) if ser is not None else []
        self.bandFmts = bandFmts
        super().__init__(**kw)


class SurfaceChart3D(_SurfaceChartBase, _3DBase):
    tagname = "surface3DChart"

    view3D = FIELD_VIEW3D_ON_CHART
    floor = FIELD_FLOOR_ON_CHART
    sideWall = FIELD_SIDE_WALL_ON_CHART
    backWall = FIELD_BACK_WALL_ON_CHART

    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False
    )

    xml_order = _SurfaceChartBase.xml_order + ("axId",)

    def __init__(self, extLst=None, **kw):
        self.extLst = extLst
        self.x_axis = TextAxis()
        self.y_axis = NumericAxis()
        self.z_axis = SeriesAxis()
        _SurfaceChartBase.__init__(self, **kw)
        _3DBase.__init__(self)


class SurfaceChart(SurfaceChart3D):
    tagname = "surfaceChart"

    def __init__(self, **kw):
        super().__init__(**kw)
        self.y_axis.delete = True
        self.view3D.x_rotation = 90
        self.view3D.y_rotation = 0
        self.view3D.perspective = False
        self.view3D.right_angle_axes = False
