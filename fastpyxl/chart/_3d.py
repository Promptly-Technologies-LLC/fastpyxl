# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.descriptors.excel import ExtensionList

from fastpyxl.typed_serialisable.errors import FieldValidationError

from .marker import PictureOptions
from .shapes import GraphicalProperties


def _view_rot_int(v, *, lo: int, hi: int, name: str):
    if v is None:
        return None
    try:
        n = int(v)
    except (TypeError, ValueError) as exc:
        raise FieldValidationError(f"{name} rejected value {v!r}") from exc
    if n < lo or n > hi:
        raise FieldValidationError(f"{name} rejected value {v!r}")
    return n


def _view_rot_float(v, *, lo: float, hi: float, name: str):
    if v is None:
        return None
    try:
        x = float(v)
    except (TypeError, ValueError) as exc:
        raise FieldValidationError(f"{name} rejected value {v!r}") from exc
    if x < lo or x > hi:
        raise FieldValidationError(f"{name} rejected value {v!r}")
    return x


class View3D(Serialisable):
    tagname = "view3D"

    rotX: float | None = Field.nested_value(
        expected_type=float,
        allow_none=True,
        converter=lambda v: _view_rot_float(v, lo=-90, hi=90, name="rotX"), default=None,
    )
    x_rotation = AliasField("rotX", default=None)
    hPercent: float | None = Field.nested_value(
        expected_type=float,
        allow_none=True,
        converter=lambda v: _view_rot_float(v, lo=5, hi=500, name="hPercent"), default=None,
    )
    height_percent = AliasField("hPercent", default=None)
    rotY: int | None = Field.nested_value(
        expected_type=int,
        allow_none=True,
        converter=lambda v: _view_rot_int(v, lo=-90, hi=90, name="rotY"), default=None,
    )
    y_rotation = AliasField("rotY", default=None)
    depthPercent: int | None = Field.nested_value(expected_type=int, allow_none=True, default=None)
    rAngAx: bool | None = Field.nested_bool(allow_none=True, default=None)
    right_angle_axes = AliasField("rAngAx", default=None)
    perspective: int | None = Field.nested_value(expected_type=int, allow_none=True, default=None)
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False, default=None
    )

    xml_order = ("rotX", "hPercent", "rotY", "depthPercent", "rAngAx", "perspective")

    def __init__(
        self,
        rotX=15,
        hPercent=None,
        rotY=20,
        depthPercent=None,
        rAngAx=True,
        perspective=None,
        extLst=None,
    ):
        self.rotX = rotX
        self.hPercent = hPercent
        self.rotY = rotY
        self.depthPercent = depthPercent
        self.rAngAx = rAngAx
        self.perspective = perspective
        self.extLst = extLst


class Surface(Serialisable):
    tagname = "surface"

    thickness: int | None = Field.nested_value(expected_type=int, allow_none=True, default=None)
    spPr: GraphicalProperties | None = Field.element(
        expected_type=GraphicalProperties, allow_none=True, default=None
    )
    graphicalProperties = AliasField("spPr", default=None)
    pictureOptions: PictureOptions | None = Field.element(
        expected_type=PictureOptions, allow_none=True, default=None
    )
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False, default=None
    )

    xml_order = ("thickness", "spPr", "pictureOptions")

    def __init__(
        self,
        thickness=None,
        spPr=None,
        pictureOptions=None,
        extLst=None,
    ):
        self.thickness = thickness
        self.spPr = spPr
        self.pictureOptions = pictureOptions
        self.extLst = extLst

FIELD_VIEW3D = Field.element(expected_type=View3D, allow_none=True, default=None)
FIELD_FLOOR = Field.element(expected_type=Surface, allow_none=True, default=None)
FIELD_SIDE_WALL = Field.element(expected_type=Surface, allow_none=True, default=None)
FIELD_BACK_WALL = Field.element(expected_type=Surface, allow_none=True, default=None)

FIELD_VIEW3D_ON_CHART = Field.element(expected_type=View3D, allow_none=True, serialize=False, default=None)
FIELD_FLOOR_ON_CHART = Field.element(expected_type=Surface, allow_none=True, serialize=False, default=None)
FIELD_SIDE_WALL_ON_CHART = Field.element(expected_type=Surface, allow_none=True, serialize=False, default=None)
FIELD_BACK_WALL_ON_CHART = Field.element(expected_type=Surface, allow_none=True, serialize=False, default=None)


class _3DBase(Serialisable):
    """
    Base class for 3D charts
    """

    tagname = "ChartBase"

    view3D = FIELD_VIEW3D
    floor = FIELD_FLOOR
    sideWall = FIELD_SIDE_WALL
    backWall = FIELD_BACK_WALL

    xml_order = ("backWall", "floor", "sideWall", "view3D")

    def __init__(
        self,
        view3D=None,
        floor=None,
        sideWall=None,
        backWall=None,
        **kw,
    ):
        if view3D is None:
            view3D = View3D()
        self.view3D = view3D
        if floor is None:
            floor = Surface()
        self.floor = floor
        if sideWall is None:
            sideWall = Surface()
        self.sideWall = sideWall
        if backWall is None:
            backWall = Surface()
        self.backWall = backWall
