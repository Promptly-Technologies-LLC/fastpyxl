# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.drawing.spreadsheet_drawing import AnchorMarker
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import Field
from fastpyxl.xml.constants import SHEET_DRAWING_NS


def _enum_converter(value, allowed, field_name):
    if value is None:
        return None
    if value not in allowed:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return value


class ObjectAnchor(Serialisable):

    tagname = "anchor"

    _from: AnchorMarker | None = Field.element(
        expected_type=AnchorMarker, allow_none=True, xml_name="from", namespace=SHEET_DRAWING_NS
    )
    to: AnchorMarker | None = Field.element(
        expected_type=AnchorMarker, allow_none=True, namespace=SHEET_DRAWING_NS
    )
    moveWithCells: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    sizeWithCells: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    z_order: int | None = Field.attribute(expected_type=int, allow_none=True, hyphenated=True)

    xml_order = ("_from", "to")

    def __init__(
        self,
        _from=None,
        to=None,
        moveWithCells=False,
        sizeWithCells=False,
        z_order=None,
    ):
        self._from = _from
        self.to = to
        self.moveWithCells = moveWithCells
        self.sizeWithCells = sizeWithCells
        self.z_order = z_order

    def to_tree(self, tagname=None, idx=None, namespace=None):
        if self._from is not None:
            self._from.namespace = SHEET_DRAWING_NS
        if self.to is not None:
            self.to.namespace = SHEET_DRAWING_NS
        return super().to_tree(tagname, idx, namespace)


class ObjectPr(Serialisable):

    tagname = "objectPr"

    anchor: ObjectAnchor | None = Field.element(expected_type=ObjectAnchor, allow_none=True)
    locked: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    defaultSize: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    _print: bool | None = Field.attribute(expected_type=bool, allow_none=True, xml_name="print")
    disabled: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    uiObject: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    autoFill: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    autoLine: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    autoPict: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    macro: str | None = Field.attribute(expected_type=str, allow_none=True)
    altText: str | None = Field.attribute(expected_type=str, allow_none=True)
    dde: bool | None = Field.attribute(expected_type=bool, allow_none=True)

    xml_order = ("anchor",)

    def __init__(
        self,
        anchor=None,
        locked=True,
        defaultSize=True,
        _print=True,
        disabled=False,
        uiObject=False,
        autoFill=True,
        autoLine=True,
        autoPict=True,
        macro=None,
        altText=None,
        dde=False,
    ):
        self.anchor = anchor
        self.locked = locked
        self.defaultSize = defaultSize
        self._print = _print
        self.disabled = disabled
        self.uiObject = uiObject
        self.autoFill = autoFill
        self.autoLine = autoLine
        self.autoPict = autoPict
        self.macro = macro
        self.altText = altText
        self.dde = dde


class OleObject(Serialisable):

    tagname = "oleObject"

    objectPr: ObjectPr | None = Field.element(expected_type=ObjectPr, allow_none=True)
    progId: str | None = Field.attribute(expected_type=str, allow_none=True)
    dvAspect: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("DVASPECT_CONTENT", "DVASPECT_ICON"), "dvAspect"),
    )
    link: str | None = Field.attribute(expected_type=str, allow_none=True)
    oleUpdate: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(v, ("OLEUPDATE_ALWAYS", "OLEUPDATE_ONCALL"), "oleUpdate"),
    )
    autoLoad: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    shapeId: int | None = Field.attribute(expected_type=int, allow_none=True)

    xml_order = ("objectPr",)

    def __init__(
        self,
        objectPr=None,
        progId=None,
        dvAspect="DVASPECT_CONTENT",
        link=None,
        oleUpdate=None,
        autoLoad=False,
        shapeId=None,
    ):
        self.objectPr = objectPr
        self.progId = progId
        self.dvAspect = dvAspect
        self.link = link
        self.oleUpdate = oleUpdate
        self.autoLoad = autoLoad
        self.shapeId = shapeId


class OleObjects(Serialisable):

    tagname = "oleObjects"

    oleObject: list[OleObject] = Field.sequence(expected_type=OleObject, default=list)

    xml_order = ("oleObject",)

    def __init__(
        self,
        oleObject=(),
    ):
        self.oleObject = oleObject
