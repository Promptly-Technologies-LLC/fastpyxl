# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import Field
from fastpyxl.xml.constants import REL_NS

from .ole import ObjectAnchor


class ControlProperty(Serialisable):

    tagname = "controlPr"

    anchor: ObjectAnchor | None = Field.element(expected_type=ObjectAnchor, allow_none=True)
    locked: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    defaultSize: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    _print: bool | None = Field.attribute(expected_type=bool, allow_none=True, xml_name="print")
    disabled: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    recalcAlways: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    uiObject: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    autoFill: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    autoLine: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    autoPict: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    macro: str | None = Field.attribute(expected_type=str, allow_none=True)
    altText: str | None = Field.attribute(expected_type=str, allow_none=True)
    linkedCell: str | None = Field.attribute(expected_type=str, allow_none=True)
    listFillRange: str | None = Field.attribute(expected_type=str, allow_none=True)
    cf: str | None = Field.attribute(expected_type=str, allow_none=True)
    id: str | None = Field.attribute(expected_type=str, allow_none=True, namespace=REL_NS)

    xml_order = ("anchor",)

    def __init__(self,
                 anchor=None,
                 locked=True,
                 defaultSize=True,
                 _print=True,
                 disabled=False,
                 recalcAlways=False,
                 uiObject=False,
                 autoFill=True,
                 autoLine=True,
                 autoPict=True,
                 macro=None,
                 altText=None,
                 linkedCell=None,
                 listFillRange=None,
                 cf='pict',
                 id=None,
                ):
        self.anchor = anchor
        self.locked = locked
        self.defaultSize = defaultSize
        self._print = _print
        self.disabled = disabled
        self.recalcAlways = recalcAlways
        self.uiObject = uiObject
        self.autoFill = autoFill
        self.autoLine = autoLine
        self.autoPict = autoPict
        self.macro = macro
        self.altText = altText
        self.linkedCell = linkedCell
        self.listFillRange = listFillRange
        self.cf = cf
        self.id = id


class Control(Serialisable):

    tagname = "control"

    controlPr: ControlProperty | None = Field.element(expected_type=ControlProperty, allow_none=True)
    shapeId: int | None = Field.attribute(expected_type=int, allow_none=True)
    name: str | None = Field.attribute(expected_type=str, allow_none=True)

    xml_order = ("controlPr",)

    def __init__(self,
                 controlPr=None,
                 shapeId=None,
                 name=None,
                ):
        self.controlPr = controlPr
        self.shapeId = shapeId
        self.name = name


class Controls(Serialisable):

    tagname = "controls"

    control: list[Control] = Field.sequence(expected_type=Control, default=list)

    xml_order = ("control",)

    def __init__(self,
                 control=(),
                ):
        self.control = control
