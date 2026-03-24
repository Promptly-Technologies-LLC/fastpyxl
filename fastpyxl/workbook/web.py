# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import Field


class WebPublishObject(Serialisable):

    tagname = "webPublishingObject"

    id: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    divId: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    sourceObject: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    destinationFile: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    title: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    autoRepublish: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)

    def __init__(self,
                 id=None,
                 divId=None,
                 sourceObject=None,
                 destinationFile=None,
                 title=None,
                 autoRepublish=None,
                ):
        self.id = id
        self.divId = divId
        self.sourceObject = sourceObject
        self.destinationFile = destinationFile
        self.title = title
        self.autoRepublish = autoRepublish


class WebPublishObjectList(Serialisable):

    tagname ="webPublishingObjects"

    webPublishObject: list[WebPublishObject] = Field.sequence(
        expected_type=WebPublishObject,
        xml_name="webPublishingObject", default=list,
    )
    xml_order = ('webPublishObject',)

    def __init__(self,
                 count=None,
                 webPublishObject=(),
                ):
        del count
        self.webPublishObject = list(webPublishObject)


    @property
    def count(self):
        return len(self.webPublishObject)


    def __iter__(self):
        if self.webPublishObject:
            yield "count", str(self.count)


class WebPublishing(Serialisable):

    tagname = "webPublishing"

    css: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    thicket: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    longFileNames: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    vml: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    allowPng: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    targetScreenSize: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_converter(
            v,
            ("544x376", "640x480", "720x512", "800x600", "1024x768", "1152x882", "1152x900",
             "1280x1024", "1600x1200", "1800x1440", "1920x1200"),
            "targetScreenSize",
        ), default=None,
    )
    dpi: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    codePage: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    characterSet: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)

    def __init__(self,
                 css=None,
                 thicket=None,
                 longFileNames=None,
                 vml=None,
                 allowPng=None,
                 targetScreenSize='800x600',
                 dpi=None,
                 codePage=None,
                 characterSet=None,
                ):
        self.css = css
        self.thicket = thicket
        self.longFileNames = longFileNames
        self.vml = vml
        self.allowPng = allowPng
        self.targetScreenSize = targetScreenSize
        self.dpi = dpi
        self.codePage = codePage
        self.characterSet = characterSet


def _enum_converter(value, allowed_values, field_name: str):
    if value is None:
        return None
    if value not in allowed_values:
        raise FieldValidationError(f"{field_name} rejected value {value!r}")
    return value
