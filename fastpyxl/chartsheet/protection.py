
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import Field

from fastpyxl.worksheet.protection import (
    _Protected
)


class ChartsheetProtection(Serialisable, _Protected):
    tagname = "sheetProtection"

    algorithmName: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    hashValue: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    saltValue: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    spinCount: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    content: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    objects: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)

    def __iter__(self):
        for key in ("content", "objects", "password", "hashValue", "spinCount", "saltValue", "algorithmName"):
            value = getattr(self, key, None)
            if value is not None:
                yield key, str(int(value)) if isinstance(value, bool) else str(value)

    def __init__(self,
                 content=None,
                 objects=None,
                 hashValue=None,
                 spinCount=None,
                 saltValue=None,
                 algorithmName=None,
                 password=None,
                 ):
        self.content = content
        self.objects = objects
        self.hashValue = hashValue
        self.spinCount = spinCount
        self.saltValue = saltValue
        self.algorithmName = algorithmName
        if password is not None:
            self.password = password
