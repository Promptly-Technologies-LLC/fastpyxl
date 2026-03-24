# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import Field

class FunctionGroup(Serialisable):

    tagname = "functionGroup"

    name: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)

    def __init__(self,
                 name=None,
                ):
        self.name = name


class FunctionGroupList(Serialisable):

    tagname = "functionGroups"

    builtInGroupCount: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    functionGroup: list[FunctionGroup] | None = Field.sequence(expected_type=FunctionGroup, allow_none=True, default=list)
    xml_order = ("functionGroup",)

    def __init__(self,
                 builtInGroupCount=16,
                 functionGroup=(),
                ):
        self.builtInGroupCount = builtInGroupCount
        self.functionGroup = list(functionGroup)
