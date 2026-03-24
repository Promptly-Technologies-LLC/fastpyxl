# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import Field

from .cell_range import MultiCellRange


class InputCells(Serialisable):

    tagname = "inputCells"

    r: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    deleted: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    undone: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    val: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    numFmtId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

    def __init__(self,
                 r=None,
                 deleted=False,
                 undone=False,
                 val=None,
                 numFmtId=None,
                ):
        self.r = r
        self.deleted = deleted
        self.undone = undone
        self.val = val
        self.numFmtId = numFmtId


class Scenario(Serialisable):

    tagname = "scenario"

    inputCells: list[InputCells] = Field.sequence(expected_type=InputCells, default=list)
    name: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    locked: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    hidden: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    user: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    comment: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)

    xml_order = ("inputCells",)

    def __init__(self,
                 inputCells=(),
                 name=None,
                 locked=False,
                 hidden=False,
                 count=None,
                 user=None,
                 comment=None,
                ):
        del count
        self.inputCells = list(inputCells)
        self.name = name
        self.locked = locked
        self.hidden = hidden
        self.user = user
        self.comment = comment

    @property
    def count(self):
        return len(self.inputCells)

    def __iter__(self):
        for k, v in super().__iter__():
            yield k, v
        yield "count", str(self.count)


class ScenarioList(Serialisable):

    tagname = "scenarios"

    scenario: list[Scenario] = Field.sequence(expected_type=Scenario, default=list)
    current: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    show: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    sqref: MultiCellRange | None = Field.attribute(expected_type=MultiCellRange, allow_none=True, default=None)

    xml_order = ("scenario",)

    def __init__(self,
                 scenario=(),
                 current=None,
                 show=None,
                 sqref=None,
                ):
        self.scenario = list(scenario)
        self.current = current
        self.show = show
        self.sqref = sqref

    def append(self, scenario):
        s = list(self.scenario)
        s.append(scenario)
        self.scenario = s

    def __bool__(self):
        return bool(self.scenario)
