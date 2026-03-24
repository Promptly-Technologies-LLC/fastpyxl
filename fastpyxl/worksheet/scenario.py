# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import Field

from .cell_range import MultiCellRange


class InputCells(Serialisable):

    tagname = "inputCells"

    r: str | None = Field.attribute(expected_type=str, allow_none=True)
    deleted: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    undone: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    val: str | None = Field.attribute(expected_type=str, allow_none=True)
    numFmtId: int | None = Field.attribute(expected_type=int, allow_none=True)

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
    name: str | None = Field.attribute(expected_type=str, allow_none=True)
    locked: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    hidden: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    user: str | None = Field.attribute(expected_type=str, allow_none=True)
    comment: str | None = Field.attribute(expected_type=str, allow_none=True)

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
        self.inputCells = inputCells
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
    current: int | None = Field.attribute(expected_type=int, allow_none=True)
    show: int | None = Field.attribute(expected_type=int, allow_none=True)
    sqref: MultiCellRange | None = Field.attribute(expected_type=MultiCellRange, allow_none=True)

    xml_order = ("scenario",)

    def __init__(self,
                 scenario=(),
                 current=None,
                 show=None,
                 sqref=None,
                ):
        self.scenario = scenario
        self.current = current
        self.show = show
        self.sqref = sqref

    def append(self, scenario):
        s = list(self.scenario)
        s.append(scenario)
        self.scenario = s

    def __bool__(self):
        return bool(self.scenario)
