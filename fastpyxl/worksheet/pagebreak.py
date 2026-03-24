# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import Field


class Break(Serialisable):

    tagname = "brk"

    id: int | None = Field.attribute(expected_type=int, allow_none=True)
    min: int | None = Field.attribute(expected_type=int, allow_none=True)
    max: int | None = Field.attribute(expected_type=int, allow_none=True)
    man: bool | None = Field.attribute(expected_type=bool, allow_none=True)
    pt: bool | None = Field.attribute(expected_type=bool, allow_none=True)

    def __init__(
        self,
        id=0,
        min=0,
        max=16383,
        man=True,
        pt=None,
    ):
        self.id = id
        self.min = min
        self.max = max
        self.man = man
        self.pt = pt


class RowBreak(Serialisable):

    tagname = "rowBreaks"

    brk: list[Break] = Field.sequence(expected_type=Break, allow_none=True, default=list)

    def __init__(
        self,
        count=None,
        manualBreakCount=None,
        brk=(),
    ):
        del count, manualBreakCount
        self.brk = list(brk)

    def __bool__(self):
        return len(self.brk) > 0

    def __len__(self):
        return len(self.brk)

    @property
    def count(self):
        return len(self)

    @property
    def manualBreakCount(self):
        return len(self)

    def __iter__(self):
        yield "count", str(self.count)
        yield "manualBreakCount", str(self.manualBreakCount)

    def append(self, brk=None):
        """
        Add a page break
        """
        vals = list(self.brk)
        if not isinstance(brk, Break):
            brk = Break(id=self.count + 1)
        vals.append(brk)
        self.brk = vals


PageBreak = RowBreak


class ColBreak(RowBreak):

    tagname = "colBreaks"
