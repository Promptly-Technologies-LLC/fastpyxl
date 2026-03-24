# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import Field


class Protection(Serialisable):
    """Protection options for use in styles."""

    tagname = "protection"

    locked: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    hidden: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)

    def __init__(self, locked=True, hidden=False):
        self.locked = locked
        self.hidden = hidden
