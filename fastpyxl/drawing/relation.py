# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.xml.constants import CHART_NS

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import Field


class ChartRelation(Serialisable):

    tagname = "chart"
    namespace = CHART_NS

    id: str | None = Field.attribute(expected_type=str, allow_none=True)

    def __init__(self, id):
        self.id = id
