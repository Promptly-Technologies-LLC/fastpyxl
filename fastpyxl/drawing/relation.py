# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.xml.constants import CHART_NS, REL_NS

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import Field


class ChartRelation(Serialisable):

    tagname = "chart"
    namespace = CHART_NS

    id: str | None = Field.attribute(expected_type=str, allow_none=True, namespace=REL_NS, default=None)

    def __init__(self, id=None):
        self.id = id
