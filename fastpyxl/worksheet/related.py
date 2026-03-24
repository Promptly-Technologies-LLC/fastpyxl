# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import Field
from fastpyxl.xml.constants import REL_NS


class Related(Serialisable):

    id: str | None = Field.attribute(expected_type=str, allow_none=True, namespace=REL_NS)

    def __init__(self, id=None):
        self.id = id

    def to_tree(self, tagname, idx=None, namespace=None):
        return super().to_tree(tagname, idx, namespace)
