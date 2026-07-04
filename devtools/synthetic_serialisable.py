"""Typed :class:`~fastpyxl.typed_serialisable.base.Serialisable` models for benchmarks and profiling."""

from __future__ import annotations

from fastpyxl.typed_serialisable.base import Serialisable as TypedSerialisable
from fastpyxl.typed_serialisable.fields import Field
from fastpyxl.xml.functions import fromstring, tostring


class TypedChild(TypedSerialisable):
    tagname = "child"
    value: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

    def __init__(self, value=None):
        self.value = value


class TypedParent(TypedSerialisable):
    tagname = "parent"
    id: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    children: list[TypedChild] = Field.sequence(expected_type=TypedChild, default=list)

    def __init__(self, id=None, children=()):
        self.id = id
        self.children = list(children)


def build_typed(n: int = 2000) -> list[TypedParent]:
    return [
        TypedParent(id=i, children=[TypedChild(value=i), TypedChild(value=i + 1)])
        for i in range(n)
    ]


def render_typed(n: int = 1200) -> int:
    items = build_typed(n)
    size = 0
    for item in items:
        size += len(tostring(item.to_tree()))
    return size


def parse_typed(n: int = 1200) -> int:
    xml = tostring(TypedParent(id=1, children=[TypedChild(value=1), TypedChild(value=2)]).to_tree())
    size = 0
    for _ in range(n):
        obj = TypedParent.from_tree(fromstring(xml))
        size += obj.id or 0
    return size
