"""Legacy vs typed :class:`~fastpyxl.descriptors.serialisable.Serialisable` models for benchmarks and profiling."""

from __future__ import annotations

from fastpyxl.descriptors import Integer, Sequence
from fastpyxl.descriptors.serialisable import Serialisable as LegacySerialisable
from fastpyxl.typed_serialisable.base import Serialisable as TypedSerialisable
from fastpyxl.typed_serialisable.fields import Field
from fastpyxl.xml.functions import fromstring, tostring


class LegacyChild(LegacySerialisable):
    tagname = "child"
    value = Integer()

    def __init__(self, value=None):
        self.value = value


class LegacyParent(LegacySerialisable):
    tagname = "parent"
    id = Integer(allow_none=True)
    children = Sequence(expected_type=LegacyChild)

    def __init__(self, id=None, children=()):
        self.id = id
        self.children = children


class TypedChild(TypedSerialisable):
    tagname = "child"
    value: int | None = Field.attribute(expected_type=int, allow_none=True)


class TypedParent(TypedSerialisable):
    tagname = "parent"
    id: int | None = Field.attribute(expected_type=int, allow_none=True)
    children: list[TypedChild] = Field.sequence(expected_type=TypedChild, default=list)


def build_legacy(n: int = 2000) -> list[LegacyParent]:
    return [
        LegacyParent(id=i, children=[LegacyChild(value=i), LegacyChild(value=i + 1)])
        for i in range(n)
    ]


def build_typed(n: int = 2000) -> list[TypedParent]:
    return [
        TypedParent(id=i, children=[TypedChild(value=i), TypedChild(value=i + 1)])
        for i in range(n)
    ]


def render_legacy(n: int = 1200) -> int:
    items = build_legacy(n)
    size = 0
    for item in items:
        size += len(tostring(item.to_tree()))
    return size


def render_typed(n: int = 1200) -> int:
    items = build_typed(n)
    size = 0
    for item in items:
        size += len(tostring(item.to_tree()))
    return size


def parse_legacy(n: int = 1200) -> int:
    xml = tostring(LegacyParent(id=1, children=[LegacyChild(1), LegacyChild(2)]).to_tree())
    size = 0
    for _ in range(n):
        obj = LegacyParent.from_tree(fromstring(xml))
        size += obj.id or 0
    return size


def parse_typed(n: int = 1200) -> int:
    xml = tostring(TypedParent(id=1, children=[TypedChild(value=1), TypedChild(value=2)]).to_tree())
    size = 0
    for _ in range(n):
        obj = TypedParent.from_tree(fromstring(xml))
        size += obj.id or 0
    return size
