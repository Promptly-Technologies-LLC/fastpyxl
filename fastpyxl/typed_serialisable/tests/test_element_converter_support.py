import pytest

from fastpyxl.tests.helper import compare_xml
from fastpyxl.typed_serialisable.base import Serialisable as TypedSerialisable
from fastpyxl.typed_serialisable.fields import Field
from fastpyxl.xml.functions import Element, fromstring, tostring


class LegacyThing:
    tagname = "thing"

    def __init__(self, value: int):
        self.value = value

    def to_tree(self, tagname=None, idx=None, namespace=None):
        del idx, namespace
        if tagname is None:
            tagname = self.tagname
        return Element(tagname, {"value": str(self.value)})

    @classmethod
    def from_tree(cls, node):
        return cls(int(node.get("value")))

    def __eq__(self, other):
        return isinstance(other, LegacyThing) and self.value == other.value


def test_field_element_accepts_converter_and_converts_before_render_and_parse():
    # Red until typed runtime supports `converter=` on Field.element.
    def convert(v):
        if isinstance(v, LegacyThing):
            return v
        return LegacyThing(int(v))

    class Parent(TypedSerialisable):
        tagname = "parent"
        thing: LegacyThing | None = Field.element(
            expected_type=LegacyThing,
            allow_none=True,
            converter=convert,
            xml_name="thing",
        )

    obj = Parent(thing="5")
    xml = tostring(obj.to_tree())
    expected = """
    <parent>
      <thing value="5"/>
    </parent>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff

    parsed = Parent.from_tree(fromstring(xml))
    assert parsed.thing == LegacyThing(5)

