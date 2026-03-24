import pytest

from fastpyxl.descriptors import Integer, Typed
from fastpyxl.descriptors.serialisable import Serialisable as LegacySerialisable
from fastpyxl.tests.helper import compare_xml
from fastpyxl.xml.functions import Element, fromstring, tostring

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import AliasField, Field


class Child(Serialisable):
    tagname = "child"
    value: int | None = Field.attribute(expected_type=int, allow_none=True)


class LegacyChild(LegacySerialisable):
    tagname = "legacyChild"
    value = Integer()

    def __init__(self, value):
        self.value = value


class MixedNewParent(Serialisable):
    tagname = "mixedNewParent"
    old_child: LegacyChild | None = Field.element(expected_type=LegacyChild, allow_none=True)


class MixedLegacyParent(LegacySerialisable):
    tagname = "mixedLegacyParent"
    new_child = Typed(expected_type=Child, allow_none=True)

    def __init__(self, new_child=None):
        self.new_child = new_child


class MixedGraphNewParent(Serialisable):
    tagname = "mixedGraphNewParent"
    old_children: list[LegacyChild] = Field.sequence(expected_type=LegacyChild, default=list)
    new_child: Child | None = Field.element(expected_type=Child, allow_none=True)


class Demo(Serialisable):
    tagname = "demo"

    attr_text: str | None = Field.attribute(expected_type=str, allow_none=True)
    attr_id: int | None = Field.attribute(expected_type=int, allow_none=True, xml_name="id")
    hyphen_name: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        hyphenated=True,
        xml_name="hyphen-name",
    )
    _class: str | None = Field.attribute(expected_type=str, allow_none=True, xml_name="class")
    val_item: int | None = Field.nested_value(expected_type=int, allow_none=True)
    text_item: str | None = Field.nested_text(expected_type=str, allow_none=True)
    child: Child | None = Field.element(expected_type=Child, allow_none=True)
    items: list[int] = Field.sequence(expected_type=int, default=list)
    nested_items: list[Child] = Field.nested_sequence(expected_type=Child, count=True, default=list)
    flag: bool | None = Field.nested_bool(allow_none=True)
    short: bool | None = Field.nested_bool(allow_none=True, renderer=lambda tag, value, ns=None: None if value is None else Element(tag))
    alias_attr: int | None = AliasField("attr_id")


class NamespaceDemo(Serialisable):
    tagname = "namespaceDemo"
    rel_id: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        xml_name="id",
        namespace="urn:test-rel",
    )


class MultiA(Serialisable):
    tagname = "a"
    value: int | None = Field.attribute(expected_type=int, allow_none=True)


class MultiB(Serialisable):
    tagname = "b"
    label: str | None = Field.attribute(expected_type=str, allow_none=True)


class MultiDemo(Serialisable):
    tagname = "multiDemo"
    entries: list[MultiA | MultiB] = Field.multi_sequence(
        parts={
            "a": MultiA,
            "b": MultiB,
        },
        default=list,
    )


class OrderedDemo(Serialisable):
    tagname = "orderedDemo"
    xml_order = ("second", "first")
    first: int | None = Field.nested_value(expected_type=int, allow_none=True)
    second: int | None = Field.nested_value(expected_type=int, allow_none=True)


def test_field_compilation_and_caches():
    assert "attr_id" in Demo.__fields__
    assert Demo.__aliases__["alias_attr"] == "attr_id"
    assert "attr_id" in Demo.__attrs__
    assert "child" in Demo.__elements__
    assert "val_item" in Demo.__nested__


def test_namespace_attribute_roundtrip():
    obj = NamespaceDemo(rel_id="rId1")
    tree = obj.to_tree()
    assert tree.get("{urn:test-rel}id") == "rId1"

    parsed = NamespaceDemo.from_tree(tree)
    assert parsed.rel_id == "rId1"


def test_multi_sequence_roundtrip_tag_dispatch():
    model = MultiDemo(entries=[MultiA(value=1), MultiB(label="x"), MultiA(value=2)])
    xml = tostring(model.to_tree())
    expected = """
    <multiDemo>
      <a value="1"></a>
      <b label="x"></b>
      <a value="2"></a>
    </multiDemo>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff

    parsed = MultiDemo.from_tree(fromstring(xml))
    assert len(parsed.entries) == 3
    assert isinstance(parsed.entries[0], MultiA)
    assert isinstance(parsed.entries[1], MultiB)
    assert isinstance(parsed.entries[2], MultiA)
    assert parsed.entries[0].value == 1
    assert parsed.entries[1].label == "x"
    assert parsed.entries[2].value == 2


def test_xml_order_override_for_elements():
    model = OrderedDemo(first=1, second=2)
    xml = tostring(model.to_tree())
    expected = """
    <orderedDemo>
      <second val="2"></second>
      <first val="1"></first>
    </orderedDemo>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_scalar_assignment_validation_and_alias_forwarding():
    demo = Demo()
    demo.attr_id = "7"  # ty: ignore[invalid-assignment]
    assert demo.attr_id == 7
    demo.alias_attr = 9
    assert demo.attr_id == 9
    with pytest.raises(TypeError):
        demo.items = "not-a-sequence"  # ty: ignore[invalid-assignment]


def test_parse_and_render_roundtrip_for_field_strategies():
    demo = Demo(
        attr_text="inline",
        attr_id=3,
        hyphen_name="x",
        _class="y",
        val_item=5,
        text_item="text",
        child=Child(value=8),
        items=[1, "2", 3],
        nested_items=[Child(value=10), Child(value=11)],
        flag=True,
    )
    xml = tostring(demo.to_tree())
    expected = """
    <demo id="3" hyphen-name="x" class="y" attr_text="inline">
      <val_item val="5"></val_item>
      <text_item>text</text_item>
      <child value="8"></child>
      <items>1</items>
      <items>2</items>
      <items>3</items>
      <nested_items count="2">
        <child value="10"></child>
        <child value="11"></child>
      </nested_items>
      <flag val="1"></flag>
    </demo>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff

    parsed = Demo.from_tree(fromstring(xml))
    assert parsed == demo


def test_interop_new_parent_with_legacy_child():
    src = "<mixedNewParent><old_child value='4'></old_child></mixedNewParent>"
    parsed = MixedNewParent.from_tree(fromstring(src))
    assert isinstance(parsed.old_child, LegacyChild)
    assert parsed.old_child.value == 4


def test_interop_legacy_parent_with_new_child():
    parent = MixedLegacyParent(new_child=Child(value=12))
    xml = tostring(parent.to_tree())
    expected = """
    <mixedLegacyParent>
      <new_child value="12"></new_child>
    </mixedLegacyParent>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_mixed_graph_roundtrip_new_parent_with_legacy_sequence():
    parent = MixedGraphNewParent(
        old_children=[LegacyChild(1), LegacyChild(2)],
        new_child=Child(value=9),
    )
    xml = tostring(parent.to_tree())
    expected = """
    <mixedGraphNewParent>
      <old_children value="1"></old_children>
      <old_children value="2"></old_children>
      <new_child value="9"></new_child>
    </mixedGraphNewParent>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff

    parsed = MixedGraphNewParent.from_tree(fromstring(xml))
    assert isinstance(parsed.old_children[0], LegacyChild)
    assert parsed.old_children[0].value == 1
    assert parsed.old_children[1].value == 2
    assert isinstance(parsed.new_child, Child)
    assert parsed.new_child.value == 9
