from fastpyxl.pivot.cache import GroupMember, LevelGroup
from fastpyxl.tests.helper import compare_xml
from fastpyxl.typed_serialisable.base import Serialisable as TypedSerialisable
from fastpyxl.xml.functions import fromstring, tostring


def test_levelgroup_migrated_to_typed_serialisable_base():
    assert issubclass(LevelGroup, TypedSerialisable)


def test_levelgroup_nested_sequence_count_serialization_stability():
    model = LevelGroup(
        name="Group A",
        uniqueName="[Dim].[Group A]",
        caption="Group A",
        uniqueParent="[Dim].[All]",
        id=1,
        groupMembers=[
            GroupMember(uniqueName="[Dim].[Item 1]", group=True),
            GroupMember(uniqueName="[Dim].[Item 2]", group=False),
        ],
    )
    xml = tostring(model.to_tree())
    expected = """
    <group name="Group A" uniqueName="[Dim].[Group A]" caption="Group A" uniqueParent="[Dim].[All]" id="1">
      <groupMembers count="2">
        <groupMember uniqueName="[Dim].[Item 1]" group="1"></groupMember>
        <groupMember uniqueName="[Dim].[Item 2]" group="0"></groupMember>
      </groupMembers>
    </group>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_levelgroup_nested_sequence_count_parse_roundtrip():
    src = """
    <group name="Group A" uniqueName="[Dim].[Group A]" caption="Group A" uniqueParent="[Dim].[All]" id="1">
      <groupMembers count="1">
        <groupMember uniqueName="[Dim].[Item 1]" group="1"/>
      </groupMembers>
    </group>
    """
    model = LevelGroup.from_tree(fromstring(src))
    assert model.id == 1
    assert model.groupMembers is not None
    assert len(model.groupMembers) == 1
    assert model.groupMembers[0].uniqueName == "[Dim].[Item 1]"
    assert model.groupMembers[0].group is True
