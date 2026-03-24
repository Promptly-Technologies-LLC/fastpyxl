from fastpyxl.pivot.cache import SharedItems
from fastpyxl.pivot.record import Record, RecordList
from fastpyxl.pivot.table import (
    AutoSortScope,
    ColHierarchiesUsage,
    HierarchyUsage,
    MemberList,
    MemberProperty,
    PivotTableStyle,
    RowColField,
    RowHierarchiesUsage,
)
from fastpyxl.tests.helper import compare_xml
from fastpyxl.typed_serialisable.base import Serialisable as TypedSerialisable
from fastpyxl.xml.functions import fromstring, tostring


def test_phase6_record_models_subclass_typed_serialisable():
    assert issubclass(Record, TypedSerialisable)
    assert issubclass(RecordList, TypedSerialisable)


def test_phase6_record_multi_sequence_roundtrip():
    src = "<r><n v=\"1\"/><s v=\"hello\"/></r>"
    model = Record.from_tree(fromstring(src))
    assert len(model._fields) == 2
    out = tostring(model.to_tree())
    diff = compare_xml(out, src)
    assert diff is None, diff


def test_phase6_shared_items_multi_sequence_roundtrip():
    src = """
    <sharedItems count="2" containsNumber="1" containsString="1">
      <n v="2"/>
      <s v="done"/>
    </sharedItems>
    """
    model = SharedItems.from_tree(fromstring(src))
    assert len(model._fields) == 2
    out = tostring(model.to_tree())
    diff = compare_xml(out, src)
    assert diff is None, diff


def test_phase6_table_slice_subclasses_typed_serialisable():
    for cls in (
        HierarchyUsage,
        ColHierarchiesUsage,
        RowHierarchiesUsage,
        PivotTableStyle,
        MemberList,
        MemberProperty,
        RowColField,
        AutoSortScope,
    ):
        assert issubclass(cls, TypedSerialisable)


def test_phase6_table_member_list_count_attribute_roundtrip():
    src = '<members count="2" level="1"><member name="A"/><member name="B"/></members>'
    model = MemberList.from_tree(fromstring(src))
    assert model.member == ["A", "B"]
    out = tostring(model.to_tree())
    diff = compare_xml(out, src)
    assert diff is None, diff
