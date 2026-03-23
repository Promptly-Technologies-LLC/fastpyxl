from fastpyxl.chart.data_source import NumVal
from fastpyxl.tests.helper import compare_xml
from fastpyxl.typed_serialisable.base import Serialisable as TypedSerialisable
from fastpyxl.xml.functions import fromstring, tostring


def test_numval_migrated_to_typed_serialisable_base():
    assert issubclass(NumVal, TypedSerialisable)


def test_numval_dynamic_value_preserves_na_sentinel():
    model = NumVal(idx=0, v="#N/A")
    assert model.v == "#N/A"

    xml = tostring(model.to_tree("pt"))
    expected = """
    <pt idx="0">
      <v>#N/A</v>
    </pt>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_numval_dynamic_value_coerces_numeric_to_float():
    model = NumVal(idx=1, v="3.5")
    assert model.v == 3.5

    parsed = NumVal.from_tree(fromstring("<pt idx='2'><v>7.25</v></pt>"))
    assert parsed.idx == 2
    assert parsed.v == 7.25

