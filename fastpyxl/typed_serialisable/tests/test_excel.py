from fastpyxl.tests.helper import compare_xml
from fastpyxl.typed_serialisable.excel import (
    Extension,
    ExtensionList,
    NestedValInt,
    explicit_none_element,
)
from fastpyxl.xml.constants import CHART_NS
from fastpyxl.xml.functions import fromstring, tostring


def test_extension_list_roundtrip():
    ext_lst = ExtensionList(ext=[Extension(uri="http://example.com/ext")])
    xml = tostring(ext_lst.to_tree())
    expected = """
    <extLst>
      <ext uri="http://example.com/ext"/>
    </extLst>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff

    parsed = ExtensionList.from_tree(fromstring(xml))
    assert parsed.ext == [Extension(uri="http://example.com/ext")]


def test_nested_val_int_roundtrip():
    node = NestedValInt(val=7)
    xml = tostring(node.to_tree())
    expected = """
    <x>
      <val val="7"/>
    </x>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff

    assert NestedValInt.from_tree(fromstring(xml)) == NestedValInt(val=7)


def test_explicit_none_element_without_namespace():
    el = explicit_none_element("tickLblSkip", "none")
    assert el.tag == "tickLblSkip"
    assert el.get("val") == "none"


def test_explicit_none_element_with_namespace():
    el = explicit_none_element("tickLblSkip", "none", namespace=CHART_NS)
    assert el.tag == "{%s}tickLblSkip" % CHART_NS
    assert el.get("val") == "none"
