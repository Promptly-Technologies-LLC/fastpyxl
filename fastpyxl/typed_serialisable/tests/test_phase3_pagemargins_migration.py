from fastpyxl.tests.helper import compare_xml
from fastpyxl.typed_serialisable.base import Serialisable as TypedSerialisable
from fastpyxl.worksheet.page import PageMargins
from fastpyxl.xml.functions import fromstring, tostring


def test_pagemargins_migrated_to_typed_serialisable_base():
    assert issubclass(PageMargins, TypedSerialisable)


def test_pagemargins_default_serialization_stability():
    model = PageMargins()
    xml = tostring(model.to_tree())
    expected = """
    <pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"></pageMargins>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_pagemargins_from_tree_roundtrip():
    src = "<pageMargins left='2' right='2' top='1' bottom='1' header='0.3' footer='0.3'/>"
    model = PageMargins.from_tree(fromstring(src))
    assert model.left == 2.0
    assert model.right == 2.0
    assert model.header == 0.3
    assert model.footer == 0.3

