import importlib
import pkgutil

import fastpyxl.chart as chart_pkg
from fastpyxl.chart.axis import ChartLines
from fastpyxl.chart.data_source import AxDataSource, NumRef, StrRef
from fastpyxl.chart.layout import Layout
from fastpyxl.chart.title import Title
from fastpyxl.tests.helper import compare_xml
from fastpyxl.typed_serialisable.base import Serialisable as TypedSerialisable
from fastpyxl.xml.functions import fromstring, tostring


def test_phase5b_chart_package_models_subclass_typed_serialisable():
    skip = frozenset({"tests", "reader"})
    for mi in pkgutil.iter_modules(chart_pkg.__path__):
        if mi.name in skip or mi.ispkg:
            continue
        mod = importlib.import_module(f"fastpyxl.chart.{mi.name}")
        for attr_name in dir(mod):
            if attr_name.startswith("__"):
                continue
            obj = getattr(mod, attr_name)
            if not isinstance(obj, type):
                continue
            if getattr(obj, "__module__", None) != mod.__name__:
                continue
            if not issubclass(obj, TypedSerialisable):
                continue
            assert issubclass(obj, TypedSerialisable)


def test_phase5b_num_ref_roundtrip():
    src = """
    <numRef>
      <f>Sheet1!$A$1:$A$4</f>
    </numRef>
    """
    m = NumRef.from_tree(fromstring(src))
    assert m.f == "Sheet1!$A$1:$A$4"
    out = tostring(m.to_tree("numRef"))
    diff = compare_xml(out, src)
    assert diff is None, diff


def test_phase5b_str_ref_roundtrip():
    src = """
    <strRef>
      <f>Sheet1!$B$1:$B$2</f>
    </strRef>
    """
    m = StrRef.from_tree(fromstring(src))
    assert m.f == "Sheet1!$B$1:$B$2"
    out = tostring(m.to_tree("strRef"))
    diff = compare_xml(out, src)
    assert diff is None, diff


def test_phase5b_layout_roundtrip():
    src = "<layout/>"
    m = Layout.from_tree(fromstring(src))
    out = tostring(m.to_tree())
    diff = compare_xml(out, src)
    assert diff is None, diff


def test_phase5b_title_overlay_roundtrip():
    src = "<title><overlay val=\"0\"/></title>"
    m = Title.from_tree(fromstring(src))
    assert m.overlay is False
    out = tostring(m.to_tree())
    back = Title.from_tree(fromstring(out))
    assert back.overlay is False


def test_phase5b_chart_lines_roundtrip():
    src = "<chartLines/>"
    m = ChartLines.from_tree(fromstring(src))
    out = tostring(m.to_tree("majorGridlines"))
    expected = "<majorGridlines/>"
    diff = compare_xml(out, expected)
    assert diff is None, diff


def test_phase5b_ax_data_source_num_ref_roundtrip():
    src = """
    <cat>
      <numRef>
        <f>'S'!$A$1:$A$2</f>
      </numRef>
    </cat>
    """
    m = AxDataSource.from_tree(fromstring(src))
    assert m.numRef is not None
    assert m.numRef.f == "'S'!$A$1:$A$2"
    out = tostring(m.to_tree())
    diff = compare_xml(out, src)
    assert diff is None, diff
