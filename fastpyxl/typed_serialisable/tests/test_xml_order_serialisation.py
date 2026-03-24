"""Tests for XML element emission order (ordering parity)."""

import pytest

from fastpyxl.xml.functions import localname, tostring

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import Field


class _OrderProbe(Serialisable):
    tagname = "probe"

    xml_order = ("b", "a")
    a: int | None = Field.nested_value(expected_type=int, allow_none=True)
    b: int | None = Field.nested_value(expected_type=int, allow_none=True)


def test_default_element_names_follow_class_elements():
    p = _OrderProbe(a=1, b=2)
    assert p._element_names_for_serialize() == ("b", "a")


def test_series_attribute_mapping_uses_serializable_fields():
    from fastpyxl.chart.series import Series, attribute_mapping

    for kind, names in attribute_mapping.items():
        for name in names:
            assert name in Series.__fields__, f"{kind}: unknown field {name!r}"
            info = Series.__fields__[name]
            assert info.kind != "alias", f"{kind}: {name} is alias"
            assert info.serialize, f"{kind}: {name} does not serialize"


def test_bar_series_roundtrip_preserves_child_tag_order():
    """Regression: bar <ser> fragment order matches Excel schema expectations."""
    from fastpyxl.chart.series import Series, attribute_mapping
    from fastpyxl.tests.helper import compare_xml
    from fastpyxl.xml.functions import fromstring

    src = """
        <ser>
          <idx val="0"/>
          <order val="0"/>
          <spPr>
              <a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                <a:prstDash val="solid" />
              </a:ln>
            </spPr>
          <val>
            <numRef>
                <f>Blatt1!$A$1:$A$12</f>
            </numRef>
          </val>
        </ser>
        """
    node = fromstring(src)
    ser = Series.from_tree(node)
    ser._serialize_element_order = attribute_mapping["bar"]
    xml = tostring(ser.to_tree())
    diff = compare_xml(xml, src)
    assert diff is None, diff


def test_child_tag_order_matches_mapping_when_all_present():
    """Emit order follows attribute_mapping when each mapped field produces a node."""
    from fastpyxl.chart.data_source import NumDataSource, NumRef
    from fastpyxl.chart.series import Series

    ser = Series()
    ser._serialize_element_order = ("idx", "order", "val")
    ser.idx = 0
    ser.order = 0
    ser.val = NumDataSource(numRef=NumRef(f="'S'!$A$1:$A$4"))
    root = ser.to_tree()
    tags = [localname(ch) for ch in root]
    assert tags == ["idx", "order", "val"]


@pytest.mark.parametrize(
    "chart_mod, chart_name, expected_kind, use_xy_series",
    [
        ("fastpyxl.chart.bar_chart", "BarChart", "bar", False),
        ("fastpyxl.chart.line_chart", "LineChart", "line", False),
        ("fastpyxl.chart.area_chart", "AreaChart", "area", False),
        ("fastpyxl.chart.bubble_chart", "BubbleChart", "bubble", True),
        ("fastpyxl.chart.pie_chart", "PieChart", "pie", False),
        ("fastpyxl.chart.radar_chart", "RadarChart", "radar", False),
        ("fastpyxl.chart.scatter_chart", "ScatterChart", "scatter", True),
        ("fastpyxl.chart.surface_chart", "SurfaceChart", "surface", False),
        ("fastpyxl.chart.stock_chart", "StockChart", "line", False),
    ],
)
def test_chart_to_tree_sets_series_serialize_order(
    chart_mod, chart_name, expected_kind, use_xy_series
):
    import importlib

    from fastpyxl.chart.series import Series, XYSeries, attribute_mapping

    mod = importlib.import_module(chart_mod)
    chart_cls = getattr(mod, chart_name)
    ser = XYSeries() if use_xy_series else Series()
    chart = chart_cls()
    chart.ser = [ser]
    chart.to_tree()
    assert ser._serialize_element_order == attribute_mapping[expected_kind]
