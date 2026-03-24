from fastpyxl.styles.cell_style import CellStyle, CellStyleList
from fastpyxl.styles.differential import DifferentialStyle, DifferentialStyleList
from fastpyxl.styles.borders import Border, Side
from fastpyxl.styles.colors import Color, ColorList, RgbColor
from fastpyxl.styles.fills import GradientFill, PatternFill, Stop
from fastpyxl.styles.table import TableStyle, TableStyleElement, TableStyleList
from fastpyxl.styles.alignment import Alignment
from fastpyxl.styles.numbers import NumberFormat, NumberFormatList
from fastpyxl.styles.protection import Protection
from fastpyxl.tests.helper import compare_xml
from fastpyxl.typed_serialisable.base import Serialisable as TypedSerialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.xml.functions import fromstring, tostring


def test_phase4_styles_core_models_use_typed_serialisable_base():
    assert issubclass(Alignment, TypedSerialisable)
    assert issubclass(Protection, TypedSerialisable)
    assert issubclass(NumberFormat, TypedSerialisable)
    assert issubclass(NumberFormatList, TypedSerialisable)
    assert issubclass(Color, TypedSerialisable)
    assert issubclass(ColorList, TypedSerialisable)
    assert issubclass(PatternFill, TypedSerialisable)
    assert issubclass(GradientFill, TypedSerialisable)
    assert issubclass(Side, TypedSerialisable)
    assert issubclass(Border, TypedSerialisable)
    assert issubclass(TableStyleElement, TypedSerialisable)
    assert issubclass(TableStyle, TypedSerialisable)
    assert issubclass(TableStyleList, TypedSerialisable)
    assert issubclass(DifferentialStyle, TypedSerialisable)
    assert issubclass(DifferentialStyleList, TypedSerialisable)
    assert issubclass(CellStyle, TypedSerialisable)
    assert issubclass(CellStyleList, TypedSerialisable)


def test_alignment_serialization_preserves_alias_and_zero_omission_behavior():
    model = Alignment(text_rotation=45, wrap_text=True, shrink_to_fit=False, indent=0)
    xml = tostring(model.to_tree())
    expected = """
    <alignment textRotation="45" wrapText="1" shrinkToFit="0"></alignment>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_alignment_text_rotation_rejects_out_of_range_value():
    try:
        Alignment(textRotation=500)
    except FieldValidationError:
        return
    assert False, "Expected FieldValidationError for invalid textRotation"


def test_number_format_list_renders_count_attribute():
    model = NumberFormatList(numFmt=[NumberFormat(numFmtId=164, formatCode="0.00")])
    xml = tostring(model.to_tree("numFmts"))
    expected = """
    <numFmts count="1">
      <numFmt numFmtId="164" formatCode="0.00"></numFmt>
    </numFmts>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_number_format_list_roundtrip_preserves_items():
    src = "<numFmts count='1'><numFmt numFmtId='164' formatCode='0.00'/></numFmts>"
    model = NumberFormatList.from_tree(fromstring(src))
    assert model.count == 1
    assert model.numFmt[0].numFmtId == 164
    assert model.numFmt[0].formatCode == "0.00"


def test_color_accepts_rgb_hex_and_emits_argb():
    model = Color(rgb="FF0000")
    xml = tostring(model.to_tree())
    expected = """<color rgb="00FF0000"></color>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_color_list_nested_sequence_roundtrip():
    model = ColorList(indexedColors=[RgbColor(rgb="000000")], mruColors=[Color(rgb="FFFFFF")])
    xml = tostring(model.to_tree())
    expected = """
    <colors>
      <indexedColors><rgbColor rgb="00000000"></rgbColor></indexedColors>
      <mruColors><color rgb="00FFFFFF"></color></mruColors>
    </colors>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_pattern_fill_supports_string_color_assignment():
    model = PatternFill(fill_type="solid", start_color="FF0000")
    xml = tostring(model.to_tree())
    expected = """
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="00FF0000"></fgColor>
      </patternFill>
    </fill>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_gradient_fill_assigns_stop_positions_for_color_list():
    fill = GradientFill(stop=[Color(rgb="000000"), Color(rgb="FFFFFF")])
    assert isinstance(fill.stop[0], Stop)
    assert fill.stop[0].position == 0.0
    assert fill.stop[1].position == 1.0


def test_border_side_alias_and_outline_rendering():
    model = Border(left=Side(border_style="thin", color="000000"), outline=False)
    xml = tostring(model.to_tree())
    expected = """
    <border outline="0">
      <left style="thin"><color rgb="00000000"></color></left>
    </border>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_table_style_list_serialization_includes_computed_count():
    model = TableStyleList(
        tableStyle=[
            TableStyle(
                name="MyStyle",
                pivot=True,
                table=True,
                tableStyleElement=[TableStyleElement(type="wholeTable", dxfId=1)],
            )
        ]
    )
    xml = tostring(model.to_tree())
    expected = """
    <tableStyles count="1" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16">
      <tableStyle name="MyStyle" pivot="1" table="1">
        <tableStyleElement type="wholeTable" dxfId="1"></tableStyleElement>
      </tableStyle>
    </tableStyles>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_differential_style_list_alias_and_count_behavior():
    dxf = DifferentialStyle(border=Border(left=Side(style="thin")))
    dxfs = DifferentialStyleList()
    dxfs.append(dxf)
    assert dxfs.count == 1
    assert dxfs.styles[0] == dxf
    xml = tostring(dxfs.to_tree())
    expected = """
    <dxfs count="1">
      <dxf>
        <border>
          <left style="thin"></left>
        </border>
      </dxf>
    </dxfs>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_cell_style_serializes_apply_flags_from_child_presence():
    model = CellStyle(
        numFmtId=1,
        fontId=2,
        fillId=3,
        borderId=4,
        alignment=Alignment(horizontal="left"),
    )
    xml = tostring(model.to_tree())
    expected = """
    <xf numFmtId="1" fontId="2" fillId="3" borderId="4" applyAlignment="1">
      <alignment horizontal="left"></alignment>
    </xf>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_cell_style_list_serializes_computed_count():
    model = CellStyleList(xf=[CellStyle(), CellStyle()])
    xml = tostring(model.to_tree())
    expected = """
    <cellXfs count="2">
      <xf numFmtId="0" fontId="0" fillId="0" borderId="0"></xf>
      <xf numFmtId="0" fontId="0" fillId="0" borderId="0"></xf>
    </cellXfs>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff

