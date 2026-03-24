from fastpyxl.drawing.colors import ColorChoice, HSLColor, RGBPercent
from fastpyxl.drawing.effect import EffectList, GlowEffect, OuterShadow, ReflectionEffect
from fastpyxl.drawing.geometry import GroupTransform2D, Point2D, PositiveSize2D, Transform2D
from fastpyxl.drawing.graphic import GraphicData, GraphicFrame, GraphicObject, GroupShape, NonVisualGraphicFrame
from fastpyxl.drawing.line import DashStop, DashStopList, LineEndProperties, LineProperties
from fastpyxl.drawing.picture import PictureFrame
from fastpyxl.drawing.relation import ChartRelation
from fastpyxl.drawing.connector import ConnectorShape, Shape
from fastpyxl.drawing.fill import Blip, BlipFillProperties, GradientFillProperties, GradientStop, SolidColorFillProperties
from fastpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor, SpreadsheetDrawing
from fastpyxl.drawing.properties import GroupShapeProperties, NonVisualDrawingProps, NonVisualGroupShape
from fastpyxl.tests.helper import compare_xml
from fastpyxl.typed_serialisable.base import Serialisable as TypedSerialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.xml.functions import tostring


def test_phase4_drawing_primitives_use_typed_serialisable_base():
    assert issubclass(Point2D, TypedSerialisable)
    assert issubclass(PositiveSize2D, TypedSerialisable)
    assert issubclass(Transform2D, TypedSerialisable)
    assert issubclass(GroupTransform2D, TypedSerialisable)
    assert issubclass(HSLColor, TypedSerialisable)
    assert issubclass(RGBPercent, TypedSerialisable)
    assert issubclass(ColorChoice, TypedSerialisable)
    assert issubclass(GlowEffect, TypedSerialisable)
    assert issubclass(OuterShadow, TypedSerialisable)
    assert issubclass(ReflectionEffect, TypedSerialisable)
    assert issubclass(EffectList, TypedSerialisable)
    assert issubclass(LineEndProperties, TypedSerialisable)
    assert issubclass(DashStop, TypedSerialisable)
    assert issubclass(DashStopList, TypedSerialisable)
    assert issubclass(LineProperties, TypedSerialisable)
    assert issubclass(ChartRelation, TypedSerialisable)
    assert issubclass(GroupShapeProperties, TypedSerialisable)
    assert issubclass(NonVisualDrawingProps, TypedSerialisable)
    assert issubclass(NonVisualGroupShape, TypedSerialisable)
    assert issubclass(GraphicData, TypedSerialisable)
    assert issubclass(GraphicObject, TypedSerialisable)
    assert issubclass(GraphicFrame, TypedSerialisable)
    assert issubclass(NonVisualGraphicFrame, TypedSerialisable)
    assert issubclass(GroupShape, TypedSerialisable)
    assert issubclass(PictureFrame, TypedSerialisable)
    assert issubclass(ConnectorShape, TypedSerialisable)
    assert issubclass(Shape, TypedSerialisable)
    assert issubclass(AnchorMarker, TypedSerialisable)
    assert issubclass(OneCellAnchor, TypedSerialisable)
    assert issubclass(SpreadsheetDrawing, TypedSerialisable)
    assert issubclass(GradientStop, TypedSerialisable)
    assert issubclass(GradientFillProperties, TypedSerialisable)
    assert issubclass(SolidColorFillProperties, TypedSerialisable)
    assert issubclass(Blip, TypedSerialisable)
    assert issubclass(BlipFillProperties, TypedSerialisable)


def test_transform2d_serialization_order_stability():
    model = Transform2D(
        rot=5400000,
        flipH=True,
        off=Point2D(x=1, y=2),
        ext=PositiveSize2D(cx=3, cy=4),
    )
    xml = tostring(model.to_tree())
    expected = """
    <xfrm xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" rot="5400000" flipH="1">
      <off x="1" y="2"></off>
      <ext cx="3" cy="4"></ext>
    </xfrm>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_positive_size_aliases_roundtrip_values():
    model = PositiveSize2D()
    model.width = 7
    model.height = 8
    assert model.cx == 7
    assert model.cy == 8


def test_hsl_color_rejects_out_of_range_values():
    try:
        HSLColor(hue=10, sat=150, lum=50)
    except FieldValidationError:
        return
    assert False, "Expected FieldValidationError for sat outside [0, 100]"


def test_color_choice_alias_and_nested_value_rendering():
    model = ColorChoice(srgbClr="FF00AA")
    xml = tostring(model.to_tree("colorChoice"))
    expected = """
    <colorChoice xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
      <srgbClr val="FF00AA"></srgbClr>
    </colorChoice>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_effect_list_serialization_order_and_colorchoice_integration():
    model = EffectList(
        glow=GlowEffect(rad=1000, srgbClr="FF0000"),
        outerShdw=OuterShadow(blurRad=10, algn="ctr", srgbClr="00FF00"),
    )
    xml = tostring(model.to_tree("effectLst"))
    expected = """
    <effectLst>
      <glow rad="1000" xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
        <srgbClr val="FF0000"></srgbClr>
      </glow>
      <outerShdw blurRad="10" algn="ctr" xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
        <srgbClr val="00FF00"></srgbClr>
      </outerShdw>
    </effectLst>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_effect_enum_validation_rejects_invalid_alignment():
    try:
        OuterShadow(algn="middle")
    except FieldValidationError:
        return
    assert False, "Expected FieldValidationError for invalid algn"


def test_line_end_and_dash_stop_serialization():
    end = LineEndProperties(type="triangle", w="lg", len="sm")
    end_xml = tostring(end.to_tree())
    end_expected = """
    <end xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" type="triangle" w="lg" len="sm"></end>
    """
    diff = compare_xml(end_xml, end_expected)
    assert diff is None, diff

    second = DashStop(d=2000, sp=1000)
    second.length = 2000
    second.space = 1000
    dash = DashStopList(ds=[DashStop(d=1000, sp=500), second])
    dash_xml = tostring(dash.to_tree("custDash"))
    dash_expected = """
    <custDash xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
      <ds d="1000" sp="500"></ds>
      <ds d="2000" sp="1000"></ds>
    </custDash>
    """
    diff = compare_xml(dash_xml, dash_expected)
    assert diff is None, diff

    line = LineProperties(w=12700, cap="rnd", prstDash="solid", headEnd=LineEndProperties(type="triangle"))
    line_xml = tostring(line.to_tree())
    line_expected = """
    <ln xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" w="12700" cap="rnd">
      <prstDash val="solid"></prstDash>
      <headEnd type="triangle"></headEnd>
    </ln>
    """
    diff = compare_xml(line_xml, line_expected)
    assert diff is None, diff


def test_chart_relation_serialization():
    rel = ChartRelation(id="rId1")
    xml = tostring(rel.to_tree())
    expected = """
    <chart xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"></chart>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_non_visual_group_shape_serialization():
    model = NonVisualGroupShape(
        cNvPr=NonVisualDrawingProps(id=1, name="grp"),
    )
    xml = tostring(model.to_tree())
    expected = """
    <nvGrpSpPr>
      <cNvPr id="1" name="grp"></cNvPr>
    </nvGrpSpPr>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_graphic_data_and_frame_serialization():
    model = GraphicFrame(
        graphic=GraphicObject(graphicData=GraphicData(uri="urn:test", chart=ChartRelation(id="rId1"))),
    )
    xml = tostring(model.to_tree())
    expected = """
    <graphicFrame>
      <nvGraphicFramePr>
        <cNvPr id="0" name="Chart 0"></cNvPr>
        <cNvGraphicFramePr></cNvGraphicFramePr>
      </nvGraphicFramePr>
      <xfrm></xfrm>
      <graphic xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
        <graphicData uri="urn:test">
          <chart xmlns="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"></chart>
        </graphicData>
      </graphic>
    </graphicFrame>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_spreadsheet_anchor_serialization():
    drawing = SpreadsheetDrawing(
        oneCellAnchor=[OneCellAnchor(_from=AnchorMarker(col=1, colOff=2, row=3, rowOff=4))]
    )
    xml = tostring(drawing.to_tree())
    expected = """
    <wsDr>
      <oneCellAnchor>
        <from>
          <col>1</col>
          <colOff>2</colOff>
          <row>3</row>
          <rowOff>4</rowOff>
        </from>
        <ext cx="0" cy="0"></ext>
        <clientData></clientData>
      </oneCellAnchor>
    </wsDr>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_fill_models_serialization():
    grad = GradientFillProperties(gsLst=[GradientStop(pos=25000, srgbClr="FF0000")])
    grad_xml = tostring(grad.to_tree())
    grad_expected = """
    <gradFill xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
      <gsLst>
        <gs pos="25000">
          <srgbClr val="FF0000"></srgbClr>
        </gs>
      </gsLst>
    </gradFill>
    """
    diff = compare_xml(grad_xml, grad_expected)
    assert diff is None, diff

    solid = SolidColorFillProperties(srgbClr="00FF00")
    solid_xml = tostring(solid.to_tree())
    solid_expected = """
    <solidFill>
      <srgbClr val="00FF00"></srgbClr>
    </solidFill>
    """
    diff = compare_xml(solid_xml, solid_expected)
    assert diff is None, diff
