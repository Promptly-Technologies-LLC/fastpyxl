import fastpyxl.drawing.text as drawing_text

from fastpyxl.drawing.colors import ColorChoice
from fastpyxl.drawing.text import (
    CharacterProperties,
    Hyperlink,
    LineBreak,
    ListStyle,
    Paragraph,
    ParagraphProperties,
    PresetTextShape,
    RegularTextRun,
    RichTextProperties,
    TextField,
)
from fastpyxl.tests.helper import compare_xml
from fastpyxl.typed_serialisable.base import Serialisable as TypedSerialisable
from fastpyxl.xml.functions import fromstring, tostring


def test_phase5_drawing_text_core_models_use_typed_serialisable_base():
    assert issubclass(RegularTextRun, TypedSerialisable)
    assert issubclass(LineBreak, TypedSerialisable)
    assert issubclass(TextField, TypedSerialisable)
    assert issubclass(Paragraph, TypedSerialisable)
    assert issubclass(CharacterProperties, TypedSerialisable)
    assert issubclass(ParagraphProperties, TypedSerialisable)
    assert issubclass(RichTextProperties, TypedSerialisable)


def test_phase5a_drawing_text_defines_only_typed_serialisable_models():
    for name in dir(drawing_text):
        obj = getattr(drawing_text, name)
        if not isinstance(obj, type):
            continue
        if getattr(obj, "__module__", None) != drawing_text.__name__:
            continue
        if name.startswith("_"):
            continue
        assert issubclass(obj, TypedSerialisable), (
            f"drawing.text.{name} must subclass typed_serialisable.base.Serialisable"
        )


def test_character_properties_coerces_string_to_color_choice_and_serializes():
    model = CharacterProperties(b=True, solidFill="FF00AA")
    assert isinstance(model.solidFill, ColorChoice)
    xml = tostring(model.to_tree())
    expected = """
    <defRPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" b="1">
      <solidFill>
        <srgbClr val="FF00AA"></srgbClr>
      </solidFill>
    </defRPr>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_paragraph_properties_bu_char_value_attribute_roundtrip():
    xml = """
    <pPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
      <buChar char="*"/>
    </pPr>
    """
    model = ParagraphProperties.from_tree(fromstring(xml))
    assert model.buChar == "*"
    out = tostring(model.to_tree())
    diff = compare_xml(out, xml)
    assert diff is None, diff


def test_richtext_properties_attributes_and_empty_autofit_element():
    model = RichTextProperties(rot=5400000, noAutofit=True, wrap="square")
    xml = tostring(model.to_tree())
    expected = """
    <bodyPr xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" rot="5400000" wrap="square">
      <noAutofit></noAutofit>
    </bodyPr>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_regular_text_run_alias_and_nested_text_serialization():
    run = RegularTextRun(t="hello")
    run.value = "world"
    assert run.t == "world"
    xml = tostring(run.to_tree())
    expected = """
    <r xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
      <t>world</t>
    </r>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_paragraph_sequence_and_alias_preserve_default_and_order():
    paragraph = Paragraph(r=[RegularTextRun(t="a"), RegularTextRun(t="b")])
    paragraph.text = [RegularTextRun(t="x")]
    assert len(paragraph.r) == 1
    xml = tostring(paragraph.to_tree())
    expected = """
    <p xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
      <r>
        <t>x</t>
      </r>
    </p>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_text_field_attributes_and_child_elements_render():
    model = TextField(id="f1", type="slidenum", t="display")
    xml = tostring(model.to_tree("fld"))
    expected = """
    <fld id="f1" type="slidenum" t="display"></fld>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_hyperlink_relationship_id_roundtrip():
    ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xml = f'<hlinkClick xmlns="{ns}" xmlns:r="{rel}" r:id="rId2"/>'
    model = Hyperlink.from_tree(fromstring(xml))
    assert model.id == "rId2"
    out = tostring(model.to_tree())
    diff = compare_xml(out, xml)
    assert diff is None, diff


def test_preset_text_shape_prst_and_guide_list_roundtrip():
    ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    xml = f"""
    <prstTxWarp xmlns="{ns}">
      <prst val="textPlain"/>
      <avLst>
        <gd name="adj" fmla="val 1"/>
      </avLst>
    </prstTxWarp>
    """
    model = PresetTextShape.from_tree(fromstring(xml))
    assert model.prst == "textPlain"
    assert model.avLst is not None and len(model.avLst.gd) == 1
    assert model.avLst.gd[0].name == "adj"
    out = tostring(model.to_tree("prstTxWarp"))
    diff = compare_xml(out, xml)
    assert diff is None, diff


def test_list_style_serializes_level_placeholders():
    model = ListStyle(lvl1pPr=ParagraphProperties(lvl=1))
    xml = tostring(model.to_tree())
    expected = """
    <lstStyle xmlns="http://schemas.openxmlformats.org/drawingml/2006/main">
      <lvl1pPr lvl="1"></lvl1pPr>
    </lstStyle>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
