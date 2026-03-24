from fastpyxl.styles.colors import Color
from fastpyxl.styles.fonts import DEFAULT_FONT, Font as FontModel
from fastpyxl.tests.helper import compare_xml
from fastpyxl.typed_serialisable.base import Serialisable as TypedSerialisable
from fastpyxl.xml.constants import SHEET_MAIN_NS
from fastpyxl.xml.functions import fromstring, tostring


def test_font_migrated_to_typed_serialisable_base():
    # Red test: until `fastpyxl.styles.fonts.Font` is migrated, it subclasses the legacy
    # descriptor-based Serialisable and will not be a typed runtime model.
    assert issubclass(FontModel, TypedSerialisable)


def test_font_xml_stability_for_default_font():
    xml = tostring(DEFAULT_FONT.to_tree())
    expected = """
    <font>
      <name val="Calibri"></name>
      <family val="2"></family>
      <color theme="1"></color>
      <sz val="11"></sz>
      <scheme val="minor"></scheme>
    </font>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_font_xml_stability_for_custom_renderer_and_underline():
    f = FontModel(
        name="Calibri",
        sz=11,
        family=2,
        b=True,
        i=True,
        u="double",
        color=Color(theme=1),
        scheme="minor",
    )
    xml = tostring(f.to_tree())
    expected = """
    <font>
      <name val="Calibri"></name>
      <family val="2"></family>
      <b val="1"></b>
      <i val="1"></i>
      <color theme="1"></color>
      <sz val="11"></sz>
      <u val="double"></u>
      <scheme val="minor"></scheme>
    </font>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_font_from_tree_sets_underline_default_when_val_missing():
    src = f'<font xmlns="{SHEET_MAIN_NS}"><u/></font>'
    obj = FontModel.from_tree(fromstring(src))
    assert obj.u == "single"

