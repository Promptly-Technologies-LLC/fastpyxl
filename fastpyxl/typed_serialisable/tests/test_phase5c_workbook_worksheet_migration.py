import importlib
import pkgutil

import fastpyxl.comments as comments_pkg
import fastpyxl.worksheet as worksheet_pkg
from fastpyxl.packaging.workbook import WorkbookPackage
from fastpyxl.tests.helper import compare_xml
from fastpyxl.typed_serialisable.base import Serialisable as TypedSerialisable
from fastpyxl.xml.functions import fromstring, tostring


def _iter_module_types(pkg, *, skip=frozenset()):
    for mi in pkgutil.iter_modules(pkg.__path__):
        if mi.name in skip or mi.ispkg:
            continue
        mod = importlib.import_module(f"{pkg.__name__}.{mi.name}")
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
            yield obj


def test_phase5c_worksheet_models_subclass_typed_serialisable():
    skip = frozenset({"_read_only", "_reader", "_writer", "copier", "formula", "worksheet"})
    for cls in _iter_module_types(worksheet_pkg, skip=skip):
        assert issubclass(cls, TypedSerialisable)


def test_phase5c_comments_models_subclass_typed_serialisable():
    skip = frozenset({"comments", "shape_writer", "comment_sheet", "author"})
    for cls in _iter_module_types(comments_pkg, skip=skip):
        assert issubclass(cls, TypedSerialisable)


def test_phase5c_cell_text_roundtrip():
    from fastpyxl.cell.text import Text

    src = "<text><t>hello</t></text>"
    model = Text.from_tree(fromstring(src))
    out = tostring(model.to_tree())
    diff = compare_xml(out, src)
    assert diff is None, diff


def test_phase5c_workbook_package_roundtrip_sheet_id_namespace():
    src = """
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <workbookPr/>
      <bookViews><workbookView/></bookViews>
      <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
    </workbook>
    """
    model = WorkbookPackage.from_tree(fromstring(src))
    out = tostring(model.to_tree())
    back = WorkbookPackage.from_tree(fromstring(out))
    assert back.sheets[0].id == "rId1"


def test_phase5c_stylesheet_roundtrip_minimal():
    from fastpyxl.styles.stylesheet import Stylesheet

    src = """
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <numFmts count="0"/>
      <cellStyleXfs count="0"/>
      <cellXfs count="0"/>
      <cellStyles count="0"/>
    </styleSheet>
    """
    model = Stylesheet.from_tree(fromstring(src))
    out = tostring(model.to_tree())
    diff = compare_xml(out, src)
    assert diff is None, diff


def test_phase5c_formatting_rule_roundtrip():
    from fastpyxl.formatting.rule import Rule

    src = '<cfRule type="expression" priority="1"><formula>A1&gt;0</formula></cfRule>'
    model = Rule.from_tree(fromstring(src))
    out = tostring(model.to_tree())
    back = Rule.from_tree(fromstring(out))
    assert back.type == "expression"
    assert back.formula == ["A1>0"]
