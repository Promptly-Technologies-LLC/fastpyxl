# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

import inspect

import pytest


pytestmark = pytest.mark.openpyxl_compat


PUBLIC_TOP_LEVEL_SYMBOLS = (
    "Workbook",
    "load_workbook",
    "open",
    "DEFUSEDXML",
    "LXML",
    "NUMPY",
    "__version__",
    "__author__",
    "__license__",
    "__url__",
)


PUBLIC_SUBMODULES = (
    "cell",
    "chart",
    "chartsheet",
    "comments",
    "drawing",
    "formatting",
    "formula",
    "styles",
    "utils",
    "workbook",
    "worksheet",
    "writer",
)


def test_public_top_level_symbols_exist(fastpyxl, openpyxl):
    for name in PUBLIC_TOP_LEVEL_SYMBOLS:
        assert hasattr(fastpyxl, name), f"fastpyxl missing public symbol {name!r}"
        assert hasattr(openpyxl, name), f"openpyxl missing public symbol {name!r}"


def test_public_submodules_exist(fastpyxl, openpyxl):
    for name in PUBLIC_SUBMODULES:
        assert hasattr(fastpyxl, name), f"fastpyxl missing submodule {name!r}"
        assert hasattr(openpyxl, name), f"openpyxl missing submodule {name!r}"


def test_load_workbook_signature_matches(fastpyxl, openpyxl):
    fast_sig = inspect.signature(fastpyxl.load_workbook)
    openpyxl_sig = inspect.signature(openpyxl.load_workbook)
    assert list(fast_sig.parameters) == list(openpyxl_sig.parameters)
    for name, fast_param in fast_sig.parameters.items():
        openpyxl_param = openpyxl_sig.parameters[name]
        assert fast_param.default == openpyxl_param.default, name
        assert fast_param.kind == openpyxl_param.kind, name


def test_workbook_constructor_signature_matches(fastpyxl, openpyxl):
    fast_sig = inspect.signature(fastpyxl.Workbook)
    openpyxl_sig = inspect.signature(openpyxl.Workbook)
    assert list(fast_sig.parameters) == list(openpyxl_sig.parameters)


def test_workbook_module_exports_match(fastpyxl, openpyxl):
    assert hasattr(fastpyxl.workbook, "Workbook")
    assert hasattr(openpyxl.workbook, "Workbook")


def test_styles_module_exports_match(fastpyxl, openpyxl):
    for name in ("Font", "PatternFill", "Alignment", "Border", "Side", "Protection"):
        assert hasattr(fastpyxl.styles, name), name
        assert hasattr(openpyxl.styles, name), name
