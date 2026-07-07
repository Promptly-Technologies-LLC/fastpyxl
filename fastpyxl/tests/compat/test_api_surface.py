# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

import importlib
import inspect

import pytest


pytestmark = pytest.mark.openpyxl_compat


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

# openpyxl symbols fastpyxl intentionally does not expose at the top level.
OPENPYXL_TOP_LEVEL_ALLOWLIST = frozenset({"DEBUG", "descriptors"})

# Per-submodule openpyxl symbols fastpyxl intentionally does not expose.
OPENPYXL_SUBMODULE_ALLOWLIST: dict[str, frozenset[str]] = {
    "utils": frozenset({"bound_dictionary"}),
}


def _public_names(module: object) -> set[str]:
    return {name for name in dir(module) if not name.startswith("_")}


def test_openpyxl_top_level_exports_available_in_fastpyxl(fastpyxl, openpyxl):
    missing = _public_names(openpyxl) - _public_names(fastpyxl) - OPENPYXL_TOP_LEVEL_ALLOWLIST
    assert not missing, f"fastpyxl missing openpyxl top-level exports: {sorted(missing)}"


@pytest.mark.parametrize("submodule", PUBLIC_SUBMODULES)
def test_openpyxl_submodule_exports_available_in_fastpyxl(
    submodule: str,
    fastpyxl,
    openpyxl,
):
    fast_module = importlib.import_module(f"{fastpyxl.__name__}.{submodule}")
    openpyxl_module = importlib.import_module(f"{openpyxl.__name__}.{submodule}")
    allowlist = OPENPYXL_SUBMODULE_ALLOWLIST.get(submodule, frozenset())
    missing = _public_names(openpyxl_module) - _public_names(fast_module) - allowlist
    assert not missing, (
        f"fastpyxl.{submodule} missing openpyxl.{submodule} exports: {sorted(missing)}"
    )


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
    for name, fast_param in fast_sig.parameters.items():
        openpyxl_param = openpyxl_sig.parameters[name]
        assert fast_param.default == openpyxl_param.default, name
        assert fast_param.kind == openpyxl_param.kind, name
