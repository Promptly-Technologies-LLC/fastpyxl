# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from pathlib import Path

import pytest

from .conftest import READ_FIXTURES
from .helpers import assert_workbooks_match, load_fixture_both


pytestmark = pytest.mark.openpyxl_compat


@pytest.mark.parametrize("fixture_path", READ_FIXTURES, ids=lambda p: p.name)
def test_both_libraries_read_fixture_equivalently(fixture_path: Path):
    fast_wb, openpyxl_wb = load_fixture_both(fixture_path)
    try:
        assert_workbooks_match(fast_wb, openpyxl_wb)
    finally:
        if hasattr(fast_wb, "close"):
            fast_wb.close()
        if hasattr(openpyxl_wb, "close"):
            openpyxl_wb.close()


@pytest.mark.parametrize("fixture_path", READ_FIXTURES, ids=lambda p: p.name)
def test_fastpyxl_save_fixture_openpyxl_can_read(
    fixture_path: Path,
    compat_tmp_path,
    fastpyxl,
    openpyxl,
):
    fast_wb, _ = load_fixture_both(fixture_path)
    try:
        out = compat_tmp_path / f"resaved-{fixture_path.name}"
        fast_wb.save(out)
        openpyxl_reloaded = openpyxl.load_workbook(out)
        fast_reloaded = fastpyxl.load_workbook(out)
        assert_workbooks_match(fast_reloaded, openpyxl_reloaded)
    finally:
        if hasattr(fast_wb, "close"):
            fast_wb.close()


@pytest.mark.parametrize("fixture_path", READ_FIXTURES, ids=lambda p: p.name)
def test_openpyxl_save_fixture_fastpyxl_can_read(
    fixture_path: Path,
    compat_tmp_path,
    fastpyxl,
    openpyxl,
):
    _, openpyxl_wb = load_fixture_both(fixture_path)
    try:
        out = compat_tmp_path / f"resaved-{fixture_path.name}"
        openpyxl_wb.save(out)
        fast_reloaded = fastpyxl.load_workbook(out)
        openpyxl_reloaded = openpyxl.load_workbook(out)
        assert_workbooks_match(fast_reloaded, openpyxl_reloaded)
    finally:
        if hasattr(openpyxl_wb, "close"):
            openpyxl_wb.close()
