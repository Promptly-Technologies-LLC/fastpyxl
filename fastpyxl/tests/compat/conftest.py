# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

from pathlib import Path

import pytest


@pytest.fixture
def fastpyxl():
    import fastpyxl

    return fastpyxl


@pytest.fixture
def openpyxl():
    import openpyxl

    return openpyxl


@pytest.fixture
def compat_tmp_path(tmp_path: Path) -> Path:
    return tmp_path


_TESTS_DIR = Path(__file__).resolve().parents[1]
_FASTPYXL_DIR = _TESTS_DIR.parent

READ_FIXTURES = (
    _FASTPYXL_DIR / "reader/tests/data/sample.xlsx",
    _FASTPYXL_DIR / "reader/tests/data/empty_with_no_properties.xlsx",
    _FASTPYXL_DIR / "reader/tests/data/hidden_sheets.xlsx",
    _TESTS_DIR / "data/genuine/sample.xlsx",
    _TESTS_DIR / "data/genuine/empty.xlsx",
)

GENUINE_SAMPLE_FIXTURE = _TESTS_DIR / "data/genuine/sample.xlsx"
