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


READ_FIXTURES = (
    Path("fastpyxl/reader/tests/data/sample.xlsx"),
    Path("fastpyxl/reader/tests/data/empty_with_no_properties.xlsx"),
    Path("fastpyxl/reader/tests/data/hidden_sheets.xlsx"),
    Path("fastpyxl/tests/data/genuine/sample.xlsx"),
    Path("fastpyxl/tests/data/genuine/empty.xlsx"),
)
