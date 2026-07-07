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


_TESTS_DIR = Path(__file__).resolve().parents[1]
_FASTPYXL_DIR = _TESTS_DIR.parent


def _fixture_id(path: Path) -> str:
    for base, prefix in (
        (_FASTPYXL_DIR / "reader/tests/data", "reader"),
        (_TESTS_DIR / "data/genuine", "genuine"),
    ):
        try:
            return f"{prefix}/{path.relative_to(base)}"
        except ValueError:
            continue
    return path.name


READ_FIXTURES = (
    (_fixture_id(_FASTPYXL_DIR / "reader/tests/data/sample.xlsx"), _FASTPYXL_DIR / "reader/tests/data/sample.xlsx"),
    (
        _fixture_id(_FASTPYXL_DIR / "reader/tests/data/empty_with_no_properties.xlsx"),
        _FASTPYXL_DIR / "reader/tests/data/empty_with_no_properties.xlsx",
    ),
    (
        _fixture_id(_FASTPYXL_DIR / "reader/tests/data/hidden_sheets.xlsx"),
        _FASTPYXL_DIR / "reader/tests/data/hidden_sheets.xlsx",
    ),
    (_fixture_id(_TESTS_DIR / "data/genuine/sample.xlsx"), _TESTS_DIR / "data/genuine/sample.xlsx"),
    (_fixture_id(_TESTS_DIR / "data/genuine/empty.xlsx"), _TESTS_DIR / "data/genuine/empty.xlsx"),
    (
        _fixture_id(_TESTS_DIR / "data/genuine/empty-with-styles.xlsx"),
        _TESTS_DIR / "data/genuine/empty-with-styles.xlsx",
    ),
)

READ_FIXTURE_IDS = tuple(fixture_id for fixture_id, _ in READ_FIXTURES)
READ_FIXTURE_PATHS = tuple(fixture_path for _, fixture_path in READ_FIXTURES)

GENUINE_SAMPLE_FIXTURE = _TESTS_DIR / "data/genuine/sample.xlsx"
