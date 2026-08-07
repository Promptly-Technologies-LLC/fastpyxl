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


def _reader_fixture(name: str) -> tuple[str, Path]:
    path = _FASTPYXL_DIR / "reader/tests/data" / name
    return _fixture_id(path), path


def _genuine_fixture(name: str) -> tuple[str, Path]:
    path = _TESTS_DIR / "data/genuine" / name
    return _fixture_id(path), path


READ_FIXTURES = (
    _reader_fixture("sample.xlsx"),
    _reader_fixture("empty_with_no_properties.xlsx"),
    _reader_fixture("hidden_sheets.xlsx"),
    _reader_fixture("complex-styles.xlsx"),
    _reader_fixture("pivot.xlsx"),
    _reader_fixture("legacy_drawing.xlsm"),
    _reader_fixture("contains_chartsheets.xlsx"),
    _reader_fixture("sample_with_images.xlsx"),
    _genuine_fixture("sample.xlsx"),
    _genuine_fixture("empty.xlsx"),
    _genuine_fixture("empty-with-styles.xlsx"),
    _genuine_fixture("libreoffice_nrt.xlsx"),
    _genuine_fixture("mac_date.xlsx"),
)

READ_FIXTURE_IDS = tuple(fixture_id for fixture_id, _ in READ_FIXTURES)
READ_FIXTURE_PATHS = tuple(fixture_path for _, fixture_path in READ_FIXTURES)

GENUINE_SAMPLE_FIXTURE = _TESTS_DIR / "data/genuine/sample.xlsx"
LEGACY_DRAWING_XLSM = _FASTPYXL_DIR / "reader/tests/data/legacy_drawing.xlsm"
EXTERNAL_LINKS_FIXTURE = (
    _FASTPYXL_DIR / "workbook/external_link/tests/data/book1.xlsx"
)
