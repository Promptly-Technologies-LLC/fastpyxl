# Copyright (c) 2010-2024 fastpyxl

from __future__ import annotations

import warnings

import pytest

from fastpyxl.xml import _env_flag


def test_env_flag_uses_primary_name(monkeypatch):
    monkeypatch.setenv("FASTPYXL_LXML", "False")
    monkeypatch.delenv("OPENPYXL_LXML", raising=False)
    assert _env_flag("FASTPYXL_LXML", legacy_name="OPENPYXL_LXML") is False


def test_env_flag_falls_back_to_legacy_name_with_warning(monkeypatch):
    monkeypatch.delenv("FASTPYXL_LXML", raising=False)
    monkeypatch.setenv("OPENPYXL_LXML", "False")
    with pytest.warns(DeprecationWarning, match="OPENPYXL_LXML is deprecated"):
        assert _env_flag("FASTPYXL_LXML", legacy_name="OPENPYXL_LXML") is False


def test_env_flag_prefers_primary_over_legacy_without_warning(monkeypatch):
    monkeypatch.setenv("FASTPYXL_LXML", "True")
    monkeypatch.setenv("OPENPYXL_LXML", "False")
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        assert _env_flag("FASTPYXL_LXML", legacy_name="OPENPYXL_LXML") is True
    assert caught == []


def test_env_flag_default_when_unset(monkeypatch):
    monkeypatch.delenv("FASTPYXL_LXML", raising=False)
    monkeypatch.delenv("OPENPYXL_LXML", raising=False)
    assert _env_flag("FASTPYXL_LXML", legacy_name="OPENPYXL_LXML", default="True") is True
