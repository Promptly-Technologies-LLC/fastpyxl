# Copyright (c) 2010-2024 fastpyxl


"""Collection of XML resources compatible across different Python versions"""
import os


def _env_flag(name, *, legacy_name=None, default="True"):
    value = os.environ.get(name)
    if value is None and legacy_name is not None:
        value = os.environ.get(legacy_name)
    if value is None:
        value = default
    return value == "True"


def lxml_available():
    try:
        from lxml.etree import LXML_VERSION
        LXML = LXML_VERSION >= (3, 3, 1, 0)
        if not LXML:
            import warnings
            warnings.warn("The installed version of lxml is too old to be used with fastpyxl")
            return False  # we have it, but too old
        else:
            return True  # we have it, and recent enough
    except ImportError:
        return False  # we don't even have it


def lxml_env_set():
    return _env_flag("FASTPYXL_LXML", legacy_name="OPENPYXL_LXML")


LXML = lxml_available() and lxml_env_set()


def defusedxml_available():
    try:
        import defusedxml # noqa
    except ImportError:
        return False
    else:
        return True


def defusedxml_env_set():
    return _env_flag("FASTPYXL_DEFUSEDXML", legacy_name="OPENPYXL_DEFUSEDXML")


DEFUSEDXML = defusedxml_available() and defusedxml_env_set()
