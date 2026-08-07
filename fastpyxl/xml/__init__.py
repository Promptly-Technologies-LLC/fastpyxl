# Copyright (c) 2010-2024 fastpyxl


"""Collection of XML resources compatible across different Python versions"""
import os
import warnings


def _env_flag(name, *, legacy_name=None, default="True"):
    value = os.environ.get(name)
    if value is None and legacy_name is not None:
        legacy_value = os.environ.get(legacy_name)
        if legacy_value is not None:
            warnings.warn(
                f"{legacy_name} is deprecated; use {name} instead. "
                "Legacy OPENPYXL_* environment aliases will be removed in a future release.",
                DeprecationWarning,
                stacklevel=2,
            )
            value = legacy_value
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
