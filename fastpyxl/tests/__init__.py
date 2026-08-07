# Copyright (c) 2010-2024 fastpyxl

import os
import warnings


def _env_flag(name, *, legacy_name=None, default="False"):
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


KEEP_VBA = _env_flag("FASTPYXL_KEEP_VBA", legacy_name="OPENPYXL_KEEP_VBA")
