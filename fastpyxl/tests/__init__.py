# Copyright (c) 2010-2024 fastpyxl

import os


def _env_flag(name, *, legacy_name=None, default="False"):
    value = os.environ.get(name)
    if value is None and legacy_name is not None:
        value = os.environ.get(legacy_name)
    if value is None:
        value = default
    return value == "True"


KEEP_VBA = _env_flag("FASTPYXL_KEEP_VBA", legacy_name="OPENPYXL_KEEP_VBA")
