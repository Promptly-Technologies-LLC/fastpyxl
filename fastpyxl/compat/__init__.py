# Copyright (c) 2010-2024 fastpyxl

import sys

from .numbers import NUMERIC_TYPES
from .strings import safe_string

if sys.version_info >= (3, 13):
    from warnings import deprecated
else:
    from typing_extensions import deprecated

__all__ = ["NUMERIC_TYPES", "deprecated", "safe_string"]
