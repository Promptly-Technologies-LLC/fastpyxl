# Copyright (c) 2010-2024 fastpyxl

from typing import Any

ABC: Any
try:
    from abc import ABC as _stdlib_abc
except ImportError:
    from abc import ABCMeta
    ABC = ABCMeta("ABC", (object,), {})
else:
    ABC = _stdlib_abc
