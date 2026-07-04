from __future__ import annotations

import re

from .errors import FieldValidationError

_HEX_BINARY = re.compile(r"[0-9a-fA-F]+$")


def hex_binary_converter(value):
    if value is None:
        return None
    if not _HEX_BINARY.match(value):
        raise FieldValidationError(f"hex binary rejected value {value!r}")
    return value
