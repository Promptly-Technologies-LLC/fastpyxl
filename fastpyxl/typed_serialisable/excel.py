# Copyright (c) 2010-2024 fastpyxl

"""Excel-specific serialisable types and XML helpers."""

from fastpyxl.compat import safe_string
from fastpyxl.xml.constants import REL_NS
from fastpyxl.xml.functions import Element

from .base import Serialisable
from .fields import Field


class Extension(Serialisable):

    tagname = "ext"

    uri: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)

    def __init__(self, uri=None):
        self.uri = uri


class ExtensionList(Serialisable):

    tagname = "extLst"

    ext: list[Extension] = Field.sequence(expected_type=Extension, default=list)

    def __init__(self, ext=()):
        self.ext = list(ext)


class NestedValInt(Serialisable):
    """Nested integer element with a val attribute."""

    tagname = "x"

    val: int | None = Field.nested_value(expected_type=int, allow_none=True, default=None)

    def __init__(self, val=None):
        self.val = val


def explicit_none_element(tagname, value, namespace=None):
    """Serialise explicit none values required by some chart elements."""
    if namespace is not None:
        tagname = "{%s}%s" % (namespace, tagname)
    return Element(tagname, val=safe_string(value))
