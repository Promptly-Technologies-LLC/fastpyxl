# Copyright (c) 2010-2024 fastpyxl

from typing import Any, cast


def test_color_descriptor():
    from ..colors import ColorChoiceDescriptor

    class DummyStyle:

        value = ColorChoiceDescriptor('value')

    style = DummyStyle()
    style.value = "efefef"
    assert cast(Any, style.value).RGB == "efefef"
