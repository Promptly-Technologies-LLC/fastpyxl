# Copyright (c) 2010-2024 fastpyxl


def test_color_descriptor():
    from ..colors import ColorChoiceDescriptor

    class DummyStyle:

        value = ColorChoiceDescriptor('value')

    style = DummyStyle()
    style.value = "efefef"
    style.value.RGB == "efefef"
