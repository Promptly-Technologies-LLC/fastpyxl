# Copyright (c) 2010-2024 fastpyxl

"""
Richtext definition
"""

from fastpyxl.descriptors.serialisable import Serialisable
from fastpyxl.descriptors import (
    Alias,
    Typed,
    Integer,
    NoneSet,
    Sequence,
)
from fastpyxl.descriptors.nested import (
    NestedString,
    NestedText,
    NestedBool,
    NestedNoneSet,
    NestedMinMax,
    NestedInteger,
    NestedFloat,
)
from fastpyxl.styles.colors import ColorDescriptor
from fastpyxl.styles.fonts import _no_value


class PhoneticProperties(Serialisable):

    tagname = "phoneticPr"

    fontId = Integer()
    type = NoneSet(values=(['halfwidthKatakana', 'fullwidthKatakana',
                            'Hiragana', 'noConversion']))
    alignment = NoneSet(values=(['noControl', 'left', 'center', 'distributed']))

    def __init__(self,
                 fontId=None,
                 type=None,
                 alignment=None,
                ):
        self.fontId = fontId
        self.type = type
        self.alignment = alignment


class PhoneticText(Serialisable):

    tagname = "rPh"

    sb = Integer()
    eb = Integer()
    t = NestedText(expected_type=str)
    text = Alias('t')

    def __init__(self,
                 sb=None,
                 eb=None,
                 t=None,
                ):
        self.sb = sb
        self.eb = eb
        self.t = t


class InlineFont(Serialisable):

    """
    Font for inline text because, yes what you need are different objects with the same elements but different constraints.
    """

    tagname = "RPrElt"

    rFont = NestedString(allow_none=True)

    charset = NestedInteger(allow_none=True)
    family = NestedMinMax(min=0, max=14, allow_none=True)
    b = NestedBool(to_tree=_no_value)
    i = NestedBool(to_tree=_no_value)
    strike = NestedBool(allow_none=True)
    outline = NestedBool(allow_none=True)
    shadow = NestedBool(allow_none=True)
    condense = NestedBool(allow_none=True)
    extend = NestedBool(allow_none=True)
    color = ColorDescriptor(allow_none=True)
    sz = NestedFloat(allow_none=True)
    u = NestedNoneSet(
        values=(
            "single",
            "double",
            "singleAccounting",
            "doubleAccounting",
        )
    )
    vertAlign = NestedNoneSet(values=("superscript", "subscript", "baseline"))
    scheme = NestedNoneSet(values=("major", "minor"))

    __elements__ = ('rFont', 'charset', 'family', 'b', 'i', 'strike',
                    'outline', 'shadow', 'condense', 'extend', 'color', 'sz', 'u',
                    'vertAlign', 'scheme')

    def __init__(self,
                 rFont=None,
                 charset=None,
                 family=None,
                 b=None,
                 i=None,
                 strike=None,
                 outline=None,
                 shadow=None,
                 condense=None,
                 extend=None,
                 color=None,
                 sz=None,
                 u=None,
                 vertAlign=None,
                 scheme=None,
                ):
        self.rFont = rFont
        self.charset = charset
        self.family = family
        self.b = b
        self.i = i
        self.strike = strike
        self.outline = outline
        self.shadow = shadow
        self.condense = condense
        self.extend = extend
        self.color = color
        self.sz = sz
        self.u = u
        self.vertAlign = vertAlign
        self.scheme = scheme


class RichText(Serialisable):

    tagname = "RElt"

    rPr = Typed(expected_type=InlineFont, allow_none=True)
    font = Alias("rPr")
    t = NestedText(expected_type=str, allow_none=True)
    text = Alias("t")

    __elements__ = ('rPr', 't')

    def __init__(self,
                 rPr=None,
                 t=None,
                ):
        self.rPr = rPr
        self.t = t


class Text(Serialisable):

    tagname = "text"

    t = NestedText(allow_none=True, expected_type=str)
    plain = Alias("t")
    r = Sequence(expected_type=RichText, allow_none=True)
    formatted = Alias("r")
    rPh = Sequence(expected_type=PhoneticText, allow_none=True)
    phonetic = Alias("rPh")
    phoneticPr = Typed(expected_type=PhoneticProperties, allow_none=True)
    PhoneticProperties = Alias("phoneticPr")

    __elements__ = ('t', 'r', 'rPh', 'phoneticPr')

    def __init__(self,
                 t=None,
                 r=(),
                 rPh=(),
                 phoneticPr=None,
                ):
        self.t = t
        self.r = r
        self.rPh = rPh
        self.phoneticPr = phoneticPr


    @property
    def content(self):
        """
        Text stripped of all formatting
        """
        snippets = []
        if self.plain is not None:
            snippets.append(self.plain)
        for block in self.formatted:
            if block.t is not None:
                snippets.append(block.t)
        return u"".join(snippets)
