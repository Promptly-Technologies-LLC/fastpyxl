# Copyright (c) 2010-2024 fastpyxl

from array import array

from fastpyxl.compat import safe_string
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import Field
from fastpyxl.descriptors.excel import ExtensionList
from fastpyxl.utils.indexed_list import IndexedList


from .alignment import Alignment
from .protection import Protection


class ArrayDescriptor:

    def __init__(self, key):
        self.key = key

    def __get__(self, instance, cls):
        return instance[self.key]

    def __set__(self, instance, value):
        instance[self.key] = value


class StyleArray(array):
    """
    Simplified named tuple with an array
    """

    __slots__ = ('_hash',)
    tagname = 'xf'

    fontId = ArrayDescriptor(0)
    fillId = ArrayDescriptor(1)
    borderId = ArrayDescriptor(2)
    numFmtId = ArrayDescriptor(3)
    protectionId = ArrayDescriptor(4)
    alignmentId = ArrayDescriptor(5)
    pivotButton = ArrayDescriptor(6)
    quotePrefix = ArrayDescriptor(7)
    xfId = ArrayDescriptor(8)


    def __new__(cls, args=[0]*9):
        obj = array.__new__(cls, 'i', args)
        obj._hash = None
        return obj


    def __setitem__(self, index, value):
        self._hash = None
        array.__setitem__(self, index, value)


    def __hash__(self):
        h = self._hash
        if h is None:
            h = hash(tuple(self))
            self._hash = h
        return h


    def __copy__(self):
        return StyleArray((self))


    def __deepcopy__(self, memo):
        return StyleArray((self))


class CellStyle(Serialisable):

    tagname = "xf"

    numFmtId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    fontId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    fillId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    borderId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    xfId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    quotePrefix: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    pivotButton: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    alignment: Alignment | None = Field.element(expected_type=Alignment, allow_none=True, default=None)
    protection: Protection | None = Field.element(expected_type=Protection, allow_none=True, default=None)
    extLst: ExtensionList | None = Field.element(expected_type=ExtensionList, allow_none=True, default=None)

    xml_order = ("alignment", "protection")

    def __init__(self,
                 numFmtId=0,
                 fontId=0,
                 fillId=0,
                 borderId=0,
                 xfId=None,
                 quotePrefix=None,
                 pivotButton=None,
                 applyNumberFormat=None,
                 applyFont=None,
                 applyFill=None,
                 applyBorder=None,
                 applyAlignment=None,
                 applyProtection=None,
                 alignment=None,
                 protection=None,
                 extLst=None,
                ):
        self.numFmtId = numFmtId
        self.fontId = fontId
        self.fillId = fillId
        self.borderId = borderId
        self.xfId = xfId
        self.quotePrefix = quotePrefix
        self.pivotButton = pivotButton
        self.applyNumberFormat = applyNumberFormat
        self.applyFont = applyFont
        self.applyFill = applyFill
        self.applyBorder = applyBorder
        self.alignment = alignment
        self.protection = protection
        self.extLst = extLst


    def to_array(self):
        """
        Convert to StyleArray
        """
        style = StyleArray()
        for idx, v in (
            (0, self.fontId), (1, self.fillId), (2, self.borderId),
            (3, self.numFmtId), (6, self.pivotButton),
            (7, self.quotePrefix), (8, self.xfId),
        ):
            if v is not None:
                style[idx] = v
        return style


    @classmethod
    def from_array(cls, style):
        """
        Convert from StyleArray
        """
        return cls(numFmtId=style.numFmtId, fontId=style.fontId,
                   fillId=style.fillId, borderId=style.borderId, xfId=style.xfId,
                   quotePrefix=style.quotePrefix, pivotButton=style.pivotButton,)


    @property
    def applyProtection(self):
        return self.protection is not None or None


    @property
    def applyAlignment(self):
        return self.alignment is not None or None


    def __iter__(self):
        attrs = (
            ("numFmtId", self.numFmtId),
            ("fontId", self.fontId),
            ("fillId", self.fillId),
            ("borderId", self.borderId),
            ("applyAlignment", self.applyAlignment),
            ("applyProtection", self.applyProtection),
            ("pivotButton", self.pivotButton),
            ("quotePrefix", self.quotePrefix),
            ("xfId", self.xfId),
        )
        for key, value in attrs:
            if value is None:
                continue
            yield key, safe_string(value)


class CellStyleList(Serialisable):

    tagname = "cellXfs"

    xf: list[CellStyle] = Field.sequence(expected_type=CellStyle, default=list)

    def __init__(self,
                 count=None,
                 xf=(),
                ):
        self.xf = list(xf)


    @property
    def count(self):
        return len(self.xf)


    def __iter__(self):
        yield "count", safe_string(self.count)


    def __getitem__(self, idx):
        try:
            return self.xf[idx]
        except IndexError:
            print((f"{idx} is out of range"))
        return self.xf[idx]


    def _to_array(self):
        """
        Extract protection and alignments, convert to style array
        """
        self.prots = IndexedList([Protection()])
        self.alignments = IndexedList([Alignment()])
        styles = [] # allow duplicates
        for xf in self.xf:
            style = xf.to_array()
            if xf.alignment is not None:
                style.alignmentId = self.alignments.add(xf.alignment)
            if xf.protection is not None:
                style.protectionId = self.prots.add(xf.protection)
            styles.append(style)
        return IndexedList(styles)
