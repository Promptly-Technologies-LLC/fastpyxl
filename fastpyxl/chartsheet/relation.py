# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import AliasField, Field
from fastpyxl.xml.constants import REL_NS


class SheetBackgroundPicture(Serialisable):
    tagname = "picture"
    id: str | None = Field.attribute(expected_type=str, allow_none=True, namespace=REL_NS, default=None)

    def __init__(self, id=None):
        self.id = id


class DrawingHF(Serialisable):
    id: str | None = Field.attribute(expected_type=str, allow_none=True, namespace=REL_NS, default=None)
    lho: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    leftHeaderOddPages = AliasField("lho", default=None)
    lhe: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    leftHeaderEvenPages = AliasField("lhe", default=None)
    lhf: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    leftHeaderFirstPage = AliasField("lhf", default=None)
    cho: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    centerHeaderOddPages = AliasField("cho", default=None)
    che: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    centerHeaderEvenPages = AliasField("che", default=None)
    chf: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    centerHeaderFirstPage = AliasField("chf", default=None)
    rho: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    rightHeaderOddPages = AliasField("rho", default=None)
    rhe: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    rightHeaderEvenPages = AliasField("rhe", default=None)
    rhf: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    rightHeaderFirstPage = AliasField("rhf", default=None)
    lfo: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    leftFooterOddPages = AliasField("lfo", default=None)
    lfe: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    leftFooterEvenPages = AliasField("lfe", default=None)
    lff: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    leftFooterFirstPage = AliasField("lff", default=None)
    cfo: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    centerFooterOddPages = AliasField("cfo", default=None)
    cfe: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    centerFooterEvenPages = AliasField("cfe", default=None)
    cff: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    centerFooterFirstPage = AliasField("cff", default=None)
    rfo: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    rightFooterOddPages = AliasField("rfo", default=None)
    rfe: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    rightFooterEvenPages = AliasField("rfe", default=None)
    rff: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    rightFooterFirstPage = AliasField("rff", default=None)

    def __init__(self,
                 id=None,
                 lho=None,
                 lhe=None,
                 lhf=None,
                 cho=None,
                 che=None,
                 chf=None,
                 rho=None,
                 rhe=None,
                 rhf=None,
                 lfo=None,
                 lfe=None,
                 lff=None,
                 cfo=None,
                 cfe=None,
                 cff=None,
                 rfo=None,
                 rfe=None,
                 rff=None,
                 ):
        self.id = id
        self.lho = lho
        self.lhe = lhe
        self.lhf = lhf
        self.cho = cho
        self.che = che
        self.chf = chf
        self.rho = rho
        self.rhe = rhe
        self.rhf = rhf
        self.lfo = lfo
        self.lfe = lfe
        self.lff = lff
        self.cfo = cfo
        self.cfe = cfe
        self.cff = cff
        self.rfo = rfo
        self.rfe = rfe
        self.rff = rff
