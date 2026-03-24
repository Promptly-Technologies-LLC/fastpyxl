# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import AliasField, Field
from fastpyxl.xml.constants import REL_NS


class SheetBackgroundPicture(Serialisable):
    tagname = "picture"
    id: str | None = Field.attribute(expected_type=str, allow_none=True, namespace=REL_NS)

    def __init__(self, id=None):
        self.id = id


class DrawingHF(Serialisable):
    id: str | None = Field.attribute(expected_type=str, allow_none=True, namespace=REL_NS)
    lho: int | None = Field.attribute(expected_type=int, allow_none=True)
    leftHeaderOddPages = AliasField("lho")
    lhe: int | None = Field.attribute(expected_type=int, allow_none=True)
    leftHeaderEvenPages = AliasField("lhe")
    lhf: int | None = Field.attribute(expected_type=int, allow_none=True)
    leftHeaderFirstPage = AliasField("lhf")
    cho: int | None = Field.attribute(expected_type=int, allow_none=True)
    centerHeaderOddPages = AliasField("cho")
    che: int | None = Field.attribute(expected_type=int, allow_none=True)
    centerHeaderEvenPages = AliasField("che")
    chf: int | None = Field.attribute(expected_type=int, allow_none=True)
    centerHeaderFirstPage = AliasField("chf")
    rho: int | None = Field.attribute(expected_type=int, allow_none=True)
    rightHeaderOddPages = AliasField("rho")
    rhe: int | None = Field.attribute(expected_type=int, allow_none=True)
    rightHeaderEvenPages = AliasField("rhe")
    rhf: int | None = Field.attribute(expected_type=int, allow_none=True)
    rightHeaderFirstPage = AliasField("rhf")
    lfo: int | None = Field.attribute(expected_type=int, allow_none=True)
    leftFooterOddPages = AliasField("lfo")
    lfe: int | None = Field.attribute(expected_type=int, allow_none=True)
    leftFooterEvenPages = AliasField("lfe")
    lff: int | None = Field.attribute(expected_type=int, allow_none=True)
    leftFooterFirstPage = AliasField("lff")
    cfo: int | None = Field.attribute(expected_type=int, allow_none=True)
    centerFooterOddPages = AliasField("cfo")
    cfe: int | None = Field.attribute(expected_type=int, allow_none=True)
    centerFooterEvenPages = AliasField("cfe")
    cff: int | None = Field.attribute(expected_type=int, allow_none=True)
    centerFooterFirstPage = AliasField("cff")
    rfo: int | None = Field.attribute(expected_type=int, allow_none=True)
    rightFooterOddPages = AliasField("rfo")
    rfe: int | None = Field.attribute(expected_type=int, allow_none=True)
    rightFooterEvenPages = AliasField("rfe")
    rff: int | None = Field.attribute(expected_type=int, allow_none=True)
    rightFooterFirstPage = AliasField("rff")

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
