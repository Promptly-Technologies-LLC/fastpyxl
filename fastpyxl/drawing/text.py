# Copyright (c) 2010-2024 fastpyxl

import re

from fastpyxl.xml.constants import DRAWING_NS, REL_NS
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import AliasField, Field

from .colors import ColorChoice
from .effect import (
    EffectList,
    EffectContainer,
)
from .fill import(
    GradientFillProperties,
    BlipFillProperties,
    PatternFillProperties,
    Blip
)
from .geometry import (
    LineProperties,
    Color,
    Scene3D
)

from fastpyxl.descriptors.excel import ExtensionList as OfficeArtExtensionList


def _none_set_member(v, allowed: frozenset[str], field_name: str) -> str | None:
    if v is None:
        return None
    if v == "none":
        return None
    if v not in allowed:
        raise FieldValidationError(f"{field_name} rejected value {v!r}")
    return v


def _defrpr_sz(v):
    if v is None:
        return None
    iv = int(v)
    if iv < 100 or iv > 400000:
        raise FieldValidationError(f"sz rejected value {v!r}")
    return iv


def _solid_fill_choice(v):
    if v is None:
        return None
    if isinstance(v, str):
        return ColorChoice(srgbClr=v)
    return v


def _bu_blip_embed(v):
    if v is None:
        return None
    if isinstance(v, Blip):
        return v
    return Blip(embed=v)


_RE_DEFRPR_U = frozenset({
    "words", "sng", "dbl", "heavy", "dotted", "dottedHeavy", "dash", "dashHeavy", "dashLong",
    "dashLongHeavy", "dotDash", "dotDashHeavy", "dotDotDash", "dotDotDashHeavy", "wavy",
    "wavyHeavy", "wavyDbl",
})
_RE_DEFRPR_STRIKE = frozenset({"noStrike", "sngStrike", "dblStrike"})
_RE_DEFRPR_CAP = frozenset({"small", "all"})

_RE_PPR_ALGN = frozenset({"l", "ctr", "r", "just", "justLow", "dist", "thaiDist"})
_RE_PPR_FONT_ALGN = frozenset({"auto", "t", "ctr", "base", "b"})

_RE_BODY_VERT_OVERFLOW = frozenset({"overflow", "ellipsis", "clip"})
_RE_BODY_HORZ_OVERFLOW = frozenset({"overflow", "clip"})
_RE_BODY_VERT = frozenset({
    "horz", "vert", "vert270", "wordArtVert", "eaVert", "mongolianVert", "wordArtVertRtl",
})
_RE_BODY_WRAP = frozenset({"none", "square"})
_RE_BODY_ANCHOR = frozenset({"t", "ctr", "b", "just", "dist"})

_PANOSE_RE = re.compile(r"^[0-9a-fA-F]+$")

_RE_TAB_STOP_ALGN = frozenset({"l", "ctr", "r", "dec"})

_RE_AUTONUM_TYPES = frozenset({
    "alphaLcParenBoth", "alphaUcParenBoth", "alphaLcParenR", "alphaUcParenR", "alphaLcPeriod",
    "alphaUcPeriod", "arabicParenBoth", "arabicParenR", "arabicPeriod", "arabicPlain",
    "romanLcParenBoth", "romanUcParenBoth", "romanLcParenR", "romanUcParenR", "romanLcPeriod",
    "romanUcPeriod", "circleNumDbPlain", "circleNumWdBlackPlain", "circleNumWdWhitePlain",
    "arabicDbPeriod", "arabicDbPlain", "ea1ChsPeriod", "ea1ChsPlain", "ea1ChtPeriod",
    "ea1ChtPlain", "ea1JpnChsDbPeriod", "ea1JpnKorPlain", "ea1JpnKorPeriod", "arabic1Minus",
    "arabic2Minus", "hebrew2Minus", "thaiAlphaPeriod", "thaiAlphaParenR", "thaiAlphaParenBoth",
    "thaiNumPeriod", "thaiNumParenR", "thaiNumParenBoth", "hindiAlphaPeriod", "hindiNumPeriod",
    "hindiNumParenR", "hindiAlpha1Period",
})

_RE_PRST_TEXT_SHAPE = frozenset({
    "textNoShape", "textPlain", "textStop", "textTriangle", "textTriangleInverted", "textChevron",
    "textChevronInverted", "textRingInside", "textRingOutside", "textArchUp", "textArchDown",
    "textCircle", "textButton", "textArchUpPour", "textArchDownPour", "textCirclePour",
    "textButtonPour", "textCurveUp", "textCurveDown", "textCanUp", "textCanDown", "textWave1",
    "textWave2", "textDoubleWave1", "textWave4", "textInflate", "textDeflate", "textInflateBottom",
    "textDeflateBottom", "textInflateTop", "textDeflateTop", "textDeflateInflate",
    "textDeflateInflateDeflate", "textFadeRight", "textFadeLeft", "textFadeUp", "textFadeDown",
    "textSlantUp", "textSlantDown", "textCascadeUp", "textCascadeDown",
})


def _font_panose(v):
    if v is None:
        return None
    if not _PANOSE_RE.match(v):
        raise FieldValidationError(f"panose rejected value {v!r}")
    return v


def _font_pitch_family(v):
    if v is None:
        return None
    iv = int(v)
    if iv < 0 or iv > 52:
        raise FieldValidationError(f"pitchFamily rejected value {v!r}")
    return iv


def _enum_member(v, allowed: frozenset[str], field_name: str) -> str | None:
    if v is None:
        return None
    if v not in allowed:
        raise FieldValidationError(f"{field_name} rejected value {v!r}")
    return v


class _CharNoFill(Serialisable):
    tagname = "noFill"
    namespace = DRAWING_NS


class _CharGrpFill(Serialisable):
    tagname = "grpFill"
    namespace = DRAWING_NS


class _CharULnTx(Serialisable):
    tagname = "uLnTx"
    namespace = DRAWING_NS


class _CharUFillTx(Serialisable):
    tagname = "uFillTx"
    namespace = DRAWING_NS


class _CharUFill(Serialisable):
    tagname = "uFill"
    namespace = DRAWING_NS


class _PPrBuClrTx(Serialisable):
    tagname = "buClrTx"
    namespace = DRAWING_NS


class _PPrBuSzTx(Serialisable):
    tagname = "buSzTx"
    namespace = DRAWING_NS


class _PPrBuFontTx(Serialisable):
    tagname = "buFontTx"
    namespace = DRAWING_NS


class _PPrBuNone(Serialisable):
    tagname = "buNone"
    namespace = DRAWING_NS


class _PPrBuAutoNum(Serialisable):
    tagname = "buAutoNum"
    namespace = DRAWING_NS


class _BodyNoAutofit(Serialisable):
    tagname = "noAutofit"
    namespace = DRAWING_NS


class _BodyNormAutofit(Serialisable):
    tagname = "normAutofit"
    namespace = DRAWING_NS


class _BodySpAutoFit(Serialisable):
    tagname = "spAutoFit"
    namespace = DRAWING_NS


class EmbeddedWAVAudioFile(Serialisable):

    tagname = "snd"
    namespace = DRAWING_NS

    name: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)

    def __init__(self,
                 name=None,
                ):
        self.name = name


class Hyperlink(Serialisable):

    tagname = "hlinkClick"
    namespace = DRAWING_NS

    invalidUrl: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    action: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    tgtFrame: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    tooltip: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    history: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    highlightClick: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    endSnd: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    snd: EmbeddedWAVAudioFile | None = Field.element(expected_type=EmbeddedWAVAudioFile, allow_none=True, default=None)
    extLst: OfficeArtExtensionList | None = Field.element(
        expected_type=OfficeArtExtensionList,
        allow_none=True,
        serialize=False, default=None,
    )
    id: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        namespace=REL_NS, default=None,
    )

    xml_order = ("snd",)

    def __init__(self,
                 invalidUrl=None,
                 action=None,
                 tgtFrame=None,
                 tooltip=None,
                 history=None,
                 highlightClick=None,
                 endSnd=None,
                 snd=None,
                 extLst=None,
                 id=None,
                ):
        self.invalidUrl = invalidUrl
        self.action = action
        self.tgtFrame = tgtFrame
        self.tooltip = tooltip
        self.history = history
        self.highlightClick = highlightClick
        self.endSnd = endSnd
        self.snd = snd
        self.extLst = extLst
        self.id = id


class Font(Serialisable):

    tagname = "latin"
    namespace = DRAWING_NS

    typeface: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    panose: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=_font_panose, default=None,
    )
    pitchFamily: int | None = Field.attribute(
        expected_type=int,
        allow_none=True,
        converter=_font_pitch_family, default=None,
    )
    charset: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

    def __init__(self,
                 typeface=None,
                 panose=None,
                 pitchFamily=None,
                 charset=None,
                ):
        self.typeface = typeface
        self.panose = panose
        self.pitchFamily = pitchFamily
        self.charset = charset


class CharacterProperties(Serialisable):

    tagname = "defRPr"
    namespace = DRAWING_NS

    kumimoji: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    lang: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    altLang: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    sz: int | None = Field.attribute(expected_type=int, allow_none=True, converter=_defrpr_sz, default=None)
    b: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    i: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    u: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _none_set_member(v, _RE_DEFRPR_U, "u"), default=None,
    )
    strike: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _none_set_member(v, _RE_DEFRPR_STRIKE, "strike"), default=None,
    )
    kern: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    cap: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _none_set_member(v, _RE_DEFRPR_CAP, "cap"), default=None,
    )
    spc: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    normalizeH: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    baseline: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    noProof: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    dirty: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    err: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    smtClean: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    smtId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    bmk: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    ln: LineProperties | None = Field.element(expected_type=LineProperties, allow_none=True, default=None)
    highlight: Color | None = Field.element(expected_type=Color, allow_none=True, default=None)
    latin: Font | None = Field.element(expected_type=Font, allow_none=True, default=None)
    ea: Font | None = Field.element(expected_type=Font, allow_none=True, default=None)
    cs: Font | None = Field.element(expected_type=Font, allow_none=True, default=None)
    sym: Font | None = Field.element(expected_type=Font, allow_none=True, default=None)
    hlinkClick: Hyperlink | None = Field.element(expected_type=Hyperlink, allow_none=True, default=None)
    hlinkMouseOver: Hyperlink | None = Field.element(expected_type=Hyperlink, allow_none=True, default=None)
    rtl: bool | None = Field.nested_bool(allow_none=True, default=None)
    extLst: OfficeArtExtensionList | None = Field.element(
        expected_type=OfficeArtExtensionList,
        allow_none=True,
        serialize=False, default=None,
    )
    noFill: _CharNoFill | None = Field.element(expected_type=_CharNoFill, allow_none=True, default=None)
    solidFill: ColorChoice | None = Field.element(
        expected_type=ColorChoice,
        allow_none=True,
        converter=_solid_fill_choice, default=None,
    )
    gradFill: GradientFillProperties | None = Field.element(
        expected_type=GradientFillProperties,
        allow_none=True, default=None,
    )
    blipFill: BlipFillProperties | None = Field.element(
        expected_type=BlipFillProperties,
        allow_none=True, default=None,
    )
    pattFill: PatternFillProperties | None = Field.element(
        expected_type=PatternFillProperties,
        allow_none=True, default=None,
    )
    grpFill: _CharGrpFill | None = Field.element(expected_type=_CharGrpFill, allow_none=True, default=None)
    effectLst: EffectList | None = Field.element(expected_type=EffectList, allow_none=True, default=None)
    effectDag: EffectContainer | None = Field.element(expected_type=EffectContainer, allow_none=True, default=None)
    uLnTx: _CharULnTx | None = Field.element(expected_type=_CharULnTx, allow_none=True, default=None)
    uLn: LineProperties | None = Field.element(expected_type=LineProperties, allow_none=True, default=None)
    uFillTx: _CharUFillTx | None = Field.element(expected_type=_CharUFillTx, allow_none=True, default=None)
    uFill: _CharUFill | None = Field.element(expected_type=_CharUFill, allow_none=True, default=None)

    xml_order = (
        "ln", "noFill", "solidFill", "gradFill", "blipFill",
        "pattFill", "grpFill", "effectLst", "effectDag", "highlight", "uLnTx",
        "uLn", "uFillTx", "uFill", "latin", "ea", "cs", "sym", "hlinkClick",
        "hlinkMouseOver", "rtl",
    )

    def __init__(self,
                 kumimoji=None,
                 lang=None,
                 altLang=None,
                 sz=None,
                 b=None,
                 i=None,
                 u=None,
                 strike=None,
                 kern=None,
                 cap=None,
                 spc=None,
                 normalizeH=None,
                 baseline=None,
                 noProof=None,
                 dirty=None,
                 err=None,
                 smtClean=None,
                 smtId=None,
                 bmk=None,
                 ln=None,
                 highlight=None,
                 latin=None,
                 ea=None,
                 cs=None,
                 sym=None,
                 hlinkClick=None,
                 hlinkMouseOver=None,
                 rtl=None,
                 extLst=None,
                 noFill=None,
                 solidFill=None,
                 gradFill=None,
                 blipFill=None,
                 pattFill=None,
                 grpFill=None,
                 effectLst=None,
                 effectDag=None,
                 uLnTx=None,
                 uLn=None,
                 uFillTx=None,
                 uFill=None,
                ):
        self.kumimoji = kumimoji
        self.lang = lang
        self.altLang = altLang
        self.sz = sz
        self.b = b
        self.i = i
        self.u = u
        self.strike = strike
        self.kern = kern
        self.cap = cap
        self.spc = spc
        self.normalizeH = normalizeH
        self.baseline = baseline
        self.noProof = noProof
        self.dirty = dirty
        self.err = err
        self.smtClean = smtClean
        self.smtId = smtId
        self.bmk = bmk
        self.ln = ln
        self.highlight = highlight
        self.latin = latin
        self.ea = ea
        self.cs = cs
        self.sym = sym
        self.hlinkClick = hlinkClick
        self.hlinkMouseOver = hlinkMouseOver
        self.rtl = rtl
        self.extLst = extLst
        self.noFill = _CharNoFill() if noFill else None
        self.solidFill = solidFill
        self.gradFill = gradFill
        self.blipFill = blipFill
        self.pattFill = pattFill
        self.grpFill = _CharGrpFill() if grpFill else None
        self.effectLst = effectLst
        self.effectDag = effectDag
        self.uLnTx = _CharULnTx() if uLnTx else None
        self.uLn = uLn
        self.uFillTx = _CharUFillTx() if uFillTx else None
        self.uFill = _CharUFill() if uFill else None


class TabStop(Serialisable):

    tagname = "tab"
    namespace = DRAWING_NS

    pos: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    algn: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_member(v, _RE_TAB_STOP_ALGN, "algn"), default=None,
    )

    def __init__(self,
                 pos=None,
                 algn=None,
                ):
        self.pos = pos
        self.algn = algn


class TabStopList(Serialisable):

    tagname = "tabLst"
    namespace = DRAWING_NS

    tab: TabStop | None = Field.element(expected_type=TabStop, allow_none=True, default=None)

    def __init__(self,
                 tab=None,
                ):
        self.tab = tab


class Spacing(Serialisable):

    spcPct: int | None = Field.nested_value(expected_type=int, allow_none=True, default=None)
    spcPts: int | None = Field.nested_value(expected_type=int, allow_none=True, default=None)

    xml_order = ("spcPct", "spcPts")

    def __init__(self,
                 spcPct=None,
                 spcPts=None,
                 ):
        self.spcPct = spcPct
        self.spcPts = spcPts


class AutonumberBullet(Serialisable):

    tagname = "buAutoNum"
    namespace = DRAWING_NS

    type: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _none_set_member(v, _RE_AUTONUM_TYPES, "type"), default=None,
    )
    startAt: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

    def __init__(self,
                 type=None,
                 startAt=None,
                ):
        self.type = type
        self.startAt = startAt


class ParagraphProperties(Serialisable):

    tagname = "pPr"
    namespace = DRAWING_NS

    marL: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    marR: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    lvl: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    indent: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    algn: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _none_set_member(v, _RE_PPR_ALGN, "algn"), default=None,
    )
    defTabSz: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    rtl: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    eaLnBrk: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    fontAlgn: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _none_set_member(v, _RE_PPR_FONT_ALGN, "fontAlgn"), default=None,
    )
    latinLnBrk: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    hangingPunct: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)

    lnSpc: Spacing | None = Field.element(expected_type=Spacing, allow_none=True, default=None)
    spcBef: Spacing | None = Field.element(expected_type=Spacing, allow_none=True, default=None)
    spcAft: Spacing | None = Field.element(expected_type=Spacing, allow_none=True, default=None)
    tabLst: TabStopList | None = Field.element(expected_type=TabStopList, allow_none=True, default=None)
    defRPr: CharacterProperties | None = Field.element(expected_type=CharacterProperties, allow_none=True, default=None)
    extLst: OfficeArtExtensionList | None = Field.element(
        expected_type=OfficeArtExtensionList,
        allow_none=True,
        serialize=False, default=None,
    )
    buClrTx: _PPrBuClrTx | None = Field.element(expected_type=_PPrBuClrTx, allow_none=True, default=None)
    buClr: Color | None = Field.element(expected_type=Color, allow_none=True, default=None)
    buSzTx: _PPrBuSzTx | None = Field.element(expected_type=_PPrBuSzTx, allow_none=True, default=None)
    buSzPct: int | None = Field.nested_value(expected_type=int, allow_none=True, default=None)
    buSzPts: int | None = Field.nested_value(expected_type=int, allow_none=True, default=None)
    buFontTx: _PPrBuFontTx | None = Field.element(expected_type=_PPrBuFontTx, allow_none=True, default=None)
    buFont: Font | None = Field.element(expected_type=Font, allow_none=True, default=None)
    buNone: _PPrBuNone | None = Field.element(expected_type=_PPrBuNone, allow_none=True, default=None)
    buAutoNum: _PPrBuAutoNum | None = Field.element(expected_type=_PPrBuAutoNum, allow_none=True, default=None)
    buChar: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        value_attribute="char", default=None,
    )
    buBlip: Blip | None = Field.nested_value(
        expected_type=Blip,
        allow_none=True,
        value_attribute="blip",
        converter=_bu_blip_embed, default=None,
    )

    xml_order = (
        "lnSpc", "spcBef", "spcAft", "tabLst", "defRPr",
        "buClrTx", "buClr", "buSzTx", "buSzPct", "buSzPts", "buFontTx", "buFont",
        "buNone", "buAutoNum", "buChar", "buBlip",
    )

    def __init__(self,
                 marL=None,
                 marR=None,
                 lvl=None,
                 indent=None,
                 algn=None,
                 defTabSz=None,
                 rtl=None,
                 eaLnBrk=None,
                 fontAlgn=None,
                 latinLnBrk=None,
                 hangingPunct=None,
                 lnSpc=None,
                 spcBef=None,
                 spcAft=None,
                 tabLst=None,
                 defRPr=None,
                 extLst=None,
                 buClrTx=None,
                 buClr=None,
                 buSzTx=None,
                 buSzPct=None,
                 buSzPts=None,
                 buFontTx=None,
                 buFont=None,
                 buNone=None,
                 buAutoNum=None,
                 buChar=None,
                 buBlip=None,
                 ):
        self.marL = marL
        self.marR = marR
        self.lvl = lvl
        self.indent = indent
        self.algn = algn
        self.defTabSz = defTabSz
        self.rtl = rtl
        self.eaLnBrk = eaLnBrk
        self.fontAlgn = fontAlgn
        self.latinLnBrk = latinLnBrk
        self.hangingPunct = hangingPunct
        self.lnSpc = lnSpc
        self.spcBef = spcBef
        self.spcAft = spcAft
        self.tabLst = tabLst
        self.defRPr = defRPr
        self.extLst = extLst
        self.buClrTx = _PPrBuClrTx() if buClrTx else None
        self.buClr = buClr
        self.buSzTx = _PPrBuSzTx() if buSzTx else None
        self.buSzPct = buSzPct
        self.buSzPts = buSzPts
        self.buFontTx = _PPrBuFontTx() if buFontTx else None
        self.buFont = buFont
        self.buNone = _PPrBuNone() if buNone else None
        self.buAutoNum = _PPrBuAutoNum() if buAutoNum else None
        self.buChar = buChar
        self.buBlip = buBlip


class ListStyle(Serialisable):

    tagname = "lstStyle"
    namespace = DRAWING_NS

    defPPr: ParagraphProperties | None = Field.element(expected_type=ParagraphProperties, allow_none=True, default=None)
    lvl1pPr: ParagraphProperties | None = Field.element(expected_type=ParagraphProperties, allow_none=True, default=None)
    lvl2pPr: ParagraphProperties | None = Field.element(expected_type=ParagraphProperties, allow_none=True, default=None)
    lvl3pPr: ParagraphProperties | None = Field.element(expected_type=ParagraphProperties, allow_none=True, default=None)
    lvl4pPr: ParagraphProperties | None = Field.element(expected_type=ParagraphProperties, allow_none=True, default=None)
    lvl5pPr: ParagraphProperties | None = Field.element(expected_type=ParagraphProperties, allow_none=True, default=None)
    lvl6pPr: ParagraphProperties | None = Field.element(expected_type=ParagraphProperties, allow_none=True, default=None)
    lvl7pPr: ParagraphProperties | None = Field.element(expected_type=ParagraphProperties, allow_none=True, default=None)
    lvl8pPr: ParagraphProperties | None = Field.element(expected_type=ParagraphProperties, allow_none=True, default=None)
    lvl9pPr: ParagraphProperties | None = Field.element(expected_type=ParagraphProperties, allow_none=True, default=None)
    extLst: OfficeArtExtensionList | None = Field.element(
        expected_type=OfficeArtExtensionList,
        allow_none=True,
        serialize=False, default=None,
    )

    xml_order = (
        "defPPr", "lvl1pPr", "lvl2pPr", "lvl3pPr", "lvl4pPr",
        "lvl5pPr", "lvl6pPr", "lvl7pPr", "lvl8pPr", "lvl9pPr",
    )

    def __init__(self,
                 defPPr=None,
                 lvl1pPr=None,
                 lvl2pPr=None,
                 lvl3pPr=None,
                 lvl4pPr=None,
                 lvl5pPr=None,
                 lvl6pPr=None,
                 lvl7pPr=None,
                 lvl8pPr=None,
                 lvl9pPr=None,
                 extLst=None,
                ):
        self.defPPr = defPPr
        self.lvl1pPr = lvl1pPr
        self.lvl2pPr = lvl2pPr
        self.lvl3pPr = lvl3pPr
        self.lvl4pPr = lvl4pPr
        self.lvl5pPr = lvl5pPr
        self.lvl6pPr = lvl6pPr
        self.lvl7pPr = lvl7pPr
        self.lvl8pPr = lvl8pPr
        self.lvl9pPr = lvl9pPr
        self.extLst = extLst


class RegularTextRun(Serialisable):

    tagname = "r"
    namespace = DRAWING_NS

    rPr: CharacterProperties | None = Field.element(expected_type=CharacterProperties, allow_none=True, default=None)
    properties = AliasField("rPr", default=None)
    t: str = Field.nested_text(expected_type=str, default=None)
    value = AliasField("t", default=None)

    xml_order = ("rPr", "t")

    def __init__(self,
                 rPr=None,
                 t="",
                ):
        self.rPr = rPr
        self.t = t


class LineBreak(Serialisable):

    tagname = "br"
    namespace = DRAWING_NS

    rPr: CharacterProperties | None = Field.element(expected_type=CharacterProperties, allow_none=True, default=None)

    def __init__(self,
                 rPr=None,
                ):
        self.rPr = rPr


class TextField(Serialisable):

    id: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    type: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    rPr: CharacterProperties | None = Field.element(expected_type=CharacterProperties, allow_none=True, default=None)
    pPr: ParagraphProperties | None = Field.element(expected_type=ParagraphProperties, allow_none=True, default=None)
    t: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)

    xml_order = ("rPr", "pPr")

    def __init__(self,
                 id=None,
                 type=None,
                 rPr=None,
                 pPr=None,
                 t=None,
                ):
        self.id = id
        self.type = type
        self.rPr = rPr
        self.pPr = pPr
        self.t = t


class Paragraph(Serialisable):

    tagname = "p"
    namespace = DRAWING_NS

    # uses element group EG_TextRun
    pPr: ParagraphProperties | None = Field.element(expected_type=ParagraphProperties, allow_none=True, default=None)
    properties = AliasField("pPr", default=None)
    endParaRPr: CharacterProperties | None = Field.element(expected_type=CharacterProperties, allow_none=True, default=None)
    r: list[RegularTextRun] = Field.sequence(expected_type=RegularTextRun, default=list)
    text = AliasField("r", default=None)
    br: LineBreak | None = Field.element(expected_type=LineBreak, allow_none=True, default=None)
    fld: TextField | None = Field.element(expected_type=TextField, allow_none=True, default=None)

    xml_order = ("pPr", "r", "br", "fld", "endParaRPr")

    def __init__(self,
                 pPr=None,
                 endParaRPr=None,
                 r=None,
                 br=None,
                 fld=None,
                 ):
        self.pPr = pPr
        self.endParaRPr = endParaRPr
        if r is None:
            r = [RegularTextRun()]
        self.r = r
        self.br = br
        self.fld = fld


class GeomGuide(Serialisable):

    tagname = "gd"
    namespace = DRAWING_NS

    name: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    fmla: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)

    def __init__(self,
                 name=None,
                 fmla=None,
                ):
        self.name = name
        self.fmla = fmla


class GeomGuideList(Serialisable):

    tagname = "avLst"
    namespace = DRAWING_NS

    gd: list[GeomGuide] | None = Field.sequence(expected_type=GeomGuide, allow_none=True, default=list)

    def __init__(self,
                 gd=None,
                ):
        self.gd = gd


class PresetTextShape(Serialisable):

    namespace = DRAWING_NS

    prst: str | None = Field.nested_value(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _enum_member(v, _RE_PRST_TEXT_SHAPE, "prst"), default=None,
    )
    avLst: GeomGuideList | None = Field.element(expected_type=GeomGuideList, allow_none=True, default=None)

    xml_order = ("prst", "avLst")

    def __init__(self,
                 prst=None,
                 avLst=None,
                ):
        self.prst = prst
        self.avLst = avLst


class TextNormalAutofit(Serialisable):

    tagname = "normAutofit"
    namespace = DRAWING_NS

    fontScale: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    lnSpcReduction: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)

    def __init__(self,
                 fontScale=None,
                 lnSpcReduction=None,
                ):
        self.fontScale = fontScale
        self.lnSpcReduction = lnSpcReduction


class RichTextProperties(Serialisable):

    tagname = "bodyPr"
    namespace = DRAWING_NS

    rot: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    spcFirstLastPara: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    vertOverflow: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _none_set_member(v, _RE_BODY_VERT_OVERFLOW, "vertOverflow"), default=None,
    )
    horzOverflow: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _none_set_member(v, _RE_BODY_HORZ_OVERFLOW, "horzOverflow"), default=None,
    )
    vert: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _none_set_member(v, _RE_BODY_VERT, "vert"), default=None,
    )
    wrap: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _none_set_member(v, _RE_BODY_WRAP, "wrap"), default=None,
    )
    lIns: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    tIns: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    rIns: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    bIns: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    numCol: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    spcCol: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    rtlCol: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    fromWordArt: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    anchor: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=lambda v: _none_set_member(v, _RE_BODY_ANCHOR, "anchor"), default=None,
    )
    anchorCtr: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    forceAA: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    upright: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    compatLnSpc: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    prstTxWarp: PresetTextShape | None = Field.element(expected_type=PresetTextShape, allow_none=True, default=None)
    scene3d: Scene3D | None = Field.element(expected_type=Scene3D, allow_none=True, default=None)
    extLst: OfficeArtExtensionList | None = Field.element(
        expected_type=OfficeArtExtensionList,
        allow_none=True,
        serialize=False, default=None,
    )
    noAutofit: _BodyNoAutofit | None = Field.element(expected_type=_BodyNoAutofit, allow_none=True, default=None)
    normAutofit: _BodyNormAutofit | None = Field.element(expected_type=_BodyNormAutofit, allow_none=True, default=None)
    spAutoFit: _BodySpAutoFit | None = Field.element(expected_type=_BodySpAutoFit, allow_none=True, default=None)
    flatTx: int | None = Field.nested_value(
        expected_type=int,
        allow_none=True,
        value_attribute="z",
        serialize=False, default=None,
    )

    xml_order = ("prstTxWarp", "scene3d", "noAutofit", "normAutofit", "spAutoFit")

    def __init__(self,
                 rot=None,
                 spcFirstLastPara=None,
                 vertOverflow=None,
                 horzOverflow=None,
                 vert=None,
                 wrap=None,
                 lIns=None,
                 tIns=None,
                 rIns=None,
                 bIns=None,
                 numCol=None,
                 spcCol=None,
                 rtlCol=None,
                 fromWordArt=None,
                 anchor=None,
                 anchorCtr=None,
                 forceAA=None,
                 upright=None,
                 compatLnSpc=None,
                 prstTxWarp=None,
                 scene3d=None,
                 extLst=None,
                 noAutofit=None,
                 normAutofit=None,
                 spAutoFit=None,
                 flatTx=None,
                ):
        self.rot = rot
        self.spcFirstLastPara = spcFirstLastPara
        self.vertOverflow = vertOverflow
        self.horzOverflow = horzOverflow
        self.vert = vert
        self.wrap = wrap
        self.lIns = lIns
        self.tIns = tIns
        self.rIns = rIns
        self.bIns = bIns
        self.numCol = numCol
        self.spcCol = spcCol
        self.rtlCol = rtlCol
        self.fromWordArt = fromWordArt
        self.anchor = anchor
        self.anchorCtr = anchorCtr
        self.forceAA = forceAA
        self.upright = upright
        self.compatLnSpc = compatLnSpc
        self.prstTxWarp = prstTxWarp
        self.scene3d = scene3d
        self.extLst = extLst
        self.noAutofit = _BodyNoAutofit() if noAutofit else None
        self.normAutofit = _BodyNormAutofit() if normAutofit else None
        self.spAutoFit = _BodySpAutoFit() if spAutoFit else None
        self.flatTx = flatTx
