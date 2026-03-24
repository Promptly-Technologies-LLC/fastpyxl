# Copyright (c) 2010-2024 fastpyxl

from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import Field
from fastpyxl.xml.functions import Element

from fastpyxl.xml.constants import XPROPS_NS
from fastpyxl import __version__


def _invert_xml_bool(value):
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.lower() not in ("true", "1", "t")
    return not bool(value)


def _invert_xml_bool_renderer(tagname, value, namespace=None):
    if value is None:
        return None
    if namespace is not None:
        tagname = "{%s}%s" % (namespace, tagname)
    el = Element(tagname)
    el.text = "false" if value else "true"
    return el


class DigSigBlob(Serialisable):

    __elements__ = __attrs__ = ()


class VectorLpstr(Serialisable):

    __elements__ = __attrs__ = ()


class VectorVariant(Serialisable):

    __elements__ = __attrs__ = ()


class ExtendedProperties(Serialisable):

    """
    See 22.2

    Most of this is irrelevant but Excel is very picky about the version number

    It uses XX.YYYY (Version.Build) and expects everyone else to

    We provide Major.Minor and the full version in the application name
    """

    tagname = "Properties"

    Template: str | None = Field.nested_text(expected_type=str, allow_none=True)
    Manager: str | None = Field.nested_text(expected_type=str, allow_none=True)
    Company: str | None = Field.nested_text(expected_type=str, allow_none=True)
    Pages: int | None = Field.nested_text(expected_type=int, allow_none=True)
    Words: int | None = Field.nested_text(expected_type=int, allow_none=True)
    Characters: int | None = Field.nested_text(expected_type=int, allow_none=True)
    PresentationFormat: str | None = Field.nested_text(expected_type=str, allow_none=True)
    Lines: int | None = Field.nested_text(expected_type=int, allow_none=True)
    Paragraphs: int | None = Field.nested_text(expected_type=int, allow_none=True)
    Slides: int | None = Field.nested_text(expected_type=int, allow_none=True)
    Notes: int | None = Field.nested_text(expected_type=int, allow_none=True)
    TotalTime: int | None = Field.nested_text(expected_type=int, allow_none=True)
    HiddenSlides: int | None = Field.nested_text(expected_type=int, allow_none=True)
    MMClips: int | None = Field.nested_text(expected_type=int, allow_none=True)
    ScaleCrop: bool | None = Field.nested_text(
        expected_type=bool,
        allow_none=True,
        converter=_invert_xml_bool,
        renderer=_invert_xml_bool_renderer,
    )
    HeadingPairs: VectorVariant | None = Field.element(expected_type=VectorVariant, allow_none=True)
    TitlesOfParts: VectorLpstr | None = Field.element(expected_type=VectorLpstr, allow_none=True)
    LinksUpToDate: bool | None = Field.nested_text(
        expected_type=bool,
        allow_none=True,
        converter=_invert_xml_bool,
        renderer=_invert_xml_bool_renderer,
    )
    CharactersWithSpaces: int | None = Field.nested_text(expected_type=int, allow_none=True)
    SharedDoc: bool | None = Field.nested_text(
        expected_type=bool,
        allow_none=True,
        converter=_invert_xml_bool,
        renderer=_invert_xml_bool_renderer,
    )
    HyperlinkBase: str | None = Field.nested_text(expected_type=str, allow_none=True)
    HLinks: VectorVariant | None = Field.element(expected_type=VectorVariant, allow_none=True)
    HyperlinksChanged: bool | None = Field.nested_text(
        expected_type=bool,
        allow_none=True,
        converter=_invert_xml_bool,
        renderer=_invert_xml_bool_renderer,
    )
    DigSig: DigSigBlob | None = Field.element(expected_type=DigSigBlob, allow_none=True)
    Application: str | None = Field.nested_text(expected_type=str, allow_none=True)
    AppVersion: str | None = Field.nested_text(expected_type=str, allow_none=True)
    DocSecurity: int | None = Field.nested_text(expected_type=int, allow_none=True)

    xml_order = ('Application', 'AppVersion', 'DocSecurity', 'ScaleCrop',
                 'LinksUpToDate', 'SharedDoc', 'HyperlinksChanged')

    def __init__(self,
                 Template=None,
                 Manager=None,
                 Company=None,
                 Pages=None,
                 Words=None,
                 Characters=None,
                 PresentationFormat=None,
                 Lines=None,
                 Paragraphs=None,
                 Slides=None,
                 Notes=None,
                 TotalTime=None,
                 HiddenSlides=None,
                 MMClips=None,
                 ScaleCrop=None,
                 HeadingPairs=None,
                 TitlesOfParts=None,
                 LinksUpToDate=None,
                 CharactersWithSpaces=None,
                 SharedDoc=None,
                 HyperlinkBase=None,
                 HLinks=None,
                 HyperlinksChanged=None,
                 DigSig=None,
                 Application=None,
                 AppVersion=None,
                 DocSecurity=None,
                ):
        self.Template = Template
        self.Manager = Manager
        self.Company = Company
        self.Pages = Pages
        self.Words = Words
        self.Characters = Characters
        self.PresentationFormat = PresentationFormat
        self.Lines = Lines
        self.Paragraphs = Paragraphs
        self.Slides = Slides
        self.Notes = Notes
        self.TotalTime = TotalTime
        self.HiddenSlides = HiddenSlides
        self.MMClips = MMClips
        self.ScaleCrop = ScaleCrop
        self.HeadingPairs = HeadingPairs
        self.TitlesOfParts = TitlesOfParts
        self.LinksUpToDate = LinksUpToDate
        self.CharactersWithSpaces = CharactersWithSpaces
        self.SharedDoc = SharedDoc
        self.HyperlinkBase = HyperlinkBase
        self.HLinks = HLinks
        self.HyperlinksChanged = HyperlinksChanged
        self.DigSig = DigSig
        if Application is None:
            self.Application = f"Microsoft Excel Compatible / Openpyxl {__version__}"
        else:
            self.Application = Application
        if AppVersion is None:
            self.AppVersion = ".".join(__version__.split(".")[:-1])
        else:
            self.AppVersion = AppVersion
        self.DocSecurity = DocSecurity


    def to_tree(self):
        tree = super().to_tree()
        tree.set("xmlns", XPROPS_NS)
        return tree
