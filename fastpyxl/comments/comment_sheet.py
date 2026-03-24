# Copyright (c) 2010-2024 fastpyxl

## Incomplete!

import re

from fastpyxl.descriptors.excel import ExtensionList

from fastpyxl.utils.indexed_list import IndexedList
from fastpyxl.xml.constants import SHEET_MAIN_NS

from fastpyxl.cell.text import Text
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.errors import FieldValidationError
from fastpyxl.typed_serialisable.fields import Field

from .author import AuthorList
from .comments import Comment
from .shape_writer import ShapeWriter

_GUID_RE = re.compile(
    r"^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$"
)


def _guid_converter(value):
    if value is None:
        return None
    if not isinstance(value, str):
        raise FieldValidationError(f"guid rejected value {value!r}")
    v = value.upper()
    if _GUID_RE.match(v) is None:
        raise FieldValidationError(f"guid rejected value {value!r}")
    return v


_H_ALIGN = frozenset({"left", "center", "right", "justify", "distributed"})
_V_ALIGN = frozenset({"top", "center", "bottom", "justify", "distributed"})


def _h_align_converter(value):
    if value is None:
        return None
    if value not in _H_ALIGN:
        raise FieldValidationError(f"textHAlign rejected value {value!r}")
    return value


def _v_align_converter(value):
    if value is None:
        return None
    if value not in _V_ALIGN:
        raise FieldValidationError(f"textVAlign rejected value {value!r}")
    return value


class Properties(Serialisable):

    tagname = "commentPr"

    locked: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    defaultSize: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    _print: bool | None = Field.attribute(
        expected_type=bool, allow_none=True, xml_name="print", default=None
    )
    disabled: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    uiObject: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    autoFill: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    autoLine: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    altText: str | None = Field.attribute(expected_type=str, allow_none=True, default=None)
    textHAlign: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=_h_align_converter, default=None,
    )
    textVAlign: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        converter=_v_align_converter, default=None,
    )
    lockText: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    justLastX: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    autoScale: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    rowHidden: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)
    colHidden: bool | None = Field.attribute(expected_type=bool, allow_none=True, default=None)

    def __init__(
        self,
        locked=None,
        defaultSize=None,
        _print=None,
        disabled=None,
        uiObject=None,
        autoFill=None,
        autoLine=None,
        altText=None,
        textHAlign=None,
        textVAlign=None,
        lockText=None,
        justLastX=None,
        autoScale=None,
        rowHidden=None,
        colHidden=None,
        anchor=None,
    ):
        self.locked = locked
        self.defaultSize = defaultSize
        self._print = _print
        self.disabled = disabled
        self.uiObject = uiObject
        self.autoFill = autoFill
        self.autoLine = autoLine
        self.altText = altText
        self.textHAlign = textHAlign
        self.textVAlign = textVAlign
        self.lockText = lockText
        self.justLastX = justLastX
        self.autoScale = autoScale
        self.rowHidden = rowHidden
        self.colHidden = colHidden
        self.anchor = anchor


class CommentRecord(Serialisable):

    tagname = "comment"

    ref: str = Field.attribute(expected_type=str, default=None)
    authorId: int = Field.attribute(expected_type=int, default=None)
    guid: str | None = Field.attribute(
        expected_type=str, allow_none=True, converter=_guid_converter, default=None
    )
    shapeId: int | None = Field.attribute(expected_type=int, allow_none=True, default=None)
    text: Text = Field.element(expected_type=Text, default=None)
    commentPr: Properties | None = Field.element(expected_type=Properties, allow_none=True, default=None)

    xml_order = ("text", "commentPr")

    def __init__(
        self,
        ref="",
        authorId=0,
        guid=None,
        shapeId=0,
        text=None,
        commentPr=None,
        author=None,
        height=79,
        width=144,
    ):
        if text is None:
            text = Text()
        self.ref = ref
        self.authorId = authorId
        self.guid = guid
        self.shapeId = shapeId
        self.text = text
        self.commentPr = commentPr
        self.author = author
        self.height = height
        self.width = width

    @classmethod
    def from_cell(cls, cell):
        """
        Class method to convert cell comment
        """
        comment = cell._comment
        ref = cell.coordinate
        self = cls(ref=ref, author=comment.author)
        self.text.t = comment.content
        self.height = comment.height
        self.width = comment.width
        return self

    @property
    def content(self):
        """
        Remove all inline formatting and stuff
        """
        return self.text.content


class CommentSheet(Serialisable):

    tagname = "comments"

    authors: AuthorList = Field.element(expected_type=AuthorList, default=None)
    commentList: list[CommentRecord] = Field.nested_sequence(
        expected_type=CommentRecord, count=False, default=list
    )
    extLst: ExtensionList | None = Field.element(
        expected_type=ExtensionList, allow_none=True, serialize=False, default=None
    )

    _id = None
    _path = "/xl/comments/comment{0}.xml"
    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
    _rel_type = "comments"
    _rel_id = None

    xml_order = ("authors", "commentList")

    def __init__(
        self,
        authors=None,
        commentList=None,
        extLst=None,
    ):
        self.authors = authors or AuthorList()
        self.commentList = commentList or []
        self.extLst = extLst

    def to_tree(self, tagname=None, idx=None, namespace=None):
        tree = super().to_tree(tagname, idx, namespace)
        tree.set("xmlns", SHEET_MAIN_NS)
        return tree

    @property
    def comments(self):
        """
        Return a dictionary of comments keyed by coord
        """
        authors = self.authors.author

        for c in self.commentList:
            yield c.ref, Comment(c.content, authors[c.authorId], c.height, c.width)

    @classmethod
    def from_comments(cls, comments):
        """
        Create a comment sheet from a list of comments for a particular worksheet
        """
        authors = IndexedList()

        # dedupe authors and get indexes
        for comment in comments:
            comment.authorId = authors.add(comment.author)

        return cls(authors=AuthorList(authors), commentList=comments)

    def write_shapes(self, vml=None):
        """
        Create the VML for comments
        """
        sw = ShapeWriter(self.comments)
        return sw.write(vml)

    @property
    def path(self):
        """
        Return path within the archive
        """
        return self._path.format(self._id)
