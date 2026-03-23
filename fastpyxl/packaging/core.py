# Copyright (c) 2010-2024 fastpyxl

import datetime

from fastpyxl.descriptors import (
    DateTime,
)
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import AliasField, Field
from fastpyxl.descriptors.nested import NestedText
from fastpyxl.xml.functions import (
    Element,
    QName,
)
from fastpyxl.xml.constants import (
    COREPROPS_NS,
    DCORE_NS,
    XSI_NS,
    DCTERMS_NS,
)


class NestedDateTime(DateTime, NestedText):

    expected_type = datetime.datetime

    def to_tree(self, tagname=None, value=None, namespace=None):
        namespace = getattr(self, "namespace", namespace)
        if namespace is not None:
            tagname = "{%s}%s" % (namespace, tagname)
        el = Element(tagname)
        if value is not None:
            value = value.replace(tzinfo=None)
            el.text = value.isoformat(timespec="seconds") + 'Z'
            return el


class QualifiedDateTime(NestedDateTime):

    """In certain situations Excel will complain if the additional type
    attribute isn't set"""

    def to_tree(self, tagname=None, value=None, namespace=None):
        el = super().to_tree(tagname, value, namespace)
        el.set("{%s}type" % XSI_NS, QName(DCTERMS_NS, "W3CDTF"))
        return el


def _datetime_converter(value):
    if value is None:
        return None
    if isinstance(value, datetime.datetime):
        return value
    if isinstance(value, str):
        return datetime.datetime.fromisoformat(value.replace("Z", "+00:00")).replace(tzinfo=None)
    raise TypeError(f"datetime rejected value {value!r}")


def _datetime_renderer(tagname, value, namespace=None):
    if value is None:
        return None
    ns = namespace
    if ns is not None:
        tagname = "{%s}%s" % (ns, tagname)
    el = Element(tagname)
    text = value.replace(tzinfo=None).isoformat(timespec="seconds") + "Z"
    el.text = text
    return el


def _qualified_datetime_renderer(tagname, value, namespace=None):
    el = _datetime_renderer(tagname, value, namespace)
    if el is None:
        return None
    el.set("{%s}type" % XSI_NS, QName(DCTERMS_NS, "W3CDTF"))
    return el


class DocumentProperties(Serialisable):
    """High-level properties of the document.
    Defined in ECMA-376 Par2 Annex D
    """

    tagname = "coreProperties"
    namespace = COREPROPS_NS

    category: str | None = Field.nested_text(expected_type=str, allow_none=True)
    contentStatus: str | None = Field.nested_text(expected_type=str, allow_none=True)
    keywords: str | None = Field.nested_text(expected_type=str, allow_none=True)
    lastModifiedBy: str | None = Field.nested_text(expected_type=str, allow_none=True)
    lastPrinted: datetime.datetime | None = Field.nested_text(
        expected_type=object,
        allow_none=True,
        converter=_datetime_converter,
        renderer=_datetime_renderer,
    )
    revision: str | None = Field.nested_text(expected_type=str, allow_none=True)
    version: str | None = Field.nested_text(expected_type=str, allow_none=True)
    last_modified_by = AliasField("lastModifiedBy")

    # Dublin Core Properties
    subject: str | None = Field.nested_text(expected_type=str, allow_none=True, namespace=DCORE_NS)
    title: str | None = Field.nested_text(expected_type=str, allow_none=True, namespace=DCORE_NS)
    creator: str | None = Field.nested_text(expected_type=str, allow_none=True, namespace=DCORE_NS)
    description: str | None = Field.nested_text(expected_type=str, allow_none=True, namespace=DCORE_NS)
    identifier: str | None = Field.nested_text(expected_type=str, allow_none=True, namespace=DCORE_NS)
    language: str | None = Field.nested_text(expected_type=str, allow_none=True, namespace=DCORE_NS)
    # Dublin Core Terms
    created: datetime.datetime | None = Field.nested_text(
        expected_type=object,
        allow_none=True,
        namespace=DCTERMS_NS,
        converter=_datetime_converter,
        renderer=_qualified_datetime_renderer,
    )  # assumed UTC
    modified: datetime.datetime | None = Field.nested_text(
        expected_type=object,
        allow_none=True,
        namespace=DCTERMS_NS,
        converter=_datetime_converter,
        renderer=_qualified_datetime_renderer,
    )  # assumed UTC

    xml_order = ("creator", "title", "description", "subject", "identifier",
                 "language", "created", "modified", "lastModifiedBy", "category",
                 "contentStatus", "version", "revision", "keywords", "lastPrinted")


    def __init__(self,
                 category=None,
                 contentStatus=None,
                 keywords=None,
                 lastModifiedBy=None,
                 lastPrinted=None,
                 revision=None,
                 version=None,
                 created=None,
                 creator="fastpyxl",
                 description=None,
                 identifier=None,
                 language=None,
                 modified=None,
                 subject=None,
                 title=None,
                 ):
        now = datetime.datetime.now(tz=datetime.timezone.utc).replace(tzinfo=None)
        self.contentStatus = contentStatus
        self.lastPrinted = lastPrinted
        self.revision = revision
        self.version = version
        self.creator = creator
        self.lastModifiedBy = lastModifiedBy
        self.modified = modified or now
        self.created = created or now
        self.title = title
        self.subject = subject
        self.description = description
        self.identifier = identifier
        self.language = language
        self.keywords = keywords
        self.category = category
