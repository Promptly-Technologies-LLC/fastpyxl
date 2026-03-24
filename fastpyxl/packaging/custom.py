# Copyright (c) 2010-2024 fastpyxl

"""Implementation of custom properties see § 22.3 in the specification"""

import datetime
from typing import cast

from warnings import warn

from fastpyxl.descriptors import Strict
from fastpyxl.descriptors.sequence import Sequence
from fastpyxl.descriptors import (
    String,
    Integer,
    Float,
    DateTime,
    Bool,
)
from fastpyxl.descriptors.nested import NestedText
from fastpyxl.typed_serialisable.base import Serialisable
from fastpyxl.typed_serialisable.fields import AliasField, Field

from fastpyxl.xml.constants import (
    CUSTPROPS_NS,
    VTYPES_NS,
    CPROPS_FMTID,
)
from fastpyxl.xml.functions import Element

from .core import _datetime_converter


def _filetime_nested_renderer(tagname, value, namespace=None):
    if value is None:
        return None
    if namespace is not None:
        tagname = "{%s}%s" % (namespace, tagname)
    el = Element(tagname)
    el.text = value.replace(tzinfo=None).isoformat(timespec="seconds") + "Z"
    return el


class NestedBoolText(Bool, NestedText):
    """
    Descriptor for handling nested elements with the value stored in the text part
    """

    pass


class _CustomDocumentProperty(Serialisable):

    """
    Low-level representation of a Custom Document Property.
    Not used directly
    Must always contain a child element, even if this is empty
    """

    tagname = "property"
    _typ = None

    name: str | None = Field.attribute(expected_type=str, allow_none=True)
    lpwstr: str | None = Field.nested_text(expected_type=str, allow_none=True, namespace=VTYPES_NS)
    i4: int | None = Field.nested_text(expected_type=int, allow_none=True, namespace=VTYPES_NS)
    r8: float | None = Field.nested_text(expected_type=float, allow_none=True, namespace=VTYPES_NS)
    filetime: datetime.datetime | None = Field.nested_text(
        expected_type=datetime.datetime,
        allow_none=True,
        namespace=VTYPES_NS,
        converter=_datetime_converter,
        renderer=_filetime_nested_renderer,
    )
    bool = Field.nested_text(expected_type=bool, allow_none=True, namespace=VTYPES_NS)
    linkTarget: str | None = Field.attribute(expected_type=str, allow_none=True)
    fmtid: str | None = Field.attribute(expected_type=str, allow_none=True)
    pid: int | None = Field.attribute(expected_type=int, allow_none=True)

    def __init__(self,
                 name=None,
                 pid=0,
                 fmtid=CPROPS_FMTID,
                 linkTarget=None,
                 **kw):
        self.fmtid = fmtid
        self.pid = pid
        self.name = name
        self._typ = None
        self.linkTarget = linkTarget

        for k, v in kw.items():
            setattr(self, k, v)
            setattr(self, "_typ", k) # ugh!
        for e in self.__elements__:
            if e not in kw:
                setattr(self, e, None)


    @property
    def type(self):
        if self._typ is not None:
            return self._typ
        for a in self.__elements__:
            if getattr(self, a) is not None:
                return a
        if self.linkTarget is not None:
            return "linkTarget"


    def to_tree(self, tagname=None, idx=None, namespace=None):
        typ = self._typ
        child = getattr(self, typ, None) if typ is not None else None
        if typ is not None and child is None:
            setattr(self, typ, "")

        return super().to_tree(tagname=None, idx=None, namespace=None)


class _CustomDocumentPropertyList(Serialisable):

    """
    Parses and seriliases property lists but is not used directly
    """

    tagname = "Properties"

    property: list[_CustomDocumentProperty] = Field.sequence(expected_type=_CustomDocumentProperty)
    customProps = AliasField("property")


    def __init__(self, property=()):
        self.property = list(property)


    def __len__(self):
        return len(self.property)


    def to_tree(self, tagname=None, idx=None, namespace=None):
        for idx, p in enumerate(self.property, 2):
            p.pid = idx
        tree = super().to_tree(tagname, idx, namespace)
        tree.set("xmlns", CUSTPROPS_NS)

        return tree


class _TypedProperty(Strict):

    name = String()

    def __init__(self,
                 name,
                 value):
        self.name = name
        self.value = value


    def __eq__(self, other):
        return self.name == other.name and self.value == other.value


    def __repr__(self):
        return f"{self.__class__.__name__}, name={self.name}, value={self.value}"


class IntProperty(_TypedProperty):

    value = Integer()


class FloatProperty(_TypedProperty):

    value = Float()


class StringProperty(_TypedProperty):

    value = String(allow_none=True)


class DateTimeProperty(_TypedProperty):

    value = DateTime()


class BoolProperty(_TypedProperty):

    value = Bool()


class LinkProperty(_TypedProperty):

    value = String()


# from Python
CLASS_MAPPING = {
    StringProperty: "lpwstr",
    IntProperty: "i4",
    FloatProperty: "r8",
    DateTimeProperty: "filetime",
    BoolProperty: "bool",
    LinkProperty: "linkTarget"
}

XML_MAPPING = {v:k for k,v in CLASS_MAPPING.items()}


class CustomPropertyList(Strict):


    props = Sequence(expected_type=_TypedProperty)

    def __init__(self):
        self.props = []


    @classmethod
    def from_tree(cls, tree):
        """
        Create list from OOXML element
        """
        prop_list = _CustomDocumentPropertyList.from_tree(tree)
        props = []

        for prop in prop_list.property:
            attr = prop.type

            typ = XML_MAPPING.get(attr, None)
            if not typ:
                warn(f"Unknown type for {prop.name}")
                continue
            value = getattr(prop, attr)
            link = prop.linkTarget
            if link is not None:
                typ = LinkProperty
                value = prop.linkTarget

            new_prop = typ(name=prop.name, value=value)
            props.append(new_prop)

        new_prop_list = cls()
        new_prop_list.props = props
        return new_prop_list


    def append(self, prop):
        if prop.name in self.names:
            raise ValueError(f"Property with name {prop.name} already exists")

        cast(list, self.props).append(prop)


    def to_tree(self):
        props = []

        for p in cast(list, self.props):
            attr = CLASS_MAPPING.get(p.__class__, None)
            if not attr:
                raise TypeError("Unknown adapter for {p}")
            np = _CustomDocumentProperty(name=p.name, **{attr:p.value})
            if isinstance(p, LinkProperty):
                np._typ = "lpwstr"
                #np.lpwstr = ""
            props.append(np)

        prop_list = _CustomDocumentPropertyList(property=props)
        return prop_list.to_tree()


    def __len__(self):
        return len(cast(list, self.props))


    @property
    def names(self):
        """List of property names"""
        return [p.name for p in cast(list, self.props)]


    def __getitem__(self, name):
        """
        Get property by name
        """
        for p in cast(list, self.props):
            if p.name == name:
                return p
        raise KeyError(f"Property with name {name} not found")


    def __delitem__(self, name):
        """
        Delete a propery by name
        """
        for idx, p in enumerate(cast(list, self.props)):
            if p.name == name:
                cast(list, self.props).pop(idx)
                return
        raise KeyError(f"Property with name {name} not found")


    def __repr__(self):
        return f"{self.__class__.__name__} containing {cast(list, self.props)}"


    def __iter__(self):
        return iter(cast(list, self.props))
