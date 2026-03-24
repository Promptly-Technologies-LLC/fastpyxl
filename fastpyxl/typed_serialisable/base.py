from __future__ import annotations

from copy import copy
from keyword import kwlist
from typing import Any

try:
    from typing import dataclass_transform
except ImportError:
    from typing_extensions import dataclass_transform

from fastpyxl.compat import safe_string
from fastpyxl.xml.functions import Element

from .compat import supports_to_tree
from .errors import FieldCoercionError, FieldValidationError
from .field_info import FieldInfo
from .fields import _FieldFactory, _FieldSpec, AliasField
from .parse import child_tag, normalize_attrib
from .render import namespaced_tag, nested_text_node, nested_value_node

KEYWORDS = frozenset(kwlist)
SEQ_TYPES = (list, tuple)

_SERIALISABLE_FIELD_SPECIFIERS = (
    _FieldFactory.attribute,
    _FieldFactory.nested_value,
    _FieldFactory.nested_text,
    _FieldFactory.nested_bool,
    _FieldFactory.element,
    _FieldFactory.sequence,
    _FieldFactory.nested_sequence,
    _FieldFactory.multi_sequence,
    AliasField,
)


@dataclass_transform(field_specifiers=_SERIALISABLE_FIELD_SPECIFIERS)
class MetaSerialisable(type):
    def __new__(mcls, name, bases, namespace):
        annotations = namespace.get("__annotations__", {})
        declared: dict[str, FieldInfo] = {}
        aliases: dict[str, str] = {}
        ordered_names = list(annotations.keys()) + [
            key for key in namespace.keys() if key not in annotations
        ]
        seen: set[str] = set()
        for key in ordered_names:
            if key in seen:
                continue
            seen.add(key)
            value = namespace.get(key)
            if isinstance(value, _FieldSpec):
                info = value.bind(key)
                declared[key] = info
                namespace.pop(key, None)
                if info.kind == "alias" and info.alias_target:
                    aliases[key] = info.alias_target

        cls = super().__new__(mcls, name, bases, namespace)

        base_fields: dict[str, FieldInfo] = {}
        for base in reversed(cls.__mro__[1:]):
            base_fields.update(getattr(base, "__fields__", {}))
        base_fields.update(declared)
        cls.__fields__ = base_fields
        cls.__xml_field_map__ = {}
        cls.__multi_sequence_tag_map__ = {}
        cls.__attribute_xml_name_map__ = {}

        attrs = []
        nested = []
        elements = []
        namespaced = []
        all_aliases = dict(getattr(cls, "__aliases__", {}))
        all_aliases.update(aliases)
        cls.__aliases__ = all_aliases

        for field in cls.__fields__.values():
            if field.kind == "alias":
                target = field.alias_target
                if target:
                    setattr(
                        cls,
                        field.name,
                        property(
                            lambda instance, t=target: getattr(instance, t),
                            lambda instance, value, t=target: setattr(instance, t, value),
                        ),
                    )
                continue
            cls.__xml_field_map__[field.tag] = field
            if field.kind == "attribute":
                cls.__attribute_xml_name_map__[field.tag] = field.name
            if field.kind == "multi_sequence":
                for part_tag in (field.parts or {}):
                    cls.__multi_sequence_tag_map__[part_tag] = field

            if field.namespace:
                namespaced.append((field.name, namespaced_tag(field.tag, field.namespace)))
            if field.kind == "attribute":
                if field.serialize:
                    attrs.append(field.name)
            elif field.kind in {"nested_value", "nested_text", "nested_bool"}:
                if field.serialize:
                    nested.append(field.name)
                    elements.append(field.name)
            else:
                if field.serialize:
                    elements.append(field.name)

        cls.__attrs__ = tuple(attrs)
        cls.__nested__ = tuple(nested)
        override = namespace.get("xml_order")
        if override is not None:
            ordered = [el for el in override if el in elements]
            remaining = [el for el in elements if el not in ordered]
            cls.__elements__ = tuple(ordered + remaining)
        else:
            cls.__elements__ = tuple(elements)
        cls.__namespaced__ = tuple(namespaced)

        if "__init__" not in namespace:
            cls.__init__ = _build_init(cls.__fields__)

        return cls


def _build_init(fields: dict[str, FieldInfo]):
    # Precompute field order and skip aliases to reduce per-instance overhead.
    init_fields: list[tuple[str, FieldInfo]] = [
        (name, field) for name, field in fields.items() if field.kind != "alias"
    ]

    def __init__(self, **kwargs):
        # Avoid routing through Serialisable.__setattr__ for construction:
        # __setattr__ does field lookups on every call; here we already have the FieldInfo.
        for name, field in init_fields:
            if name in kwargs:
                value = kwargs[name]
            else:
                default = field.default
                if callable(default) and default in (list, dict, set, tuple):
                    default = default()
                value = default

            value = _coerce(field, value)
            object.__setattr__(self, name, value)

    return __init__


def _convert_scalar(expected_type: Any, value: Any):
    if value is None:
        return None
    if expected_type is Any or expected_type is object:
        return value
    if isinstance(value, expected_type):
        return value
    if expected_type is bool and isinstance(value, str):
        if value.lower() in {"0", "false", "f"}:
            return False
        if value.lower() in {"1", "true", "t"}:
            return True
    try:
        return expected_type(value)
    except Exception as exc:
        raise FieldCoercionError(f"expected {expected_type} for value {value!r}") from exc


def _coerce(field: FieldInfo, value: Any):
    if value is None:
        if field.allow_none or field.default is None:
            return None
        raise FieldValidationError(f"{field.name} rejected value {value!r}")

    if field.converter is not None:
        value = field.converter(value)

    if field.kind in {"attribute", "nested_value", "nested_text", "nested_bool"}:
        value = _convert_scalar(field.expected_type, value)
    elif field.kind == "element":
        if not isinstance(value, field.expected_type):
            raise FieldValidationError(f"{field.name} rejected value {value!r}")
    elif field.kind in {"sequence", "nested_sequence", "multi_sequence"}:
        if not isinstance(value, SEQ_TYPES):
            raise TypeError(f"{field.name} expected a sequence but got {type(value)}")
        converted = []
        for item in value:
            if field.kind == "multi_sequence":
                parts = tuple((field.parts or {}).values())
                if parts and not isinstance(item, parts):
                    raise FieldValidationError(f"{field.name} rejected value {item!r}")
                converted.append(item)
            elif field.sequence_item_is_model:
                if not isinstance(item, field.expected_type):
                    raise FieldValidationError(f"{field.name} rejected value {item!r}")
                converted.append(item)
            else:
                converted.append(_convert_scalar(field.expected_type, item))
        container = field.container_factory or list
        value = container(converted)

    if field.validator is not None:
        field.validator(value)
    return value


class Serialisable(metaclass=MetaSerialisable):
    __fields__ = {}
    __attrs__ = ()
    __nested__ = ()
    __elements__ = ()
    __namespaced__ = ()
    __aliases__ = {}
    __xml_field_map__ = {}
    __multi_sequence_tag_map__ = {}
    __attribute_xml_name_map__ = {}
    namespace = None

    @property
    def tagname(self):
        raise NotImplementedError

    def __setattr__(self, name, value):
        field = self.__fields__.get(name)
        if field is not None and field.kind != "alias":
            value = _coerce(field, value)
        super().__setattr__(name, value)

    @classmethod
    def from_tree(cls, node):
        attrib: Any = dict(node.attrib)
        for key, ns_key in cls.__namespaced__:
            if ns_key in attrib:
                attrib[key] = attrib.pop(ns_key)
        attrib = normalize_attrib(attrib)
        for xml_name, py_name in cls.__attribute_xml_name_map__.items():
            if xml_name in attrib and py_name not in attrib:
                attrib[py_name] = attrib.pop(xml_name)
        if node.text and "attr_text" in cls.__attrs__:
            attrib["attr_text"] = node.text

        for child in node:
            tag = child_tag(child)
            field = cls.__xml_field_map__.get(tag) or cls.__fields__.get(tag)
            if field is None:
                field = cls.__multi_sequence_tag_map__.get(tag)
            if field is None:
                continue

            if field.kind == "nested_value":
                obj = child.get(field.value_attribute)
            elif field.kind == "nested_text":
                obj = child.text
            elif field.kind == "nested_bool":
                obj = child.get("val", True)
            elif field.kind == "element":
                obj = field.expected_type.from_tree(child)
            elif field.kind == "sequence":
                if field.sequence_item_is_model:
                    obj = field.expected_type.from_tree(child)
                elif field.sequence_primitive_attribute:
                    obj = child.get(field.sequence_primitive_attribute)
                    if obj is not None and not field.sequence_item_is_model:
                        obj = _convert_scalar(field.expected_type, obj)
                else:
                    obj = child.text
                attrib.setdefault(field.name, [])
                attrib[field.name].append(obj)
                continue
            elif field.kind == "nested_sequence":
                obj = [field.expected_type.from_tree(el) for el in child]
            elif field.kind == "multi_sequence":
                parts = field.parts or {}
                model = parts.get(tag)
                if model is None:
                    continue
                obj = model.from_tree(child)
                attrib.setdefault(field.name, [])
                attrib[field.name].append(obj)
                continue
            else:
                continue

            attrib[field.name] = obj
        return cls(**attrib)

    def to_tree(self, tagname=None, idx=None, namespace=None):
        del idx
        if tagname is None:
            tagname = self.tagname
        if tagname.startswith("_"):
            tagname = tagname[1:]

        namespace = getattr(self, "namespace", namespace)
        root = Element(namespaced_tag(tagname, namespace))
        attrs = dict(self)
        for key, ns_key in self.__namespaced__:
            field = self.__fields__.get(key)
            if field is None or field.kind != "attribute":
                continue
            tag = field.tag
            if tag in attrs:
                attrs[ns_key] = attrs.pop(tag)
        for key, value in attrs.items():
            root.set(key, value)

        for name in self.__elements__:
            field = self.__fields__[name]
            value = getattr(self, name)
            tag = field.tag
            if field.kind == "nested_value":
                if field.renderer is not None:
                    node = field.renderer(tag, value, field.namespace or namespace)
                else:
                    node = nested_value_node(
                        tag, value, field.namespace or namespace, field.value_attribute
                    )
                if node is not None:
                    root.append(node)
            elif field.kind == "nested_text":
                if field.renderer is not None:
                    node = field.renderer(tag, value, field.namespace or namespace)
                else:
                    node = nested_text_node(tag, value, field.namespace or namespace)
                if node is not None:
                    root.append(node)
            elif field.kind == "nested_bool":
                if field.renderer is not None:
                    node = field.renderer(tag, value, field.namespace or namespace)
                else:
                    node = nested_value_node(tag, bool(value), field.namespace or namespace) if value is not None else None
                if node is not None:
                    root.append(node)
            elif field.kind == "element":
                if value is not None:
                    node = value.to_tree(tag)
                    if node is not None:
                        root.append(node)
            elif field.kind == "sequence":
                if value is None:
                    continue
                pattr = field.sequence_primitive_attribute
                idx_base = getattr(self, "idx_base", 0)
                for idx, item in enumerate(value, idx_base):
                    if supports_to_tree(item):
                        root.append(item.to_tree(tag, idx))
                    else:
                        el = Element(namespaced_tag(tag, field.namespace or namespace))
                        if pattr is not None:
                            el.set(pattr, safe_string(item))
                        else:
                            el.text = safe_string(item)
                        root.append(el)
            elif field.kind == "nested_sequence":
                if not value:
                    continue
                container = Element(namespaced_tag(tag, field.namespace or namespace))
                if field.count:
                    container.set("count", str(len(value)))
                for item in value:
                    container.append(item.to_tree())
                root.append(container)
            elif field.kind == "multi_sequence":
                if value is None:
                    continue
                for item in value:
                    root.append(item.to_tree())
        return root

    def __iter__(self):
        for attr in self.__attrs__:
            field = self.__fields__[attr]
            value = getattr(self, attr)
            if value is None:
                continue
            xml_attr = field.tag
            if xml_attr.startswith("_"):
                xml_attr = xml_attr[1:]
            if field.hyphenated:
                xml_attr = xml_attr.replace("_", "-")
            yield xml_attr, safe_string(value)

    def __eq__(self, other):
        if not self.__class__ == other.__class__:
            return False
        if dict(self) != dict(other):
            return False
        for el in self.__elements__:
            if getattr(self, el) != getattr(other, el):
                return False
        return True

    def __hash__(self):
        vals = []
        for attr in self.__attrs__ + self.__elements__:
            v = getattr(self, attr, None)
            if isinstance(v, list):
                v = tuple(v)
            vals.append(v)
        return hash(tuple(vals))

    def __add__(self, other):
        if type(self) is not type(other):
            raise TypeError("Cannot combine instances of different types")
        vals = {}
        for attr in self.__attrs__:
            vals[attr] = getattr(self, attr) or getattr(other, attr)
        for el in self.__elements__:
            a = getattr(self, el)
            b = getattr(other, el)
            vals[el] = (a + b) if (a and b) else (a or b)
        return self.__class__(**vals)

    def __copy__(self):
        xml = self.to_tree(tagname="dummy")
        cp = self.__class__.from_tree(xml)
        for k in self.__dict__:
            if k not in self.__attrs__ + self.__elements__:
                setattr(cp, k, copy(getattr(self, k)))
        return cp
