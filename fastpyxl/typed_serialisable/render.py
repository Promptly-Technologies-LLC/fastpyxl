from __future__ import annotations

from fastpyxl.compat import safe_string
from fastpyxl.xml.functions import Element, whitespace


def namespaced_tag(tag: str, namespace: str | None) -> str:
    if namespace is None:
        return tag
    return f"{{{namespace}}}{tag}"


def nested_value_node(tag: str, value, namespace: str | None = None, value_attribute: str = "val"):
    if value is None:
        return None
    return Element(namespaced_tag(tag, namespace), {value_attribute: safe_string(value)})


def nested_text_node(tag: str, value, namespace: str | None = None):
    if value is None:
        return None
    el = Element(namespaced_tag(tag, namespace))
    el.text = safe_string(value)
    whitespace(el)
    return el
