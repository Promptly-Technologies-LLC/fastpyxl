from __future__ import annotations

from keyword import kwlist
from typing import Any

KEYWORDS = frozenset(kwlist)


def normalize_attrib(attrib: dict[str, Any]) -> dict[str, Any]:
    for key in list(attrib):
        if key.startswith("{"):
            del attrib[key]
            continue
        if key in KEYWORDS:
            attrib["_" + key] = attrib.pop(key)
            continue
        if "-" in key:
            attrib[key.replace("-", "_")] = attrib.pop(key)
    return attrib


def child_tag(node) -> str:
    tag = node.tag
    if callable(tag):
        return "comment"
    if tag.startswith("{"):
        tag = tag.rsplit("}", 1)[1]
    if tag in KEYWORDS:
        return "_" + tag
    return tag
