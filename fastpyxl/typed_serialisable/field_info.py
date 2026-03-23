from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable, Literal


FieldKind = Literal[
    "attribute",
    "nested_value",
    "nested_text",
    "nested_bool",
    "element",
    "sequence",
    "nested_sequence",
    "multi_sequence",
    "alias",
]


@dataclass(frozen=True)
class FieldInfo:
    name: str
    kind: FieldKind
    expected_type: Any = object
    allow_none: bool = False
    default: Any = None
    xml_name: str | None = None
    namespace: str | None = None
    hyphenated: bool = False
    count: bool = False
    container_factory: Callable[[], Any] | type | None = list
    validator: Callable[[Any], None] | None = None
    converter: Callable[[Any], Any] | None = None
    parser: Callable[[Any], Any] | None = None
    renderer: Callable[[str, Any, str | None], Any] | None = None
    alias_target: str | None = None
    parts: dict[str, Any] | None = None
    sequence_item_is_model: bool = False

    @property
    def tag(self) -> str:
        return self.xml_name or self.name
