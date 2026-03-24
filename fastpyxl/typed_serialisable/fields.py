from __future__ import annotations

from dataclasses import replace
from typing import Any

from .field_info import FieldInfo


class _FieldSpec:
    def __init__(self, info: FieldInfo):
        self.info = info

    def bind(self, name: str) -> FieldInfo:
        return replace(self.info, name=name)


class _FieldFactory:
    @staticmethod
    def _is_model_type(expected_type: Any) -> bool:
        return hasattr(expected_type, "from_tree") or hasattr(expected_type, "to_tree")

    @staticmethod
    def attribute(
        *,
        expected_type: Any = str,
        allow_none: bool = False,
        default: Any = None,
        xml_name: str | None = None,
        namespace: str | None = None,
        hyphenated: bool = False,
        converter=None,
        validator=None,
        serialize: bool = True,
    ) -> _FieldSpec:
        return _FieldSpec(
            FieldInfo(
                name="",
                kind="attribute",
                expected_type=expected_type,
                allow_none=allow_none,
                default=default,
                xml_name=xml_name,
                namespace=namespace,
                hyphenated=hyphenated,
                converter=converter,
                validator=validator,
                serialize=serialize,
            )
        )

    @staticmethod
    def nested_value(
        *,
        expected_type: Any = str,
        allow_none: bool = False,
        default: Any = None,
        xml_name: str | None = None,
        namespace: str | None = None,
        value_attribute: str = "val",
        serialize: bool = True,
        converter=None,
        parser=None,
        renderer=None,
    ) -> _FieldSpec:
        return _FieldSpec(
            FieldInfo(
                name="",
                kind="nested_value",
                expected_type=expected_type,
                allow_none=allow_none,
                default=default,
                xml_name=xml_name,
                namespace=namespace,
                value_attribute=value_attribute,
                serialize=serialize,
                converter=converter,
                parser=parser,
                renderer=renderer,
            )
        )

    @staticmethod
    def nested_text(
        *,
        expected_type: Any = str,
        allow_none: bool = False,
        default: Any = None,
        xml_name: str | None = None,
        namespace: str | None = None,
        converter=None,
        parser=None,
        renderer=None,
    ) -> _FieldSpec:
        return _FieldSpec(
            FieldInfo(
                name="",
                kind="nested_text",
                expected_type=expected_type,
                allow_none=allow_none,
                default=default,
                xml_name=xml_name,
                namespace=namespace,
                converter=converter,
                parser=parser,
                renderer=renderer,
            )
        )

    @staticmethod
    def nested_bool(
        *,
        allow_none: bool = False,
        default: Any = None,
        xml_name: str | None = None,
        namespace: str | None = None,
        renderer=None,
    ) -> _FieldSpec:
        return _FieldSpec(
            FieldInfo(
                name="",
                kind="nested_bool",
                expected_type=bool,
                allow_none=allow_none,
                default=default,
                xml_name=xml_name,
                namespace=namespace,
                renderer=renderer,
            )
        )

    @staticmethod
    def element(
        *,
        expected_type: Any,
        allow_none: bool = False,
        default: Any = None,
        xml_name: str | None = None,
        namespace: str | None = None,
        converter=None,
        validator=None,
        serialize: bool = True,
    ) -> _FieldSpec:
        return _FieldSpec(
            FieldInfo(
                name="",
                kind="element",
                expected_type=expected_type,
                allow_none=allow_none,
                default=default,
                xml_name=xml_name,
                namespace=namespace,
                converter=converter,
                validator=validator,
                serialize=serialize,
            )
        )

    @staticmethod
    def sequence(
        *,
        expected_type: Any,
        allow_none: bool = False,
        default: Any = list,
        xml_name: str | None = None,
        namespace: str | None = None,
        container_factory=list,
        primitive_attribute: str | None = None,
    ) -> _FieldSpec:
        return _FieldSpec(
            FieldInfo(
                name="",
                kind="sequence",
                expected_type=expected_type,
                allow_none=allow_none,
                default=default,
                xml_name=xml_name,
                namespace=namespace,
                container_factory=container_factory,
                sequence_item_is_model=_FieldFactory._is_model_type(expected_type),
                sequence_primitive_attribute=primitive_attribute,
            )
        )

    @staticmethod
    def nested_sequence(
        *,
        expected_type: Any,
        allow_none: bool = False,
        default: Any = list,
        xml_name: str | None = None,
        namespace: str | None = None,
        count: bool = False,
        container_factory=list,
    ) -> _FieldSpec:
        return _FieldSpec(
            FieldInfo(
                name="",
                kind="nested_sequence",
                expected_type=expected_type,
                allow_none=allow_none,
                default=default,
                xml_name=xml_name,
                namespace=namespace,
                count=count,
                container_factory=container_factory,
                sequence_item_is_model=_FieldFactory._is_model_type(expected_type),
            )
        )

    @staticmethod
    def multi_sequence(
        *,
        parts: dict[str, Any],
        allow_none: bool = False,
        default: Any = list,
        xml_name: str | None = None,
        namespace: str | None = None,
        container_factory=list,
    ) -> _FieldSpec:
        return _FieldSpec(
            FieldInfo(
                name="",
                kind="multi_sequence",
                expected_type=tuple(parts.values()),
                allow_none=allow_none,
                default=default,
                xml_name=xml_name,
                namespace=namespace,
                container_factory=container_factory,
                parts=parts,
            )
        )


def AliasField(target: str, *, xml_name: str | None = None) -> _FieldSpec:
    return _FieldSpec(
        FieldInfo(
            name="",
            kind="alias",
            alias_target=target,
            xml_name=xml_name,
        )
    )


Field = _FieldFactory()
