from __future__ import annotations


def supports_from_tree(expected_type) -> bool:
    return hasattr(expected_type, "from_tree")


def supports_to_tree(value) -> bool:
    return hasattr(value, "to_tree")
