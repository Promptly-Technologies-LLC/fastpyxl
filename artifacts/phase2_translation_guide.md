# Phase 2 Translation Guide

This guide defines the mechanical translation rules from legacy descriptors to the typed runtime.

## Descriptor-to-Field Mapping

| Legacy descriptor/pattern | Typed declaration | Notes |
|---|---|---|
| `Alias("target")` | `alias_name: T = AliasField("target")` | Alias has no independent XML output and no own storage. |
| `Sequence(expected_type=T)` | `values: list[T] = Field.sequence(expected_type=T, default=list)` | Repeated homogeneous child elements. |
| `NestedSequence(expected_type=T, count=True)` | `values: list[T] = Field.nested_sequence(expected_type=T, count=True, default=list)` | Wrapped container with child elements and optional `count`. |
| `Set(values=...)` | `value: str = Field.attribute(..., validator=...)` | Prefer `Literal[...]` when option set is compact. |
| `NoneSet(values=...)` | `value: str \| None = Field.attribute(..., allow_none=True, validator=...)` | Same validation contract, nullable storage. |
| `NumberValueDescriptor` | `value: float \| str \| None = Field.nested_text(expected_type=object, allow_none=True, converter=...)` | Converter preserves sentinel strings like `"#N/A"`, otherwise coerces to float. |
| `MultiSequence` + `MultiSequencePart` | `entries: list[A \| B] = Field.multi_sequence(parts={"a": A, "b": B}, default=list)` | Parse dispatch is driven by child XML tag. |
| Style side-effect descriptor | Keep `Field.*` pure; move registration to explicit writer/workbook API | Assignment stores value only; registration runs explicitly at save-time. |

## Reference Translations

### Alias

```python
class SeriesRef(Serialisable):
    f: str | None = Field.nested_text(expected_type=str, allow_none=True)
    ref: str | None = AliasField("f")
```

### Sequence

```python
class Labels(Serialisable):
    item: list[str] = Field.sequence(expected_type=str, default=list)
```

### NestedSequence

```python
class GroupMembers(Serialisable):
    member: list[Member] = Field.nested_sequence(expected_type=Member, count=True, default=list)
```

### Set / NoneSet

```python
ALLOWED = {"major", "minor", "none"}

def validate_axis_mode(value: str | None) -> None:
    if value is None:
        return
    if value not in ALLOWED:
        raise ValueError(f"axis_mode rejected value {value!r}")


class AxisMode(Serialisable):
    mode: str | None = Field.attribute(
        expected_type=str,
        allow_none=True,
        validator=validate_axis_mode,
    )
```

### NumberValueDescriptor

```python
def convert_number_value(value: object) -> float | str | None:
    if value is None:
        return None
    if value == "#N/A":
        return "#N/A"
    return float(value)


class NumValue(Serialisable):
    v: float | str | None = Field.nested_text(
        expected_type=object,
        allow_none=True,
        converter=convert_number_value,
    )
```

### MultiSequence / MultiSequencePart

```python
class Group(Serialisable):
    parts: list[Axis | Plot] = Field.multi_sequence(
        parts={"axis": Axis, "plot": Plot},
        default=list,
    )
```

### Style Side Effects

```python
class StyledCell(Serialisable):
    font: Font | None = Field.element(expected_type=Font, allow_none=True)


class WorkbookStyleRegistry:
    def register_cell_style(self, cell: StyledCell) -> None:
        # Writer/save layer calls this explicitly.
        ...
```

## Regression Coverage

Executable examples for each hard case are in:

- `fastpyxl/typed_serialisable/tests/test_phase2_translation_playbooks.py`
