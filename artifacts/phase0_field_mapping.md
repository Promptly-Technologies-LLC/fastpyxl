# Phase 0 Field Mapping Table

This table captures the descriptor-to-typed-field mapping used to bootstrap Phase 1.

| Legacy descriptor/pattern | Typed runtime target | Notes |
|---|---|---|
| `Typed` / `Convertible` / scalar subclasses | `Field.attribute(...)` + converter hook | Scalar XML attributes in `__attrs__`.
| `NestedValue` (`val` attr) | `Field.nested_value(...)` | Nested node with value on `val` attribute.
| `NestedText` | `Field.nested_text(...)` | Nested node with text content.
| `NestedBool` / `EmptyTag` | `Field.nested_bool(...)` | Support empty-tag truthy semantics and custom renderer hooks.
| `Sequence` | `Field.sequence(...)` | Repeated homogeneous children; container defaults to `list`.
| `NestedSequence` | `Field.nested_sequence(...)` | Wrapped collection with optional `count` attribute.
| `MultiSequence` + `MultiSequencePart` | `Field.multi_sequence(...)` | Tagged heterogeneous child list with tag dispatch.
| `Alias` | `AliasField(target=...)` | No storage; forwards get/set to the target field.
| `Set` / `NoneSet` | `Field.attribute(...)` + validator | Prefer `Literal[...]` when practical, else validator hook.
| `MatchPattern` | `Field.attribute(...)` + validator | Regex validation remains eager.
| Namespaced fields (`namespace=...`) | `Field.*(..., namespace=...)` | Keep namespaced cache entries precomputed.
| Hyphenated names | `Field.*(..., hyphenated=True)` | Preserve underscore<->hyphen conversion parity.

