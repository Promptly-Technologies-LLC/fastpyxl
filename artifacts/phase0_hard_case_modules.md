# Phase 0 Hard-Case Module List

## Inventory Snapshot

- Serialisable class count: 388
- Top package distribution:
  - `drawing`: 119
  - `worksheet`: 65
  - `pivot`: 62
  - `chart`: 53
  - `workbook`: 26
  - `styles`: 23
  - `packaging`: 15
  - `chartsheet`: 11
  - `comments`: 4
  - `formatting`: 4
- Hard-pattern descriptor call counts:
  - `Typed`: 599
  - `Alias`: 172
  - `Sequence`: 97
  - `NoneSet`: 72
  - `Set`: 56
  - `NestedSequence`: 50
  - `MultiSequencePart`: 33
  - `MultiSequence`: 4

## Work Package Assignment

| Module | Package | Work package | Notes |
|---|---|---|---|
| `fastpyxl.chart.data_source` | `chart` | `C` | Dynamic number/string value coercion (`#N/A` compatibility). |
| `fastpyxl.chart.plotarea` | `chart` | `A` | Tag-dispatched heterogeneous chart and axis sequences. |
| `fastpyxl.pivot.cache` | `pivot` | `A` | Mixed multi-sequence and deep nested-sequence coverage. |
| `fastpyxl.pivot.record` | `pivot` | `A` | Heterogeneous multi-sequence records. |
| `fastpyxl.styles.styleable` | `styles` | `B` | Style assignment side effects currently hidden in descriptor set. |

## Notes

- Work package keys follow `artifacts/plan.md` (A: multi-sequence, B: style side effects, C: dynamic values, D: ordering parity).
- Work package D classes should be added during migration waves as explicit `__elements__`/`__attrs__` overrides are encountered.
