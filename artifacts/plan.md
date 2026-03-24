# Typed Serialisable Rewrite: Execution Plan

## Decision

Rewrite the `Serialisable` system in-house around typed field metadata.

Do not adopt `pydantic-xml`.

Rationale:
- preserve a library-appropriate dependency footprint
- preserve fast direct attribute reads for normal fields
- model OOXML-specific field semantics directly instead of fitting them into a generic XML library
- make the runtime architecture match the type-checking story instead of layering stubs or generic descriptors on top of the old system

## Success Criteria

The rewrite is complete when all of the following are true:
- migrated classes expose real public attribute types via annotations
- migrated classes preserve existing XML read/write behavior
- migrated classes pass golden round-trip tests and the normal test suite
- hot-path `to_tree()` and `from_tree()` throughput are within 10-15% of baseline
- descriptor-driven false positives are gone for migrated packages
- the legacy descriptor/metaclass stack can be removed cleanly

## TDD Rules

This project should be executed as strict test-driven development, not just test-backed migration.

### Global rule

No production code change is allowed unless it is preceded by at least one failing test that expresses the behavior being added or preserved.

### Required loop

Every task follows:

1. write or extend a failing test
2. implement the smallest change that makes it pass
3. refactor while keeping the test suite green

### Test layers

Use four test layers deliberately:

1. unit tests
- field semantics
- parse/render helpers
- validation/coercion behavior

2. golden tests
- XML round-trip stability
- old/new serializer equivalence for migrated classes

3. integration tests
- workbook load/save behavior
- reader/writer interactions
- cross-module object graphs

4. performance tests
- parse/serialize throughput
- object construction and assignment

### Pull-request rule

Every meaningful change set should say which failing tests were written first and which layer they belong to.

## Fixed Design Decisions

These are no longer open questions for planning purposes.

### Validation model

Use eager validation and coercion.

- `__init__`, `from_tree()`, and normal attribute assignment all validate
- validation lives in `Serialisable.__setattr__` plus field-specific converters/validators
- this preserves current behavior better than init-only validation

### Storage model

Use plain instance storage in `__dict__` for v1.

- normal field reads must stay direct attribute reads
- no `__slots__` in the first implementation
- revisit `__slots__` only after benchmarks, not before

### Alias model

Implement aliases as generated properties plus metadata.

- alias fields do not own storage
- alias fields do not emit XML independently
- alias access forwards to the target field
- alias class access must be well-defined

### Ordering model

Use declaration order as the default XML order, with explicit override metadata where needed.

- preserve an `xml_order=` escape hatch for classes that currently rely on custom `__elements__`
- do not require every class to restate ordering manually

### Sequence containers

Use `list[T]` as the default runtime container, with per-field container overrides where public behavior requires something else.

- `list` for normal sequences
- configurable container factory for set/tuple-like cases
- preserve uniqueness behavior explicitly, not implicitly

### Dynamic field cases

Represent dynamic value fields as typed unions plus custom parser/converter hooks.

- example: `NumberValueDescriptor` becomes `float | str | None` with a custom converter
- do not allow runtime mutation of field metadata like `expected_type`

### Old/new interop contract

During the coexistence period (Phases 3–6), migrated and unmigrated classes must interoperate seamlessly. The rules:

- **New class referencing old class:** A new-style `Field.element(expected_type=OldClass)` must work. The parse pipeline calls `OldClass.from_tree(node)` and the render pipeline calls `old_instance.to_tree()` — both still exist on legacy classes. The field stores the raw `OldClass` instance as its runtime value. Any compatibility adapters are confined to parse/render boundaries and helper utilities; they are not the stored public value type.
- **Old class referencing new class:** A legacy `Typed(expected_type=NewClass)` must work. New-style classes must expose `from_tree(cls, node)` and `to_tree(self, tagname, idx, namespace)` with signatures matching the legacy `Serialisable` interface. The new base class provides these as thin wrappers around the native parse/render pipeline.
- **Duck-typing bridge:** Both old and new classes must support `__iter__` (for the attribute-dict pattern used in `to_tree`), `__eq__`, `__hash__`, and `__copy__` with compatible semantics. The new base class implements these in terms of `__fields__`.
- **Import stability:** During migration, existing import paths (e.g., `from fastpyxl.styles.fonts import Font`) continue to resolve to the class regardless of whether it has been migrated. No re-export shims — the class stays in its original module.
- **Identity rule:** Mixed old/new interop must preserve object identity and `isinstance` expectations for the stored model values. A field annotated as `OldClass` stores an `OldClass`; a field annotated as `NewClass` stores a `NewClass`.

This contract must be validated in Phase 1 with at least one synthetic test: a new-style parent model containing an old-style child, and an old-style parent containing a new-style child.

### `__init__` synthesis rules

Most existing `Serialisable` subclasses (~88%) define explicit `__init__` methods.

- **Classes with explicit `__init__`:** Leave the `__init__` untouched during migration. The metaclass does not generate or replace it. Validation on assignment via `__setattr__` ensures field semantics are enforced regardless of how construction happens.
- **Classes without explicit `__init__`:** The metaclass synthesizes an `__init__` from `__fields__`, accepting all declared fields as keyword arguments with their declared defaults. This replaces the implicit behavior the old metaclass provided.
- **"Non-trivial" defined:** A constructor is non-trivial if it does more than straight `self.x = x` assignment — e.g., computed defaults, conditional logic, alias redirection, or type-selection (like `Color`). These must be preserved as-is and are not candidates for synthesis.
- **Post-migration cleanup:** After a full module is migrated, constructors that are purely mechanical `self.x = x` for every field may optionally be removed in a separate refactoring pass, letting the synthesized `__init__` take over. This is never done in the same change as the migration itself.

### Error contract

Validation errors must preserve compatibility with existing error types and messages where downstream code or tests depend on them.

- Field validation raises `ValueError` or `TypeError` consistent with the current descriptor behavior.
- The error message must include the field name and the rejected value.
- `errors.py` defines a small set of concrete exception classes (`FieldValidationError`, `FieldCoercionError`, `ParseError`, `RenderError`) that all subclass the standard types (`ValueError`, `TypeError`) so existing `except ValueError` handlers continue to work.
- New-style errors may include additional context (field kind, expected type) but must not change the exception type hierarchy.

### Thread safety

`__fields__` and all class-level caches (`__attrs__`, `__elements__`, etc.) are immutable after class creation. They are safe to read from multiple threads without synchronization. Instance-level `__dict__` mutation is not thread-safe, consistent with standard Python object behavior — no additional guarantees are provided or needed.

## Target Architecture

### Core primitives

Create a new package, tentatively `fastpyxl/typed_serialisable/`, with:

- `field_info.py`
  - immutable `FieldInfo`
  - field kind enum/literals
  - validator/converter/parser/renderer hooks
- `fields.py`
  - `Field.*` factories
  - `AliasField`
- `base.py`
  - new `MetaSerialisable`
  - new `Serialisable`
- `parse.py`
  - XML -> object pipeline
- `render.py`
  - object -> XML pipeline
- `compat.py`
  - helper adapters for old/new interop during migration
- `errors.py`
  - validation and parse/render errors

### Canonical field kinds

The runtime must support these first-class field strategies:

1. attribute field
2. nested `val` field
3. nested text field
4. serialisable element field
5. repeated sequence field
6. wrapped nested sequence field
7. heterogeneous multi-sequence field
8. alias field

### Field declaration style

Use annotations as the source of public types and field objects as the source of XML semantics.

Preferred style:

```python
class Font(Serialisable):
    name: str | None = Field.nested_text()
    sz: float | None = Field.nested_float()
    b: bool | None = Field.nested_bool(renderer=_no_value)
    bold: bool | None = AliasField("b")
```

Avoid a weak boolean API like `Field(nested=True)` as the long-term public surface.

### Required runtime caches

At class creation, compute and cache:

- `__fields__`
- `__attrs__`
- `__nested__`
- `__elements__`
- `__namespaced__`
- `__aliases__`

These remain part of the runtime because the serializer benefits from precomputed layout.

## Golden Test Strategy

This is the first execution phase, not an optional later step.

### Deliverables

- `scripts/generate_serialisable_goldens.py`
- `tests/golden/` XML fixtures for migrated packages
- `tests/test_serialisable_goldens.py`

### Generation strategy

Generate two kinds of coverage:

1. Class round-trip goldens
- discover `Serialisable` subclasses recursively
- instantiate classes from existing defaults when possible
- serialize -> deserialize -> assert equality
- store the serialized XML as the golden for stable classes

2. Existing XML fixture goldens
- scan existing `tests/data/` XML fixtures
- map fixtures to target model classes where possible
- assert `from_tree(fixture) -> to_tree()` remains stable

### Rules

- do not block the rewrite on perfect auto-instantiation coverage for all 393 classes
- classify goldens into:
  - `auto_roundtrip`
  - `fixture_roundtrip`
  - `manual_required`
- hard classes can enter later waves with targeted manual goldens

### Exit gate

Before any migration work starts:
- golden test generator exists
- at least the simple/leaf classes have generated coverage
- every migration wave below has explicit golden coverage for its target modules

## Test Matrix

The rewrite needs explicit test targets per runtime feature.

### Field strategy unit tests

Add unit tests for each field kind before using it in migrated production models:

1. attribute field
2. nested `val` field
3. nested text field
4. serialisable child field
5. repeated sequence field
6. wrapped nested sequence field
7. heterogeneous multi-sequence field
8. alias field

Each field strategy test suite must cover:
- construction
- normal assignment
- invalid assignment
- parse from XML
- render to XML
- `None` handling
- ordering and namespace behavior where relevant

### Cross-cutting unit tests

Add dedicated unit tests for:
- keyword normalization
- hyphenated XML names
- namespace mapping
- explicit ordering overrides
- validator and converter hooks
- equality, hashing, copying, and addition semantics where preserved

### Migration test requirement

No class is eligible for migration until it has:
- targeted unit tests for the field strategies it uses
- golden coverage for its XML behavior
- integration coverage if it participates in workbook load/save flows

## Execution Phases

## Phase 0: Inventory and Baseline

### Deliverables

- finalized field mapping table
- baseline parse/serialize benchmarks
- list of hard-case modules with notes on which work package each belongs to

### Required inventory

From the current codebase scan:
- ~393 `Serialisable` classes
- package distribution:
  - `drawing`: 119
  - `worksheet`: 65
  - `pivot`: 62
  - `chart`: 53
  - `workbook`: 26
  - `styles`: 23
- hard-pattern counts:
  - `Typed`: 601
  - `Alias`: 172
  - `Sequence`: 100
  - `NoneSet`: 73
  - `Set`: 58
  - `NestedSequence`: 51

### Exit gate

- baseline benchmark script exists
- field mapping doc exists
- hard-case list exists

## Phase 1: Infrastructure Build

Build the new typed runtime before migrating any real module.

### Work items

1. Implement `FieldInfo`
2. Implement `Field.*` factories
3. Implement alias handling
4. Implement class creation pipeline
5. Implement `__setattr__` validation/coercion
6. Implement parse/render pipeline
7. Implement compatibility helpers for old/new coexistence

### TDD sequence

Build infrastructure in this exact order:

1. write failing tests for `FieldInfo` compilation and class creation
2. implement the minimum metadata model
3. write failing tests for scalar attribute fields
4. implement scalar field assignment + render/parse
5. repeat for each remaining field kind one at a time
6. write failing tests for alias forwarding
7. implement alias behavior
8. write failing tests for namespace, keyword, and ordering logic
9. implement serializer edge-case handling
10. refactor after all infrastructure tests are green

### Minimum unit-test targets

- scalar attribute field
- nested `val` field
- nested text field
- serialisable child field
- sequence field
- nested sequence field
- alias field
- namespace + hyphenated + keyword normalization

### Exit gate

- a small synthetic test model suite passes
- old and new implementations can coexist in the same runtime
- the interop contract smoke test passes for both old-parent/new-child and new-parent/old-child cases
- benchmark on synthetic models is within target range
- every infrastructure feature was introduced via a failing test first

## Phase 2: Translation Guide and Hard-Case Playbooks

Produce the mechanical translation rules before bulk migration.

### Deliverables

- descriptor-to-field mapping table
- one reference translation example per field kind
- playbooks for hard cases

### Required playbooks

1. `Alias`
- source field + generated property

2. `Sequence`
- `list[T]` plus element parser/renderer

3. `NestedSequence`
- wrapped element with child item renderer

4. `Set` and `NoneSet`
- `Literal[...] | None` or validator-backed `str | None`
- use validators when literal unions become too large or awkward

5. `NumberValueDescriptor`
- `float | str | None` with custom converter preserving `"#N/A"`

6. `MultiSequence` / `MultiSequencePart`
- tagged heterogeneous child list
- parser dispatches by XML tag

7. style side-effect descriptors
- separate serialization model from workbook registration logic

### Exit gate

- translation guide exists in the repo
- at least one committed example exists for each hard case
- each example is backed by a failing-then-passing regression test

## Phase 3: Pilot Migration

Migrate a small but representative set before attempting bulk conversion.

### Pilot targets

Use four deliberately different classes spanning four packages:

1. `fastpyxl/styles/fonts.py:Font`
- nested primitives
- aliases
- custom renderer

2. `fastpyxl/chart/data_source.py:NumVal`
- dynamic value field replacement

3. `fastpyxl/worksheet/page.py:PageMargins`
- simple leaf class in the worksheet package
- exercises the integration test surface early (workbook load/save)

4. `fastpyxl/pivot/cache.py:LevelGroup`
- nested sequence (`groupMembers`) with `count=True`
- exercises the wrapped sequence pattern concretely

### Per-class process

1. add or update a failing golden or regression test for the class
2. add targeted failing unit tests for any field strategies or custom hooks the class relies on
3. translate descriptors to typed fields
4. preserve public constructor behavior
5. preserve XML output order
6. get unit tests green
7. get goldens green
8. run relevant integration tests
9. benchmark if the class sits on a hot path
10. refactor only after all tests are green

### Exit gate

- all pilot classes are migrated
- no unresolved design blockers remain
- mechanical translation recipe is validated
- each pilot migration has a documented red/green/refactor history

## Phase 4: Leaf and Low-Complexity Waves

Bulk migration starts only after the pilot passes.

**Phase 4 closeout (done):** Foundation, styles core, and the drawing *spine* are largely on `fastpyxl.typed_serialisable`. Some files are **hybrid** (typed and legacy `Serialisable` subclasses in the same module). A slice of work originally listed under Phase 5A (drawing composites) was completed while unblocking primitives and integration tests.

### Wave 4A: Foundation models — status

**Delivered**

- `fastpyxl/packaging/` — typed migration for core document properties, relationships, manifest, extended/core metadata, custom property XML layer; hybrid where noted below.
- `fastpyxl/workbook/` — simple workbook XML models (properties, views, protection, defined names, external link, smart tags, function groups, etc.) on typed `Serialisable`.
- `fastpyxl/chartsheet/` — chartsheet-facing models migrated; **no** remaining `from fastpyxl.descriptors.serialisable import Serialisable` in `chartsheet/` as of Phase 4 closeout.

**Still open (carry to Phase 5C or later)**

- `fastpyxl/packaging/workbook.py` — helper types (`ChildSheet`, `PivotCache`, `FileRecoveryProperties`, …) are typed; **`WorkbookPackage` remains legacy `Serialisable`** (root package type).
- `fastpyxl/descriptors/excel.py` — still the home of **descriptor** helpers (`Relation`, etc.); not a “typed Serialisable migration” target in the same sense as leaf models. Revisit only if the plan later inlines or replaces these.

### Wave 4B: Styles core — status

**Delivered**

- `styles/alignment.py`, `protection.py`, `numbers.py`, `colors.py`, `borders.py`, `fills.py`, `fonts.py`, `table.py`, `differential.py`, `cell_style.py` — migrated to typed fields.

**Explicitly deferred (unchanged)**

- `styles/stylesheet.py` — still legacy; depends on workbook/formatting/cell integration (Phase 5C).
- `styles/named_styles.py` — still legacy (Phase 5C / adjacent).

### Wave 4C: Drawing primitives — status

**Delivered (module-level typed migration)**

- `drawing/fill.py`, `drawing/line.py`, `drawing/properties.py`, `drawing/relation.py`, `drawing/xdr.py` — on typed `Serialisable` (with `xml_order` / validators where needed).

**Hybrid: high-value types typed, remainder still legacy `Serialisable` in-file**

- `drawing/geometry.py` — **typed:** `Point2D`, `PositiveSize2D`, `Transform2D`, `GroupTransform2D`. **Still legacy:** 3D/path/preset geometry, `ShapeStyle`, `FontReference`, and related types (large tail). *Schedule: continue as Phase 5A/5B prep or a dedicated “drawing tail” pass.*
- `drawing/colors.py` — **typed:** `HSLColor`, `RGBPercent`, `ColorChoice`. **Still legacy:** e.g. `Transform`, `SystemColor`, `SchemeColor`, `ColorMapping`, …
- `drawing/effect.py` — **typed:** most effect list / shadow / blur stack used by charts and fills. **Still legacy:** e.g. `HSLEffect`, `FillOverlayEffect`, `DuotoneEffect`, `ColorReplaceEffect`, `Color`, several alpha helper effects.

**Completed early (was Phase 5A on the old plan)**

To avoid a compatibility layer and to fix class-level descriptor access, these were migrated during Phase 4:

- `drawing/graphic.py`, `drawing/connector.py`, `drawing/picture.py`, `drawing/spreadsheet_drawing.py`

### Phase 4 tests added

- `typed_serialisable/tests/test_phase4_styles_core_migration.py`
- `typed_serialisable/tests/test_phase4_foundation_models_migration.py`
- `typed_serialisable/tests/test_phase4_drawing_primitives_migration.py`
- (pilot / Phase 3 adjacent) `test_phase3_*`, `test_element_converter_support.py`

### Exit gate per wave (Phase 4)

- Targeted regression tests and smoke checks for migrated classes pass.
- Full “all goldens / all package tests” remains the **Phase 5–6** bar until CI is run with pytest on the full tree.

---

## Phase 5: Interior Model Waves

These waves depend on migrated foundations. **Several drawing composite files are already done;** the lists below are **what remains**, not the original aspirational targets.

### Wave 5A: Drawing composites — remaining

**Delivered (Phase 5 closeout)**

- `drawing/text.py` — fully on `fastpyxl.typed_serialisable.base.Serialisable` (no remaining legacy `Serialisable` subclasses in-file; `OfficeArtExtensionList` remains a legacy *referenced* type for interop). Chart `RichText` / `Text` continue to compose these models via legacy `Typed` fields.

**Already completed (during Phase 4)**

- `drawing/graphic.py`, `drawing/connector.py`, `drawing/picture.py`, `drawing/spreadsheet_drawing.py`, `drawing/relation.py`, `drawing/xdr.py` (coordinate with `drawing/text.py` when migrating rich text).

### Wave 5B: Chart models

**Delivered**

- `chart/*` production modules use `fastpyxl.typed_serialisable.base.Serialisable` with `Field` / `AliasField`; legacy `ExtensionList` and related types from `fastpyxl.descriptors.excel` remain referenced where needed for interop.
- Regression coverage: `fastpyxl/chart/tests/`, `fastpyxl/typed_serialisable/tests/test_phase5b_chart_migration.py`.

**Chartsheet**

- `chartsheet/*` — **migrated in Phase 4**; further chart-*sheet* work is mostly chart package integration, not duplicating chartsheet XML models.

### Wave 5C: Workbook / worksheet / comments / formatting / cell / stylesheet

**Delivered**

- `worksheet/*` production modules (including previously hybrid `worksheet/page.py`) now use `fastpyxl.typed_serialisable.base.Serialisable` with `Field` / `AliasField`; non-model utility classes remain unchanged where appropriate.
- `comments/*`, `formatting/*`, and `cell/text.py` migrated to typed `Serialisable`, including behavior-parity handling for value coercion, XML namespace attributes, and selective serialization (`serialize=False`) where needed.
- `styles/named_styles.py` and `styles/stylesheet.py` migrated to typed `Serialisable` while preserving style binding and workbook integration behavior.
- `packaging/workbook.py` `WorkbookPackage` and nested workbook models migrated to typed `Serialisable` with relationship namespace parity.
- Regression coverage expanded with `typed_serialisable/tests/test_phase5c_workbook_worksheet_migration.py`, plus targeted suites across worksheet/styles/formatting/comments/cell.

### Exit gate per wave

- targeted module suite passes
- workbook load/save integration tests pass
- representative real-file fixtures still round-trip
- newly discovered behaviors are captured first in failing tests, not post hoc fixes

## Phase 6: Complex Tail

Leave the hardest modules for last.

**Note:** `pivot/cache.py` may already contain **pilot / partial** typed migrations (e.g. nested-sequence patterns); treat the file as **hybrid** until a full pass removes legacy `Serialisable` subclasses.

### Primary targets

- `pivot/cache.py` (complete any remaining legacy classes)
- `pivot/table.py`
- `pivot/record.py`
- any remaining `MultiSequence` users
- style side-effect consumers in `styles/styleable.py`

### Required tactics

- isolate heterogeneous list parsing behind dedicated field strategies
- split style registration side effects out of serialization code
- add manual golden fixtures before translating each complex module
- add bespoke failing regression tests for each hard case before implementing its runtime support

### Exit gate

- all pivot tests pass
- all remaining goldens pass
- no remaining legacy-only blockers exist

## Phase 7: Remove Legacy Infrastructure

Remove the old descriptor system only after all production modules are migrated.

### Delete or deprecate

- `descriptors/base.py`
- `descriptors/nested.py`
- `descriptors/sequence.py`
- `descriptors/serialisable.py`
- legacy metaclass logic in `descriptors/__init__.py`

### Keep temporarily if needed

- compatibility shims for import stability

### Exit gate

- no production code imports legacy descriptors
- tests pass without legacy runtime dependencies

## Phase 8: Tighten Static Checking

Remove `ty` suppressions gradually, not all at once.

### Removal order

1. `invalid-assignment`
2. `unresolved-attribute`
3. `invalid-argument-type`
4. `unsupported-operator`
5. `not-iterable`
6. `not-subscriptable`
7. `call-non-callable`
8. `no-matching-overload`
9. `index-out-of-bounds`
10. `invalid-method-override`

Keep `unresolved-import` and legacy-related suppressions separate until optional dependency typing is handled.

### Exit gate

- each rule is enabled only after the current tree is clean
- no broad descriptor-related ignores remain
- any diagnostics fixed during rule enablement are accompanied by tests when behavior changes

## Mechanical Per-Class Recipe

Use this exact checklist during bulk migration.

1. Write or update a failing golden test for the class.
2. Write failing unit tests for any class-specific parsing, rendering, aliasing, or validation behavior.
3. Translate descriptors to annotations + `Field.*` declarations.
4. Replace aliases with `AliasField` declarations.
5. Preserve the existing constructor signature if public and non-trivial.
6. Preserve explicit XML name, namespace, and order.
7. Add custom parser/renderer hooks only where the translation guide requires them.
8. Get targeted unit tests green.
9. Get golden tests green.
10. Run focused integration tests for the module.
11. If the class is on a hot path, run targeted benchmarks.
12. Refactor only after the test pyramid is green.
13. Remove unused descriptor imports from the file.

Do not redesign public constructors during migration unless there is a compelling correctness reason.

If a class cannot be migrated test-first because behavior is still unclear, stop and capture that behavior with a characterization test before changing implementation.

## Hard-Case Work Packages

These need explicit ownership because they can block whole waves.

### Work package A: Heterogeneous sequences

Targets:
- `pivot/record.py`
- `pivot/cache.py`
- `chart/plotarea.py`

Output:
- reusable multi-sequence field strategy
- tests proving tag-based parse/render equivalence

### Work package B: Style side effects

Targets:
- `styles/styleable.py`
- callers that rely on implicit workbook registry mutation

The core problem: `StyleDescriptor.__set__` currently registers styles in the workbook collection as a side effect of attribute assignment (e.g., `cell.font = Font(bold=True)` silently adds the font to the workbook's font table). This conflates serialization with application logic.

Migration path:
- **Step 1:** Extract the registration logic from `StyleDescriptor.__set__` into explicit methods on the workbook/worksheet layer (e.g., `workbook._register_style(cell)`).
- **Step 2:** During migration, the new-style field assignment stores the value without side effects. The registration call moves to the save path — the writer calls registration explicitly when serializing.
- **Step 3:** For public API compatibility, `cell.font = ...` must still "just work" at save time. The key change is *when* registration happens (at save time vs. at assignment time), not *whether* it happens.
- **Public API impact:** Code that assigns styles and saves will behave identically. Code that inspects the workbook's style registry *between* assignment and save may see different intermediate state. Document this as a known behavioral change.

Output:
- explicit registration/application boundary
- characterization tests proving save-time behavior is identical
- migration note documenting the timing change for style registry inspection
- explicit compatibility decision: registry inspection between assignment and save is treated as a documented behavioral change unless a later compatibility layer restores the old timing

### Work package C: Dynamic value fields

Targets:
- `chart/data_source.py:NumberValueDescriptor`

Output:
- union-typed field strategy with parser hook

### Work package D: Ordering parity

Targets:
- classes with explicit `__elements__` / `__attrs__` overrides

Output:
- ordering override mechanism
- regression tests for XML ordering

## Benchmarks and Acceptance Thresholds

Run these benchmarks at baseline, after infrastructure, after pilot, and after every major wave.

### Required benchmarks

- object construction
- attribute reads
- attribute assignment
- parse representative workbook XML
- serialize representative workbook XML
- load/save representative workbook

### Thresholds

- attribute reads for normal fields: no material regression
- parse/serialize throughput: within 10-15%
- object construction: within 15-20% unless justified

If a wave misses thresholds, stop and profile before continuing.

## Immediate Next Steps

The next concrete tasks should be:

1. Add `scripts/generate_serialisable_goldens.py`.
2. Add a benchmark harness for parse/serialize/object construction.
3. Write the first failing test for `FieldInfo` compilation — this creates `fastpyxl/typed_serialisable/` organically.
4. Implement `FieldInfo`, `Field.*`, and alias metadata one tested feature at a time.
5. Write and pass the interop contract smoke test (new-style child in old-style parent and vice versa).
6. Translate the four pilot classes.

That is the shortest path from architecture to executable work without committing to the wrong runtime.
