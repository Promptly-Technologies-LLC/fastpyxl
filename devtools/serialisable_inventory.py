"""AST inventory of Serialisable subclasses and descriptor usage (optional maintenance helper)."""

from __future__ import annotations

import argparse
import ast
from collections import Counter
from dataclasses import dataclass
from pathlib import Path


def _repo_root() -> Path:
    return Path(__file__).resolve().parents[1]


DESCRIPTOR_NAMES = {
    "Typed",
    "Alias",
    "Sequence",
    "NoneSet",
    "Set",
    "NestedSequence",
    "MultiSequence",
    "MultiSequencePart",
}


@dataclass(frozen=True)
class ModuleSummary:
    module: str
    package: str
    serialisable_classes: int
    descriptor_counts: Counter[str]


def _is_serialisable_base(node: ast.expr) -> bool:
    if isinstance(node, ast.Name):
        return node.id == "Serialisable"
    if isinstance(node, ast.Attribute):
        return node.attr == "Serialisable"
    return False


def _called_name(node: ast.AST) -> str | None:
    if isinstance(node, ast.Name):
        return node.id
    if isinstance(node, ast.Attribute):
        return node.attr
    return None


def _parse_module(path: Path, root: Path) -> ModuleSummary:
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    rel = path.relative_to(root)
    module = rel.with_suffix("").as_posix().replace("/", ".")
    package = rel.parts[1] if len(rel.parts) > 2 else rel.parts[0]

    serialisable_classes = 0
    descriptor_counts: Counter[str] = Counter()

    for node in ast.walk(tree):
        if isinstance(node, ast.ClassDef):
            if any(_is_serialisable_base(base) for base in node.bases):
                serialisable_classes += 1
        if isinstance(node, ast.Call):
            name = _called_name(node.func)
            if name in DESCRIPTOR_NAMES:
                descriptor_counts[name] += 1

    return ModuleSummary(
        module=module,
        package=package,
        serialisable_classes=serialisable_classes,
        descriptor_counts=descriptor_counts,
    )


def inventory(root: Path, package_root: Path) -> list[ModuleSummary]:
    summaries: list[ModuleSummary] = []
    for path in package_root.rglob("*.py"):
        if "tests" in path.parts:
            continue
        summaries.append(_parse_module(path, root))
    return summaries


def write_field_mapping(path: Path) -> None:
    lines = [
        "# Descriptor → typed field mapping",
        "",
        "| Legacy descriptor/pattern | Typed runtime target | Notes |",
        "|---|---|---|",
        "| `Typed` / `Convertible` / scalar subclasses | `Field.attribute(...)` + converter hook | Scalar XML attributes in `__attrs__`.",
        "| `NestedValue` (`val` attr) | `Field.nested_value(...)` | Nested node with value on `val` attribute.",
        "| `NestedText` | `Field.nested_text(...)` | Nested node with text content.",
        "| `NestedBool` / `EmptyTag` | `Field.nested_bool(...)` | Support empty-tag truthy semantics and custom renderer hooks.",
        "| `Sequence` | `Field.sequence(...)` | Repeated homogeneous children; container defaults to `list`.",
        "| `NestedSequence` | `Field.nested_sequence(...)` | Wrapped collection with optional `count` attribute.",
        "| `MultiSequence` + `MultiSequencePart` | `Field.multi_sequence(...)` | Tagged heterogeneous child list with tag dispatch.",
        "| `Alias` | `AliasField(target=...)` | No storage; forwards get/set to the target field.",
        "| `Set` / `NoneSet` | `Field.attribute(...)` + validator | Prefer `Literal[...]` when practical, else validator hook.",
        "| `MatchPattern` | `Field.attribute(...)` + validator | Regex validation remains eager.",
        "| Namespaced fields (`namespace=...`) | `Field.*(..., namespace=...)` | Keep namespaced cache entries precomputed.",
        "| Hyphenated names | `Field.*(..., hyphenated=True)` | Preserve underscore<->hyphen conversion parity.",
        "",
    ]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def _work_package_for_module(module: str) -> tuple[str, str] | None:
    package_map = {
        "fastpyxl.pivot.record": ("A", "Heterogeneous multi-sequence records."),
        "fastpyxl.pivot.cache": ("A", "Mixed multi-sequence and deep nested-sequence coverage."),
        "fastpyxl.chart.plotarea": ("A", "Tag-dispatched heterogeneous chart and axis sequences."),
        "fastpyxl.styles.styleable": ("B", "Style assignment side effects currently hidden in descriptor set."),
        "fastpyxl.chart.data_source": ("C", "Dynamic number/string value coercion (`#N/A` compatibility)."),
    }
    if module in package_map:
        return package_map[module]
    return None


def write_hard_cases(path: Path, summaries: list[ModuleSummary]) -> None:
    package_counts: Counter[str] = Counter()
    descriptor_totals: Counter[str] = Counter()
    hard_rows: list[tuple[str, str, str, str]] = []

    for summary in summaries:
        if summary.serialisable_classes:
            package_counts[summary.package] += summary.serialisable_classes
        descriptor_totals.update(summary.descriptor_counts)
        package = _work_package_for_module(summary.module)
        if package is not None:
            work_package, note = package
            hard_rows.append((summary.module, summary.package, work_package, note))

    hard_rows.sort()
    top_packages = package_counts.most_common(10)
    descriptor_lines = [
        f"- `{name}`: {count}" for name, count in descriptor_totals.most_common()
    ]

    lines = [
        "# Serialisable hard-case modules",
        "",
        "## Inventory snapshot",
        "",
        f"- Serialisable class count: {sum(package_counts.values())}",
        "- Top package distribution:",
    ]
    lines.extend([f"  - `{name}`: {count}" for name, count in top_packages])
    lines.append("- Hard-pattern descriptor call counts:")
    lines.extend([f"  {line}" for line in descriptor_lines])
    lines.extend(
        [
            "",
            "## Work package assignment",
            "",
            "| Module | Package | Work package | Notes |",
            "|---|---|---|---|",
        ]
    )
    for module, package, work_package, note in hard_rows:
        lines.append(f"| `{module}` | `{package}` | `{work_package}` | {note} |")
    lines.append("")
    lines.append("## Notes")
    lines.append("")
    lines.append(
        "- Work package keys: A multi-sequence, B style side effects, C dynamic values, D ordering parity."
    )
    lines.append(
        "- Work package D: add during migration when explicit `__elements__`/`__attrs__` overrides appear."
    )
    lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    root = _repo_root()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=root / "artifacts",
        help="Where to write markdown tables (default: ./artifacts)",
    )
    args = parser.parse_args()
    out = args.output_dir
    out.mkdir(parents=True, exist_ok=True)
    summaries = inventory(root, root / "fastpyxl")
    write_field_mapping(out / "serialisable_field_mapping.md")
    write_hard_cases(out / "serialisable_hard_cases.md", summaries)
    print(f"Wrote {out / 'serialisable_field_mapping.md'}")
    print(f"Wrote {out / 'serialisable_hard_cases.md'}")


if __name__ == "__main__":
    main()
