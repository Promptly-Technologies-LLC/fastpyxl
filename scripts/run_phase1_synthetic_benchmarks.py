#!/usr/bin/env python3
"""Run synthetic legacy-vs-typed serializer benchmarks (Phase 1 harness)."""

from __future__ import annotations

import argparse
import json
import statistics
import sys
import time
from dataclasses import asdict, dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fastpyxl.descriptors import Integer, Sequence
from fastpyxl.descriptors.serialisable import Serialisable as LegacySerialisable
from fastpyxl.typed_serialisable.base import Serialisable as TypedSerialisable
from fastpyxl.typed_serialisable.fields import Field
from fastpyxl.xml.functions import fromstring, tostring


ARTIFACTS_DIR = ROOT / "artifacts"

_REPORT_TITLE = {
    "phase1_synthetic_benchmarks": "Phase 1 Synthetic Benchmarks",
    "phase8_synthetic_benchmarks": "Phase 8 Synthetic Benchmarks",
}


def _markdown_title(output_stem: str) -> str:
    return _REPORT_TITLE.get(
        output_stem, output_stem.replace("_", " ").title().replace("Phase8", "Phase 8")
    )


class LegacyChild(LegacySerialisable):
    tagname = "child"
    value = Integer()

    def __init__(self, value=None):
        self.value = value


class LegacyParent(LegacySerialisable):
    tagname = "parent"
    id = Integer(allow_none=True)
    children = Sequence(expected_type=LegacyChild)

    def __init__(self, id=None, children=()):
        self.id = id
        self.children = children


class TypedChild(TypedSerialisable):
    tagname = "child"
    value: int | None = Field.attribute(expected_type=int, allow_none=True)


class TypedParent(TypedSerialisable):
    tagname = "parent"
    id: int | None = Field.attribute(expected_type=int, allow_none=True)
    children: list[TypedChild] = Field.sequence(expected_type=TypedChild, default=list)


@dataclass
class Result:
    name: str
    mean_seconds: float
    median_seconds: float
    stdev_seconds: float
    samples: list[float]


def _measure(name: str, fn, iterations: int = 7) -> Result:
    samples: list[float] = []
    for _ in range(iterations):
        start = time.perf_counter()
        fn()
        samples.append(time.perf_counter() - start)
    return Result(
        name=name,
        mean_seconds=statistics.mean(samples),
        median_seconds=statistics.median(samples),
        stdev_seconds=statistics.pstdev(samples),
        samples=samples,
    )


def _build_legacy(n: int = 2000) -> list[LegacyParent]:
    return [
        LegacyParent(id=i, children=[LegacyChild(value=i), LegacyChild(value=i + 1)])
        for i in range(n)
    ]


def _build_typed(n: int = 2000) -> list[TypedParent]:
    return [
        TypedParent(id=i, children=[TypedChild(value=i), TypedChild(value=i + 1)])
        for i in range(n)
    ]


def _render_legacy(n: int = 1200) -> int:
    items = _build_legacy(n)
    size = 0
    for item in items:
        size += len(tostring(item.to_tree()))
    return size


def _render_typed(n: int = 1200) -> int:
    items = _build_typed(n)
    size = 0
    for item in items:
        size += len(tostring(item.to_tree()))
    return size


def _parse_legacy(n: int = 1200) -> int:
    xml = tostring(LegacyParent(id=1, children=[LegacyChild(1), LegacyChild(2)]).to_tree())
    size = 0
    for _ in range(n):
        obj = LegacyParent.from_tree(fromstring(xml))
        size += (obj.id or 0)
    return size


def _parse_typed(n: int = 1200) -> int:
    xml = tostring(TypedParent(id=1, children=[TypedChild(value=1), TypedChild(value=2)]).to_tree())
    size = 0
    for _ in range(n):
        obj = TypedParent.from_tree(fromstring(xml))
        size += (obj.id or 0)
    return size


def _pct_change(new: float, old: float) -> float:
    if old == 0:
        return 0.0
    return ((new - old) / old) * 100.0


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output-stem",
        default="phase1_synthetic_benchmarks",
        help="Basename for artifacts/STEM.json and artifacts/STEM.md (default: phase1_synthetic_benchmarks)",
    )
    args = parser.parse_args()
    output_stem = args.output_stem
    if output_stem.endswith((".json", ".md")):
        output_stem = output_stem.rsplit(".", 1)[0]

    ARTIFACTS_DIR.mkdir(exist_ok=True)

    benchmark_results = [
        _measure("construct_legacy_2k", lambda: _build_legacy(2000)),
        _measure("construct_typed_2k", lambda: _build_typed(2000)),
        _measure("render_legacy_1200", lambda: _render_legacy(1200)),
        _measure("render_typed_1200", lambda: _render_typed(1200)),
        _measure("parse_legacy_1200", lambda: _parse_legacy(1200)),
        _measure("parse_typed_1200", lambda: _parse_typed(1200)),
    ]

    by_name = {r.name: r for r in benchmark_results}
    comparisons = {
        "construction_typed_vs_legacy_percent": _pct_change(
            by_name["construct_typed_2k"].mean_seconds,
            by_name["construct_legacy_2k"].mean_seconds,
        ),
        "render_typed_vs_legacy_percent": _pct_change(
            by_name["render_typed_1200"].mean_seconds,
            by_name["render_legacy_1200"].mean_seconds,
        ),
        "parse_typed_vs_legacy_percent": _pct_change(
            by_name["parse_typed_1200"].mean_seconds,
            by_name["parse_legacy_1200"].mean_seconds,
        ),
    }

    payload = {
        "results": [asdict(r) for r in benchmark_results],
        "comparisons": comparisons,
    }
    out_json = ARTIFACTS_DIR / f"{output_stem}.json"
    out_json.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    lines = [
        f"# {_markdown_title(output_stem)}",
        "",
        "| Benchmark | Mean (s) | Median (s) | StdDev (s) |",
        "|---|---:|---:|---:|",
    ]
    for r in benchmark_results:
        lines.append(
            f"| `{r.name}` | {r.mean_seconds:.6f} | {r.median_seconds:.6f} | {r.stdev_seconds:.6f} |"
        )
    lines.extend(
        [
            "",
            "## Typed vs Legacy",
            "",
            f"- Construction delta: {comparisons['construction_typed_vs_legacy_percent']:.2f}%",
            f"- Render delta: {comparisons['render_typed_vs_legacy_percent']:.2f}%",
            f"- Parse delta: {comparisons['parse_typed_vs_legacy_percent']:.2f}%",
            "",
            "Positive means typed is slower; negative means typed is faster.",
            "",
        ]
    )
    out_md = ARTIFACTS_DIR / f"{output_stem}.md"
    out_md.write_text("\n".join(lines), encoding="utf-8")

    print(f"Wrote {out_json.relative_to(ROOT)}")
    print(f"Wrote {out_md.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
