#!/usr/bin/env python3
"""Run Phase 0 baseline benchmarks and persist results."""

from __future__ import annotations

import io
import json
import statistics
import sys
import tempfile
import time
from dataclasses import dataclass, asdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
ARTIFACTS_DIR = ROOT / "artifacts"
DEFAULT_XLSX = ROOT / "fastpyxl" / "reader" / "tests" / "data" / "sample.xlsx"
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from fastpyxl import Workbook, load_workbook
from fastpyxl.styles.fonts import Font


@dataclass
class BenchmarkResult:
    name: str
    unit: str
    iterations: int
    samples: list[float]
    mean: float
    stdev: float
    median: float


def _measure(name: str, unit: str, iterations: int, fn) -> BenchmarkResult:
    samples: list[float] = []
    for _ in range(iterations):
        start = time.perf_counter()
        fn()
        elapsed = time.perf_counter() - start
        samples.append(elapsed)
    return BenchmarkResult(
        name=name,
        unit=unit,
        iterations=iterations,
        samples=samples,
        mean=statistics.mean(samples),
        stdev=statistics.pstdev(samples),
        median=statistics.median(samples),
    )


def bench_object_construction(n: int = 100_000) -> float:
    for _ in range(n):
        Font(name="Calibri", sz=11, b=True, i=False)
    return 0.0


def bench_attribute_read(n: int = 300_000) -> float:
    f = Font(name="Calibri", sz=11, b=True, i=False)
    sink = 0.0
    for _ in range(n):
        if f.name:
            sink += f.sz or 0.0
        if f.b:
            sink += 1.0
    if sink < 0:
        raise RuntimeError("Unreachable sink guard")
    return sink


def bench_attribute_assignment(n: int = 150_000) -> float:
    f = Font(name="Calibri", sz=11, b=True, i=False)
    for idx in range(n):
        f.sz = 10.0 + (idx % 5)
        f.b = bool(idx % 2)
    return float(n)


def bench_parse_workbook(path: Path) -> float:
    wb = load_workbook(path)
    wb.close()
    return 0.0


def bench_serialize_workbook() -> float:
    wb = Workbook()
    ws = wb.active
    for row in range(1, 51):
        ws.append([row, row * 2, row * 3, f"row-{row}"])
    buffer = io.BytesIO()
    wb.save(buffer)
    return float(buffer.tell())


def bench_load_save_roundtrip(path: Path) -> float:
    with tempfile.NamedTemporaryFile(suffix=".xlsx") as tmp:
        wb = load_workbook(path)
        wb.save(tmp.name)
        wb.close()
    return 0.0


def main() -> None:
    ARTIFACTS_DIR.mkdir(exist_ok=True)
    xlsx = DEFAULT_XLSX
    if not xlsx.exists():
        raise FileNotFoundError(f"Fixture not found: {xlsx}")

    results: list[BenchmarkResult] = []
    results.append(_measure("object_construction_100k", "seconds", 5, bench_object_construction))
    results.append(_measure("attribute_read_300k", "seconds", 5, bench_attribute_read))
    results.append(_measure("attribute_assignment_150k", "seconds", 5, bench_attribute_assignment))
    results.append(_measure("parse_representative_workbook", "seconds", 5, lambda: bench_parse_workbook(xlsx)))
    results.append(_measure("serialize_representative_workbook", "seconds", 5, bench_serialize_workbook))
    results.append(_measure("load_save_roundtrip", "seconds", 3, lambda: bench_load_save_roundtrip(xlsx)))

    payload = {
        "fixture": str(xlsx.relative_to(ROOT)),
        "results": [asdict(r) for r in results],
    }

    out_json = ARTIFACTS_DIR / "phase0_baseline_benchmarks.json"
    out_json.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    lines = [
        "# Phase 0 Baseline Benchmarks",
        "",
        f"- Fixture workbook: `{payload['fixture']}`",
        "",
        "| Benchmark | Mean (s) | Median (s) | StdDev (s) | Iterations |",
        "|---|---:|---:|---:|---:|",
    ]
    for result in results:
        lines.append(
            f"| `{result.name}` | {result.mean:.8f} | {result.median:.8f} | {result.stdev:.8f} | {result.iterations} |"
        )
    lines.append("")

    out_md = ARTIFACTS_DIR / "phase0_baseline_benchmarks.md"
    out_md.write_text("\n".join(lines), encoding="utf-8")
    print(f"Wrote {out_json.relative_to(ROOT)}")
    print(f"Wrote {out_md.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
