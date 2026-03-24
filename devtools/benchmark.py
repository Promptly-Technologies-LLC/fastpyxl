"""Timing helpers and report writers for local performance work."""

from __future__ import annotations

import json
import statistics
import time
from collections.abc import Callable, Iterable, Mapping, Sequence
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class BenchmarkResult:
    """One benchmark: repeated wall-clock samples (``time.perf_counter``)."""

    name: str
    unit: str
    iterations: int
    samples: list[float]
    mean: float
    median: float
    stdev: float


def measure(
    name: str,
    fn: Callable[[], Any],
    *,
    iterations: int,
    unit: str = "seconds",
) -> BenchmarkResult:
    """Run ``fn`` ``iterations`` times and collect per-iteration elapsed times."""
    samples: list[float] = []
    for _ in range(iterations):
        start = time.perf_counter()
        fn()
        samples.append(time.perf_counter() - start)
    return BenchmarkResult(
        name=name,
        unit=unit,
        iterations=iterations,
        samples=samples,
        mean=statistics.mean(samples),
        median=statistics.median(samples),
        stdev=statistics.pstdev(samples) if len(samples) > 1 else 0.0,
    )


def percent_change(new: float, old: float) -> float:
    """Percentage delta: ``((new - old) / old) * 100``; 0 if ``old`` is 0."""
    if old == 0:
        return 0.0
    return ((new - old) / old) * 100.0


def compare_by_name(
    left: Sequence[BenchmarkResult],
    right: Sequence[BenchmarkResult],
    *,
    left_label: str = "A",
    right_label: str = "B",
) -> list[dict[str, Any]]:
    """Align results by ``name`` and compute mean/median deltas between runs."""
    by_right = {r.name: r for r in right}
    rows: list[dict[str, Any]] = []
    for a in left:
        b = by_right.get(a.name)
        if b is None:
            continue
        rows.append(
            {
                "name": a.name,
                f"mean_{left_label}": a.mean,
                f"mean_{right_label}": b.mean,
                "mean_delta_percent": percent_change(b.mean, a.mean),
                f"median_{left_label}": a.median,
                f"median_{right_label}": b.median,
                "median_delta_percent": percent_change(b.median, a.median),
            }
        )
    return rows


def write_json(path: Path, payload: Mapping[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(dict(payload), indent=2), encoding="utf-8")


def results_payload(
    results: Iterable[BenchmarkResult],
    *,
    extra: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """Serialize benchmark results for JSON (includes raw samples)."""
    base: dict[str, Any] = {"results": [asdict(r) for r in results]}
    if extra:
        base.update(extra)
    return base


def result_from_dict(d: Mapping[str, Any]) -> BenchmarkResult:
    """Restore :class:`BenchmarkResult` from :func:`results_payload` JSON."""
    return BenchmarkResult(
        name=str(d["name"]),
        unit=str(d["unit"]),
        iterations=int(d["iterations"]),
        samples=[float(x) for x in d["samples"]],
        mean=float(d["mean"]),
        median=float(d["median"]),
        stdev=float(d["stdev"]),
    )


def markdown_results_table(
    results: Sequence[BenchmarkResult],
    *,
    include_iterations: bool = False,
    float_fmt: str = ".8f",
) -> list[str]:
    """Markdown table lines (no leading title)."""
    if include_iterations:
        header = "| Benchmark | Mean (s) | Median (s) | StdDev (s) | Iterations |"
        sep = "|---|---:|---:|---:|---:|"
    else:
        header = "| Benchmark | Mean (s) | Median (s) | StdDev (s) |"
        sep = "|---|---:|---:|---:|"
    lines = [header, sep]
    for r in results:
        if include_iterations:
            lines.append(
                f"| `{r.name}` | {r.mean:{float_fmt}} | {r.median:{float_fmt}} | "
                f"{r.stdev:{float_fmt}} | {r.iterations} |"
            )
        else:
            lines.append(
                f"| `{r.name}` | {r.mean:{float_fmt}} | {r.median:{float_fmt}} | {r.stdev:{float_fmt}} |"
            )
    return lines


def markdown_compare_table(
    rows: Sequence[Mapping[str, Any]],
    *,
    left_label: str = "A",
    right_label: str = "B",
    float_fmt: str = ".8f",
) -> list[str]:
    """Markdown table for :func:`compare_by_name` output."""
    lines = [
        "| Benchmark | "
        f"Mean {left_label} (s) | Mean {right_label} (s) | Δ mean % | "
        f"Median {left_label} (s) | Median {right_label} (s) | Δ median % |",
        "|---|---:|---:|---:|---:|---:|---:|",
    ]
    for row in rows:
        lines.append(
            "| `{name}` | {ma:{fmt}} | {mb:{fmt}} | {d_mean:.2f} | "
            "{meda:{fmt}} | {medb:{fmt}} | {d_med:.2f} |".format(
                name=row["name"],
                ma=row[f"mean_{left_label}"],
                mb=row[f"mean_{right_label}"],
                d_mean=row["mean_delta_percent"],
                meda=row[f"median_{left_label}"],
                medb=row[f"median_{right_label}"],
                d_med=row["median_delta_percent"],
                fmt=float_fmt,
            )
        )
    return lines
