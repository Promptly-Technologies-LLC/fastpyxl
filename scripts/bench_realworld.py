#!/usr/bin/env python3
"""Benchmark load/save of a real-world Excel file: openpyxl (PyPI) vs fastpyxl (local)."""

from __future__ import annotations

import io
import json
import subprocess
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def _bench(package: str, fixture: Path, iterations: int) -> dict:
    """Run benchmarks using the given package name in an isolated venv."""
    script = f"""
import io, time, statistics, json

fixture = {str(fixture)!r}
iterations = {iterations}
pkg = {package!r}

exec(f"from {{pkg}} import load_workbook")

# --- load benchmark ---
load_times = []
for _ in range(iterations):
    t0 = time.perf_counter()
    wb = load_workbook(fixture)
    t1 = time.perf_counter()
    load_times.append(t1 - t0)
    wb.close()

# --- save benchmark ---
wb = load_workbook(fixture)
save_times = []
for _ in range(iterations):
    buf = io.BytesIO()
    t0 = time.perf_counter()
    wb.save(buf)
    t1 = time.perf_counter()
    save_times.append(t1 - t0)
wb.close()

# --- read-only iteration ---
ro_times = []
for _ in range(iterations):
    t0 = time.perf_counter()
    wb = load_workbook(fixture, read_only=True)
    count = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                _ = cell.value
                count += 1
    wb.close()
    t1 = time.perf_counter()
    ro_times.append(t1 - t0)

def stats(samples):
    return {{
        "mean": statistics.mean(samples),
        "median": statistics.median(samples),
        "stdev": statistics.stdev(samples) if len(samples) > 1 else 0.0,
        "samples": samples,
    }}

print(json.dumps({{
    "load": stats(load_times),
    "save": stats(save_times),
    "read_only_iter": stats(ro_times),
    "cell_count": count,
}}))
"""
    with tempfile.TemporaryDirectory(prefix="fpyxl-bench-") as tmp:
        venv = Path(tmp) / "venv"
        subprocess.run(
            ["uv", "venv", str(venv)],
            check=True, capture_output=True,
        )
        python = venv / "bin" / "python"

        if package == "fastpyxl":
            subprocess.run(
                ["uv", "pip", "install", "--python", str(python), str(ROOT)],
                check=True, capture_output=True,
            )
        else:
            subprocess.run(
                ["uv", "pip", "install", "--python", str(python), "openpyxl"],
                check=True, capture_output=True,
            )

        result = subprocess.run(
            [str(python), "-c", script],
            capture_output=True, text=True, timeout=600,
        )
        if result.returncode != 0:
            print(f"STDERR for {package}:\n{result.stderr}", file=sys.stderr)
            raise RuntimeError(f"Benchmark failed for {package}")
        return json.loads(result.stdout.strip())


def fmt(seconds: float) -> str:
    if seconds < 1:
        return f"{seconds * 1000:.0f}ms"
    return f"{seconds:.2f}s"


def pct(old: float, new: float) -> str:
    delta = (new - old) / old * 100
    sign = "+" if delta > 0 else ""
    return f"{sign}{delta:.1f}%"


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("fixture", type=Path, help="Path to .xlsx/.xlsm file")
    parser.add_argument("--iterations", type=int, default=3, help="Iterations per benchmark")
    parser.add_argument("--output", type=Path, default=ROOT / "artifacts" / "realworld_compare.md")
    args = parser.parse_args()

    fixture = args.fixture.resolve()
    if not fixture.exists():
        raise FileNotFoundError(f"Fixture not found: {fixture}")

    print(f"Benchmarking {fixture.name} ({fixture.stat().st_size / 1024 / 1024:.1f} MB)")
    print(f"  Iterations: {args.iterations}")
    print()

    print("Running openpyxl (from PyPI) ...", flush=True)
    data_openpyxl = _bench("openpyxl", fixture, args.iterations)

    print("Running fastpyxl (local) ...", flush=True)
    data_fastpyxl = _bench("fastpyxl", fixture, args.iterations)

    benchmarks = ["load", "save", "read_only_iter"]
    labels = ["Load workbook", "Save workbook", "Read-only iteration"]

    lines = [
        f"# Real-world benchmark: {fixture.name}",
        "",
        f"File size: {fixture.stat().st_size / 1024 / 1024:.1f} MB, "
        f"cells: {data_fastpyxl.get('cell_count', 'N/A'):,}, "
        f"{args.iterations} iterations (median)",
        "",
        "| Benchmark | openpyxl | fastpyxl | Change |",
        "|---|--:|--:|--:|",
    ]
    for key, label in zip(benchmarks, labels):
        old = data_openpyxl[key]["median"]
        new = data_fastpyxl[key]["median"]
        lines.append(f"| {label} | {fmt(old)} | {fmt(new)} | **{pct(old, new)}** |")

    lines.extend([
        "",
        "Negative % = fastpyxl is faster.",
    ])

    report = "\n".join(lines)
    print()
    print(report)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(report + "\n", encoding="utf-8")

    json_out = args.output.with_suffix(".json")
    json_out.write_text(json.dumps({
        "fixture": fixture.name,
        "openpyxl": data_openpyxl,
        "fastpyxl": data_fastpyxl,
    }, indent=2) + "\n", encoding="utf-8")
    print(f"\nWrote {args.output.relative_to(ROOT)}")
    print(f"Wrote {json_out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
