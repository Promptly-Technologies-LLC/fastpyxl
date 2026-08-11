#!/usr/bin/env python3
"""Microbenchmarks for keep_formula_cache dual-load (issue #71).

Reports:
1. Parse wall time: default vs keep_formula_cache=True
2. Peak RSS: default vs dual-load
3. Wall/RSS: double load_workbook (False + True) vs single dual-load
"""

from __future__ import annotations

import os
import resource
import statistics
import time
from pathlib import Path

from fastpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
FIXTURE = ROOT / "fastpyxl/tests/data/genuine/sample.xlsx"
ITERATIONS = 20


def _rss_kb() -> int:
    return resource.getrusage(resource.RUSAGE_SELF).ru_maxrss


def _bench_load(**kwargs) -> tuple[list[float], int]:
    times: list[float] = []
    for _ in range(ITERATIONS):
        t0 = time.perf_counter()
        wb = load_workbook(FIXTURE, **kwargs)
        t1 = time.perf_counter()
        times.append(t1 - t0)
        wb.close()
    return times, _rss_kb()


def _bench_double_load() -> tuple[list[float], int]:
    times: list[float] = []
    for _ in range(ITERATIONS):
        t0 = time.perf_counter()
        wb_f = load_workbook(FIXTURE, data_only=False)
        wb_v = load_workbook(FIXTURE, data_only=True)
        t1 = time.perf_counter()
        times.append(t1 - t0)
        wb_f.close()
        wb_v.close()
    return times, _rss_kb()


def _fmt(times: list[float]) -> str:
    return (
        f"mean={statistics.mean(times)*1000:.2f}ms "
        f"median={statistics.median(times)*1000:.2f}ms "
        f"stdev={statistics.pstdev(times)*1000:.2f}ms"
    )


def main() -> None:
    print(f"fixture={FIXTURE}")
    print(f"iterations={ITERATIONS} pid={os.getpid()}")

    default_t, default_rss = _bench_load()
    print(f"1) default load:          {_fmt(default_t)}  maxRSS={default_rss}KB")

    dual_t, dual_rss = _bench_load(keep_formula_cache=True)
    print(f"2) keep_formula_cache:    {_fmt(dual_t)}  maxRSS={dual_rss}KB")

    double_t, double_rss = _bench_double_load()
    print(f"3) double load (F+T):     {_fmt(double_t)}  maxRSS={double_rss}KB")

    mean_default = statistics.mean(default_t)
    mean_dual = statistics.mean(dual_t)
    mean_double = statistics.mean(double_t)
    print()
    print(f"dual/default wall ratio:  {mean_dual / mean_default:.3f}")
    print(f"dual/double wall ratio:   {mean_dual / mean_double:.3f}")
    print(f"dual-default RSS delta:   {dual_rss - default_rss}KB")
    print(f"double-dual RSS delta:    {double_rss - dual_rss}KB")


if __name__ == "__main__":
    main()
