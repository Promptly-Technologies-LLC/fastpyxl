#!/usr/bin/env python3
"""Microbenchmarks for access='indexed' (issue #75).

Compares normal, read_only, and indexed modes on:
1. sparse random cell hits
2. narrow column strip
3. full sheet scan

Reports wall time and max RSS. Index/load cost is included in each timed
load+access measurement (not split), matching typical caller workflows.
"""

from __future__ import annotations

import random
import resource
import tempfile
import time
from pathlib import Path

from fastpyxl import Workbook, load_workbook


def _rss_kb() -> int:
    return resource.getrusage(resource.RUSAGE_SELF).ru_maxrss


def _make_sparse(path: Path, rows: int = 5000, hits: int = 200) -> list[tuple[int, int]]:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sparse"
    coords = [(r, 1) for r in range(1, rows + 1, max(1, rows // hits))]
    for r, c in coords:
        ws.cell(r, c, value=f"r{r}")
        ws.cell(r, 3, value=r)
    wb.save(path)
    return coords


def _make_dense_strings(path: Path, rows: int = 2000, cols: int = 10) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Dense"
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            ws.cell(r, c, value=f"r{r}c{c}")
    wb.save(path)


def _bench(label: str, path: Path, mode: str, work) -> None:
    kwargs = {}
    if mode == "read_only":
        kwargs["read_only"] = True
    elif mode == "indexed":
        kwargs["access"] = "indexed"

    # Warm one load so import/zip caches are less noisy.
    wb0 = load_workbook(path, **kwargs)
    wb0.close()

    t0 = time.perf_counter()
    rss0 = _rss_kb()
    wb = load_workbook(path, **kwargs)
    work(wb)
    wb.close()
    elapsed = time.perf_counter() - t0
    rss1 = _rss_kb()
    print(f"{label:28} {mode:10} {_fmt(elapsed)}  maxRSS={max(rss0, rss1)}KB")


def _fmt(seconds: float) -> str:
    if seconds < 1:
        return f"{seconds * 1000:8.2f}ms"
    return f"{seconds:8.3f}s "


def main() -> None:
    random.seed(0)
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        sparse = tmp_path / "sparse.xlsx"
        dense = tmp_path / "dense.xlsx"
        coords = _make_sparse(sparse)
        _make_dense_strings(dense)

        sample = coords[:: max(1, len(coords) // 100)]
        if len(sample) < 100:
            sample = coords[:100]

        def sparse_random(wb):
            ws = wb.active
            for r, c in sample:
                _ = ws.cell(r, c).value

        def column_strip(wb):
            ws = wb.active
            for row in ws.iter_rows(min_col=1, max_col=1, values_only=True):
                _ = row[0]

        def full_scan(wb):
            ws = wb.active
            for row in ws.iter_rows(values_only=True):
                _ = row

        print("=== sparse workbook ===")
        for mode in ("normal", "read_only", "indexed"):
            _bench("sparse random cells", sparse, mode, sparse_random)
        for mode in ("normal", "read_only", "indexed"):
            _bench("narrow column strip", sparse, mode, column_strip)
        for mode in ("normal", "read_only", "indexed"):
            _bench("full sheet scan", sparse, mode, full_scan)

        print("=== dense string workbook ===")
        for mode in ("normal", "read_only", "indexed"):
            _bench("full sheet scan", dense, mode, full_scan)
        for mode in ("normal", "read_only", "indexed"):
            _bench("sparse random cells", dense, mode, sparse_random)


if __name__ == "__main__":
    main()
