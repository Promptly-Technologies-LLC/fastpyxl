#!/usr/bin/env python3
"""Microbenchmarks for access='indexed' (issue #75).

Compares normal, read_only, and indexed modes on:
1. sparse random cell hits
2. narrow column strip
3. full sheet scan

On both a large sparse workbook and a string-heavy dense workbook that uses
a real shared-string table (not inlineStr).

Reports wall time and max RSS. Index/load cost is included in each timed
load+access measurement (not split), matching typical caller workflows.
"""

from __future__ import annotations

import random
import resource
import tempfile
import time
from io import BytesIO
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile

from fastpyxl import Workbook, load_workbook
from fastpyxl.utils import get_column_letter


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


def _make_dense_shared_strings(path: Path, rows: int = 2000, cols: int = 10) -> list[tuple[int, int]]:
    """Build a dense workbook whose cells reference xl/sharedStrings.xml."""
    # Start from a normal empty package so content types / rels stay valid.
    skeleton = BytesIO()
    Workbook().save(skeleton)
    skeleton.seek(0)

    # Reuse a pool so uniqueCount << cell count (string-heavy but compressible SST).
    strings = [f"token-{i}" for i in range(500)]
    sheet_rows: list[str] = []
    coords: list[tuple[int, int]] = []
    sst_body = "".join(f"<si><t>{escape(s)}</t></si>" for s in strings)
    sst = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        f'count="{rows * cols}" uniqueCount="{len(strings)}">{sst_body}</sst>'
    )

    for r in range(1, rows + 1):
        cells = []
        for c in range(1, cols + 1):
            idx = (r * cols + c) % len(strings)
            ref = f"{get_column_letter(c)}{r}"
            cells.append(f'<c r="{ref}" t="s"><v>{idx}</v></c>')
            if c == 1 and r % max(1, rows // 100) == 0:
                coords.append((r, c))
        sheet_rows.append(f'<row r="{r}" spans="1:{cols}">{"".join(cells)}</row>')

    sheet = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<dimension ref="A1:{get_column_letter(cols)}{rows}"/>'
        f'<sheetData>{"".join(sheet_rows)}</sheetData></worksheet>'
    )

    out = BytesIO()
    with ZipFile(skeleton, "r") as zin, ZipFile(out, "w", ZIP_DEFLATED) as zout:
        names = set(zin.namelist())
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "xl/worksheets/sheet1.xml":
                data = sheet.encode("utf-8")
            elif info.filename == "[Content_Types].xml" and "sharedStrings" not in data.decode(
                "utf-8", "ignore"
            ):
                text = data.decode("utf-8")
                override = (
                    '<Override PartName="/xl/sharedStrings.xml" '
                    'ContentType="application/vnd.openxmlformats-officedocument.'
                    'spreadsheetml.sharedStrings+xml"/>'
                )
                text = text.replace("</Types>", override + "</Types>")
                data = text.encode("utf-8")
            elif info.filename == "xl/_rels/workbook.xml.rels" and "sharedStrings" not in data.decode(
                "utf-8", "ignore"
            ):
                text = data.decode("utf-8")
                # Pick a free rId.
                rid = "rId99"
                rel = (
                    f'<Relationship Id="{rid}" '
                    'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                    'relationships/sharedStrings" Target="sharedStrings.xml"/>'
                )
                text = text.replace("</Relationships>", rel + "</Relationships>")
                data = text.encode("utf-8")
            zout.writestr(info.filename, data)
        if "xl/sharedStrings.xml" not in names:
            zout.writestr("xl/sharedStrings.xml", sst.encode("utf-8"))

    path.write_bytes(out.getvalue())
    if len(coords) < 100:
        coords = [(r, 1) for r in range(1, rows + 1, max(1, rows // 100))]
    return coords


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
        sparse_coords = _make_sparse(sparse)
        dense_coords = _make_dense_shared_strings(dense)

        def make_sample(coords):
            sample = coords[:: max(1, len(coords) // 100)]
            if len(sample) < 100:
                sample = coords[:100]
            return sample

        sparse_sample = make_sample(sparse_coords)
        dense_sample = make_sample(dense_coords)

        def sparse_random_on(sample):
            def work(wb):
                ws = wb.active
                for r, c in sample:
                    _ = ws.cell(r, c).value

            return work

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
            _bench("sparse random cells", sparse, mode, sparse_random_on(sparse_sample))
        for mode in ("normal", "read_only", "indexed"):
            _bench("narrow column strip", sparse, mode, column_strip)
        for mode in ("normal", "read_only", "indexed"):
            _bench("full sheet scan", sparse, mode, full_scan)

        print("=== dense SST workbook ===")
        for mode in ("normal", "read_only", "indexed"):
            _bench("full sheet scan", dense, mode, full_scan)
        for mode in ("normal", "read_only", "indexed"):
            _bench("sparse random cells", dense, mode, sparse_random_on(dense_sample))


if __name__ == "__main__":
    main()
