#!/usr/bin/env python3
"""cProfile a real-world workbook load to identify hot paths."""

from __future__ import annotations

import cProfile
import pstats
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("fixture", type=Path, help="Path to .xlsx/.xlsm file")
    parser.add_argument("--top", type=int, default=40, help="Number of top functions to show")
    parser.add_argument("--output-dir", type=Path, default=ROOT / "artifacts")
    args = parser.parse_args()

    fixture = args.fixture.resolve()
    if not fixture.exists():
        raise FileNotFoundError(f"Fixture not found: {fixture}")

    from fastpyxl import load_workbook

    print(f"Profiling load of {fixture.name} ({fixture.stat().st_size / 1024 / 1024:.1f} MB)")
    print()

    pr = cProfile.Profile()
    pr.enable()
    wb = load_workbook(fixture)
    pr.disable()
    wb.close()

    args.output_dir.mkdir(parents=True, exist_ok=True)

    for sort_key in ("tottime", "cumtime"):
        print(f"\n{'=' * 80}")
        print(f"  Top {args.top} by {sort_key}")
        print(f"{'=' * 80}\n")

        stats = pstats.Stats(pr, stream=sys.stdout)
        stats.sort_stats(sort_key)
        stats.print_stats(args.top)

        out = args.output_dir / f"profile_load_{sort_key}.txt"
        with open(out, "w") as f:
            stats_file = pstats.Stats(pr, stream=f)
            stats_file.sort_stats(sort_key)
            stats_file.print_stats(args.top)
        print(f"Wrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
