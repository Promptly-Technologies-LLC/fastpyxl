#!/usr/bin/env python3
"""cProfile construction path for synthetic legacy vs typed benchmark models."""

from __future__ import annotations

import argparse
import importlib.util
import pstats
import sys
from io import StringIO
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def _load_benchmark_module():
    path = ROOT / "scripts" / "run_phase1_synthetic_benchmarks.py"
    spec = importlib.util.spec_from_file_location("_synthetic_benchmark_models", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load {path}")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = mod
    spec.loader.exec_module(mod)
    return mod


def _format_stats(profile, sort_key: str, limit: int) -> str:
    s = StringIO()
    stats = pstats.Stats(profile, stream=s)
    stats.strip_dirs()
    stats.sort_stats(sort_key)
    stats.print_stats(limit)
    return s.getvalue()


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--n", type=int, default=2000, help="parents per build (default 2000)")
    parser.add_argument("--repeat", type=int, default=12, help="times to run each build loop")
    parser.add_argument(
        "--sort",
        action="append",
        choices=("tottime", "cumtime"),
        dest="sorts",
        help="pstats sort key (default: both tottime and cumtime); may be passed twice",
    )
    parser.add_argument("--top", type=int, default=45, help="lines of print_stats output")
    args = parser.parse_args()
    sorts = args.sorts or ["tottime", "cumtime"]

    bench = _load_benchmark_module()
    artifacts = ROOT / "artifacts"
    artifacts.mkdir(exist_ok=True)

    from cProfile import Profile

    for label, builder in (
        ("typed", lambda: bench._build_typed(args.n)),
        ("legacy", lambda: bench._build_legacy(args.n)),
    ):
        pr = Profile()
        pr.enable()
        for _ in range(args.repeat):
            builder()
        pr.disable()
        for sort_key in sorts:
            body = _format_stats(pr, sort_key, args.top)
            header = (
                f"# construct profile ({label})\n"
                f"# n={args.n} repeat={args.repeat} sort={sort_key} top={args.top}\n\n"
            )
            out = artifacts / f"construct_profile_{label}_{sort_key}.txt"
            out.write_text(header + body, encoding="utf-8")
            print(f"Wrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    sys.path.insert(0, str(ROOT))
    main()
