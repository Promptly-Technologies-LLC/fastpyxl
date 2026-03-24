#!/usr/bin/env python3
"""Performance tooling: micro-benchmarks, synthetic serializer benchmarks, git compare, and profiling."""

from __future__ import annotations

import argparse
import io
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from devtools import synthetic_serialisable as synth
from devtools.benchmark import (
    BenchmarkResult,
    compare_by_name,
    markdown_compare_table,
    markdown_results_table,
    measure,
    percent_change,
    result_from_dict,
    results_payload,
    write_json,
)
from devtools.paths import try_relative
from devtools.profile import profile_repeat, write_profile_text
from fastpyxl import Workbook, load_workbook
from fastpyxl.styles.fonts import Font

DEFAULT_XLSX = ROOT / "fastpyxl" / "reader" / "tests" / "data" / "sample.xlsx"

_SUITES: dict[str, dict[str, tuple[str, ...] | list[str]]] = {
    "baseline": {
        "json_candidates": ("baseline_benchmarks.json", "phase0_baseline_benchmarks.json"),
        "extra_args": [],
    },
    "synthetic": {
        "json_candidates": ("synthetic_benchmarks.json", "phase1_synthetic_benchmarks.json"),
        "extra_args": ["--output-stem", "synthetic_benchmarks"],
    },
}


def _bench_object_construction(n: int = 100_000) -> float:
    for _ in range(n):
        Font(name="Calibri", sz=11, b=True, i=False)
    return 0.0


def _bench_attribute_read(n: int = 300_000) -> float:
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


def _bench_attribute_assignment(n: int = 150_000) -> float:
    f = Font(name="Calibri", sz=11, b=True, i=False)
    for idx in range(n):
        f.sz = 10.0 + (idx % 5)
        f.b = bool(idx % 2)
    return float(n)


def _bench_parse_workbook(path: Path) -> float:
    wb = load_workbook(path)
    wb.close()
    return 0.0


def _bench_serialize_workbook() -> float:
    wb = Workbook()
    ws = wb.active
    for row in range(1, 51):
        ws.append([row, row * 2, row * 3, f"row-{row}"])
    buffer = io.BytesIO()
    wb.save(buffer)
    return float(buffer.tell())


def _bench_load_save_roundtrip(path: Path) -> float:
    with tempfile.NamedTemporaryFile(suffix=".xlsx") as tmp:
        wb = load_workbook(path)
        wb.save(tmp.name)
        wb.close()
    return 0.0


def _run_baseline_micro(xlsx: Path) -> list[BenchmarkResult]:
    results: list[BenchmarkResult] = []
    results.append(
        measure("object_construction_100k", _bench_object_construction, iterations=5)
    )
    results.append(measure("attribute_read_300k", _bench_attribute_read, iterations=5))
    results.append(
        measure("attribute_assignment_150k", _bench_attribute_assignment, iterations=5)
    )
    results.append(
        measure(
            "parse_representative_workbook",
            lambda: _bench_parse_workbook(xlsx),
            iterations=5,
        )
    )
    results.append(
        measure("serialize_representative_workbook", _bench_serialize_workbook, iterations=5)
    )
    results.append(
        measure(
            "load_save_roundtrip",
            lambda: _bench_load_save_roundtrip(xlsx),
            iterations=3,
        )
    )
    return results


def _cmd_baseline(args: argparse.Namespace) -> None:
    stem = args.output_stem
    if stem.endswith((".json", ".md")):
        stem = stem.rsplit(".", 1)[0]

    xlsx = args.fixture
    if not xlsx.exists():
        raise FileNotFoundError(f"Fixture not found: {xlsx}")

    results = _run_baseline_micro(xlsx)
    payload = results_payload(
        results,
        extra={"fixture": try_relative(xlsx, ROOT)},
    )

    args.output_dir.mkdir(parents=True, exist_ok=True)
    out_json = args.output_dir / f"{stem}.json"
    write_json(out_json, payload)

    title = stem.replace("_", " ").title()
    lines = [
        f"# {title}",
        "",
        f"- Fixture workbook: `{payload['fixture']}`",
        "",
        *markdown_results_table(results, include_iterations=True),
        "",
    ]
    out_md = args.output_dir / f"{stem}.md"
    out_md.write_text("\n".join(lines), encoding="utf-8")
    print(f"Wrote {try_relative(out_json, ROOT)}")
    print(f"Wrote {try_relative(out_md, ROOT)}")


def _run_synthetic() -> tuple[list[BenchmarkResult], dict[str, float]]:
    benchmark_results = [
        measure("construct_legacy_2k", lambda: synth.build_legacy(2000), iterations=7),
        measure("construct_typed_2k", lambda: synth.build_typed(2000), iterations=7),
        measure("render_legacy_1200", lambda: synth.render_legacy(1200), iterations=7),
        measure("render_typed_1200", lambda: synth.render_typed(1200), iterations=7),
        measure("parse_legacy_1200", lambda: synth.parse_legacy(1200), iterations=7),
        measure("parse_typed_1200", lambda: synth.parse_typed(1200), iterations=7),
    ]

    by_name = {r.name: r for r in benchmark_results}
    comparisons = {
        "construction_typed_vs_legacy_percent": percent_change(
            by_name["construct_typed_2k"].mean,
            by_name["construct_legacy_2k"].mean,
        ),
        "render_typed_vs_legacy_percent": percent_change(
            by_name["render_typed_1200"].mean,
            by_name["render_legacy_1200"].mean,
        ),
        "parse_typed_vs_legacy_percent": percent_change(
            by_name["parse_typed_1200"].mean,
            by_name["parse_legacy_1200"].mean,
        ),
    }
    return benchmark_results, comparisons


def _cmd_synthetic(args: argparse.Namespace) -> None:
    output_stem = args.output_stem
    if output_stem.endswith((".json", ".md")):
        output_stem = output_stem.rsplit(".", 1)[0]

    args.output_dir.mkdir(parents=True, exist_ok=True)

    benchmark_results, comparisons = _run_synthetic()
    payload = results_payload(benchmark_results, extra={"comparisons": comparisons})

    out_json = args.output_dir / f"{output_stem}.json"
    write_json(out_json, payload)

    title = output_stem.replace("_", " ").title()
    lines = [
        f"# {title}",
        "",
        *markdown_results_table(benchmark_results, float_fmt=".6f"),
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
    out_md = args.output_dir / f"{output_stem}.md"
    out_md.write_text("\n".join(lines), encoding="utf-8")

    print(f"Wrote {try_relative(out_json, ROOT)}")
    print(f"Wrote {try_relative(out_md, ROOT)}")


def _git(repo: Path, *args: str, capture: bool = False) -> str | None:
    if capture:
        return subprocess.check_output(["git", *args], cwd=repo, text=True).strip()
    subprocess.run(["git", *args], cwd=repo, check=True)
    return None


def _resolve_short(repo: Path, ref: str) -> str:
    return _git(repo, "rev-parse", "--short", ref, capture=True) or ref


def _worktree_add(repo: Path, path: Path, ref: str) -> None:
    _git(repo, "worktree", "add", "--detach", str(path), ref)


def _worktree_remove(repo: Path, path: Path) -> None:
    _git(repo, "worktree", "remove", "--force", str(path))


def _find_json(worktree: Path, out_dir: Path, candidates: tuple[str, ...]) -> Path:
    for name in candidates:
        p = out_dir / name
        if p.is_file():
            return p
    for name in candidates:
        p = worktree / "artifacts" / name
        if p.is_file():
            return p
    tried = ", ".join(candidates)
    raise FileNotFoundError(
        f"Expected benchmark JSON ({tried}) under {out_dir} or {worktree / 'artifacts'}."
    )


def _run_suite(worktree: Path, suite: str, out_dir: Path) -> dict:
    spec = _SUITES[suite]
    unified = worktree / "scripts" / "benchmark.py"
    legacy = (
        worktree / "scripts" / "run_phase0_baseline_benchmarks.py"
        if suite == "baseline"
        else worktree / "scripts" / "run_phase1_synthetic_benchmarks.py"
    )
    if unified.is_file():
        script = unified
        cmd: list[str] = [
            "uv",
            "run",
            "python",
            str(script.relative_to(worktree)),
            suite,
            "--output-dir",
            str(out_dir),
        ]
        cmd.extend(spec["extra_args"])
    elif legacy.is_file():
        script = legacy
        cmd = [
            "uv",
            "run",
            "python",
            str(script.relative_to(worktree)),
            "--output-dir",
            str(out_dir),
        ]
        cmd.extend(spec["extra_args"])
    else:
        raise FileNotFoundError(
            f"Neither scripts/benchmark.py nor {legacy.name} found in worktree."
        )

    subprocess.run(cmd, cwd=worktree, check=True)
    candidates = tuple(spec["json_candidates"])
    json_path = _find_json(worktree, out_dir, candidates)
    return json.loads(json_path.read_text(encoding="utf-8"))


def _cmd_compare(args: argparse.Namespace) -> None:
    stem = args.output_stem
    if stem.endswith((".json", ".md")):
        stem = stem.rsplit(".", 1)[0]

    repo = ROOT
    short_a = _resolve_short(repo, args.ref_a)
    short_b = _resolve_short(repo, args.ref_b)

    tmp_root = Path(tempfile.mkdtemp(prefix="fastpyxl-bench-"))
    wt_a = tmp_root / "wt_a"
    wt_b = tmp_root / "wt_b"
    out_a = tmp_root / "out_a"
    out_b = tmp_root / "out_b"

    data_a: dict | None = None
    data_b: dict | None = None
    failed = False
    try:
        _worktree_add(repo, wt_a, args.ref_a)
        _worktree_add(repo, wt_b, args.ref_b)
        out_a.mkdir(parents=True)
        out_b.mkdir(parents=True)
        data_a = _run_suite(wt_a, args.suite, out_a)
        data_b = _run_suite(wt_b, args.suite, out_b)
    except BaseException as e:
        failed = True
        if isinstance(e, subprocess.CalledProcessError):
            print(f"Benchmark subprocess failed: {e}", file=sys.stderr)
        raise
    finally:
        if failed and args.keep_worktrees:
            print(f"Left worktrees under {tmp_root}", file=sys.stderr)
        else:
            for wt in (wt_a, wt_b):
                if wt.exists():
                    try:
                        _worktree_remove(repo, wt)
                    except subprocess.CalledProcessError:
                        pass
            shutil.rmtree(tmp_root, ignore_errors=True)

    if data_a is None or data_b is None:
        raise RuntimeError("No benchmark data collected")

    left = [result_from_dict(r) for r in data_a["results"]]
    right = [result_from_dict(r) for r in data_b["results"]]
    rows = compare_by_name(left, right, left_label="A", right_label="B")

    payload = {
        "ref_a": args.ref_a,
        "ref_b": args.ref_b,
        "short_a": short_a,
        "short_b": short_b,
        "suite": args.suite,
        "run_a": data_a,
        "run_b": data_b,
        "comparison": rows,
    }

    args.output_dir.mkdir(parents=True, exist_ok=True)
    out_json = args.output_dir / f"{stem}.json"
    write_json(out_json, payload)

    title = f"Benchmark compare {short_a} → {short_b}"
    lines = [
        f"# {title}",
        "",
        f"- Suite: `{args.suite}`",
        f"- A (baseline column): `{args.ref_a}` (`{short_a}`)",
        f"- B (comparison column): `{args.ref_b}` (`{short_b}`)",
        "",
        "Δ columns are percent change from A to B: **positive means B is slower**, negative means B is faster.",
        "",
        *markdown_compare_table(rows, left_label="A", right_label="B"),
        "",
    ]
    out_md = args.output_dir / f"{stem}.md"
    out_md.write_text("\n".join(lines), encoding="utf-8")

    print(f"Wrote {try_relative(out_json, ROOT)}")
    print(f"Wrote {try_relative(out_md, ROOT)}")


def _cmd_profile(args: argparse.Namespace) -> None:
    sorts = tuple(args.sorts or ["tottime", "cumtime"])
    args.output_dir.mkdir(parents=True, exist_ok=True)

    for label, builder in (
        ("typed", lambda: synth.build_typed(args.n)),
        ("legacy", lambda: synth.build_legacy(args.n)),
    ):
        pr = profile_repeat(builder, repeat=args.repeat)
        base = args.output_dir / f"construct_profile_{label}"
        write_profile_text(
            base,
            label=label,
            profile=pr,
            sort_keys=sorts,
            top=args.top,
            meta={"n": args.n, "repeat": args.repeat},
        )
        for sk in sorts:
            out = args.output_dir / f"construct_profile_{label}_{sk}.txt"
            print(f"Wrote {try_relative(out, ROOT)}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  uv run python scripts/benchmark.py baseline\n"
            "  uv run python scripts/benchmark.py synthetic\n"
            "  uv run python scripts/benchmark.py compare main HEAD --suite baseline\n"
            "  uv run python scripts/benchmark.py profile\n"
        ),
    )
    sub = parser.add_subparsers(dest="command", required=True)

    p_base = sub.add_parser("baseline", help="Fonts, workbook parse/serialize micro-benchmarks")
    p_base.add_argument(
        "--output-dir",
        type=Path,
        default=ROOT / "artifacts",
        help="Directory for JSON/Markdown (default: ./artifacts)",
    )
    p_base.add_argument(
        "--output-stem",
        default="baseline_benchmarks",
        help="Basename for output files (default: baseline_benchmarks)",
    )
    p_base.add_argument(
        "--fixture",
        type=Path,
        default=DEFAULT_XLSX,
        help="Sample .xlsx path (default: reader test fixture)",
    )
    p_base.set_defaults(func=_cmd_baseline)

    p_syn = sub.add_parser("synthetic", help="Legacy vs typed synthetic serializer benchmarks")
    p_syn.add_argument(
        "--output-dir",
        type=Path,
        default=ROOT / "artifacts",
        help="Directory for JSON/Markdown (default: ./artifacts)",
    )
    p_syn.add_argument(
        "--output-stem",
        default="synthetic_benchmarks",
        help="Basename for output files (default: synthetic_benchmarks)",
    )
    p_syn.set_defaults(func=_cmd_synthetic)

    p_cmp = sub.add_parser(
        "compare",
        help="Run a suite at two git refs (worktrees) and write a comparison report",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Each ref is checked out in a detached worktree. Both commits should include "
            "this CLI (and devtools) for comparable runs; older trees may only have legacy "
            "scripts and JSON names—those filenames are still detected when present under artifacts/."
        ),
    )
    p_cmp.add_argument("ref_a", help="First ref (e.g. main or abc1234)")
    p_cmp.add_argument("ref_b", help="Second ref (e.g. HEAD)")
    p_cmp.add_argument(
        "--suite",
        choices=tuple(_SUITES),
        default="baseline",
        help="Which subcommand suite to run in each worktree (default: baseline)",
    )
    p_cmp.add_argument(
        "--output-dir",
        type=Path,
        default=ROOT / "artifacts",
        help="Where to write compare JSON/Markdown (default: ./artifacts)",
    )
    p_cmp.add_argument(
        "--output-stem",
        default="benchmark_compare",
        help="Basename for compare outputs (default: benchmark_compare)",
    )
    p_cmp.add_argument(
        "--keep-worktrees",
        action="store_true",
        help="On failure, do not remove temporary worktrees (path printed to stderr)",
    )
    p_cmp.set_defaults(func=_cmd_compare)

    p_prof = sub.add_parser("profile", help="cProfile synthetic model construction (typed vs legacy)")
    p_prof.add_argument("--n", type=int, default=2000, help="parents per build (default 2000)")
    p_prof.add_argument("--repeat", type=int, default=12, help="times to run each build loop")
    p_prof.add_argument(
        "--sort",
        action="append",
        choices=("tottime", "cumtime"),
        dest="sorts",
        help="pstats sort key (default: tottime and cumtime)",
    )
    p_prof.add_argument("--top", type=int, default=45, help="print_stats line limit")
    p_prof.add_argument(
        "--output-dir",
        type=Path,
        default=ROOT / "artifacts",
        help="Directory for profile text files (default: ./artifacts)",
    )
    p_prof.set_defaults(func=_cmd_profile)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
