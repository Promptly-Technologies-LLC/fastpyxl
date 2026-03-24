"""Helpers around :mod:`cProfile` and :mod:`pstats` for hot-path investigation."""

from __future__ import annotations

from collections.abc import Callable
from io import StringIO
from pathlib import Path
from typing import Any

from cProfile import Profile


def format_pstats(profile: Profile, sort_key: str, limit: int) -> str:
    """Return ``pstats.print_stats`` text for ``profile``."""
    import pstats

    s = StringIO()
    stats = pstats.Stats(profile, stream=s)
    stats.strip_dirs()
    stats.sort_stats(sort_key)
    stats.print_stats(limit)
    return s.getvalue()


def profile_repeat(
    fn: Callable[[], Any],
    *,
    repeat: int,
) -> Profile:
    """Enable profiling, call ``fn`` ``repeat`` times, return the :class:`Profile`."""
    pr = Profile()
    pr.enable()
    for _ in range(repeat):
        fn()
    pr.disable()
    return pr


def write_profile_text(
    path: Path,
    *,
    label: str,
    profile: Profile,
    sort_keys: tuple[str, ...],
    top: int,
    meta: dict[str, Any] | None = None,
) -> None:
    """Write one text file per ``sort_key`` with a short header."""
    path.parent.mkdir(parents=True, exist_ok=True)
    meta = meta or {}
    for sort_key in sort_keys:
        body = format_pstats(profile, sort_key, top)
        parts = [
            f"# profile ({label})",
            "# " + " ".join(f"{k}={v}" for k, v in sorted(meta.items())),
            f"# sort={sort_key} top={top}",
            "",
            body,
        ]
        out = path.parent / f"{path.stem}_{sort_key}{path.suffix or '.txt'}"
        out.write_text("\n".join(parts), encoding="utf-8")
