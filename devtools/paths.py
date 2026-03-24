"""Small path helpers for CLI scripts."""

from __future__ import annotations

from pathlib import Path


def try_relative(path: Path, root: Path) -> str:
    """Return ``path`` relative to ``root`` when possible, else the absolute path string."""
    try:
        return str(path.relative_to(root))
    except ValueError:
        return str(path)
