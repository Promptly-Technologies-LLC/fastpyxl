"""Validate conventional-commit subjects used by python-semantic-release.

Merge commits are ignored by PSR (`ignore_merge_commits = true`). Only
`feat` / `fix` / `perf` subjects bump a version (`default_bump_level = 0`).
Prose titles such as ``Add keep_formula_cache…`` merge CI-green and skip PyPI.
"""

from __future__ import annotations

import argparse
import subprocess
import sys
from dataclasses import dataclass
from re import compile as regexp

MINOR_TAGS = ("feat",)
PATCH_TAGS = ("fix", "perf")
OTHER_ALLOWED_TAGS = (
    "build",
    "chore",
    "ci",
    "docs",
    "style",
    "refactor",
    "test",
)
ALLOWED_TAGS = (*MINOR_TAGS, *PATCH_TAGS, *OTHER_ALLOWED_TAGS)

_TYPE_PATTERN = regexp(r"(?P<type>%s)" % "|".join(ALLOWED_TAGS))
_SUBJECT_PATTERN = regexp(
    r"^"
    + _TYPE_PATTERN.pattern
    + r"(?:\((?P<scope>[^\n]+)\))?"
    + r"(?P<break>!)?:\s+"
    + r"(?P<subject>[^\n]+)"
)


@dataclass(frozen=True, slots=True)
class ParsedSubject:
    type: str
    bumps_version: bool


def parse_subject(message: str) -> ParsedSubject | None:
    """Parse the first line of a commit message or PR title."""
    first_line = message.split("\n", 1)[0].strip()
    match = _SUBJECT_PATTERN.match(first_line)
    if match is None:
        return None
    commit_type = match.group("type")
    breaking = bool(match.group("break"))
    return ParsedSubject(
        type=commit_type,
        bumps_version=breaking or commit_type in MINOR_TAGS or commit_type in PATCH_TAGS,
    )


def validate_subjects(subjects: tuple[str, ...] | list[str]) -> list[str]:
    """Return one error string per subject that is not conventional."""
    errors: list[str] = []
    for subject in subjects:
        text = subject.strip()
        if not text:
            continue
        if parse_subject(text) is None:
            errors.append(f"not a conventional commit: {text!r}")
    return errors


def _subjects_from_git_range(git_range: str) -> tuple[str, ...]:
    result = subprocess.run(
        ["git", "log", "--format=%s", "--no-merges", git_range],
        check=True,
        capture_output=True,
        text=True,
    )
    return tuple(line for line in result.stdout.splitlines() if line.strip())


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Fail if commit subjects / PR titles are not conventional.",
    )
    parser.add_argument(
        "--subject",
        action="append",
        default=[],
        help="A single subject (repeatable). Use for PR titles.",
    )
    parser.add_argument(
        "--git-range",
        help="git revision range whose non-merge subjects are checked, e.g. origin/master..HEAD",
    )
    args = parser.parse_args(argv)

    subjects: list[str] = list(args.subject)
    if args.git_range:
        subjects.extend(_subjects_from_git_range(args.git_range))
    if not subjects:
        parser.error("provide --subject and/or --git-range")

    errors = validate_subjects(subjects)
    if not errors:
        return 0
    for error in errors:
        print(error, file=sys.stderr)
    print(
        "Use feat:/fix:/perf: to bump PyPI; other allowed types: "
        + ", ".join(OTHER_ALLOWED_TAGS),
        file=sys.stderr,
    )
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
