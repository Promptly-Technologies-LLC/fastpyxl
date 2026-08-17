"""Tests for conventional-commit subject checks used by release CI."""

from __future__ import annotations

import pytest

from devtools.conventional_commits import (
    ALLOWED_TAGS,
    PATCH_TAGS,
    MINOR_TAGS,
    parse_subject,
    validate_subjects,
)

PROSE_SUBJECTS = (
    "Add keep_formula_cache dual-load for formulas and caches",
    "Store formula caches on a worksheet side map",
    "Fix formula-cache side-map sync for detached cells and moves",
    "Add access=\"indexed\" for on-demand workbook reads",
    "Merge pull request #72 from Promptly-Technologies-LLC/cursor/keep-formula-cache-a543",
)


@pytest.mark.parametrize(
    ("subject", "expected_type", "bumps"),
    [
        ("feat: keep_formula_cache dual-load and indexed access", "feat", True),
        ("feat(reader): dual-load formulas and caches", "feat", True),
        ("feat!: breaking load_workbook change", "feat", True),
        ("fix: accept a:spLocks noTextEdit via ShapeLocking (#69)", "fix", True),
        ("perf: speed up worksheet parse", "perf", True),
        ("ci: require conventional PR titles", "ci", False),
        ("docs: document keep_formula_cache", "docs", False),
        ("test: cover formula-cache side map", "test", False),
        ("chore: ignore merge commits in release notes", "chore", False),
    ],
)
def test_conventional_subjects_parse(subject: str, expected_type: str, bumps: bool) -> None:
    parsed = parse_subject(subject)
    assert parsed is not None
    assert parsed.type == expected_type
    assert parsed.bumps_version is bumps
    assert expected_type in ALLOWED_TAGS


@pytest.mark.parametrize("subject", PROSE_SUBJECTS)
def test_prose_subjects_do_not_parse(subject: str) -> None:
    assert parse_subject(subject) is None


def test_validate_subjects_rejects_prose() -> None:
    errors = validate_subjects(PROSE_SUBJECTS)
    assert len(errors) == len(PROSE_SUBJECTS)
    assert all("not a conventional commit" in err for err in errors)


def test_validate_subjects_accepts_conventional() -> None:
    assert validate_subjects(
        (
            "feat: keep_formula_cache dual-load and indexed access",
            "ci: require conventional commit subjects",
        )
    ) == []


def test_bump_tag_sets_match_psr_config() -> None:
    assert MINOR_TAGS == ("feat",)
    assert PATCH_TAGS == ("fix", "perf")
    assert "feat" in ALLOWED_TAGS
    assert "fix" in ALLOWED_TAGS
    assert "ci" in ALLOWED_TAGS
