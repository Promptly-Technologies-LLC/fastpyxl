Always run Python code with `uv run`.

Practice TDD. First write RED phase tests and watch them fail for the right reason, then turn them GREEN.

Before pushing new optimizations to the remote, run microbenchmarks to ensure performance has actually improved.

## Cursor Cloud specific instructions

`fastpyxl` is a pure-Python library (a drop-in replacement for `openpyxl`) with no runtime services, web server, database, or CLI. "Running" it means importing the API or exercising the test/benchmark scripts.

- Dependencies are managed by `uv`. The startup update script runs `uv sync --all-groups`, so the `.venv` with the full `dev` + `docs` groups is already provisioned. Prefix all Python commands with `uv run`.
- Tests: `uv run pytest` (config lives in `pyproject.toml`; warnings are errors and markers are strict). Optional-dependency tests auto-skip via markers in `conftest.py`, but the update script installs all optional deps so nothing skips for that reason.
- Lint/type-check: `uv run ruff check .` and `uv run ty check`.
- openpyxl drop-in contract: `uv run pytest fastpyxl/tests/compat -m openpyxl_compat`.
- Benchmarks/profiling scripts live in `devtools/` and `scripts/`; run them with `uv run` (e.g. `uv run python devtools/benchmark.py`).