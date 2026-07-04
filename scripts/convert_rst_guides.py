"""Convert legacy Sphinx RST guides to Quarto QMD for great-docs."""

from __future__ import annotations

import re
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DOC = ROOT / "doc"
OUT = ROOT / "user_guide"

GUIDES: list[tuple[str, str, str, str]] = [
    ("Introduction", "01", "tutorial.rst", "Tutorial"),
    ("Introduction", "02", "usage.rst", "Usage"),
    ("Styling", "10", "styles.rst", "Styles"),
    ("Styling", "11", "rich_text.rst", "Rich Text"),
    ("Styling", "12", "formatting.rst", "Conditional Formatting"),
    ("Worksheets", "20", "editing_worksheets.rst", "Editing Worksheets"),
    ("Worksheets", "21", "worksheet_properties.rst", "Worksheet Properties"),
    ("Worksheets", "22", "validation.rst", "Data Validation"),
    ("Worksheets", "23", "worksheet_tables.rst", "Worksheet Tables"),
    ("Worksheets", "24", "filters.rst", "Filters"),
    ("Worksheets", "25", "print_settings.rst", "Print Settings"),
    ("Worksheets", "26", "pivot.rst", "Pivot Tables"),
    ("Worksheets", "27", "comments.rst", "Comments"),
    ("Worksheets", "28", "datetime.rst", "Dates and Times"),
    ("Worksheets", "29", "simple_formulae.rst", "Simple Formulae"),
    ("Workbooks", "30", "defined_names.rst", "Defined Names"),
    ("Workbooks", "31", "workbook_custom_doc_props.rst", "Custom Document Properties"),
    ("Workbooks", "32", "protection.rst", "Protection"),
    ("Charts", "40", "charts/introduction.rst", "Charts"),
    ("Charts", "41", "charts/anchors.rst", "Chart Anchors"),
    ("Charts", "42", "charts/area.rst", "Area Charts"),
    ("Charts", "43", "charts/bar.rst", "Bar Charts"),
    ("Charts", "44", "charts/bubble.rst", "Bubble Charts"),
    ("Charts", "45", "charts/chart_layout.rst", "Chart Layout"),
    ("Charts", "46", "charts/chartsheet.rst", "Chartsheets"),
    ("Charts", "47", "charts/doughnut.rst", "Doughnut Charts"),
    ("Charts", "48", "charts/gauge.rst", "Gauge Charts"),
    ("Charts", "49", "charts/graphical.rst", "Graphical Properties"),
    ("Charts", "50", "charts/line.rst", "Line Charts"),
    ("Charts", "51", "charts/limits_and_scaling.rst", "Limits and Scaling"),
    ("Charts", "52", "charts/pattern.rst", "Pattern Fills"),
    ("Charts", "53", "charts/pie.rst", "Pie Charts"),
    ("Charts", "54", "charts/radar.rst", "Radar Charts"),
    ("Charts", "55", "charts/scatter.rst", "Scatter Charts"),
    ("Charts", "56", "charts/secondary.rst", "Secondary Axes"),
    ("Charts", "57", "charts/stock.rst", "Stock Charts"),
    ("Charts", "58", "charts/surface.rst", "Surface Charts"),
    ("Images", "60", "images.rst", "Images"),
    ("Pandas", "70", "pandas.rst", "Pandas Integration"),
    ("Performance", "80", "optimized.rst", "Optimized Modes"),
    ("Performance", "81", "performance.rst", "Performance"),
    ("Developers", "90", "development.rst", "Development"),
    ("Developers", "91", "formula.rst", "Formula Tokenizer"),
    ("Release Notes", "95", "changes.rst", "Release Notes"),
]

RE_SPHINX = [
    (re.compile(r":doc:`([^`]+)`"), r"[\1](\1.qmd)"),
    (re.compile(r":ref:`([^`]+)`"), r"\1"),
    (re.compile(r"\|release\|"), "1.0.9"),
    (re.compile(r"\|version\|"), "1.0"),
    (re.compile(r"\|today\|"), "2026-07-04"),
    (re.compile(r"\.\. parsed-literal::"), ".. code-block:: text"),
]


def convert_rst(src: Path) -> str:
    result = subprocess.run(
        ["pandoc", "-f", "rst", "-t", "markdown", str(src)],
        check=True,
        capture_output=True,
        text=True,
    )
    body = result.stdout
    for pattern, repl in RE_SPHINX:
        body = pattern.sub(repl, body)
    return body.strip() + "\n"


def main() -> None:
    if OUT.exists():
        for path in OUT.rglob("*.qmd"):
            path.unlink()
    else:
        OUT.mkdir()

    assets = OUT / "images"
    assets.mkdir(exist_ok=True)
    for name in ("logo.svg", "benchmark_chart.svg"):
        src = DOC / name
        if src.exists():
            (assets / name).write_bytes(src.read_bytes())

    for section, prefix, rel_path, title in GUIDES:
        src = DOC / rel_path
        slug = Path(rel_path).stem
        dest = OUT / f"{prefix}-{slug}.qmd"
        body = convert_rst(src)
        frontmatter = f'---\ntitle: "{title}"\nguide-section: "{section}"\n---\n\n'
        dest.write_text(frontmatter + body, encoding="utf-8")
        print(f"Wrote {dest.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
