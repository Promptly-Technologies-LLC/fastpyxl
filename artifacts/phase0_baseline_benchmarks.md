# Phase 0 Baseline Benchmarks

- Fixture workbook: `fastpyxl/reader/tests/data/sample.xlsx`

| Benchmark | Mean (s) | Median (s) | StdDev (s) | Iterations |
|---|---:|---:|---:|---:|
| `object_construction_100k` | 1.91021108 | 1.82604062 | 0.14247283 | 5 |
| `attribute_read_300k` | 0.06757152 | 0.06775566 | 0.00122573 | 5 |
| `attribute_assignment_150k` | 0.33453520 | 0.33174346 | 0.01830035 | 5 |
| `parse_representative_workbook` | 0.05411865 | 0.05398494 | 0.00258218 | 5 |
| `serialize_representative_workbook` | 0.01350043 | 0.01168228 | 0.00308721 | 5 |
| `load_save_roundtrip` | 0.09314495 | 0.09390240 | 0.00181239 | 3 |
