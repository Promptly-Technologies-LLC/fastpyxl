# Phase 8 Synthetic Benchmarks

| Benchmark | Mean (s) | Median (s) | StdDev (s) |
|---|---:|---:|---:|
| `construct_legacy_2k` | 0.013196 | 0.010724 | 0.006171 |
| `construct_typed_2k` | 0.023909 | 0.021726 | 0.009380 |
| `render_legacy_1200` | 0.055418 | 0.058128 | 0.010597 |
| `render_typed_1200` | 0.038937 | 0.036416 | 0.006912 |
| `parse_legacy_1200` | 0.038549 | 0.037838 | 0.001752 |
| `parse_typed_1200` | 0.038201 | 0.037635 | 0.001104 |

## Typed vs Legacy

- Construction delta: 81.18%
- Render delta: -29.74%
- Parse delta: -0.90%

Positive means typed is slower; negative means typed is faster.
