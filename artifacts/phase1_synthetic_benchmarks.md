# Phase 1 Synthetic Benchmarks

| Benchmark | Mean (s) | Median (s) | StdDev (s) |
|---|---:|---:|---:|
| `construct_legacy_2k` | 0.014976 | 0.013010 | 0.004418 |
| `construct_typed_2k` | 0.013889 | 0.012259 | 0.004510 |
| `render_legacy_1200` | 0.044937 | 0.042874 | 0.002952 |
| `render_typed_1200` | 0.036047 | 0.035986 | 0.001376 |
| `parse_legacy_1200` | 0.055850 | 0.054547 | 0.003177 |
| `parse_typed_1200` | 0.049823 | 0.049042 | 0.003963 |

## Typed vs Legacy

- Construction delta: -7.25%
- Render delta: -19.78%
- Parse delta: -10.79%

Positive means typed is slower; negative means typed is faster.
