# CHANGELOG

<!-- version list -->

## v1.0.6 (2026-03-30)

### Performance Improvements

- Optimize formula translation hot path (fixes #25)
  ([`359d0bd`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/359d0bdd1578c5c82a18333a3c537a3c6abcf0fa))


## v1.0.5 (2026-03-30)

### Performance Improvements

- Reduce per-call overhead in parse_cell (fixes #27)
  ([`05b8cde`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/05b8cde4e283351632c39f5144742f56f7e0b8fc))


## v1.0.4 (2026-03-30)

### Performance Improvements

- Cache sheetnames and eliminate list allocs in __contains__/__getitem__ (fixes #36)
  ([`5f287bc`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/5f287bc14070b896ea28532b8f8a94c656685733))

### Testing

- Add regression tests for sheetnames cache invalidation
  ([#36](https://github.com/Promptly-Technologies-LLC/fastpyxl/pull/36),
  [`d215181`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/d2151819dc8b448d8eefbc828228fa18a103aa20))


## v1.0.3 (2026-03-30)

### Performance Improvements

- Skip find/findtext for empty cells in parse_cell (fixes #30)
  ([`c15d959`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/c15d959bd177d503d029d93f0435211ede85db26))


## v1.0.2 (2026-03-30)

### Performance Improvements

- Skip row re-parse in parse_cell by using row_counter (fixes #31)
  ([`4881180`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/4881180c471c464029b1938326509d4e8b765658))


## v1.0.1 (2026-03-30)

### Documentation

- Rebrand documentation from openpyxl to fastpyxl
  ([`b07ecab`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/b07ecabd5ee3628b81510881c5b34db642859639))

- Rewrite README with benchmarks, rebrand docs for fastpyxl
  ([`47421c3`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/47421c331ab6b33bebcaf2097ecec91f8d7a38c7))

### Performance Improvements

- Filter vba_archive to VBA-relevant files only (fixes #29)
  ([`5da5b0a`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/5da5b0a2ce5f63d4521cabad4838c5f655d89eee))


## v1.0.0 (2026-03-29)

- Initial Release
