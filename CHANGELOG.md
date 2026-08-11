# CHANGELOG

<!-- version list -->

## v1.0.12 (2026-08-08)

### Bug Fixes

- Accept a:spLocks noTextEdit via ShapeLocking
  ([#69](https://github.com/Promptly-Technologies-LLC/fastpyxl/pull/69),
  [`809a692`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/809a69252a6d1d9f07d67b36baa1d16f02caf7a6))

### Documentation

- Formalize openpyxl >=3.0 compatibility contract
  ([`2f1ff39`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/2f1ff39fbbdb2c641ba5f823cd0ef5a0ef25043a))

### Testing

- Add load_workbook kwargs behavioral parity coverage
  ([`87e112b`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/87e112b57d3eca5f9a337de2c2e574a53e9583e8))

- Add write_only mode parity coverage to compat suite
  ([`11ee797`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/11ee7978167318c6755b9ad00c8aebb5df3445cf))

- Cover chart sibling parse with noTextEdit locks
  ([`7fceb60`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/7fceb60145779c71e28a4eb93face58e3b7f9b96))

- Defer openpyxl imports in rich-text fixture setup
  ([`6adb4f2`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/6adb4f2693d308a5b3b151e8afdd11dfeb806b94))

- Ignore unset openpyxl color descriptors in snapshots
  ([`2df7238`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/2df7238114a376aa0314eff8be6631b2996d7294))

- Normalize ArrayFormula and CellRichText in compat snapshots
  ([`b3b70b8`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/b3b70b8e37dbbda1121dbe5088bec7d699c54741))


## Unreleased

### Features

- Add `keep_formula_cache=True` to dual-load formulas and cached values in
  one parse (fixes #71)

- Add `keep_vba="full"` escape hatch and document filtered `vba_archive`
  semantics for default `keep_vba=True` (fixes #67)

## v1.0.11 (2026-08-07)

### Bug Fixes

- Load x14 data validations from worksheet extLst
  ([`9234030`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/9234030fc87f8ed0c5fa95ec8fbdf10bf905c616))

### Build System

- Upgrade ruff to 0.16 and enable ISC004
  ([`b1c01d6`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/b1c01d6379675e78a1b3a3272569c4de8c316517))

### Continuous Integration

- Add a lint job running ruff and ty
  ([`e60c56d`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/e60c56d72159f75719e29a4f55261af194f23182))

### Documentation

- Add Cursor Cloud environment setup notes to AGENTS.md
  ([`387192f`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/387192f15104694a2ea40d28d695346150873901))


## v1.0.10 (2026-07-04)

### Bug Fixes

- Suppress benign data validation extension warnings on load
  ([`4b54d98`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/4b54d9810b34458f1c3869741ce419f12a323cb3))


## v1.0.9 (2026-04-22)

### Bug Fixes

- Derive workbook content type from filename extension (fixes #46)
  ([`53bfb23`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/53bfb236a302dc4fd0671f5cd507e8a173d688be))


## v1.0.8 (2026-04-06)

### Bug Fixes

- Allow multi-digit vmlDrawing indices in ARC_VBA regex (fixes #44)
  ([`070a7e5`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/070a7e5fbcface591d1bf9d4e3e10c384e3ff8bf))

### Documentation

- Update benchmark chart and README with fresh v1.0.7 numbers (fixes #42)
  ([`f010081`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/f010081a828371736b39466108d94c23598a6037))


## v1.0.7 (2026-03-30)

### Bug Fixes

- Handle unhashable keys in __getitem__ and __contains__
  ([`080959a`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/080959a1a9ee7010a5d880aac15505b03cdabf0f))

### Performance Improvements

- O(1) sheet lookup via cached title map (fixes #40)
  ([`2d6b3e6`](https://github.com/Promptly-Technologies-LLC/fastpyxl/commit/2d6b3e6d521110bcb2612f815463805d0001a486))


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
