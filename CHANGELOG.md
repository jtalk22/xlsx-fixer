# Changelog

## 1.1.0 (2026-03-09)

- `fix_batch()` — Fix all xlsx files in a directory (in-place or to output dir)
- `batch` CLI command — `xlsx-fixer batch <dir> --key <KEY>` (licensed feature)
- License key validation via Lemon Squeezy with activation caching
  - 7-day cache TTL, 30-day offline grace period
  - Supports `--key`, `--key-file`, and default `~/.xlsx-fixer/license-key`
  - Zero new dependencies (urllib.request)
- Recursive directory scanning with `--recursive` flag
- Skips Excel temp files (~$ prefix)
- 60 tests (27 → 60)

## 1.0.0 (2026-03-08)

Initial release.

- `fix()` — Fix openpyxl inline strings, fullCalcOnLoad, calcChain.xml, control chars
- `check()` — Non-destructive corruption detection with Issue objects
- CLI: `xlsx-fixer fix`, `xlsx-fixer check`
- Zero dependencies beyond Python standard library
- Tested on Python 3.9-3.14, openpyxl 3.0.x-3.1.x, Mac Excel 16.x

Derived from 34 versions of a production Excel generator (448 verification
checks, 11 sheets, 195 properties).
