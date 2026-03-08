# Changelog

## 1.0.0 (2026-03-08)

Initial release.

- `fix()` — Fix openpyxl inline strings, fullCalcOnLoad, calcChain.xml, control chars
- `check()` — Non-destructive corruption detection with Issue objects
- CLI: `xlsx-fixer fix`, `xlsx-fixer check`
- Zero dependencies beyond Python standard library
- Tested on Python 3.9-3.14, openpyxl 3.0.x-3.1.x, Mac Excel 16.x

Derived from 34 versions of a production Excel generator (448 verification
checks, 11 sheets, 195 properties).
