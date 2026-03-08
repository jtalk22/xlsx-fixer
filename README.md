# xlsx-fixer

**Fix the Mac Excel "We found a problem with some content" error in openpyxl-generated .xlsx files.**

openpyxl hardcodes inline strings (`t="inlineStr"`) for every text cell. Mac Excel's strict OOXML parser rejects this, showing a recovery dialog on *every open*. Windows Excel silently accepts it, which is why the bug goes unnoticed during development.

`xlsx-fixer` rewrites the ZIP to use a proper shared string table (`xl/sharedStrings.xml`), removes inconsistent calc state, and strips illegal control characters. One function call. Zero dependencies beyond the standard library.

## The Problem

If you generate `.xlsx` files with Python's `openpyxl` library and open them on Mac, you see:

> **"We found a problem with some content in 'file.xlsx'. Do you want us to try to recover as much as we can?"**

After clicking "Yes":

> **"Removed Records: String properties from /xl/sharedStrings.xml part"**

This happens because openpyxl writes every text cell as an inline string (`<c t="inlineStr"><is><t>text</t></is></c>`) instead of referencing a shared string table. The openpyxl maintainer has [declined to fix this](https://foss.heptapod.net/openpyxl/openpyxl/-/issues/1804) for 17+ years.

## Install

```bash
pip install xlsx-fixer
```

**Zero dependencies.** Uses only Python standard library (`zipfile`, `xml.etree.ElementTree`).

## Usage

### Python API

```python
from xlsx_fixer import fix, check

# Fix in-place (after openpyxl wb.save())
from openpyxl import Workbook
wb = Workbook()
ws = wb.active
ws["A1"] = "Hello, Mac Excel!"
wb.save("report.xlsx")

fix("report.xlsx")  # That's it. Mac-safe now.

# Fix to a new file
fix("report.xlsx", output="report_fixed.xlsx")

# Check without modifying
issues = check("report.xlsx")
for issue in issues:
    print(f"[{issue.severity}] {issue.code}: {issue.message}")
```

### CLI

```bash
# Fix in-place
xlsx-fixer fix report.xlsx

# Fix to new file
xlsx-fixer fix report.xlsx -o fixed.xlsx

# Check without modifying
xlsx-fixer check report.xlsx

# Version
xlsx-fixer --version
```

### Integration with openpyxl

Add two lines to your existing code:

```python
from openpyxl import Workbook
from xlsx_fixer import fix

wb = Workbook()
# ... your existing workbook code ...

wb.calculation.fullCalcOnLoad = None  # Remove inconsistent calc state
wb.save("output.xlsx")
fix("output.xlsx")  # Convert inline strings to shared string table
```

## What It Fixes

| # | Issue | Root Cause | Impact |
|---|-------|-----------|--------|
| 1 | **Inline strings** | openpyxl hardcodes `t="inlineStr"` in `cell/_writer.py` | "We found a problem" dialog on every Mac open |
| 2 | **fullCalcOnLoad** | openpyxl sets `fullCalcOnLoad="1"` without generating `calcChain.xml` | Recovery dialog; formulas show 0 |
| 3 | **Stale calcChain.xml** | References cells that no longer contain formulas | "Removed Records: Formula" in repair log |
| 4 | **Control characters** | Illegal XML 1.0 chars (U+0000-U+0008, etc.) in cell values | ST_Xstring validation failures |

## How It Works

1. Reads the entire `.xlsx` ZIP into memory
2. Parses every `xl/worksheets/sheet*.xml`
3. Finds all `<c t="inlineStr"><is><t>TEXT</t></is></c>` cells
4. Builds a deduplicated shared string table
5. Rewrites each cell as `<c t="s"><v>INDEX</v></c>`
6. Creates `xl/sharedStrings.xml` with `<sst>` root element
7. Adds relationship to `xl/_rels/workbook.xml.rels`
8. Adds Override to `[Content_Types].xml`
9. Removes `fullCalcOnLoad` from `<calcPr>` in `xl/workbook.xml`
10. Optionally removes stale `calcChain.xml`
11. Writes new ZIP back to the same (or specified output) path

The fix is **idempotent** — running it twice is safe and the second run is a no-op.

## Performance

The fix operates entirely in-memory. On a real production workbook (11 sheets, 195 properties, 448 verification checks, ~100KB):

- **~30ms** on Apple Silicon
- **~50ms** on Intel Mac
- Scales linearly with file size

## Why Not Just Fix openpyxl?

The maintainer has permanently refused to add shared string table support to the writer. The inline string behavior is hardcoded in `cell/_writer.py` lines 21-22 and 70-79 with no API flag to change it. This isn't a bug they plan to fix — it's a design decision from 2007.

With 249M+ openpyxl downloads per month and 7,400+ dependent packages, this affects an enormous number of Python developers who generate Excel files for Mac users.

## Tested On

- Python 3.9 - 3.14
- openpyxl 3.0.x - 3.1.x
- Mac Excel 16.x (Microsoft 365)
- Windows Excel (passes through unchanged)
- LibreOffice Calc (passes through unchanged)

## License

MIT License. See [LICENSE](LICENSE) for details.
