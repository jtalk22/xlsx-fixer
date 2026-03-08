"""xlsx-fixer: Fix Mac Excel corruption in openpyxl-generated .xlsx files.

openpyxl hardcodes inline strings (t="inlineStr") which Mac Excel's strict
OOXML parser rejects with "We found a problem with some content." This package
rewrites the ZIP to use a proper shared string table (xl/sharedStrings.xml).

Usage:
    from xlsx_fixer import fix, check

    fix("report.xlsx")                     # Fix in-place
    fix("report.xlsx", output="fixed.xlsx") # Fix to new file

    issues = check("report.xlsx")          # Check without modifying
    for issue in issues:
        print(f"[{issue.severity}] {issue.code}: {issue.message}")

CLI:
    xlsx-fixer fix report.xlsx
    xlsx-fixer check report.xlsx
"""

from xlsx_fixer.fixer import fix, check, Issue

__version__ = "1.0.0"
__all__ = ["fix", "check", "Issue", "__version__"]
