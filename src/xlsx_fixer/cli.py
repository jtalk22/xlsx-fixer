"""xlsx-fixer CLI: Fix Mac Excel corruption in openpyxl-generated files.

Usage:
    xlsx-fixer fix <file> [--output <file>] [--keep-calc-chain]
    xlsx-fixer check <file>
    xlsx-fixer --version
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from xlsx_fixer import __version__
from xlsx_fixer.fixer import Issue, check, fix


def _severity_marker(severity: str) -> str:
    if severity == "error":
        return "ERROR"
    return "WARN "


def cmd_check(args: argparse.Namespace) -> int:
    """Check a file for corruption issues."""
    path = Path(args.file)
    try:
        issues = check(path)
    except (FileNotFoundError, ValueError) as e:
        print(f"  Error: {e}", file=sys.stderr)
        return 1

    if not issues:
        print(f"  {path.name}: No issues found")
        return 0

    errors = sum(1 for i in issues if i.severity == "error")
    warnings = sum(1 for i in issues if i.severity == "warning")

    print(f"  {path.name}: {errors} error(s), {warnings} warning(s)")
    print()
    for issue in issues:
        marker = _severity_marker(issue.severity)
        print(f"  [{marker}] {issue.code}: {issue.message}")
        if issue.detail:
            for line in issue.detail.split(". "):
                print(f"          {line.strip()}.")
        print()

    return 1 if errors > 0 else 0


def cmd_fix(args: argparse.Namespace) -> int:
    """Fix corruption issues in a file."""
    path = Path(args.file)
    output = Path(args.output) if args.output else None

    try:
        stats = fix(
            path,
            output=output,
            remove_calc_chain=not args.keep_calc_chain,
        )
    except (FileNotFoundError, ValueError) as e:
        print(f"  Error: {e}", file=sys.stderr)
        return 1

    target = output or path
    fixes_applied = []

    if stats["inline_strings"] > 0:
        fixes_applied.append(
            f"Converted {stats['inline_strings']} inline strings -> "
            f"{stats['unique_strings']} unique shared strings"
        )

    if stats["full_calc_removed"]:
        fixes_applied.append("Removed fullCalcOnLoad (inconsistent calc state)")

    if stats["calc_chain_removed"]:
        fixes_applied.append("Removed stale calcChain.xml")

    if stats["control_chars_stripped"] > 0:
        fixes_applied.append(
            f"Stripped {stats['control_chars_stripped']} illegal control character(s)"
        )

    if fixes_applied:
        print(f"  Fixed: {target.name}")
        for f in fixes_applied:
            print(f"    - {f}")
    else:
        print(f"  {target.name}: No fixes needed")

    return 0


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="xlsx-fixer",
        description="Fix Mac Excel corruption in openpyxl-generated .xlsx files",
    )
    parser.add_argument(
        "--version", action="version", version=f"xlsx-fixer {__version__}"
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    # fix command
    fix_parser = subparsers.add_parser(
        "fix", help="Fix corruption issues in an xlsx file"
    )
    fix_parser.add_argument("file", help="Path to the xlsx file to fix")
    fix_parser.add_argument(
        "-o", "--output", help="Output path (default: fix in-place)"
    )
    fix_parser.add_argument(
        "--keep-calc-chain",
        action="store_true",
        help="Don't remove calcChain.xml",
    )

    # check command
    check_parser = subparsers.add_parser(
        "check", help="Check an xlsx file for corruption issues"
    )
    check_parser.add_argument("file", help="Path to the xlsx file to check")

    args = parser.parse_args()

    if args.command == "fix":
        sys.exit(cmd_fix(args))
    elif args.command == "check":
        sys.exit(cmd_check(args))


if __name__ == "__main__":
    main()
