"""Core xlsx-fixer logic: detect and repair openpyxl corruption patterns.

Five corruption patterns are addressed:
1. Inline strings (t="inlineStr") -> shared string table conversion
2. fullCalcOnLoad="1" without calcChain.xml -> inconsistent calc state
3. Stale calcChain.xml referencing non-existent formula cells
4. Table + ConditionalFormatting overlap -> invalid OOXML
5. Control characters in string values -> ST_Xstring violations

Derived from 34 versions of a production Excel generator (448 verification
checks, 11 sheets, 195 properties) tested on Mac Excel 16.x.
"""

from __future__ import annotations

import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Union

# OOXML namespace constants
SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
SST_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
)
SST_CT = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
)

# Register namespaces to prevent ns0: prefix mangling in output XML.
# CRITICAL: Must happen before any ET.fromstring() / ET.parse() calls.
ET.register_namespace("", SHEET_NS)
ET.register_namespace(
    "r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)

# Characters illegal in OOXML ST_Xstring (XML 1.0 restricted chars)
_ILLEGAL_XML_RE = re.compile(
    "[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x84\x86-\x9f]"
)


@dataclass
class Issue:
    """A detected corruption issue in an xlsx file."""

    code: str
    severity: str  # "error" or "warning"
    message: str
    detail: str = ""


def check(path: Union[str, Path]) -> list[Issue]:
    """Check an xlsx file for known corruption patterns without modifying it.

    Returns a list of Issue objects describing problems found.
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    if not zipfile.is_zipfile(path):
        raise ValueError(f"Not a valid ZIP/xlsx file: {path}")

    issues: list[Issue] = []

    with zipfile.ZipFile(path, "r") as zf:
        names = zf.namelist()
        sheet_names = sorted(
            n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml", n)
        )

        # Check 1: Inline strings
        inline_count = 0
        for sheet_name in sheet_names:
            data = zf.read(sheet_name)
            inline_count += data.count(b"inlineStr")

        if inline_count > 0:
            issues.append(
                Issue(
                    code="INLINE_STRINGS",
                    severity="error",
                    message=f"Found {inline_count} inline string references across {len(sheet_names)} sheets",
                    detail=(
                        "openpyxl hardcodes t=\"inlineStr\" for all text cells. "
                        "Mac Excel rejects this with \"We found a problem with some content.\""
                    ),
                )
            )

        has_sst = "xl/sharedStrings.xml" in names
        if not has_sst and inline_count > 0:
            issues.append(
                Issue(
                    code="MISSING_SST",
                    severity="error",
                    message="No xl/sharedStrings.xml found",
                    detail="File relies entirely on inline strings with no shared string table.",
                )
            )

        # Check 2: fullCalcOnLoad
        if "xl/workbook.xml" in names:
            wb_xml = zf.read("xl/workbook.xml")
            if b"fullCalcOnLoad" in wb_xml:
                has_calc_chain = "xl/calcChain.xml" in names
                if not has_calc_chain:
                    issues.append(
                        Issue(
                            code="FULL_CALC_NO_CHAIN",
                            severity="error",
                            message='fullCalcOnLoad="1" set without calcChain.xml',
                            detail=(
                                "openpyxl sets fullCalcOnLoad in <calcPr> but never generates "
                                "calcChain.xml. This inconsistent state triggers the recovery dialog."
                            ),
                        )
                    )

        # Check 3: Stale calcChain.xml
        if "xl/calcChain.xml" in names:
            try:
                cc_data = zf.read("xl/calcChain.xml")
                cc_tree = ET.fromstring(cc_data)
                cc_refs = set()
                for c_el in cc_tree.iter(f"{{{SHEET_NS}}}c"):
                    ref = c_el.get("r")
                    sheet_id = c_el.get("i")
                    if ref and sheet_id:
                        cc_refs.add((sheet_id, ref))

                if cc_refs:
                    issues.append(
                        Issue(
                            code="CALC_CHAIN_PRESENT",
                            severity="warning",
                            message=f"calcChain.xml references {len(cc_refs)} formula cells",
                            detail=(
                                "If any referenced cell no longer contains a formula, "
                                "Excel will show 'Removed Records: Formula from /xl/calcChain.xml'."
                            ),
                        )
                    )
            except ET.ParseError:
                issues.append(
                    Issue(
                        code="CALC_CHAIN_CORRUPT",
                        severity="error",
                        message="calcChain.xml is not valid XML",
                    )
                )

        # Check 4: Table objects (Table + CF overlap risk)
        table_count = sum(
            1 for n in names if re.match(r"xl/tables/table\d+\.xml", n)
        )
        if table_count > 0:
            issues.append(
                Issue(
                    code="TABLE_OBJECTS",
                    severity="warning",
                    message=f"Found {table_count} Table object(s)",
                    detail=(
                        "Table XML + conditionalFormatting on overlapping ranges = invalid OOXML. "
                        "Use ws.auto_filter.ref instead of ws.add_table() for the same filter UX."
                    ),
                )
            )

        # Check 5: Control characters in sheet XML
        control_char_sheets = []
        for sheet_name in sheet_names:
            data = zf.read(sheet_name)
            if _ILLEGAL_XML_RE.search(data.decode("utf-8", errors="replace")):
                control_char_sheets.append(sheet_name)

        if control_char_sheets:
            issues.append(
                Issue(
                    code="CONTROL_CHARS",
                    severity="warning",
                    message=f"Control characters found in {len(control_char_sheets)} sheet(s)",
                    detail=(
                        "XML 1.0 restricts certain control characters (U+0000-U+0008, U+000B, "
                        "U+000C, U+000E-U+001F). These violate the ST_Xstring type in OOXML."
                    ),
                )
            )

        # Check XML parseability
        for name in names:
            if name.endswith(".xml"):
                try:
                    ET.fromstring(zf.read(name))
                except ET.ParseError as e:
                    issues.append(
                        Issue(
                            code="XML_PARSE_ERROR",
                            severity="error",
                            message=f"Invalid XML in {name}",
                            detail=str(e),
                        )
                    )

    return issues


def fix(
    path: Union[str, Path],
    output: Union[str, Path, None] = None,
    *,
    remove_calc_chain: bool = True,
    strip_control_chars: bool = True,
) -> dict:
    """Fix known corruption patterns in an openpyxl-generated xlsx file.

    Args:
        path: Path to the input xlsx file.
        output: Path for the fixed file. If None, fixes in-place.
        remove_calc_chain: Remove calcChain.xml if present (default True).
        strip_control_chars: Remove illegal XML control characters (default True).

    Returns:
        Dict with fix statistics:
            - inline_strings: number of inline string references converted
            - unique_strings: number of unique strings in the new SST
            - full_calc_removed: whether fullCalcOnLoad was removed
            - calc_chain_removed: whether calcChain.xml was removed
            - control_chars_stripped: number of control characters removed
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    if not zipfile.is_zipfile(path):
        raise ValueError(f"Not a valid ZIP/xlsx file: {path}")

    stats = {
        "inline_strings": 0,
        "unique_strings": 0,
        "full_calc_removed": False,
        "calc_chain_removed": False,
        "control_chars_stripped": 0,
    }

    # Read entire archive into memory
    with zipfile.ZipFile(path, "r") as zin:
        names = zin.namelist()
        members = {name: zin.read(name) for name in names}

    # === Fix 1: Convert inline strings to shared string table ===
    sst: list[str] = []
    sst_index: dict[str, int] = {}
    total_count = 0

    def intern_string(s: str) -> int:
        nonlocal total_count
        total_count += 1
        if s not in sst_index:
            sst_index[s] = len(sst)
            sst.append(s)
        return sst_index[s]

    sheet_names = sorted(
        n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml", n)
    )

    new_sheets: dict[str, bytes] = {}
    for sheet_name in sheet_names:
        # Re-register sheet namespace before parsing each sheet
        ET.register_namespace("", SHEET_NS)
        tree = ET.fromstring(members[sheet_name])

        for cell in tree.iter(f"{{{SHEET_NS}}}c"):
            if cell.get("t") == "inlineStr":
                is_el = cell.find(f"{{{SHEET_NS}}}is")
                if is_el is not None:
                    t_el = is_el.find(f"{{{SHEET_NS}}}t")
                    text = (t_el.text or "") if t_el is not None else ""
                else:
                    text = ""

                # Strip control characters from the string value
                if strip_control_chars:
                    cleaned = _ILLEGAL_XML_RE.sub("", text)
                    if cleaned != text:
                        stats["control_chars_stripped"] += len(text) - len(cleaned)
                        text = cleaned

                idx = intern_string(text)

                # Replace: inlineStr -> shared string ref
                cell.set("t", "s")
                for child in list(cell):
                    if child.tag == f"{{{SHEET_NS}}}is":
                        cell.remove(child)
                v_el = cell.find(f"{{{SHEET_NS}}}v")
                if v_el is None:
                    v_el = ET.SubElement(cell, f"{{{SHEET_NS}}}v")
                v_el.text = str(idx)

        new_sheets[sheet_name] = ET.tostring(tree, xml_declaration=True, encoding="UTF-8")

    stats["inline_strings"] = total_count
    stats["unique_strings"] = len(sst)

    # Build xl/sharedStrings.xml (only if there were inline strings)
    sst_xml: bytes | None = None
    if total_count > 0:
        ET.register_namespace("", SHEET_NS)
        sst_root = ET.Element(f"{{{SHEET_NS}}}sst")
        sst_root.set("count", str(total_count))
        sst_root.set("uniqueCount", str(len(sst)))
        for s in sst:
            si = ET.SubElement(sst_root, f"{{{SHEET_NS}}}si")
            t = ET.SubElement(si, f"{{{SHEET_NS}}}t")
            t.text = s
            # Preserve whitespace for strings with leading/trailing spaces
            if s and s != s.strip():
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        sst_xml = ET.tostring(sst_root, xml_declaration=True, encoding="UTF-8")

        # Patch xl/_rels/workbook.xml.rels - add sharedStrings relationship
        rels_path = "xl/_rels/workbook.xml.rels"
        if rels_path in members:
            ET.register_namespace("", REL_NS)
            rels_tree = ET.fromstring(members[rels_path])
            # Check if SST relationship already exists
            has_sst_rel = any(
                r.get("Type") == SST_REL_TYPE for r in rels_tree
            )
            if not has_sst_rel:
                existing_ids = {r.get("Id") for r in rels_tree}
                rid = "rId99"
                while rid in existing_ids:
                    rid = f"rId{int(rid[3:]) + 1}"
                rel_el = ET.SubElement(rels_tree, f"{{{REL_NS}}}Relationship")
                rel_el.set("Id", rid)
                rel_el.set("Type", SST_REL_TYPE)
                rel_el.set("Target", "sharedStrings.xml")
            members[rels_path] = ET.tostring(
                rels_tree, xml_declaration=True, encoding="UTF-8"
            )

        # Patch [Content_Types].xml - add sharedStrings override
        ET.register_namespace("", CT_NS)
        ct_tree = ET.fromstring(members["[Content_Types].xml"])
        already = any(
            el.get("PartName") == "/xl/sharedStrings.xml"
            for el in ct_tree
            if el.get("PartName")
        )
        if not already:
            override = ET.SubElement(ct_tree, f"{{{CT_NS}}}Override")
            override.set("PartName", "/xl/sharedStrings.xml")
            override.set("ContentType", SST_CT)
        members["[Content_Types].xml"] = ET.tostring(
            ct_tree, xml_declaration=True, encoding="UTF-8"
        )

    # === Fix 2: Remove fullCalcOnLoad ===
    if "xl/workbook.xml" in members:
        ET.register_namespace("", SHEET_NS)
        wb_tree = ET.fromstring(members["xl/workbook.xml"])
        for calc_pr in wb_tree.iter(f"{{{SHEET_NS}}}calcPr"):
            if "fullCalcOnLoad" in calc_pr.attrib:
                del calc_pr.attrib["fullCalcOnLoad"]
                stats["full_calc_removed"] = True
        members["xl/workbook.xml"] = ET.tostring(
            wb_tree, xml_declaration=True, encoding="UTF-8"
        )

    # === Fix 3: Remove stale calcChain.xml ===
    removed_names: set[str] = set()
    if remove_calc_chain and "xl/calcChain.xml" in members:
        del members["xl/calcChain.xml"]
        removed_names.add("xl/calcChain.xml")
        stats["calc_chain_removed"] = True

        # Also remove from [Content_Types].xml
        ET.register_namespace("", CT_NS)
        ct_tree = ET.fromstring(members["[Content_Types].xml"])
        for child in list(ct_tree):
            if child.get("PartName") == "/xl/calcChain.xml":
                ct_tree.remove(child)
        members["[Content_Types].xml"] = ET.tostring(
            ct_tree, xml_declaration=True, encoding="UTF-8"
        )

    # === Write the fixed archive ===
    output_path = Path(output) if output else path
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in names:
            if name in removed_names:
                continue
            if name in new_sheets:
                zout.writestr(name, new_sheets[name])
            elif name in members:
                zout.writestr(name, members[name])
        # Add the new sharedStrings.xml
        if sst_xml is not None:
            zout.writestr("xl/sharedStrings.xml", sst_xml)

    output_path.write_bytes(buf.getvalue())

    return stats


def fix_batch(
    directory: Union[str, Path],
    output_dir: Union[str, Path, None] = None,
    *,
    remove_calc_chain: bool = True,
    strip_control_chars: bool = True,
    recursive: bool = False,
) -> list[dict]:
    """Fix all xlsx files in a directory.

    Args:
        directory: Path to scan for xlsx files.
        output_dir: Directory for fixed files. If None, fixes in-place.
        remove_calc_chain: Remove calcChain.xml if present.
        strip_control_chars: Remove illegal XML control characters.
        recursive: If True, scan subdirectories.

    Returns:
        List of dicts, each with 'file' (Path) and 'stats' (fix result) or 'error' (str).
    """
    directory = Path(directory)
    if not directory.is_dir():
        raise NotADirectoryError(f"Not a directory: {directory}")

    pattern = "**/*.xlsx" if recursive else "*.xlsx"
    files = sorted(directory.glob(pattern))

    if output_dir is not None:
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

    results = []
    for f in files:
        # Skip temp files (Excel lock files start with ~$)
        if f.name.startswith("~$"):
            continue
        out = output_dir / f.name if output_dir else None
        try:
            stats = fix(f, output=out, remove_calc_chain=remove_calc_chain,
                        strip_control_chars=strip_control_chars)
            results.append({"file": f, "stats": stats})
        except Exception as e:
            results.append({"file": f, "error": str(e)})

    return results
