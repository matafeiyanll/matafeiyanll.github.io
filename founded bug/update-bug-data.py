"""Rebuild data/bug-data.js from the Excel workbooks in data/.

This script uses only the Python standard library. It reads the first worksheet
from every workbook registered in data/workbooks.json and keeps the Excel files
as the source of truth for the website fallback data.
"""

from __future__ import annotations

import json
import posixpath
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
MANIFEST_PATH = DATA_DIR / "workbooks.json"
OUTPUT_PATH = DATA_DIR / "bug-data.js"
EXPECTED_HEADERS = ["Program", "ID", "Bug ID", "Bug Link", "Bug Type", "Status", "Method"]

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def column_index(cell_reference: str) -> int:
    match = re.match(r"[A-Z]+", cell_reference.upper())
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_reference}")
    index = 0
    for character in match.group(0):
        index = index * 26 + ord(character) - ord("A") + 1
    return index - 1


def parse_number(value: str):
    if re.fullmatch(r"-?\d+", value):
        return int(value)
    try:
        return float(value)
    except ValueError:
        return value


def cell_text(cell: ET.Element, shared_strings: list[str]):
    cell_type = cell.get("t", "")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.findall(f".//{{{MAIN_NS}}}t"))

    value_node = cell.find(f"{{{MAIN_NS}}}v")
    if value_node is None or value_node.text is None:
        return ""

    value = value_node.text
    if cell_type == "s":
        return shared_strings[int(value)]
    if cell_type == "b":
        return value == "1"
    if cell_type in {"str", "e"}:
        return value
    return parse_number(value)


def first_worksheet_path(archive: zipfile.ZipFile) -> str:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    first_sheet = workbook.find(f".//{{{MAIN_NS}}}sheet")
    if first_sheet is None:
        raise ValueError("Workbook has no worksheets")

    relationship_id = first_sheet.get(f"{{{REL_NS}}}id")
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    for relationship in relationships.findall(f"{{{PACKAGE_REL_NS}}}Relationship"):
        if relationship.get("Id") == relationship_id:
            target = relationship.get("Target", "")
            if target.startswith("/"):
                return target.lstrip("/")
            return posixpath.normpath(posixpath.join("xl", target))
    raise ValueError("Could not locate the first worksheet")


def read_workbook(path: Path) -> list[dict]:
    with zipfile.ZipFile(path) as archive:
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            shared_root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            shared_strings = [
                "".join(node.text or "" for node in item.findall(f".//{{{MAIN_NS}}}t"))
                for item in shared_root.findall(f"{{{MAIN_NS}}}si")
            ]

        worksheet = ET.fromstring(archive.read(first_worksheet_path(archive)))
        parsed_rows: list[list] = []
        for row in worksheet.findall(f".//{{{MAIN_NS}}}sheetData/{{{MAIN_NS}}}row"):
            cells: dict[int, object] = {}
            for cell in row.findall(f"{{{MAIN_NS}}}c"):
                reference = cell.get("r", "")
                cells[column_index(reference)] = cell_text(cell, shared_strings)
            if cells:
                parsed_rows.append([cells.get(index, "") for index in range(max(cells) + 1)])

    if not parsed_rows:
        raise ValueError(f"{path.name}: worksheet is empty")

    headers = [str(value).strip() for value in parsed_rows[0]]
    missing = [header for header in EXPECTED_HEADERS if header not in headers]
    if missing:
        raise ValueError(f"{path.name}: missing columns {', '.join(missing)}")

    records: list[dict] = []
    for values in parsed_rows[1:]:
        record = {
            header: values[index] if index < len(values) else ""
            for index, header in enumerate(headers)
            if header
        }
        normalized = {header: record.get(header, "") for header in EXPECTED_HEADERS}
        if normalized["Program"] and normalized["Bug ID"]:
            records.append(normalized)
    return records


def write_text_atomic(path: Path, text: str) -> None:
    temporary = path.with_name(f".{path.name}.tmp")
    temporary.write_text(text, encoding="utf-8", newline="\n")
    temporary.replace(path)


def main() -> int:
    manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8-sig"))
    all_records: list[dict] = []

    for entry in manifest:
        workbook_path = DATA_DIR / entry["file"]
        if not workbook_path.is_file():
            raise FileNotFoundError(f"Missing workbook: {workbook_path}")
        records = read_workbook(workbook_path)
        entry["count"] = len(records)
        all_records.extend(records)
        print(f"{entry['method']}: {len(records)} records")

    manifest_text = json.dumps(manifest, ensure_ascii=False, indent=2) + "\n"
    fallback_text = "window.BUG_DATA = " + json.dumps(all_records, ensure_ascii=False, indent=2) + ";\n"
    write_text_atomic(MANIFEST_PATH, manifest_text)
    write_text_atomic(OUTPUT_PATH, fallback_text)
    print(f"Updated data/bug-data.js with {len(all_records)} records.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(f"Update failed: {error}", file=sys.stderr)
        raise SystemExit(1)
