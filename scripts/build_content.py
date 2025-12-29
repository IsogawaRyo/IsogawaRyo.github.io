#!/usr/bin/env python3
"""Convert content.xlsx into content.json for the static site."""
from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from typing import Dict, List
import xml.etree.ElementTree as ET

NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def load_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    try:
        data = zf.read("xl/sharedStrings.xml")
    except KeyError:
        return []
    root = ET.fromstring(data)
    strings: List[str] = []
    for si in root.findall("main:si", NS):
        text_parts = [node.text or "" for node in si.findall(".//main:t", NS)]
        strings.append("".join(text_parts))
    return strings


def column_to_index(ref: str) -> int:
    letters = re.match(r"([A-Z]+)", ref)
    if not letters:
        return 0
    value = 0
    for ch in letters.group(1):
        value = value * 26 + (ord(ch) - 64)
    return value - 1


def load_rows(zf: zipfile.ZipFile, shared_strings: List[str]) -> List[List[str]]:
    sheet_data = zf.read("xl/worksheets/sheet1.xml")
    root = ET.fromstring(sheet_data)
    rows_xml = root.findall("main:sheetData/main:row", NS)
    rows: List[List[str]] = []
    for row in rows_xml:
        cells: Dict[int, str] = {}
        for cell in row.findall("main:c", NS):
            ref = cell.get("r", "A1")
            idx = column_to_index(ref)
            cell_type = cell.get("t")
            value = ""
            v = cell.find("main:v", NS)
            if cell_type == "s" and v is not None:
                value = shared_strings[int(v.text)]
            elif cell_type == "inlineStr":
                inline = cell.find("main:is/main:t", NS)
                value = inline.text if inline is not None else ""
            elif v is not None:
                value = v.text or ""
            cells[idx] = value
        if not cells:
            continue
        max_idx = max(cells)
        row_values = [cells.get(i, "") for i in range(max_idx + 1)]
        rows.append(row_values)
    return rows


def convert_records(rows: List[List[str]]) -> Dict[str, object]:
    if not rows:
        raise SystemExit("Excel sheet is empty")
    header = rows[0]
    records = []
    for values in rows[1:]:
        record = {header[i]: values[i] if i < len(values) else "" for i in range(len(header))}
        if any(value.strip() for value in record.values()):
            records.append(record)

    data = {
        "hero": {"name": "", "tagline": "", "description": ""},
        "selfIntro": [],
        "currentSelf": [],
        "direction": [],
        "stack": [],
        "basicInfo": [],
        "recentWork": [],
        "studies": [],
        "favorites": [],
        "contact": [],
        "memo": [],
    }

    for record in records:
        section = record.get("Section", "").strip().upper()
        key = record.get("Key", "").strip()
        value = record.get("Value", "").strip()
        link = record.get("Link", "").strip()

        if not section:
            continue

        if section == "HERO":
            lowered = key.lower()
            if lowered == "name":
                data["hero"]["name"] = value
            elif lowered == "tagline":
                data["hero"]["tagline"] = value
            elif lowered == "description":
                data["hero"]["description"] = value
        elif section == "SELF_INTRO" and value:
            data["selfIntro"].append(value)
        elif section == "CURRENT_SELF" and value:
            data["currentSelf"].append(value)
        elif section == "DIRECTION" and value:
            data["direction"].append(value)
        elif section == "STACK" and value:
            data["stack"].append(value)
        elif section == "BASIC_INFO" and value:
            data["basicInfo"].append({"label": key or "", "value": value})
        elif section == "RECENT_WORK" and value:
            data["recentWork"].append(value)
        elif section == "STUDIES" and value:
            data["studies"].append(value)
        elif section == "FAVORITES" and value:
            data["favorites"].append(value)
        elif section == "CONTACT" and value:
            data["contact"].append({
                "label": key or value,
                "value": value,
                "link": link or value,
            })
        elif section == "MEMO" and value:
            data["memo"].append(value)

    return data


def build(input_path: Path, output_path: Path) -> None:
    if not input_path.exists():
        raise SystemExit(f"Excel file not found: {input_path}")

    with zipfile.ZipFile(input_path, "r") as zf:
        shared_strings = load_shared_strings(zf)
        rows = load_rows(zf, shared_strings)

    data = convert_records(rows)
    output_path.write_text(json.dumps(data, ensure_ascii=False, indent=2))
    print(f"Wrote {output_path} from {input_path}")


def main(argv: List[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", default="content.xlsx", help="Path to the Excel source file")
    parser.add_argument("--output", default="content.json", help="Where to write the generated JSON")
    args = parser.parse_args(argv)

    build(Path(args.input), Path(args.output))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
