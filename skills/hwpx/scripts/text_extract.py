#!/usr/bin/env python3
"""Extract text from an HWPX document.

Reads HWPX section XML directly for convenient CLI use.

Usage:
    python text_extract.py document.hwpx
    python text_extract.py document.hwpx --format markdown
    python text_extract.py document.hwpx --include-tables
"""

import argparse
import re
import sys
from pathlib import Path
from zipfile import ZipFile

from lxml import etree

NS = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
}


def section_sort_key(name: str) -> tuple[int, str]:
    match = re.search(r"section(\d+)\.xml$", name)
    if match:
        return int(match.group(1)), name
    return 10_000, name


def iter_sections(hwpx_path: str):
    with ZipFile(hwpx_path, "r") as zf:
        section_names = sorted(
            (
                name
                for name in zf.namelist()
                if name.startswith("Contents/section") and name.endswith(".xml")
            ),
            key=section_sort_key,
        )
        for name in section_names:
            yield name, etree.fromstring(zf.read(name))


def is_nested_paragraph(para) -> bool:
    return bool(para.xpath("ancestor::hp:tbl", namespaces=NS))


def paragraph_text(para) -> str:
    parts: list[str] = []
    for node in para.xpath(".//hp:t | .//hp:lineBreak", namespaces=NS):
        if etree.QName(node).localname == "lineBreak":
            parts.append("\n")
        elif node.text:
            parts.append(node.text)
    return "".join(parts)


def extract_plain(hwpx_path: str, *, include_tables: bool = False) -> str:
    """Extract plain text from HWPX file."""

    lines: list[str] = []
    for _, section in iter_sections(hwpx_path):
        for para in section.xpath(".//hp:p", namespaces=NS):
            if not include_tables and is_nested_paragraph(para):
                continue
            text = paragraph_text(para).strip()
            if text:
                lines.append(text)
    return "\n".join(lines)


def extract_markdown(hwpx_path: str) -> str:
    """Extract text as Markdown with section separators."""

    lines: list[str] = []

    for _, section in iter_sections(hwpx_path):
        if lines:
            lines.append("")
            lines.append("---")
            lines.append("")

        for para in section.xpath(".//hp:p", namespaces=NS):
            text = paragraph_text(para).strip()
            if text:
                if is_nested_paragraph(para):
                    lines.append(f"  {text}")
                else:
                    lines.append(text)

    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract text from an HWPX document"
    )
    parser.add_argument("input", help="Path to .hwpx file")
    parser.add_argument(
        "--format", "-f",
        choices=["plain", "markdown"],
        default="plain",
        help="Output format (default: plain)",
    )
    parser.add_argument(
        "--include-tables",
        action="store_true",
        help="Include text from tables and nested objects (plain mode)",
    )
    parser.add_argument(
        "--output", "-o",
        help="Output file path (default: stdout)",
    )
    args = parser.parse_args()

    if not Path(args.input).is_file():
        print(f"Error: File not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    if args.format == "markdown":
        result = extract_markdown(args.input)
    else:
        result = extract_plain(args.input, include_tables=args.include_tables)

    if args.output:
        Path(args.output).write_text(result, encoding="utf-8")
        print(f"Extracted to: {args.output}", file=sys.stderr)
    else:
        print(result)


if __name__ == "__main__":
    main()
