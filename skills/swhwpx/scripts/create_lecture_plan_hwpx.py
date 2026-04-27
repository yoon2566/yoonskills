#!/usr/bin/env python3
"""Create a 6-column Korean lecture-plan HWPX from JSON.

This helper follows the XML-first workflow used by the hwpx-2 skill:
it creates UTF-8 OWPML section XML, packages it with build_hwpx.py, and
runs validate.py on the final HWPX.
"""

from __future__ import annotations

import argparse
import copy
import json
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

from lxml import etree


SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
REPORT_SECTION = SKILL_DIR / "templates" / "report" / "section0.xml"
BUILD = SCRIPT_DIR / "build_hwpx.py"
VALIDATE = SCRIPT_DIR / "validate.py"

HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
HS = "http://www.hancom.co.kr/hwpml/2011/section"


def hp(tag: str, **attrs: str) -> etree._Element:
    return etree.Element(f"{{{HP}}}{tag}", **attrs)


def qn(tag: str) -> str:
    return f"{{{HP}}}{tag}"


class Ids:
    def __init__(self, start: int = 1000000002) -> None:
        self.value = start

    def next(self) -> str:
        current = self.value
        self.value += 1
        return str(current)


def as_text(value: Any, default: str = "") -> str:
    if value is None:
        return default
    if isinstance(value, list):
        return "\n".join(str(item) for item in value)
    return str(value)


def load_base_root() -> etree._Element:
    tree = etree.parse(str(REPORT_SECTION))
    template_root = tree.getroot()
    root = etree.Element(f"{{{HS}}}sec", nsmap=template_root.nsmap)
    first_paragraph = copy.deepcopy(template_root.find(qn("p")))
    if first_paragraph is None:
        raise RuntimeError("report template has no first paragraph")
    root.append(first_paragraph)
    return root


def make_text_paragraph(ids: Ids, text: str, para: str = "0", char: str = "0") -> etree._Element:
    paragraph = hp(
        "p",
        id=ids.next(),
        paraPrIDRef=para,
        styleIDRef="0",
        pageBreak="0",
        columnBreak="0",
        merged="0",
    )
    run = hp("run", charPrIDRef=char)
    text_node = hp("t")
    if text:
        text_node.text = text
    run.append(text_node)
    paragraph.append(run)
    return paragraph


def row_height(row: list[tuple[str, int, str]]) -> int:
    max_lines = max((len(text.splitlines()) or 1) for text, _, _ in row)
    return max(1900, 850 * max_lines + 900)


def make_cell(
    ids: Ids,
    row: int,
    col: int,
    width: int,
    height: int,
    col_span: int,
    text: str,
    kind: str,
) -> etree._Element:
    border = "4" if kind in {"label", "header", "title"} else "3"
    para = "21" if kind in {"label", "header", "title", "center"} else "22"
    char = {"title": "7", "label": "9", "header": "9", "center": "0"}.get(kind, "0")

    cell = hp(
        "tc",
        name="",
        header="0",
        hasMargin="0",
        protect="0",
        editable="0",
        dirty="1",
        borderFillIDRef=border,
    )
    cell.append(hp("cellAddr", colAddr=str(col), rowAddr=str(row)))
    cell.append(hp("cellSpan", colSpan=str(col_span), rowSpan="1"))
    cell.append(hp("cellSz", width=str(width), height=str(height)))
    cell.append(hp("cellMargin", left="283", right="283", top="141", bottom="141"))

    sub_list = hp(
        "subList",
        id="",
        textDirection="HORIZONTAL",
        lineWrap="BREAK",
        vertAlign="CENTER",
        linkListIDRef="0",
        linkListNextIDRef="0",
        textWidth=str(max(0, width - 566)),
        fieldName="",
    )
    for line in text.splitlines() or [""]:
        sub_list.append(make_text_paragraph(ids, line, para=para, char=char))
    cell.append(sub_list)
    return cell


def build_rows(data: dict[str, Any]) -> list[list[tuple[str, int, str]]]:
    title = as_text(data.get("title"), "2026년 여름방학 AI 교육 강의계획서")
    teacher = as_text(data.get("teacher"), "")
    fee = as_text(data.get("fee"), "1인 기준 0원")

    rows: list[list[tuple[str, int, str]]] = [
        [(title, 6, "title")],
        [("강좌명", 1, "label"), (as_text(data.get("course_name")), 1, "body"), ("강사명", 2, "label"), (teacher, 2, "body")],
        [("강의목표 및 개요", 1, "label"), (as_text(data.get("goal")), 5, "body")],
        [("수강 대상", 1, "label"), (as_text(data.get("target")), 5, "body")],
        [("사용 매체", 1, "label"), (as_text(data.get("media"), "강사용 컴퓨터, 학생 컴퓨터, 인터넷"), 5, "body")],
        [("교재", 1, "label"), (as_text(data.get("materials"), "자체 제작 교안 및 활동지"), 1, "body"), ("재료비", 1, "label"), (fee, 3, "body")],
        [("차시", 1, "header"), ("강의내용", 4, "header"), ("학습준비물", 1, "header")],
    ]

    for session in data.get("sessions", []):
        no = as_text(session.get("no"))
        topic = as_text(session.get("topic"))
        activity = as_text(session.get("activity"))
        result = as_text(session.get("result"))
        prep = as_text(session.get("prep"))
        content_lines = [f"[{topic}]" if topic else "", f"- {activity}" if activity else ""]
        if result:
            content_lines.append(f"- 결과물: {result}")
        content = "\n".join(line for line in content_lines if line)
        rows.append([(no, 1, "center"), (content, 4, "body"), (prep, 1, "body")])

    for extra in data.get("extra_rows", []):
        rows.append([(as_text(extra.get("label")), 1, "label"), (as_text(extra.get("text")), 5, "body")])

    return rows


def make_table_paragraph(ids: Ids, data: dict[str, Any]) -> etree._Element:
    col_widths = [5200, 9800, 5200, 7600, 7600, 7120]
    rows = build_rows(data)
    heights = [row_height(row) for row in rows]

    paragraph = hp(
        "p",
        id=ids.next(),
        paraPrIDRef="0",
        styleIDRef="0",
        pageBreak="0",
        columnBreak="0",
        merged="0",
    )
    run = hp("run", charPrIDRef="0")
    table = hp(
        "tbl",
        id="1000000099",
        zOrder="0",
        numberingType="TABLE",
        textWrap="TOP_AND_BOTTOM",
        textFlow="BOTH_SIDES",
        lock="0",
        dropcapstyle="None",
        pageBreak="CELL",
        repeatHeader="0",
        rowCnt=str(len(rows)),
        colCnt="6",
        cellSpacing="0",
        borderFillIDRef="3",
        noAdjust="0",
    )
    table.append(hp("sz", width="42520", widthRelTo="ABSOLUTE", height=str(sum(heights)), heightRelTo="ABSOLUTE", protect="0"))
    table.append(hp("pos", treatAsChar="1", affectLSpacing="0", flowWithText="1", allowOverlap="0", holdAnchorAndSO="0", vertRelTo="PARA", horzRelTo="COLUMN", vertAlign="TOP", horzAlign="LEFT", vertOffset="0", horzOffset="0"))
    table.append(hp("outMargin", left="0", right="0", top="0", bottom="0"))
    table.append(hp("inMargin", left="0", right="0", top="0", bottom="0"))

    for row_idx, (row, height) in enumerate(zip(rows, heights)):
        tr = hp("tr")
        col_idx = 0
        for text, span, kind in row:
            width = sum(col_widths[col_idx : col_idx + span])
            tr.append(make_cell(ids, row_idx, col_idx, width, height, span, text, kind))
            col_idx += span
        table.append(tr)

    run.append(table)
    paragraph.append(run)
    return paragraph


def make_section(data: dict[str, Any], output: Path) -> Path:
    ids = Ids()
    root = load_base_root()
    root.append(make_table_paragraph(ids, data))
    etree.indent(root, space="  ")
    etree.ElementTree(root).write(str(output), xml_declaration=True, encoding="UTF-8", pretty_print=True)
    return output


def main() -> int:
    parser = argparse.ArgumentParser(description="Create a 6-column lecture-plan HWPX from JSON")
    parser.add_argument("--input", "-i", required=True, help="Input JSON path")
    parser.add_argument("--output", "-o", required=True, help="Output .hwpx path")
    parser.add_argument("--creator", default="Codex", help="Metadata creator")
    args = parser.parse_args()

    input_path = Path(args.input).resolve()
    output_path = Path(args.output).resolve()
    data = json.loads(input_path.read_text(encoding="utf-8"))
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as tmp:
        section_path = make_section(data, Path(tmp) / "section0.xml")
        subprocess.run(
            [
                sys.executable,
                str(BUILD),
                "--template",
                "report",
                "--section",
                str(section_path),
                "--title",
                as_text(data.get("title"), output_path.stem),
                "--creator",
                args.creator,
                "--output",
                str(output_path),
            ],
            check=True,
        )

    subprocess.run([sys.executable, str(VALIDATE), str(output_path)], check=True)
    print(f"Created lecture-plan HWPX: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
