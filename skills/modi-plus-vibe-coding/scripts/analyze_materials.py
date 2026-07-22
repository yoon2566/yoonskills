"""Analyze MODI+ teaching materials exported as PDF or PowerPoint.

The script is intentionally self-contained so a student or maintainer can
repeat a material survey without depending on the original workstation. PDF
text extraction uses pypdf when available; PPTX extraction uses only the
standard library. The report keeps page/slide previews and writes complete
plain-text extracts beside the JSON/Markdown summary.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple
from xml.etree import ElementTree


KEYWORD_GROUPS = {
    "python": ("python", "파이썬"),
    "vibe_coding": ("바이브코딩", "vibe coding", "chatgpt", "챗gpt"),
    "connection": ("연결", "네트워크", "network", "usb", "serial", "직렬", "블루투스", "ble"),
    "button": ("버튼", "button", "pressed", "clicked", "toggled"),
    "led_rgb": ("led", "엘이디", "rgb", "빨강", "초록", "파랑"),
    "motor_car": ("모터", "motor", "자동차", "차", "주행", "바퀴", "wheel"),
    "joystick": ("조이스틱", "joystick", "x축", "y축", "deadzone"),
    "display": ("디스플레이", "display", "화면", "글자", "text"),
    "sensor": ("센서", "sensor", "환경", "env", "온도", "습도", "조도", "tof", "거리", "imu"),
    "speaker": ("스피커", "speaker", "소리", "음악", "music"),
    "programming": ("변수", "조건문", "반복문", "함수", "리스트", "딕셔너리", "클래스", "알고리즘"),
    "debugging": ("오류", "에러", "디버깅", "debug", "예외", "문제 해결", "troubleshoot"),
    "safety": ("안전", "정지", "stop", "주의", "보호", "위험", "공중", "바닥"),
    "project": ("프로젝트", "작품", "실습", "과제", "게임", "미션"),
}


def normalize_text(value: str) -> str:
    """Collapse extraction whitespace while preserving readable text."""

    value = value.replace("\x00", " ").replace("\u00a0", " ")
    value = re.sub(r"[ \t\r\f\v]+", " ", value)
    value = re.sub(r"\n{3,}", "\n\n", value)
    return value.strip()


def preview(value: str, limit: int = 320) -> str:
    """Return a one-line preview suitable for Markdown tables."""

    value = normalize_text(value).replace("\n", " / ")
    if len(value) <= limit:
        return value
    return value[: limit - 1].rstrip() + "…"


def safe_filename(value: str) -> str:
    value = re.sub(r"[^0-9A-Za-z가-힣._-]+", "_", value)
    return value.strip("._") or "material"


def extract_pdf(path: Path) -> List[str]:
    try:
        from pypdf import PdfReader
    except ImportError as exc:  # pragma: no cover - environment-specific
        raise RuntimeError("PDF 분석에는 pypdf가 필요합니다.") from exc

    reader = PdfReader(str(path))
    pages = []
    for page in reader.pages:
        pages.append(normalize_text(page.extract_text() or ""))
    return pages


def extract_pptx(path: Path) -> List[str]:
    """Extract text from PPTX slide XML without requiring python-pptx."""

    slide_pattern = re.compile(r"^ppt/slides/slide(\d+)\.xml$")
    slide_files: List[Tuple[int, str]] = []
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            match = slide_pattern.match(name)
            if match:
                slide_files.append((int(match.group(1)), name))
        pages = []
        for _, name in sorted(slide_files):
            root = ElementTree.fromstring(archive.read(name))
            text_parts = [element.text or "" for element in root.iter() if element.tag.endswith("}t")]
            pages.append(normalize_text("\n".join(text_parts)))
    return pages


def extract_pages(path: Path) -> Tuple[str, List[str]]:
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return "pdf", extract_pdf(path)
    if suffix == ".pptx":
        return "pptx", extract_pptx(path)
    raise ValueError(f"지원하지 않는 파일 형식: {path}")


def count_keyword_groups(text: str) -> Dict[str, int]:
    lowered = text.lower()
    counts = {}
    for group, terms in KEYWORD_GROUPS.items():
        counts[group] = sum(lowered.count(term.lower()) for term in terms)
    return {key: value for key, value in counts.items() if value}


def first_title(pages: Sequence[str], fallback: str) -> str:
    for page in pages:
        for line in page.splitlines():
            candidate = line.strip()
            if candidate:
                return candidate[:180]
    return fallback


def session_number(name: str) -> Optional[int]:
    match = re.search(r"[_ -](\d+)차시", name)
    return int(match.group(1)) if match else None


def analyze_file(path: Path, root: Path, text_dir: Path) -> Dict[str, object]:
    kind, pages = extract_pages(path)
    relative = path.relative_to(root).as_posix()
    text_path = text_dir / (safe_filename(path.stem) + ".txt")
    text_path.write_text(
        "\n\n".join(f"--- {kind.upper()} {index} ---\n{page}" for index, page in enumerate(pages, 1)),
        encoding="utf-8",
    )
    all_text = "\n".join(pages)
    page_records = [
        {
            "number": index,
            "characters": len(page),
            "title_or_first_line": first_title([page], ""),
            "preview": preview(page),
            "keyword_groups": count_keyword_groups(page),
        }
        for index, page in enumerate(pages, 1)
    ]
    aggregate = Counter()
    for page in pages:
        aggregate.update(count_keyword_groups(page))
    return {
        "file": relative,
        "format": kind,
        "bytes": path.stat().st_size,
        "session": session_number(path.name),
        "pages_or_slides": len(pages),
        "title": first_title(pages, path.stem),
        "characters": len(all_text),
        "keyword_groups": dict(aggregate),
        "text_extract": text_path.relative_to(text_dir.parent).as_posix(),
        "pages": page_records,
    }


def markdown_report(data: Dict[str, object]) -> str:
    files = data["files"]
    format_counts = Counter(item["format"] for item in files)
    lines = [
        "# MODI+ 수업 자료 분석 보고서",
        "",
        f"- 분석 기준 폴더: `{data['input_dir']}`",
        f"- 분석 시각(로컬 실행): `{data['generated_at']}`",
        f"- 문서 수: **{len(files)}** ({format_counts.get('pdf', 0)} PDF, {format_counts.get('pptx', 0)} PPTX)",
        "- PDF는 PPTX가 없는 경우에도 슬라이드 PDF 내보내기로 간주하여 페이지별 텍스트와 주제를 분석했다.",
        "- `text_extracts/`에는 각 문서의 페이지/슬라이드 전체 추출 텍스트가 있다.",
        "",
        "## 문서별 개요",
        "",
        "| 순서 | 파일 | 페이지/슬라이드 | 추출 문자 수 | 주요 주제 | 제목/첫 문장 |",
        "|---:|---|---:|---:|---|---|",
    ]
    ordered = sorted(files, key=lambda item: (item["session"] is None, item["session"] or 999, item["file"]))
    for index, item in enumerate(ordered, 1):
        topics = ", ".join(item["keyword_groups"].keys()) or "추출 텍스트 없음"
        title = item["title"].replace("|", "¦")
        lines.append(
            f"| {index} | `{item['file']}` | {item['pages_or_slides']} | {item['characters']} | {topics} | {title} |"
        )

    lines.extend(["", "## 페이지/슬라이드별 분석", ""])
    for item in ordered:
        lines.extend(
            [
                f"### `{item['file']}`",
                "",
                f"- 형식: {item['format'].upper()} / {item['pages_or_slides']}페이지(슬라이드)",
                f"- 문서 주제 키워드 빈도: `{json.dumps(item['keyword_groups'], ensure_ascii=False)}`",
                f"- 전체 추출 텍스트: [`{item['text_extract']}`]({item['text_extract']})",
                "",
                "| 번호 | 문자 수 | 키워드 그룹 | 내용 미리보기 |",
                "|---:|---:|---|---|",
            ]
        )
        for page in item["pages"]:
            groups = ", ".join(page["keyword_groups"].keys()) or "-"
            page_preview = page["preview"].replace("|", "¦")
            lines.append(f"| {page['number']} | {page['characters']} | {groups} | {page_preview} |")
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input_dir", type=Path, help="PDF/PPTX가 들어 있는 폴더")
    parser.add_argument("--output-dir", type=Path, required=True, help="JSON/Markdown/텍스트를 쓸 폴더")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    input_dir = args.input_dir.resolve()
    output_dir = args.output_dir.resolve()
    if not input_dir.is_dir():
        print(f"입력 폴더가 없습니다: {input_dir}", file=sys.stderr)
        return 2

    output_dir.mkdir(parents=True, exist_ok=True)
    text_dir = output_dir / "text_extracts"
    text_dir.mkdir(parents=True, exist_ok=True)
    paths = sorted(
        path for path in input_dir.rglob("*")
        if path.is_file() and path.suffix.lower() in {".pdf", ".pptx"}
    )
    if not paths:
        print("PDF 또는 PPTX를 찾지 못했습니다.", file=sys.stderr)
        return 2

    try:
        files = [analyze_file(path, input_dir, text_dir) for path in paths]
    except Exception as exc:
        print(f"자료 분석 실패: {exc}", file=sys.stderr)
        return 1

    from datetime import datetime

    data = {
        "input_dir": str(input_dir),
        "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "file_count": len(files),
        "files": files,
    }
    (output_dir / "materials_analysis.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (output_dir / "materials_analysis.md").write_text(markdown_report(data), encoding="utf-8")
    print(f"분석 완료: {len(files)}개 문서")
    print(f"JSON: {output_dir / 'materials_analysis.json'}")
    print(f"Markdown: {output_dir / 'materials_analysis.md'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
