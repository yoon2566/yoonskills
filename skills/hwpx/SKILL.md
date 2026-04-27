---
name: hwpx
description: Create, inspect, edit, and validate Korean Hangul HWPX documents without Hancom Office. Use for .hwp/.hwpx reference forms, direct OWPML XML assembly, lecture-plan generation, table/form preservation, text extraction, HWPX package validation, and older prompts that mention hwpx-2.
---

# HWPX

Use this skill when the user wants Korean Hangul documents created, inspected, edited, or checked without launching Hancom Office. The default output is `.hwpx`, not binary `.hwp`.

Core rule: do not use Hancom Office, COM automation, or GUI conversion. Build HWPX as a ZIP/XML package and validate the result.

## Environment

- On Windows, use PowerShell syntax.
- Use the workspace `.venv` explicitly, for example `.\.venv\Scripts\python.exe`.
- Do not use global Python or `pip install --user`.
- Install only the missing packages needed by the scripts, usually `lxml`; use `kordoc` only when `.hwp` parsing is required and available.

## Workflow

1. **Reference analysis**
   - If the reference is `.hwp`, parse it with `kordoc` to Markdown/JSON and use the extracted text/table structure as layout guidance.
   - If the reference is `.hwpx`, use `scripts/analyze_template.py` to extract `Contents/header.xml` and `Contents/section0.xml`.
   - For form-reference work, preserve table structure, cell spans, margins, paragraph IDs, character style references, and section properties where possible.

2. **Generate content**
   - Prefer direct UTF-8 XML generation for reliable Korean text.
   - For the common 6-column lecture plan form, use `scripts/create_lecture_plan_hwpx.py`.
   - For general documents, either use a bundled template or create/edit `section0.xml` directly.
   - Limit structural changes unless the user explicitly asks for them; for provided forms, replace only the requested text/data where possible.

3. **Package**
   - Use `scripts/build_hwpx.py`.
   - Bundled templates:
     - `report` for general reports and lecture plans.
     - `gonmun` for official-document style.
     - `minutes` for meeting minutes.
     - `proposal` for proposals or business-overview documents.
   - The build script keeps `mimetype` first and uncompressed.

4. **Validate**
   - Always run `scripts/validate.py` on generated HWPX files.
   - When preserving a reference document and page-count fidelity matters, also run `scripts/page_guard.py`.
   - If parser output displays mojibake, inspect `Contents/section0.xml` inside the HWPX directly before concluding the file is broken. Parser output can be display-encoding-sensitive even when the internal XML is correct.

## Common Commands

Use paths relative to this skill directory when running bundled scripts.

```powershell
$py = "C:\path\to\workspace\.venv\Scripts\python.exe"
$skill = "C:\Users\gyu\.codex\skills\hwpx"

# Analyze an HWPX reference form.
& $py "$skill\scripts\analyze_template.py" "reference.hwpx" `
  --extract-header "ref_header.xml" `
  --extract-section "ref_section0.xml"

# Build from a bundled template.
& $py "$skill\scripts\build_hwpx.py" `
  --template report `
  --title "문서 제목" `
  --creator "Codex" `
  --output "result.hwpx"

# Build from custom XML.
& $py "$skill\scripts\build_hwpx.py" `
  --template report `
  --section "section0.xml" `
  --output "result.hwpx"

# Validate.
& $py "$skill\scripts\validate.py" "result.hwpx"

# Compare page count against a reference when needed.
& $py "$skill\scripts\page_guard.py" `
  --reference "reference.hwpx" `
  --output "result.hwpx"
```

## Lecture Plan Helper

Use `scripts/create_lecture_plan_hwpx.py` for a table like the `2026년 여름방학 AI 교육 강의계획서` workflow.

Input JSON shape:

```json
{
  "title": "2026년 여름방학 AI 교육 강의계획서",
  "course_name": "AI 융합 교육: AI 미디어 만들기",
  "teacher": "윤설정",
  "target": "초등 1-2학년",
  "goal": "1. 목표...\n2. 목표...",
  "media": "강사용 컴퓨터, 학생 컴퓨터, 인터넷, AI 도구",
  "materials": "자체 제작 교안 및 활동지",
  "fee": "1인 기준 0원",
  "sessions": [
    {
      "no": "1",
      "topic": "AI는 무엇을 도와줄까?",
      "activity": "AI가 그림, 글, 목소리를 만드는 예시 보기",
      "result": "AI 이해 활동지",
      "prep": "AI 사례 자료"
    }
  ],
  "extra_rows": [
    {"label": "추천 수업명", "text": "AI랑 함께 만드는 상상 그림책"},
    {"label": "최종 결과물", "text": "AI 그림책 또는 AI 캐릭터 미디어 카드"}
  ]
}
```

Run:

```powershell
& $py "$skill\scripts\create_lecture_plan_hwpx.py" --input "data.json" --output "result.hwpx"
& $py "$skill\scripts\validate.py" "result.hwpx"
```

## Custom XML

When generating a custom document, write `section0.xml` using the namespaces and structural patterns in `references/hwpx-format.md`, then package it with `scripts/build_hwpx.py`.

Read `references/hwpx-format.md` only when exact OWPML element patterns are needed for paragraphs, tables, section properties, or package structure.

For HWPX documents, the key files are:

- `Contents/header.xml`: styles, fonts, paragraph and character properties, border fills, metadata.
- `Contents/section0.xml`: body content, paragraphs, tables, sections.
- `Contents/content.hpf`: package manifest.
- `mimetype`: must be the first ZIP entry and contain `application/hwp+zip`.

## Validation Checklist

- Generated file is a non-empty `.hwpx` ZIP archive.
- First ZIP entry is `mimetype`.
- Required files exist: `Contents/content.hpf`, `Contents/header.xml`, `Contents/section0.xml`.
- XML and HPF files are well-formed.
- Korean text is present in internal `Contents/section0.xml`.
- For form-reference work, the result follows the reference table/section shape closely enough for the user's stated purpose.
- If page count is important, `scripts/page_guard.py` passes or the mismatch is explained to the user.
