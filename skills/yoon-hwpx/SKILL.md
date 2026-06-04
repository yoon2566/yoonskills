---
name: yoon-hwpx
description: Use this skill for Korean HWPX, Hangul document, 원본 양식 보존, linesegarray, security warning, safe text-node editing, validation, and troubleshooting without local Hancom Office.
---

# Yoon HWPX Skill

Use this skill when the user asks to edit, generate, validate, repair, or troubleshoot a Korean HWPX document while preserving the original Hangul document layout.

## Core Rules

- Do not use local Hangul software.
- Do not use `Hwp.exe`, `HwpConverter.exe`, Hancom Office, COM automation, or GUI conversion.
- Do not modify the original HWPX in place.
- Treat HWPX as a ZIP/XML package.
- Preserve the original form: tables, merged cells, styles, margins, images, and section structure.
- Prefer replacing only `hp:t` text nodes inside `Contents/section*.xml`.
- Remove `hp:linesegarray` from every section XML after text changes.
- Do not claim success without validation logs.

## Workflow

1. Inspect the source HWPX.
2. Validate required package entries and XML well-formedness.
3. Extract a text-node map from `Contents/section*.xml`.
4. Prepare replacement text in JSON or Markdown-derived JSON.
5. Create a new output HWPX from the source.
6. Replace only target `hp:t` text nodes.
7. Remove `hp:linesegarray`.
8. Preserve ZIP entry order and original `ZipInfo` metadata as far as Python stdlib allows.
9. Validate the final HWPX.
10. Save logs and clearly report what changed.

## Recommended Commands

Resolve bundled scripts relative to this skill directory.

```powershell
$env:PYTHONUTF8='1'
$env:PYTHONIOENCODING='utf-8'
$py = ".\.venv\Scripts\python.exe"
$skill = "C:\path\to\yoon-hwpx"

& $py "$skill\scripts\analyze_hwpx.py" .\input\template.hwpx --json .\work\analysis.json
& $py "$skill\scripts\extract_text_map.py" .\input\template.hwpx --output .\work\text-map.json
& $py "$skill\scripts\apply_text_map.py" --source .\input\template.hwpx --map .\work\text-map.json --output .\output\result.hwpx
& $py "$skill\scripts\validate_hwpx.py" .\output\result.hwpx --expect-no-linesegarray
```

## Red Flags

- The file validates as ZIP/XML but Hangul shows a damage, tamper, or security warning.
- The original form was rebuilt from scratch instead of copied.
- Rows, columns, images, styles, or merged cells changed without explicit approval.
- The final output has no verification summary.

If the user reports a Hangul warning after text edits, check whether old `hp:linesegarray` nodes remain.

## References

- Read [hwpx-safe-edit-workflow.md](references/hwpx-safe-edit-workflow.md) for the detailed safe-edit procedure.
- Read [troubleshooting.md](references/troubleshooting.md) when validation succeeds but Hangul reports damage, tampering, or layout problems.
