---
name: edukit
description: Coordinate end-to-end education production workflows that span multiple tools or artifact types, such as Google Workspace, lesson plans, presentations, HWPX/PDF documents, browser research, 3D assets, and video. Use when the user explicitly asks for EduKit or when an education task requires orchestration across two or more domains. Prefer dedicated specialist skills for a single narrow task.
---

# EduKit

Coordinate multi-tool education work while preserving each specialist tool's workflow and validation rules.

## Core Rules

1. Read workspace instructions before acting.
2. Assume Windows PowerShell unless the environment says otherwise.
3. Route narrow tasks to a dedicated skill instead of duplicating its instructions.
4. Use the workspace `.venv\Scripts\python.exe`; create `.venv` first when Python is required and it does not exist.
5. Inspect schemas or command help before using unfamiliar or version-sensitive CLI commands.
6. Preview and verify file moves, renames, uploads, writes, and conversions.
7. Keep user-specific paths, account names, IDs, and credentials out of reusable outputs.

## Routing

| Task | Preferred route | Read |
|---|---|---|
| Google Drive, Sheets, Docs, Slides, Gmail, Calendar | Use a dedicated Google Workspace skill or `gws`; use the bundled runner for JSON-heavy PowerShell calls | [references/google-workspace.md](references/google-workspace.md) |
| Lesson plans, presentations, HWPX, PDF, HTML learning materials | Use the relevant document or presentation skill | [references/documents-content.md](references/documents-content.md) |
| STL, STEP/STP, OpenSCAD, Blender, browser checks, video | Use the relevant 3D, browser, or media skill | [references/3d-media-browser.md](references/3d-media-browser.md) |
| CLI setup, Python environments, file operations, Codex configuration | Follow Windows-safe setup and mutation rules | [references/environment-safety.md](references/environment-safety.md) |

## Workflow

1. Clarify the requested final artifact and delivery location from the conversation and workspace.
2. Break the request into domain-specific stages.
3. Load only the reference files and specialist skills needed for those stages.
4. Check installed tools and authentication before relying on them.
5. Produce intermediate artifacts in the workspace or another user-approved location.
6. Validate each stage before using its output in the next stage.
7. Report final artifact paths, remote links, checks performed, and unresolved limitations.

## Google Workspace Guardrail

Run `gws auth status` before non-trivial Google Workspace work. If authentication is unhealthy, stop and report the failure.

On Windows PowerShell, do not rely on shell-quoted inline JSON for non-trivial `--params` or `--json` values. Use [scripts/run_gws.py](scripts/run_gws.py) with UTF-8 JSON files. Check unfamiliar routes with `gws schema <service.resource.method>`.

## Specialist Skill Guardrail

When a dedicated skill is available, read it and follow its workflow:

- HWPX editing: preserve the package and validate the result; do not parse an `.hwpx` file directly as XML.
- STL to STEP/STP: use a deterministic local converter and preserve the requested output extension.
- Browser verification: use the configured browser skill or plugin.
- Presentations and documents: use the configured artifact skill and verify the rendered output.

## Mutation Guardrail

Before a recursive move, bulk rename, overwrite, remote upload, or remote write:

1. Resolve and inspect source and destination paths.
2. Preview the affected items.
3. Avoid overwriting existing outputs unless explicitly requested.
4. Perform the operation.
5. Re-read or re-list the result.

## Resources

- [scripts/run_gws.py](scripts/run_gws.py): pass JSON safely to the native `gws` executable from Windows.
- [references/google-workspace.md](references/google-workspace.md): authenticated and schema-checked Workspace patterns.
- [references/documents-content.md](references/documents-content.md): education content and document workflows.
- [references/3d-media-browser.md](references/3d-media-browser.md): 3D, browser, and media workflows.
- [references/environment-safety.md](references/environment-safety.md): portable Windows and setup rules.
