# Documents And Education Content

Use this reference for lesson plans, curricula, presentations, HWPX, PDF, and HTML learning materials.

## Education Content Workflow

1. Identify audience, duration, learning objectives, delivery format, and evaluation method.
2. Gather and cite source material when factual accuracy matters.
3. Draft the instructional structure before styling.
4. Build the requested artifact with a dedicated document or presentation skill when available.
5. Render or open the artifact and verify content, layout, and links.

## Lesson Plan Structure

Use this as a starting point, adapting it to the request:

1. Overview: audience, duration, topic, prerequisites
2. Learning objectives
3. Session sequence
4. Learner activities
5. Required materials and tools
6. Assessment and reflection

## Presentations

- Use the configured presentation skill for substantial decks.
- Separate content outline, visual design, and final verification.
- When using Google Slides, create and inspect the presentation through the verified Workspace route.
- Export native Google Slides through `drive files export`, not `drive files get`.
- Verify the rendered or exported deck before reporting completion.

## HWPX

Treat HWPX as a ZIP/XML package, but do not use `Expand-Archive` directly on an `.hwpx` path and do not parse the entire `.hwpx` file as XML.

Prefer a dedicated HWPX skill. Preserve the source package, edit only the required XML content, write a new output file, and validate the final package.

When no dedicated helper is available, use a ZIP-aware library such as Python `zipfile` to inspect package entries. Never modify the original in place.

## PDF

- Use the configured PDF skill when available.
- Use the workspace `.venv` for Python-based extraction or generation.
- Verify page count, extracted text quality, and rendered output.
- Preserve the original unless the user explicitly requests replacement.

## HTML Learning Materials

1. Inspect the existing HTML, CSS, JavaScript, and asset structure.
2. Keep changes scoped to the requested behavior.
3. Open the material with the configured browser tool.
4. Check layout, console errors, navigation, and interactive behavior.
