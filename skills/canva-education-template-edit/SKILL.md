---
name: canva-education-template-edit
description: Edit an existing Canva presentation directly while preserving the original template design, layout, colors, decorations, and overall visual style. Use when the user shares a Canva design link or design ID and wants the original deck updated in place, especially for Korean education materials, classroom slides, lesson decks, or presentation templates where only the content should change and the template look should stay the same.
---

# Canva Education Template Edit

## Overview

Edit an existing Canva deck in place instead of generating a new presentation. Preserve the template look and replace only the content needed for the new lesson or education-focused presentation.

## Workflow

1. Identify the Canva design to edit from the provided design URL or design ID.
2. Start an editing transaction for that design before inspecting or changing all content.
3. Read the existing text and use the current page structure to map new content onto the existing slides.
4. Prefer replacing text only. Keep layout, background, colors, decorations, frames, and media unchanged unless the user explicitly asks otherwise.
5. Apply all requested text updates in draft mode with `perform_editing_operations`.
6. Show the preview thumbnail returned for the first updated page directly in chat.
7. If more than one page changed, fetch every other updated page thumbnail one page at a time and show each preview directly in chat.
8. Summarize the changes briefly and ask for explicit approval before committing.
9. Commit only after approval, then return the Canva link for review.

## Editing Rules

- Prefer direct editing of the original Canva design over creating a new design.
- Preserve the original template style unless the user explicitly asks to redesign it.
- Treat text replacement as the default operation.
- Do not replace or move images, videos, stickers, fills, or decorative elements unless the user asks for that change.
- If the user asks to replace an image on a page with multiple media elements, inspect all page assets first and confirm the correct target before updating it.
- If the request is underspecified, clarify the missing content before committing changes.
- If the link points to a Canva shortlink, resolve it first when the available Canva tools support that flow.

## Education Deck Defaults

- Write in Korean unless the user asks for another language.
- Use short, clear, student-friendly sentences for lesson slides.
- Compress long paragraphs into brief teaching points, labels, or checklist items.
- Favor text that fits the existing layout over literal completeness.
- If text looks crowded, shorten the copy before resizing or repositioning elements.
- Keep slide flow easy to teach from: introduction, concept explanation, examples, activity or practice, summary.

## Output Conventions

- Show every thumbnail returned by Canva directly in chat with image rendering.
- Keep change summaries concise and grouped by slide purpose rather than by raw operation list.
- Ask for final save approval exactly once, after previews are available.
- After commit, return the original Canva design link so the user can review the saved result.

## Example Requests

- "이 Canva 링크 직접 수정해줘. 템플릿은 그대로 두고 수업 PPT로 바꿔줘."
- "원본 Canva 디자인에 바로 반영해줘. 내용만 교육용 발표자료로 교체해줘."
- "이 템플릿 스타일 유지하고 한국어 수업 자료로 바꿔줘."
