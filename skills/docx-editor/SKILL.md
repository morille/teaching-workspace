---
name: docx-editor
description: Use when reading, creating, modifying, formatting, or validating Microsoft Word .docx files, especially Chinese application forms, reports, tables, proposals, and official documents.
---

# DOCX Editor

Use this skill for `.docx` and Microsoft Word document editing.

## Core principles

- Preserve the original document structure, formatting, tables, numbering, headers, footers, page breaks, and styles whenever possible.
- Never overwrite the original file. Always save an edited copy with a clear suffix, such as `_edited`, `_filled`, or `_reviewed`.
- For table-based forms, edit text inside existing cells instead of rebuilding the table.
- For Chinese official documents, project applications, teaching documents, and review forms, use formal, concise, institutionally appropriate language.
- Before making large changes, inspect the document structure: paragraphs, tables, headings, section breaks, and embedded objects.

## Preferred tools

- Use `python-docx` for structured `.docx` editing.
- Use LibreOffice headless mode to convert `.docx` to PDF for layout checking.
- Use PDF or image rendering to inspect whether tables overflow, pagination changes unexpectedly, or fonts are missing.
- Use ZIP/XML inspection only when `python-docx` cannot access the needed structure.

## Workflow

1. Make a backup copy of the original file.
2. Inspect the document structure:
   - Paragraph count
   - Table count
   - Heading styles
   - Section and page setup
   - Headers and footers
3. Identify the exact paragraphs or table cells to modify.
4. Apply minimal edits while preserving formatting.
5. Save as a new `.docx`.
6. Convert to PDF or rendered page images for layout verification.
7. Report:
   - What was changed
   - Where it was changed
   - Whether formatting or layout issues remain

## Special handling for application forms

- Do not delete form labels or instructions.
- Keep answers within the required word limits.
- If a cell has limited space, compress the wording rather than changing the table layout.
- For research or project applications, maintain consistency among title, research object, objectives, methods, expected outcomes, and schedule.
- For member contribution sections, make each person's role concrete and differentiated.

## Safety

- Do not fabricate official facts, awards, project levels, publication details, dates, or funding amounts.
- If source material is incomplete, mark uncertain content clearly instead of inventing it.
