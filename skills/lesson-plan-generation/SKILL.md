---
name: lesson-plan-generation
description: Generate and revise Chinese course lesson-plan DOCX files from approved teaching-plan templates. Use when creating, updating, comparing, or formatting 教案/lesson plan Word documents, especially 中国文化实验动画创作 lesson plans that must preserve prior-year table layout, fonts, line height, paragraph density, section headings, page breaks, and DOCX template style.
---

# Lesson Plan Generation

Use this skill when producing or revising formal course lesson-plan `.docx` files. The priority is template fidelity: preserve the hand-tuned Word format first, then update content.

## Core Workflow

1. Locate the current teaching workspace at `D:\codex\teaching-workspace`.
2. Use the hand-adjusted first lesson plan as the style source when available: `archive\lesson-plan\中国文化实验动画创作25-26（2）第一课次教案-沿用24-25格式 - 副本.docx`.
3. Align the new lesson content with the current syllabus/progress table before writing.
4. Generate or revise `.docx` by replacing text inside the existing template cells. Do not rebuild the table, margins, row heights, or page geometry.
5. Verify paragraph counts and key fields after generation.

## Formatting Rules

- Preserve existing DOCX template styles instead of globally resetting fonts.
- Keep cover/header/course-info text in the inherited Microsoft YaHei style.
- Keep正文 table content in inherited Songti/宋体 style unless the template already says otherwise.
- Use the hand-tuned template paragraph density: 教学目标 about 22 paragraphs, 教学内容 about 90-95 paragraphs, 教学设计 about 70-75 paragraphs, 教学后记 blank.
- In 教学内容, keep major section headings as the template's 12 pt bold Songti style.
- Keep `教师：` and `学生：` prefixes and their emphasis style when present.
- Avoid adding many blank paragraphs. If content is too long, merge natural prose rather than expanding the template.
- Keep full Chinese words and phrases together when inserting manual line breaks. Do not split a word to force wrapping.

## Reusable Script

Prefer the repository script when the task matches its scope:

```powershell
& 'C:\Users\Morille\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' `
  'D:\codex\teaching-workspace\lesson-plan\scripts\generate_lesson_docx.py' `
  <lesson_no> `
  'D:\codex\teaching-workspace\archive\lesson-plan\中国文化实验动画创作25-26（2）第一课次教案-沿用24-25格式 - 副本.docx' `
  <output.docx>
```

Read [references/lesson-plan-docx.md](references/lesson-plan-docx.md) before changing the script, adding a new lesson, or manually editing generated DOCX formatting rules.

## QA Checklist

After generating a lesson plan:

- Confirm `课次`, `授课题目`, `学时`, and header semester are correct.
- Confirm the lesson follows the current syllabus/progress table, not the previous year's sequence.
- Confirm paragraph counts are close to the hand-tuned template.
- Open or inspect the DOCX to ensure Chinese text is not garbled.
- Compare key headings for bold/size consistency: `一、教学目标`, `二、重点难点`, `三、学情分析`, `【课前自主学习】`, `【课中讲析讨论】`, `【课后实践应用】`.
