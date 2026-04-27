# Lesson Plan DOCX Reference

## Workspace Paths

Use the repository at `D:\codex\teaching-workspace`.

Important files:

- Style source: `archive\lesson-plan\中国文化实验动画创作25-26（2）第一课次教案-沿用24-25格式 - 副本.docx`
- Generator: `lesson-plan\scripts\generate_lesson_docx.py`
- Original generation notes: `lesson-plan\scripts\教案生成脚本规范.md`
- Generated lesson plans: `lesson-plan\lesson-XX\docs\`

## Template Fidelity

The approved DOCX style comes from the user's hand-adjusted first lesson plan. Treat it as the source of truth for cover/header typography, course-info table typography, main table dimensions, row heights, line spacing, page breaks, blank-line rhythm, and heading emphasis.

Do not create a fresh Word document for official lesson plans unless the user explicitly asks.

## Content Density

The hand-adjusted first lesson has these practical paragraph targets:

- Teaching objectives cell: 22 paragraphs.
- Teaching content cell: 93 paragraphs.
- Teaching design cell: 73 paragraphs.
- Teaching note cell: blank template space.

When adding new lessons, stay close to these counts. If content is too long, compress prose and combine related details. Do not carry over huge unused blank areas from older templates.

## Fixed Structure

Every lesson plan should keep:

1. 课次、授课题目、学时。
2. 教学目标：知识目标、能力目标、素养目标。
3. 重点难点。
4. 学情分析。
5. 教学内容。
6. 教学设计。
7. 教学后记。

## 教学内容 Structure

Use this recurring structure:

- `【课前自主学习】`
- `明确目标，发布任务`
- `线上预习，总结归纳`
- `师生讨论，辅导答疑`
- `【课中讲析讨论】`
- `一、课程内容分享/上节实践作业展示`
- `二、教师讲评，评价学习效果`
- `三、教师精讲，引导合作探究`
- `四、课终小结`
- `课堂测试`
- `五、作业布置`
- `【课后实践应用】`

Classroom tests should be directly usable in class. Prefer full multiple-choice questions with options and answers over placeholder prompts.

## 教学设计 Structure

Keep the design cell concise and activity-focused:

- 课前启发式教学。
- 课中 180 分钟分配。
- 教师活动。
- 学生活动。
- 课程思政设计。
- 内容总结。
- 作业布置。

Keep `教师：` and `学生：` prefixes. Preserve their emphasis style where inherited from the template.

## Course-Specific Rule

For `中国文化实验动画创作`, always align each lesson to the 2025-2026 syllabus/progress table before using the previous year's lesson content. Previous-year lessons are reference material, not the source of sequencing.

Example: the 2025-2026 second lesson is about material classification and Chinese cultural experimental animation exploration. It should not blindly reuse the previous year's second lesson on full production principles and workflow.
