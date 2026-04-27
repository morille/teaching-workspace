---
name: presentation-guidelines
description: Presentation production guidelines for readable, reference-driven, classroom, lecture, training, workshop, and style-matched slide decks. Use with presentation, PowerPoint, PPT, PPTX, slide deck, deck refresh, HTML slide prototype, or visual QA tasks when the user asks for clear projection readability, reference screenshot style adaptation, layout variety, semantic image use, text wrapping checks, or long-term presentation quality standards.
---

# Presentation Guidelines

Apply this skill as an overlay on top of the normal presentation/PowerPoint workflow. It does not replace the deck-building runtime; it adds quality gates for readability, reference-driven layout, asset choice, and QA.

## Core Rules

- Prioritize projection readability over visual density. Body text must be at least 18 pt; for live teaching, lecture, workshop, or training decks, prefer 22-26 pt body/callout text and split content across slides instead of shrinking copy.
- Keep each slide focused on one information center. Separate title, body, and visual regions clearly, and avoid turning a slide into a dense note page.
- Default normal body copy to left alignment. Center text only for intentional display objects such as cover titles, section numbers, badges, step markers, compact labels, or diagram nodes.
- Design text boxes for the actual language and expected wrapping, especially CJK text. Leave generous width and height slack so text wraps completely inside its shape without clipped lines, orphaned punctuation, or collisions.
- Do not use bottom hint/tip bars, status strips, or footer callouts as a default slide template. Add them only when the user explicitly asks for persistent tips, citations, progress, or handout-style notes.
- Cover slides must pass a contrast check. Do not place light text on pale images, pale gradients, or low-contrast fields without a deliberate scrim, solid text field, darker type color, or different crop.

## Reference Style Use

When the user supplies screenshots, a sample deck, or visual references:

- Treat references as style evidence, not content assets. Infer layout rhythm, title/body proportion, whitespace, alignment, color-block behavior, image-to-text ratio, and density.
- Do not reuse reference screenshots as content imagery unless the slide is explicitly analyzing that screenshot.
- Do not mechanically clone the reference. Adapt the visual logic to the new topic, audience, and content shape.
- If the reference set contains multiple body-slide layout types, extract several composition archetypes and vary them across the deck.
- If the user requests layout variety, use at least three distinct body-slide archetypes when the slide count allows, while preserving a coherent visual system.

## Asset Rules

- Every formal content slide should have a semantic visual when the user requests visuals or when the topic benefits from visual explanation. The asset must help explain the slide, not merely fill space.
- Do not make all visuals SVG/diagram-like by default. Use real or generated raster images when the slide meaning benefits from mood boards, photographic references, software screenshots, workspace/file examples, material libraries, or concrete visual evidence.
- Use diagrams, native shapes, charts, tables, and structured graphics when relationships, process, hierarchy, comparison, or classification are the point.
- Before accepting an asset, answer: "Why does this asset belong on this slide, and what information does it carry?" If the answer is decorative or vague, choose a better asset or redesign the slide.
- Maintain visual consistency across a deck. Avoid mixing unrelated styles such as flat icons, realistic photos, and futuristic gradients unless the contrast is deliberate and explained by the content.

## Layout QA

Before delivery or preview:

- Inspect every slide at readable size, not only as a montage/contact sheet.
- Check that title, subtitle, body, labels, captions, and visual regions do not overlap.
- Check that all text fits its frame after actual rendering. For important text, verify that rendered text height fits the box and the widest rendered line fits the available width when layout metrics are available.
- Check that required one-line labels remain one line; if they wrap unintentionally, widen the region, shorten the text, or intentionally redesign the label as a multi-line object.
- Check that images are not stretched, low-resolution, accidentally cropped, or clipped through important subjects, labels, faces, products, or diagrams.
- When text feels too small or crowded, shorten copy, split the slide, or enlarge the layout region. Do not solve readability problems by shrinking text below the legibility floor.

## HTML Slide Prototypes

When the user asks for an HTML slide prototype instead of PPTX:

- Preserve the same readability, wrapping, contrast, and semantic-asset rules.
- Use a fixed 16:9 stage unless the user asks otherwise.
- Make text boxes left-aligned by default and ensure the CSS allows complete wrapping without overflow.
- Remove default bottom tip bars unless explicitly requested.
- Test the prototype at the intended presentation size and check for overlap, clipping, and low contrast.
