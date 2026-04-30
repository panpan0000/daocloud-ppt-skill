---
name: "ppt-template-builder"
description: "Builds brand-consistent PPTX by reusing template layouts and a template knowledge base. Invoke when user asks to generate slides from a corporate PPT template."
---

# PPT Template Builder

## Purpose
This skill generates new `.pptx` files by reusing slide layouts from a corporate template instead of manually rebuilding visuals.
It supports reusable page semantics: `cover`, `toc`, `section`, `sub_section`, `content`, `end`.

## When To Invoke
Invoke this skill when:
- The user asks to generate or edit a deck based on an existing company template.
- The user wants consistent brand style (fonts, spacing, colors, page structures).
- The user asks for standard page types (cover, TOC, section page, content page, end page).
- The user asks to extract company assets (logos/backgrounds/icons) from a template.

Do not invoke this skill for:
- Pure data analysis with no slide generation.
- Ad-hoc custom design where no template reuse is required.

## Required Inputs
- `template_path`: path to source template `.pptx`.
- `output_path`: path to generated output `.pptx`.
- `page_plan`: ordered page list with semantic types and text payload.

Recommended:
- `knowledge_base_path`: `template_knowledge_base.json` generated from the template.

## Output Contract
- A generated deck that uses template layouts directly.
- A `template_knowledge_base.json` containing:
- Layout catalog.
- Semantic layout mapping.
- Sample slide references.
- End-page (`Thanks`) reference.
- Optional extracted media assets under `extracted_company_assets/`.

## Standard Workflow
1. Open template and build/update `template_knowledge_base.json`.
2. Resolve layout by semantic key (`cover`, `toc`, `section`, `content`, `end`).
3. Create slides with template layouts.
4. Fill placeholders by placeholder index order.
5. Save output deck.
6. Validate slide count and key pages (especially `end` / `Thanks`).

## Semantic Page Mapping (Default)
- `cover` -> `封面_Cover`
- `toc` -> `目录页_Content`
- `section` -> `章节页_Section page`
- `sub_section` -> `副章节页_Sub section page`
- `content` -> `标准内容页_Standard page`
- `content_with_subtitle` -> `标准内容页（小标题）_Standard page with subtitle`
- `end` -> `封底_End page`

If names differ in another template, regenerate the knowledge base and remap semantics by layout name.

## Page Authoring Rules
- Prefer modifying placeholder text only.
- Do not manually redraw brand components already provided by the template.
- Keep one semantic purpose per slide.
- Preserve original layout geometry and style primitives.
- Add end page (`Thanks`) for complete narrative closure unless user asks otherwise.

## Asset Reuse Rules
- Use extracted assets only when placeholders cannot satisfy requirements.
- Keep logos and brand marks in original aspect ratio.
- Avoid mixing external style assets unless explicitly requested.

## Packaging For Publishing
Use one of two distribution modes:

### Mode A: Lite Package (Recommended)
Includes:
- `SKILL.md`
- generation script(s)
- `template_knowledge_base.json` (structure/sample only)

Does not include:
- original `PPT_Template.pptx`
- full brand media dump

Use this mode when sharing publicly or when template ownership is restricted.

### Mode B: Full Internal Package
Includes:
- everything in Lite
- original `PPT_Template.pptx`
- approved brand assets

Use only when you have redistribution rights.

## Suggested Skill Bundle Structure
```text
.trae/skills/ppt-template-builder/
  SKILL.md
  assets/
    template_knowledge_base.sample.json
  scripts/
    generate_from_template.py
```

## Invocation Example
User intent:
"Create a 6-page deck from our template with cover, toc, two sections, one content page, and thanks page."

Expected behavior:
- Reuse template layouts from semantic mapping.
- Fill placeholders with provided text.
- Produce `.pptx` output and report generated path.

## Quality Checklist
- Output opens in PowerPoint/Keynote without corruption.
- Cover/TOC/Section/Content/End pages all present if requested.
- End page uses template end layout and contains `Thanks`.
- No placeholder overflow for major title/body fields.
- Style consistency preserved from template.
