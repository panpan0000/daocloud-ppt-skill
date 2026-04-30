---
name: ppt-template-builder
version: 1.0.0
description: Build brand-consistent PPTX from a corporate template. Invoke when user asks to generate cover/toc/section/content/end slides with template style.
categories:
  - productivity
  - presentation
---

# PPT Template Builder

## Description
Generate a new `.pptx` by reusing layouts from `PPT_Template.pptx`.
This skill keeps visual consistency for:
- cover
- toc
- section
- content
- end (`Thanks`)

## When To Use
Use this skill when the user asks:
- "Generate slides based on our template"
- "Keep company style and layout"
- "Create cover/toc/section/content/end quickly"

## Inputs
- `title`: deck title
- `toc_items`: toc lines
- `section_title`: section page title
- `content_title`: content page title
- `content_body`: content body text
- `output_filename`: output pptx filename

## Output
- `output_path`: absolute path to generated pptx
- `slide_count`: number of slides generated
- `message`: status message

## Notes
- Requires `python-pptx`.
- Uses local `PPT_Template.pptx` in skill directory by default.
