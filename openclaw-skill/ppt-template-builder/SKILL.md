---
name: ppt-template-builder
version: 1.0.0
description: Build brand-consistent PPTX from block/xml using a corporate PPT template.
categories:
  - productivity
  - presentation
---

# PPT Template Builder

## Description
Generate a new `.pptx` by reusing existing template slides from `PPT_Template.pptx`.
The runtime is XML-only and prioritizes template-first rendering.

## When To Use
Use this skill when the user asks:
- "Render PPT from structured block/xml"
- "Keep company template style (background/icon/font/layout)"
- "Avoid html-to-ppt export as primary path"

## Inputs
- `title`: optional deck title fallback
- `output_filename`: output pptx filename
- `template_file`: optional template filename in skill directory
- `block_xml`: block/xml payload (required)
- `render_strategy`: `template_first` | `template_only`
- `allow_html_fallback`: allow html/richtext block fallback to template content slides

## Output
- `output_path`: absolute path to generated pptx
- `slide_count`: number of slides generated
- `message`: status message

## Notes
- Requires `python-pptx`.
- Uses local `PPT_Template.pptx` in skill directory by default.
- Use `assets/demo_blocks.xml` for a full XML demo (toc + matrix + pie/bar/line + table + slogan).
