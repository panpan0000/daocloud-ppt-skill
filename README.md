# PPT Builder Skill

`Corporate PPT Generator` is an OpenClaw skill for XML-driven PPTX generation with strict brand-style reuse.

## Scope

- Runtime is XML-only (`block_xml` input).
- Rendering is template-first and keeps corporate style (background/icon/font/layout).
- HTML blocks are optional fallback, not the primary path.

## Repository Structure

- `openclaw-skill/corporate-ppt-generator/src/index.py`: XML renderer runtime.
- `openclaw-skill/corporate-ppt-generator/manifest.yaml`: OpenClaw manifest.
- `openclaw-skill/corporate-ppt-generator/SKILL.md`: skill description.
- `openclaw-skill/corporate-ppt-generator/assets/demo_blocks.xml`: demo XML payload.
- `openclaw-skill/corporate-ppt-generator/tools/extract_page_catalog.py`: template catalog utility.
- `Makefile`: demo and packaging commands.

## Quick Start

```bash
cd PPT_Builder_Skill
python3 -m pip install -r openclaw-skill/corporate-ppt-generator/requirements.txt
make demo
make package-openclaw
```

Outputs:

- `openclaw-skill/corporate-ppt-generator/examples_demo_xml.pptx`
- `dist/corporate-ppt-generator-openclaw-official.zip`

## Release Download

- Tag push matching `v*` triggers GitHub Actions release packaging automatically.
- Download from GitHub Releases asset: `corporate-ppt-generator-openclaw-official.zip`.
- First public release tag: `v0.1`.
