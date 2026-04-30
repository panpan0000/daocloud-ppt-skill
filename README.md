# PPT Builder Skill

`PPT Builder Skill` is an OpenClaw skill for XML-driven PPTX generation with strict template reuse.

## Scope

- Runtime is XML-only (`block_xml` input).
- Rendering is template-first and keeps corporate style (background/icon/font/layout).
- HTML blocks are optional fallback, not the primary path.

## Repository Structure

- `openclaw-skill/ppt-template-builder/src/index.py`: XML renderer runtime.
- `openclaw-skill/ppt-template-builder/manifest.yaml`: OpenClaw manifest.
- `openclaw-skill/ppt-template-builder/SKILL.md`: skill description.
- `openclaw-skill/ppt-template-builder/assets/demo_blocks.xml`: demo XML payload.
- `openclaw-skill/ppt-template-builder/tools/extract_page_catalog.py`: template catalog utility.
- `Makefile`: demo and packaging commands.

## Quick Start

```bash
cd /PPT_Builder_Skill
python3 -m pip install -r openclaw-skill/ppt-template-builder/requirements.txt
make demo-pages-xml
make package-openclaw
```

Outputs:

- `openclaw-skill/ppt-template-builder/examples_demo_xml.pptx`
- `dist/ppt-template-builder-openclaw-official.zip`

## Optional Utility

```bash
cd /Users/peterpan/go/src/PPT_Builder_Skill
make extract-catalog
```

## Git Push (with clash proxy)

```bash
cd /Users/peterpan/go/src/PPT_Builder_Skill
export https_proxy=http://127.0.0.1:7897 http_proxy=http://127.0.0.1:7897 all_proxy=socks5://127.0.0.1:7897
git add README.md Makefile openclaw-skill/ppt-template-builder
git commit -m "refactor: keep xml-only template-first skill path"
git push
```
