# PPT Builder Skill

`PPT Builder Skill` is a reusable workflow for generating brand-consistent slides from a corporate template.

It focuses on:
- Reusing template layouts instead of manual redraw.
- Building a lightweight template knowledge base for semantic page mapping.
- Generating common page types: cover, toc, section, content, end (`Thanks`).

## Repository Structure

- `daocloud-ppt.py`: project demo script that extracts assets, builds knowledge base, and generates `cc.pptx`.
- `openclaw-skill/ppt-template-builder/SKILL.md`: OpenClaw skill description.
- `openclaw-skill/ppt-template-builder/manifest.yaml`: OpenClaw machine-readable manifest.
- `openclaw-skill/ppt-template-builder/src/index.py`: skill runtime entry.
- `Makefile`: one-command package target for OpenClaw skill zip.

## Prerequisites

- Python `3.10+`
- Install dependency:

```bash
python3 -m pip install python-pptx
```

## Quick Start

### 1. Generate demo deck from current project script

```bash
cd /Users/peterpan/go/src/PPT_Builder_Skill
python3 daocloud-ppt.py
```

Expected outputs:
- `cc.pptx`
- `template_knowledge_base.json`
- `extracted_company_assets/`

### 2. Generate deck via OpenClaw skill entry

```bash
cd /Users/peterpan/go/src/PPT_Builder_Skill
python3 - <<'PY'
import asyncio
import importlib.util
from pathlib import Path
p = Path("openclaw-skill/ppt-template-builder/src/index.py")
spec = importlib.util.spec_from_file_location("skill_index", p)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)
res = asyncio.run(mod.handler({"title": "OpenClaw Demo", "output_filename": "cc_skill.pptx"}, None))
print(res)
PY
```

Expected outputs:
- `openclaw-skill/ppt-template-builder/cc_skill.pptx`

## Semantic Layout Mapping

Default semantic mapping in this project:
- `cover` -> `封面_Cover`
- `toc` -> `目录页_Content`
- `section` -> `章节页_Section page`
- `sub_section` -> `副章节页_Sub section page`
- `content` -> `标准内容页_Standard page`
- `content_with_subtitle` -> `标准内容页（小标题）_Standard page with subtitle`
- `end` -> `封底_End page`

The end page (`Thanks`) reference is included in the knowledge base as:
- `thanks_page_reference.expected_layout = "封底_End page"`
- `thanks_page_reference.expected_text = "Thanks."`

## Packaging And Publishing

Single package mode (always include `PPT_Template.pptx`):

```bash
cd /Users/peterpan/go/src/PPT_Builder_Skill
make package-openclaw
make demo-pages
make extract-catalog
```

Output:
- `dist/ppt-template-builder-openclaw-official.zip`

Skill source directory:
- `openclaw-skill/ppt-template-builder/`

## Git Workflow (with clash proxy env)

If you need to push to remote:

```bash
cd /Users/peterpan/go/src/PPT_Builder_Skill
export https_proxy=http://127.0.0.1:7897 http_proxy=http://127.0.0.1:7897 all_proxy=socks5://127.0.0.1:7897
git add README.md Makefile openclaw-skill/ppt-template-builder
git commit -m "feat: add openclaw skill package and make target"
git push
```

## Notes

- Keep template ownership/compliance in mind before publishing package artifacts.
- For a new template from another company, regenerate the knowledge base and update semantic mapping instead of hardcoding slide indices.
