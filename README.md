# PPT Builder Skill

`PPT Builder Skill` is a reusable workflow for generating brand-consistent slides from a corporate template.

It focuses on:
- Reusing template layouts instead of manual redraw.
- Building a lightweight template knowledge base for semantic page mapping.
- Generating common page types: cover, toc, section, content, end (`Thanks`).

## Repository Structure

- `daocloud-ppt.py`: project demo script that extracts assets, builds knowledge base, and generates `cc.pptx`.
- `.trae/skills/ppt-template-builder/SKILL.md`: reusable skill definition for agent invocation.
- `.trae/skills/ppt-template-builder/scripts/generate_from_template.py`: standalone generator script for packaging.
- `.trae/skills/ppt-template-builder/assets/template_knowledge_base.sample.json`: sample knowledge base for distribution.

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

### 2. Generate deck via skill script

```bash
cd /Users/peterpan/go/src/PPT_Builder_Skill
python3 .trae/skills/ppt-template-builder/scripts/generate_from_template.py \
  --template PPT_Template.pptx \
  --output cc_skill.pptx \
  --kb template_knowledge_base_skill.json
```

Expected outputs:
- `cc_skill.pptx`
- `template_knowledge_base_skill.json`

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
mkdir -p dist
zip -r dist/ppt-template-builder-openclaw.zip \
  README.md \
  .trae/skills/ppt-template-builder \
  PPT_Template.pptx
```

Send `dist/ppt-template-builder-openclaw.zip` to other OpenClaw users.

## Git Workflow (with clash proxy env)

If you need to push to remote:

```bash
cd /Users/peterpan/go/src/PPT_Builder_Skill
export https_proxy=http://127.0.0.1:7897 http_proxy=http://127.0.0.1:7897 all_proxy=socks5://127.0.0.1:7897
git add README.md daocloud-ppt.py .trae/skills/ppt-template-builder
git commit -m "feat: add reusable ppt-template-builder skill package"
git push
```

## Notes

- Keep template ownership/compliance in mind before publishing package artifacts.
- For a new template from another company, regenerate the knowledge base and update semantic mapping instead of hardcoding slide indices.
