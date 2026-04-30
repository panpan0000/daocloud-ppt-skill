# OpenClaw Skill: ppt-template-builder

## Install (local path)

```bash
# Put this folder under your OpenClaw skills directory, e.g.
# ~/.openclaw/workspace/skills/ppt-template-builder
```

## Test

```bash
openclaw chat
```

Try:
- "Use ppt-template-builder to generate a deck with title Quarterly Review."
- "Use ppt-template-builder in examples mode."
- "Use ppt-template-builder in complex mode."

## Dependencies

- python-pptx

```bash
python3 -m pip install -r requirements.txt
```

## Local Commands

```bash
cd /Users/peterpan/go/src/PPT_Builder_Skill
make demo-pages
make demo-pages-complex
make extract-catalog
make package-openclaw
```
