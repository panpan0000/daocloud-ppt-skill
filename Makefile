SHELL := /bin/bash

SKILL_NAME := ppt-template-builder
OPENCLAW_SKILL_DIR := openclaw-skill/$(SKILL_NAME)
DIST_DIR := dist
OPENCLAW_ZIP := $(DIST_DIR)/$(SKILL_NAME)-openclaw-official.zip

.PHONY: package-openclaw clean-dist verify-skill-tree

verify-skill-tree:
	@test -f "$(OPENCLAW_SKILL_DIR)/SKILL.md"
	@test -f "$(OPENCLAW_SKILL_DIR)/manifest.yaml"
	@test -f "$(OPENCLAW_SKILL_DIR)/src/index.py"
	@test -f "$(OPENCLAW_SKILL_DIR)/PPT_Template.pptx"

package-openclaw: verify-skill-tree
	@mkdir -p "$(DIST_DIR)"
	@rm -f "$(OPENCLAW_ZIP)"
	@cd openclaw-skill && zip -r "../$(OPENCLAW_ZIP)" "$(SKILL_NAME)" >/dev/null
	@echo "Built: $(OPENCLAW_ZIP)"

clean-dist:
	@rm -rf "$(DIST_DIR)"
	@echo "Cleaned: $(DIST_DIR)"
