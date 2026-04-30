SHELL := /bin/bash

SKILL_NAME := corporate-ppt-generator
OPENCLAW_SKILL_DIR := openclaw-skill/$(SKILL_NAME)
DIST_DIR := dist
OPENCLAW_ZIP := $(DIST_DIR)/$(SKILL_NAME)-openclaw-official.zip
XML_DEMO_OUTPUT := $(OPENCLAW_SKILL_DIR)/examples_demo_xml.pptx
XML_DEMO_INPUT := $(OPENCLAW_SKILL_DIR)/assets/demo_blocks.xml
PAGE_CATALOG := $(OPENCLAW_SKILL_DIR)/assets/page_catalog.json
TEMPLATE_PIPELINE := tools/template_pipeline.py

.PHONY: package-openclaw clean-dist verify-skill-tree demo extract-catalog run-template-pipeline

verify-skill-tree:
	@test -f "$(OPENCLAW_SKILL_DIR)/SKILL.md"
	@test -f "$(OPENCLAW_SKILL_DIR)/manifest.yaml"
	@test -f "$(OPENCLAW_SKILL_DIR)/src/index.py"
	@test -f "$(OPENCLAW_SKILL_DIR)/PPT_Template.pptx"

package-openclaw: verify-skill-tree
	@mkdir -p "$(DIST_DIR)"
	@rm -f "$(OPENCLAW_ZIP)"
	@cd openclaw-skill && zip -r "../$(OPENCLAW_ZIP)" "$(SKILL_NAME)" \
		-x "*/__pycache__/*" \
		-x "*/examples_demo.pptx" \
		-x "*/examples_demo_complex.pptx" \
		-x "*/examples_demo_xml.pptx" >/dev/null
	@echo "Built: $(OPENCLAW_ZIP)"

demo: verify-skill-tree
	@python3 "$(OPENCLAW_SKILL_DIR)/src/index.py" \
		--output "$(notdir $(XML_DEMO_OUTPUT))" \
		--block-xml-file "$(XML_DEMO_INPUT)" \
		--render-strategy template_first \
		--allow-html-fallback
	@echo "XML demo deck: $(XML_DEMO_OUTPUT)"

extract-catalog: verify-skill-tree
	@python3 "$(OPENCLAW_SKILL_DIR)/tools/extract_page_catalog.py" \
		--template "$(OPENCLAW_SKILL_DIR)/PPT_Template.pptx" \
		--output "$(PAGE_CATALOG)"
	@echo "Catalog: $(PAGE_CATALOG)"

run-template-pipeline:
	@python3 "$(TEMPLATE_PIPELINE)"

clean-dist:
	@rm -rf "$(DIST_DIR)"
	@echo "Cleaned: $(DIST_DIR)"
