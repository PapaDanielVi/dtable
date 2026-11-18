SHELL := /bin/zsh
.DEFAULT_GOAL := help

NAME := dtable
VERSION := $(shell node -p "require('./manifest.json').version")
DIST_DIR := dist
BUILD_DIR := $(DIST_DIR)/$(NAME)
ZIP_FILE := $(DIST_DIR)/$(NAME)-v$(VERSION).zip

SHEETJS_VERSION := 0.20.3
SHEETJS_URL := https://cdn.sheetjs.com/xlsx-$(SHEETJS_VERSION)/package/dist/xlsx.full.min.js

RUNTIME_FILES := \
	manifest.json \
	background.js \
	content.js \
	content.css \
	popup.html \
	popup.css \
	popup.js

.PHONY: help check lint test build package zip clean update-sheetjs open-chrome

help: ## Show available commands
	@awk 'BEGIN {FS = ":.*##"; printf "Usage: make <target>\n\nTargets:\n"} /^[a-zA-Z0-9_-]+:.*?##/ {printf "  %-16s %s\n", $$1, $$2}' $(MAKEFILE_LIST)

check: ## Validate manifest JSON, JS syntax, and required assets
	@node -e "JSON.parse(require('fs').readFileSync('manifest.json', 'utf8'))"
	@node --check background.js
	@node --check content.js
	@node --check popup.js
	@test -f lib/xlsx.full.min.js || (echo "Missing lib/xlsx.full.min.js. Run 'make update-sheetjs'." && exit 1)
	@test -f icons/icon16.png && test -f icons/icon32.png && test -f icons/icon48.png && test -f icons/icon128.png
	@echo "Syntax and asset checks passed."

lint: check ## Alias for check

test: check ## Run syntax checks and background export unit tests
	@node scripts/verify.js

build: check ## Assemble unpacked extension files into dist/dtable
	@rm -rf $(BUILD_DIR)
	@mkdir -p $(BUILD_DIR)/icons $(BUILD_DIR)/lib
	@cp $(RUNTIME_FILES) $(BUILD_DIR)/
	@cp icons/*.png $(BUILD_DIR)/icons/
	@cp lib/xlsx.full.min.js $(BUILD_DIR)/lib/
	@echo "Built unpacked extension at $(BUILD_DIR)"

package: test build ## Create distributable .zip archive in dist/
	@rm -f $(ZIP_FILE)
	@cd $(BUILD_DIR) && zip -qr ../$(NAME)-v$(VERSION).zip .
	@echo "Packaged $(ZIP_FILE)"

zip: package ## Alias for package

update-sheetjs: ## Download SheetJS xlsx.full.min.js locally into lib/
	@mkdir -p lib
	@curl -fsSL -o lib/xlsx.full.min.js $(SHEETJS_URL)
	@echo "Updated lib/xlsx.full.min.js to v$(SHEETJS_VERSION)"

open-chrome: ## Open Chrome Extensions page (macOS)
	@open -a "Google Chrome" "chrome://extensions"

clean: ## Remove build artifacts in dist/
	@rm -rf $(DIST_DIR)
	@echo "Cleaned $(DIST_DIR)/"
