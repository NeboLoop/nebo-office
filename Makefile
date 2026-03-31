.PHONY: build manifest upload clean help

build: ## Cross-compile for all platforms and update plugin.json
	./build.sh

manifest: ## Show plugin.json
	@cat plugin.json | jq .

upload: ## Upload plugin binaries to NeboLoop (requires PLUGIN_ID and TOKEN env vars)
	@if [ -z "$$PLUGIN_ID" ] || [ -z "$$TOKEN" ]; then echo "Usage: PLUGIN_ID=... TOKEN=... make upload"; exit 1; fi
	@for PLATFORM in darwin-arm64 darwin-amd64 linux-arm64 linux-amd64 windows-amd64; do \
		BINARY=$$(ls dist/plugin/$$PLATFORM/nebo-office* 2>/dev/null | head -1); \
		if [ -n "$$BINARY" ]; then \
			echo "Uploading $$PLATFORM..."; \
			curl --http1.1 -s -X POST "https://neboloop.com/api/v1/developer/apps/$$PLUGIN_ID/binaries" \
				-H "Authorization: Bearer $$TOKEN" \
				-F "file=@$$BINARY" \
				-F "platform=$$PLATFORM" \
				-F "skill=@PLUGIN.md" \
			| jq -r 'if .id then "OK" else . end'; \
		fi; \
	done

clean: ## Remove build artifacts
	rm -rf dist/plugin/

help: ## Show this help
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | sort | awk 'BEGIN {FS = ":.*?## "}; {printf "  \033[36m%-15s\033[0m %s\n", $$1, $$2}'
