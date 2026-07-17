SHELL := /bin/bash

DOTNET_ARTIFACTS ?= /tmp/pptxgo-dotnet-artifacts
COVERAGE_PROFILE ?= coverage.out
COVERAGE_HTML ?= coverage.html
DEMO_PPTX ?= examples/01_basic/01_basic_demo.pptx

.PHONY: help test build coverage dotnet-build examples validate check clean

help: ## Show available targets
	@awk 'BEGIN {FS = ":.*## "; printf "Available targets:\n"} /^[a-zA-Z0-9_-]+:.*## / {printf "  %-14s %s\n", $$1, $$2}' $(MAKEFILE_LIST)

test: ## Run all Go tests
	go test ./...

build: ## Build every package
	go build ./...

coverage: ## Generate Go coverage profile and HTML report
	go test -coverprofile=$(COVERAGE_PROFILE) ./...
	go tool cover -html=$(COVERAGE_PROFILE) -o $(COVERAGE_HTML)

dotnet-build: ## Build the Open XML validator
	dotnet build PptxValidator/PptxValidator.csproj --artifacts-path $(DOTNET_ARTIFACTS)

examples: ## Run example programs
	cd examples/01_basic && go run main.go

# Validation is not optional CI polish here: docxgo shipped nine months
# with a working OpenXmlValidator that CI never invoked. pptxgo runs it on
# every check from day one.
validate: examples dotnet-build ## Generate a demo .pptx and validate it against the OpenXML SDK
	dotnet $(DOTNET_ARTIFACTS)/bin/PptxValidator/debug/PptxValidator.dll $(DEMO_PPTX)

check: test build validate ## Run the full validation suite (tests, build, schema validation)

clean: ## Remove local build, coverage, and demo artifacts
	rm -rf $(COVERAGE_PROFILE) $(COVERAGE_HTML) $(DEMO_PPTX)
