---
name: nebo-office
description: "Document generation binary for Word (.docx), Excel (.xlsx), and PowerPoint (.pptx). Converts JSON specs into formatted Office documents and unpacks existing documents back to JSON. Skills depend on this plugin to get the binary via $NEBO_OFFICE_BIN."
version: "0.1.0"
license: MIT
---

# nebo-office — Document Generation

Compiled Rust binary for creating, editing, and validating Word, Excel, and PowerPoint files. Converts JSON specs to Office documents and unpacks existing documents back to the same JSON format.

## Env Var

Skills access the binary path via `$NEBO_OFFICE_BIN`.

## Supported Formats

| Format | Commands |
|--------|----------|
| Word (.docx) | `nebo-office docx create/unpack/validate` |
| Excel (.xlsx) | `nebo-office xlsx create/unpack/validate` |
| PowerPoint (.pptx) | `nebo-office pptx create/unpack/validate` |
