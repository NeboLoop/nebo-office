---
name: nebo-office
description: "Document generation binary for Word (.docx), Excel (.xlsx), and PowerPoint (.pptx). Converts JSON specs into formatted Office documents and unpacks existing documents back to JSON. Skills access the binary via $NEBO_OFFICE_BIN."
version: "0.1.0"
license: MIT
---

# nebo-office — Document Generation

Compiled Rust binary for creating, editing, and validating Word, Excel, and PowerPoint files. Converts JSON specs to Office documents and unpacks existing documents back to the same JSON format. Skills access the binary path via `$NEBO_OFFICE_BIN`.

## Services

| Skill | Capability |
|-------|-----------|
| `docx` | Create, unpack, and validate Word documents (.docx) |
| `xlsx` | Create, unpack, and validate Excel spreadsheets (.xlsx) |
| `pptx` | Create, unpack, and validate PowerPoint presentations (.pptx) |
