---
name: nebo-office
description: "Use this skill whenever the user wants to create, read, edit, or manipulate Word documents (.docx files). Triggers include: any mention of 'Word doc', 'word document', '.docx', or requests to produce professional documents with formatting like headings, tables, page numbers, cover pages, or letterheads. Also use when extracting or reorganizing content from .docx files, working with tracked changes or comments, or converting content into a polished Word document. If the user asks for a 'report', 'memo', 'letter', 'proposal', 'template', or similar deliverable as a Word file, use this skill. Do NOT use for PDFs, spreadsheets, or Google Docs."
license: Proprietary. LICENSE.txt has complete terms
---

# nebo-office: DOCX Creation, Editing, and Analysis

## Overview

`nebo-office` is a precompiled Rust binary for working with Word documents. It converts a JSON spec into a fully-formatted DOCX file, and unpacks existing DOCX files back into the same JSON format for editing. No JavaScript, no Python, no code generation.

## Quick Reference

| Task | Approach |
|------|----------|
| Create new document | Write JSON spec → `nebo-office docx create` |
| Read/analyze content | `nebo-office docx unpack` → read JSON |
| Edit existing document | Unpack → edit JSON → create |
| Validate | `nebo-office docx validate` |

## Setup

Download the binary if not already present:
```bash
# Check if installed
which nebo-office || nebo-office version

# If not found, download for your platform
# macOS ARM: curl -L -o /usr/local/bin/nebo-office <download-url>/darwin-arm64/nebo-office && chmod +x /usr/local/bin/nebo-office
# macOS Intel: curl -L -o /usr/local/bin/nebo-office <download-url>/darwin-amd64/nebo-office && chmod +x /usr/local/bin/nebo-office
# Linux ARM: curl -L -o /usr/local/bin/nebo-office <download-url>/linux-arm64/nebo-office && chmod +x /usr/local/bin/nebo-office
# Linux Intel: curl -L -o /usr/local/bin/nebo-office <download-url>/linux-amd64/nebo-office && chmod +x /usr/local/bin/nebo-office
```

## CLI Usage

```bash
nebo-office docx create spec.json -o output.docx        # JSON → DOCX
nebo-office docx create spec.json -o output.docx --validate  # create + validate
nebo-office docx unpack input.docx -o spec.json          # DOCX → JSON
nebo-office docx validate spec.json                      # check spec
nebo-office docx validate output.docx                    # check DOCX structure
nebo-office version
```

Use `-` for stdin/stdout piping.

---

## JSON Spec Format

### Root Structure

```json
{
  "version": 1,
  "page": { },
  "styles": { },
  "metadata": { },
  "headers": { },
  "footers": { },
  "body": [ ]
}
```

Only `version` and `body` are required. All units: **inches** for dimensions, **points** for font sizes and spacing.

### Page Setup

```json
"page": {
  "size": "letter",
  "margin": { "top": 1.0, "bottom": 1.0, "left": 1.25, "right": 1.25 }
}
```

Sizes: `"letter"` (8.5×11), `"a4"`, `"legal"`, or `{ "width": 8.5, "height": 11 }`.
Margin: uniform number or `{ "top", "bottom", "left", "right" }`.

### Styles

```json
"styles": {
  "font": "Calibri",
  "size": 11,
  "color": "333333",
  "headings": {
    "font": "Calibri",
    "color": "1A3C5E",
    "h1": { "size": 28, "bold": true },
    "h2": { "size": 18, "bold": true },
    "h3": { "size": 14, "bold": true }
  },
  "custom": {
    "Callout": { "font": "Calibri", "size": 10, "italic": true, "color": "5A6A7A", "indent": { "left": 0.5, "right": 0.5 } },
    "CoverTitle": { "font": "Calibri", "size": 36, "bold": true, "color": "1A3C5E", "align": "center" }
  }
}
```

Custom style properties: `font`, `size`, `color`, `bold`, `italic`, `align`, `indent` `{ left, right }`, `spacing` `{ before, after }`.

### Metadata

```json
"metadata": {
  "title": "Document Title",
  "creator": "Author Name",
  "subject": "Subject",
  "keywords": ["key1", "key2"],
  "category": "Category",
  "description": "Description"
}
```

### Headers and Footers

```json
"headers": {
  "default": [{ "paragraph": { "text": "Company Name", "align": "right" } }],
  "first": []
},
"footers": {
  "default": [{
    "paragraph": {
      "runs": [{ "text": "Page " }, { "field": "page-number" }, { "text": " of " }, { "field": "total-pages" }],
      "align": "center"
    }
  }]
}
```

`"first": []` suppresses header/footer on the first page.

---

## Body Elements

### Headings

```json
{ "heading": 1, "text": "Title" }
{ "heading": 2, "text": "Section", "id": "bookmark-id" }
```

### Paragraphs

**Simple (with markdown-like formatting):**
```json
{ "paragraph": "Revenue grew **15%** year-over-year." }
```

Supported: `**bold**`, `*italic*`, `__underline__`, `~~strike~~`, `` `code` ``, `[text](url)`, `[^1]` (footnote ref).

**With properties:**
```json
{ "paragraph": { "text": "Centered text.", "align": "center", "style": "Callout", "spacing": { "before": 12, "after": 6 } } }
```

**Explicit runs (complex formatting):**
```json
{
  "paragraph": {
    "runs": [
      { "text": "Normal " },
      { "text": "red bold", "bold": true, "color": "DC2626" },
      { "tab": true },
      { "field": "page-number" }
    ],
    "align": "right"
  }
}
```

Run properties: `bold`, `italic`, `underline`, `strike`, `superscript`, `subscript`, `font`, `size`, `color`, `highlight`, `link`, `all-caps`, `small-caps`.

### Bullets

```json
{ "bullets": ["Item one", "Item **two**", { "text": "Parent", "children": ["Nested A", "Nested B"] }] }
```

### Numbered Lists

```json
{ "numbered": ["First", "Second", "Third"] }
{ "numbered": ["Step 1", "Step 2"], "restart": true }
```

Use `"restart": true` when a new numbered list should start from 1.

### Tables

**Simple:**
```json
{ "table": [["Name", "Role"], ["Alice", "Engineer"]], "header-rows": 1 }
```

**Full (with styling):**
```json
{
  "table": {
    "columns": [{ "width": 2.0 }, { "width": 4.0 }],
    "header-rows": 1,
    "rows": [
      { "cells": [
        { "text": "Header", "bold": true, "shading": "1A3C5E", "color": "FFFFFF" },
        { "text": "Header 2", "bold": true, "shading": "1A3C5E", "color": "FFFFFF" }
      ]},
      { "cells": [
        { "text": "Value" },
        { "text": "Description" }
      ]}
    ]
  }
}
```

Cell properties: `text`, `runs`, `body`, `bold`, `color`, `shading`, `align`, `valign`, `colspan`, `rowspan`.

**Zebra striping:** Alternate rows with `"shading": "F5F7FA"` for readability.

### Images

```json
{ "image": "chart.png", "width": 4.0, "height": 3.0, "alt": "Chart", "align": "center", "caption": "Figure 1" }
```

Images referenced by filename from the assets directory (same dir as JSON by default, or `--assets <dir>`).

### Page Break

```json
{ "page-break": true }
```

### Section Break

```json
{ "section-break": { "type": "next-page", "valign": "center", "columns": 2, "column-gap": 0.5 } }
```

Types: `"next-page"`, `"continuous"`, `"even-page"`, `"odd-page"`.
`"valign"`: `"top"`, `"center"`, `"bottom"` — vertically aligns content on the page.

### Table of Contents

```json
{ "toc": true }
{ "toc": { "title": "Contents", "depth": 3 } }
```

### Bookmark

```json
{ "bookmark": "anchor-name" }
```

---

## Cover Page Pattern

Use a section break with `valign: center` to create a vertically centered cover page:

```json
{
  "version": 1,
  "styles": {
    "font": "Calibri", "size": 11, "color": "333333",
    "headings": { "color": "1A3C5E", "h2": { "size": 18, "bold": true } },
    "custom": {
      "CoverTitle": { "size": 36, "bold": true, "color": "1A3C5E", "align": "center" },
      "CoverSubtitle": { "size": 16, "color": "5A6A7A", "align": "center" },
      "CoverMeta": { "size": 11, "color": "888888", "align": "center" }
    }
  },
  "headers": {
    "default": [{ "paragraph": { "text": "Company  |  Confidential", "align": "right" } }],
    "first": []
  },
  "footers": {
    "default": [{ "paragraph": { "runs": [{ "text": "Page " }, { "field": "page-number" }], "align": "center" } }]
  },
  "body": [
    { "paragraph": { "text": "Document Title", "style": "CoverTitle" } },
    { "paragraph": { "text": "Subtitle", "style": "CoverSubtitle" } },
    { "paragraph": { "text": "", "spacing": { "before": 200 } } },
    { "paragraph": { "text": "Author Name", "style": "CoverMeta" } },
    { "paragraph": { "text": "Date", "style": "CoverMeta" } },
    { "section-break": { "type": "next-page", "valign": "center" } },

    { "heading": 2, "text": "First Section" },
    { "paragraph": "Content begins here." }
  ]
}
```

The spacer paragraph (200pt before) pushes the author info down from the title, and `valign: center` centers the whole block on the page.

---

## Professional Table Pattern

Dark header with zebra-striped body:

```json
{
  "table": {
    "columns": [{ "width": 1.5 }, { "width": 4.5 }],
    "header-rows": 1,
    "rows": [
      { "cells": [
        { "text": "Column 1", "bold": true, "shading": "1A3C5E", "color": "FFFFFF" },
        { "text": "Column 2", "bold": true, "shading": "1A3C5E", "color": "FFFFFF" }
      ]},
      { "cells": [{ "text": "Row 1" }, { "text": "Value" }] },
      { "cells": [{ "text": "Row 2", "shading": "F5F7FA" }, { "text": "Value", "shading": "F5F7FA" }] },
      { "cells": [{ "text": "Row 3" }, { "text": "Value" }] },
      { "cells": [{ "text": "Row 4", "shading": "F5F7FA" }, { "text": "Value", "shading": "F5F7FA" }] }
    ]
  }
}
```

---

## Tracked Changes

**Inline:**
```json
{ "paragraph": { "runs": [
  { "text": "The term is " },
  { "delete": "30", "author": "Claude", "date": "2025-10-01T00:00:00Z" },
  { "insert": "60", "author": "Claude", "date": "2025-10-01T00:00:00Z" },
  { "text": " days." }
]}}
```

**Block-level:**
```json
{ "paragraph": { "text": "New paragraph.", "inserted": { "author": "Claude", "date": "..." } } }
{ "paragraph": { "text": "Removed paragraph.", "deleted": { "author": "Claude", "date": "..." } } }
```

## Comments

```json
"comments": {
  "c1": { "author": "Claude", "date": "2025-10-15T12:00:00Z", "text": "Is this confirmed?",
    "replies": [{ "author": "Jane", "date": "...", "text": "Yes." }] }
}
```

Anchor in runs: `{ "comment-start": "c1" }` ... `{ "comment-end": "c1" }`

## Footnotes

Reference in text with `[^1]` syntax or explicit `{ "footnote": "1" }` run.

Define in spec root:
```json
"footnotes": { "1": "Source: Annual Report 2024" }
```

---

## Editing Existing Documents

### Step 1: Unpack

```bash
nebo-office docx unpack document.docx -o spec.json --pretty
```

This extracts the DOCX into a JSON spec — the same format used for creation. Images are extracted to an assets directory alongside the JSON. Adjacent runs with identical formatting are automatically merged.

### Step 2: Edit the JSON

Read `spec.json`, modify the body, styles, metadata, headers, footers — any part of the spec. Since it's the same JSON format used for creation, all the same features are available.

Common edits:
- **Change text**: Edit paragraph text directly in the JSON
- **Add sections**: Insert new body elements (headings, paragraphs, tables, etc.)
- **Restyle**: Modify `styles` to change fonts, colors, sizes
- **Add tracked changes**: Use `insert`/`delete` runs with author and date
- **Add comments**: Define in `comments` object, anchor with `comment-start`/`comment-end` runs
- **Reorganize**: Move, reorder, or delete body elements

### Step 3: Create

```bash
nebo-office docx create spec.json -o output.docx --validate
```

The `--validate` flag checks the output DOCX for structural issues.

### Round-Trip Example

```bash
# Unpack existing document
nebo-office docx unpack report.docx -o report.json --pretty

# Edit report.json (agent modifies the JSON)

# Rebuild
nebo-office docx create report.json -o report-updated.docx --validate
```

---

## Critical Rules

- **All dimensions in inches**, font sizes in **points**, spacing in **points**
- **Use `"restart": true`** on numbered lists that should start from 1
- **Cover pages**: Use `section-break` with `"valign": "center"`, not spacing hacks
- **First-page header suppression**: Set `"first": []` in headers
- **Markdown-like text** supports `**bold**`, `*italic*`, `__underline__`, `~~strike~~`, `` `code` ``, `[link](url)`
- **Tables**: Always set `"columns"` with widths for consistent rendering
- **Zebra striping**: Use `"shading": "F5F7FA"` on alternate row cells
- **Custom styles**: Define in `styles.custom`, reference with `"style": "StyleName"` on paragraphs
