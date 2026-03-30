---
name: docx
description: "Use this skill whenever the user wants to create, read, edit, or manipulate Word documents (.docx files). Triggers include: any mention of 'Word doc', 'word document', '.docx', or requests to produce professional documents with formatting like tables of contents, headings, page numbers, or letterheads. Also use when extracting or reorganizing content from .docx files. If the user asks for a 'report', 'memo', 'letter', 'template', or similar deliverable as a Word or .docx file, use this skill."
license: Proprietary. LICENSE.txt has complete terms
---

# DOCX — Document Generation & Manipulation

Generate and manipulate Word documents (.docx) from JSON specifications using the `nebo-office` binary. Compiled Rust — no JavaScript or Python dependencies.

## Commands

```bash
nebo-office docx create spec.json -o output.docx [--assets <dir>]
nebo-office docx unpack input.docx -o spec.json [--assets <dir>] [--pretty]
nebo-office docx validate spec.json
```

## JSON Spec Format

```json
{
  "version": 1,
  "metadata": { "title": "Report Title", "creator": "Alma Tuck" },
  "page": {
    "size": "letter",
    "orientation": "portrait",
    "margin": { "top": 1, "bottom": 1, "left": 1, "right": 1 }
  },
  "styles": {
    "font": "Arial",
    "size": 12,
    "color": "333333"
  },
  "body": [
    { "heading": 1, "text": "Main Title" },
    { "paragraph": "Regular text with **bold** and *italic* support." },
    { "table": [["Item", "Amount"], ["Service", "$500"]], "header-rows": 1 },
    { "image": "logo.png", "width": 2, "height": 1 }
  ]
}
```

## Page Sizes

| Size | Dimensions |
|------|-----------|
| `letter` | 8.5" x 11" — default |
| `a4` | 210mm x 297mm |
| `legal` | 8.5" x 14" |

Custom: `"size": { "width": 8.5, "height": 11 }`

## Margins

All margins in inches (default: 1 inch each):

```json
"margin": { "top": 1, "bottom": 1, "left": 1, "right": 1 }
```

Or a single number for uniform margins: `"margin": 1`

## Block Types

### Heading
```json
{ "heading": 1, "text": "Main Title" }
{ "heading": 2, "text": "Section" }
{ "heading": 3, "text": "Subsection", "id": "bookmark-id" }
```

Levels 1-6 supported. Optional `id` creates a bookmark anchor.

### Paragraph

Simple text with inline markdown:
```json
{ "paragraph": "Text with **bold**, *italic*, __underline__, ~~strike~~, `code`, and [links](https://example.com)." }
```

Full paragraph with formatting:
```json
{
  "paragraph": {
    "text": "Aligned and spaced text",
    "align": "center",
    "spacing": { "before": 12, "after": 6 },
    "indent": { "left": 0.5 }
  }
}
```

Paragraph with runs (mixed formatting):
```json
{
  "paragraph": {
    "runs": [
      { "text": "Normal text " },
      { "text": "bold red", "bold": true, "color": "DC2626" },
      { "tab": true },
      { "text": "after tab" }
    ]
  }
}
```

### Run Properties

| Property | Type | Description |
|----------|------|-------------|
| `text` | string | Text content |
| `bold` | bool | Bold text |
| `italic` | bool | Italic text |
| `underline` | bool | Underlined text |
| `strike` | bool | Strikethrough |
| `font` | string | Font name |
| `size` | number | Font size in points |
| `color` | string | Font color (hex, `"FF0000"`) |
| `highlight` | string | Highlight color |
| `link` | string | Hyperlink URL |
| `superscript` | bool | Superscript |
| `subscript` | bool | Subscript |
| `all-caps` | bool | All capitals |
| `small-caps` | bool | Small capitals |

### Special Runs

```json
{ "tab": true }
{ "break": "line" }
{ "break": "page" }
{ "field": "page-number" }
{ "field": "total-pages" }
{ "footnote": "1" }
```

### Bullets
```json
{
  "bullets": [
    "First item",
    "Second item",
    { "text": "Third item", "children": ["Nested A", "Nested B"] }
  ]
}
```

### Numbered Lists
```json
{ "numbered": ["Step 1", "Step 2", "Step 3"] }
{ "numbered": ["New list", "Continues"], "restart": true }
```

### Table

Simple form:
```json
{
  "table": [
    ["Item", "Qty", "Price"],
    ["Widget A", "10", "$25"],
    ["Widget B", "5", "$50"]
  ],
  "header-rows": 1
}
```

Full form with formatting:
```json
{
  "table": {
    "columns": [{ "width": 3 }, { "width": 2 }, { "width": 2 }],
    "header-rows": 1,
    "rows": [
      {
        "cells": [
          { "text": "Header", "bold": true, "shading": "4472C4", "color": "FFFFFF" },
          { "text": "Qty", "bold": true, "shading": "4472C4", "color": "FFFFFF", "align": "center" },
          { "text": "Price", "bold": true, "shading": "4472C4", "color": "FFFFFF", "align": "right" }
        ]
      },
      {
        "cells": ["Widget A", "10", "$25"]
      }
    ]
  }
}
```

### Cell Properties

| Property | Type | Description |
|----------|------|-------------|
| `text` | string | Cell text (supports **bold** markdown) |
| `bold` | bool | Bold text |
| `color` | string | Text color (hex) |
| `shading` | string | Background color (hex) |
| `align` | string | `"left"`, `"center"`, `"right"` |
| `valign` | string | `"top"`, `"center"`, `"bottom"` |
| `colspan` | number | Column span |
| `rowspan` | number | Row span |

### Image
```json
{ "image": "photo.png", "width": 4, "height": 3 }
{ "image": "logo.png", "width": 2, "height": 1, "align": "center", "caption": "Figure 1" }
```

- Width/height in inches
- Images loaded from `--assets` directory
- Supported: png, jpg, jpeg, gif, bmp

### Page Break
```json
{ "page-break": true }
```

### Section Break
```json
{
  "section-break": {
    "type": "next-page",
    "columns": 2,
    "column-gap": 0.5,
    "page": { "orientation": "landscape" }
  }
}
```

Types: `next-page`, `continuous`, `even-page`, `odd-page`

### Table of Contents
```json
{ "toc": true }
{ "toc": { "title": "Contents", "depth": 3 } }
```

## Headers and Footers

```json
{
  "headers": {
    "default": [{ "paragraph": { "runs": [{ "text": "Company Name" }, { "tab": true }, { "field": "page-number" }] } }],
    "first": []
  },
  "footers": {
    "default": [{ "paragraph": { "runs": [{ "text": "Page " }, { "field": "page-number" }, { "text": " of " }, { "field": "total-pages" }], "align": "center" } }]
  }
}
```

Empty array `[]` suppresses header/footer on that page type (first, even).

## Styles

```json
"styles": {
  "font": "Arial",
  "size": 12,
  "color": "333333",
  "headings": {
    "font": "Arial",
    "color": "1A3C5E",
    "h1": { "size": 28, "bold": true },
    "h2": { "size": 22, "bold": true },
    "h3": { "size": 16, "bold": true, "italic": true }
  },
  "custom": {
    "quote": {
      "italic": true,
      "indent": { "left": 0.5, "right": 0.5 },
      "color": "666666"
    }
  }
}
```

## Comments

```json
{
  "comments": {
    "c1": {
      "author": "Claude",
      "date": "2026-01-15T12:00:00Z",
      "text": "Review this section",
      "replies": [
        { "author": "Alma", "date": "2026-01-16T09:00:00Z", "text": "Looks good" }
      ]
    }
  }
}
```

Reference in text with runs: `{ "comment-start": "c1" }` and `{ "comment-end": "c1" }`

## Footnotes

```json
{
  "footnotes": {
    "1": "Source: Annual Report 2025",
    "2": "See appendix for methodology"
  }
}
```

Reference in text: `"Revenue grew 15%[^1] using adjusted metrics[^2]."` or with runs: `{ "footnote": "1" }`

## Tracked Changes

```json
{
  "paragraph": {
    "runs": [
      { "text": "The term is " },
      { "delete": "30", "author": "Claude", "date": "2026-01-15T00:00:00Z" },
      { "insert": "60", "author": "Claude", "date": "2026-01-15T00:00:00Z" },
      { "text": " days." }
    ]
  }
}
```

## Metadata

```json
"metadata": {
  "title": "Document Title",
  "subject": "Subject Line",
  "creator": "Author Name",
  "description": "Document description",
  "keywords": ["keyword1", "keyword2"],
  "category": "Reports"
}
```

## Example: Business Report

```json
{
  "version": 1,
  "metadata": { "title": "Q4 Performance Report", "creator": "Acme Corp" },
  "page": { "size": "letter", "margin": { "top": 1, "bottom": 1, "left": 1, "right": 1 } },
  "styles": {
    "font": "Arial",
    "size": 11,
    "headings": { "color": "1A3C5E", "h1": { "size": 28, "bold": true }, "h2": { "size": 20, "bold": true } }
  },
  "headers": {
    "default": [{ "paragraph": { "runs": [{ "text": "Acme Corp — Q4 Report", "italic": true, "color": "999999" }] } }]
  },
  "footers": {
    "default": [{ "paragraph": { "runs": [{ "text": "Page " }, { "field": "page-number" }], "align": "center" } }]
  },
  "body": [
    { "heading": 1, "text": "Q4 Performance Report" },
    { "paragraph": "Prepared by Acme Corp — January 2026" },
    { "heading": 2, "text": "Executive Summary" },
    { "paragraph": "Revenue grew **15%** year-over-year to **$12.5M**, driven by strong growth in the Asia Pacific region." },
    { "heading": 2, "text": "Revenue by Region" },
    {
      "table": {
        "columns": [{ "width": 2.5 }, { "width": 2 }, { "width": 1.5 }],
        "header-rows": 1,
        "rows": [
          { "cells": [
            { "text": "Region", "bold": true, "shading": "1A3C5E", "color": "FFFFFF" },
            { "text": "Revenue", "bold": true, "shading": "1A3C5E", "color": "FFFFFF", "align": "right" },
            { "text": "Growth", "bold": true, "shading": "1A3C5E", "color": "FFFFFF", "align": "right" }
          ]},
          { "cells": ["North America", "$5.2M", "+12%"] },
          { "cells": ["Europe", "$3.1M", "+8%"] },
          { "cells": ["Asia Pacific", "$2.2M", "+22%"] },
          { "cells": [{ "text": "**Total**", "bold": true }, "$10.5M", "+14%"] }
        ]
      }
    },
    { "heading": 2, "text": "Key Achievements" },
    { "bullets": ["Launched 3 new products", "Expanded to 5 new markets", "Reduced churn to **1.8%**"] },
    { "heading": 2, "text": "Next Steps" },
    { "numbered": ["Finalize 2026 budget", "Hire 20 new engineers", "Launch enterprise tier"] }
  ]
}
```

## Example: Letter

```json
{
  "version": 1,
  "page": { "size": "letter", "margin": { "top": 1.5, "bottom": 1, "left": 1, "right": 1 } },
  "styles": { "font": "Times-Roman", "size": 12 },
  "body": [
    { "paragraph": { "text": "January 15, 2026", "spacing": { "after": 24 } } },
    { "paragraph": "Dear Mr. Smith," },
    { "paragraph": "" },
    { "paragraph": "Thank you for your interest in our services. We are pleased to offer the following proposal for your consideration." },
    { "paragraph": "" },
    { "paragraph": "We look forward to hearing from you." },
    { "paragraph": "" },
    { "paragraph": "Sincerely," },
    { "paragraph": "" },
    { "paragraph": { "runs": [{ "text": "Alma Tuck", "bold": true }] } },
    { "paragraph": "Director of Operations" }
  ]
}
```

## Round-Trip

```bash
nebo-office docx unpack existing.docx -o spec.json --pretty
# Edit spec.json
nebo-office docx create spec.json -o modified.docx
```

## Critical Rules

1. **All dimensions in inches** — margins, image sizes, column widths, indentation
2. **Font sizes in points** — `"size": 12` means 12pt
3. **Colors are 6-char hex without #** — `"1A3C5E"` not `"#1A3C5E"`
4. **Markdown in paragraph strings** — `**bold**`, `*italic*`, `__underline__`, `~~strike~~`, `` `code` ``, `[text](url)`
5. **Use runs for mixed formatting** — when you need different styles in one paragraph
6. **Images from assets dir** — pass `--assets <dir>` or place next to spec file
7. **Tables need header-rows** — set `"header-rows": 1` for proper header styling
