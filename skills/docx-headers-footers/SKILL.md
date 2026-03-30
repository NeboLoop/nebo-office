---
name: docx-headers-footers
description: "DOCX headers and footers: page numbers, fields, first-page suppression."
license: MIT
triggers:
  - page header
  - page footer
  - page number
  - header footer
---

# docx +headers-footers

> **PREREQUISITE:** Read [`../docx/SKILL.md`](../docx/SKILL.md) for commands and JSON spec basics.

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

## Available Fields

```json
{ "field": "page-number" }
{ "field": "total-pages" }
```

## See Also

- [docx](../docx/SKILL.md) — Service overview
- [docx-styles](../docx-styles/SKILL.md) — Run properties and special runs
