---
name: docx-tables
description: "DOCX tables: simple and full-form tables, cell properties, column widths, colspan, rowspan, header rows."
license: MIT
---

# docx +tables

> **PREREQUISITE:** Read [`../docx/SKILL.md`](../docx/SKILL.md) for commands and JSON spec basics.

## Simple Form

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

## Full Form with Formatting

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

## Cell Properties

| Property | Type | Description |
|----------|------|-------------|
| `text` | string | Cell text (supports `**bold**` markdown) |
| `bold` | bool | Bold text |
| `color` | string | Text color (hex) |
| `shading` | string | Background color (hex) |
| `align` | string | `"left"`, `"center"`, `"right"` |
| `valign` | string | `"top"`, `"center"`, `"bottom"` |
| `colspan` | number | Column span |
| `rowspan` | number | Row span |

## See Also

- [docx](../docx/SKILL.md) — Service overview
