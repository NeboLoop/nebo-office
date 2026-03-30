---
name: xlsx
description: "Use this skill any time a spreadsheet file is the primary input or output. This means any task where the user wants to: open, read, edit, or fix an existing .xlsx, .xlsm, .csv, or .tsv file (e.g., adding columns, computing formulas, formatting, charting, cleaning messy data); create a new spreadsheet from scratch or from other data sources; or convert between tabular file formats. Trigger especially when the user references a spreadsheet file by name or path — even casually (like \"the xlsx in my downloads\") — and wants something done to it or produced from it. Also trigger for cleaning or restructuring messy tabular data files (malformed rows, misplaced headers, junk data) into proper spreadsheets. The deliverable must be a spreadsheet file. Do NOT trigger when the primary deliverable is a Word document, HTML report, standalone Python script, database pipeline, or Google Sheets API integration, even if tabular data is involved."
license: MIT
triggers:
  - xlsx
  - .xlsx
  - excel
  - spreadsheet
  - workbook
  - .csv
  - .tsv
---

# XLSX — Spreadsheet Generation

Generate Excel spreadsheets (.xlsx) from JSON specifications using the `nebo-office` binary. Compiled Rust — no Python dependencies.

## Helper Skills

| Skill | What it covers |
|-------|---------------|
| [`xlsx-formulas`](../xlsx-formulas/SKILL.md) | Formulas, named ranges |
| [`xlsx-formatting`](../xlsx-formatting/SKILL.md) | Rich cell properties, row defaults, column definitions, conditional formatting, number formats |
| [`xlsx-features`](../xlsx-features/SKILL.md) | Freeze panes, merged cells, auto-filter, data validation, print setup |

## Commands

```bash
nebo-office xlsx create spec.json -o output.xlsx [--assets <dir>]
nebo-office xlsx unpack input.xlsx -o spec.json [--assets <dir>] [--pretty]
nebo-office xlsx validate spec.json
```

## JSON Spec Format

```json
{
  "version": 1,
  "metadata": { "title": "Q4 Report", "creator": "Alma Tuck" },
  "styles": { "font": "Calibri", "size": 11 },
  "sheets": [
    {
      "name": "Revenue",
      "columns": [
        { "width": 20 },
        { "width": 15, "format": "$#,##0" }
      ],
      "freeze": { "row": 1 },
      "rows": [
        {
          "cells": ["Region", "Q4 Revenue"],
          "bold": true, "shading": "4472C4", "color": "FFFFFF"
        },
        { "cells": ["North America", 1250000] },
        {
          "cells": [
            "**Total**",
            { "formula": "=SUM(B2:B5)", "format": "$#,##0", "bold": true }
          ]
        }
      ],
      "merged": ["A10:D10"],
      "autofilter": { "range": "A1:D1" }
    }
  ]
}
```

## Cell Types

| Type | Example | Notes |
|------|---------|-------|
| String | `"Hello"` | Supports `**bold**` markdown |
| Number | `1250000` | Numeric, not quoted |
| Boolean | `true` | |
| Formula | `{ "formula": "=SUM(B2:B5)" }` | Excel formula string |
| Rich | `{ "value": 100, "bold": true, "format": "$#,##0" }` | Full formatting control |

## Round-Trip

```bash
nebo-office xlsx unpack existing.xlsx -o spec.json --pretty
# Edit spec.json
nebo-office xlsx create spec.json -o modified.xlsx
```

## Example: Financial Model

```json
{
  "version": 1,
  "metadata": { "title": "Financial Model" },
  "styles": { "font": "Calibri", "size": 11 },
  "sheets": [
    {
      "name": "Assumptions",
      "columns": [{ "width": 30 }, { "width": 15 }, { "width": 20 }],
      "freeze": { "row": 1 },
      "rows": [
        { "cells": ["Assumption", "Value", "Source"], "bold": true, "shading": "4472C4", "color": "FFFFFF" },
        { "cells": ["Revenue Growth", { "value": 0.15, "format": "0.0%" }, "Management"] },
        { "cells": ["Operating Margin", { "value": 0.22, "format": "0.0%" }, "Industry avg"] }
      ]
    },
    {
      "name": "P&L",
      "columns": [{ "width": 25 }, { "width": 15, "format": "$#,##0" }, { "width": 15, "format": "$#,##0" }],
      "freeze": { "row": 1 },
      "rows": [
        { "cells": ["", "2025", "2026"], "bold": true, "shading": "4472C4", "color": "FFFFFF" },
        { "cells": ["Revenue", 10000000, { "formula": "=B2*1.15" }] },
        { "cells": ["COGS", { "formula": "=-B2*0.6" }, { "formula": "=-C2*0.6" }] },
        { "cells": ["**Gross Profit**", { "formula": "=B2+B3", "bold": true }, { "formula": "=C2+C3", "bold": true }] }
      ]
    }
  ]
}
```

## Critical Rules

1. **Use formulas, not calculated values** — let Excel compute
2. **All assumptions in separate cells** — reference with formulas
3. **Number formats on the cell or column** — not in the value itself
4. **Color values are 6-char hex without #** — `"4472C4"` not `"#4472C4"`
5. **Markdown bold (`**text**`) in strings** — renders as bold
