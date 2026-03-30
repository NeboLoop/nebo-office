---
name: xlsx-formatting
description: "XLSX formatting: rich cell properties, row defaults, column definitions, conditional formatting, number formats."
license: MIT
triggers:
  - cell format
  - conditional format
  - number format
  - column width
---

# xlsx +formatting

> **PREREQUISITE:** Read [`../xlsx/SKILL.md`](../xlsx/SKILL.md) for commands and JSON spec basics.

## Rich Cell Properties

| Property | Type | Description |
|----------|------|-------------|
| `value` | any | Cell value (string, number, boolean) |
| `formula` | string | Excel formula (overrides value) |
| `format` | string | Number format (`"$#,##0"`, `"0.00%"`) |
| `bold` | bool | Bold text |
| `italic` | bool | Italic text |
| `underline` | bool | Underlined text |
| `font` | string | Font name |
| `size` | number | Font size in points |
| `color` | string | Font color (hex, `"FF0000"`) |
| `shading` | string | Background color (hex) |
| `align` | string | `"left"`, `"center"`, `"right"` |
| `valign` | string | `"top"`, `"center"`, `"bottom"` |
| `wrap` | bool | Wrap text in cell |

## Row-Level Defaults

Properties on a row apply to all cells unless overridden:
```json
{ "cells": ["A", "B"], "bold": true, "shading": "4472C4", "color": "FFFFFF", "height": 20 }
```

## Column Definitions

```json
"columns": [
  { "width": 20 },
  { "width": 15, "format": "$#,##0" },
  { "width": 10, "hidden": true }
]
```

## Conditional Formatting

```json
"conditional": [
  { "range": "B2:B10", "rule": "greater-than", "value": 1000000, "style": { "color": "00B050" } }
]
```
Rules: `greater-than`, `less-than`, `equal`, `not-equal`, `between`

## Number Formats

| Format | Example | Description |
|--------|---------|-------------|
| `$#,##0` | $1,250,000 | Currency |
| `$#,##0.00` | $1,250,000.00 | Currency + cents |
| `0.00%` | 15.00% | Percentage |
| `#,##0` | 1,250,000 | Thousands separator |
| `mm-dd-yy` | 01-15-26 | Date |

## See Also

- [xlsx](../xlsx/SKILL.md) — Service overview
