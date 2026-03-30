---
name: xlsx-features
description: "XLSX features: freeze panes, merged cells, auto-filter, data validation, print setup."
license: MIT
triggers:
  - freeze pane
  - merged cell
  - auto-filter
  - autofilter
  - data validation
  - print setup
---

# xlsx +features

> **PREREQUISITE:** Read [`../xlsx/SKILL.md`](../xlsx/SKILL.md) for commands and JSON spec basics.

## Freeze Panes

```json
"freeze": { "row": 1 }
"freeze": { "row": 1, "col": 1 }
```

## Merged Cells

```json
"merged": ["A1:D1", "A10:D10"]
```

## Auto-filter

```json
"autofilter": { "range": "A1:D10" }
```

## Data Validation

```json
"validations": [
  { "range": "C2:C10", "type": "list", "values": ["Yes", "No"] },
  { "range": "D2:D10", "type": "whole", "min": 0, "max": 100 }
]
```

## Print Setup

```json
"print": { "orientation": "landscape", "fit-to-page": true }
```

## See Also

- [xlsx](../xlsx/SKILL.md) — Service overview
