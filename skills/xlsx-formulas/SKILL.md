---
name: xlsx-formulas
description: "XLSX formulas: Excel formula cells, named ranges."
license: MIT
---

# xlsx +formulas

> **PREREQUISITE:** Read [`../xlsx/SKILL.md`](../xlsx/SKILL.md) for commands and JSON spec basics.

## Formula Cells

```json
{ "formula": "=SUM(B2:B5)" }
{ "formula": "=SUM(B2:B5)", "format": "$#,##0", "bold": true }
```

Formulas go inside a cell object. They can be combined with any rich cell properties (format, bold, italic, etc.).

## Common Patterns

```json
{ "cells": ["Revenue", 10000000, { "formula": "=B2*1.15" }] }
{ "cells": ["COGS", { "formula": "=-B2*0.6" }, { "formula": "=-C2*0.6" }] }
{ "cells": ["**Gross Profit**", { "formula": "=B2+B3", "bold": true }, { "formula": "=C2+C3", "bold": true }] }
```

## Named Ranges

```json
"named_ranges": [{ "name": "revenue_total", "range": "Revenue!B6" }]
```

Named ranges can be referenced in formulas across sheets.

## See Also

- [xlsx](../xlsx/SKILL.md) — Service overview
