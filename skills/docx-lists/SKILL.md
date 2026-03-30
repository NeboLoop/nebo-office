---
name: docx-lists
description: "DOCX lists: bullets, numbered lists, nesting, restart numbering."
license: MIT
triggers:
  - bullet list
  - numbered list
  - bulleted
  - list items
---

# docx +lists

> **PREREQUISITE:** Read [`../docx/SKILL.md`](../docx/SKILL.md) for commands and JSON spec basics.

## Bullets

```json
{
  "bullets": [
    "First item",
    "Second item",
    { "text": "Third item", "children": ["Nested A", "Nested B"] }
  ]
}
```

## Numbered Lists

```json
{ "numbered": ["Step 1", "Step 2", "Step 3"] }
{ "numbered": ["New list", "Continues"], "restart": true }
```

Use `"restart": true` to reset numbering for a new list.

## See Also

- [docx](../docx/SKILL.md) — Service overview
