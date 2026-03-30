---
name: pptx-themes
description: "PPTX themes: color palettes, font settings."
license: MIT
---

# pptx +themes

> **PREREQUISITE:** Read [`../pptx/SKILL.md`](../pptx/SKILL.md) for commands and JSON spec basics.

## Theme Colors

```json
"theme": {
  "colors": {
    "primary": "1F4E79",
    "accent1": "4472C4",
    "accent2": "ED7D31",
    "background": "FFFFFF",
    "text": "333333"
  },
  "font": "Calibri"
}
```

## Color Roles

| Key | Usage |
|-----|-------|
| `primary` | Title backgrounds, key accents |
| `accent1` | Charts, table headers, emphasis |
| `accent2` | Secondary highlights, callouts |
| `background` | Slide background (default white) |
| `text` | Body text color |

## See Also

- [pptx](../pptx/SKILL.md) — Service overview
