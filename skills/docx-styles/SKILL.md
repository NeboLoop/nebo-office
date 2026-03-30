---
name: docx-styles
description: "DOCX styles: fonts, colors, heading styles, run properties, custom styles, mixed formatting with runs."
license: MIT
---

# docx +styles

> **PREREQUISITE:** Read [`../docx/SKILL.md`](../docx/SKILL.md) for commands and JSON spec basics.

## Document Styles

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

## Paragraph with Runs (Mixed Formatting)

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

## Run Properties

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

## Special Runs

```json
{ "tab": true }
{ "break": "line" }
{ "break": "page" }
{ "field": "page-number" }
{ "field": "total-pages" }
{ "footnote": "1" }
```

## See Also

- [docx](../docx/SKILL.md) — Service overview
