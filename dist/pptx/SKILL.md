---
name: pptx
description: "Use this skill any time a .pptx file is involved in any way — as input, output, or both. This includes: creating slide decks, pitch decks, or presentations; reading, parsing, or extracting text from any .pptx file (even if the extracted content will be used elsewhere, like in an email or summary); editing, modifying, or updating existing presentations; combining or splitting slide files; working with templates, layouts, speaker notes, or comments. Trigger whenever the user mentions \"deck,\" \"slides,\" \"presentation,\" or references a .pptx filename, regardless of what they plan to do with the content afterward. If a .pptx file needs to be opened, created, or touched, use this skill."
license: Proprietary. LICENSE.txt has complete terms
---

# PPTX — Presentation Generation

Generate PowerPoint presentations (.pptx) from JSON specifications using the `nebo-office` binary. Compiled Rust — no JavaScript, no PptxGenJS.

## Commands

```bash
nebo-office pptx create spec.json -o output.pptx [--assets <dir>]
nebo-office pptx unpack input.pptx -o spec.json [--assets <dir>] [--pretty]
nebo-office pptx validate spec.json
```

## JSON Spec Format

```json
{
  "version": 1,
  "metadata": { "title": "Q4 Review", "creator": "Alma Tuck" },
  "theme": {
    "colors": {
      "primary": "1F4E79",
      "accent1": "4472C4",
      "accent2": "ED7D31",
      "background": "FFFFFF",
      "text": "333333"
    },
    "font": "Calibri"
  },
  "size": "16:9",
  "slides": [
    {
      "layout": "title",
      "title": "Q4 Business Review",
      "subtitle": "Acme Corporation — January 2026",
      "background": { "color": "1F4E79" },
      "notes": "Welcome everyone."
    },
    {
      "layout": "content",
      "title": "Key Metrics",
      "body": [
        { "bullets": ["Revenue: **$12.5M** (+15%)", "Customers: **1,200** (+8%)"] }
      ]
    }
  ]
}
```

## Slide Layouts

| Layout | Description | Fields |
|--------|-------------|--------|
| `title` | Title slide | `title`, `subtitle` |
| `content` | Title + body content | `title`, `body` |
| `section` | Section header | `title`, `subtitle` |
| `two-column` | Two-column layout | `title`, `left`, `right` |
| `blank` | Empty slide | `shapes` |
| `title-only` | Title, no body | `title`, `shapes` |
| `comparison` | Side-by-side comparison | `title`, `left`, `right` |

## Slide Sizes

```json
"size": "16:9"     // 10" × 5.625" (default)
"size": "16:10"    // 10" × 6.25"
"size": "4:3"      // 10" × 7.5"
"size": { "width": 13.3, "height": 7.5 }  // custom
```

## Body Blocks

Used in `body`, `left`, and `right` arrays:

```json
{ "paragraph": "Some text with **bold**" }
{ "bullets": ["Point 1", "**Bold** point 2", "Point 3"] }
{ "numbered": ["Step 1", "Step 2", "Step 3"] }
{ "table": [["Metric", "Value"], ["Revenue", "$12.5M"]], "header-rows": 1 }
{ "image": "chart.png", "width": 4, "height": 3 }
```

## Shapes

All coordinates in inches (0,0 = top-left):

```json
"shapes": [
  { "type": "rect", "x": 0.5, "y": 0.5, "w": 4, "h": 3, "fill": "4472C4", "opacity": 0.1 },
  { "type": "text", "x": 1, "y": 1, "w": 3, "h": 2, "text": "Hello", "font-size": 24, "color": "1F4E79", "align": "center" },
  { "type": "image", "image": "logo.png", "x": 5, "y": 0.5, "w": 4.5, "h": 4 },
  { "type": "oval", "x": 2, "y": 2, "w": 1, "h": 1, "fill": "ED7D31" },
  { "type": "line", "x": 0, "y": 3, "w": 10, "h": 0, "line-color": "CCCCCC", "line-width": 2 },
  { "type": "rounded-rect", "x": 1, "y": 1, "w": 3, "h": 2, "fill": "E7E6E6", "corner-radius": 0.2 }
]
```

### Shape Types

| Type | Description | Extra Properties |
|------|-------------|-----------------|
| `rect` | Rectangle | `fill`, `opacity`, `line-color`, `line-width` |
| `rounded-rect` | Rounded rectangle | + `corner-radius` |
| `oval` | Ellipse | `fill`, `opacity` |
| `line` | Line connector | `line-color`, `line-width` |
| `text` | Text box | `text`, `font-size`, `color`, `bold`, `align`, `valign` |
| `image` | Image | `image` (filename from assets dir) |

## Backgrounds

```json
{ "color": "1F4E79" }
{ "image": "bg.jpg" }
{ "gradient": { "from": "1F4E79", "to": "4472C4", "angle": 90 } }
```

Dark backgrounds automatically get white text.

## Transitions

```json
"transition": { "type": "fade", "duration": 0.5 }
```

Types: `fade`, `push`, `wipe`

## Speaker Notes

```json
"notes": "Remember to mention the Q4 highlights here."
```

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

## Example: Pitch Deck

```json
{
  "version": 1,
  "metadata": { "title": "Series A Pitch", "creator": "Alma Tuck" },
  "theme": {
    "colors": { "primary": "1F4E79", "accent1": "4472C4", "background": "FFFFFF", "text": "333333" },
    "font": "Calibri"
  },
  "size": "16:9",
  "slides": [
    {
      "layout": "title",
      "title": "Acme Inc.",
      "subtitle": "Series A — $10M Raise",
      "background": { "color": "1F4E79" },
      "notes": "30-second intro"
    },
    {
      "layout": "content",
      "title": "The Problem",
      "body": [
        { "bullets": ["Manual processes waste 20 hrs/week", "Error rates above 15%", "No visibility into pipeline"] }
      ]
    },
    {
      "layout": "content",
      "title": "Traction",
      "body": [
        { "table": [["Metric", "Q3", "Q4"], ["ARR", "$2M", "$3.5M"], ["Customers", "80", "120"], ["NRR", "115%", "125%"]], "header-rows": 1 }
      ]
    },
    {
      "layout": "two-column",
      "title": "Before & After",
      "left": [{ "paragraph": "**Before**" }, { "bullets": ["Manual entry", "3-day turnaround"] }],
      "right": [{ "paragraph": "**After**" }, { "bullets": ["Automated", "Real-time"] }]
    },
    {
      "layout": "content",
      "title": "The Ask",
      "body": [
        { "paragraph": "**$10M Series A** to accelerate growth" },
        { "bullets": ["50% → Engineering", "30% → Sales", "20% → Operations"] }
      ]
    }
  ]
}
```

## Example: Quarterly Review

```json
{
  "version": 1,
  "metadata": { "title": "Q4 2025 Review" },
  "theme": {
    "colors": { "primary": "2D5F2D", "accent1": "4CAF50", "accent2": "FF9800", "background": "FFFFFF", "text": "333333" },
    "font": "Calibri"
  },
  "size": "16:9",
  "slides": [
    {
      "layout": "title",
      "title": "Q4 2025 Business Review",
      "subtitle": "All Hands — January 2026",
      "background": { "gradient": { "from": "2D5F2D", "to": "4CAF50", "angle": 135 } }
    },
    {
      "layout": "content",
      "title": "Highlights",
      "body": [
        { "bullets": ["Revenue: **$12.5M** (+15% YoY)", "New customers: **45**", "NPS: **72** (up from 65)"] }
      ]
    },
    {
      "layout": "blank",
      "shapes": [
        { "type": "rect", "x": 0.3, "y": 0.3, "w": 9.4, "h": 5, "fill": "F5F5F5" },
        { "type": "text", "x": 0.5, "y": 0.5, "w": 9, "h": 1, "text": "Revenue by Region", "font-size": 28, "bold": true, "color": "2D5F2D" }
      ]
    }
  ]
}
```

## Round-Trip

```bash
nebo-office pptx unpack existing.pptx -o spec.json --pretty
# Edit spec.json
nebo-office pptx create spec.json -o modified.pptx
```

## Critical Rules

1. **All coordinates in inches** — `"x": 1` means 1 inch from left edge
2. **Colors are 6-char hex without #** — `"4472C4"` not `"#4472C4"`
3. **Dark backgrounds get auto white text** — based on luminance
4. **Markdown bold in text** — `**text**` renders bold in slides
5. **Images from assets dir** — pass `--assets <dir>` or place next to spec file
6. **Vary slide layouts** — don't use the same layout for every slide
