---
name: pptx-shapes
description: "PPTX shapes: rectangles, ovals, lines, text boxes, images, backgrounds, transitions."
license: MIT
triggers:
  - slide shape
  - text box
  - slide background
  - slide transition
---

# pptx +shapes

> **PREREQUISITE:** Read [`../pptx/SKILL.md`](../pptx/SKILL.md) for commands and JSON spec basics.

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

## Shape Types

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

## See Also

- [pptx](../pptx/SKILL.md) — Service overview
