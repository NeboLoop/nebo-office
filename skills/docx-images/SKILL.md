---
name: docx-images
description: "DOCX images: embedding, sizing, alignment, captions."
license: MIT
---

# docx +images

> **PREREQUISITE:** Read [`../docx/SKILL.md`](../docx/SKILL.md) for commands and JSON spec basics.

## Image Block

```json
{ "image": "photo.png", "width": 4, "height": 3 }
{ "image": "logo.png", "width": 2, "height": 1, "align": "center", "caption": "Figure 1" }
```

- Width/height in inches
- Images loaded from `--assets` directory
- Supported formats: png, jpg, jpeg, gif, bmp

## Usage

Pass `--assets <dir>` to specify the directory containing images, or place them next to the spec file.

```bash
nebo-office docx create spec.json -o output.docx --assets ./images
```

## See Also

- [docx](../docx/SKILL.md) — Service overview
