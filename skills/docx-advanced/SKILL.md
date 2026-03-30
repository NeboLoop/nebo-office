---
name: docx-advanced
description: "DOCX advanced features: table of contents, comments, tracked changes, footnotes, section breaks."
license: MIT
---

# docx +advanced

> **PREREQUISITE:** Read [`../docx/SKILL.md`](../docx/SKILL.md) for commands and JSON spec basics.

## Table of Contents

```json
{ "toc": true }
{ "toc": { "title": "Contents", "depth": 3 } }
```

## Section Break

```json
{
  "section-break": {
    "type": "next-page",
    "columns": 2,
    "column-gap": 0.5,
    "page": { "orientation": "landscape" }
  }
}
```

Types: `next-page`, `continuous`, `even-page`, `odd-page`

## Comments

```json
{
  "comments": {
    "c1": {
      "author": "Claude",
      "date": "2026-01-15T12:00:00Z",
      "text": "Review this section",
      "replies": [
        { "author": "Alma", "date": "2026-01-16T09:00:00Z", "text": "Looks good" }
      ]
    }
  }
}
```

Reference in text with runs: `{ "comment-start": "c1" }` and `{ "comment-end": "c1" }`

## Footnotes

```json
{
  "footnotes": {
    "1": "Source: Annual Report 2025",
    "2": "See appendix for methodology"
  }
}
```

Reference in text: `"Revenue grew 15%[^1] using adjusted metrics[^2]."` or with runs: `{ "footnote": "1" }`

## Tracked Changes

```json
{
  "paragraph": {
    "runs": [
      { "text": "The term is " },
      { "delete": "30", "author": "Claude", "date": "2026-01-15T00:00:00Z" },
      { "insert": "60", "author": "Claude", "date": "2026-01-15T00:00:00Z" },
      { "text": " days." }
    ]
  }
}
```

## See Also

- [docx](../docx/SKILL.md) — Service overview
- [docx-styles](../docx-styles/SKILL.md) — Run properties
