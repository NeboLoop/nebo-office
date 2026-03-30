---
name: pdf
description: Use this skill whenever the user wants to do anything with PDF files. This includes reading or extracting text/tables from PDFs, combining or merging multiple PDFs into one, splitting PDFs apart, rotating pages, adding watermarks, creating new PDFs, filling PDF forms, encrypting/decrypting PDFs, extracting images, and OCR on scanned PDFs to make them searchable. If the user mentions a .pdf file or asks to produce one, use this skill.
license: Proprietary. LICENSE.txt has complete terms
---

# PDF — Document Generation & Manipulation

Generate and manipulate PDF files from JSON specifications using the `nebo-pdf` binary. Compiled Rust with `printpdf` (create) and `lopdf` (fill/extract).

## Commands

```bash
nebo-pdf create spec.json -o output.pdf [--assets <dir>]
nebo-pdf fill input.pdf fields.json -o output.pdf
nebo-pdf extract input.pdf -o spec.json [--pretty]
nebo-pdf validate spec.json
```

## JSON Spec Format

```json
{
  "version": 1,
  "metadata": { "title": "Invoice #1234", "creator": "Alma Tuck" },
  "page": {
    "size": "letter",
    "margin": { "top": 1, "bottom": 1, "left": 1, "right": 1 }
  },
  "styles": {
    "font": "Helvetica",
    "size": 12,
    "color": "000000"
  },
  "pages": [
    {
      "body": [
        { "heading": 1, "text": "Invoice" },
        { "paragraph": "Date: January 15, 2026" },
        { "table": [["Item", "Amount"], ["Service", "$500"]], "header-rows": 1 },
        { "image": "logo.png", "x": 5, "y": 0, "width": 2, "height": 1 }
      ]
    }
  ]
}
```

## Page Sizes

| Size | Dimensions (points) |
|------|-------------------|
| `letter` | 612 × 792 (8.5" × 11") — default |
| `a4` | 595 × 842 (210mm × 297mm) |
| `legal` | 612 × 1008 (8.5" × 14") |

## Margins

All margins in inches (default: 1 inch each):

```json
"margin": { "top": 1, "bottom": 1, "left": 1, "right": 1 }
```

## Block Types

### Heading
```json
{ "heading": 1, "text": "Main Title" }
{ "heading": 2, "text": "Section" }
{ "heading": 3, "text": "Subsection" }
```

Sizes: h1 = 2× base, h2 = 1.5×, h3 = 1.25×, h4 = 1.1×

### Paragraph
```json
{ "paragraph": "Regular text that wraps automatically within the page margins." }
```

### Table
```json
{
  "table": [
    ["Item", "Qty", "Price", "Total"],
    ["Widget A", "10", "$25", "$250"],
    ["Widget B", "5", "$50", "$250"]
  ],
  "header-rows": 1
}
```

Header rows render in bold. Table draws grid lines automatically.

### Image
```json
{ "image": "logo.png", "width": 2, "height": 1 }
{ "image": "chart.png", "x": 5, "y": 2, "width": 3, "height": 2 }
```

- Width/height in inches
- Optional `x`, `y` for absolute positioning (inches from left/top)
- Images loaded from `--assets` directory

## Positioned Elements

Omit `x`/`y` for flow mode (content stacks top-to-bottom). Include them for absolute positioning:

```json
{ "paragraph": "This flows normally" },
{ "paragraph": "This is positioned", "x": 5, "y": 2 }
```

## Watermark

```json
"watermark": {
  "text": "DRAFT",
  "color": "CCCCCC",
  "rotate": 45,
  "size": 72
}
```

Applied to every page.

## Built-in Fonts

Available without embedding:
- **Helvetica** (default) — clean sans-serif
- **Times-Roman** — classic serif
- **Courier** — monospace

## Form Filling

Fill an existing PDF form with values:

```bash
nebo-pdf fill template.pdf fields.json -o filled.pdf
```

`fields.json`:
```json
{
  "name": "John Doe",
  "date": "2026-01-15",
  "amount": "500.00",
  "approved": true
}
```

Keys match PDF form field names (the `T` attribute in AcroForm fields).

## Text Extraction

Extract text from an existing PDF to JSON:

```bash
nebo-pdf extract document.pdf -o output.json --pretty
```

Returns a PdfSpec with paragraphs extracted per page.

## Example: Invoice

```json
{
  "version": 1,
  "metadata": { "title": "Invoice #1234", "creator": "Acme Corp" },
  "page": { "size": "letter", "margin": { "top": 1, "bottom": 1, "left": 1, "right": 1 } },
  "styles": { "font": "Helvetica", "size": 11 },
  "pages": [
    {
      "body": [
        { "heading": 1, "text": "INVOICE" },
        { "paragraph": "Invoice #: 1234" },
        { "paragraph": "Date: January 15, 2026" },
        { "paragraph": "Due: February 15, 2026" },
        { "paragraph": "" },
        { "heading": 3, "text": "Bill To" },
        { "paragraph": "Jane Smith" },
        { "paragraph": "456 Oak Ave, Springfield, IL 62701" },
        { "paragraph": "" },
        {
          "table": [
            ["Description", "Qty", "Rate", "Amount"],
            ["Consulting Services", "40 hrs", "$150/hr", "$6,000"],
            ["Travel Expenses", "1", "$500", "$500"],
            ["Software License", "1", "$1,200", "$1,200"]
          ],
          "header-rows": 1
        },
        { "paragraph": "" },
        { "heading": 3, "text": "Total: $7,700" },
        { "paragraph": "" },
        { "paragraph": "Payment Terms: Net 30" },
        { "paragraph": "Please remit payment to Acme Corp, 123 Main St, Chicago, IL 60601" }
      ]
    }
  ]
}
```

## Example: Simple Report

```json
{
  "version": 1,
  "metadata": { "title": "Monthly Report", "creator": "Alma Tuck" },
  "page": { "size": "letter" },
  "styles": { "font": "Helvetica", "size": 12 },
  "watermark": { "text": "CONFIDENTIAL", "color": "DDDDDD", "size": 60 },
  "pages": [
    {
      "body": [
        { "heading": 1, "text": "Monthly Performance Report" },
        { "heading": 2, "text": "January 2026" },
        { "paragraph": "This report summarizes key metrics for the month of January." },
        { "heading": 3, "text": "Revenue Summary" },
        {
          "table": [
            ["Region", "Revenue", "Growth"],
            ["North America", "$5.2M", "+12%"],
            ["Europe", "$3.1M", "+8%"],
            ["Asia Pacific", "$2.2M", "+22%"],
            ["Total", "$10.5M", "+14%"]
          ],
          "header-rows": 1
        },
        { "paragraph": "North America continues to lead in absolute revenue while Asia Pacific shows the strongest growth trajectory." }
      ]
    },
    {
      "body": [
        { "heading": 2, "text": "Customer Metrics" },
        {
          "table": [
            ["Metric", "December", "January", "Change"],
            ["Active Customers", "1,150", "1,200", "+50"],
            ["NPS Score", "68", "72", "+4"],
            ["Churn Rate", "2.1%", "1.8%", "-0.3%"]
          ],
          "header-rows": 1
        },
        { "paragraph": "Customer satisfaction improved across all segments." }
      ]
    }
  ]
}
```

## Critical Rules

1. **Page-oriented** — content is organized by page, not auto-paginated
2. **Inches for positioning** — all x/y/width/height values are in inches
3. **Flow mode by default** — omit x/y and content stacks top-to-bottom
4. **Positioned elements use x/y from top-left** — (0,0) is top-left corner
5. **Colors are 6-char hex** — `"000000"` for black, `"FF0000"` for red
6. **Form filling preserves structure** — only fills form fields, doesn't modify layout
