# Formatting Rules for Si-Som DOCX

Professional Word document formatting requires precision. Use these units and rules for consistent results.

## Unit Conversion (DXA / Twips)

Word internal units are usually **Twips** (or **DXA**), where:
- **1,440 DXA = 1 inch**
- **1/20th of a point** (72 points = 1 inch)
- In `python-docx`, use `docx.shared.Twips`.

| Physical | DXA / Twips | Points |
|----------|-------------|--------|
| 1 inch   | 1,440       | 72     |
| 0.5 inch | 720         | 36     |
| 1 cm     | 567         | 28.35  |
| 10 pt    | 200         | 10     |
| 12 pt    | 240         | 12     |

## Common Paper Sizes (DXA)

| Paper | Width (DXA) | Height (DXA) | Content Width (1" margins) |
|-------|-------------|--------------|---------------------------|
| **A4** | **11,906** | **16,838** | **9,026** |
| US Letter | 12,240 | 15,840 | 9,360 |

## Typography Entities (Smart Quotes)

When modifying XML or injecting text into documents, prefer these professional characters over plain quotes:

| Character | Entity (XML/HTML) | Unicode |
|-----------|-------------------|---------|
| ‘ (left single) | `&#x2018;` | `\u2018` |
| ’ (right single / apostrophe) | `&#x2019;` | `\u2019` |
| “ (left double) | `&#x201C;` | `\u201C` |
| ” (right double) | `&#x201D;` | `\u201D` |

## Table Best Practices

1. **Dual Widths**: Always set table width AND cell width in DXA/Twips.
2. **Shading**: Use `ShadingType.CLEAR` (XML equivalent: `<w:shd w:val="clear" .../>`) to avoid black backgrounds on some viewers.
3. **Margins**: Set default cell margins: `top: 80`, `bottom: 80`, `left: 120`, `right: 120` (in Twips) for readable padding.
