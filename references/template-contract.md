# DOCX template contract

The reference DOCX is an immutable layout authority. Copy it before editing.

## Inventory scope

Record these values in `template_contract.json` and confirm important ones visually:

- Original absolute path, byte size, SHA-256, and modification time.
- Section count, page size, orientation, margins, columns, header/footer distance, and section breaks.
- Header/footer text, page-number fields, first-page and odd/even settings.
- Built-in and custom paragraph styles: local name, base style, next style, outline level, font families, size, weight, color, alignment, indentation, spacing, line spacing, pagination controls, numbering, and borders.
- Representative direct formatting for Title, every heading level, body, captions, references, lists, table text, cover metadata, and appendices.
- Heading and caption numbering method: literal text, list template, `SEQ`, `STYLEREF`, `REF`, or other fields.
- Table count, row/column dimensions, widths, alignment, cell margins, merges, repeating headers, borders, shading, fixed heights, and vertical alignment.
- Inline and floating images, physical size, wrap mode, anchor paragraph, and caption adjacency.
- Content controls, bookmarks, form fields, TOC, cross-references, and protected regions.
- Sample paragraphs for every material style, including effective formatting rather than style definitions alone.

Use `scripts/extract_docx_contract.py` for the first pass. Word COM inspection and rendered pages remain authoritative for properties that `python-docx` cannot resolve reliably.

## Edit map

Before authoring, classify nodes as:

- `replace`: exact placeholder or source content that will be replaced.
- `insert`: a defined insertion point with inherited style.
- `preserve`: layout/content that must remain unchanged.
- `blank`: formal fields intentionally left empty.
- `generated`: figures, captions, TOC, fields, and cross-references created by automation.

Prefer bookmarks, content controls, exact placeholder text, and allowlisted paragraph/table nodes. Paragraph indices are a last resort and must be verified immediately before the VBS run.

## Fidelity rules

- Preserve sections, tables, merged cells, row heights, column widths, page settings, headers, footers, signature areas, and protected text.
- Apply both style definitions and effective paragraph/run formatting when the template relies on direct formatting.
- Do not clear all direct formatting globally; it can destroy intentional emphasis, field formatting, and cover design.
- Avoid global search-and-replace for short or repeated tokens.
- Inline figures are preferred. Floating shapes require renderer-specific inspection.
- If content length conflicts with the template, rewrite or redistribute content inside the existing structure before altering layout.
- "Same as the reference" means the same formatting system and layout behavior. Content-driven pagination may differ, but unexplained style or spacing drift is a failure.
