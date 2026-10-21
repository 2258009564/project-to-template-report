# Academic table rules

Avoid decorative or product-dashboard table styling. Do not add colored headers, gradients, cards, accent fills, zebra striping, or ornamental icons unless the supplied reference explicitly requires them and the user asks for fidelity.

## Three-line tables

Use a three-line table when the template or user requires it.

- No background fill in header or body cells.
- No vertical borders.
- No body-row horizontal gridlines.
- One top border, one border below the header row, and one bottom border.
- Match the reference's line weights exactly; if the reference cannot be measured, use a visually stronger top/bottom line and a lighter header separator.
- Repeat the header row when the table spans pages.
- Keep the caption outside the table and adjacent to it.
- Copy the reference's font, size, paragraph spacing, and alignment. Do not assume all three-line-table body text is centered if the reference uses semantic alignment.

## When the user says not to use three-line tables

Use a plain table with:

- no background color or shading anywhere;
- all cell text horizontally centered;
- all cell text vertically centered;
- consistent plain borders copied from the template, or thin black/gray grid borders when the template has no model;
- no decorative header treatment, zebra striping, icons, badges, or callout styling.

The user's explicit request overrides the template's table style.

## Structural safety

- Preserve merged cells, widths, row heights, cell margins, and repeating headers when filling an existing template table.
- Never rebuild a strict template table just to change its text.
- Do not use fixed row heights that clip content; when the template has fixed heights, shorten content or verify Word expansion behavior before changing them.
- Center text vertically through Word cell properties, not by adding blank paragraphs.

## QA

Inspect table XML and rendered pages. Confirm:

- no unexpected `w:shd` fill remains;
- expected borders are present and forbidden borders absent;
- horizontal and vertical alignment match the selected mode;
- text wraps without clipping;
- captions remain attached;
- multi-page tables repeat headers and do not lose their bottom border.
