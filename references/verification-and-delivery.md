# Verification and delivery

## Structural checks

- Original template hash is unchanged.
- Output opens in Word without repair prompts.
- Required and protected text checks both pass.
- Section, table, caption, image, heading, bookmark, content-control, and field counts match the content map.
- No placeholder tokens, internal notes, tool citations, temporary paths, or stale version text remain.
- Internal Drawio and PNG working sets contain one matching pair per inserted diagram when editable diagram sources are part of the workflow.

Run:

```powershell
python scripts/extract_docx_contract.py template.docx --out template_contract.json
python scripts/extract_docx_contract.py output.docx --out output_contract.json
python scripts/compare_docx_contracts.py template_contract.json output_contract.json --out contract_diff.json
```

Review every reported drift. Some content counts change when sections are filled; unexplained page setup, style, header/footer, table geometry, or protected-text drift is not acceptable.

## Console screenshot checks

When the report includes command-line evidence:

- Every requested result has a capture; no monospaced text block stands in for a run.
- Each capture shows the prompt, the executed command, the input and the program output.
- Captures are inline at a readable template-compatible width (often 10-11 cm), directly below their result caption, and the caption keeps with the image.
- No capture is blank, black, cropped mid-line or split across pages.
- The report names the terminal and the shell that were actually used.
- Capture completion and exit status passed; old output after a failed capture is not reused. Raw screenshots remain available for checking crop boundaries.

## Word checks

- Create a separate hidden Word instance.
- Update each TOC and every field, call `Repaginate`, save, and record `ComputeStatistics(wdStatisticPages)`.
- Check page numbers, caption numbering, cross-references, and the TOC after the update.
- Close only the instance and document created by the task.

## Render checks

This repository does not bundle a renderer. If the Codex documents skill supplies `render_docx.py`, use it without `--emit_pdf` unless a PDF is requested. Otherwise export an internal PDF with the isolated Word instance and render each page using an available PDF renderer. Keep that preview internal unless requested. If neither route is available, report the visual verification gap instead of claiming layout passed. Inspect every final page at readable resolution; Word remains the pagination authority.

Use a new render directory for each layout-changing repair and maintain a concise defect ledger. Inspect every final page; after local changes, focus intermediate reviews on changed pages and spillover, reusing unchanged structural evidence. Do not repeat source research or template extraction. A contact sheet locates issues but does not replace readable page inspection.

Check:

- cover and formal metadata;
- heading hierarchy and spacing;
- body alignment, indentation, and line spacing;
- table widths, merges, borders, row breaks, repeated headers, shading, and alignment;
- complete figures, connector readability, caption adjacency, and image sharpness;
- headers, footers, page numbers, TOC, references, and final page;
- blank or nearly blank pages that were not present in the reference;
- clipping, overlap, missing glyphs, shifted shapes, orphans, and large unexplained gaps.

A GBK `UnicodeDecodeError` from the renderer's output reader does not prove failure. If page PNGs were produced, inspect them before deciding.

## Comparison strategy

Compare reference and output in three layers:

1. Contract comparison for styles, sections, headers/footers, tables, and object types.
2. Side-by-side page review for visual system and layout behavior.
3. Pixel or perceptual diff for pages expected to be unchanged, such as cover, approval, or signature pages.

## Delivery

Deliver only the final artifacts the user requested, to the requested path. Inserting diagrams into a DOCX does not by itself authorize delivery of `.drawio`, exported PNGs, render previews, a figure manifest, VBS, contracts, or QA logs; keep those working artifacts internal unless the user explicitly requests them.
