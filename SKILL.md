---
name: project-to-template-report
description: Create or update Word reports from a supplied template and factual materials, or transfer template formatting to structured content. Extract effective formatting, preserve the cover, batch author with Word COM, and verify rendered pages. Project evidence and UML diagrams are conditional modules.
---

# Template report with reusable Word automation

## Execution

Follow the current user's execution and delegation preferences, including locally stored preferences when available.

Separate content from formatting. Copy the reference; preserve its cover, sections, headers, footers and objects outside authorized edit ranges. Resolve required personal information through the mandatory identity gate below; other unknown formal metadata must not be invented. User updates override older career plans and drafts.

## Scope

- Fill/report: inspect template and supplied facts, author structured blocks, batch format.
- Format transfer: inspect reference and target; preserve target content and map only authorized roles. Do not generate a replacement report. The bundled build engine supports tail replacement, not arbitrary format transfer; use a bounded COM adapter for the latter.
- Software-project report: additionally read [project evidence](references/project-evidence-and-writing.md); inspect relevant code and tests, not every file by default.
- Diagrams requested/required: read [figures](references/drawio-and-figure-qa.md). Create only useful figures. UML completeness rules apply only to relevant assignments.
- Console run evidence requested/required: read [console screenshots](references/console-screenshots.md). Capture a real terminal run; monospaced text pasted from a console is not evidence.

Do not load all references or infer project/diagram requirements from this skill's name.

## One request, template-driven completion

A request such as “根据我的 Word 制作 skill，帮我完成以下报告” authorizes the full report workflow within the supplied template's edit boundaries. Do not require separate instructions to draft sections, generate required figures/tables, review prose or verify pages.

Read the template as both a layout reference and a task specification. Extract requirements from headings, instructional text, captions, figure/table placeholders and supplied assignment directions. Map each requirement to a content block and its verification in the task-local plan. A cue such as “此处应有流程图” or “插入实验结果表” is an actionable requirement, not text to preserve in the finished report. Generate the supported content, insert it with its caption and explanation, and verify it; do not stop at describing what the user should draw or fill in. Read figure/table references when the template requires them even if the user's message never mentions diagrams or tables.

When the template lacks explicit figure/table cues, choose useful figures and tables where they clarify the requested content and the allowed template structure permits them. Avoid decorative additions and preserve protected areas. The user may also request an additional figure/table without having to redesign the workflow.

Use the template topic, assignment instructions, relevant local records and user-authorized materials already available before requesting more input. Draft general explanation, analysis and evidence-supported diagrams autonomously when the task is sufficiently specified. A template alone does not establish real implementation, measurements, personal experience or test results: do not invent these. Mark illustrative examples as such when the assignment permits them; resolve only genuinely blocking missing facts with a focused question, and continue independent sections. Missing required evidence is an explicit gap, not a completed figure/table or a passed report. Stop after a reviewed, verified report is delivered or the remaining concrete gap is reported.

## Local reusable information

At every task start, load relevant records using [local memory and feedback learning](references/local-learning.md). Keep identity data and personalized rules outside this skill directory. Apply current instructions before older memory and avoid asking again for consistent known fields.

Before asking for name, student ID, class, university or major, read [local profile lookup](references/local-profile.md). Reuse sourced facts from the local profile and user-specified prior documents. Distinguish the report author from teammates; document order alone does not identify the user. Current explicit user information overrides older records. If required personal fields remain missing or conflicting after lookup, stop and ask the user in one concise batch, then wait for the answer before report generation or delivery. Do not silently leave those fields blank, guess, or bypass the gate by continuing independent sections. Only an explicit user instruction to omit or leave specific fields blank waives those fields. After the reply, apply the linked local profile and consent rules. This mandatory pause takes precedence over the general instruction to continue independent report sections.

## Read template once

Runtime selection happens before Word COM initialization: default `-Runtime Auto` prefers PowerShell 7+ and falls back to Windows PowerShell 5.1 when unavailable. Use `-Runtime PowerShell7` or `-Runtime WindowsPowerShell` to pin it. Keep `scripts/word_runtime.ps1` beside the engine. Both versions write BOM-free UTF-8 JSON; script sources retain a BOM for 5.1 Chinese parsing. See [runtime and encoding](references/vbs-word-automation.md#runtime-selection-and-encoding).

Run `scripts/word_report.ps1 -Mode inspect -Source <template> -Out <profile.json>` on Windows with Word. Cache uses source SHA-256 and schema version; `-Force` refreshes. Inspection is read-only and records effective paragraph/run formatting, sections, tables, fields and object counts. Mixed values stay explicit; never use Word's undefined sentinel as a size.

Visually inspect template patterns. Select role samples outside the cover, TOC and tables for body/heading/caption roles. A style name alone is not effective formatting. Save sample indices and approved corrections in the plan. Inspect ambiguity once; avoid repeated whole-document dumps. For legacy binary files with a DOCX suffix, Word converts only the output copy.

`extract_docx_contract.py` is an OOXML inventory, not an effective-format authority. `compare_docx_contracts.py` checks stable invariants; changing content counts alone is not failure.

## Structured content then one formatting pass

Before inserting generated prose into Word, complete [content review](references/content-review.md), fix findings and recheck changed passages. Formatting-only work preserves supplied wording unless rewriting is authorized.

Read [plan schema and automation](references/vbs-word-automation.md). Prefer reusable `word_report.ps1`; PowerShell and VBS both call Word COM. VBS is an optional adapter, not a reason to regenerate boilerplate. `emit_word_vbs.py` remains a legacy scaffold for specialized edits.

1. Map facts and required sections. Preserve template heading topics/order/hierarchy; remove assignment hints and word-count suffixes when required by the template.
2. Write UTF-8 JSON blocks with semantic roles. Do not use an unstructured blob and regexes to rediscover headings.
3. Specify unique exact-paragraph edit boundary, role samples, cover cell edits, blocks and required/forbidden text. Explicitly normalize red/highlighted instruction formatting and inconsistent indents.
4. Run `word_report.ps1 -Mode build -Source <template> -Plan <plan.json> -Out <result.docx>`. It retains the template, batches text insertion, applies effective roles, makes a real TOC, updates fields, repaginates and verifies generated paragraphs. Unsupported structures fail explicitly.
5. Keep original hashes, profile, plan and concise QA ledger task-local. Request summaries and failures rather than dumping XML/JSON, all source files or entire page text into model context.

Tables: read [table rules](references/academic-tables.md); copy reference examples where present. New plain tables use unshaded cells, readable borders, consistent padding and repeated headers. Follow explicit template directions. Figures use structured roles and inspected exports; mark implemented/planned boundaries in diagrams and prose.

## Bounded verification

Read [verification](references/verification-and-delivery.md). Stabilize content, run structural/format assertions, then a full-page final render review. When the report contains console captures, confirm each capture shows the prompt, command, input and output, and that no result caption is split from its image. Every final page must be inspected. Word PDF export is native pagination evidence; use an available renderer for page images as described in verification. Zero format mismatches is not visual proof.

Repairs reuse the plan and cached profile. Rerender affected pages plus pagination spillover; global layout changes require rechecking all affected pages. Finish with a fresh complete render and review coverage for the delivered version. Do not repeat project research, diagram generation or unchanged extraction after local edits. Record concrete defects and stop when requirements/layout pass; do not add speculative rejection loops.

No invented accomplishments, tests, dates or metrics. Conditional upgrades and learning plans stay separate from completed résumé claims. Never attach to or terminate the user's Word instance. Failed edits, unresolved mixed role formats and verification errors return nonzero. Deliver requested artifacts following the user's organization preferences.
