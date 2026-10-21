# Adversarial review and iteration loop

Final review is an evidence-producing loop, not a last glance. Review the document as if rejecting a submission for any visible or factual defect.

## QA ledger

Maintain a task-local `qa.md` with one row per finding:

| Iteration | Document and page | Object | Finding | Severity | Repair | Fresh evidence | Status |
|---|---|---|---|---|---|---|---|

Use `open`, `fixed`, `rechecked`, or `accepted limitation`. A finding closes only after a new render or structural check proves the repair.

## Full-page review

1. Refresh Word fields and repaginate in a separate hidden Word instance.
2. Render to a new empty iteration directory. Never inspect stale PNGs.
3. Confirm the rendered page count against Word's page count and explain any renderer difference.
4. Inspect every page at readable resolution. Contact sheets may locate density or blank-page problems, but do not replace full-resolution inspection.
5. Review each page once against these criteria, revisiting concrete failures rather than running separate complete passes for every criterion:
   - content and factual consistency with the project;
   - heading hierarchy, TOC and page numbering;
   - paragraph spacing, indentation and orphan/widow behavior;
   - every table's width, borders, shading, alignment, wrapping and page breaks;
   - every figure's content, connector geometry, sharpness, caption and surrounding explanation;
   - cover, signatures, dates and unknown formal fields;
   - final-page completeness and absence of placeholders or internal notes.

## Cross-document review

For a document set, compare shared metadata, terminology, module names, dates, roles, requirements, figure names and version identifiers. A fact corrected in one document must be searched across the entire set.

Verify phase traceability:

- task and schedule commitments appear in the log;
- requirements map to overview components and interfaces;
- overview components map to detailed classes/functions/data structures;
- detailed design maps to implementation symbols and tests;
- figures do not claim semantics absent from code.

## Adversarial rejection conditions

Reopen the iteration if any of these occurs:

- a page was not inspected at full resolution;
- a mandatory figure is absent, misclassified or unreadable at Word size;
- a connector touches text, an arrow is detached, or a class/component has unsupported semantics;
- a table uses decorative fill, clips text, loses its header or violates the selected three-line/plain-grid rule;
- TOC, caption, cross-reference or page fields are stale;
- a formal field is invented, a placeholder remains, or project facts conflict across documents;
- a repair changes pagination but affected spillover pages are not rechecked.

## Iteration stop condition

Stop when the ledger has no open finding, the delivered version has fresh full-page renders and complete review coverage, structural checks pass, source template hashes remain unchanged, and claimed behavior has evidence. A local repair reuses cached extraction and research; recheck changed pages and pagination spillover. Global layout changes require full affected-document review. Report accepted limitations explicitly.

## Skill improvement

When the user has authorized ongoing evolution of this installed skill, review the QA ledger after delivery preparation:

1. Identify failures caused by missing or weak workflow guidance, not by one-off document content.
2. Patch the narrowest instruction, reference, or deterministic script that would prevent recurrence.
3. Preserve existing scope and authorization boundaries.
4. Run `quick_validate.py` and reproduce the previously failing check.
5. Record the skill change in the task QA ledger; do not create a generic changelog.
