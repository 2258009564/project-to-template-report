# Word COM plan and automation

The reusable implementation is `scripts/word_report.ps1`, requiring Windows PowerShell and Microsoft Word. No embedded macro permissions are needed. VBS, VBA and PowerShell share Word's object model; PowerShell supplies JSON support without a generated VBS parser.

## Inspect

`powershell -NoProfile -File scripts/word_report.ps1 -Mode inspect -Source template.docx -Out profile.json`

Cache invalidates on SHA-256/schema change. Profiles include Word table paragraphs: do not mix these indices with python-docx indices. Font and paragraph dictionaries contain effective values. Mixed properties include run samples. Body samples must exclude cover, TOC, tables and empty paragraphs.

## Plan

```json
{
  "replace_from": "目录",
  "roles": {
    "h1": {"sample": 55, "outline": 1},
    "body": {"sample": 59, "font": {"Color": 0, "Bold": 0}, "paragraph": {"FirstLineIndent": 24}}
  },
  "cover_cells": [{"table": 1, "row": 4, "column": 2, "text": "：真实姓名"}],
  "blocks": [{"kind": "text", "role": "h1", "text": "一、引言"}],
  "required": ["一、引言"],
  "forbidden": ["XXXX"],
  "pdf": "absolute-internal-preview.pdf"
}
```

`replace_from` must match exactly one standalone paragraph. Its tail through the final paragraph is replaced; preceding cover and section break are retained. This is an authorized boundary, not automatic cover detection. Arbitrary partial edits need a bounded bookmark/content-control adapter.

Role `sample` is mandatory. Optional `font`, `paragraph`, `outline` override only named properties. Explicit corrections may remove red instructions or normalize an inconsistent body indent. Undefined mixed values (9999999) must be resolved. Resume headings can use outline 10 to stay outside TOC.

Blocks:
- `text`: role, text and optional page_break_before.
- `toc`: heading role and text; engine makes an actual Word TOC for outline 1–2.
- `image`: role, absolute path, width_cm; caption follows as a separate text block.
- `table`: role and rectangular rows of strings; text-area width, unshaded cells, repeated header; caption precedes as a text block.

The engine bookmarks generated blocks before TOC expansion, so range movement does not invalidate verification. Caption/image keep behavior is set via roles. Cover cells preserve existing font/paragraph formatting and only their text changes.

## Failure and lifecycle

Open reference read-only in a new hidden Word instance. Save output to a separate temporary DOCX, verify, then replace requested output. Close created objects in finally. Never use global error suppression or terminate all Word processes. Verify required/forbidden text, TOC, original hash, section geometry and effective role properties. Write concise QA JSON beside result.

Enumerate once for extraction; insert text in one operation and format known ranges once. Cache hits and verification counts are reported instead of full profiles. Shell duration and model token use are separate costs. Legacy emit_word_vbs.py is for unsupported specialized edits only, not the default path.
