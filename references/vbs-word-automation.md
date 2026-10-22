# Word COM plan and automation

The reusable implementation is `scripts/word_report.ps1`, requiring Windows and Microsoft Word. PowerShell 7+ is preferred; Windows PowerShell 5.1 remains a compatibility option. No embedded macro permissions are needed. VBS, VBA and PowerShell share Word's object model.

## Runtime selection and encoding

Both the engine and `test_word_report.ps1` accept `-Runtime Auto|PowerShell7|WindowsPowerShell`:

- `Auto` (default): keep a running PowerShell 7+ host; otherwise switch to `pwsh.exe` when available. If it is missing, continue in Windows PowerShell 5.1 with a warning.
- `PowerShell7`: require PowerShell 7+. Fail clearly if it cannot be found; do not silently downgrade.
- `WindowsPowerShell`: explicitly run in Windows PowerShell 5.1, even if started from `pwsh`.

The runtime is selected before Word is started. `word_runtime.ps1` forwards parameter values and switches via `-File`, keeps the current working directory, reports the selected runtime, and propagates child exit codes. It also saves DOCX through typed by-reference `SaveAs2` arguments to avoid the implicit-conversion problem reproduced under 5.1. Keep this helper beside the engine when copying files. Runtime selection does not provide a COM timeout or guarantee that all Word hangs disappear; no version speedup is claimed without a controlled measurement.

PowerShell source files carry a UTF-8 BOM so 5.1 can parse Chinese text. Generated profiles, QA records and test JSON use explicit UTF-8 **without** BOM in both runtimes. Legacy BOM-bearing inspect caches are normalized on a cache hit without invoking Word. JSON serialization explicitly sets depth 30; PowerShell 7 also defaults to depth 2, so upgrading alone does not prevent truncation. See Microsoft's [encoding documentation](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_character_encoding) and [ConvertTo-Json documentation](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/convertto-json).

## Inspect

```powershell
# Prefer PowerShell 7 explicitly.
pwsh -NoProfile -File scripts/word_report.ps1 -Runtime PowerShell7 -Mode inspect -Source template.docx -Out profile.json

# Windows built-in entrypoint: automatically switch to 7 when installed, otherwise use 5.1.
powershell -NoProfile -File scripts/word_report.ps1 -Mode inspect -Source template.docx -Out profile.json

# Explicit compatibility mode.
powershell -NoProfile -File scripts/word_report.ps1 -Runtime WindowsPowerShell -Mode inspect -Source template.docx -Out profile.json
```

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
