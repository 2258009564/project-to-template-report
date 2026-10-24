# Console screenshots as run evidence

Use this module when the assignment, template, teacher's slides or the user ask for
command-line output evidence: “运行结果截图”, “命令行截图”, “贴上运行截图” or a slide that
says “结果截图”. A screenshot is proof of a real run. Text pasted into a monospace
paragraph is not a screenshot and does not satisfy such a requirement.

## When a capture is required

- The template or assignment has a result section that asks for a screenshot, a run,
  a terminal or “实拍”.
- The user asks to replace text results with screenshots.
- The task involves a compiled program whose behaviour is part of the evidence.

When only the source code is requested, do not add captures; they lengthen the report
without adding evidence.

## Environment

| Part | Requirement |
| --- | --- |
| Terminal | Windows Terminal (`wt.exe`); a legacy console window only when Windows Terminal is unavailable |
| Shell | PowerShell 7 (`pwsh.exe`) preferred; Windows PowerShell 5.1 is an acceptable fallback |
| Capture | `scripts/console_shot.ps1` captures the window client area as PNG (PrintWindow first, screen copy fallback) |
| Crop | `scripts/trim_console_shot.py` removes the empty rows below the last text row |
| Insertion | inline picture at 10–11 cm width, directly below its result caption |

Check the shell before promising it: `Get-Command pwsh.exe`. If only Windows
PowerShell is available, say so in the report's environment line and in the delivery
note; do not describe a 5.1 session as PowerShell 7.

## Workflow

1. Keep a clean work directory: `src/`, `bin/`, `cases/`. Compile with the real
   toolchain and record its version.
2. Store one input file per instance. Write a demo script per instance that sets a
   unique window title, prints the prompt and the input lines, then pipes exactly
   those lines into the binary. The binary's own stdout follows.
3. Launch the demo script in Windows Terminal through `console_shot.ps1`, with a
   dedicated title, a fixed window size and a timeout.
4. Trim the empty area below the last text row. Keep the tab strip when it identifies
   the terminal; it is cheap authenticity for the reader.
5. Insert the PNG below “结果N：” at 10–11 cm width, set that caption paragraph to keep
   with the next paragraph, and verify every rendered page.
6. Keep captures and demo scripts task-local; deliver only the report.

## Capture commands

```powershell
# PowerShell 7 preferred; pass -Shell powershell when only 5.1 exists.
pwsh -NoProfile -File scripts/console_shot.ps1 `
  -ScriptFile C:\work\lab\tools\case-1.ps1 `
  -WorkingDirectory C:\work\lab `
  -Title ALG-LAB-1 -Out C:\work\lab\shots\raw-1.png

python scripts/trim_console_shot.py `
  C:\work\lab\shots\raw-1.png C:\work\lab\shots\shot-1.png --inset 2
```

`console_shot.ps1` fails when the window is not found, the client area is invalid or
the capture is blank; do not ignore those failures. `-KeepOpen` leaves the window open
for manual inspection.

## Rules

- Never fabricate, edit or reuse a capture: no synthetic output, no patched numbers, no
  borrowed screenshot from another run.
- A capture must contain the prompt, the executed command, the input and the output.
- One capture per instance, numbered as the assignment numbers them.
- If a run fails, fix the program and capture again instead of hiding the failure.
- Long output lines may wrap; the command itself must not wrap mid-token, otherwise
  widen the window.
- Name the terminal and shell in the report, for example “Windows Terminal + PowerShell 7”.
- Restore any terminal setting changed for the capture, and close the windows you opened.

## Pitfalls

- DPI: the capture process must be per-monitor DPI aware, or the crop is offset and
  blurred. `console_shot.ps1` sets that awareness itself.
- Occlusion: a screen copy captures whatever is on top. `console_shot.ps1` renders the
  window through PrintWindow first and only falls back to a screen copy, so a window
  that another application covers still yields a full image.
- Short windows: content scrolls off the top and the prompt disappears. Size the window
  to the expected line count before capturing, then trim.
- Blank images: a uniform capture means the window was hidden or not yet painted;
  rerun the capture instead of shipping it.
- File locks: Word keeps the previous report open while it is being reviewed, so a
  rebuild may fail to replace it. Write the next build to a new name and say so.
- Terminal settings are user data. Back up `settings.json` before changing font or
  colour for a capture, and restore it afterwards.
