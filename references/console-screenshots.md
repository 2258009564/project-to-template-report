# Console screenshots as run evidence

Use this module when the assignment, template, teacher's slides or the user ask for
command-line output evidence: “运行结果截图”, “命令行截图”, “贴上运行截图” or a slide that
says “结果截图”. Capture an actual execution and retain its script, input and exit
status; a screenshot alone does not establish program correctness. Text pasted into
a monospace paragraph does not satisfy an explicit screenshot requirement.

## When a capture is required

- The template or assignment has a result section that asks for a screenshot, a run,
  a terminal or “实拍”.
- The user asks to replace text results with screenshots.
- The assignment requires visual run evidence for a compiled program.

When only the source code is requested, do not add captures; they lengthen the report
without adding evidence.

## Environment

| Part | Requirement |
| --- | --- |
| Terminal | Windows Terminal (`wt.exe`) is required; there is no legacy-window capture backend |
| Shell | `-Shell Auto` prefers PowerShell 7, then 5.1; `-Shell pwsh` and `-Shell powershell` pin the choice |
| Capture | `console_shot.ps1`, `console_shot_helpers.ps1` and `console_shot_runner.ps1` must remain together |
| Crop | Python + Pillow; `scripts/trim_console_shot.py` conservatively trims trailing whitespace to a separate PNG |
| Insertion | inline picture sized to the template and readability, often 10–11 cm, directly below its result caption |

Check the shell before promising it: `Get-Command pwsh.exe`. If only Windows
PowerShell is available, say so in the report's environment line and in the delivery
note; do not describe a 5.1 session as PowerShell 7.

## Workflow

1. Keep a clean work directory: `src/`, `bin/`, `cases/`. Compile with the real
   toolchain and record its version.
2. Store one input file per instance. Write a demo script per instance that sets a
   unique window title, prints the prompt and the input lines, then pipes exactly
   those lines into the binary. The binary's own stdout follows.
3. Request the capture with `console_shot.ps1`. It always opens a new WT window,
   assigns a unique session token and waits for the runner's atomic completion
   record. `TimeoutSec` bounds launch/execution/window discovery; `SettleMs` is an
   additional rendering delay after successful execution, not a completion signal.
4. Keep the raw PNG and trim only empty margins into a separate file. Inspect that
   no prompt, command, input or output was removed. Tab chrome is not run evidence.
5. Insert the PNG below “结果N：” at a readable template-compatible width, set that caption paragraph to keep
   with the next paragraph, and verify every rendered page.
6. Keep captures and demo scripts task-local; deliver only the report.

## Capture commands

```powershell
# Auto prefers PowerShell 7 and falls back to 5.1. Both require Windows Terminal.
pwsh -NoProfile -File scripts/console_shot.ps1 `
  -ScriptFile C:\work\lab\tools\case-1.ps1 `
  -WorkingDirectory C:\work\lab `
  -Title ALG-LAB-1 -Out C:\work\lab\shots\raw-1.png

python scripts/trim_console_shot.py `
  C:\work\lab\shots\raw-1.png C:\work\lab\shots\shot-1.png --inset 2
```

`console_shot.ps1` fails when the window is not found, the client area is invalid or
the run fails, times out, is obscured or has a blank terminal body. It validates
before replacing the requested output. Existing output from an earlier run remains
unchanged on failure: never treat that old file as a new successful capture.
`-KeepOpen` leaves only this run's dedicated window open for manual inspection.
Launching, sizing and foregrounding this new window are visible operations; follow
the current user's foreground-interaction preferences. No terminal settings are edited.

Keep demo scripts noninteractive and let them return normally. Print actual input
before passing it to the program, check native exit codes and propagate errors.
Do not use `exit` in an inline `-Command` to terminate its capture host. The tool
cannot prove that long output fits in the viewport: visually review every result.

Pillow is an optional dependency for cropping (`python -m pip install Pillow`).
Omit the crop step if the raw capture is already suitable. With no destination,
the crop tool creates `<stem>-trimmed.png`; it rejects in-place writes, uniform
images and invalid margins. Crop thresholds are heuristics and require visual review.

## Rules

- Never fabricate output or reuse a capture from another run: no synthetic output, no patched numbers, no
  borrowed screenshot from another run.
- A capture must contain the prompt, the executed command, the input and the output.
- One capture per instance, numbered as the assignment numbers them.
- If a run fails, fix the program and capture again instead of hiding the failure.
- Long output lines may wrap; the command itself must not wrap mid-token, otherwise
  widen the window.
- Name the terminal and shell in the report, for example “Windows Terminal + PowerShell 7”.
- Leave user terminal settings untouched; close only the dedicated window created for this run.

## Pitfalls

- DPI: the capture thread temporarily uses per-monitor awareness and restores it
  afterwards, rather than silently ignoring a process-wide DPI-setting failure.
- Occlusion: a covered window captures whatever is on top. Bring the window to the
  foreground and wait for rendering before capturing.
  This implementation uses a checked, visible client-area screen copy. It does
  not promise an occluded-window capture through `PrintWindow`; a successful API
  return alone would not verify GPU-rendered terminal contents.
- Short windows: content scrolls off the top and the prompt disappears. Size the window
  to the expected line count before capturing, then trim.
- Blank images: the body is checked independently of the tab strip. Five arbitrary
  background points are not sufficient to distinguish sparse text from an empty screen.
- File locks: Word keeps the previous report open while it is being reviewed, so a
  rebuild may fail to replace it. Write the next build to a new name and say so.
- Shell selection is not a guarantee of correctness or completion: trust the exit
  status and completion record, then review the actual pixels.

## Regression checks

```powershell
# Background executions and content detection; no windows are opened.
pwsh -NoProfile -File scripts/test_console_shot.ps1 -WorkDir C:\work\capture-tests
python scripts/test_trim_console_shot.py

# Explicitly opt into visible capture checks for both installed shells.
pwsh -NoProfile -File scripts/test_console_shot.ps1 -WorkDir C:\work\capture-tests-ui -IncludeWindowCapture
```
