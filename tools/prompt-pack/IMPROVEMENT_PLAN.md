# PromptPack Improvement Plan

This document records the PromptPack performance and UX improvement plan implemented after the initial v1 version.

The goal is to make PromptPack faster, safer around slow Office/PDF conversions, and easier to launch in normal Windows file workflows.

---

## 1. Scope

Implemented changes:

- Add faster extraction paths for files that do not need Office COM.
- Keep PDF extraction on the Word PDF import / OCR path for consistent output formatting.
- Avoid custom PDF text-layer parsing because it can produce disordered text.
- Improve console UX with clearer phases, progress, spinner updates, and final results.
- Add optional Explorer context menu registration.
- Move generated output files into an `output` directory under the PromptPack script directory.

Non-goals:

- No automatic summarization.
- No automatic compression.
- No truncation by file size or token estimate.
- No external OCR dependency.
- No Tesseract integration.
- No Azure OCR integration.
- No GUI application.
- No installer package.
- No SendTo shortcut.

---

## 2. Performance Plan

### 2.1 C# helper through `Add-Type`

PromptPack loads a small C# helper with PowerShell `Add-Type`.

Expected useful areas:

- Fast text file reads.
- Encoding detection helpers.
- Lightweight string processing.

Areas where C# will not solve the bottleneck:

- Word COM startup.
- Excel COM startup.
- PowerPoint COM startup.
- Word PDF import / OCR itself.
- Office dialogs or Office-side hangs.

Design rule:

- Fast, deterministic work should stay in the parent process.
- Slow or risky Office COM work should remain isolated in worker processes.

### 2.2 Text file fast path

Text-like formats are handled in the parent process by the fast helper instead of spawning a worker for every simple file.

Examples:

- `.txt`
- `.md`
- `.csv`
- `.json`
- `.xml`
- `.html`
- `.css`
- `.js`
- `.ts`
- `.py`
- `.ps1`
- `.bat`
- `.cmd`
- `.sql`
- `.yaml`
- `.yml`
- `.ini`
- `.log`

Rules:

- Preserve source content.
- Do not summarize.
- Do not truncate.
- If decoding fails, record a failure instead of silently skipping the file.

### 2.3 PDF Word import path

PDF handling uses Word PDF import / OCR directly.

Reason:

- Custom direct PDF text-layer extraction can produce disordered text.
- Word already provides the desired reading and OCR behavior in the target environment.
- Keeping PDF on the Word path makes text-layer PDFs and scanned PDFs follow one consistent extraction format.

Rules:

- Do not add external PDF libraries.
- Do not use a separate OCR engine.
- Do not do PDF-only temporary path rewriting.
- Do not pass a special PDF converter format to Word.
- Call Word with the same resolved source path used by other file types.
- Word PDF handling should stay close to the proven production pattern:
  - `New-Object -ComObject 'Word.Application'`
  - `$word.Visible = $false`
  - `$doc = $word.Documents.Open($Path)`
  - `$text = $doc.Content.Text`

### 2.4 Worker boundary

Worker processes remain required for operations that can hang:

- Word files.
- Excel files.
- PowerPoint files.
- PDF files through Word PDF import / OCR.

Timeout behavior remains:

- Main pass timeout defaults to 120 seconds per file.
- Retry timeout defaults to 300 seconds per file.
- A timed-out file becomes `DEFERRED_TIMEOUT`.
- The main pass continues.
- Deferred retry is selectable, not automatic.

---

## 3. Console UX Plan

The console makes long operations visibly alive.

### 3.1 Phases

Display these phases:

- `SCAN`
- `PLAN`
- `EXTRACT`
- `WRITE`
- `DEFERRED`
- `DONE`

### 3.2 Progress display

Show:

- Current phase.
- Output path.
- Total file count.
- Processed file count.
- OK count.
- Deferred count.
- Failed count.
- Unsupported count.
- Current file number.
- Current file type.
- Current file name.
- Extraction method.
- Elapsed time.
- Timeout limit for worker-backed operations.

Example:

```text
PromptPack
----------

Phase: EXTRACT
Output: C:\repos\sakana-oss\tools\prompt-pack\output\promptpack_20260702_143000.txt

Files:     17 / 48
OK:        15
Deferred:  1
Failed:    1

Current:
[18/48] PDF  report.pdf
Method: Word OCR
Elapsed: 00:00:12 / 120s
Status: Working /
```

### 3.3 Spinner

Use an ASCII spinner for waiting states:

```text
|
/
-
\
```

Rules:

- No emoji.
- English-only script-generated text.
- Keep the final summary readable even if the live display becomes distorted.
- Use colors where helpful, but do not depend on colors for meaning.

### 3.4 Deferred retry prompt

If deferred files exist, show them after the main output file is written.

Example:

```text
Deferred files found: 2

[1] PDF  scanned-form.pdf
[2] DOCX large-document.docx

Retry deferred files now?
[A] Retry all
[S] Skip
Choice:
```

Default choice:

- `S`

Retry must not start automatically unless the command-line option explicitly requests it.

---

## 4. Output Directory Plan

Generated bundles should be written under the PromptPack script directory:

```text
tools/prompt-pack/output/
```

Rules:

- Create the `output` directory if it does not exist.
- Use `Join-Path` for output path construction.
- Use a timestamped file name.

Example:

```text
promptpack_20260702_143000.txt
```

Repository hygiene:

- Add `tools/prompt-pack/output/.gitignore`.
- Keep generated `.txt` output files out of git.

---

## 5. Context Menu Plan

PromptPack should support optional Explorer right-click registration.

Files to add:

- `Install-PromptPackContextMenu.bat`
- `Uninstall-PromptPackContextMenu.bat`

No SendTo shortcut should be created.

### 5.1 Registry locations

Use `HKCU` only. Administrator privileges must not be required.

File context menu:

```text
HKCU\Software\Classes\*\shell\PromptPack
HKCU\Software\Classes\*\shell\PromptPack\command
```

Directory context menu:

```text
HKCU\Software\Classes\Directory\shell\PromptPack
HKCU\Software\Classes\Directory\shell\PromptPack\command
```

Optional future target:

```text
HKCU\Software\Classes\Drive\shell\PromptPack
HKCU\Software\Classes\Drive\shell\PromptPack\command
```

### 5.2 Install behavior

The install BAT should:

- Resolve the current `PromptPack.bat` path.
- Write fixed registry keys.
- Set a readable menu label such as `Run PromptPack`.
- Set the command to call `PromptPack.bat "%1"`.
- Store metadata values where useful:
  - `InstallPath`
  - `ScriptPath`
  - `InstalledAt`

If PromptPack is moved, the menu command can become stale. This is expected for a registry-backed shell verb.

Mitigation:

- Re-running install overwrites the command with the new current path.
- Uninstall does not depend on the original install path.

### 5.3 Uninstall behavior

The uninstall BAT should delete fixed keys only:

```bat
reg delete "HKCU\Software\Classes\*\shell\PromptPack" /f
reg delete "HKCU\Software\Classes\Directory\shell\PromptPack" /f
```

This means a newly downloaded PromptPack can remove an old stale registration even if the original installed folder is gone.

### 5.4 Windows 11 note

The registry shell verbs may appear under "Show more options" on Windows 11.

PromptPack should not implement a modern Windows 11 shell extension in this phase, because that would add complexity and move away from the lightweight, no-admin design.

---

## 6. Path Handling Rules

Existing path safety rules remain required:

- Use `Get-Item -LiteralPath`.
- Use `Join-Path`.
- Quote paths in BAT files.
- Do not use `Invoke-Expression`.
- Do not concatenate paths into executable command strings.
- Support spaces.
- Support Japanese characters.
- Support parentheses.
- Support `#`.
- Support `&`.
- Support OneDrive and SharePoint synchronized paths.

PDF handling must use the same path rules as other file types. It should not introduce PDF-only path escaping or temporary path relocation.

---

## 7. Test Plan

Before push, verify:

- Script files remain ASCII-only where required.
- PowerShell syntax is valid.
- Text files use the fast path.
- PDFs use Word PDF import / OCR directly.
- Word files still extract.
- Excel files still extract.
- PowerPoint files still extract.
- Timeout produces `DEFERRED_TIMEOUT`.
- Deferred retry prompt appears only when needed.
- Paths with spaces work.
- Paths with Japanese characters work.
- Paths with `#` work.
- Paths with `&` work.
- Output is written to `output/`.
- Generated output files do not dirty the git worktree.
- Context menu install writes only HKCU keys.
- Context menu uninstall removes fixed keys even if the install path changed.

---

## 8. Suggested Implementation Order

1. Add `output/` creation and `.gitignore`.
2. Add context menu install/uninstall BAT files.
3. Improve console phase display and final summary.
4. Add spinner display for worker waits.
5. Add C# helper skeleton through `Add-Type`.
6. Move text files to the fast parent-process path.
7. Keep PDFs on the Word PDF import / OCR path.
8. Do not use custom direct PDF text extraction.
9. Run Windows Helm tests.
10. Commit and push.

---

## 9. Acceptance Criteria

The improvement set is done when:

- Existing drag-and-drop usage still works.
- Optional right-click registration works without administrator privileges.
- Uninstall works without knowing the original install folder.
- Output files are created under `tools/prompt-pack/output/`.
- Text files are faster than the all-worker v1 behavior.
- PDFs use the same Word PDF import / OCR path.
- A hung Office/PDF operation cannot block the whole run.
- The user can see progress, current file, elapsed time, and spinner activity.
- Deferred retry remains user-selectable.
- Script-generated text remains English-only.
