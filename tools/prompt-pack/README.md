# PromptPack

PromptPack bundles files and folders into one structured text file for generative AI chat tools.

## Files

- `PromptPack.bat` - drag-and-drop entry point for Windows
- `PromptPack.ps1` - extraction and bundling script
- `Install-PromptPackContextMenu.bat` - starts the optional current-user Explorer context menu installer
- `Install-PromptPackContextMenu.ps1` - writes the HKCU context menu registry keys
- `Uninstall-PromptPackContextMenu.bat` - starts the context menu uninstaller
- `Uninstall-PromptPackContextMenu.ps1` - removes the HKCU context menu registry keys
- `SPEC.md` - v1 specification
- `IMPROVEMENT_PLAN.md` - performance and UX improvement record
- `output/` - generated bundles; contents are ignored by git

## Usage

Drag files or folders onto `PromptPack.bat`.

The BAT file starts PowerShell without administrator privileges:

```bat
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "PromptPack.ps1"
```

Generated bundles are written under the tool directory:

```text
tools/prompt-pack/output/promptpack_yyyyMMdd_HHmmss.txt
```

You can override the output path with `-OutFile` when running `PromptPack.ps1` directly.

## Optional Explorer context menu

Run this once to add `Run PromptPack` to the current user's file and folder right-click menus:

```bat
Install-PromptPackContextMenu.bat
```

The installer opens a persistent command window when launched from Explorer so that errors and success messages do not disappear immediately. Close the window manually after reading the result.

Run this to remove it:

```bat
Uninstall-PromptPackContextMenu.bat
```

The uninstaller uses the same persistent-window behavior.

The installer writes only to HKCU:

```text
HKCU\Software\Classes\*\shell\PromptPack
HKCU\Software\Classes\Directory\shell\PromptPack
```

No SendTo shortcut is created. If the tool directory is moved, run the installer again to update the command path. The uninstaller deletes fixed keys, so it can remove old registrations even from a newly downloaded copy.

Manual removal:

```bat
reg delete "HKCU\Software\Classes\*\shell\PromptPack" /f
reg delete "HKCU\Software\Classes\Directory\shell\PromptPack" /f
```

## Supported inputs

- Text files: `.txt`, `.md`, `.csv`, `.json`, `.xml`, `.html`, `.css`, `.js`, `.ts`, `.py`, `.ps1`, `.bat`, `.sql`, `.yaml`, `.ini`, `.log`, and related formats
- Word: `.docx`, `.doc`, `.docm`, `.rtf`
- Excel: `.xlsx`, `.xls`, `.xlsm`, `.xlsb`
- PowerPoint: `.pptx`, `.ppt`, `.pptm`
- PDF: `.pdf`

## PDF behavior

PDF extraction is two-stage:

1. Try direct text-layer extraction with a built-in C# helper loaded through `Add-Type`.
2. If no usable text is found, fall back to Word PDF import / OCR.

The Word fallback uses the same resolved source path as other files and keeps the call close to the proven pattern:

```powershell
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open($Path)
$text = $doc.Content.Text
```

There is no PDF-only temporary path rewrite, no converter format override, and no external PDF/OCR dependency.

## Performance and timeout behavior

Fast paths run in the parent process:

- Text-like files
- PDF files with directly extractable text layers

Office-backed work remains isolated in a per-file worker PowerShell process:

- Word files
- Excel files
- PowerPoint files
- PDFs that need Word PDF import / OCR

Defaults:

- Main pass timeout: 120 seconds per worker-backed file
- Retry timeout: 300 seconds per worker-backed file
- Deferred retry mode: ask the user after the main output is written

If a file times out, the main pass continues and records the file as `DEFERRED_TIMEOUT`.

Command-line example:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\PromptPack.ps1 -TimeoutSeconds 120 -RetryTimeoutSeconds 300 -DeferredAction Ask "C:\path\to\folder"
```

For unattended tests, use `-DeferredAction Skip`.

## Console UX

PromptPack shows phase markers and visible progress:

- `SCAN`
- `PLAN`
- `EXTRACT`
- `WRITE`
- `DEFERRED`
- `DONE`

Worker-backed operations show an ASCII spinner with current stage, elapsed seconds, and timeout limit.

## Notes

- The script does not summarize, compress, or truncate source content.
- Unsupported, failed, and deferred files are recorded in the output.
- Script-generated labels and console messages are English-only.
- Microsoft Office is required for Office files and for PDF OCR/import fallback.
