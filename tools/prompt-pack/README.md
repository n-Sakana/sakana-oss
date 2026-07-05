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
- `logs/` - per-run diagnostic logs; contents are ignored by git

## Usage

Drag files or folders onto `PromptPack.bat`.

When launched from Explorer, the BAT file reopens itself in a persistent command window so that progress, errors, and final results remain visible. Close the window manually after reading the result.

The BAT file starts PowerShell without administrator privileges:

```bat
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "PromptPack.ps1"
```

Generated bundles and diagnostic logs are written under the tool directory:

```text
tools/prompt-pack/output/promptpack_yyyyMMdd_HHmmss.txt
tools/prompt-pack/logs/promptpack_yyyyMMdd_HHmmss.log
```

The bundle is kept clean for AI chat input: file tree, extracted contents, and failed/deferred/unsupported file lists. Detailed runtime information is written to the log file instead.

You can override the bundle path with `-OutFile` when running `PromptPack.ps1` directly. The log file still goes under `logs/`.

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

## Office and PDF behavior

Modern Office files are read directly from their Open XML package with the built-in C# helper loaded through PowerShell `Add-Type`.

Fast Open XML paths:

- `.docx`, `.docm` - read `word/*.xml` parts from the package
- `.xlsx`, `.xlsm` - read workbook, shared strings, and worksheet XML parts
- `.pptx`, `.pptm` - read slide and notes XML parts

These files are ZIP/XML packages already. PromptPack does not create a new ZIP file and does not extract folders to disk; it opens the package in memory and reads the XML entries.

Legacy Office and PDF paths still use Office COM:

- `.doc`, `.rtf` - Word COM
- `.xls`, `.xlsb` - Excel COM
- `.ppt` - PowerPoint COM
- `.pdf` - Word PDF import / OCR

If an Open XML read fails, PromptPack records the fallback reason and retries that file through the relevant Office COM worker.

PDF files are always extracted through Microsoft Word COM. The call is kept close to the proven pattern:


```powershell
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open($Path)
$text = $doc.Content.Text
```

There is no direct PDF text-layer parser, no PDF-only temporary path rewrite, no converter format override, and no external PDF/OCR dependency.

## Performance and timeout behavior

Fast paths run in the parent process:

- Text-like files
- Modern Office Open XML files: `.docx`, `.docm`, `.xlsx`, `.xlsm`, `.pptx`, `.pptm`

COM-backed work remains isolated in a per-file worker PowerShell process:

- Legacy Word files
- Legacy Excel files
- Legacy PowerPoint files
- Open XML fallback cases
- PDF files through Word PDF import / OCR

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

## Console UX and logs

The console is intentionally minimal:

```text
PromptPack
Running...
Done.

Output:
...\output\promptpack_yyyyMMdd_HHmmss.txt

Log:
...\logs\promptpack_yyyyMMdd_HHmmss.log
```

Detailed progress is written to the run log instead of the console or bundle. The log includes input paths, collected file list, route decisions, worker start/exit, stage transitions, elapsed time, timeout stage, fallback reason, stdout/stderr, and final output path.

## Notes

- The script does not summarize, compress, or truncate source content.
- Unsupported, failed, and deferred files are recorded in the output.
- Script-generated labels and console messages are English-only.
- Microsoft Office is required for legacy Office formats, Open XML fallback, and PDF OCR/import.
