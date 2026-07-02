# PromptPack

PromptPack bundles files and folders into one structured text file for generative AI chat tools.

## Files

- `PromptPack.bat` — drag-and-drop entry point for Windows
- `PromptPack.ps1` — extraction and bundling script
- `SPEC.md` — v1 specification

## Usage

Drag files or folders onto `PromptPack.bat`.

The BAT file starts PowerShell without administrator privileges:

```bat
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "PromptPack.ps1"
```

The output file is created next to the first dropped file, or inside the first dropped folder.

Example output name:

```text
prompt-pack_20260702_113000.txt
```

## Supported inputs

- Text files: `.txt`, `.md`, `.csv`, `.json`, `.xml`, `.html`, `.css`, `.js`, `.ts`, `.py`, `.ps1`, `.bat`, `.sql`, `.yaml`, `.ini`, `.log`, and related formats
- Word: `.docx`, `.doc`, `.docm`, `.rtf`
- Excel: `.xlsx`, `.xls`, `.xlsm`, `.xlsb`
- PowerPoint: `.pptx`, `.ppt`, `.pptm`
- PDF: `.pdf`, direct Word PDF import / OCR

## Timeout and deferred retry

Each file is extracted in a separate worker PowerShell process.

Defaults:

- Main pass timeout: 120 seconds per file
- Retry timeout: 300 seconds per file
- Deferred retry mode: ask the user after the main output is written

If a file times out, the main pass continues and records the file as `DEFERRED_TIMEOUT`.
After the main pass finishes, PromptPack writes the output file and then asks whether to retry deferred files.

Command-line example:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\PromptPack.ps1 -TimeoutSeconds 120 -RetryTimeoutSeconds 300 -DeferredAction Ask "C:\path\to\folder"
```

For unattended tests, use `-DeferredAction Skip`.

## Notes

- The script does not summarize, compress, or truncate source content.
- Unsupported, failed, and deferred files are recorded in the output.
- PDF files are opened directly with Word COM using the same resolved input path as other file types. No PDF-only temporary path rewrite, direct text pre-pass, or converter format override is used.
- Script-generated labels and console messages are English-only.
- Microsoft Office is required for Office files and for PDF OCR/import.
