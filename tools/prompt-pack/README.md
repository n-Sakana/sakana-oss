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
- PDF: `.pdf`, opened through Word PDF import / OCR

## Notes

- The script does not summarize, compress, or truncate source content.
- Unsupported and failed files are recorded in the output.
- Script-generated labels and console messages are English-only.
- Microsoft Office is required for Office files and PDF import.
