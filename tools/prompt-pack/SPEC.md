# PromptPack Specification v1

## 1. Purpose

PromptPack converts multiple files and folders into a single structured text file that can be attached to generative AI chat tools.

Typical use cases:

- Provide a set of documents to Copilot / ChatGPT / Claude / Gemini
- Work around limits on attachment count, file formats, or file size
- Bundle Office files, PDFs, and text files into one `.txt`
- Preserve file paths, file structure, extracted contents, and extraction failures
- Do not summarize, compress, truncate, or modify the source content

This tool is a bundler, not a summarizer.

---

## 2. Deliverables

The tool contains these files:

- `PromptPack.bat`
- `PromptPack.ps1`
- `Install-PromptPackContextMenu.bat`
- `Install-PromptPackContextMenu.ps1`
- `Uninstall-PromptPackContextMenu.bat`
- `Uninstall-PromptPackContextMenu.ps1`
- `README.md`
- `SPEC.md`
- `output/.gitignore`

Roles:

- `PromptPack.bat`
  - Receives files and folders by drag-and-drop
  - Requires no administrator privileges
  - Starts PowerShell with `ExecutionPolicy Bypass`

- `PromptPack.ps1`
  - Collects input files
  - Extracts text content
  - Builds a structured output text file
  - Shows progress and final results in the console

This directory contains the specification and v1 implementation files.

---

## 3. Execution Flow

User operation:

1. Drag files or folders onto `PromptPack.bat`
2. PowerShell starts
3. The script scans input paths
4. The script extracts fast-path text and modern Office Open XML files in the parent process
5. COM-backed files and PDF Word import/OCR run in per-file worker processes
6. A per-file timeout prevents a single hung worker from blocking the whole run
7. The script writes a single `.txt` output file under `output/` after the main pass
8. If timeout-deferred files exist, the user can choose whether to retry them
9. The console shows the output path and extraction summary

Expected `.bat` shape:

```bat
@echo off
setlocal

set "SCRIPT=%~dp0PromptPack.ps1"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*

pause
```

The `.bat` file itself must contain English text only.

---

## 4. Target Environment

- Windows
- PowerShell 5.1 or later
- Microsoft Office installed
- No administrator privileges required

Office COM automation is used for legacy Office documents, Open XML fallback, and PDF OCR/import through Word.

---

## 5. Supported File Types

### 5.1 Text Files

Supported examples:

- `.txt`
- `.md`
- `.csv`
- `.json`
- `.xml`
- `.html`
- `.htm`
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

Processing rules:

- Read as text
- Preserve content as much as possible
- Try common encodings such as UTF-8, UTF-8 BOM, UTF-16 LE, and CP932
- If the file cannot be decoded, record it as a failed file

### 5.2 Word

Supported:

- `.docx`
- `.doc`
- `.docm`
- `.rtf`

Processing rules:

- `.docx` and `.docm` are read through the built-in C# Open XML helper
- The helper reads ZIP/XML package entries directly without extracting folders to disk
- `.doc` and `.rtf` are opened with Word COM
- If Open XML extraction fails, retry through Word COM as a fallback and record the fallback reason
- Add the extracted text under a clear file boundary

### 5.3 Excel

Supported:

- `.xlsx`
- `.xls`
- `.xlsm`
- `.xlsb`

Processing rules:

- `.xlsx` and `.xlsm` are read through the built-in C# Open XML helper
- The helper reads workbook, shared strings, and worksheet XML entries directly
- `.xls` and `.xlsb` are opened with Excel COM
- If Open XML extraction fails, retry through Excel COM as a fallback and record the fallback reason
- Extract each worksheet separately
- Write sheet names
- Output cell values in a tab-separated or CSV-like text form

### 5.4 PowerPoint

Supported:

- `.pptx`
- `.ppt`
- `.pptm`

Processing rules:

- `.pptx` and `.pptm` are read through the built-in C# Open XML helper
- The helper reads slide and notes XML entries directly
- `.ppt` is opened with PowerPoint COM
- If Open XML extraction fails, retry through PowerPoint COM as a fallback and record the fallback reason
- Extract text from slides
- Write slide numbers clearly

### 5.5 PDF

Supported:

- `.pdf`

Processing rules:

- Resolve and validate the input path using the same path handling used for other file types
- Open the original PDF path directly with Word COM
- Do not use a direct PDF text-layer parser
- Do not use PDF-only temporary path rewriting
- Do not pass a special PDF converter format to Word
- Let Word handle PDF conversion and OCR for text, scanned, or handwritten PDFs
- Extract text from the opened Word document
- Keep the original path in the output

Rules:

- No separate OCR engine in v1
- No Tesseract integration in v1
- No Azure OCR integration in v1

---

## 6. Folder Handling

When a folder is dropped:

- Recursively collect supported files
- Preserve folder structure in the output
- Record unsupported or failed files when relevant

Default exclusions should be minimal.

Possible obvious exclusions:

- `.git`
- `node_modules`
- `.venv`
- `__pycache__`

v1 must not silently reduce the content by file size, importance, or estimated token count.

---

## 7. Path Handling

This is a critical requirement.

The tool must work with OneDrive and SharePoint synchronized folders.

Path cases to support:

- Spaces
- Japanese characters
- Parentheses
- `#`
- `&`
- Long folder names
- Deep SharePoint-derived paths

Implementation rules:

- Always quote the PowerShell script path in the `.bat`
- In PowerShell, use `-LiteralPath` for path-based operations
- Do not build executable command strings by concatenating paths
- Do not use `Invoke-Expression`
- Use `Join-Path` when constructing output paths
- Treat input paths as literal strings

---

## 8. Output File

Output format:

- Single `.txt`
- UTF-8
- Structured text
- Directly attachable to AI chat tools

Example file name:

```text
tools/prompt-pack/output/promptpack_20260702_111530.txt
```

Output location:

- Use `output/` under the PromptPack script directory by default
- Create the directory automatically when needed
- Keep `output/.gitignore` so generated bundles do not dirty the repository
- Use safe path construction with `Join-Path`
- Allow direct script users to override the location with `-OutFile`

---

## 9. Output Text Structure

All output labels and headings generated by the script must be English.

Example:

```text
# PromptPack Bundle

Generated At: 2026-07-02 11:15:30
Tool: PromptPack.ps1
Total Files: 12
Succeeded: 10
Failed: 2

---

## Input Paths

- C:\Users\...\Documents\Project
- C:\Users\...\Desktop\memo.docx

---

## File Tree

Project/
  proposal.pdf
  notes.md
  data.xlsx

---

## Extraction Summary

| No | Status | Type | Path | Method | Notes |
|---:|---|---|---|---|---|
| 1 | OK | PDF | C:\...\proposal.pdf | Word PDF Import | Opened directly with Word |
| 2 | OK | Markdown | C:\...\notes.md | Text Read | UTF-8 |
| 3 | FAIL | Excel | C:\...\locked.xlsx | Excel COM | Password protected |

---

## Contents

### File 1

Path: C:\...\proposal.pdf
Type: PDF
Method: Word PDF Import
Status: OK

[Extracted text here]

---

### File 2

Path: C:\...\notes.md
Type: Markdown
Method: Text Read
Status: OK

[Extracted text here]

---

## Failed Files

### Failed File 1

Path: C:\...\locked.xlsx
Type: Excel
Status: FAIL
Reason: Password protected or cannot be opened
```

---

## 10. Language Rules for Scripts

The script files themselves must not contain Japanese text.

English-only targets:

- Variable names
- Function names
- Comments
- Console logs
- Error messages
- Output headings
- Output labels
- `.bat` text

Extracted source document text may contain Japanese or any other language.

---

## 11. Console Display

The console must show visible progress while running.

Rules:

- English only
- Modern and readable
- Color-coded where practical
- No emoji
- Use ASCII spinner characters for worker-backed waiting states
- Show the final output path
- Show the final extraction summary

Displayed phases:

- `SCAN`
- `PLAN`
- `EXTRACT`
- `WRITE`
- `DEFERRED`
- `DONE`

Status examples:

- `OK_TEXT`
- `OK_DOCX_XML`
- `OK_XLSX_XML`
- `OK_PPTX_XML`
- `OK_WORD_COM`
- `OK_EXCEL_COM`
- `OK_POWERPOINT_COM`
- `OK_WORD_COM_FALLBACK`
- `OK_EXCEL_COM_FALLBACK`
- `OK_POWERPOINT_COM_FALLBACK`
- `OK_PDF_WORD`
- `FAIL`
- `UNSUPPORTED`
- `DEFERRED_TIMEOUT`

Example:

```text
PromptPack
----------

Phase: EXTRACT
Output: C:\repos\sakana-oss\tools\prompt-pack\output\promptpack_20260702_143000.txt

[18/48] PDF    Working / Stage: OPEN_PDF Elapsed: 12s / 120s  report.pdf
[18/48] PDF    OK_PDF_WORD      C:\path\report.pdf

Summary:
OK:          42
Failed:      2
Deferred:    3
Unsupported: 1
Total:       48
```

PowerShell should use `Write-Host` with colors and same-line console updates. Do not use `Write-Progress` for the main spinner because it can produce repeated lines in some hosts.

---

## 12. Timeout and Deferred Retry

Fast-path text extraction and modern Office Open XML extraction run in the parent process. COM-backed extraction and all PDF Word import/OCR work run in separate worker PowerShell processes.

Default settings:

- Main pass timeout: 120 seconds per worker-backed file
- Retry timeout: 300 seconds per file
- Deferred retry: ask the user after the main pass

Timeout behavior:

- If a worker exceeds the timeout, the parent process stops that worker
- The file is recorded as `DEFERRED_TIMEOUT`
- The main pass continues with the next file
- The output file is written after the main pass
- Deferred files are listed in the summary and in a dedicated section
- The user can choose whether to retry deferred files
- Retry is never automatic unless explicitly requested with a command-line option

Worker diagnostics:

- Workers write stage markers such as `START_WORD`, `OPEN_PDF`, `READ_CONTENT`, and `CLOSE_DOCUMENT`
- Timeout notes include the last known stage when available
- PDF handling uses the same resolved original path as the source input and opens it directly with Word COM

Command-line controls:

- `-TimeoutSeconds <n>`
- `-RetryTimeoutSeconds <n>`
- `-DeferredAction Ask|Retry|Skip`

---

## 13. Error Handling

The tool must not silently skip failed files.

For each failed file, record:

- Path
- File type
- Method
- Status
- Failure reason

Expected error cases:

- File cannot be opened
- Password-protected file
- Office COM failure
- Broken file
- Unsupported extension
- Encoding detection failure
- Permission problem

Failure records must appear both in the extraction summary and in the failed files section.

---

## 14. Optional Explorer Context Menu

PromptPack includes optional current-user Explorer context menu registration.

Included files:

- `Install-PromptPackContextMenu.bat`
- `Install-PromptPackContextMenu.ps1`
- `Uninstall-PromptPackContextMenu.bat`
- `Uninstall-PromptPackContextMenu.ps1`

Registry locations:

```text
HKCU\Software\Classes\*\shell\PromptPack
HKCU\Software\Classes\Directory\shell\PromptPack
```

Rules:

- HKCU only
- No administrator privileges
- No SendTo shortcut
- Registration is persistent until uninstall
- The uninstall script deletes fixed keys and does not depend on the original install path
- Re-running install updates the command path to the current tool directory

---

## 15. Non-Goals for v1

v1 does not do the following:

- Automatic summarization
- Automatic compression
- Truncation by size
- Importance ranking
- Content rewriting for AI
- Dedicated PDF OCR engine
- Tesseract integration
- Azure OCR integration
- Image embedding
- GUI app
- Installer

The tool only bundles source content into one structured text file.

---

## 16. Implementation Outline

Expected PowerShell functions:

- `Read-TextFile`
- `Read-WordOpenXmlFile`
- `Read-ExcelOpenXmlFile`
- `Read-PowerPointOpenXmlFile`
- `Read-WordFile`
- `Read-ExcelFile`
- `Read-PowerPointFile`
- `Read-PdfFileWithWord`
- `Invoke-FileExtractionWithTimeout`
- `Invoke-WorkerMode`
- `Write-ConsoleStatus`
- `Build-OutputText`

Expected high-level process:

1. Receive drag-and-drop paths through `param`
2. Validate inputs with `Get-Item -LiteralPath`
3. Recursively collect supported files from folders
4. Build file tree data
5. Dispatch fast text and Open XML files in the parent process
6. Dispatch COM-backed files to a worker process with a timeout
7. Store extraction result objects
8. Mark timed-out files as `DEFERRED_TIMEOUT`
9. Build structured output text
10. Write the `.txt` file
11. Ask whether to retry deferred files when needed
12. Show final summary in the console

---

## 17. Completion Criteria

v1 is complete when:

- Drag-and-drop of files onto `.bat` works
- Drag-and-drop of folders onto `.bat` works
- OneDrive / SharePoint synchronized paths work
- Word, Excel, PowerPoint, PDF, and text files are processed
- Modern Office Open XML files use the C# ZIP/XML reader
- Legacy Office files use Office COM
- PDF text is extracted through Word PDF import / OCR
- A single `.txt` file is generated
- A hung file is deferred instead of blocking the whole run
- Deferred retry is selectable by the user
- The output contains file tree, summary, contents, deferred files, and failed files
- The console shows progress and final results
- Script-generated labels are English-only
- Source content is not silently summarized, compressed, or truncated
- Script-generated labels are English-only
