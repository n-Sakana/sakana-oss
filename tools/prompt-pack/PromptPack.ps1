[CmdletBinding()]
param(
    [Parameter(Position = 0, Mandatory = $false)]
    [string[]]$InputPaths,

    [string]$OutFile = "",

    [int]$TimeoutSeconds = 120,

    [int]$RetryTimeoutSeconds = 300,

    [ValidateSet("Ask", "Retry", "Skip")]
    [string]$DeferredAction = "Ask",

    [switch]$Worker,

    [string]$WorkerRequestPath = ""
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$script:TextExtensions = @(
    ".txt", ".md", ".markdown", ".csv", ".tsv",
    ".json", ".xml", ".html", ".htm",
    ".css", ".js", ".jsx", ".ts", ".tsx",
    ".py", ".ps1", ".bat", ".cmd", ".sql",
    ".yaml", ".yml", ".ini", ".log",
    ".config", ".toml", ".vbs", ".bas"
)

$script:WordExtensions = @(".docx", ".doc", ".docm", ".rtf")
$script:ExcelExtensions = @(".xlsx", ".xls", ".xlsm", ".xlsb")
$script:PowerPointExtensions = @(".pptx", ".ppt", ".pptm")
$script:PdfExtensions = @(".pdf")
$script:ExcludedDirectoryNames = @(".git", "node_modules", ".venv", "__pycache__")

$script:WordApp = $null
$script:ExcelApp = $null
$script:PowerPointApp = $null
$script:ScanMessages = New-Object "System.Collections.Generic.List[string]"
$script:ExitCode = 0
$script:WorkerStagePath = ""
$script:WorkerPidPath = ""
$script:OwnedOfficePids = New-Object "System.Collections.Generic.HashSet[int]"
$script:NativeHelperLoaded = $false

function Initialize-NativeHelper {
    if ($script:NativeHelperLoaded) {
        return
    }

    if ("PromptPackNativeHelper" -as [type]) {
        $script:NativeHelperLoaded = $true
        return
    }

    $source = @'
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;

public sealed class PromptPackTextResult
{
    public string Content;
    public string EncodingName;
}

public static class PromptPackNativeHelper
{
    public static PromptPackTextResult ReadTextFile(string path)
    {
        byte[] bytes = File.ReadAllBytes(path);
        PromptPackTextResult result = new PromptPackTextResult();

        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
        {
            result.Content = Encoding.UTF8.GetString(bytes);
            result.EncodingName = "UTF-8 BOM";
            return result;
        }

        if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
        {
            result.Content = Encoding.Unicode.GetString(bytes);
            result.EncodingName = "UTF-16 LE";
            return result;
        }

        if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
        {
            result.Content = Encoding.BigEndianUnicode.GetString(bytes);
            result.EncodingName = "UTF-16 BE";
            return result;
        }

        try
        {
            UTF8Encoding utf8 = new UTF8Encoding(false, true);
            result.Content = utf8.GetString(bytes);
            result.EncodingName = "UTF-8";
            return result;
        }
        catch
        {
        }

        try
        {
            result.Content = Encoding.GetEncoding(932).GetString(bytes);
            result.EncodingName = "CP932";
            return result;
        }
        catch
        {
            result.Content = Encoding.Default.GetString(bytes);
            result.EncodingName = "Default";
            return result;
        }
    }

    public static string ExtractPdfText(string path)
    {
        byte[] bytes = File.ReadAllBytes(path);
        Encoding latin = Encoding.GetEncoding(28591);
        string raw = latin.GetString(bytes);
        StringBuilder output = new StringBuilder();
        int searchFrom = 0;

        while (true)
        {
            int streamMarker = raw.IndexOf("stream", searchFrom, StringComparison.Ordinal);
            if (streamMarker < 0)
            {
                break;
            }

            int dataStart = streamMarker + 6;
            if (dataStart < bytes.Length && raw[dataStart] == '\r')
            {
                dataStart++;
                if (dataStart < bytes.Length && raw[dataStart] == '\n')
                {
                    dataStart++;
                }
            }
            else if (dataStart < bytes.Length && raw[dataStart] == '\n')
            {
                dataStart++;
            }

            int dataEnd = raw.IndexOf("endstream", dataStart, StringComparison.Ordinal);
            if (dataEnd < 0)
            {
                break;
            }

            int dictStart = Math.Max(0, streamMarker - 1200);
            string dictionary = raw.Substring(dictStart, streamMarker - dictStart);
            byte[] streamBytes = Slice(bytes, dataStart, dataEnd - dataStart);
            streamBytes = TrimTrailingLineEndings(streamBytes);

            byte[] decoded = streamBytes;
            if (dictionary.IndexOf("/FlateDecode", StringComparison.Ordinal) >= 0 ||
                dictionary.IndexOf("/Fl", StringComparison.Ordinal) >= 0)
            {
                decoded = TryInflate(streamBytes);
            }

            if (decoded != null && decoded.Length > 0)
            {
                string content = latin.GetString(decoded);
                string text = ExtractTextFromContentStream(content);
                if (!String.IsNullOrWhiteSpace(text))
                {
                    output.AppendLine(text.Trim());
                    output.AppendLine();
                }
            }

            searchFrom = dataEnd + 9;
        }

        return CleanExtractedText(output.ToString());
    }

    private static byte[] Slice(byte[] bytes, int start, int length)
    {
        if (length <= 0)
        {
            return new byte[0];
        }

        byte[] result = new byte[length];
        Buffer.BlockCopy(bytes, start, result, 0, length);
        return result;
    }

    private static byte[] TrimTrailingLineEndings(byte[] bytes)
    {
        int length = bytes.Length;
        while (length > 0 && (bytes[length - 1] == 0x0A || bytes[length - 1] == 0x0D))
        {
            length--;
        }

        if (length == bytes.Length)
        {
            return bytes;
        }

        return Slice(bytes, 0, length);
    }

    private static byte[] TryInflate(byte[] data)
    {
        byte[] result = TryInflateWithOffset(data, 0);
        if (result != null)
        {
            return result;
        }

        if (data.Length > 2)
        {
            result = TryInflateWithOffset(data, 2);
            if (result != null)
            {
                return result;
            }
        }

        return null;
    }

    private static byte[] TryInflateWithOffset(byte[] data, int offset)
    {
        try
        {
            using (MemoryStream input = new MemoryStream(data, offset, data.Length - offset))
            using (DeflateStream deflate = new DeflateStream(input, CompressionMode.Decompress))
            using (MemoryStream output = new MemoryStream())
            {
                CopyStream(deflate, output);
                return output.ToArray();
            }
        }
        catch
        {
            return null;
        }
    }

    private static void CopyStream(Stream input, Stream output)
    {
        byte[] buffer = new byte[8192];
        while (true)
        {
            int read = input.Read(buffer, 0, buffer.Length);
            if (read <= 0)
            {
                break;
            }
            output.Write(buffer, 0, read);
        }
    }

    private static string ExtractTextFromContentStream(string content)
    {
        StringBuilder output = new StringBuilder();
        bool inTextObject = false;
        int index = 0;

        while (index < content.Length)
        {
            if (IsTokenAt(content, index, "BT"))
            {
                inTextObject = true;
                index += 2;
                continue;
            }

            if (IsTokenAt(content, index, "ET"))
            {
                inTextObject = false;
                output.AppendLine();
                index += 2;
                continue;
            }

            if (inTextObject && content[index] == '(')
            {
                int endIndex;
                string value = ReadPdfLiteralString(content, index, out endIndex);
                if (!String.IsNullOrEmpty(value))
                {
                    output.Append(value);
                }
                index = Math.Max(index + 1, endIndex);
                continue;
            }

            if (inTextObject && content[index] == '<' && index + 1 < content.Length && content[index + 1] != '<')
            {
                int endIndex;
                string value = ReadPdfHexString(content, index, out endIndex);
                if (!String.IsNullOrEmpty(value))
                {
                    output.Append(value);
                }
                index = Math.Max(index + 1, endIndex);
                continue;
            }

            if (inTextObject && IsTokenAt(content, index, "T*"))
            {
                output.AppendLine();
                index += 2;
                continue;
            }

            if (inTextObject && IsTokenAt(content, index, "Td"))
            {
                output.AppendLine();
                index += 2;
                continue;
            }

            if (inTextObject && IsTokenAt(content, index, "TD"))
            {
                output.AppendLine();
                index += 2;
                continue;
            }

            index++;
        }

        return output.ToString();
    }

    private static bool IsTokenAt(string text, int index, string token)
    {
        if (index < 0 || index + token.Length > text.Length)
        {
            return false;
        }

        if (String.CompareOrdinal(text, index, token, 0, token.Length) != 0)
        {
            return false;
        }

        bool leftOk = index == 0 || IsDelimiter(text[index - 1]);
        bool rightOk = index + token.Length >= text.Length || IsDelimiter(text[index + token.Length]);
        return leftOk && rightOk;
    }

    private static bool IsDelimiter(char value)
    {
        return Char.IsWhiteSpace(value) || value == '[' || value == ']' || value == '<' || value == '>' || value == '/' || value == '(' || value == ')';
    }

    private static string ReadPdfLiteralString(string text, int start, out int endIndex)
    {
        List<byte> bytes = new List<byte>();
        int depth = 1;
        int index = start + 1;

        while (index < text.Length)
        {
            char current = text[index];
            if (current == '\\')
            {
                index++;
                if (index >= text.Length)
                {
                    break;
                }

                char escaped = text[index];
                if (escaped == 'n') bytes.Add(0x0A);
                else if (escaped == 'r') bytes.Add(0x0D);
                else if (escaped == 't') bytes.Add(0x09);
                else if (escaped == 'b') bytes.Add(0x08);
                else if (escaped == 'f') bytes.Add(0x0C);
                else if (escaped == '(') bytes.Add((byte)'(');
                else if (escaped == ')') bytes.Add((byte)')');
                else if (escaped == '\\') bytes.Add((byte)'\\');
                else if (escaped == '\r')
                {
                    if (index + 1 < text.Length && text[index + 1] == '\n')
                    {
                        index++;
                    }
                }
                else if (escaped == '\n')
                {
                }
                else if (escaped >= '0' && escaped <= '7')
                {
                    int value = escaped - '0';
                    int count = 1;
                    while (count < 3 && index + 1 < text.Length && text[index + 1] >= '0' && text[index + 1] <= '7')
                    {
                        index++;
                        value = (value * 8) + (text[index] - '0');
                        count++;
                    }
                    bytes.Add((byte)(value & 0xFF));
                }
                else
                {
                    bytes.Add((byte)(escaped & 0xFF));
                }
            }
            else if (current == '(')
            {
                depth++;
                bytes.Add((byte)'(');
            }
            else if (current == ')')
            {
                depth--;
                if (depth == 0)
                {
                    index++;
                    break;
                }
                bytes.Add((byte)')');
            }
            else
            {
                bytes.Add((byte)(current & 0xFF));
            }

            index++;
        }

        endIndex = index;
        return DecodePdfBytes(bytes.ToArray());
    }

    private static string ReadPdfHexString(string text, int start, out int endIndex)
    {
        StringBuilder hex = new StringBuilder();
        int index = start + 1;

        while (index < text.Length)
        {
            char current = text[index];
            if (current == '>')
            {
                index++;
                break;
            }

            if (IsHex(current))
            {
                hex.Append(current);
            }

            index++;
        }

        if ((hex.Length % 2) == 1)
        {
            hex.Append('0');
        }

        byte[] bytes = new byte[hex.Length / 2];
        for (int i = 0; i < bytes.Length; i++)
        {
            bytes[i] = Convert.ToByte(hex.ToString(i * 2, 2), 16);
        }

        endIndex = index;
        return DecodePdfBytes(bytes);
    }

    private static bool IsHex(char value)
    {
        return (value >= '0' && value <= '9') ||
            (value >= 'a' && value <= 'f') ||
            (value >= 'A' && value <= 'F');
    }

    private static string DecodePdfBytes(byte[] bytes)
    {
        if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
        {
            return Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2);
        }

        return Encoding.GetEncoding(28591).GetString(bytes);
    }

    private static string CleanExtractedText(string text)
    {
        if (String.IsNullOrWhiteSpace(text))
        {
            return "";
        }

        string[] lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        StringBuilder output = new StringBuilder();
        bool previousBlank = false;

        foreach (string rawLine in lines)
        {
            string line = rawLine.Trim();
            if (line.Length == 0)
            {
                if (!previousBlank)
                {
                    output.AppendLine();
                }
                previousBlank = true;
            }
            else
            {
                output.AppendLine(line);
                previousBlank = false;
            }
        }

        return output.ToString().Trim();
    }
}
'@

    Add-Type -TypeDefinition $source -Language CSharp
    $script:NativeHelperLoaded = $true
}

function Write-Section {
    param([string]$Text)

    Write-Host ""
    Write-Host $Text -ForegroundColor Cyan
    Write-Host (("-" * $Text.Length)) -ForegroundColor DarkCyan
}

function Write-Info {
    param([string]$Text)
    Write-Host ("[INFO] " + $Text) -ForegroundColor Gray
}

function Write-Warn {
    param([string]$Text)
    Write-Host ("[WARN] " + $Text) -ForegroundColor Yellow
}

function Write-Fail {
    param([string]$Text)
    Write-Host ("[FAIL] " + $Text) -ForegroundColor Red
}

function Set-WorkerStage {
    param([string]$Stage)

    if ([string]::IsNullOrWhiteSpace($script:WorkerStagePath)) {
        return
    }

    try {
        $line = "{0}`t{1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff zzz"), $Stage
        Add-Content -LiteralPath $script:WorkerStagePath -Value $line -Encoding UTF8
    }
    catch {
    }
}

function Get-ProcessIdsByName {
    param([string]$Name)

    try {
        return @((Get-Process -Name $Name -ErrorAction SilentlyContinue) | ForEach-Object { [int]$_.Id })
    }
    catch {
        return @()
    }
}

function Add-OwnedOfficePid {
    param([int]$ProcessId)

    if ($ProcessId -le 0) {
        return
    }

    if ($script:OwnedOfficePids.Add($ProcessId)) {
        if (-not [string]::IsNullOrWhiteSpace($script:WorkerPidPath)) {
            try {
                Add-Content -LiteralPath $script:WorkerPidPath -Value ([string]$ProcessId) -Encoding ASCII
            }
            catch {
            }
        }
    }
}

function Register-NewOfficeProcesses {
    param(
        [string]$ProcessName,
        [int[]]$Before
    )

    $after = @(Get-ProcessIdsByName -Name $ProcessName)
    foreach ($processId in $after) {
        if ($Before -notcontains $processId) {
            Add-OwnedOfficePid -ProcessId $processId
        }
    }
}

function Release-ComObjectSafe {
    param([object]$Object)

    if ($null -ne $Object) {
        try {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object)
        }
        catch {
        }
    }
}

function Stop-OfficeApplications {
    if ($null -ne $script:WordApp) {
        try { $script:WordApp.Quit() } catch { }
        Release-ComObjectSafe -Object $script:WordApp
        $script:WordApp = $null
    }

    if ($null -ne $script:ExcelApp) {
        try { $script:ExcelApp.Quit() } catch { }
        Release-ComObjectSafe -Object $script:ExcelApp
        $script:ExcelApp = $null
    }

    if ($null -ne $script:PowerPointApp) {
        try { $script:PowerPointApp.Quit() } catch { }
        Release-ComObjectSafe -Object $script:PowerPointApp
        $script:PowerPointApp = $null
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Stop-OwnedOfficeProcessesFromFile {
    param([string]$PidPath)

    if ([string]::IsNullOrWhiteSpace($PidPath) -or -not (Test-Path -LiteralPath $PidPath -PathType Leaf)) {
        return
    }

    $processIds = @()
    try {
        $processIds = @(Get-Content -LiteralPath $PidPath -ErrorAction Stop | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ } | Sort-Object -Unique)
    }
    catch {
        return
    }

    foreach ($processId in $processIds) {
        try {
            $process = Get-Process -Id $processId -ErrorAction Stop
            Stop-Process -Id $process.Id -Force -ErrorAction Stop
        }
        catch {
        }
    }
}

function Get-WordApplication {
    if ($null -eq $script:WordApp) {
        $before = @(Get-ProcessIdsByName -Name "WINWORD")
        Set-WorkerStage -Stage "START_WORD"
        $script:WordApp = New-Object -ComObject Word.Application
        Register-NewOfficeProcesses -ProcessName "WINWORD" -Before $before
        $script:WordApp.Visible = $false
    }

    return $script:WordApp
}

function Get-ExcelApplication {
    if ($null -eq $script:ExcelApp) {
        $before = @(Get-ProcessIdsByName -Name "EXCEL")
        Set-WorkerStage -Stage "START_EXCEL"
        $script:ExcelApp = New-Object -ComObject Excel.Application
        Register-NewOfficeProcesses -ProcessName "EXCEL" -Before $before
        $script:ExcelApp.Visible = $false
        $script:ExcelApp.DisplayAlerts = $false
        try { $script:ExcelApp.AskToUpdateLinks = $false } catch { }
        try { $script:ExcelApp.AutomationSecurity = 3 } catch { }
    }

    return $script:ExcelApp
}

function Get-PowerPointApplication {
    if ($null -eq $script:PowerPointApp) {
        $before = @(Get-ProcessIdsByName -Name "POWERPNT")
        Set-WorkerStage -Stage "START_POWERPOINT"
        $script:PowerPointApp = New-Object -ComObject PowerPoint.Application
        Register-NewOfficeProcesses -ProcessName "POWERPNT" -Before $before
        try { $script:PowerPointApp.DisplayAlerts = 1 } catch { }
        try { $script:PowerPointApp.AutomationSecurity = 3 } catch { }
    }

    return $script:PowerPointApp
}

function Test-ExcludedDirectory {
    param([System.IO.DirectoryInfo]$Directory)

    if (($Directory.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
        return $true
    }

    foreach ($name in $script:ExcludedDirectoryNames) {
        if ($Directory.Name -ieq $name) {
            return $true
        }
    }

    return $false
}

function Get-SafeChildItems {
    param([string]$DirectoryPath)

    try {
        return @(Get-ChildItem -LiteralPath $DirectoryPath -Force -ErrorAction Stop |
            Sort-Object @{ Expression = { if ($_.PSIsContainer) { 0 } else { 1 } } }, Name)
    }
    catch {
        $script:ScanMessages.Add("Cannot read directory: $DirectoryPath :: $($_.Exception.Message)") | Out-Null
        return @()
    }
}

function Get-FileKind {
    param([string]$Path)

    $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()

    if ($script:TextExtensions -contains $ext) { return "Text" }
    if ($script:WordExtensions -contains $ext) { return "Word" }
    if ($script:ExcelExtensions -contains $ext) { return "Excel" }
    if ($script:PowerPointExtensions -contains $ext) { return "PowerPoint" }
    if ($script:PdfExtensions -contains $ext) { return "PDF" }

    return "Unsupported"
}

function Get-TypeCode {
    param([string]$Path)

    $ext = [System.IO.Path]::GetExtension($Path)
    if ([string]::IsNullOrWhiteSpace($ext)) {
        return "FILE"
    }

    return $ext.TrimStart(".").ToUpperInvariant()
}

function New-ExtractionResult {
    param(
        [int]$No,
        [string]$Status,
        [string]$Type,
        [string]$Path,
        [string]$Method,
        [string]$Notes,
        [string]$Content
    )

    return [pscustomobject]@{
        No = $No
        Status = $Status
        Type = $Type
        TypeCode = Get-TypeCode -Path $Path
        Path = $Path
        Method = $Method
        Notes = $Notes
        Content = $Content
    }
}

function Get-SuccessStatus {
    param(
        [string]$Kind,
        [string]$Method
    )

    if ($Kind -eq "Text") { return "OK_TEXT" }
    if ($Kind -eq "Word") { return "OK_WORD" }
    if ($Kind -eq "Excel") { return "OK_EXCEL" }
    if ($Kind -eq "PowerPoint") { return "OK_POWERPOINT" }
    if ($Kind -eq "PDF" -and $Method -eq "PDF Text Layer") { return "OK_PDF_TEXT" }
    if ($Kind -eq "PDF") { return "OK_WORD_OCR" }

    return "OK"
}

function Read-TextFile {
    param([string]$Path)

    Set-WorkerStage -Stage "READ_TEXT"
    Initialize-NativeHelper
    $data = [PromptPackNativeHelper]::ReadTextFile($Path)

    return [pscustomobject]@{
        Content = $data.Content
        Method = "Text Read"
        Notes = $data.EncodingName
    }
}

function Read-WordLikeFile {
    param(
        [string]$Path,
        [string]$Method,
        [string]$OpenStage = "OPEN_DOCUMENT"
    )

    $word = Get-WordApplication
    $document = $null

    try {
        Set-WorkerStage -Stage $OpenStage
        $document = $word.Documents.Open($Path)
        Set-WorkerStage -Stage "READ_CONTENT"
        $text = $document.Content.Text

        if ($null -eq $text) {
            $text = ""
        }

        $notes = ""
        if ([string]::IsNullOrWhiteSpace($text)) {
            $notes = "No text extracted"
        }

        return [pscustomobject]@{
            Content = $text
            Method = $Method
            Notes = $notes
        }
    }
    finally {
        if ($null -ne $document) {
            Set-WorkerStage -Stage "CLOSE_DOCUMENT"
            try { $document.Close($false) } catch { }
            Release-ComObjectSafe -Object $document
        }
    }
}

function Read-WordFile {
    param([string]$Path)
    return Read-WordLikeFile -Path $Path -Method "Word COM" -OpenStage "OPEN_DOCUMENT"
}

function Read-PdfFileWithWord {
    param([string]$Path)

    $data = Read-WordLikeFile -Path $Path -Method "Word PDF Import" -OpenStage "OPEN_PDF"
    if ([string]::IsNullOrWhiteSpace($data.Notes)) {
        $data.Notes = "Opened directly with Word"
    }
    else {
        $data.Notes = $data.Notes + "; Opened directly with Word"
    }
    return $data
}

function Test-UsableExtractedText {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $false
    }

    $visibleCount = 0
    foreach ($char in $Text.ToCharArray()) {
        if (-not [char]::IsControl($char) -and -not [char]::IsWhiteSpace($char)) {
            $visibleCount++
        }
    }

    return ($visibleCount -ge 4)
}

function Read-PdfFileDirect {
    param([string]$Path)

    Set-WorkerStage -Stage "READ_PDF_TEXT_LAYER"
    Initialize-NativeHelper
    $text = [PromptPackNativeHelper]::ExtractPdfText($Path)

    if (-not (Test-UsableExtractedText -Text $text)) {
        return $null
    }

    return [pscustomobject]@{
        Content = $text
        Method = "PDF Text Layer"
        Notes = "Direct extraction"
    }
}

function Read-PdfFile {
    param([string]$Path)

    $direct = Read-PdfFileDirect -Path $Path
    if ($null -ne $direct) {
        return $direct
    }

    $fallback = Read-PdfFileWithWord -Path $Path
    if ([string]::IsNullOrWhiteSpace($fallback.Notes)) {
        $fallback.Notes = "Fallback after direct extraction found no usable text"
    }
    else {
        $fallback.Notes = $fallback.Notes + "; Fallback after direct extraction found no usable text"
    }
    return $fallback
}

function Format-CellText {
    param([object]$Value)

    if ($null -eq $Value) {
        return ""
    }

    $text = [string]$Value
    $text = $text.Replace("`r`n", " ")
    $text = $text.Replace("`n", " ")
    $text = $text.Replace("`r", " ")
    $text = $text.Replace("`t", " ")
    return $text
}

function Read-ExcelFile {
    param([string]$Path)

    $excel = Get-ExcelApplication
    $workbook = $null
    $builder = New-Object System.Text.StringBuilder

    try {
        Set-WorkerStage -Stage "OPEN_WORKBOOK"
        $workbook = $excel.Workbooks.Open($Path, 0, $true)
        Set-WorkerStage -Stage "READ_WORKBOOK"

        for ($sheetIndex = 1; $sheetIndex -le $workbook.Worksheets.Count; $sheetIndex++) {
            $worksheet = $workbook.Worksheets.Item($sheetIndex)
            [void]$builder.AppendLine("## Sheet: $($worksheet.Name)")

            $usedRange = $worksheet.UsedRange
            $rowCount = $usedRange.Rows.Count
            $columnCount = $usedRange.Columns.Count

            if ($rowCount -eq 1 -and $columnCount -eq 1) {
                $singleCell = $usedRange.Cells.Item(1, 1)
                $singleValue = $singleCell.Text
                Release-ComObjectSafe -Object $singleCell
                if ([string]::IsNullOrWhiteSpace([string]$singleValue)) {
                    [void]$builder.AppendLine("(empty)")
                    [void]$builder.AppendLine("")
                    Release-ComObjectSafe -Object $usedRange
                    Release-ComObjectSafe -Object $worksheet
                    continue
                }
            }

            for ($row = 1; $row -le $rowCount; $row++) {
                $cells = New-Object "System.Collections.Generic.List[string]"

                for ($column = 1; $column -le $columnCount; $column++) {
                    $cell = $usedRange.Cells.Item($row, $column)
                    $cellText = Format-CellText -Value $cell.Text
                    Release-ComObjectSafe -Object $cell
                    $cells.Add($cellText) | Out-Null
                }

                [void]$builder.AppendLine(($cells.ToArray() -join "`t"))
            }

            [void]$builder.AppendLine("")
            Release-ComObjectSafe -Object $usedRange
            Release-ComObjectSafe -Object $worksheet
        }

        return [pscustomobject]@{
            Content = $builder.ToString()
            Method = "Excel COM"
            Notes = "Worksheets: $($workbook.Worksheets.Count)"
        }
    }
    finally {
        if ($null -ne $workbook) {
            Set-WorkerStage -Stage "CLOSE_WORKBOOK"
            try { $workbook.Close($false) } catch { }
            Release-ComObjectSafe -Object $workbook
        }
    }
}

function Append-PowerPointShapeText {
    param(
        [object]$Shape,
        [System.Text.StringBuilder]$Builder
    )

    try {
        if ($Shape.Type -eq 6) {
            for ($i = 1; $i -le $Shape.GroupItems.Count; $i++) {
                $childShape = $Shape.GroupItems.Item($i)
                Append-PowerPointShapeText -Shape $childShape -Builder $Builder
                Release-ComObjectSafe -Object $childShape
            }
        }
    }
    catch {
    }

    try {
        if ($Shape.HasTextFrame -ne 0 -and $Shape.TextFrame.HasText -ne 0) {
            $text = $Shape.TextFrame.TextRange.Text
            if (-not [string]::IsNullOrWhiteSpace($text)) {
                [void]$Builder.AppendLine($text)
            }
        }
    }
    catch {
    }

    try {
        if ($Shape.HasTable -ne 0) {
            $table = $Shape.Table
            for ($row = 1; $row -le $table.Rows.Count; $row++) {
                $cells = New-Object "System.Collections.Generic.List[string]"
                for ($column = 1; $column -le $table.Columns.Count; $column++) {
                    $cellText = $table.Cell($row, $column).Shape.TextFrame.TextRange.Text
                    $cells.Add((Format-CellText -Value $cellText)) | Out-Null
                }
                [void]$Builder.AppendLine(($cells.ToArray() -join "`t"))
            }
        }
    }
    catch {
    }
}

function Read-PowerPointFile {
    param([string]$Path)

    $powerPoint = Get-PowerPointApplication
    $presentation = $null
    $builder = New-Object System.Text.StringBuilder

    try {
        Set-WorkerStage -Stage "OPEN_PRESENTATION"
        $presentation = $powerPoint.Presentations.Open($Path, -1, 0, 0)
        Set-WorkerStage -Stage "READ_PRESENTATION"

        for ($slideIndex = 1; $slideIndex -le $presentation.Slides.Count; $slideIndex++) {
            $slide = $presentation.Slides.Item($slideIndex)
            [void]$builder.AppendLine("## Slide $($slide.SlideIndex)")

            for ($shapeIndex = 1; $shapeIndex -le $slide.Shapes.Count; $shapeIndex++) {
                $shape = $slide.Shapes.Item($shapeIndex)
                Append-PowerPointShapeText -Shape $shape -Builder $builder
                Release-ComObjectSafe -Object $shape
            }

            [void]$builder.AppendLine("")
            Release-ComObjectSafe -Object $slide
        }

        return [pscustomobject]@{
            Content = $builder.ToString()
            Method = "PowerPoint COM"
            Notes = "Slides: $($presentation.Slides.Count)"
        }
    }
    finally {
        if ($null -ne $presentation) {
            Set-WorkerStage -Stage "CLOSE_PRESENTATION"
            try { $presentation.Close() } catch { }
            Release-ComObjectSafe -Object $presentation
        }
    }
}

function Invoke-FileExtraction {
    param(
        [System.IO.FileInfo]$File,
        [int]$No
    )

    $kind = Get-FileKind -Path $File.FullName

    if ($kind -eq "Unsupported") {
        return New-ExtractionResult -No $No -Status "UNSUPPORTED" -Type $kind -Path $File.FullName -Method "Not processed" -Notes "Extension is not supported in v1" -Content ""
    }

    try {
        if ($kind -eq "Text") {
            $data = Read-TextFile -Path $File.FullName
        }
        elseif ($kind -eq "Word") {
            $data = Read-WordFile -Path $File.FullName
        }
        elseif ($kind -eq "Excel") {
            $data = Read-ExcelFile -Path $File.FullName
        }
        elseif ($kind -eq "PowerPoint") {
            $data = Read-PowerPointFile -Path $File.FullName
        }
        elseif ($kind -eq "PDF") {
            $data = Read-PdfFile -Path $File.FullName
        }
        else {
            return New-ExtractionResult -No $No -Status "UNSUPPORTED" -Type $kind -Path $File.FullName -Method "Not processed" -Notes "Extension is not supported in v1" -Content ""
        }

        $status = Get-SuccessStatus -Kind $kind -Method $data.Method
        return New-ExtractionResult -No $No -Status $status -Type $kind -Path $File.FullName -Method $data.Method -Notes $data.Notes -Content $data.Content
    }
    catch {
        return New-ExtractionResult -No $No -Status "FAIL" -Type $kind -Path $File.FullName -Method "$kind extraction" -Notes $_.Exception.Message -Content ""
    }
}

function ConvertTo-JsonFile {
    param(
        [object]$Data,
        [string]$Path
    )

    $json = $Data | ConvertTo-Json -Depth 12
    $utf8NoBom = New-Object System.Text.UTF8Encoding -ArgumentList $false
    [System.IO.File]::WriteAllText($Path, $json, $utf8NoBom)
}

function Invoke-WorkerMode {
    if ([string]::IsNullOrWhiteSpace($WorkerRequestPath)) {
        throw "Worker request path was not provided."
    }

    $request = $null
    $requestJson = [System.IO.File]::ReadAllText($WorkerRequestPath, [System.Text.Encoding]::UTF8)
    $request = $requestJson | ConvertFrom-Json
    $script:WorkerStagePath = [string]$request.StagePath
    $script:WorkerPidPath = [string]$request.PidPath

    try {
        Set-WorkerStage -Stage "START_WORKER"
        $file = Get-Item -LiteralPath ([string]$request.InputPath) -Force -ErrorAction Stop
        $result = Invoke-FileExtraction -File $file -No ([int]$request.No)
        Set-WorkerStage -Stage "WRITE_RESULT"
        ConvertTo-JsonFile -Data $result -Path ([string]$request.ResultPath)
        Set-WorkerStage -Stage "END_WORKER"
    }
    catch {
        $path = ""
        $no = 0
        $resultPath = ""
        try {
            if ($null -ne $request) {
                $path = [string]$request.InputPath
                $no = [int]$request.No
                $resultPath = [string]$request.ResultPath
            }
        }
        catch {
        }

        if ([string]::IsNullOrWhiteSpace($resultPath)) {
            throw
        }

        $result = New-ExtractionResult -No $no -Status "FAIL" -Type (Get-FileKind -Path $path) -Path $path -Method "Worker extraction" -Notes $_.Exception.Message -Content ""
        ConvertTo-JsonFile -Data $result -Path $resultPath
        exit 1
    }
    finally {
        Stop-OfficeApplications
    }
}

function Add-FileIfNew {
    param(
        [System.IO.FileInfo]$File,
        [System.Collections.Generic.List[System.IO.FileInfo]]$Files,
        [System.Collections.Generic.HashSet[string]]$Seen
    )

    $fullPath = [System.IO.Path]::GetFullPath($File.FullName)
    if ($Seen.Add($fullPath)) {
        $Files.Add($File) | Out-Null
    }
}

function Collect-FilesFromDirectory {
    param(
        [string]$DirectoryPath,
        [System.Collections.Generic.List[System.IO.FileInfo]]$Files,
        [System.Collections.Generic.HashSet[string]]$Seen
    )

    foreach ($item in Get-SafeChildItems -DirectoryPath $DirectoryPath) {
        if ($item.PSIsContainer) {
            if (-not (Test-ExcludedDirectory -Directory $item)) {
                Collect-FilesFromDirectory -DirectoryPath $item.FullName -Files $Files -Seen $Seen
            }
        }
        else {
            Add-FileIfNew -File $item -Files $Files -Seen $Seen
        }
    }
}

function Collect-InputFiles {
    param([string[]]$Paths)

    $files = New-Object "System.Collections.Generic.List[System.IO.FileInfo]"
    $seen = New-Object "System.Collections.Generic.HashSet[string]" ([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($path in $Paths) {
        try {
            $item = Get-Item -LiteralPath $path -Force -ErrorAction Stop

            if ($item.PSIsContainer) {
                Collect-FilesFromDirectory -DirectoryPath $item.FullName -Files $files -Seen $seen
            }
            else {
                Add-FileIfNew -File $item -Files $files -Seen $seen
            }
        }
        catch {
            $script:ScanMessages.Add("Cannot read input path: $path :: $($_.Exception.Message)") | Out-Null
        }
    }

    return @($files | Sort-Object FullName)
}

function Append-TreeItem {
    param(
        [object]$Item,
        [string]$Indent,
        [System.Text.StringBuilder]$Builder
    )

    if ($Item.PSIsContainer) {
        if (Test-ExcludedDirectory -Directory $Item) {
            return
        }

        [void]$Builder.AppendLine("$Indent[D] $($Item.Name)/")
        foreach ($child in Get-SafeChildItems -DirectoryPath $Item.FullName) {
            Append-TreeItem -Item $child -Indent ($Indent + "  ") -Builder $Builder
        }
    }
    else {
        [void]$Builder.AppendLine("$Indent[F] $($Item.Name)")
    }
}

function Build-FileTree {
    param([string[]]$Paths)

    $builder = New-Object System.Text.StringBuilder

    foreach ($path in $Paths) {
        try {
            $item = Get-Item -LiteralPath $path -Force -ErrorAction Stop
            Append-TreeItem -Item $item -Indent "" -Builder $builder
        }
        catch {
            [void]$builder.AppendLine("[MISSING] $path")
        }
    }

    return $builder.ToString()
}

function Get-OutputBaseDirectory {
    param([string]$FirstPath)

    $item = Get-Item -LiteralPath $FirstPath -Force -ErrorAction Stop

    if ($item.PSIsContainer) {
        return $item.FullName
    }

    return $item.DirectoryName
}

function ConvertTo-TableCell {
    param([string]$Value)

    if ($null -eq $Value) {
        return ""
    }

    $text = $Value.Replace("|", "\|")
    $text = $text.Replace("`r`n", "<br>")
    $text = $text.Replace("`n", "<br>")
    $text = $text.Replace("`r", "<br>")
    return $text
}

function Write-StatusLine {
    param(
        [int]$Index,
        [int]$Total,
        [object]$Result
    )

    $statusColor = "Gray"
    if ($Result.Status -like "OK*") { $statusColor = "Green" }
    elseif ($Result.Status -eq "FAIL") { $statusColor = "Red" }
    elseif ($Result.Status -eq "UNSUPPORTED") { $statusColor = "DarkYellow" }
    elseif ($Result.Status -eq "DEFERRED_TIMEOUT") { $statusColor = "Yellow" }

    $prefix = "[{0}/{1}] {2,-6} " -f $Index, $Total, $Result.TypeCode
    Write-Host $prefix -NoNewline -ForegroundColor DarkGray
    Write-Host ("{0,-17}" -f $Result.Status) -NoNewline -ForegroundColor $statusColor
    Write-Host (" {0}" -f $Result.Path) -ForegroundColor Gray
}

function Get-LatestWorkerStage {
    param([string]$StagePath)

    if ([string]::IsNullOrWhiteSpace($StagePath) -or -not (Test-Path -LiteralPath $StagePath -PathType Leaf)) {
        return "Unknown"
    }

    try {
        $line = Get-Content -LiteralPath $StagePath -Tail 1 -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($line)) {
            return "Unknown"
        }
        $parts = ([string]$line).Split("`t")
        if ($parts.Count -ge 2) {
            return $parts[1]
        }
        return [string]$line
    }
    catch {
        return "Unknown"
    }
}

function ConvertTo-WindowsCommandLineArgument {
    param([string]$Argument)

    if ($null -eq $Argument) {
        return '""'
    }

    if ($Argument.Length -eq 0) {
        return '""'
    }

    if ($Argument -notmatch '[\s"]') {
        return $Argument
    }

    $builder = New-Object System.Text.StringBuilder
    [void]$builder.Append('"')
    $backslashCount = 0

    foreach ($char in $Argument.ToCharArray()) {
        if ($char -eq '\') {
            $backslashCount++
        }
        elseif ($char -eq '"') {
            [void]$builder.Append('\' * (($backslashCount * 2) + 1))
            [void]$builder.Append('"')
            $backslashCount = 0
        }
        else {
            if ($backslashCount -gt 0) {
                [void]$builder.Append('\' * $backslashCount)
                $backslashCount = 0
            }
            [void]$builder.Append($char)
        }
    }

    if ($backslashCount -gt 0) {
        [void]$builder.Append('\' * ($backslashCount * 2))
    }

    [void]$builder.Append('"')
    return $builder.ToString()
}

function Join-WindowsCommandLine {
    param([string[]]$Arguments)

    return (($Arguments | ForEach-Object { ConvertTo-WindowsCommandLineArgument -Argument $_ }) -join " ")
}

function New-DeferredTimeoutResult {
    param(
        [System.IO.FileInfo]$File,
        [int]$No,
        [int]$TimeoutSeconds,
        [string]$Stage
    )

    $kind = Get-FileKind -Path $File.FullName
    return New-ExtractionResult -No $No -Status "DEFERRED_TIMEOUT" -Type $kind -Path $File.FullName -Method "$kind extraction" -Notes ("Timeout after {0} seconds. Last stage: {1}" -f $TimeoutSeconds, $Stage) -Content ""
}

function Clear-CurrentConsoleLine {
    try {
        $width = [Console]::WindowWidth
        if ($width -lt 20) {
            $width = 120
        }
    }
    catch {
        $width = 120
    }

    Write-Host ("`r" + (" " * ([Math]::Min($width - 1, 180))) + "`r") -NoNewline
}

function Write-WorkerSpinner {
    param(
        [int]$Index,
        [int]$Total,
        [System.IO.FileInfo]$File,
        [string]$Stage,
        [int]$ElapsedSeconds,
        [int]$TimeoutSeconds,
        [string]$Spinner
    )

    $name = $File.Name
    if ($name.Length -gt 48) {
        $name = $name.Substring(0, 45) + "..."
    }

    $line = "[{0}/{1}] {2,-6} Working {3} Stage: {4} Elapsed: {5}s / {6}s  {7}" -f `
        $Index,
        $Total,
        (Get-TypeCode -Path $File.FullName),
        $Spinner,
        $Stage,
        $ElapsedSeconds,
        $TimeoutSeconds,
        $name

    try {
        $width = [Console]::WindowWidth
        if ($width -gt 20 -and $line.Length -gt ($width - 1)) {
            $line = $line.Substring(0, $width - 2)
        }
    }
    catch {
    }

    Write-Host ("`r" + $line) -NoNewline -ForegroundColor DarkCyan
}

function Invoke-FileExtractionWithTimeout {
    param(
        [System.IO.FileInfo]$File,
        [int]$No,
        [int]$TimeoutSeconds,
        [int]$Index = 0,
        [int]$Total = 0
    )

    $jobRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "PromptPack"
    $jobDir = Join-Path -Path $jobRoot -ChildPath ("job_{0}" -f [System.Guid]::NewGuid().ToString("N"))
    [void](New-Item -ItemType Directory -Path $jobDir -Force)

    $requestPath = Join-Path -Path $jobDir -ChildPath "request.json"
    $resultPath = Join-Path -Path $jobDir -ChildPath "result.json"
    $stagePath = Join-Path -Path $jobDir -ChildPath "stage.log"
    $pidPath = Join-Path -Path $jobDir -ChildPath "office-pids.txt"
    $stdoutPath = Join-Path -Path $jobDir -ChildPath "stdout.txt"
    $stderrPath = Join-Path -Path $jobDir -ChildPath "stderr.txt"

    try {
        $request = [pscustomobject]@{
            InputPath = $File.FullName
            No = $No
            ResultPath = $resultPath
            StagePath = $stagePath
            PidPath = $pidPath
        }
        ConvertTo-JsonFile -Data $request -Path $requestPath

        $powerShellPath = Join-Path -Path $PSHOME -ChildPath "powershell.exe"
        if (-not (Test-Path -LiteralPath $powerShellPath -PathType Leaf)) {
            $powerShellPath = "powershell.exe"
        }

        $arguments = @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-STA",
            "-File", $PSCommandPath,
            "-Worker",
            "-WorkerRequestPath", $requestPath
        )

        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = $powerShellPath
        $startInfo.Arguments = Join-WindowsCommandLine -Arguments $arguments
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true
        $startInfo.WorkingDirectory = Split-Path -Parent $PSCommandPath

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $startInfo
        [void]$process.Start()

        $deadline = (Get-Date).AddSeconds([Math]::Max(1, $TimeoutSeconds))
        $startTime = Get-Date
        $spinnerChars = @("|", "/", "-", "\")
        $spinnerIndex = 0
        while (-not $process.HasExited -and (Get-Date) -lt $deadline) {
            if ($Index -gt 0 -and $Total -gt 0) {
                $stage = Get-LatestWorkerStage -StagePath $stagePath
                $elapsedSeconds = [int]((Get-Date) - $startTime).TotalSeconds
                Write-WorkerSpinner -Index $Index -Total $Total -File $File -Stage $stage -ElapsedSeconds $elapsedSeconds -TimeoutSeconds $TimeoutSeconds -Spinner $spinnerChars[$spinnerIndex % $spinnerChars.Count]
                $spinnerIndex++
            }
            Start-Sleep -Milliseconds 250
        }

        if ($Index -gt 0 -and $Total -gt 0) {
            Clear-CurrentConsoleLine
        }

        $completed = $process.HasExited
        if (-not $completed) {
            $stage = Get-LatestWorkerStage -StagePath $stagePath
            Stop-OwnedOfficeProcessesFromFile -PidPath $pidPath
            try { $process.Kill() } catch { }
            try { [void]$process.WaitForExit(5000) } catch { }
            return New-DeferredTimeoutResult -File $File -No $No -TimeoutSeconds $TimeoutSeconds -Stage $stage
        }

        $stdout = $process.StandardOutput.ReadToEnd()
        $stderr = $process.StandardError.ReadToEnd()
        if (-not [string]::IsNullOrWhiteSpace($stdout)) {
            [System.IO.File]::WriteAllText($stdoutPath, $stdout)
        }
        if (-not [string]::IsNullOrWhiteSpace($stderr)) {
            [System.IO.File]::WriteAllText($stderrPath, $stderr)
        }

        if (Test-Path -LiteralPath $resultPath -PathType Leaf) {
            $json = [System.IO.File]::ReadAllText($resultPath, [System.Text.Encoding]::UTF8)
            return ($json | ConvertFrom-Json)
        }

        $kind = Get-FileKind -Path $File.FullName
        $notes = "Worker completed without a result file."
        if (-not [string]::IsNullOrWhiteSpace($stderr)) {
            $notes = $notes + " " + $stderr.Trim()
        }
        return New-ExtractionResult -No $No -Status "FAIL" -Type $kind -Path $File.FullName -Method "$kind extraction" -Notes $notes -Content ""
    }
    catch {
        $kind = Get-FileKind -Path $File.FullName
        return New-ExtractionResult -No $No -Status "FAIL" -Type $kind -Path $File.FullName -Method "$kind extraction" -Notes $_.Exception.Message -Content ""
    }
    finally {
        try { Remove-Item -LiteralPath $jobDir -Recurse -Force -ErrorAction SilentlyContinue } catch { }
    }
}

function Get-ResultCounts {
    param([object[]]$Results)

    return [pscustomobject]@{
        OK = @($Results | Where-Object { $_.Status -like "OK*" }).Count
        Failed = @($Results | Where-Object { $_.Status -eq "FAIL" }).Count
        Unsupported = @($Results | Where-Object { $_.Status -eq "UNSUPPORTED" }).Count
        Deferred = @($Results | Where-Object { $_.Status -eq "DEFERRED_TIMEOUT" }).Count
        Total = $Results.Count
    }
}

function Build-OutputText {
    param(
        [string[]]$InputPaths,
        [string]$FileTree,
        [object[]]$Results,
        [string]$GeneratedAt,
        [string]$OutputPath,
        [int]$TimeoutSeconds,
        [int]$RetryTimeoutSeconds
    )

    $builder = New-Object System.Text.StringBuilder
    $counts = Get-ResultCounts -Results $Results

    [void]$builder.AppendLine("# PromptPack Bundle")
    [void]$builder.AppendLine("")
    [void]$builder.AppendLine("Generated At: $GeneratedAt")
    [void]$builder.AppendLine("Tool: PromptPack.ps1")
    [void]$builder.AppendLine("Output Path: $OutputPath")
    [void]$builder.AppendLine("Main Timeout Seconds: $TimeoutSeconds")
    [void]$builder.AppendLine("Retry Timeout Seconds: $RetryTimeoutSeconds")
    [void]$builder.AppendLine("Total Files: $($counts.Total)")
    [void]$builder.AppendLine("Succeeded: $($counts.OK)")
    [void]$builder.AppendLine("Failed: $($counts.Failed)")
    [void]$builder.AppendLine("Deferred: $($counts.Deferred)")
    [void]$builder.AppendLine("Unsupported: $($counts.Unsupported)")
    [void]$builder.AppendLine("")
    [void]$builder.AppendLine("---")
    [void]$builder.AppendLine("")
    [void]$builder.AppendLine("## Input Paths")
    [void]$builder.AppendLine("")

    foreach ($path in $InputPaths) {
        [void]$builder.AppendLine("- $path")
    }

    [void]$builder.AppendLine("")
    [void]$builder.AppendLine("---")
    [void]$builder.AppendLine("")
    [void]$builder.AppendLine("## File Tree")
    [void]$builder.AppendLine("")
    [void]$builder.AppendLine('```text')
    [void]$builder.Append($FileTree)
    [void]$builder.AppendLine('```')
    [void]$builder.AppendLine("")
    [void]$builder.AppendLine("---")
    [void]$builder.AppendLine("")
    [void]$builder.AppendLine("## Extraction Summary")
    [void]$builder.AppendLine("")
    [void]$builder.AppendLine("| No | Status | Type | Path | Method | Notes |")
    [void]$builder.AppendLine("|---:|---|---|---|---|---|")

    foreach ($result in $Results) {
        [void]$builder.AppendLine(("| {0} | {1} | {2} | {3} | {4} | {5} |" -f `
            $result.No,
            (ConvertTo-TableCell -Value $result.Status),
            (ConvertTo-TableCell -Value $result.Type),
            (ConvertTo-TableCell -Value $result.Path),
            (ConvertTo-TableCell -Value $result.Method),
            (ConvertTo-TableCell -Value $result.Notes)))
    }

    if ($script:ScanMessages.Count -gt 0) {
        [void]$builder.AppendLine("")
        [void]$builder.AppendLine("---")
        [void]$builder.AppendLine("")
        [void]$builder.AppendLine("## Scan Messages")
        [void]$builder.AppendLine("")
        foreach ($message in $script:ScanMessages) {
            [void]$builder.AppendLine("- $message")
        }
    }

    [void]$builder.AppendLine("")
    [void]$builder.AppendLine("---")
    [void]$builder.AppendLine("")
    [void]$builder.AppendLine("## Contents")

    foreach ($result in $Results) {
        [void]$builder.AppendLine("")
        [void]$builder.AppendLine("### File $($result.No)")
        [void]$builder.AppendLine("")
        [void]$builder.AppendLine("Path: $($result.Path)")
        [void]$builder.AppendLine("Type: $($result.Type)")
        [void]$builder.AppendLine("Method: $($result.Method)")
        [void]$builder.AppendLine("Status: $($result.Status)")
        if (-not [string]::IsNullOrWhiteSpace($result.Notes)) {
            [void]$builder.AppendLine("Notes: $($result.Notes)")
        }
        [void]$builder.AppendLine("")

        if ($result.Status -like "OK*") {
            [void]$builder.AppendLine($result.Content)
        }
        elseif ($result.Status -eq "DEFERRED_TIMEOUT") {
            [void]$builder.AppendLine("Extraction was deferred because the file timed out during the main pass.")
        }
        else {
            [void]$builder.AppendLine("No content extracted.")
        }

        [void]$builder.AppendLine("")
        [void]$builder.AppendLine("---")
    }

    [void]$builder.AppendLine("")
    [void]$builder.AppendLine("## Deferred Files")
    [void]$builder.AppendLine("")
    $deferred = @($Results | Where-Object { $_.Status -eq "DEFERRED_TIMEOUT" })
    if ($deferred.Count -eq 0) {
        [void]$builder.AppendLine("None")
    }
    else {
        foreach ($result in $deferred) {
            [void]$builder.AppendLine("### Deferred File $($result.No)")
            [void]$builder.AppendLine("")
            [void]$builder.AppendLine("Path: $($result.Path)")
            [void]$builder.AppendLine("Type: $($result.Type)")
            [void]$builder.AppendLine("Method: $($result.Method)")
            [void]$builder.AppendLine("Reason: $($result.Notes)")
            [void]$builder.AppendLine("")
        }
    }

    [void]$builder.AppendLine("")
    [void]$builder.AppendLine("## Failed Files")
    [void]$builder.AppendLine("")
    $failed = @($Results | Where-Object { $_.Status -eq "FAIL" })
    if ($failed.Count -eq 0) {
        [void]$builder.AppendLine("None")
    }
    else {
        foreach ($result in $failed) {
            [void]$builder.AppendLine("### Failed File $($result.No)")
            [void]$builder.AppendLine("")
            [void]$builder.AppendLine("Path: $($result.Path)")
            [void]$builder.AppendLine("Type: $($result.Type)")
            [void]$builder.AppendLine("Method: $($result.Method)")
            [void]$builder.AppendLine("Reason: $($result.Notes)")
            [void]$builder.AppendLine("")
        }
    }

    [void]$builder.AppendLine("")
    [void]$builder.AppendLine("## Unsupported Files")
    [void]$builder.AppendLine("")
    $unsupported = @($Results | Where-Object { $_.Status -eq "UNSUPPORTED" })
    if ($unsupported.Count -eq 0) {
        [void]$builder.AppendLine("None")
    }
    else {
        foreach ($result in $unsupported) {
            [void]$builder.AppendLine("- $($result.Path)")
        }
    }

    return $builder.ToString()
}

function Write-BundleOutput {
    param(
        [string[]]$InputPaths,
        [string]$FileTree,
        [object[]]$Results,
        [string]$OutputPath,
        [int]$TimeoutSeconds,
        [int]$RetryTimeoutSeconds
    )

    $generatedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss zzz"
    $outputText = Build-OutputText -InputPaths $InputPaths -FileTree $FileTree -Results $Results -GeneratedAt $generatedAt -OutputPath $OutputPath -TimeoutSeconds $TimeoutSeconds -RetryTimeoutSeconds $RetryTimeoutSeconds
    $utf8NoBom = New-Object System.Text.UTF8Encoding -ArgumentList $false
    [System.IO.File]::WriteAllText($OutputPath, $outputText, $utf8NoBom)
}

function Write-Summary {
    param(
        [string]$Title,
        [string]$OutputPath,
        [object[]]$Results
    )

    $counts = Get-ResultCounts -Results $Results

    Write-Section -Text $Title
    Write-Host "Output:" -ForegroundColor Gray
    Write-Host $OutputPath -ForegroundColor Green
    Write-Host ""
    Write-Host "Summary:" -ForegroundColor Gray

    $failColor = "Gray"
    if ($counts.Failed -gt 0) {
        $failColor = "Red"
    }

    $deferredColor = "Gray"
    if ($counts.Deferred -gt 0) {
        $deferredColor = "Yellow"
    }

    Write-Host ("OK:          {0}" -f $counts.OK) -ForegroundColor Green
    Write-Host ("Failed:      {0}" -f $counts.Failed) -ForegroundColor $failColor
    Write-Host ("Deferred:    {0}" -f $counts.Deferred) -ForegroundColor $deferredColor
    Write-Host ("Unsupported: {0}" -f $counts.Unsupported) -ForegroundColor DarkYellow
    Write-Host ("Total:       {0}" -f $counts.Total) -ForegroundColor Gray

    if ($script:ScanMessages.Count -gt 0) {
        Write-Warn -Text ("Scan messages: {0}" -f $script:ScanMessages.Count)
    }
}

function Invoke-OptimizedFileExtraction {
    param(
        [System.IO.FileInfo]$File,
        [int]$No,
        [int]$TimeoutSeconds,
        [int]$Index,
        [int]$Total
    )

    $kind = Get-FileKind -Path $File.FullName

    if ($kind -eq "Unsupported") {
        return New-ExtractionResult -No $No -Status "UNSUPPORTED" -Type $kind -Path $File.FullName -Method "Not processed" -Notes "Extension is not supported in v1" -Content ""
    }

    if ($kind -eq "Text") {
        return Invoke-FileExtraction -File $File -No $No
    }

    if ($kind -eq "PDF") {
        try {
            $direct = Read-PdfFileDirect -Path $File.FullName
            if ($null -ne $direct) {
                return New-ExtractionResult -No $No -Status "OK_PDF_TEXT" -Type $kind -Path $File.FullName -Method $direct.Method -Notes $direct.Notes -Content $direct.Content
            }
        }
        catch {
            Write-Warn -Text ("Direct PDF extraction failed, falling back to Word: {0}" -f $_.Exception.Message)
        }
    }

    return Invoke-FileExtractionWithTimeout -File $File -No $No -TimeoutSeconds $TimeoutSeconds -Index $Index -Total $Total
}

function Write-Phase {
    param([string]$Name)

    Write-Host ""
    Write-Host ("Phase: {0}" -f $Name) -ForegroundColor Cyan
}

function Invoke-MainExtractionPass {
    param(
        [System.IO.FileInfo[]]$Files,
        [int]$TimeoutSeconds
    )

    $results = New-Object "System.Collections.Generic.List[object]"
    $index = 0

    foreach ($file in $Files) {
        $index++
        $percent = [int](($index / [double]$Files.Count) * 100)
        Write-Progress -Activity "PromptPack" -Status ("Extracting {0} of {1}" -f $index, $Files.Count) -PercentComplete $percent

        $result = Invoke-OptimizedFileExtraction -File $file -No $index -TimeoutSeconds $TimeoutSeconds -Index $index -Total $Files.Count
        $results.Add($result) | Out-Null
        Write-StatusLine -Index $index -Total $Files.Count -Result $result
    }

    Write-Progress -Activity "PromptPack" -Completed
    return ,$results
}

function Invoke-DeferredRetryPass {
    param(
        [System.Collections.Generic.List[object]]$Results,
        [int]$TimeoutSeconds
    )

    $deferred = @($Results | Where-Object { $_.Status -eq "DEFERRED_TIMEOUT" })
    if ($deferred.Count -eq 0) {
        return
    }

    Write-Section -Text "Retry Deferred Files"
    $retryIndex = 0
    foreach ($oldResult in $deferred) {
        $retryIndex++
        $percent = [int](($retryIndex / [double]$deferred.Count) * 100)
        Write-Progress -Activity "PromptPack Retry" -Status ("Retrying {0} of {1}" -f $retryIndex, $deferred.Count) -PercentComplete $percent

        try {
            $file = Get-Item -LiteralPath $oldResult.Path -Force -ErrorAction Stop
            $newResult = Invoke-FileExtractionWithTimeout -File $file -No ([int]$oldResult.No) -TimeoutSeconds $TimeoutSeconds -Index $retryIndex -Total $deferred.Count
        }
        catch {
            $newResult = New-ExtractionResult -No ([int]$oldResult.No) -Status "FAIL" -Type $oldResult.Type -Path $oldResult.Path -Method $oldResult.Method -Notes $_.Exception.Message -Content ""
        }

        $Results[[int]$oldResult.No - 1] = $newResult
        Write-StatusLine -Index $retryIndex -Total $deferred.Count -Result $newResult
    }

    Write-Progress -Activity "PromptPack Retry" -Completed
}

function Get-DeferredRetryChoice {
    param([object[]]$Results)

    $deferred = @($Results | Where-Object { $_.Status -eq "DEFERRED_TIMEOUT" })
    if ($deferred.Count -eq 0) {
        return "Skip"
    }

    if ($DeferredAction -eq "Retry") {
        return "Retry"
    }

    if ($DeferredAction -eq "Skip") {
        return "Skip"
    }

    Write-Section -Text "Deferred Files"
    Write-Host ("Deferred files found: {0}" -f $deferred.Count) -ForegroundColor Yellow
    foreach ($result in $deferred) {
        Write-Host ("[{0}] {1}  {2}" -f $result.No, $result.TypeCode, $result.Path) -ForegroundColor Gray
        Write-Host ("    {0}" -f $result.Notes) -ForegroundColor DarkGray
    }

    Write-Host ""
    Write-Host "Retry deferred files now? [R] Retry / [S] Skip (default: S): " -NoNewline -ForegroundColor Cyan
    $choice = Read-Host
    if ($choice -match '^[Rr]') {
        return "Retry"
    }

    return "Skip"
}

function Invoke-PromptPack {
    if ($null -eq $InputPaths -or $InputPaths.Count -eq 0) {
        throw "No input paths were provided."
    }

    Write-Section -Text "PromptPack"
    Write-Info -Text ("Main timeout seconds: {0}" -f $TimeoutSeconds)
    Write-Info -Text ("Retry timeout seconds: {0}" -f $RetryTimeoutSeconds)

    if ([string]::IsNullOrWhiteSpace($OutFile)) {
        $scriptDirectory = Split-Path -Parent $PSCommandPath
        $outputDirectory = Join-Path -Path $scriptDirectory -ChildPath "output"
        $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $OutFile = Join-Path -Path $outputDirectory -ChildPath ("promptpack_{0}.txt" -f $stamp)
    }

    $outputFullPath = [System.IO.Path]::GetFullPath($OutFile)
    $outputDirectory = Split-Path -Parent $outputFullPath
    if (-not [string]::IsNullOrWhiteSpace($outputDirectory) -and -not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $outputDirectory -Force)
    }

    Write-Info -Text ("Output path: {0}" -f $outputFullPath)
    Write-Phase -Name "SCAN"
    Write-Info -Text "Collecting files..."

    $fileTree = Build-FileTree -Paths $InputPaths
    $files = @(Collect-InputFiles -Paths $InputPaths)

    if ($files.Count -eq 0) {
        throw "No files were found."
    }

    Write-Phase -Name "PLAN"
    Write-Info -Text ("Files found: {0}" -f $files.Count)

    Write-Phase -Name "EXTRACT"
    $results = Invoke-MainExtractionPass -Files $files -TimeoutSeconds $TimeoutSeconds

    Write-Phase -Name "WRITE"
    Write-BundleOutput -InputPaths $InputPaths -FileTree $fileTree -Results $results.ToArray() -OutputPath $outputFullPath -TimeoutSeconds $TimeoutSeconds -RetryTimeoutSeconds $RetryTimeoutSeconds
    Write-Summary -Title "Main Pass Completed" -OutputPath $outputFullPath -Results $results.ToArray()

    $retryChoice = Get-DeferredRetryChoice -Results $results.ToArray()
    if ($retryChoice -eq "Retry") {
        Write-Phase -Name "DEFERRED"
        Invoke-DeferredRetryPass -Results $results -TimeoutSeconds $RetryTimeoutSeconds
        Write-Phase -Name "WRITE"
        Write-BundleOutput -InputPaths $InputPaths -FileTree $fileTree -Results $results.ToArray() -OutputPath $outputFullPath -TimeoutSeconds $TimeoutSeconds -RetryTimeoutSeconds $RetryTimeoutSeconds
        Write-Summary -Title "Retry Completed" -OutputPath $outputFullPath -Results $results.ToArray()
    }
    elseif (@($results | Where-Object { $_.Status -eq "DEFERRED_TIMEOUT" }).Count -gt 0) {
        Write-Info -Text "Deferred retry skipped."
    }

    Write-Phase -Name "DONE"
    $counts = Get-ResultCounts -Results $results.ToArray()
    if ($counts.Failed -gt 0 -or $counts.Deferred -gt 0 -or $script:ScanMessages.Count -gt 0) {
        $script:ExitCode = 1
    }
}

try {
    if ($Worker) {
        Invoke-WorkerMode
    }
    else {
        Invoke-PromptPack
    }
}
catch {
    if (-not $Worker) {
        Write-Section -Text "Error"
        Write-Fail -Text $_.Exception.Message
    }
    $script:ExitCode = 1
}
finally {
    Stop-OfficeApplications
}

exit $script:ExitCode
