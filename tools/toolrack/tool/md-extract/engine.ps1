# MD Extract engine -- scan, classify, extract, worker, and output functions.
# Ported from extract-md.ps1; PowerShell 5.1, ASCII source only.

$script:TextExtensions = @(
    ".txt", ".md", ".markdown", ".csv", ".tsv",
    ".json", ".xml", ".html", ".htm",
    ".css", ".js", ".jsx", ".ts", ".tsx",
    ".py", ".ps1", ".bat", ".cmd", ".sql",
    ".yaml", ".yml", ".ini", ".log",
    ".config", ".toml", ".vbs", ".bas"
)
$script:WordOpenXmlExtensions = @(".docx", ".docm")
$script:ExcelOpenXmlExtensions = @(".xlsx", ".xlsm")
$script:PowerPointOpenXmlExtensions = @(".pptx", ".pptm")
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
$script:LastWorkerOwnedOfficeRecords = @()
$script:NativeHelperLoaded = $false
$script:ForceComExtraction = $false
$script:RunLogPath = ""
$script:DeferredAction = "Skip"
$script:WorkerRequestPath = ""
$script:ExtractEntryPath = Join-Path $PSScriptRoot "main.ps1"

function Initialize-NativeHelper {
    if ($script:NativeHelperLoaded) { return }
    if ("ExtractMdNativeHelper" -as [type]) {
        $script:NativeHelperLoaded = $true
        return
    }
    $helperPath = Join-Path $PSScriptRoot "helper.cs"
    Add-Type -Path $helperPath -ReferencedAssemblies @("System.IO.Compression.dll", "System.Xml.dll")
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

function Write-FailLine {
    param([string]$Text)
    Write-Host ("[FAIL] " + $Text) -ForegroundColor Red
}

function Initialize-RunLog {
    param([string]$Path)

    $script:RunLogPath = $Path
    $directory = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path -LiteralPath $directory -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $directory -Force)
    }

    $utf8NoBom = New-Object System.Text.UTF8Encoding -ArgumentList $false
    $header = @(
        "# extract-md Run Log",
        "Started At: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff zzz')",
        "Process Id: $PID",
        "Script Path: $script:ExtractEntryPath",
        ""
    ) -join "`r`n"
    [System.IO.File]::WriteAllText($Path, ($header + "`r`n"), $utf8NoBom)
}

function Write-RunLog {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($script:RunLogPath)) {
        return
    }

    try {
        $utf8NoBom = New-Object System.Text.UTF8Encoding -ArgumentList $false
        $line = "[{0}] {1}`r`n" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff zzz"), $Text
        [System.IO.File]::AppendAllText($script:RunLogPath, $line, $utf8NoBom)
    }
    catch {
    }
}

function Write-RunLogBlock {
    param(
        [string]$Title,
        [string[]]$Lines
    )

    Write-RunLog -Text $Title
    foreach ($line in $Lines) {
        Write-RunLog -Text ("  " + [string]$line)
    }
}

function Write-StageLogToRunLog {
    param(
        [string]$StagePath,
        [string]$Prefix
    )

    if ([string]::IsNullOrWhiteSpace($StagePath) -or -not (Test-Path -LiteralPath $StagePath -PathType Leaf)) {
        Write-RunLog -Text ("{0} stage log missing" -f $Prefix)
        return
    }

    try {
        $lines = @(Get-Content -LiteralPath $StagePath -ErrorAction Stop)
        Write-RunLogBlock -Title ("{0} stage log" -f $Prefix) -Lines $lines
    }
    catch {
        Write-RunLog -Text ("{0} stage log read failed: {1}" -f $Prefix, $_.Exception.Message)
    }
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
    param(
        [int]$ProcessId,
        [string]$ProcessName
    )

    if ($ProcessId -le 0) {
        return
    }

    $process = Get-Process -Id $ProcessId -ErrorAction Stop
    if ($process.ProcessName -ine $ProcessName) {
        throw "Office process identity mismatch: expected $ProcessName, got $($process.ProcessName)."
    }
    $startedUtcTicks = $process.StartTime.ToUniversalTime().Ticks
    if ($script:OwnedOfficePids.Add($ProcessId)) {
        if (-not [string]::IsNullOrWhiteSpace($script:WorkerPidPath)) {
            try {
                $record = "{0}`t{1}`t{2}" -f $ProcessId, $process.ProcessName.ToUpperInvariant(), $startedUtcTicks
                Add-Content -LiteralPath $script:WorkerPidPath -Value $record -Encoding ASCII
            }
            catch {
                try { Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue } catch { }
                throw "Cannot record the owned Office process: $($_.Exception.Message)"
            }
        }
    }
}

function Register-OfficeApplicationProcess {
    param(
        [object]$Application,
        [string]$ProcessName,
        [int[]]$Before
    )

    Initialize-NativeHelper
    [long]$windowHandle = 0
    try { $windowHandle = [long]$Application.Hwnd } catch { }
    if ($windowHandle -le 0) {
        try { $windowHandle = [long]$Application.HWND } catch { }
    }

    [int]$processId = 0
    if ($windowHandle -gt 0) {
        $processId = [ExtractMdNativeHelper]::GetWindowProcessId([IntPtr]$windowHandle)
    }
    if ($processId -le 0) {
        $after = @(Get-ProcessIdsByName -Name $ProcessName)
        $newProcesses = @($after | Where-Object { $Before -notcontains $_ })
        $automationProcesses = New-Object "System.Collections.Generic.List[int]"
        foreach ($candidateId in $newProcesses) {
            try {
                $candidateProcess = Get-CimInstance Win32_Process -Filter ("ProcessId=" + [int]$candidateId) -ErrorAction Stop
                if ([string]$candidateProcess.CommandLine -match '(?i)(/automation|-Embedding)') {
                    $automationProcesses.Add([int]$candidateId)
                }
            } catch { }
        }
        if ($automationProcesses.Count -ne 1) {
            throw "Cannot identify the Office process created for $ProcessName."
        }
        $processId = [int]$automationProcesses[0]
    }
    if ($Before -contains $processId) {
        throw "Office automation reused an existing $ProcessName process; extraction was stopped to protect the open application."
    }
    $processDetails = Get-CimInstance Win32_Process -Filter ("ProcessId=" + $processId) -ErrorAction Stop
    if ([string]$processDetails.CommandLine -notmatch '(?i)(/automation|-Embedding)') {
        throw "The identified $ProcessName process is not an isolated automation process."
    }
    Add-OwnedOfficePid -ProcessId $processId -ProcessName $ProcessName
    return $processId
}

function Read-OwnedOfficeProcessRecords {
    param([string]$PidPath)

    if ([string]::IsNullOrWhiteSpace($PidPath) -or -not (Test-Path -LiteralPath $PidPath -PathType Leaf)) {
        return @()
    }
    $records = New-Object "System.Collections.Generic.List[object]"
    try {
        foreach ($line in @(Get-Content -LiteralPath $PidPath -ErrorAction Stop)) {
            $parts = @(([string]$line) -split "`t")
            [int]$processId = 0
            [long]$startedUtcTicks = 0
            if ($parts.Count -ne 3 -or
                -not [int]::TryParse($parts[0], [ref]$processId) -or $processId -le 0 -or
                $parts[1] -notin @("WINWORD", "EXCEL", "POWERPNT") -or
                -not [long]::TryParse($parts[2], [ref]$startedUtcTicks) -or $startedUtcTicks -le 0) {
                continue
            }
            $records.Add([pscustomobject]@{
                ProcessId = $processId
                ProcessName = $parts[1]
                StartedUtcTicks = $startedUtcTicks
            })
        }
    } catch {
        return @()
    }
    return $records.ToArray()
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

    foreach ($record in @(Read-OwnedOfficeProcessRecords -PidPath $PidPath)) {
        try {
            $process = Get-Process -Id ([int]$record.ProcessId) -ErrorAction Stop
            if ($process.ProcessName.ToUpperInvariant() -cne [string]$record.ProcessName) { continue }
            if ($process.StartTime.ToUniversalTime().Ticks -ne [long]$record.StartedUtcTicks) { continue }
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
        $candidate = New-Object -ComObject Word.Application
        try {
            [void](Register-OfficeApplicationProcess -Application $candidate -ProcessName "WINWORD" -Before $before)
            $candidate.Visible = $false
            $candidate.DisplayAlerts = 0
            try { $candidate.AutomationSecurity = 3 } catch { }
            $script:WordApp = $candidate
        } catch {
            Release-ComObjectSafe -Object $candidate
            throw
        }
    }

    return $script:WordApp
}

function Get-ExcelApplication {
    if ($null -eq $script:ExcelApp) {
        $before = @(Get-ProcessIdsByName -Name "EXCEL")
        Set-WorkerStage -Stage "START_EXCEL"
        $candidate = New-Object -ComObject Excel.Application
        try {
            [void](Register-OfficeApplicationProcess -Application $candidate -ProcessName "EXCEL" -Before $before)
            $candidate.Visible = $false
            $candidate.DisplayAlerts = $false
            try { $candidate.AskToUpdateLinks = $false } catch { }
            try { $candidate.AutomationSecurity = 3 } catch { }
            $script:ExcelApp = $candidate
        } catch {
            Release-ComObjectSafe -Object $candidate
            throw
        }
    }

    return $script:ExcelApp
}

function Get-PowerPointApplication {
    if ($null -eq $script:PowerPointApp) {
        $before = @(Get-ProcessIdsByName -Name "POWERPNT")
        Set-WorkerStage -Stage "START_POWERPOINT"
        $candidate = New-Object -ComObject PowerPoint.Application
        try {
            [void](Register-OfficeApplicationProcess -Application $candidate -ProcessName "POWERPNT" -Before $before)
            try { $candidate.DisplayAlerts = 1 } catch { }
            try { $candidate.AutomationSecurity = 3 } catch { }
            $script:PowerPointApp = $candidate
        } catch {
            Release-ComObjectSafe -Object $candidate
            throw
        }
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

function Get-ExtensionLower {
    param([string]$Path)

    return [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
}

function Test-WordOpenXmlFile {
    param([string]$Path)

    return ($script:WordOpenXmlExtensions -contains (Get-ExtensionLower -Path $Path))
}

function Test-ExcelOpenXmlFile {
    param([string]$Path)

    return ($script:ExcelOpenXmlExtensions -contains (Get-ExtensionLower -Path $Path))
}

function Test-PowerPointOpenXmlFile {
    param([string]$Path)

    return ($script:PowerPointOpenXmlExtensions -contains (Get-ExtensionLower -Path $Path))
}

function Test-OfficeOpenXmlFile {
    param([string]$Path)

    return ((Test-WordOpenXmlFile -Path $Path) -or (Test-ExcelOpenXmlFile -Path $Path) -or (Test-PowerPointOpenXmlFile -Path $Path))
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
    if ($Method -eq "Word OpenXML/C#") { return "OK_DOCX_XML" }
    if ($Method -eq "Excel OpenXML/C#") { return "OK_XLSX_XML" }
    if ($Method -eq "PowerPoint OpenXML/C#") { return "OK_PPTX_XML" }
    if ($Method -eq "Word COM fallback") { return "OK_WORD_COM_FALLBACK" }
    if ($Method -eq "Excel COM fallback") { return "OK_EXCEL_COM_FALLBACK" }
    if ($Method -eq "PowerPoint COM fallback") { return "OK_POWERPOINT_COM_FALLBACK" }
    if ($Kind -eq "Word") { return "OK_WORD_COM" }
    if ($Kind -eq "Excel") { return "OK_EXCEL_COM" }
    if ($Kind -eq "PowerPoint") { return "OK_POWERPOINT_COM" }
    if ($Kind -eq "PDF") { return "OK_PDF_WORD" }

    return "OK"
}

function Read-TextFile {
    param([string]$Path)

    Set-WorkerStage -Stage "READ_TEXT"
    Initialize-NativeHelper
    $data = [ExtractMdNativeHelper]::ReadTextFile($Path)

    return [pscustomobject]@{
        Content = $data.Content
        Method = "Text Read"
        Notes = $data.EncodingName
    }
}

function Read-WordOpenXmlFile {
    param([string]$Path)

    Set-WorkerStage -Stage "READ_DOCX_XML"
    Initialize-NativeHelper
    $data = [ExtractMdNativeHelper]::ReadWordOpenXml($Path)

    return [pscustomobject]@{
        Content = $data.Content
        Method = "Word OpenXML/C#"
        Notes = $data.Notes
    }
}

function Read-ExcelOpenXmlFile {
    param([string]$Path)

    Set-WorkerStage -Stage "READ_XLSX_XML"
    Initialize-NativeHelper
    $data = [ExtractMdNativeHelper]::ReadExcelOpenXml($Path)

    return [pscustomobject]@{
        Content = $data.Content
        Method = "Excel OpenXML/C#"
        Notes = $data.Notes
    }
}

function Read-PowerPointOpenXmlFile {
    param([string]$Path)

    Set-WorkerStage -Stage "READ_PPTX_XML"
    Initialize-NativeHelper
    $data = [ExtractMdNativeHelper]::ReadPowerPointOpenXml($Path)

    return [pscustomobject]@{
        Content = $data.Content
        Method = "PowerPoint OpenXML/C#"
        Notes = $data.Notes
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
        $document = $word.Documents.Open($Path, $false, $true, $false)
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

    if ((-not $script:ForceComExtraction) -and (Test-WordOpenXmlFile -Path $Path)) {
        return Read-WordOpenXmlFile -Path $Path
    }

    $method = "Word COM"
    if ($script:ForceComExtraction -and (Test-WordOpenXmlFile -Path $Path)) {
        $method = "Word COM fallback"
    }

    return Read-WordLikeFile -Path $Path -Method $method -OpenStage "OPEN_DOCUMENT"
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

function Read-PdfFile {
    param([string]$Path)

    return Read-PdfFileWithWord -Path $Path
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

function Read-ExcelFileWithCom {
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

        $method = "Excel COM"
        if ($script:ForceComExtraction -and (Test-ExcelOpenXmlFile -Path $Path)) {
            $method = "Excel COM fallback"
        }

        return [pscustomobject]@{
            Content = $builder.ToString()
            Method = $method
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

function Read-ExcelFile {
    param([string]$Path)

    if ((-not $script:ForceComExtraction) -and (Test-ExcelOpenXmlFile -Path $Path)) {
        return Read-ExcelOpenXmlFile -Path $Path
    }

    return Read-ExcelFileWithCom -Path $Path
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

function Read-PowerPointFileWithCom {
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

        $method = "PowerPoint COM"
        if ($script:ForceComExtraction -and (Test-PowerPointOpenXmlFile -Path $Path)) {
            $method = "PowerPoint COM fallback"
        }

        return [pscustomobject]@{
            Content = $builder.ToString()
            Method = $method
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

function Read-PowerPointFile {
    param([string]$Path)

    if ((-not $script:ForceComExtraction) -and (Test-PowerPointOpenXmlFile -Path $Path)) {
        return Read-PowerPointOpenXmlFile -Path $Path
    }

    return Read-PowerPointFileWithCom -Path $Path
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
    if ([string]::IsNullOrWhiteSpace($script:WorkerRequestPath)) {
        throw "Worker request path was not provided."
    }

    $request = $null
    $requestJson = [System.IO.File]::ReadAllText($script:WorkerRequestPath, [System.Text.Encoding]::UTF8)
    $request = $requestJson | ConvertFrom-Json
    $script:WorkerStagePath = [string]$request.StagePath
    $script:WorkerPidPath = [string]$request.PidPath
    $script:ForceComExtraction = $false
    if ($request.PSObject.Properties.Name -contains "ForceCom") {
        $script:ForceComExtraction = [bool]$request.ForceCom
    }

    try {
        Set-WorkerStage -Stage "START_WORKER"
        $delayMs = 0
        if ([int]::TryParse($env:TOOLRACK_MD_EXTRACT_TEST_DELAY_MS, [ref]$delayMs) -and $delayMs -gt 0) {
            Set-WorkerStage -Stage "TEST_DELAY"
            Start-Sleep -Milliseconds $delayMs
        }
        $file = Get-Item -LiteralPath ([string]$request.InputPath) -Force -ErrorAction Stop
        $result = Invoke-FileExtraction -File $file -No ([int]$request.No)
        Set-WorkerStage -Stage "WRITE_RESULT"
        ConvertTo-JsonFile -Data $result -Path ([string]$request.ResultPath)
        Set-WorkerStage -Stage "END_WORKER"
        $hangAfterResultMs = 0
        if ([int]::TryParse($env:TOOLRACK_MD_EXTRACT_TEST_HANG_AFTER_RESULT_MS, [ref]$hangAfterResultMs) -and $hangAfterResultMs -gt 0) {
            Start-Sleep -Milliseconds $hangAfterResultMs
        }
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
    $method = "$kind extraction"
    if ($kind -eq "Word") { $method = "Word COM" }
    elseif ($kind -eq "Excel") { $method = "Excel COM" }
    elseif ($kind -eq "PowerPoint") { $method = "PowerPoint COM" }
    elseif ($kind -eq "PDF") { $method = "Word PDF Import" }

    return New-ExtractionResult -No $No -Status "DEFERRED_TIMEOUT" -Type $kind -Path $File.FullName -Method $method -Notes ("Timeout after {0} seconds. Last stage: {1}" -f $TimeoutSeconds, $Stage) -Content ""
}

function Invoke-FileExtractionWithTimeout {
    param(
        [System.IO.FileInfo]$File,
        [int]$No,
        [int]$TimeoutSeconds,
        [int]$Index = 0,
        [int]$Total = 0,
        [switch]$ForceCom
    )

    $jobRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "extract-md"
    $jobDir = Join-Path -Path $jobRoot -ChildPath ("job_{0}" -f [System.Guid]::NewGuid().ToString("N"))
    [void](New-Item -ItemType Directory -Path $jobDir -Force)

    $requestPath = Join-Path -Path $jobDir -ChildPath "request.json"
    $resultPath = Join-Path -Path $jobDir -ChildPath "result.json"
    $stagePath = Join-Path -Path $jobDir -ChildPath "stage.log"
    $pidPath = Join-Path -Path $jobDir -ChildPath "office-pids.txt"
    $stdoutPath = Join-Path -Path $jobDir -ChildPath "stdout.txt"
    $stderrPath = Join-Path -Path $jobDir -ChildPath "stderr.txt"

    try {
        Write-RunLog -Text ("WORKER_START no={0} type={1} timeout={2}s force_com={3} path={4}" -f $No, (Get-TypeCode -Path $File.FullName), $TimeoutSeconds, [bool]$ForceCom.IsPresent, $File.FullName)
        Write-RunLog -Text ("WORKER_JOB_DIR no={0} dir={1}" -f $No, $jobDir)

        $request = [pscustomobject]@{
            InputPath = $File.FullName
            No = $No
            ResultPath = $resultPath
            StagePath = $stagePath
            PidPath = $pidPath
            ForceCom = [bool]$ForceCom.IsPresent
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
            "-File", $script:ExtractEntryPath,
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
        $startInfo.WorkingDirectory = Split-Path -Parent $script:ExtractEntryPath

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $startInfo
        [void]$process.Start()

        $deadline = (Get-Date).AddSeconds([Math]::Max(1, $TimeoutSeconds))
        $startTime = Get-Date
        $lastStage = ""
        while (-not $process.HasExited -and (Get-Date) -lt $deadline) {
            $stage = Get-LatestWorkerStage -StagePath $stagePath
            if ($stage -ne $lastStage) {
                $elapsedMs = [int]((Get-Date) - $startTime).TotalMilliseconds
                Write-RunLog -Text ("WORKER_STAGE no={0} stage={1} elapsed_ms={2}" -f $No, $stage, $elapsedMs)
                $lastStage = $stage
            }
            Start-Sleep -Milliseconds 500
        }

        $completed = $process.HasExited
        if (-not $completed) {
            $stage = Get-LatestWorkerStage -StagePath $stagePath
            if ($stage -eq "END_WORKER" -and (Test-Path -LiteralPath $resultPath -PathType Leaf)) {
                $json = [System.IO.File]::ReadAllText($resultPath, [System.Text.Encoding]::UTF8)
                $workerResult = $json | ConvertFrom-Json
                Write-RunLog -Text ("WORKER_RESULT_BEFORE_CLEANUP no={0} status={1} method={2}" -f $No, $workerResult.Status, $workerResult.Method)
                Stop-OwnedOfficeProcessesFromFile -PidPath $pidPath
                try { $process.Kill() } catch { }
                try { [void]$process.WaitForExit(5000) } catch { }
                return $workerResult
            }
            $elapsedSeconds = [int]((Get-Date) - $startTime).TotalSeconds
            Write-RunLog -Text ("WORKER_TIMEOUT no={0} stage={1} elapsed_s={2} timeout_s={3}" -f $No, $stage, $elapsedSeconds, $TimeoutSeconds)
            Write-StageLogToRunLog -StagePath $stagePath -Prefix ("WORKER no={0}" -f $No)
            Stop-OwnedOfficeProcessesFromFile -PidPath $pidPath
            try { $process.Kill() } catch { }
            try { [void]$process.WaitForExit(5000) } catch { }
            return New-DeferredTimeoutResult -File $File -No $No -TimeoutSeconds $TimeoutSeconds -Stage $stage
        }

        $stdout = $process.StandardOutput.ReadToEnd()
        $stderr = $process.StandardError.ReadToEnd()
        Write-RunLog -Text ("WORKER_EXIT no={0} exit_code={1} elapsed_ms={2}" -f $No, $process.ExitCode, [int]((Get-Date) - $startTime).TotalMilliseconds)
        Write-StageLogToRunLog -StagePath $stagePath -Prefix ("WORKER no={0}" -f $No)
        if (-not [string]::IsNullOrWhiteSpace($stdout)) {
            [System.IO.File]::WriteAllText($stdoutPath, $stdout)
            Write-RunLogBlock -Title ("WORKER_STDOUT no={0}" -f $No) -Lines @($stdout -split "`r?`n")
        }
        if (-not [string]::IsNullOrWhiteSpace($stderr)) {
            [System.IO.File]::WriteAllText($stderrPath, $stderr)
            Write-RunLogBlock -Title ("WORKER_STDERR no={0}" -f $No) -Lines @($stderr -split "`r?`n")
        }

        if (Test-Path -LiteralPath $resultPath -PathType Leaf) {
            $json = [System.IO.File]::ReadAllText($resultPath, [System.Text.Encoding]::UTF8)
            $workerResult = $json | ConvertFrom-Json
            Write-RunLog -Text ("WORKER_RESULT no={0} status={1} method={2} notes={3} content_chars={4}" -f $No, $workerResult.Status, $workerResult.Method, $workerResult.Notes, ([string]$workerResult.Content).Length)
            return $workerResult
        }

        $kind = Get-FileKind -Path $File.FullName
        $notes = "Worker completed without a result file."
        if (-not [string]::IsNullOrWhiteSpace($stderr)) {
            $notes = $notes + " " + $stderr.Trim()
        }
        Write-RunLog -Text ("WORKER_NO_RESULT no={0} notes={1}" -f $No, $notes)
        return New-ExtractionResult -No $No -Status "FAIL" -Type $kind -Path $File.FullName -Method "$kind extraction" -Notes $notes -Content ""
    }
    catch {
        $kind = Get-FileKind -Path $File.FullName
        Write-RunLog -Text ("WORKER_EXCEPTION no={0} message={1}" -f $No, $_.Exception.Message)
        return New-ExtractionResult -No $No -Status "FAIL" -Type $kind -Path $File.FullName -Method "$kind extraction" -Notes $_.Exception.Message -Content ""
    }
    finally {
        $script:LastWorkerOwnedOfficeRecords = @(Read-OwnedOfficeProcessRecords -PidPath $pidPath)
        Stop-OwnedOfficeProcessesFromFile -PidPath $pidPath
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

    [void]$builder.AppendLine("# Extract Bundle")
    [void]$builder.AppendLine("")
    [void]$builder.AppendLine("Generated At: $GeneratedAt")
    [void]$builder.AppendLine("Tool: extract-md.ps1")
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
    [void]$builder.AppendLine("## Contents")

    $successful = @($Results | Where-Object { $_.Status -like "OK*" })
    if ($successful.Count -eq 0) {
        [void]$builder.AppendLine("")
        [void]$builder.AppendLine("No content extracted.")
    }
    else {
        foreach ($result in $successful) {
            [void]$builder.AppendLine("")
            [void]$builder.AppendLine("### File $($result.No)")
            [void]$builder.AppendLine("")
            [void]$builder.AppendLine("Path: $($result.Path)")
            [void]$builder.AppendLine("Type: $($result.Type)")
            [void]$builder.AppendLine("")
            [void]$builder.AppendLine($result.Content)
            [void]$builder.AppendLine("")
            [void]$builder.AppendLine("---")
        }
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
        Write-RunLog -Text ("ROUTE no={0} route=text_direct" -f $No)
        return Invoke-FileExtraction -File $File -No $No
    }

    if (Test-OfficeOpenXmlFile -Path $File.FullName) {
        Write-RunLog -Text ("ROUTE no={0} route=openxml_direct" -f $No)
        $xmlResult = Invoke-FileExtraction -File $File -No $No
        if ($xmlResult.Status -ne "FAIL") {
            return $xmlResult
        }

        Write-RunLog -Text ("OPENXML_FALLBACK no={0} status={1} notes={2}" -f $No, $xmlResult.Status, $xmlResult.Notes)
        $fallbackResult = Invoke-FileExtractionWithTimeout -File $File -No $No -TimeoutSeconds $TimeoutSeconds -Index $Index -Total $Total -ForceCom
        if ([string]::IsNullOrWhiteSpace([string]$fallbackResult.Notes)) {
            $fallbackResult.Notes = "OpenXML failed: $($xmlResult.Notes)"
        }
        else {
            $fallbackResult.Notes = "$($fallbackResult.Notes); OpenXML failed: $($xmlResult.Notes)"
        }
        return $fallbackResult
    }

    Write-RunLog -Text ("ROUTE no={0} route=worker_com_or_pdf" -f $No)
    return Invoke-FileExtractionWithTimeout -File $File -No $No -TimeoutSeconds $TimeoutSeconds -Index $Index -Total $Total
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
        Write-RunLog -Text ("FILE_START no={0} total={1} type={2} size={3} path={4}" -f $index, $Files.Count, (Get-TypeCode -Path $file.FullName), $file.Length, $file.FullName)
        $started = Get-Date

        $result = Invoke-OptimizedFileExtraction -File $file -No $index -TimeoutSeconds $TimeoutSeconds -Index $index -Total $Files.Count
        $results.Add($result) | Out-Null

        $elapsedMs = [int]((Get-Date) - $started).TotalMilliseconds
        Write-RunLog -Text ("FILE_END no={0} status={1} method={2} notes={3} content_chars={4} elapsed_ms={5}" -f $index, $result.Status, $result.Method, $result.Notes, ([string]$result.Content).Length, $elapsedMs)
    }

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

    Write-RunLog -Text ("RETRY_START count={0} timeout={1}s" -f $deferred.Count, $TimeoutSeconds)
    $retryIndex = 0
    foreach ($oldResult in $deferred) {
        $retryIndex++
        Write-RunLog -Text ("RETRY_FILE_START retry_no={0} original_no={1} path={2}" -f $retryIndex, $oldResult.No, $oldResult.Path)
        $started = Get-Date

        try {
            $file = Get-Item -LiteralPath $oldResult.Path -Force -ErrorAction Stop
            $newResult = Invoke-FileExtractionWithTimeout -File $file -No ([int]$oldResult.No) -TimeoutSeconds $TimeoutSeconds -Index $retryIndex -Total $deferred.Count
        }
        catch {
            $newResult = New-ExtractionResult -No ([int]$oldResult.No) -Status "FAIL" -Type $oldResult.Type -Path $oldResult.Path -Method $oldResult.Method -Notes $_.Exception.Message -Content ""
        }

        $Results[[int]$oldResult.No - 1] = $newResult
        Write-RunLog -Text ("RETRY_FILE_END retry_no={0} original_no={1} status={2} method={3} notes={4} content_chars={5} elapsed_ms={6}" -f $retryIndex, $oldResult.No, $newResult.Status, $newResult.Method, $newResult.Notes, ([string]$newResult.Content).Length, [int]((Get-Date) - $started).TotalMilliseconds)
    }
    Write-RunLog -Text "RETRY_END"
}

function Get-DeferredRetryChoice {
    param([object[]]$Results)

    $deferred = @($Results | Where-Object { $_.Status -eq "DEFERRED_TIMEOUT" })
    if ($deferred.Count -eq 0) {
        return "Skip"
    }

    Write-RunLog -Text ("DEFERRED_FOUND count={0}" -f $deferred.Count)
    foreach ($result in $deferred) {
        Write-RunLog -Text ("DEFERRED_ITEM no={0} type={1} path={2} notes={3}" -f $result.No, $result.TypeCode, $result.Path, $result.Notes)
    }

    if ($script:DeferredAction -eq "Retry") {
        Write-RunLog -Text "DEFERRED_ACTION preset=Retry"
        return "Retry"
    }

    if ($script:DeferredAction -eq "Skip") {
        Write-RunLog -Text "DEFERRED_ACTION preset=Skip"
        return "Skip"
    }

    Write-Host ""
    Write-Host ("Deferred files found: {0}" -f $deferred.Count) -ForegroundColor Yellow
    Write-Host "See the log file for details." -ForegroundColor Gray
    Write-Host "Retry deferred files now? [R] Retry / [S] Skip (default: S): " -NoNewline -ForegroundColor Cyan
    $choice = Read-Host
    if ($choice -match '^[Rr]') {
        Write-RunLog -Text "DEFERRED_ACTION user=Retry"
        return "Retry"
    }

    Write-RunLog -Text "DEFERRED_ACTION user=Skip"
    return "Skip"
}
