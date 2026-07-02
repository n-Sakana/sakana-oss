[CmdletBinding()]
param(
    [Parameter(Position = 0, Mandatory = $false)]
    [string[]]$InputPaths,

    [string]$OutFile = ""
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

function Get-WordApplication {
    if ($null -eq $script:WordApp) {
        $script:WordApp = New-Object -ComObject Word.Application
        $script:WordApp.Visible = $false
        $script:WordApp.DisplayAlerts = 0
        try { $script:WordApp.AutomationSecurity = 3 } catch { }
    }

    return $script:WordApp
}

function Get-ExcelApplication {
    if ($null -eq $script:ExcelApp) {
        $script:ExcelApp = New-Object -ComObject Excel.Application
        $script:ExcelApp.Visible = $false
        $script:ExcelApp.DisplayAlerts = $false
        try { $script:ExcelApp.AskToUpdateLinks = $false } catch { }
        try { $script:ExcelApp.AutomationSecurity = 3 } catch { }
    }

    return $script:ExcelApp
}

function Get-PowerPointApplication {
    if ($null -eq $script:PowerPointApp) {
        $script:PowerPointApp = New-Object -ComObject PowerPoint.Application
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

function Read-TextFile {
    param([string]$Path)

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $encodingName = "UTF-8"
    $text = $null

    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        $encodingName = "UTF-8 BOM"
        $text = [System.Text.Encoding]::UTF8.GetString($bytes)
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        $encodingName = "UTF-16 LE"
        $text = [System.Text.Encoding]::Unicode.GetString($bytes)
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
        $encodingName = "UTF-16 BE"
        $text = [System.Text.Encoding]::BigEndianUnicode.GetString($bytes)
    }
    else {
        $utf8Strict = New-Object System.Text.UTF8Encoding -ArgumentList $false, $true
        try {
            $encodingName = "UTF-8"
            $text = $utf8Strict.GetString($bytes)
        }
        catch {
            try {
                $encodingName = "CP932"
                $text = [System.Text.Encoding]::GetEncoding(932).GetString($bytes)
            }
            catch {
                $encodingName = "Default"
                $text = [System.Text.Encoding]::Default.GetString($bytes)
            }
        }
    }

    return [pscustomobject]@{
        Content = $text
        Method = "Text Read"
        Notes = $encodingName
    }
}

function Read-WordLikeFile {
    param(
        [string]$Path,
        [string]$Method
    )

    $word = Get-WordApplication
    $document = $null

    try {
        $document = $word.Documents.Open($Path, $false, $true, $false)
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
            try { $document.Close($false) } catch { }
            Release-ComObjectSafe -Object $document
        }
    }
}

function Read-WordFile {
    param([string]$Path)
    return Read-WordLikeFile -Path $Path -Method "Word COM"
}

function Read-PdfFileWithWord {
    param([string]$Path)
    return Read-WordLikeFile -Path $Path -Method "Word PDF Import"
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
        $workbook = $excel.Workbooks.Open($Path, 0, $true)

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
        $presentation = $powerPoint.Presentations.Open($Path, -1, 0, 0)

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
            $data = Read-PdfFileWithWord -Path $File.FullName
        }
        else {
            return New-ExtractionResult -No $No -Status "UNSUPPORTED" -Type $kind -Path $File.FullName -Method "Not processed" -Notes "Extension is not supported in v1" -Content ""
        }

        return New-ExtractionResult -No $No -Status "OK" -Type $kind -Path $File.FullName -Method $data.Method -Notes $data.Notes -Content $data.Content
    }
    catch {
        return New-ExtractionResult -No $No -Status "FAIL" -Type $kind -Path $File.FullName -Method "$kind extraction" -Notes $_.Exception.Message -Content ""
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
    if ($Result.Status -eq "OK") { $statusColor = "Green" }
    elseif ($Result.Status -eq "FAIL") { $statusColor = "Red" }
    elseif ($Result.Status -eq "UNSUPPORTED") { $statusColor = "DarkYellow" }

    $prefix = "[{0}/{1}] {2,-6} " -f $Index, $Total, $Result.TypeCode
    Write-Host $prefix -NoNewline -ForegroundColor DarkGray
    Write-Host ("{0,-11}" -f $Result.Status) -NoNewline -ForegroundColor $statusColor
    Write-Host (" {0}" -f $Result.Path) -ForegroundColor Gray
}

function Build-OutputText {
    param(
        [string[]]$InputPaths,
        [string]$FileTree,
        [object[]]$Results,
        [string]$GeneratedAt,
        [string]$OutputPath
    )

    $builder = New-Object System.Text.StringBuilder
    $okCount = @($Results | Where-Object { $_.Status -eq "OK" }).Count
    $failCount = @($Results | Where-Object { $_.Status -eq "FAIL" }).Count
    $unsupportedCount = @($Results | Where-Object { $_.Status -eq "UNSUPPORTED" }).Count

    [void]$builder.AppendLine("# PromptPack Bundle")
    [void]$builder.AppendLine("")
    [void]$builder.AppendLine("Generated At: $GeneratedAt")
    [void]$builder.AppendLine("Tool: PromptPack.ps1")
    [void]$builder.AppendLine("Output Path: $OutputPath")
    [void]$builder.AppendLine("Total Files: $($Results.Count)")
    [void]$builder.AppendLine("Succeeded: $okCount")
    [void]$builder.AppendLine("Failed: $failCount")
    [void]$builder.AppendLine("Unsupported: $unsupportedCount")
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

        if ($result.Status -eq "OK") {
            [void]$builder.AppendLine($result.Content)
        }
        else {
            [void]$builder.AppendLine("No content extracted.")
        }

        [void]$builder.AppendLine("")
        [void]$builder.AppendLine("---")
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

function Invoke-PromptPack {
    if ($null -eq $InputPaths -or $InputPaths.Count -eq 0) {
        throw "No input paths were provided."
    }

    Write-Section -Text "PromptPack"
    Write-Info -Text "Collecting files..."

    $baseDirectory = Get-OutputBaseDirectory -FirstPath $InputPaths[0]
    $fileTree = Build-FileTree -Paths $InputPaths
    $files = @(Collect-InputFiles -Paths $InputPaths)

    if ($files.Count -eq 0) {
        throw "No files were found."
    }

    Write-Info -Text ("Files found: {0}" -f $files.Count)

    $results = New-Object "System.Collections.Generic.List[object]"
    $index = 0

    foreach ($file in $files) {
        $index++
        $percent = [int](($index / [double]$files.Count) * 100)
        Write-Progress -Activity "PromptPack" -Status ("Processing {0} of {1}" -f $index, $files.Count) -PercentComplete $percent

        $result = Invoke-FileExtraction -File $file -No $index
        $results.Add($result) | Out-Null
        Write-StatusLine -Index $index -Total $files.Count -Result $result
    }

    Write-Progress -Activity "PromptPack" -Completed

    if ([string]::IsNullOrWhiteSpace($OutFile)) {
        $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $OutFile = Join-Path -Path $baseDirectory -ChildPath ("prompt-pack_{0}.txt" -f $stamp)
    }

    $outputFullPath = [System.IO.Path]::GetFullPath($OutFile)
    $outputDirectory = Split-Path -Parent $outputFullPath
    if (-not [string]::IsNullOrWhiteSpace($outputDirectory) -and -not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
        [void](New-Item -ItemType Directory -Path $outputDirectory -Force)
    }

    $generatedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss zzz"
    $outputText = Build-OutputText -InputPaths $InputPaths -FileTree $fileTree -Results $results.ToArray() -GeneratedAt $generatedAt -OutputPath $outputFullPath
    $utf8NoBom = New-Object System.Text.UTF8Encoding -ArgumentList $false
    [System.IO.File]::WriteAllText($outputFullPath, $outputText, $utf8NoBom)

    $okCount = @($results | Where-Object { $_.Status -eq "OK" }).Count
    $failCount = @($results | Where-Object { $_.Status -eq "FAIL" }).Count
    $unsupportedCount = @($results | Where-Object { $_.Status -eq "UNSUPPORTED" }).Count

    Write-Section -Text "Done"
    Write-Host "Output:" -ForegroundColor Gray
    Write-Host $outputFullPath -ForegroundColor Green
    Write-Host ""
    Write-Host "Summary:" -ForegroundColor Gray
    $failColor = "Gray"
    if ($failCount -gt 0) {
        $failColor = "Red"
    }

    Write-Host ("OK:          {0}" -f $okCount) -ForegroundColor Green
    Write-Host ("Failed:      {0}" -f $failCount) -ForegroundColor $failColor
    Write-Host ("Unsupported: {0}" -f $unsupportedCount) -ForegroundColor DarkYellow
    Write-Host ("Total:       {0}" -f $results.Count) -ForegroundColor Gray

    if ($script:ScanMessages.Count -gt 0) {
        Write-Warn -Text ("Scan messages: {0}" -f $script:ScanMessages.Count)
    }

    if ($failCount -gt 0 -or $script:ScanMessages.Count -gt 0) {
        $script:ExitCode = 1
    }
}

try {
    Invoke-PromptPack
}
catch {
    Write-Section -Text "Error"
    Write-Fail -Text $_.Exception.Message
    $script:ExitCode = 1
}
finally {
    Stop-OfficeApplications
}

exit $script:ExitCode
