# test/test-md-extract.ps1 -- text, OpenXML, timeout, Deferred, and bundle tests
. (Join-Path $PSScriptRoot "_assert.ps1")
$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$toolDir = Join-Path $root "tool\md-extract"
$main = Join-Path $toolDir "main.ps1"
$engine = Join-Path $toolDir "engine.ps1"
$helper = Join-Path $toolDir "helper.cs"
$manifest = Join-Path $toolDir "tool.json"

Assert-True (Test-Path -LiteralPath $main -PathType Leaf) "md-extract main.ps1 exists"
Assert-True (Test-Path -LiteralPath $engine -PathType Leaf) "md-extract engine.ps1 exists"
Assert-True (Test-Path -LiteralPath $helper -PathType Leaf) "md-extract helper.cs exists"
Assert-True (Test-Path -LiteralPath $manifest -PathType Leaf) "md-extract tool.json exists"
if (-not (Test-Path -LiteralPath $main) -or -not (Test-Path -LiteralPath $engine) -or -not (Test-Path -LiteralPath $helper) -or -not (Test-Path -LiteralPath $manifest)) { Exit-Test }
. $engine
$script:ExtractEntryPath = $main

$fx = Join-Path $env:TEMP ("toolrack_me_" + [guid]::NewGuid().ToString("N"))
$input = Join-Path $fx "input"
New-Item -ItemType Directory -Force $input | Out-Null

function Write-ExtractBytes {
    param([string]$Path, [byte[]]$Bytes)
    $parent = Split-Path $Path -Parent
    if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Force $parent | Out-Null }
    [IO.File]::WriteAllBytes($Path, $Bytes)
}

function Join-ExtractBytes {
    param([byte[]]$First, [byte[]]$Second)
    $result = New-Object byte[] ($First.Length + $Second.Length)
    [Array]::Copy($First, 0, $result, 0, $First.Length)
    [Array]::Copy($Second, 0, $result, $First.Length, $Second.Length)
    return $result
}

function New-OpenXmlFixture {
    param([string]$Path, [hashtable]$Entries)
    Add-Type -AssemblyName System.IO.Compression
    $stream = [IO.File]::Open($Path, [IO.FileMode]::Create, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    $archive = New-Object IO.Compression.ZipArchive($stream, [IO.Compression.ZipArchiveMode]::Create, $false)
    try {
        $encoding = New-Object Text.UTF8Encoding($false)
        foreach ($name in @($Entries.Keys | Sort-Object)) {
            $entry = $archive.CreateEntry($name)
            $writer = New-Object IO.StreamWriter($entry.Open(), $encoding)
            try { $writer.Write([string]$Entries[$name]) } finally { $writer.Dispose() }
        }
    } finally {
        $archive.Dispose()
        $stream.Dispose()
    }
}

function Remove-NewExtractArtifacts {
    param([object[]]$BeforeOutput, [object[]]$BeforeLog)
    $outputDir = Join-Path $root "output"
    $logDir = Join-Path $root "log"
    $afterOutput = @(Get-ChildItem $outputDir -File -Filter "extract-md_*.md" -ErrorAction SilentlyContinue)
    $afterLog = @(Get-ChildItem $logDir -File -Filter "extract-md_*.log" -ErrorAction SilentlyContinue)
    foreach ($file in @($afterOutput | Where-Object { $BeforeOutput.Name -notcontains $_.Name })) { Remove-Item -LiteralPath $file.FullName -Force }
    foreach ($file in @($afterLog | Where-Object { $BeforeLog.Name -notcontains $_.Name })) { Remove-Item -LiteralPath $file.FullName -Force }
}

function Invoke-TestComWorker {
    param([string]$Path, [string]$ExpectedText, [string]$Label, [switch]$ForceCom, [int]$TimeoutSeconds = 30, [switch]$AllowDeferred)
    if ($ForceCom) { $result = Invoke-FileExtractionWithTimeout (Get-Item $Path) 1 $TimeoutSeconds -ForceCom }
    else { $result = Invoke-FileExtractionWithTimeout (Get-Item $Path) 1 $TimeoutSeconds }
    $ownedRecords = @($script:LastWorkerOwnedOfficeRecords)
    $ownedStopped = ($ownedRecords.Count -eq 1)
    foreach ($record in $ownedRecords) {
        for ($waitOwned = 0; $waitOwned -lt 20; $waitOwned++) {
            $ownedProcess = Get-Process -Id ([int]$record.ProcessId) -ErrorAction SilentlyContinue
            if ($null -eq $ownedProcess -or
                $ownedProcess.ProcessName.ToUpperInvariant() -cne [string]$record.ProcessName -or
                $ownedProcess.StartTime.ToUniversalTime().Ticks -ne [long]$record.StartedUtcTicks) {
                break
            }
            Start-Sleep -Milliseconds 100
        }
        if ($null -ne $ownedProcess -and
            $ownedProcess.ProcessName.ToUpperInvariant() -ceq [string]$record.ProcessName -and
            $ownedProcess.StartTime.ToUniversalTime().Ticks -eq [long]$record.StartedUtcTicks) {
            $ownedStopped = $false
        }
    }
    Assert-True $ownedStopped ($Label + " stops only its recorded Office process")
    if ($AllowDeferred -and $result.Status -eq "DEFERRED_TIMEOUT") {
        Assert-True ($result.Notes -like "*OPEN_PDF*") ($Label + " COM route reaches the bounded PDF import stage")
        return
    }
    if ($result.Status -notlike "OK*") { Write-Host ("  COM detail: {0}: {1}: {2}" -f $Label, $result.Status, $result.Notes) -ForegroundColor Yellow }
    Assert-True ($result.Status -like "OK*") ($Label + " COM extraction succeeds")
    Assert-True ([string]$result.Content -like ("*" + $ExpectedText + "*")) ($Label + " COM content is present")
}

function New-LegacyOfficeFixture {
    param([string]$Kind, [string]$Path, [string]$Text)
    $ownedRecords = @()
    $ownedRecordPath = Join-Path (Split-Path $Path -Parent) (".office-owner-" + [guid]::NewGuid().ToString("N") + ".txt")
    $job = Start-Job -ArgumentList $Kind, $Path, $Text, $ownedRecordPath -ScriptBlock {
        param($fixtureKind, $fixturePath, $fixtureText, $fixtureOwnerPath)
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class ToolRackFixtureWindowProcess {
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    public static int GetProcessId(IntPtr window) {
        uint processId;
        GetWindowThreadProcessId(window, out processId);
        return processId > Int32.MaxValue ? 0 : (int)processId;
    }
}
'@
        $ownedRecord = ""
        function Set-FixtureOwnedOffice {
            param([object]$Application, [string]$Name, [int[]]$Before)
            [long]$handle = 0
            try { $handle = [long]$Application.Hwnd } catch { }
            if ($handle -le 0) { try { $handle = [long]$Application.HWND } catch { } }
            $ownedPid = [ToolRackFixtureWindowProcess]::GetProcessId([IntPtr]$handle)
            if ($ownedPid -le 0) {
                $after = @(Get-Process -Name $Name -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
                $newProcesses = @($after | Where-Object { $Before -notcontains $_ })
                if ($newProcesses.Count -ne 1) { throw "Cannot identify fixture Office process." }
                $ownedPid = [int]$newProcesses[0]
            }
            if ($Before -contains $ownedPid) { throw "Fixture Office automation reused an existing process." }
            $ownedProcess = Get-Process -Id $ownedPid -ErrorAction Stop
            if ($ownedProcess.ProcessName -ine $Name) { throw "Fixture Office process name mismatch." }
            $script:ownedRecord = "OWNED|{0}|{1}|{2}" -f $ownedPid, $ownedProcess.ProcessName.ToUpperInvariant(), $ownedProcess.StartTime.ToUniversalTime().Ticks
            [IO.File]::WriteAllText($fixtureOwnerPath, $script:ownedRecord, (New-Object Text.UTF8Encoding($false)))
        }
        $app = $null; $document = $null; $workbook = $null; $presentation = $null
        $expectedName = if ($fixtureKind -in @("word", "pdf")) { "WINWORD" } elseif ($fixtureKind -eq "excel") { "EXCEL" } else { "POWERPNT" }
        $beforeIds = @(Get-Process -Name $expectedName -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
        try {
            if ($fixtureKind -eq "word" -or $fixtureKind -eq "pdf") {
                $app = New-Object -ComObject Word.Application
                $app.Visible = $false
                $app.DisplayAlerts = 0
                $document = $app.Documents.Add()
                Set-FixtureOwnedOffice $app "WINWORD" $beforeIds
                $document.Content.Text = $fixtureText
                if ($fixtureKind -eq "word") { $document.SaveAs2($fixturePath, 0) }
                else { $document.ExportAsFixedFormat($fixturePath, 17) }
                $document.Close($false)
                $document = $null
                $app.Quit()
                $app = $null
            } elseif ($fixtureKind -eq "excel") {
                $app = New-Object -ComObject Excel.Application
                $app.Visible = $false
                $app.DisplayAlerts = $false
                $workbook = $app.Workbooks.Add()
                Set-FixtureOwnedOffice $app "EXCEL" $beforeIds
                $workbook.Worksheets.Item(1).Cells.Item(1, 1).Value2 = $fixtureText
                $workbook.SaveAs($fixturePath, 56)
                $workbook.Close($false)
                $workbook = $null
                $app.Quit()
                $app = $null
            } else {
                $app = New-Object -ComObject PowerPoint.Application
                $presentation = $app.Presentations.Add()
                Set-FixtureOwnedOffice $app "POWERPNT" $beforeIds
                $slide = $presentation.Slides.Add(1, 12)
                $shape = $slide.Shapes.AddTextbox(1, 10, 10, 400, 50)
                $shape.TextFrame.TextRange.Text = $fixtureText
                $presentation.SaveAs($fixturePath, 1)
                $presentation.Close()
                $presentation = $null
                $app.Quit()
                $app = $null
            }
        } finally {
            if ($null -ne $document) { try { $document.Close($false) } catch { } }
            if ($null -ne $workbook) { try { $workbook.Close($false) } catch { } }
            if ($null -ne $presentation) { try { $presentation.Close() } catch { } }
            if ($null -ne $app) { try { $app.Quit() } catch { } }
        }
    }
    $finished = Wait-Job $job -Timeout 30
    try {
        if ($null -eq $finished) { Stop-Job $job; return $false }
        Receive-Job $job -ErrorAction Stop | Out-Null
        return (Test-Path -LiteralPath $Path -PathType Leaf)
    } finally {
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        if (Test-Path -LiteralPath $ownedRecordPath -PathType Leaf) {
            $ownedRecords = @([IO.File]::ReadAllText($ownedRecordPath) | Where-Object { [string]$_ -match '^OWNED\|\d+\|(WINWORD|EXCEL|POWERPNT)\|\d+$' })
        }
        foreach ($record in $ownedRecords) {
            $parts = @(([string]$record) -split '\|')
            try {
                $process = Get-Process -Id ([int]$parts[1]) -ErrorAction Stop
                if ($process.ProcessName.ToUpperInvariant() -cne $parts[2]) { continue }
                if ($process.StartTime.ToUniversalTime().Ticks -ne [long]$parts[3]) { continue }
                Stop-Process -Id $process.Id -Force -ErrorAction Stop
            } catch { }
        }
        Remove-Item -LiteralPath $ownedRecordPath -Force -ErrorAction SilentlyContinue
    }
}

try {
    $utf8 = New-Object Text.UTF8Encoding($false, $true)
    $utf8Bom = New-Object Text.UTF8Encoding($true, $true)
    $utf16 = New-Object Text.UnicodeEncoding($false, $true, $true)
    $cp932 = [Text.Encoding]::GetEncoding(932)

    Write-ExtractBytes (Join-Path $input "utf8.txt") ($utf8.GetBytes("utf8 body`n"))
    Write-ExtractBytes (Join-Path $input "utf8-bom.txt") (Join-ExtractBytes ($utf8Bom.GetPreamble()) ($utf8Bom.GetBytes("bom body`r`n")))
    Write-ExtractBytes (Join-Path $input "utf16.txt") (Join-ExtractBytes ($utf16.GetPreamble()) ($utf16.GetBytes("wide body`r`n")))
    $jp = ([char]0x7D4C).ToString() + [char]0x8CBB
    Write-ExtractBytes (Join-Path $input "cp932.txt") ($cp932.GetBytes($jp + "`r`n"))
    Write-ExtractBytes (Join-Path $input "broken.txt") ([byte[]](129))
    Write-ExtractBytes (Join-Path $input "empty.txt") ([byte[]]@())
    Write-ExtractBytes (Join-Path $input "unsupported.bin") ([byte[]](1,2,3))
    Write-ExtractBytes (Join-Path $input "deep\a\b\deep.txt") ($utf8.GetBytes("deep body"))
    Write-ExtractBytes (Join-Path $input ".git\hidden.txt") ($utf8.GetBytes("must not scan"))

    $docx = Join-Path $input "sample.docx"
    $xlsx = Join-Path $input "sample.xlsx"
    $pptx = Join-Path $input "sample.pptx"
    New-OpenXmlFixture $docx @{
        "word/document.xml" = '<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:t>DOCX HELLO</w:t></w:r></w:p></w:body></w:document>'
    }
    New-OpenXmlFixture $xlsx @{
        "xl/worksheets/sheet1.xml" = '<worksheet><sheetData><row><c r="A1" t="inlineStr"><is><t>XLSX HELLO</t></is></c></row></sheetData></worksheet>'
    }
    New-OpenXmlFixture $pptx @{
        "ppt/slides/slide1.xml" = '<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><a:p><a:r><a:t>PPTX HELLO</a:t></a:r></a:p></p:cSld></p:sld>'
        "ppt/notesSlides/notesSlide1.xml" = '<p:notes xmlns:p="urn:p" xmlns:a="urn:a"><a:p><a:r><a:t>NOTES HELLO</a:t></a:r></a:p></p:notes>'
    }

    Initialize-NativeHelper
    Assert-True ($null -ne ("ExtractMdNativeHelper" -as [type])) "helper.cs compiles with .NET Framework assemblies"
    $textResult = Read-TextFile (Join-Path $input "utf8.txt")
    Assert-True ($textResult.Content -eq "utf8 body`n" -and $textResult.Notes -eq "UTF-8") "UTF-8 text extraction is unchanged"
    Assert-True ((Read-TextFile (Join-Path $input "utf8-bom.txt")).Notes -eq "UTF-8 BOM") "UTF-8 BOM detected"
    Assert-True ((Read-TextFile (Join-Path $input "utf16.txt")).Notes -eq "UTF-16 LE") "UTF-16LE detected"
    Assert-True ((Read-TextFile (Join-Path $input "cp932.txt")).Content -like "*$jp*") "CP932 text extracted"
    Assert-True ((Read-TextFile (Join-Path $input "broken.txt")).Notes -eq "CP932") "broken bytes preserve the legacy CP932 replacement fallback"
    Assert-True ((Read-TextFile (Join-Path $input "empty.txt")).Content -eq "") "empty text file extracted"
    $hugePath = Join-Path $fx "huge.txt"
    $hugeText = New-Object System.String -ArgumentList ([char]120),([int](20MB))
    [IO.File]::WriteAllText($hugePath, $hugeText, $utf8)
    $hugeResult = Read-TextFile $hugePath
    Assert-True ($hugeResult.Content.Length -eq 20MB) "20 MiB text file extracts without truncation"
    $hugeText = $null; $hugeResult = $null

    $word = Read-WordOpenXmlFile $docx
    $excel = Read-ExcelOpenXmlFile $xlsx
    $powerPoint = Read-PowerPointOpenXmlFile $pptx
    Assert-True ($word.Content -like "*DOCX HELLO*") "DOCX OpenXML extracted without Office"
    Assert-True ($excel.Content -like "*XLSX HELLO*") "XLSX OpenXML extracted without Office"
    Assert-True ($powerPoint.Content -like "*PPTX HELLO*" -and $powerPoint.Content -like "*NOTES HELLO*") "PPTX slides and notes extracted without Office"

    $dtdDocx = Join-Path $fx "dtd.docx"
    New-OpenXmlFixture $dtdDocx @{
        "word/document.xml" = '<!DOCTYPE w:document [<!ENTITY probe "EXPANDED">]><w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:t>&probe;</w:t></w:r></w:p></w:body></w:document>'
    }
    $dtdRejected = $false
    try { [void](Read-WordOpenXmlFile $dtdDocx) } catch { $dtdRejected = ($_.Exception.Message -match 'DTD') }
    Assert-True $dtdRejected "OpenXML DTD declarations are rejected before entity expansion"

    $wideXlsx = Join-Path $fx "wide-column.xlsx"
    New-OpenXmlFixture $wideXlsx @{
        "xl/worksheets/sheet1.xml" = '<worksheet><sheetData><row><c r="ZZZ1"><v>1</v></c></row></sheetData></worksheet>'
    }
    $wideColumnRejected = $false
    try { [void](Read-ExcelOpenXmlFile $wideXlsx) } catch { $wideColumnRejected = ($_.Exception.Message -match 'XFD') }
    Assert-True $wideColumnRejected "XLSX columns beyond XFD are rejected before sparse-row allocation"

    $files = @(Collect-InputFiles @($input))
    Assert-True (@($files | Where-Object { $_.Name -eq "deep.txt" }).Count -eq 1) "deep file included in scan"
    Assert-True (@($files | Where-Object { $_.FullName -like "*\.git\*" }).Count -eq 0) "excluded directory omitted from scan"
    $tree = Build-FileTree @($input)
    Assert-True ($tree -like "*deep.txt*" -and $tree -notlike "*.git*") "file tree follows scan exclusions"

    $results = @(
        (New-ExtractionResult 1 "OK" "Text" (Join-Path $input "utf8.txt") "Text Read" "UTF-8" "body"),
        (New-ExtractionResult 2 "DEFERRED_TIMEOUT" "Word" "slow.doc" "Word COM" "Timeout" ""),
        (New-ExtractionResult 3 "FAIL" "PDF" "bad.pdf" "Word PDF Import" "failure" ""),
        (New-ExtractionResult 4 "UNSUPPORTED" "Unsupported" "x.bin" "Not processed" "unsupported" "")
    )
    $bundle = Build-OutputText @($input) $tree $results "2026-01-01 00:00:00 +09:00" "out.md" 120 300
    Assert-True ($bundle -like "*# Extract Bundle*## File Tree*## Contents*## Deferred Files*## Failed Files*## Unsupported Files*") "bundle keeps all legacy output sections"
    Assert-True ($bundle -like "*Succeeded: 1*Failed: 1*Deferred: 1*Unsupported: 1*") "bundle counts remain consistent"

    $slowDoc = Join-Path $fx "slow.doc"
    Write-ExtractBytes $slowDoc ($utf8.GetBytes("dummy"))
    $env:TOOLRACK_MD_EXTRACT_TEST_DELAY_MS = "2500"
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $timeoutResult = Invoke-FileExtractionWithTimeout (Get-Item $slowDoc) 1 1
    $sw.Stop()
    $env:TOOLRACK_MD_EXTRACT_TEST_DELAY_MS = ""
    Assert-True ($timeoutResult.Status -eq "DEFERRED_TIMEOUT") "worker timeout produces Deferred status"
    Assert-True ($timeoutResult.Notes -like "*TEST_DELAY*") "Deferred result records the last worker stage"
    Assert-True ($sw.Elapsed.TotalSeconds -lt 8) "forced timeout returns promptly"

    $env:TOOLRACK_MD_EXTRACT_TEST_HANG_AFTER_RESULT_MS = "2500"
    $cleanupResult = Invoke-FileExtractionWithTimeout (Get-Item (Join-Path $input "utf8.txt")) 1 1
    $env:TOOLRACK_MD_EXTRACT_TEST_HANG_AFTER_RESULT_MS = ""
    Assert-True ($cleanupResult.Status -like "OK*") "completed result is recovered when COM-style worker cleanup hangs"
    Assert-True ($cleanupResult.Content -eq "utf8 body`n") "result recovery preserves worker content"

    $retryList = New-Object "System.Collections.Generic.List[object]"
    $retryList.Add((New-ExtractionResult 1 "DEFERRED_TIMEOUT" "Text" (Join-Path $input "utf8.txt") "Text Read" "Timeout" ""))
    Invoke-DeferredRetryPass $retryList 10
    Assert-True ($retryList[0].Status -like "OK*") "Deferred retry replaces the timed-out result"
    Assert-True ($retryList[0].Content -eq "utf8 body`n") "Deferred retry returns extracted content"

    $wordType = [type]::GetTypeFromProgID("Word.Application")
    if ($null -eq $wordType) {
        Write-Host "  skip: Word COM and PDF routes (Word not installed)" -ForegroundColor Yellow
    } else {
        $docPath = Join-Path $fx "com.doc"
        $pdfPath = Join-Path $fx "com.pdf"
        $wordMade = New-LegacyOfficeFixture "word" $docPath "WORD COM HELLO"
        $pdfMade = New-LegacyOfficeFixture "pdf" $pdfPath "PDF COM HELLO"
        Assert-True $wordMade "Word COM fixture created"
        Assert-True $pdfMade "PDF fixture created through Word COM"
        if ($wordMade) { Invoke-TestComWorker $docPath "WORD COM HELLO" "Word" }
        if ($pdfMade) { Invoke-TestComWorker $pdfPath "PDF COM HELLO" "PDF" -TimeoutSeconds 10 -AllowDeferred }
    }

    $excelType = [type]::GetTypeFromProgID("Excel.Application")
    if ($null -eq $excelType) {
        Write-Host "  skip: Excel COM route (Excel not installed)" -ForegroundColor Yellow
    } else {
        $xlsPath = Join-Path $fx "com.xls"
        $excelMade = New-LegacyOfficeFixture "excel" $xlsPath "EXCEL COM HELLO"
        Assert-True $excelMade "Excel COM fixture created"
        if ($excelMade) { Invoke-TestComWorker $xlsPath "EXCEL COM HELLO" "Excel" }
    }

    $powerPointType = [type]::GetTypeFromProgID("PowerPoint.Application")
    if ($null -eq $powerPointType) {
        Write-Host "  skip: PowerPoint COM route (PowerPoint not installed)" -ForegroundColor Yellow
    } else {
        $pptPath = Join-Path $fx "com.ppt"
        $powerPointMade = New-LegacyOfficeFixture "powerpoint" $pptPath "POWERPOINT COM HELLO"
        Assert-True $powerPointMade "PowerPoint COM fixture created"
        if ($powerPointMade) { Invoke-TestComWorker $pptPath "POWERPOINT COM HELLO" "PowerPoint" }
    }
    $beforeOutput = @(Get-ChildItem (Join-Path $root "output") -File -Filter "extract-md_*.md" -ErrorAction SilentlyContinue)
    $beforeLog = @(Get-ChildItem (Join-Path $root "log") -File -Filter "extract-md_*.log" -ErrorAction SilentlyContinue)
    try {
        $env:TOOLRACK_NOPAUSE = "1"
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main -Target $input -TimeoutSeconds 10 -DeferredAction Skip -NoOpen
        Assert-True ($LASTEXITCODE -eq 0) "main extraction route exits successfully"
        $afterOutput = @(Get-ChildItem (Join-Path $root "output") -File -Filter "extract-md_*.md" | Where-Object { $beforeOutput.Name -notcontains $_.Name })
        Assert-True ($afterOutput.Count -eq 1) "main route writes one bundle under toolrack output"
        if ($afterOutput.Count -eq 1) {
            $mainBundle = [IO.File]::ReadAllText($afterOutput[0].FullName)
            Assert-True ($mainBundle -like "*DOCX HELLO*" -and $mainBundle -like "*XLSX HELLO*" -and $mainBundle -like "*PPTX HELLO*") "main bundle contains all OpenXML fixture content"
            Assert-True ($mainBundle -notlike "*must not scan*") "main bundle excludes configured directories"
        }
    } finally {
        $env:TOOLRACK_NOPAUSE = ""
        Remove-NewExtractArtifacts $beforeOutput $beforeLog
    }

    $toolJson = [IO.File]::ReadAllText($manifest) | ConvertFrom-Json
    $customVariant = @($toolJson.variants | Where-Object { $_.label -eq "Custom..." })[0]
    Assert-True (@($customVariant.args).Count -eq 1 -and $customVariant.args[0] -eq "-Custom") "Custom variant is distinct from Default"

    $asciiFiles = @(Get-ChildItem $toolDir -File | Where-Object { $_.Extension -in @(".ps1", ".cs", ".json") })
    foreach ($file in $asciiFiles) {
        $bytes = [IO.File]::ReadAllBytes($file.FullName)
        Assert-True (@($bytes | Where-Object { $_ -gt 127 }).Count -eq 0) ("ASCII source: " + $file.Name)
    }
} finally {
    $env:TOOLRACK_MD_EXTRACT_TEST_DELAY_MS = ""
    $env:TOOLRACK_MD_EXTRACT_TEST_HANG_AFTER_RESULT_MS = ""
    $env:TOOLRACK_NOPAUSE = ""
    Stop-OfficeApplications
    Remove-Item -LiteralPath $fx -Recurse -Force -ErrorAction SilentlyContinue
}

Exit-Test
