<#
.SYNOPSIS
  Collects a phase-2 measurement set with explicitly approved controlled input.

.DESCRIPTION
  Starts the three sample tools once (seven workers), launches a hidden
  mouse-move-only activity helper for each window, and keeps taking 90-second
  windows until at least three valid windows and 30 WAIT onsets are available.
  Invalid attempts and their rejection reasons remain in the output tree.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateRange(1, 2147483647)]
    [int]$ExcelPid,

    [Parameter(Mandatory)]
    [ValidatePattern('^[A-Za-z0-9][A-Za-z0-9_.-]*$')]
    [string]$Label,

    [ValidateRange(10, 240)]
    [int]$Seconds = 90,

    [ValidateRange(1, 20)]
    [int]$RequiredValidWindows = 3,

    [ValidateRange(1, 1000)]
    [int]$RequiredWaitOnsets = 30,

    [ValidateRange(1, 20)]
    [int]$MaxAttempts = 8,

    [ValidateRange(0, 30)]
    [int]$InitialWarmupSeconds = 15,

    [ValidateRange(200, 1000)]
    [int]$SampleRateHz = 250,

    [string]$OutRoot = (Join-Path $env:TEMP 'xltoolrack-phase2-measurements')
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$measureScript = Join-Path $PSScriptRoot 'Measure-Desktop.ps1'
$activityScript = Join-Path $PSScriptRoot 'Run-ControlledExcelActivity.ps1'
foreach ($script in @($measureScript, $activityScript)) {
    if (-not (Test-Path -LiteralPath $script)) {
        throw "Required script was not found: $script"
    }
}

$excel = Get-Process -Id $ExcelPid -ErrorAction Stop
if ($excel.ProcessName -ine 'EXCEL') {
    throw "PID $ExcelPid is not Excel."
}
if ($excel.MainWindowHandle -eq 0) {
    throw "Excel PID $ExcelPid has no visible main window."
}

$setRoot = Join-Path ([IO.Path]::GetFullPath($OutRoot)) $Label
if (-not (Test-Path -LiteralPath $setRoot)) {
    New-Item -ItemType Directory -Path $setRoot -Force | Out-Null
}

function Read-SummaryField {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name
    )
    $escaped = [regex]::Escape($Name)
    $line = Get-Content -LiteralPath $Path |
        Where-Object { $_ -match ('^' + $escaped + '\s*:\s*(.*)$') } |
        Select-Object -First 1
    if ($null -eq $line) { return $null }
    return [regex]::Match($line, ('^' + $escaped + '\s*:\s*(.*)$')).Groups[1].Value.Trim()
}

$attemptRecords = New-Object Collections.Generic.List[object]
$validRecords = New-Object Collections.Generic.List[object]
$totalWaitOnsets = 0
$startedAt = Get-Date

for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
    if ($validRecords.Count -ge $RequiredValidWindows -and
        $totalWaitOnsets -ge $RequiredWaitOnsets) {
        break
    }

    $attemptName = 'attempt-{0:D2}' -f $attempt
    $attemptDir = Join-Path $setRoot $attemptName
    New-Item -ItemType Directory -Path $attemptDir -Force | Out-Null
    $readyFile = Join-Path $attemptDir 'activity.ready'
    $stopFile = Join-Path $attemptDir 'activity.stop'
    $activitySummary = Join-Path $attemptDir 'controlled-activity.json'
    $attemptMetadata = Join-Path $attemptDir 'attempt.json'
    foreach ($path in @($readyFile, $stopFile)) {
        if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
        }
    }

    $warmupSeconds = if ($attempt -eq 1) { $InitialWarmupSeconds } else { 2 }
    $helperSeconds = $Seconds + $warmupSeconds + 45
    $helperArguments = @(
        '-NoLogo'
        '-NoProfile'
        '-ExecutionPolicy'
        'Bypass'
        '-File'
        ('"' + $activityScript + '"')
        '-ExcelPid'
        [string]$ExcelPid
        '-Seconds'
        [string]$helperSeconds
        '-IntervalMilliseconds'
        '180'
        '-ReadyFile'
        ('"' + $readyFile + '"')
        '-StopFile'
        ('"' + $stopFile + '"')
        '-SummaryPath'
        ('"' + $activitySummary + '"')
    )

    $helper = $null
    $measurementError = $null
    $summaryPath = $null
    $valid = $false
    $waitOnsets = 0
    $pumpSamples = 0
    try {
        Write-Host "[$Label] starting controlled window attempt $attempt..." -ForegroundColor Cyan
        $helper = Start-Process -FilePath 'powershell.exe' `
            -ArgumentList $helperArguments -WindowStyle Hidden -PassThru
        $readyDeadline = (Get-Date).AddSeconds(15)
        while (-not (Test-Path -LiteralPath $readyFile) -and
               (Get-Date) -lt $readyDeadline) {
            if ($helper.HasExited) {
                throw "Controlled activity helper exited before becoming ready (exit=$($helper.ExitCode))."
            }
            Start-Sleep -Milliseconds 100
        }
        if (-not (Test-Path -LiteralPath $readyFile)) {
            throw 'Timed out waiting for controlled activity helper readiness.'
        }

        $measureArguments = @{
            ExcelPid = $ExcelPid
            Seconds = $Seconds
            SampleRateHz = $SampleRateHz
            WarmupSeconds = $warmupSeconds
            ForegroundWaitSeconds = 10
            ActivateExcel = $true
            ActivitySource = 'controlled'
            ExpectedJobCount = 7
            OutDir = $attemptDir
        }
        if ($attempt -eq 1) {
            $measureArguments.RestartAllTools = $true
        }
        & $measureScript @measureArguments

        $summaryPath = Get-ChildItem -LiteralPath $attemptDir `
            -Filter 'desktop-summary_*.txt' |
            Sort-Object LastWriteTimeUtc |
            Select-Object -Last 1 -ExpandProperty FullName
        if ([string]::IsNullOrWhiteSpace($summaryPath)) {
            throw 'Measure-Desktop completed without writing a desktop summary.'
        }
        $valid = ((Read-SummaryField $summaryPath 'Window valid') -eq 'True')
        $waitOnsets = [int](Read-SummaryField $summaryPath 'WAIT onset count')
        $pumpSamples = [int](Read-SummaryField $summaryPath 'Pump samples')
        if (-not $valid) {
            throw 'Measure-Desktop returned but marked the window invalid.'
        }
    }
    catch {
        $measurementError = $_.Exception.Message
        if ($null -eq $summaryPath) {
            $summary = Get-ChildItem -LiteralPath $attemptDir `
                -Filter 'desktop-summary_*.txt' -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTimeUtc |
                Select-Object -Last 1
            if ($null -ne $summary) {
                $summaryPath = $summary.FullName
                $value = Read-SummaryField $summaryPath 'WAIT onset count'
                if ($null -ne $value) { $waitOnsets = [int]$value }
                $value = Read-SummaryField $summaryPath 'Pump samples'
                if ($null -ne $value) { $pumpSamples = [int]$value }
            }
        }
        Write-Warning "[$Label] attempt $attempt rejected: $measurementError"
    }
    finally {
        [IO.File]::WriteAllText($stopFile, (Get-Date).ToString('o'))
        if ($null -ne $helper -and -not $helper.HasExited) {
            if (-not $helper.WaitForExit(5000)) {
                # PID belongs to the helper created immediately above.
                Stop-Process -Id $helper.Id -Force -ErrorAction SilentlyContinue
            }
        }
    }

    $record = [PSCustomObject][ordered]@{
        Attempt = $attempt
        Valid = $valid
        WaitOnsets = $waitOnsets
        PumpSamples = $pumpSamples
        Error = $measurementError
        Directory = $attemptDir
        DesktopSummary = $summaryPath
        ActivitySummary = $activitySummary
    }
    $attemptRecords.Add($record)
    if ($valid) {
        $validRecords.Add($record)
        $totalWaitOnsets += $waitOnsets
    }
    $record | ConvertTo-Json | Out-File -LiteralPath $attemptMetadata -Encoding UTF8
}

$completedAt = Get-Date
$setPassed = ($validRecords.Count -ge $RequiredValidWindows -and
              $totalWaitOnsets -ge $RequiredWaitOnsets)
$setSummary = [PSCustomObject][ordered]@{
    Label = $Label
    ExcelPid = $ExcelPid
    StartedAt = $startedAt.ToString('o')
    CompletedAt = $completedAt.ToString('o')
    MeasurementSeconds = $Seconds
    SampleRateHz = $SampleRateHz
    ActivitySource = 'controlled SendInput mouse-move only'
    RequiredValidWindows = $RequiredValidWindows
    RequiredWaitOnsets = $RequiredWaitOnsets
    AttemptCount = $attemptRecords.Count
    ValidWindowCount = $validRecords.Count
    TotalWaitOnsets = $totalWaitOnsets
    Passed = $setPassed
    Attempts = [object[]]$attemptRecords.ToArray()
}
$setSummaryPath = Join-Path $setRoot 'measurement-set.json'
$setSummary | ConvertTo-Json -Depth 5 | Out-File -LiteralPath $setSummaryPath -Encoding UTF8
$setSummary | ConvertTo-Json -Depth 5

if (-not $setPassed) {
    throw "Measurement set '$Label' did not meet validity requirements: valid=$($validRecords.Count)/$RequiredValidWindows, WAIT onsets=$totalWaitOnsets/$RequiredWaitOnsets."
}
