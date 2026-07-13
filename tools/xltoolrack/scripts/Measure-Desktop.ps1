<#
.SYNOPSIS
  Measures the real, visible xltoolrack Excel instance.

.DESCRIPTION
  Attaches to an existing visible Excel process without creating a test Excel
  instance. During one observation window it captures both the in-memory FE
  pump profile and the system cursor WAIT intervals. The workbook remains open
  for either normal owner activity or an explicitly approved controlled-input
  helper. A phase-2 window is rejected unless its scaled tick floor (70 ticks
  per 90 seconds), cursor sampling, target foreground, visibility, and movement
  checks pass. ActivitySource records which mode supplied the movement.

.EXAMPLE
  .\scripts\Measure-Desktop.ps1 -ExcelPid 1234 -Seconds 90 `
      -StartTools pi_race,life -OutDir .\docs\measurements\after
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateRange(1, 2147483647)]
    [int]$ExcelPid,

    [ValidateRange(10, 240)]
    [int]$Seconds = 90,

    [ValidateRange(200, 1000)]
    [int]$SampleRateHz = 250,

    [ValidateRange(0, 30)]
    [int]$WarmupSeconds = 5,

    [ValidateRange(0, 300)]
    [int]$ForegroundWaitSeconds = 60,

    [ValidateSet('stopwatch', 'pi_race', 'life')]
    [string[]]$StartTools = @(),

    [switch]$RestartAllTools,

    [switch]$ActivateExcel,

    [ValidateSet('owner', 'controlled')]
    [string]$ActivitySource = 'owner',

    [ValidateRange(0, 100)]
    [int]$ExpectedJobCount = 0,

    [string]$OutDir = (Join-Path $env:TEMP 'xltoolrack-desktop-measurement')
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

if (-not (Test-Path -LiteralPath $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
}
$OutDir = [IO.Path]::GetFullPath($OutDir)

if (-not ('XltrDesktopNative' -as [type])) {
    Add-Type @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class XltrDesktopNative
{
    public delegate bool EnumWindowProc(IntPtr hwnd, IntPtr parameter);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(
        IntPtr parent, EnumWindowProc callback, IntPtr parameter);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern int GetClassName(
        IntPtr hwnd, StringBuilder className, int capacity);

    [DllImport("oleacc.dll")]
    public static extern int AccessibleObjectFromWindow(
        IntPtr hwnd, uint objectId, ref Guid iid,
        [MarshalAs(UnmanagedType.Interface)] out object value);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(
        IntPtr hwnd, out uint processId);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();

    [DllImport("user32.dll")]
    private static extern bool AttachThreadInput(
        uint attachThread, uint attachToThread, bool attach);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hwnd);

    [DllImport("user32.dll")]
    private static extern bool BringWindowToTop(IntPtr hwnd);

    [DllImport("user32.dll")]
    private static extern bool ShowWindowAsync(IntPtr hwnd, int command);

    public static IntPtr[] FindDescendantsByClass(IntPtr parent, string wanted)
    {
        var result = new List<IntPtr>();
        EnumChildWindows(parent, delegate(IntPtr hwnd, IntPtr unused) {
            var className = new StringBuilder(256);
            GetClassName(hwnd, className, className.Capacity);
            if (String.Equals(className.ToString(), wanted,
                              StringComparison.OrdinalIgnoreCase)) {
                result.Add(hwnd);
            }
            return true;
        }, IntPtr.Zero);
        return result.ToArray();
    }

    public static bool ActivateWindow(IntPtr hwnd)
    {
        uint targetProcessId;
        GetWindowThreadProcessId(hwnd, out targetProcessId);
        uint unused;
        uint foregroundThread = GetWindowThreadProcessId(
            GetForegroundWindow(), out unused);
        uint currentThread = GetCurrentThreadId();
        bool attached = foregroundThread != currentThread &&
                        AttachThreadInput(currentThread, foregroundThread, true);
        ShowWindowAsync(hwnd, 9); // SW_RESTORE
        BringWindowToTop(hwnd);
        SetForegroundWindow(hwnd);
        if (attached) {
            AttachThreadInput(currentThread, foregroundThread, false);
        }
        return ForegroundProcessId() == targetProcessId;
    }

    public static uint ForegroundProcessId()
    {
        uint processId;
        GetWindowThreadProcessId(GetForegroundWindow(), out processId);
        return processId;
    }
}
'@
}

function Release-ComObject {
    param($Value)
    if ($null -eq $Value) { return }
    try {
        if ([Runtime.InteropServices.Marshal]::IsComObject($Value)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
        }
    } catch {}
}

function Connect-VisibleExcel {
    param([int]$ProcessId)

    $process = Get-Process -Id $ProcessId -ErrorAction Stop
    if ($process.ProcessName -ine 'EXCEL') {
        throw "PID $ProcessId is not Excel."
    }
    $mainWindow = [IntPtr]$process.MainWindowHandle
    if ($mainWindow -eq [IntPtr]::Zero) {
        throw "Excel PID $ProcessId has no visible main window."
    }

    $iidDispatch = [Guid]'00020400-0000-0000-C000-000000000046'
    $nativeObjectId = [uint32]4294967280 # OBJID_NATIVEOM (0xFFFFFFF0)
    foreach ($child in [XltrDesktopNative]::FindDescendantsByClass($mainWindow, 'EXCEL7')) {
        $nativeObject = $null
        $application = $null
        try {
            $hr = [XltrDesktopNative]::AccessibleObjectFromWindow(
                $child, $nativeObjectId, [ref]$iidDispatch, [ref]$nativeObject)
            if ($hr -ne 0 -or $null -eq $nativeObject) { continue }
            $application = $nativeObject.Application
            if ([int64]$application.Hwnd -eq $mainWindow.ToInt64()) {
                return [PSCustomObject]@{
                    Application = $application
                    NativeObject = $nativeObject
                }
            }
        } catch {
            Release-ComObject $application
            Release-ComObject $nativeObject
            continue
        }
        Release-ComObject $application
        Release-ComObject $nativeObject
    }
    throw "Could not attach to the visible workbook in Excel PID $ProcessId."
}

function QualifiedMacro {
    param([string]$Procedure)
    return "'xltoolrack.xlam'!$Procedure"
}

function Invoke-ExcelMacro {
    param($Excel, [string]$Procedure, [object[]]$Arguments = @())

    $busyErrors = @(-2146777998, -2147417846, -2147418111)
    $deadline = (Get-Date).AddSeconds(30)
    while ($true) {
        try {
            $macro = QualifiedMacro $Procedure
            switch ($Arguments.Count) {
                0 { return $Excel.Run($macro) }
                1 { return $Excel.Run($macro, $Arguments[0]) }
                2 { return $Excel.Run($macro, $Arguments[0], $Arguments[1]) }
                default { throw 'Invoke-ExcelMacro supports at most two arguments.' }
            }
        } catch {
            $hr = $_.Exception.HResult
            if ($null -ne $_.Exception.InnerException -and
                $_.Exception.InnerException.HResult -ne 0) {
                $hr = $_.Exception.InnerException.HResult
            }
            if (($busyErrors -contains [int]$hr) -and (Get-Date) -lt $deadline) {
                Start-Sleep -Milliseconds 200
                continue
            }
            throw
        }
    }
}

function Get-NearestRankPercentile {
    param([double[]]$Values, [double]$Percentile)
    if ($Values.Count -eq 0) { return 0.0 }
    $sorted = @($Values | Sort-Object)
    $index = [Math]::Max(0, [Math]::Ceiling($Percentile * $sorted.Count) - 1)
    return [double]$sorted[$index]
}

function Wait-ForNoJobs {
    param($Excel)
    $deadline = (Get-Date).AddSeconds(30)
    do {
        $count = [int](Invoke-ExcelMacro $Excel 'JobTest_RunningCount')
        if ($count -eq 0) { return }
        Start-Sleep -Milliseconds 500
    } while ((Get-Date) -lt $deadline)
    throw "Timed out waiting for xltoolrack jobs to stop (remaining=$count)."
}

$connection = $null
$excel = $null
try {
    $connection = Connect-VisibleExcel $ExcelPid
    $excel = $connection.Application

    if ($RestartAllTools) {
        [void](Invoke-ExcelMacro $excel 'JobTest_StopAll')
        Wait-ForNoJobs $excel
        $StartTools = @('stopwatch', 'pi_race', 'life')
        if ($ExpectedJobCount -eq 0) { $ExpectedJobCount = 7 }
    }

    foreach ($tool in $StartTools) {
        [void](Invoke-ExcelMacro $excel 'Infra_Run' @($tool))
    }

    if ($WarmupSeconds -gt 0) {
        Write-Host "Warming up the live Excel workload for $WarmupSeconds seconds..."
        Start-Sleep -Seconds $WarmupSeconds
    }

    if ($ActivateExcel) {
        $mainWindow = [IntPtr](Get-Process -Id $ExcelPid -ErrorAction Stop).MainWindowHandle
        if (-not [XltrDesktopNative]::ActivateWindow($mainWindow)) {
            throw "Could not activate Excel PID $ExcelPid."
        }
        Start-Sleep -Milliseconds 500
    }

    if ([int][XltrDesktopNative]::ForegroundProcessId() -ne $ExcelPid -and
        $ForegroundWaitSeconds -gt 0) {
        Write-Host "Waiting up to $ForegroundWaitSeconds seconds for the owner to bring Excel PID $ExcelPid to the foreground..."
        $foregroundDeadline = (Get-Date).AddSeconds($ForegroundWaitSeconds)
        while ([int][XltrDesktopNative]::ForegroundProcessId() -ne $ExcelPid -and
               (Get-Date) -lt $foregroundDeadline) {
            Start-Sleep -Milliseconds 200
        }
    }
    if ([int][XltrDesktopNative]::ForegroundProcessId() -ne $ExcelPid) {
        throw "Excel PID $ExcelPid was not brought to the foreground; no measurement was taken."
    }

    $activeJobsAtStart = [int](Invoke-ExcelMacro $excel 'JobTest_RunningCount')
    [void](Invoke-ExcelMacro $excel 'JobPump.Pump_ProfileResetForTest')
    [void](Invoke-ExcelMacro $excel 'JobPump.Pump_EnsureArmed')
    if ([int][XltrDesktopNative]::ForegroundProcessId() -ne $ExcelPid) {
        throw "Excel PID $ExcelPid lost the foreground while preparing the measurement."
    }

    $foregroundPidAtStart = [int][XltrDesktopNative]::ForegroundProcessId()
    $mainWindowAtStart = [IntPtr](Get-Process -Id $ExcelPid -ErrorAction Stop).MainWindowHandle
    Write-Host "Measuring live Excel PID $ExcelPid for $Seconds seconds."
    Write-Host 'Keep using the visible workbook normally during this window.'
    $cursorScript = Join-Path $PSScriptRoot 'Measure-Cursor.ps1'
    $cursorResult = & $cursorScript -Seconds $Seconds -SampleRateHz $SampleRateHz `
        -ExpectedForegroundProcessId $ExcelPid -OutDir $OutDir
    $foregroundPidAtEnd = [int][XltrDesktopNative]::ForegroundProcessId()
    $mainWindowAtEnd = [IntPtr](Get-Process -Id $ExcelPid -ErrorAction Stop).MainWindowHandle

    $profileText = [string](Invoke-ExcelMacro $excel 'JobPump.Pump_ProfileForTest')
    $activeJobsAtEnd = [int](Invoke-ExcelMacro $excel 'JobTest_RunningCount')

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $profilePath = Join-Path $OutDir "pump-profile_$stamp.csv"
    [IO.File]::WriteAllText($profilePath, $profileText, [Text.Encoding]::UTF8)

    $rows = @($profileText | ConvertFrom-Csv)
    $freshRows = @($rows | Where-Object { [int]$_.FreshRecords -gt 0 })
    $elapsed = [double[]]@($rows | ForEach-Object { [double]$_.ElapsedMs })
    $freshElapsed = [double[]]@($freshRows | ForEach-Object { [double]$_.ElapsedMs })
    $aggregateElapsed = [double[]]@($rows | ForEach-Object { [double]$_.AggregateReadMs })
    $pumpOnceElapsed = [double[]]@($rows | ForEach-Object { [double]$_.PumpOnceMs })
    $dispatchElapsed = [double[]]@($rows | ForEach-Object { [double]$_.DispatchMs })
    $errorPollElapsed = [double[]]@($rows | ForEach-Object { [double]$_.ErrorPollMs })
    $cleanupElapsed = [double[]]@($rows | ForEach-Object { [double]$_.CleanupMs })
    $freshRecordTotal = if ($rows.Count -gt 0) {
        [int](($rows | Measure-Object FreshRecords -Sum).Sum)
    } else { 0 }
    $maxFreshRecords = if ($rows.Count -gt 0) {
        [int](($rows | Measure-Object FreshRecords -Maximum).Maximum)
    } else { 0 }
    $maxFlushTools = if ($rows.Count -gt 0) {
        [int](($rows | Measure-Object FlushTools -Maximum).Maximum)
    } else { 0 }
    # The phase-2 protocol requires at least 70 executed ticks in a 90-second
    # window. Scale that floor for shorter smoke or longer evidence windows.
    $minimumPumpSamples = [Math]::Max(5, [Math]::Ceiling($Seconds * 70.0 / 90.0))
    $profileValid = ($rows.Count -ge $minimumPumpSamples)
    $excelVisibleAtEndpoints = ($mainWindowAtStart -ne [IntPtr]::Zero -and
                                $mainWindowAtEnd -ne [IntPtr]::Zero)
    $pointerActivityObserved = ([int]$cursorResult.CursorPositionChanges -gt 0)
    $jobCountValid = ($ExpectedJobCount -eq 0 -or
                      ($activeJobsAtStart -eq $ExpectedJobCount -and
                       $activeJobsAtEnd -eq $ExpectedJobCount))
    $windowValid = ($profileValid -and [bool]$cursorResult.Valid -and
                    $excelVisibleAtEndpoints -and $pointerActivityObserved -and
                    $jobCountValid)
    $summaryPath = Join-Path $OutDir "desktop-summary_$stamp.txt"
    $summary = @(
        'xltoolrack live desktop measurement'
        '==================================='
        "Excel PID                 : $ExcelPid"
        "Measurement seconds       : $Seconds"
        "Cursor target rate        : $SampleRateHz Hz"
        ('Cursor resolution (ms)   : {0:N3}' -f [double]$cursorResult.SamplingResolutionMs)
        ('Cursor effective rate    : {0:N1} Hz' -f [double]$cursorResult.EffectiveSampleRateHz)
        ('Cursor gap P95 (ms)      : {0:N3}' -f [double]$cursorResult.SampleGapP95Ms)
        "Cursor sampler priority   : $($cursorResult.Priority)"
        "Tools started by script   : $($StartTools -join ',')"
        "Foreground PID start/end  : $foregroundPidAtStart / $foregroundPidAtEnd"
        ('Excel foreground rate    : {0:P2}' -f [double]$cursorResult.ExpectedForegroundFraction)
        "Excel visible start/end   : $excelVisibleAtEndpoints"
        "Cursor position changes   : $($cursorResult.CursorPositionChanges)"
        "Activity source           : $ActivitySource"
        "Pointer activity observed : $pointerActivityObserved"
        "Active jobs start/end     : $activeJobsAtStart / $activeJobsAtEnd"
        "Expected active jobs      : $ExpectedJobCount"
        "Active job check passed   : $jobCountValid"
        "Pump samples              : $($rows.Count)"
        "Minimum pump samples      : $minimumPumpSamples"
        "Pump profile valid        : $profileValid"
        "Cursor measurement valid  : $($cursorResult.Valid)"
        "Window valid              : $windowValid"
        "Fresh-result ticks        : $($freshRows.Count)"
        "Fresh records total/max   : $freshRecordTotal / $maxFreshRecords"
        ('Tick elapsed P50 (ms)     : {0:N3}' -f (Get-NearestRankPercentile $elapsed 0.50))
        ('Tick elapsed P95 (ms)     : {0:N3}' -f (Get-NearestRankPercentile $elapsed 0.95))
        ('Tick elapsed max (ms)     : {0:N3}' -f (Get-NearestRankPercentile $elapsed 1.00))
        ('Fresh tick P95 (ms)       : {0:N3}' -f (Get-NearestRankPercentile $freshElapsed 0.95))
        ('Aggregate read P95 (ms)   : {0:N3}' -f (Get-NearestRankPercentile $aggregateElapsed 0.95))
        ('PumpOnce P95 (ms)         : {0:N3}' -f (Get-NearestRankPercentile $pumpOnceElapsed 0.95))
        ('Dispatch P95 (ms)         : {0:N3}' -f (Get-NearestRankPercentile $dispatchElapsed 0.95))
        ('Error poll P95 (ms)       : {0:N3}' -f (Get-NearestRankPercentile $errorPollElapsed 0.95))
        ('Cleanup P95 (ms)          : {0:N3}' -f (Get-NearestRankPercentile $cleanupElapsed 0.95))
        "Flush tools max/tick      : $maxFlushTools"
        "WAIT onset count          : $($cursorResult.WaitOnsets)"
        ('WAIT duration P50 (ms)    : {0:N3}' -f [double]$cursorResult.WaitP50Ms)
        ('WAIT duration P95 (ms)    : {0:N3}' -f [double]$cursorResult.WaitP95Ms)
        ('WAIT duration max (ms)    : {0:N3}' -f [double]$cursorResult.WaitMaxMs)
        "Pump profile              : $profilePath"
        "Cursor summary            : $($cursorResult.SummaryPath)"
    )
    $summary | Out-File -LiteralPath $summaryPath -Encoding UTF8
    $summary | ForEach-Object { Write-Host $_ }
    if (-not $windowValid) {
        throw "Invalid measurement window: ticks=$($rows.Count)/$minimumPumpSamples, cursorValid=$($cursorResult.Valid), visible=$excelVisibleAtEndpoints, cursorMoves=$($cursorResult.CursorPositionChanges), jobs=$activeJobsAtStart/$activeJobsAtEnd expected=$ExpectedJobCount."
    }
} finally {
    if ($null -ne $connection) {
        Release-ComObject $connection.NativeObject
        Release-ComObject $connection.Application
    }
}
