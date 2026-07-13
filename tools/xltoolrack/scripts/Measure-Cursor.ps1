<#
.SYNOPSIS
  Passive cursor-shape observer for xltoolrack cursor-flicker diagnosis.
  Records IDC_WAIT intervals, their P50/P95 duration, and IDC_APPSTARTING
  as a negative control. Sampling uses a Stopwatch spin-wait at 250 Hz by
  default and temporarily raises only the sampler process to High priority.
  It never touches Excel, activates a window, or sends input.

.PARAMETER Seconds
  Duration of the observation window in seconds.

.PARAMETER OutDir
  Directory for the transition CSV, per-second WAIT CSV, and summary.

.PARAMETER SampleRateHz
  Spin-wait sample rate. Phase-2 evidence requires at least 200 Hz.

.PARAMETER ExpectedForegroundProcessId
  Optional process id that must own at least 95% of foreground samples and
  both endpoint foreground windows.

.EXAMPLE
  .\scripts\Measure-Cursor.ps1 -Seconds 90 -OutDir .\measure-after
#>
[CmdletBinding()]
param(
    [ValidateRange(1, 86400)]
    [int]$Seconds = 60,

    [ValidateRange(200, 1000)]
    [int]$SampleRateHz = 250,

    [ValidateRange(0, 2147483647)]
    [int]$ExpectedForegroundProcessId = 0,

    [string]$OutDir = (Join-Path $env:TEMP 'xltoolrack-cursor-measurements')
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

if (-not (Test-Path -LiteralPath $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
}

if (-not ('CursorNative' -as [type])) {
    Add-Type @"
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

public struct POINT
{
    public int x;
    public int y;
}

public struct CURSORINFO
{
    public int cbSize;
    public int flags;
    public IntPtr hCursor;
    public POINT ptScreenPos;
}

public static class CursorNative
{
    [DllImport("user32.dll")]
    public static extern bool GetCursorInfo(out CURSORINFO pci);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr LoadCursor(IntPtr hInstance, IntPtr lpCursorName);

    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(
        IntPtr hwnd, out uint processId);

    [DllImport("winmm.dll")]
    public static extern uint timeBeginPeriod(uint uMilliseconds);

    [DllImport("winmm.dll")]
    public static extern uint timeEndPeriod(uint uMilliseconds);

    public static uint ForegroundProcessId()
    {
        uint processId;
        GetWindowThreadProcessId(GetForegroundWindow(), out processId);
        return processId;
    }

    public static void SpinUntil(Stopwatch stopwatch, long targetTicks)
    {
        while (stopwatch.ElapsedTicks < targetTicks) {
            Thread.SpinWait(32);
        }
    }
}
"@
}

$cursorIds = @{
    32650 = 'APPSTARTING'
    32512 = 'ARROW'
    32514 = 'WAIT'
    32513 = 'IBEAM'
    32515 = 'CROSS'
}
$cursorIdOrder = @(32650, 32512, 32514, 32513, 32515)
$handleMap = @{}
foreach ($id in $cursorIdOrder) {
    $handle = [CursorNative]::LoadCursor([IntPtr]::Zero, [IntPtr]$id)
    $handleMap[$handle.ToInt64()] = $cursorIds[$id]
}

function Get-CursorClass {
    param([IntPtr]$Handle)
    $key = $Handle.ToInt64()
    if ($handleMap.ContainsKey($key)) { return [string]$handleMap[$key] }
    return ('unknown:0x{0:X}' -f $key)
}

function Add-WaitSlice {
    param(
        [hashtable]$Buckets,
        [double]$FromMs,
        [double]$ToMs
    )
    $cursor = $FromMs
    while ($cursor -lt $ToMs) {
        $second = [int][Math]::Floor($cursor / 1000.0)
        $boundary = [Math]::Min($ToMs, ($second + 1) * 1000.0)
        if (-not $Buckets.ContainsKey($second)) { $Buckets[$second] = 0.0 }
        $Buckets[$second] = [double]$Buckets[$second] + ($boundary - $cursor)
        $cursor = $boundary
    }
}

function Get-NearestRankPercentile {
    param(
        [System.Collections.Generic.List[double]]$Values,
        [double]$Percentile
    )
    if ($Values.Count -eq 0) { return 0.0 }
    $sorted = @($Values | Sort-Object)
    $rank = [Math]::Max(1, [Math]::Ceiling($Percentile * $sorted.Count))
    return [double]$sorted[[int]$rank - 1]
}

function Get-ForegroundState {
    $handle = [CursorNative]::GetForegroundWindow()
    return [PSCustomObject]@{
        Handle = $handle.ToInt64()
        NonZero = ($handle -ne [IntPtr]::Zero)
        ProcessId = [int][CursorNative]::ForegroundProcessId()
        Time = Get-Date
    }
}

$records = New-Object System.Collections.Generic.List[object]
$waitDurations = New-Object System.Collections.Generic.List[double]
$sampleGaps = New-Object System.Collections.Generic.List[double]
$waitPerSecond = @{}
$cursorInfoSize = 16 + [IntPtr]::Size
$ci = New-Object CURSORINFO
$ci.cbSize = $cursorInfoSize

$foregroundStart = Get-ForegroundState
$startWallTime = Get-Date
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$lastHandle = [IntPtr]::Zero
$lastClass = $null
$lastFlags = -1
$lastTimeMs = 0.0
$haveLast = $false
$waitOnsets = 0
$appStartingOnsets = 0
$currentWaitOnsetMs = $null
$sampleCount = 0
$successfulSamples = 0
$failedSamples = 0
$suppressedSamples = 0
$hiddenSamples = 0
$timerRaised = $false
$priorityRaised = $false
$priorityDuringMeasurement = ''
$currentProcess = [Diagnostics.Process]::GetCurrentProcess()
$originalPriority = $currentProcess.PriorityClass
$stopwatchFrequency = [double][Diagnostics.Stopwatch]::Frequency
$samplePeriodTicks = [long][Math]::Round($stopwatchFrequency / $SampleRateHz)
$samplePeriodMs = 1000.0 / $SampleRateHz
$measurementTicks = [long][Math]::Round($Seconds * $stopwatchFrequency)
$nextSampleTicks = 0L
$previousSampleMs = $null
$missedDeadlines = 0
$expectedForegroundSamples = 0
$foregroundSamples = 0
$cursorPositionChanges = 0
$havePosition = $false
$lastPositionX = 0
$lastPositionY = 0

try {
    $currentProcess.PriorityClass = [Diagnostics.ProcessPriorityClass]::High
    $priorityRaised = $true
    $priorityDuringMeasurement = [string]$currentProcess.PriorityClass
    if ($currentProcess.PriorityClass -ne [Diagnostics.ProcessPriorityClass]::High) {
        throw 'Could not raise the cursor sampler process priority to High.'
    }
    [void][CursorNative]::timeBeginPeriod(1)
    $timerRaised = $true

    while ($stopwatch.ElapsedTicks -lt $measurementTicks) {
        $ci.cbSize = $cursorInfoSize
        $ok = [CursorNative]::GetCursorInfo([ref]$ci)
        $nowMs = $stopwatch.Elapsed.TotalMilliseconds
        $sampleCount++
        if ($null -ne $previousSampleMs) {
            $sampleGaps.Add($nowMs - [double]$previousSampleMs)
        }
        $previousSampleMs = $nowMs

        $foregroundPid = [int][CursorNative]::ForegroundProcessId()
        if ($foregroundPid -ne 0) { $foregroundSamples++ }
        if ($ExpectedForegroundProcessId -gt 0 -and
            $foregroundPid -eq $ExpectedForegroundProcessId) {
            $expectedForegroundSamples++
        }

        if ($ok) {
            $successfulSamples++
            $handle = $ci.hCursor
            $class = Get-CursorClass $handle
            $flags = [int]$ci.flags
            if (($flags -band 2) -ne 0) { $suppressedSamples++ }
            if (($flags -band 1) -eq 0) { $hiddenSamples++ }
            if ($havePosition -and
                ($ci.ptScreenPos.x -ne $lastPositionX -or
                 $ci.ptScreenPos.y -ne $lastPositionY)) {
                $cursorPositionChanges++
            }
            $lastPositionX = $ci.ptScreenPos.x
            $lastPositionY = $ci.ptScreenPos.y
            $havePosition = $true

            if ($haveLast -and $lastClass -eq 'WAIT') {
                Add-WaitSlice $waitPerSecond $lastTimeMs $nowMs
            }

            if (-not $haveLast -or $handle -ne $lastHandle -or $flags -ne $lastFlags) {
                $records.Add([PSCustomObject]@{
                    ElapsedMs = [Math]::Round($nowMs, 1)
                    WallClock = $startWallTime.AddMilliseconds($nowMs).ToString('HH:mm:ss.fff')
                    Handle = ('0x{0:X}' -f $handle.ToInt64())
                    Class = $class
                    Flags = $flags
                })

                if ($class -eq 'WAIT' -and $lastClass -ne 'WAIT') {
                    $waitOnsets++
                    $currentWaitOnsetMs = $nowMs
                }
                if ($lastClass -eq 'WAIT' -and $class -ne 'WAIT' -and $null -ne $currentWaitOnsetMs) {
                    $waitDurations.Add($nowMs - [double]$currentWaitOnsetMs)
                    $currentWaitOnsetMs = $null
                }
                if ($class -eq 'APPSTARTING' -and $lastClass -ne 'APPSTARTING') {
                    $appStartingOnsets++
                }

                $lastHandle = $handle
                $lastClass = $class
                $lastFlags = $flags
                $haveLast = $true
            }
            $lastTimeMs = $nowMs
        } else {
            $failedSamples++
        }

        $nextSampleTicks += $samplePeriodTicks
        $elapsedTicks = $stopwatch.ElapsedTicks
        if ($nextSampleTicks -le $elapsedTicks) {
            $periodsBehind = [long][Math]::Floor(
                ($elapsedTicks - $nextSampleTicks) / [double]$samplePeriodTicks) + 1L
            $missedDeadlines += $periodsBehind
            $nextSampleTicks += $periodsBehind * $samplePeriodTicks
        }
        $spinTargetTicks = [Math]::Min($nextSampleTicks, $measurementTicks)
        if ($stopwatch.ElapsedTicks -lt $spinTargetTicks) {
            [CursorNative]::SpinUntil($stopwatch, $spinTargetTicks)
        }
    }
}
finally {
    $stopwatch.Stop()
    if ($timerRaised) { [void][CursorNative]::timeEndPeriod(1) }
    if ($priorityRaised) {
        try { $currentProcess.PriorityClass = $originalPriority } catch {}
    }
}

$endTimeMs = [Math]::Min($stopwatch.Elapsed.TotalMilliseconds, $Seconds * 1000.0)
if ($lastClass -eq 'WAIT' -and $null -ne $currentWaitOnsetMs) {
    Add-WaitSlice $waitPerSecond $lastTimeMs $endTimeMs
    $waitDurations.Add($endTimeMs - [double]$currentWaitOnsetMs)
}
$foregroundEnd = Get-ForegroundState

$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$transitionsPath = Join-Path $OutDir "cursor-transitions_$stamp.csv"
$perSecondPath = Join-Path $OutDir "cursor-persec_$stamp.csv"
$summaryPath = Join-Path $OutDir "cursor-summary_$stamp.txt"

$records | Export-Csv -LiteralPath $transitionsPath -NoTypeInformation -Encoding UTF8
$perSecondRows = for ($second = 0; $second -lt $Seconds; $second++) {
    $waitMs = if ($waitPerSecond.ContainsKey($second)) { [double]$waitPerSecond[$second] } else { 0.0 }
    [PSCustomObject]@{
        SecondIndex = $second
        WaitMs = [Math]::Round($waitMs, 1)
    }
}
$perSecondRows | Export-Csv -LiteralPath $perSecondPath -NoTypeInformation -Encoding UTF8

$totalWaitMs = 0.0
if ($waitPerSecond.Count -gt 0) {
    $sum = ($waitPerSecond.Values | Measure-Object -Sum).Sum
    if ($null -ne $sum) { $totalWaitMs = [double]$sum }
}
$waitP50 = Get-NearestRankPercentile $waitDurations 0.50
$waitP95 = Get-NearestRankPercentile $waitDurations 0.95
$waitMax = Get-NearestRankPercentile $waitDurations 1.00
$sampleGapP50 = Get-NearestRankPercentile $sampleGaps 0.50
$sampleGapP95 = Get-NearestRankPercentile $sampleGaps 0.95
$sampleGapMax = Get-NearestRankPercentile $sampleGaps 1.00
$effectiveSampleRate = $sampleCount / [Math]::Max($stopwatch.Elapsed.TotalSeconds, 0.001)
$samplingValid = ($samplePeriodMs -le 5.0 -and $effectiveSampleRate -ge 200.0 -and
                  $sampleGapP95 -le 5.0)
$expectedForegroundFraction = if ($sampleCount -gt 0 -and $ExpectedForegroundProcessId -gt 0) {
    $expectedForegroundSamples / [double]$sampleCount
} else { 0.0 }
$expectedForegroundValid = if ($ExpectedForegroundProcessId -gt 0) {
    ($foregroundStart.ProcessId -eq $ExpectedForegroundProcessId -and
     $foregroundEnd.ProcessId -eq $ExpectedForegroundProcessId -and
     $expectedForegroundFraction -ge 0.95)
} else { $true }
$measurementValid = ($foregroundStart.NonZero -and $foregroundEnd.NonZero -and
                     $suppressedSamples -eq 0 -and $successfulSamples -gt 0 -and
                     $samplingValid -and $expectedForegroundValid)

$summaryLines = @(
    'xltoolrack cursor measurement summary'
    '====================================='
    "Start (wall clock)        : $($startWallTime.ToString('yyyy-MM-dd HH:mm:ss.fff'))"
    "Duration requested        : $Seconds s"
    "Target sample rate        : $SampleRateHz Hz"
    ('Sampling resolution      : {0:N3} ms' -f $samplePeriodMs)
    ('Effective sample rate    : {0:N1} Hz' -f $effectiveSampleRate)
    ('Sample gap P50/P95/max   : {0:N3} / {1:N3} / {2:N3} ms' -f $sampleGapP50, $sampleGapP95, $sampleGapMax)
    "Missed sample deadlines  : $missedDeadlines"
    "Sampler priority          : $priorityDuringMeasurement (restored to $originalPriority)"
    "Sampling checks passed   : $samplingValid"
    "Samples (ok/failed)       : $successfulSamples / $failedSamples"
    "CURSOR_SUPPRESSED samples : $suppressedSamples"
    "Cursor-hidden samples     : $hiddenSamples"
    "Cursor position changes  : $cursorPositionChanges"
    ('Foreground start         : 0x{0:X} pid={1} (nonzero={2})' -f $foregroundStart.Handle, $foregroundStart.ProcessId, $foregroundStart.NonZero)
    ('Foreground end           : 0x{0:X} pid={1} (nonzero={2})' -f $foregroundEnd.Handle, $foregroundEnd.ProcessId, $foregroundEnd.NonZero)
    "Expected foreground PID  : $ExpectedForegroundProcessId"
    ('Expected foreground rate : {0:P2}' -f $expectedForegroundFraction)
    "Foreground checks passed : $expectedForegroundValid"
    "Validity checks passed    : $measurementValid"
    ''
    "WAIT onset count          : $waitOnsets"
    ('WAIT total busy-ms       : {0:N1}' -f $totalWaitMs)
    ('WAIT busy-ms/s            : {0:N2}' -f ($totalWaitMs / [Math]::Max($Seconds, 1)))
    ('WAIT duration P50 (ms)    : {0:N1}' -f $waitP50)
    ('WAIT duration P95 (ms)    : {0:N1}' -f $waitP95)
    ('WAIT duration max (ms)    : {0:N1}' -f $waitMax)
    "APPSTARTING onset count   : $appStartingOnsets"
    ''
    "Transition log            : $transitionsPath"
    "Per-second WAIT log       : $perSecondPath"
)

$summaryLines | Out-File -LiteralPath $summaryPath -Encoding UTF8
$summaryLines | ForEach-Object { Write-Host $_ }

[PSCustomObject]@{
    Valid = $measurementValid
    SamplingValid = $samplingValid
    TargetSampleRateHz = $SampleRateHz
    SamplingResolutionMs = $samplePeriodMs
    EffectiveSampleRateHz = $effectiveSampleRate
    SampleGapP50Ms = $sampleGapP50
    SampleGapP95Ms = $sampleGapP95
    SampleGapMaxMs = $sampleGapMax
    MissedDeadlines = $missedDeadlines
    Priority = $priorityDuringMeasurement
    ForegroundStartNonZero = $foregroundStart.NonZero
    ForegroundEndNonZero = $foregroundEnd.NonZero
    ForegroundStartProcessId = $foregroundStart.ProcessId
    ForegroundEndProcessId = $foregroundEnd.ProcessId
    ExpectedForegroundFraction = $expectedForegroundFraction
    ExpectedForegroundValid = $expectedForegroundValid
    CursorPositionChanges = $cursorPositionChanges
    WaitOnsets = $waitOnsets
    WaitP50Ms = $waitP50
    WaitP95Ms = $waitP95
    WaitMaxMs = $waitMax
    AppStartingOnsets = $appStartingOnsets
    TransitionPath = $transitionsPath
    PerSecondPath = $perSecondPath
    SummaryPath = $summaryPath
}
