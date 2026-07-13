<#
.SYNOPSIS
  Keeps one test-owned Excel window foreground while moving the pointer over
  its worksheet area for a controlled cursor measurement.

.DESCRIPTION
  This helper is intentionally opt-in. It does not click, type, or modify the
  workbook. It sends only absolute mouse-move input between safe points in the
  lower-middle portion of the Excel client area. Use it only with explicit
  approval from the interactive desktop owner.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateRange(1, 2147483647)]
    [int]$ExcelPid,

    [ValidateRange(1, 1800)]
    [int]$Seconds = 120,

    [ValidateRange(50, 2000)]
    [int]$IntervalMilliseconds = 180,

    [string]$ReadyFile,

    [string]$StopFile,

    [string]$SummaryPath = (Join-Path $env:TEMP 'xltoolrack-controlled-activity.json')
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

if (-not ('XltrControlledInput' -as [type])) {
    Add-Type @'
using System;
using System.Runtime.InteropServices;
using System.Threading;

public static class XltrControlledInput
{
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MOUSEINPUT
    {
        public int dx;
        public int dy;
        public uint mouseData;
        public uint dwFlags;
        public uint time;
        public UIntPtr dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct INPUT
    {
        public uint type;
        public MOUSEINPUT mouseInput;
    }

    [DllImport("user32.dll")]
    private static extern uint SendInput(
        uint inputCount, INPUT[] inputs, int inputSize);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hwnd, out RECT rectangle);

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

    [DllImport("user32.dll")]
    private static extern int GetSystemMetrics(int index);

    public static uint ForegroundProcessId()
    {
        uint processId;
        GetWindowThreadProcessId(GetForegroundWindow(), out processId);
        return processId;
    }

    public static bool ActivateWindow(IntPtr hwnd)
    {
        uint targetProcessId;
        GetWindowThreadProcessId(hwnd, out targetProcessId);
        uint unused;
        IntPtr foreground = GetForegroundWindow();
        uint foregroundThread = GetWindowThreadProcessId(foreground, out unused);
        uint currentThread = GetCurrentThreadId();
        bool attached = foregroundThread != 0 && foregroundThread != currentThread &&
                        AttachThreadInput(currentThread, foregroundThread, true);
        try {
            ShowWindowAsync(hwnd, 9); // SW_RESTORE
            BringWindowToTop(hwnd);
            SetForegroundWindow(hwnd);
        }
        finally {
            if (attached) {
                AttachThreadInput(currentThread, foregroundThread, false);
            }
        }
        return ForegroundProcessId() == targetProcessId;
    }

    public static bool MovePointer(int screenX, int screenY)
    {
        const int SM_XVIRTUALSCREEN = 76;
        const int SM_YVIRTUALSCREEN = 77;
        const int SM_CXVIRTUALSCREEN = 78;
        const int SM_CYVIRTUALSCREEN = 79;
        const uint INPUT_MOUSE = 0;
        const uint MOUSEEVENTF_MOVE = 0x0001;
        const uint MOUSEEVENTF_VIRTUALDESK = 0x4000;
        const uint MOUSEEVENTF_ABSOLUTE = 0x8000;

        int left = GetSystemMetrics(SM_XVIRTUALSCREEN);
        int top = GetSystemMetrics(SM_YVIRTUALSCREEN);
        int width = Math.Max(2, GetSystemMetrics(SM_CXVIRTUALSCREEN));
        int height = Math.Max(2, GetSystemMetrics(SM_CYVIRTUALSCREEN));
        int boundedX = Math.Max(left, Math.Min(left + width - 1, screenX));
        int boundedY = Math.Max(top, Math.Min(top + height - 1, screenY));
        int normalizedX = (int)Math.Round((boundedX - left) * 65535.0 / (width - 1));
        int normalizedY = (int)Math.Round((boundedY - top) * 65535.0 / (height - 1));

        INPUT input = new INPUT();
        input.type = INPUT_MOUSE;
        input.mouseInput.dx = normalizedX;
        input.mouseInput.dy = normalizedY;
        input.mouseInput.dwFlags = MOUSEEVENTF_MOVE |
                                   MOUSEEVENTF_VIRTUALDESK |
                                   MOUSEEVENTF_ABSOLUTE;
        return SendInput(1, new INPUT[] { input }, Marshal.SizeOf(typeof(INPUT))) == 1;
    }
}
'@
}

$process = Get-Process -Id $ExcelPid -ErrorAction Stop
if ($process.ProcessName -ine 'EXCEL') {
    throw "PID $ExcelPid is not Excel."
}
$window = [IntPtr]$process.MainWindowHandle
if ($window -eq [IntPtr]::Zero) {
    throw "Excel PID $ExcelPid has no visible main window."
}

foreach ($path in @($ReadyFile, $SummaryPath)) {
    if ([string]::IsNullOrWhiteSpace($path)) { continue }
    $parent = Split-Path -Parent $path
    if (-not [string]::IsNullOrWhiteSpace($parent) -and
        -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
}
if (-not [string]::IsNullOrWhiteSpace($ReadyFile) -and
    (Test-Path -LiteralPath $ReadyFile)) {
    Remove-Item -LiteralPath $ReadyFile -Force
}

$deadline = (Get-Date).AddSeconds($Seconds)
$foregroundSamples = 0
$samples = 0
$movesSent = 0
$activationAttempts = 0
$activationSuccesses = 0
$pointIndex = 0
$startedAt = Get-Date
$stoppedByFile = $false

try {
    while ((Get-Date) -lt $deadline) {
        if (-not [string]::IsNullOrWhiteSpace($StopFile) -and
            (Test-Path -LiteralPath $StopFile)) {
            $stoppedByFile = $true
            break
        }

        $process.Refresh()
        $window = [IntPtr]$process.MainWindowHandle
        if ($window -eq [IntPtr]::Zero) {
            throw "Excel PID $ExcelPid lost its visible main window."
        }

        $samples++
        if ([int][XltrControlledInput]::ForegroundProcessId() -ne $ExcelPid) {
            $activationAttempts++
            if ([XltrControlledInput]::ActivateWindow($window)) {
                $activationSuccesses++
            }
        }
        if ([int][XltrControlledInput]::ForegroundProcessId() -eq $ExcelPid) {
            $foregroundSamples++
        }

        $rectangle = New-Object XltrControlledInput+RECT
        if (-not [XltrControlledInput]::GetWindowRect($window, [ref]$rectangle)) {
            throw "GetWindowRect failed for Excel PID $ExcelPid."
        }
        $width = $rectangle.Right - $rectangle.Left
        $height = $rectangle.Bottom - $rectangle.Top
        if ($width -lt 400 -or $height -lt 300) {
            throw "Excel PID $ExcelPid window is too small for controlled activity ($width x $height)."
        }

        # Lower-middle client points stay below the ribbon/formula bar and
        # above the sheet tabs/status bar in a normally sized Excel window.
        $points = @(
            @(0.35, 0.58),
            @(0.50, 0.68),
            @(0.65, 0.58),
            @(0.50, 0.76)
        )
        $point = $points[$pointIndex % $points.Count]
        $pointIndex++
        $x = $rectangle.Left + [int][Math]::Round($width * [double]$point[0])
        $y = $rectangle.Top + [int][Math]::Round($height * [double]$point[1])
        if ([XltrControlledInput]::MovePointer($x, $y)) {
            $movesSent++
        }

        if ($movesSent -eq 1 -and -not [string]::IsNullOrWhiteSpace($ReadyFile)) {
            [IO.File]::WriteAllText($ReadyFile, "excelPid=$ExcelPid`r`nstarted=$($startedAt.ToString('o'))`r`n")
        }
        Start-Sleep -Milliseconds $IntervalMilliseconds
    }
}
finally {
    $endedAt = Get-Date
    $foregroundFraction = if ($samples -gt 0) {
        $foregroundSamples / [double]$samples
    } else { 0.0 }
    $summary = [ordered]@{
        ExcelPid = $ExcelPid
        StartedAt = $startedAt.ToString('o')
        EndedAt = $endedAt.ToString('o')
        RequestedSeconds = $Seconds
        ElapsedSeconds = ($endedAt - $startedAt).TotalSeconds
        IntervalMilliseconds = $IntervalMilliseconds
        Samples = $samples
        ForegroundSamples = $foregroundSamples
        ForegroundFraction = $foregroundFraction
        MovesSent = $movesSent
        ActivationAttempts = $activationAttempts
        ActivationSuccesses = $activationSuccesses
        StoppedByFile = $stoppedByFile
        InputMode = 'SendInput mouse-move only; no clicks or keys'
    }
    $summary | ConvertTo-Json | Out-File -LiteralPath $SummaryPath -Encoding UTF8
    $summary | ConvertTo-Json -Compress | Write-Output
}
