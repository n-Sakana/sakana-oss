$ErrorActionPreference = 'Stop'
$script:Failed = 0

# The FE runs a 1-second OnTime pump while jobs are active, so COM calls from
# the tests can momentarily collide with a tick and get rejected as busy
# (0x800AC472 / RPC_E_SERVERCALL_RETRYLATER). Register the standard OLE
# message filter so those calls retry transparently instead of failing.
if (-not ('XltrMessageFilter' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

[ComImport, Guid("00000016-0000-0000-C000-000000000046"),
 InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IOleMessageFilter {
    [PreserveSig] int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
    [PreserveSig] int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
    [PreserveSig] int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
}

public class XltrMessageFilter : IOleMessageFilter {
    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);

    public static void Register() {
        IOleMessageFilter old;
        CoRegisterMessageFilter(new XltrMessageFilter(), out old);
    }

    int IOleMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo) {
        return 0; // SERVERCALL_ISHANDLED
    }

    int IOleMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType) {
        if (dwTickCount < 30000) { return 150; } // retry after 150ms for up to ~30s
        return -1; // give up: surface the error
    }

    int IOleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType) {
        return 2; // PENDINGMSG_WAITDEFPROCESS
    }
}
'@
}
try { [XltrMessageFilter]::Register() } catch {}

function Assert-True {
    param([bool]$Cond, [string]$Msg)
    if ($Cond) {
        Write-Host "  ok   $Msg" -ForegroundColor Green
    } else {
        $script:Failed++
        Write-Host "  FAIL $Msg" -ForegroundColor Red
    }
}

function Assert-Equal {
    param($Expected, $Actual, [string]$Msg)
    Assert-True ($Expected -eq $Actual) "$Msg (expected=$Expected actual=$Actual)"
}

function Exit-Test {
    if ($script:Failed) {
        Write-Host "$script:Failed failure(s)" -ForegroundColor Red
        exit 1
    }
    Write-Host 'all passed' -ForegroundColor Green
    exit 0
}

function Release-Com {
    param($Object)
    if ($null -eq $Object) { return }
    try {
        if ([Runtime.InteropServices.Marshal]::IsComObject($Object)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
        }
    } catch {}
}

function QM {
    param([string]$Workbook, [string]$Procedure)
    return "'" + ($Workbook -replace "'", "''") + "'!" + $Procedure
}

# Application.Run with busy retry. While the FE pump or a dispatch is mid
# execution Excel surfaces 0x800AC472 / RETRYLATER / CALL_REJECTED to
# automation callers; those are momentary, so retry for up to 25 seconds.
function Invoke-XlRun {
    param($Excel, [string]$Macro, $A1, $A2, $A3)
    $busy = @(-2146777998, -2147417846, -2147418111)
    $deadline = (Get-Date).AddSeconds(25)
    while ($true) {
        try {
            if ($PSBoundParameters.ContainsKey('A3')) { return $Excel.Run($Macro, $A1, $A2, $A3) }
            elseif ($PSBoundParameters.ContainsKey('A2')) { return $Excel.Run($Macro, $A1, $A2) }
            elseif ($PSBoundParameters.ContainsKey('A1')) { return $Excel.Run($Macro, $A1) }
            else { return $Excel.Run($Macro) }
        } catch {
            $hr = $null
            if ($null -ne $_.Exception.InnerException) { $hr = $_.Exception.InnerException.HResult }
            if ($null -eq $hr -or $hr -eq 0) { $hr = $_.Exception.HResult }
            if (($busy -contains [int]$hr) -and ((Get-Date) -lt $deadline)) {
                Start-Sleep -Milliseconds 200
                continue
            }
            throw
        }
    }
}

function New-TestDirectory {
    $path = Join-Path $env:TEMP ('xltoolrack_test_' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force -Path $path | Out-Null
    return $path
}

function Build-Addin {
    param([string]$Dir, [ValidateSet('xlam','xlsm','all')][string]$Format = 'xlam')
    $root = Split-Path $PSScriptRoot -Parent
    & (Join-Path $root 'scripts\Build-Addin.ps1') -OutputFormat $Format -OutputDirectory $Dir | Out-Null
    if ($Format -eq 'xlsm') { return (Join-Path $Dir 'xltoolrack-test.xlsm') }
    return (Join-Path $Dir 'xltoolrack.xlam')
}

function Open-TestAddin {
    param([string]$Path)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $true
    $previousSecurity = $excel.AutomationSecurity
    $excel.AutomationSecurity = 1
    try {
        $workbook = $excel.Workbooks.Open($Path, 0, $false)
    } finally {
        $excel.AutomationSecurity = $previousSecurity
    }
    return [pscustomobject]@{ Excel = $excel; Workbook = $workbook }
}

function Close-TestExcel {
    param($Excel, $Workbook)
    if ($null -ne $Excel -and $null -ne $Workbook) {
        try {
            $bookName = [string]$Workbook.Name
            $null = Invoke-XlRun $Excel (QM $bookName 'JobTest_StopAll')
            Start-Sleep -Seconds 2
            for ($i = 0; $i -lt 3; $i++) {
                $remaining = [int](Invoke-XlRun $Excel (QM $bookName 'JobTest_RunningCount'))
                if ($remaining -eq 0) { break }
                Start-Sleep -Seconds 1
            }
        } catch {}
    }
    if ($null -ne $Workbook) {
        try { $Workbook.Close($false) } catch {}
        Release-Com $Workbook
    }
    if ($null -ne $Excel) {
        try { $Excel.Quit() } catch {}
        Release-Com $Excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

function Remove-TestDirectory {
    param([string]$Path)
    if (-not [string]::IsNullOrWhiteSpace($Path) -and (Test-Path -LiteralPath $Path)) {
        Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue
    }
}
