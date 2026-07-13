# common/launch.ps1 -- Explorer/silent wrapper around shared launch-core.ps1.
param(
    [Parameter(Mandatory = $true)][string]$Tool,
    [int]$Variant = [int]::MinValue,
    [string]$Action,
    [string]$Target = "",
    [switch]$Gui
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$script:VariantSpecified = $PSBoundParameters.ContainsKey("Variant")
$script:ActionSpecified = $PSBoundParameters.ContainsKey("Action")
. (Join-Path $PSScriptRoot "launch-core.ps1")

# A quoted Windows path ending in '\' can arrive with the final slash changed to '"'.
# A valid Windows path cannot contain '"', so restoring it here is unambiguous.
if ($Target.EndsWith('"')) {
    $Target = $Target.Substring(0, $Target.Length - 1) + "\"
}

function Show-LaunchError {
    param([bool]$GuiMode, [string]$Title, [string]$Message)
    if ($GuiMode) {
        Add-Type -AssemblyName PresentationFramework
        [void][Windows.MessageBox]::Show($Message, "toolrack: " + $Title, "OK", "Error")
        return
    }
    Write-Host ""
    Write-Host ("  ERR  {0}" -f $Title) -ForegroundColor Red
    foreach ($line in ($Message -split "`n")) { Write-Host ("       {0}" -f $line) -ForegroundColor Red }
    if ($env:TOOLRACK_NOPAUSE -ne "1") {
        Write-Host ""
        Write-Host "Press any key to close..." -ForegroundColor DarkGray
        [void][Console]::ReadKey($true)
    }
}

function Invoke-Launch {
    $plan = Resolve-LaunchPlan -ToolDir $Tool -VariantIndex $Variant `
        -VariantSpecified $script:VariantSpecified -ActionId $Action -ActionSpecified $script:ActionSpecified
    if (-not $plan.Ok) {
        Show-LaunchError $Gui.IsPresent "cannot start tool" (($plan.Errors -join "`n") + "`nTool: " + $Tool)
        return 1
    }

    $execution = Invoke-LaunchPlan -Plan $plan -Target $Target -Route Explorer
    if (-not $execution.Ok) {
        Show-LaunchError $Gui.IsPresent $plan.Name ($execution.Errors -join "`n")
        return 1
    }

    if ($plan.Window -eq "console" -and $plan.KeepOpen -and -not $Gui.IsPresent -and $env:TOOLRACK_NOPAUSE -ne "1") {
        Write-Host ""
        Write-Host "===== Done. Press any key to close =====" -ForegroundColor DarkGray
        [void][Console]::ReadKey($true)
    }
    return [int]$execution.ExitCode
}

if ($MyInvocation.InvocationName -ne ".") {
    $exitCode = Invoke-Launch
    exit $exitCode
}
