# common/uninstall.ps1 -- stop Host and remove only ToolRack-owned state.
. (Join-Path $PSScriptRoot "install.ps1")

function Build-UninstallRegistryLines {
    param([string[]]$Owned, $Context)
    $lines = New-Object "System.Collections.Generic.List[string]"
    foreach ($line in @(Build-RegLines @() $Owned)) { $lines.Add([string]$line) }
    $lines.Add("[" + $Context.RunRegistryPath + "]")
    $lines.Add('"' + (ConvertTo-RegString $Context.RunValueName) + '"=-')
    $lines.Add("")
    return $lines.ToArray()
}

function Invoke-Uninstall {
    [CmdletBinding()]
    param($InstallContext = $null)
    $ErrorActionPreference = "Stop"
    Write-Host ""
    Write-Host "=== toolrack uninstall ===" -ForegroundColor Cyan
    Write-Host ""

    $context = $null
    $previous = $null
    $temporaryRoot = Join-Path $env:TEMP ("toolrack_uninstall_" + [guid]::NewGuid().ToString("N"))
    try {
        $context = Resolve-InstallContext $InstallContext
        $previous = Set-InstallScriptContext $context
        New-Item -ItemType Directory -Force -Path $temporaryRoot | Out-Null

        $activeState = Get-ActiveStateRoot $context.LocalRoot
        if ($null -ne $activeState) {
            $stopped = Stop-HostInternal $activeState $context.Root
            if ($stopped.WasRunning -and -not $stopped.Stopped) { throw "Host could not be stopped" }
            if ($stopped.WasRunning) { Write-Host ("  OK   stopped Host PID {0}" -f $stopped.ProcessId) -ForegroundColor Green }
        }

        $owned = @(Get-OwnedMenuRootPaths)
        foreach ($path in $owned) { Write-Host ("  OK   found {0}" -f $path) -ForegroundColor Green }
        $lines = Build-UninstallRegistryLines $owned $context
        if (-not (Import-RegistryLines $lines $temporaryRoot)) { throw "registry import failed" }

        Remove-OwnedLocalRoot $context.LocalRoot
        Write-Host ""
        Write-Host ("Removed {0} ToolRack menu page(s), Host autostart, and local Host state." -f $owned.Count) -ForegroundColor Green
        return [pscustomobject]@{ Ok = $true; RemovedMenus = $owned.Count }
    } catch {
        Write-Host ("  ERROR {0}" -f $_.Exception.Message) -ForegroundColor Red
        return [pscustomobject]@{ Ok = $false; Error = $_.Exception.Message }
    } finally {
        if ($null -ne $previous) { Restore-InstallScriptContext $previous }
        Remove-Item -LiteralPath $temporaryRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

if ($MyInvocation.InvocationName -ne ".") {
    $uninstallResult = Invoke-Uninstall
    if (-not $uninstallResult.Ok) { exit 1 }
    exit 0
}
