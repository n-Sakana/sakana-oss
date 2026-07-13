# common/host.ps1 -- hidden resident PowerShell 5.1 Host entry.
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$StateRoot,
    [switch]$SelfTest,
    [switch]$TestBadSource,
    [switch]$TestMode
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

try {
    [Console]::OutputEncoding = New-Object Text.UTF8Encoding($false)
    if ($TestBadSource) {
        try { Add-Type -TypeDefinition "this is not valid C sharp" -ErrorAction Stop }
        catch { throw ("Add-Type failed: " + $_.Exception.Message) }
    }

    $sourceRoot = Join-Path $PSScriptRoot "host"
    $sources = @(
        "BindingModel.cs",
        "NativeMethods.cs",
        "HostPipe.cs",
        "HotkeyManager.cs",
        "MouseGestureHook.cs",
        "ActivationWorker.cs",
        "HostApplication.cs"
    ) | ForEach-Object { Join-Path $sourceRoot $_ }
    foreach ($source in $sources) {
        if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "Host source not found: $source" }
    }
    try {
        Add-Type -Path $sources -ReferencedAssemblies @(
            "System.dll",
            "System.Core.dll",
            "System.Windows.Forms.dll",
            "System.Drawing.dll",
            "System.Runtime.Serialization.dll",
            "System.Xml.dll",
            [Management.Automation.PSObject].Assembly.Location
        ) -ErrorAction Stop
    } catch {
        throw ("Add-Type failed: " + $_.Exception.Message)
    }

    if ($SelfTest) {
        Write-Output ([ToolRack.HostApplication]::SelfTest([IO.Path]::GetFullPath($StateRoot)))
        exit 0
    }
    $code = [ToolRack.HostApplication]::Run([IO.Path]::GetFullPath($StateRoot), [bool]$TestMode)
    exit $code
} catch {
    Write-Error $_.Exception.Message
    exit 1
}
