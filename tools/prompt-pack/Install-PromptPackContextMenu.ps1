[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSCommandPath
$runner = Join-Path -Path $root -ChildPath "PromptPack.bat"

if (-not (Test-Path -LiteralPath $runner -PathType Leaf)) {
    throw "PromptPack.bat was not found: $runner"
}

$targets = @(
    "Software\Classes\*\shell\PromptPack",
    "Software\Classes\Directory\shell\PromptPack"
)

$command = '"' + $runner + '" "%1"'
$registryRoot = [Microsoft.Win32.Registry]::CurrentUser

foreach ($target in $targets) {
    $key = $registryRoot.CreateSubKey($target)
    if ($null -eq $key) {
        throw "Failed to create registry key: HKCU\$target"
    }

    try {
        $key.SetValue("", "Run PromptPack", [Microsoft.Win32.RegistryValueKind]::String)
        $key.SetValue("MUIVerb", "Run PromptPack", [Microsoft.Win32.RegistryValueKind]::String)
        $key.SetValue("Icon", $runner, [Microsoft.Win32.RegistryValueKind]::String)
        $key.SetValue("InstallPath", ($root + [System.IO.Path]::DirectorySeparatorChar), [Microsoft.Win32.RegistryValueKind]::String)
        $key.SetValue("ScriptPath", $runner, [Microsoft.Win32.RegistryValueKind]::String)
    }
    finally {
        $key.Close()
    }

    $commandKey = $registryRoot.CreateSubKey($target + "\command")
    if ($null -eq $commandKey) {
        throw "Failed to create registry key: HKCU\$target\command"
    }

    try {
        $commandKey.SetValue("", $command, [Microsoft.Win32.RegistryValueKind]::String)
    }
    finally {
        $commandKey.Close()
    }
}

Write-Host "PromptPack context menu installed for current user." -ForegroundColor Green
Write-Host "Registry keys:" -ForegroundColor Gray
foreach ($target in $targets) {
    Write-Host ("- HKCU\" + $target) -ForegroundColor Gray
}
