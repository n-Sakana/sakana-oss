[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$targets = @(
    "Software\Classes\*\shell\PromptPack",
    "Software\Classes\Directory\shell\PromptPack"
)

$registryRoot = [Microsoft.Win32.Registry]::CurrentUser

foreach ($target in $targets) {
    $key = $registryRoot.OpenSubKey($target)
    if ($null -ne $key) {
        $key.Close()
        $registryRoot.DeleteSubKeyTree($target)
    }
}

Write-Host "PromptPack context menu removed for current user." -ForegroundColor Green
