[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$targets = @(
    "HKCU:\Software\Classes\*\shell\PromptPack",
    "HKCU:\Software\Classes\Directory\shell\PromptPack"
)

foreach ($target in $targets) {
    if (Test-Path -Path $target) {
        Remove-Item -Path $target -Recurse -Force
    }
}

Write-Host "PromptPack context menu removed for current user." -ForegroundColor Green
