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
    "HKCU:\Software\Classes\*\shell\PromptPack",
    "HKCU:\Software\Classes\Directory\shell\PromptPack"
)

$command = '"' + $runner + '" "%1"'

foreach ($target in $targets) {
    $commandKey = Join-Path -Path $target -ChildPath "command"
    [void](New-Item -Path $target -Force)
    [void](New-Item -Path $commandKey -Force)
    Set-Item -Path $target -Value "Run PromptPack"
    New-ItemProperty -Path $target -Name "MUIVerb" -Value "Run PromptPack" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $target -Name "Icon" -Value $runner -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $target -Name "InstallPath" -Value ($root + [System.IO.Path]::DirectorySeparatorChar) -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $target -Name "ScriptPath" -Value $runner -PropertyType String -Force | Out-Null
    Set-Item -Path $commandKey -Value $command
}

Write-Host "PromptPack context menu installed for current user." -ForegroundColor Green
Write-Host "Registry keys:" -ForegroundColor Gray
foreach ($target in $targets) {
    Write-Host ("- " + $target.Replace("HKCU:", "HKCU")) -ForegroundColor Gray
}
