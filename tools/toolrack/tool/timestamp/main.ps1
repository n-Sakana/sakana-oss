# Timestamp -- append _<stamp> to a file or folder name.
param(
    [string]$Target = "$PWD",
    [ValidateSet("yyyyMMdd", "yyyyMMddHHmmss")][string]$Format = "yyyyMMdd"
)
. (Join-Path $PSScriptRoot "..\..\common\ui.ps1")

if (-not (Test-Path -LiteralPath $Target)) {
    $message = "Path not found: $Target"
    Write-Err $message
    Show-ErrorDialog -Title "Timestamp" -Message $message
    exit 1
}

Start-UI -Title "Timestamp"
$item = Get-Item -LiteralPath $Target
$stamp = Get-Date -Format $Format
if ($item.PSIsContainer) { $newName = "$($item.Name)_$stamp" }
else { $newName = "$($item.BaseName)_$stamp$($item.Extension)" }
$newPath = Join-Path (Split-Path $Target -Parent) $newName
Write-Dim ("From: {0}" -f $item.Name)
Write-Dim ("To:   {0}" -f $newName)
if (Test-Path -LiteralPath $newPath) {
    $message = "Already exists: $newName"
    Write-Err $message
    Show-ErrorDialog -Title "Timestamp" -Message $message
    Stop-UI
    exit 1
}
try {
    Rename-Item -LiteralPath $Target -NewName $newName -ErrorAction Stop
    Write-Ok "Renamed"
} catch {
    $message = $_.Exception.Message
    Write-Err $message
    Show-ErrorDialog -Title "Timestamp" -Message $message
    Stop-UI
    exit 1
}
Stop-UI
