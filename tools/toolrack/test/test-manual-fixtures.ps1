# test/test-manual-fixtures.ps1 -- generated manual examples are valid and repeatable
. (Join-Path $PSScriptRoot "_assert.ps1")
$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$generator = Join-Path $PSScriptRoot "create-manual-fixtures.ps1"
$fixtureRoot = Join-Path $env:TEMP ("toolrack_manual_" + [guid]::NewGuid().ToString("N"))

try {
    Assert-True (Test-Path -LiteralPath $generator -PathType Leaf) "manual fixture generator exists"
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $generator -Destination $fixtureRoot -NoOpen
    Assert-True ($LASTEXITCODE -eq 0) "manual fixture generator succeeds under PowerShell 5.1"
    Assert-True (Test-Path -LiteralPath (Join-Path $fixtureRoot "README.txt") -PathType Leaf) "fixture README exists"
    Assert-True (Test-Path -LiteralPath (Join-Path $fixtureRoot "ready-mirror.md") -PathType Leaf) "ready mirror exists"
    Assert-True (Test-Path -LiteralPath (Join-Path $fixtureRoot "patches\apply-demo.md") -PathType Leaf) "valid patch example exists"
    Assert-True (Test-Path -LiteralPath (Join-Path $fixtureRoot "patches\reject-path-attack.md") -PathType Leaf) "rejected patch example exists"
    Assert-True (Test-Path -LiteralPath (Join-Path $fixtureRoot "extract-source\sample.docx") -PathType Leaf) "DOCX fixture exists"
    Assert-True (Test-Path -LiteralPath (Join-Path $fixtureRoot "extract-source\sample.xlsx") -PathType Leaf) "XLSX fixture exists"
    Assert-True (Test-Path -LiteralPath (Join-Path $fixtureRoot "extract-source\sample.pptx") -PathType Leaf) "PPTX fixture exists"
    Assert-True (Test-Path -LiteralPath (Join-Path $fixtureRoot "extract-source\sample.pdf") -PathType Leaf) "PDF fixture exists"

    . (Join-Path $root "tool\md-patch\format.ps1")
    . (Join-Path $root "tool\md-patch\engine.ps1")
    $patchTarget = Join-Path $fixtureRoot "patch-target"
    $patchText = [IO.File]::ReadAllText((Join-Path $fixtureRoot "patches\apply-demo.md"))
    $prepared = Prepare-PatchApplication (Read-PatchDocument $patchText) $patchTarget
    $originalApp = [IO.File]::ReadAllBytes((Join-Path $patchTarget "src\app.ps1"))
    $originalAsset = [IO.File]::ReadAllBytes((Join-Path $patchTarget "assets\logo.png"))
    $result = Apply-PreparedPatch $prepared (Join-Path $fixtureRoot "patch-backups")
    Assert-True ($result.Applied -eq 5) "valid manual patch applies all operations"
    Assert-True ([IO.File]::ReadAllText((Join-Path $patchTarget "src\app.ps1")).Contains('$name = "toolrack"')) "manual patch modifies the named source line"
    Assert-True (Test-Path -LiteralPath (Join-Path $patchTarget "created.txt") -PathType Leaf) "manual patch creates a text file"
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $patchTarget "obsolete.txt"))) "manual patch deletes the obsolete file"
    Assert-True (Test-Path -LiteralPath (Join-Path $patchTarget "empty-created") -PathType Container) "manual patch creates an empty directory"
    $updatedAsset = [IO.File]::ReadAllBytes((Join-Path $patchTarget "assets\logo.png"))
    $backupApp = [IO.File]::ReadAllBytes((Join-Path $result.BackupPath "data\src\app.ps1"))
    $backupAsset = [IO.File]::ReadAllBytes((Join-Path $result.BackupPath "data\assets\logo.png"))
    Assert-True (-not (Test-PatchByteArraysEqual $originalAsset $updatedAsset)) "manual patch replaces a binary asset"
    Assert-True (Test-PatchByteArraysEqual $originalApp $backupApp) "manual patch backs up the original source bytes"
    Assert-True (Test-PatchByteArraysEqual $originalAsset $backupAsset) "manual patch backs up the original binary bytes"

    . (Join-Path $root "tool\md-extract\engine.ps1")
    $docx = Read-WordOpenXmlFile (Join-Path $fixtureRoot "extract-source\sample.docx")
    $xlsx = Read-ExcelOpenXmlFile (Join-Path $fixtureRoot "extract-source\sample.xlsx")
    $pptx = Read-PowerPointOpenXmlFile (Join-Path $fixtureRoot "extract-source\sample.pptx")
    Assert-True ($docx.Content -like "*DOCX MANUAL FIXTURE*") "generated DOCX is extractable"
    Assert-True ($xlsx.Content -like "*XLSX MANUAL FIXTURE*") "generated XLSX is extractable"
    Assert-True ($pptx.Content -like "*PPTX MANUAL FIXTURE*" -and $pptx.Content -like "*NOTES MANUAL FIXTURE*") "generated PPTX slides and notes are extractable"

    $sourceBytes = [IO.File]::ReadAllBytes($generator)
    Assert-True (@($sourceBytes | Where-Object { $_ -gt 127 }).Count -eq 0) "manual fixture generator is ASCII source"
} finally {
    Remove-Item -LiteralPath $fixtureRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Exit-Test
