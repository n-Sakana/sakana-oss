# test/test-tree.ps1
. (Join-Path $PSScriptRoot "_assert.ps1")
$root = Split-Path $PSScriptRoot -Parent
$main = Join-Path $root "tool\tree\main.ps1"
$env:TOOLRACK_NOPAUSE = "1"
$fx = Join-Path $env:TEMP ("toolrack_tr_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force (Join-Path $fx "a\b\c") | Out-Null
try {
"1" | Set-Content (Join-Path $fx "top.txt")
"2" | Set-Content (Join-Path $fx "a\mid.txt")
"3" | Set-Content (Join-Path $fx "a\b\c\deep.txt")
New-Item -ItemType Directory -Force (Join-Path $fx "node_modules\junk") | Out-Null

Assert-True (Test-Path -LiteralPath $main -PathType Leaf) "tree main.ps1 exists"
if (-not (Test-Path -LiteralPath $main -PathType Leaf)) {
    Exit-Test
}
$mainText = [System.IO.File]::ReadAllText($main)
Assert-True ($mainText -match 'function Get-TreeItems') "tree owns a prunable directory walker"
Assert-True ($mainText -notmatch 'Get-ChildItem[^\r\n]*-Recurse') "excluded directories are pruned before recursive enumeration"
function Invoke-Tree {
    param([string[]]$ExtraArgs)
    $before = @(Get-ChildItem (Join-Path $root "output") -Filter "tree_*.md" -ErrorAction SilentlyContinue)
    $runOutput = @(& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main -Target $fx -NoOpen @ExtraArgs 2>&1)
    $code = $LASTEXITCODE
    foreach ($line in $runOutput) { Write-Host ([string]$line) }
    if ($code -ne 0) { return $null }
    $after = @(Get-ChildItem (Join-Path $root "output") -Filter "tree_*.md" -ErrorAction SilentlyContinue)
    $new = @($after | Where-Object { $before.Name -notcontains $_.Name })
    if ($new.Count -ne 1) { return $null }
    return (Get-Content -LiteralPath $new[0].FullName -Raw)
}

$md = Invoke-Tree @("-Depth", "1")
Assert-True ($null -ne $md) "depth1: output md produced"
Assert-True ($md -like "*top.txt*") "depth1: shows top-level file"
Assert-True ($md -notlike "*mid.txt*") "depth1: hides depth-2 file"
Assert-True ($md -notlike "*node_modules*") "excluded dirs hidden"

$md = Invoke-Tree @("-Depth", "0")
Assert-True ($md -like "*deep.txt*") "unlimited: shows deepest file"

$md = Invoke-Tree @("-Depth", "0", "-DirsOnly")
Assert-True ($md -notlike "*deep.txt*") "dirs-only: no files"
Assert-True ($md -like "*- c/*") "dirs-only: still shows folders"

# behavior check: without -NoOpen, Tree asks Open-Folder for the shared output directory
$openMark = Join-Path $fx "opened-folder.txt"
$env:TOOLRACK_TEST_OPEN_FOLDER_FILE = $openMark
$runOutput = @(& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main -Target $fx -Depth 1 2>&1)
$openCode = $LASTEXITCODE
foreach ($line in $runOutput) { Write-Host ([string]$line) }
Assert-True ($openCode -eq 0) "tree completes when opening the result folder"
Assert-True (Test-Path -LiteralPath $openMark) "tree requests a folder open"
if (Test-Path -LiteralPath $openMark) {
    $opened = [System.IO.File]::ReadAllText($openMark)
    Assert-True ($opened -eq (Join-Path $root "output")) "tree opens the shared output folder"
}
$env:TOOLRACK_TEST_OPEN_FOLDER_FILE = ""

} finally {
    Remove-Item -LiteralPath $fx -Recurse -Force -ErrorAction SilentlyContinue
}
Exit-Test
