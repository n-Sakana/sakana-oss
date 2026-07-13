# test/test-timestamp.ps1
. (Join-Path $PSScriptRoot "_assert.ps1")
$root = Split-Path $PSScriptRoot -Parent
$main = Join-Path $root "tool\timestamp\main.ps1"
$env:TOOLRACK_NOPAUSE = "1"
$fx = Join-Path $env:TEMP ("toolrack_ts_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force $fx | Out-Null
try {

$f = Join-Path $fx "note.txt"; "x" | Set-Content $f
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main -Target $f -Format yyyyMMdd
$stamp = Get-Date -Format "yyyyMMdd"
Assert-True (Test-Path (Join-Path $fx "note_$stamp.txt")) "file renamed with _yyyyMMdd (ext kept)"

$d = Join-Path $fx "proj"; New-Item -ItemType Directory $d | Out-Null
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main -Target $d -Format yyyyMMddHHmmss
$renamed = @(Get-ChildItem $fx -Directory | Where-Object { $_.Name -match "^proj_\d{14}$" })
Assert-True ($renamed.Count -eq 1) "folder renamed with _yyyyMMddHHmmss"

# collision: target already exists -> abort, exit 1, source untouched
$a = Join-Path $fx "dup.txt"; "a" | Set-Content $a
$b = Join-Path $fx "dup_$stamp.txt"; "b" | Set-Content $b
$errorMark = Join-Path $fx "timestamp-error.txt"
$env:TOOLRACK_TEST_ERROR_FILE = $errorMark
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main -Target $a -Format yyyyMMdd
Assert-True ($LASTEXITCODE -eq 1) "collision -> exit 1"
Assert-True (Test-Path $a) "source untouched on collision"
Assert-True (Test-Path -LiteralPath $errorMark) "collision error is surfaced outside the hidden console"
if (Test-Path -LiteralPath $errorMark) {
    Assert-True ([System.IO.File]::ReadAllText($errorMark) -like "*Already exists*") "collision error explains the reason"
}

# missing target also uses the visible error path
Remove-Item -LiteralPath $errorMark -Force -ErrorAction SilentlyContinue
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main -Target (Join-Path $fx "missing.txt") -Format yyyyMMdd
Assert-True ($LASTEXITCODE -eq 1) "missing target -> exit 1"
Assert-True (Test-Path -LiteralPath $errorMark) "missing-target error is surfaced"
if (Test-Path -LiteralPath $errorMark) {
    Assert-True ([System.IO.File]::ReadAllText($errorMark) -like "*Path not found*") "missing-target error explains the reason"
}

# the same error path works through silent.vbs
Remove-Item -LiteralPath $errorMark -Force -ErrorAction SilentlyContinue
$hiddenDup = Join-Path $fx "hidden-dup.txt"; "h" | Set-Content $hiddenDup
$hiddenDupResult = Join-Path $fx "hidden-dup_$stamp.txt"; "existing" | Set-Content $hiddenDupResult
$vbs = Join-Path $root "common\silent.vbs"
& wscript.exe $vbs (Split-Path $main -Parent) 0 $hiddenDup
for ($i = 0; $i -lt 20 -and -not (Test-Path -LiteralPath $errorMark); $i++) {
    Start-Sleep -Milliseconds 100
}
Assert-True (Test-Path -LiteralPath $errorMark) "hidden collision produces a visible error notification"

# real no-console route: silent.vbs -> launch.ps1 -> hidden timestamp process
$silentSource = Join-Path $fx "silent.txt"; "s" | Set-Content $silentSource
& wscript.exe $vbs (Split-Path $main -Parent) 0 $silentSource
$silentResult = Join-Path $fx "silent_$stamp.txt"
for ($i = 0; $i -lt 40 -and -not (Test-Path -LiteralPath $silentResult); $i++) {
    Start-Sleep -Milliseconds 250
}
Assert-True (Test-Path -LiteralPath $silentResult) "timestamp runs through the hidden VBS route"
$env:TOOLRACK_TEST_ERROR_FILE = ""

} finally {
    Remove-Item -LiteralPath $fx -Recurse -Force -ErrorAction SilentlyContinue
}
Exit-Test
