# test/test-timer.ps1 -- timer builds its window without errors (smoke)
. (Join-Path $PSScriptRoot "_assert.ps1")
$root = Split-Path $PSScriptRoot -Parent
$main = Join-Path $root "tool\timer\main.ps1"
$env:TOOLRACK_SMOKE = "1"
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main -Target "C:\ignored" -Minutes 5
Assert-True ($LASTEXITCODE -eq 0) "preset mode builds cleanly"
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main
Assert-True ($LASTEXITCODE -eq 0) "custom (setup) mode builds cleanly"
$env:TOOLRACK_SMOKE = ""
Exit-Test
