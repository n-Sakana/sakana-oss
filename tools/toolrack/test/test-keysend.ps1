# test/test-keysend.ps1 -- resident loop exits at MaxRun; no key sent when idle threshold not met
. (Join-Path $PSScriptRoot "_assert.ps1")
$root = Split-Path $PSScriptRoot -Parent
$main = Join-Path $root "tool\keysend\main.ps1"
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$out = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main -Target "C:\" -Idle 9999 -MaxRun 2 -Key 0x07 2>&1
$sw.Stop()
$text = @($out | ForEach-Object { [string]$_ })
Assert-True ($sw.Elapsed.TotalSeconds -lt 15) "exits near MaxRun (not hanging)"
Assert-Contains $text "*Max runtime reached*" "reports MaxRun exit"
Assert-Contains $text "*Total sent: 0*" "no key sent below idle threshold"

$badOut = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main -Target "C:\" -Idle 9999 -MaxRun 1 -Key nope 2>&1
$badText = @($badOut | ForEach-Object { [string]$_ })
Assert-True ($LASTEXITCODE -eq 1) "invalid key -> exit 1"
Assert-Contains $badText "*Invalid key code*" "invalid key has a clean error"
Exit-Test
