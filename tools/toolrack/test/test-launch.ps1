# test/test-launch.ps1 -- console-type launch through the real launch.ps1 process
. (Join-Path $PSScriptRoot "_assert.ps1")
$root = Split-Path $PSScriptRoot -Parent
$launch = Join-Path $root "common\launch.ps1"
$env:TOOLRACK_NOPAUSE = "1"
$fx = Join-Path $env:TEMP ("toolrack_lt_" + [guid]::NewGuid().ToString("N"))
$echoDir = Join-Path $fx "echo-tool"
New-Item -ItemType Directory -Force $echoDir | Out-Null
try {
$mark = Join-Path $fx "mark.txt"
@'
param([string]$Target, [string]$A = "", [string]$B = "")
[System.IO.File]::WriteAllText($env:TOOLRACK_TEST_MARK, "T=$Target A=$A B=$B")
exit 7
'@ | Set-Content -LiteralPath (Join-Path $echoDir "main.ps1") -Encoding Ascii
'{"schema":1,"id":"echo-tool","name":"Echo","on":["folder"],"run":{"type":"powershell","entry":"main.ps1","keep_open":false},"variants":[{"label":"AB","args":["-A","hello world","-B","2"]}]}' |
    Set-Content -LiteralPath (Join-Path $echoDir "tool.json") -Encoding Ascii
$env:TOOLRACK_TEST_MARK = $mark

# variant 0 with spaces in target and args
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $launch -Tool $echoDir -Variant 0 -Target "C:\spa ce"
Assert-True ($LASTEXITCODE -eq 7) "tool exit code is passed through"
$content = [System.IO.File]::ReadAllText($mark)
Assert-True ($content -eq "T=C:\spa ce A=hello world B=2") ("args intact: got '" + $content + "'")

# flat launch (-Variant -1): no variant args
Remove-Item $mark
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $launch -Tool $echoDir -Variant -1 -Target "C:\x"
$content = [System.IO.File]::ReadAllText($mark)
Assert-True ($content -eq "T=C:\x A= B=") "flat launch passes no variant args"

# Runtime edits to variant args take effect without reinstall and must keep route parity.
Remove-Item $mark -ErrorAction SilentlyContinue
'{"schema":1,"id":"echo-tool","name":"Echo","on":["folder"],"run":{"type":"powershell","entry":"main.ps1","keep_open":false},"variants":[{"label":"Bad","args":[""]}]}' |
    Set-Content -LiteralPath (Join-Path $echoDir "tool.json") -Encoding Ascii
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $launch -Tool $echoDir -Variant 0 -Target "C:\x"
Assert-True ($LASTEXITCODE -eq 1) "runtime manifest rejects empty variant args"
Assert-True (-not (Test-Path -LiteralPath $mark)) "tool does not run with an empty variant arg"

# error paths exit 1 without hanging (NOPAUSE)
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $launch -Tool (Join-Path $fx "missing") -Variant -1 -Target "C:\x"
Assert-True ($LASTEXITCODE -eq 1) "missing tool dir -> exit 1"
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $launch -Tool $echoDir -Variant 9 -Target "C:\x"
Assert-True ($LASTEXITCODE -eq 1) "variant index out of range -> exit 1"
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $launch -Tool $echoDir -Variant -2 -Target "C:\x"
Assert-True ($LASTEXITCODE -eq 1) "variant index below -1 -> exit 1"

# runtime manifest changes must not escape the tool folder
$outside = Join-Path $fx "outside.ps1"
[System.IO.File]::WriteAllText($outside, "exit 0`n")
'{"schema":1,"id":"echo-tool","name":"Echo","on":["folder"],"run":{"type":"powershell","entry":"..\\outside.ps1","keep_open":false}}' |
    Set-Content -LiteralPath (Join-Path $echoDir "tool.json") -Encoding Ascii
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $launch -Tool $echoDir -Variant -1 -Target "C:\x"
Assert-True ($LASTEXITCODE -eq 1) "entry outside tool folder -> exit 1"

} finally {
    Remove-Item -LiteralPath $fx -Recurse -Force -ErrorAction SilentlyContinue
}
Exit-Test
