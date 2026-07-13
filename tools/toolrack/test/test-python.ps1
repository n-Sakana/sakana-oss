# test/test-python.ps1 -- python type: requires check + argv passthrough (skips if py absent)
. (Join-Path $PSScriptRoot "_assert.ps1")
$root = Split-Path $PSScriptRoot -Parent
$launch = Join-Path $root "common\launch.ps1"
$env:TOOLRACK_NOPAUSE = "1"
if ($null -eq (Get-Command py -ErrorAction SilentlyContinue)) {
    Write-Host "  skip: py launcher not installed" -ForegroundColor Yellow
    Exit-Test
}
$fx = Join-Path $env:TEMP ("toolrack_pt_" + [guid]::NewGuid().ToString("N"))
$pdir = Join-Path $fx "py-tool"
New-Item -ItemType Directory -Force $pdir | Out-Null
try {
$mark = Join-Path $fx "py-mark.txt"
@'
import os, sys
with open(os.environ["TOOLRACK_TEST_MARK"], "w") as f:
    f.write("|".join(sys.argv[1:]))
sys.exit(3)
'@ | Set-Content -LiteralPath (Join-Path $pdir "main.py") -Encoding Ascii
'{"schema":1,"id":"py-tool","name":"PyTool","on":["folder"],"run":{"type":"python","entry":"main.py","keep_open":false,"requires":["json"]},"variants":[{"label":"V","args":["--depth","5"]}]}' |
    Set-Content -LiteralPath (Join-Path $pdir "tool.json") -Encoding Ascii
$env:TOOLRACK_TEST_MARK = $mark

# happy path: requires satisfied (json is stdlib), argv passed through
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $launch -Tool $pdir -Variant 0 -Target "C:\p y"
Assert-True ($LASTEXITCODE -eq 3) "python exit code passthrough"
$argvLine = [System.IO.File]::ReadAllText($mark)
Assert-True ($argvLine -eq "--target|C:\p y|--depth|5") ("python argv intact: got '" + $argvLine + "'")

# py may write a harmless warning to stderr while still exiting zero
Remove-Item $mark
$realPy = (Get-Command py -ErrorAction Stop).Source
$fakeBin = Join-Path $fx "fake-bin"
New-Item -ItemType Directory -Force $fakeBin | Out-Null
$fakePy = Join-Path $fakeBin "py.cmd"
$fakeBody = "@echo off`r`nif `"%~2`"==`"-c`" echo harmless warning 1>&2`r`n@`"$realPy`" %*`r`n"
[System.IO.File]::WriteAllText($fakePy, $fakeBody)
$oldPath = $env:PATH
try {
    $env:PATH = $fakeBin + ";" + $oldPath
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $launch -Tool $pdir -Variant 0 -Target "C:\warn"
    Assert-True ($LASTEXITCODE -eq 3) "zero-exit requires probe tolerates stderr"
    Assert-True (Test-Path -LiteralPath $mark) "python tool runs after harmless probe stderr"

    # With no requires, the same condition exercises the pythonw path probe.
    Remove-Item $mark -ErrorAction SilentlyContinue
    '{"schema":1,"id":"py-tool","name":"PyTool","on":["folder"],"run":{"type":"python","entry":"main.py","window":"hidden"}}' |
        Set-Content -LiteralPath (Join-Path $pdir "tool.json") -Encoding Ascii
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $launch -Gui -Tool $pdir -Variant -1 -Target "C:\hidden-warn"
    Assert-True ($LASTEXITCODE -eq 0) "zero-exit pythonw probe tolerates stderr"
    for ($i = 0; $i -lt 50 -and -not (Test-Path -LiteralPath $mark); $i++) { Start-Sleep -Milliseconds 100 }
    Assert-True (Test-Path -LiteralPath $mark) "hidden python tool runs after harmless probe stderr"
} finally {
    $env:PATH = $oldPath
}

# missing module: tool must NOT run
Remove-Item $mark -ErrorAction SilentlyContinue
'{"schema":1,"id":"py-tool","name":"PyTool","on":["folder"],"run":{"type":"python","entry":"main.py","keep_open":false,"requires":["not_a_module_zz9"]}}' |
    Set-Content -LiteralPath (Join-Path $pdir "tool.json") -Encoding Ascii
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $launch -Tool $pdir -Variant -1 -Target "C:\x"
Assert-True ($LASTEXITCODE -eq 1) "missing module -> exit 1"
Assert-True (-not (Test-Path -LiteralPath $mark)) "tool did not run when module missing"

} finally {
    Remove-Item -LiteralPath $fx -Recurse -Force -ErrorAction SilentlyContinue
}
Exit-Test
