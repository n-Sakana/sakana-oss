# test/test-uninstall.ps1 -- register a demo tool, then uninstall must wipe all contexts
. (Join-Path $PSScriptRoot "_assert.ps1")
$root = Split-Path $PSScriptRoot -Parent
$demo = Join-Path $root "tool\zz-demo"
New-Item -ItemType Directory -Force $demo | Out-Null
[System.IO.File]::WriteAllText((Join-Path $demo "main.ps1"), "exit 0`n")
[System.IO.File]::WriteAllText((Join-Path $demo "tool.json"),
    '{"schema":1,"id":"zz-demo","name":"ZZ","on":["file","folder","background"],"run":{"type":"powershell","entry":"main.ps1"}}')
try {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root "common\install.ps1") | Out-Null
    & reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v ToolRackHost | Out-Null
    Assert-True ($LASTEXITCODE -eq 0) "Host Run value exists before uninstall"
    Assert-True (Test-Path -LiteralPath (Join-Path $env:LOCALAPPDATA "ToolRack\active.txt") -PathType Leaf) "Host local state exists before uninstall"
    foreach ($base in @("HKCU\Software\Classes\*\shell",
                        "HKCU\Software\Classes\Directory\shell",
                        "HKCU\Software\Classes\Directory\Background\shell")) {
        & reg add ($base + "\ToolRack.legacy.p9") /f | Out-Null
    }
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root "common\uninstall.ps1") | Out-Null
    foreach ($k in @("HKCU\Software\Classes\*\shell\ToolRack",
                     "HKCU\Software\Classes\Directory\shell\ToolRack",
                     "HKCU\Software\Classes\Directory\Background\shell\ToolRack",
                     "HKCU\Software\Classes\*\shell\ToolRack.ai",
                     "HKCU\Software\Classes\Directory\shell\ToolRack.general.p2",
                     "HKCU\Software\Classes\Directory\Background\shell\ToolRack.legacy.p9")) {
        & reg query $k 2>$null | Out-Null
        Assert-True ($LASTEXITCODE -ne 0) ("wiped: " + $k)
    }
    & reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v ToolRackHost 2>$null | Out-Null
    Assert-True ($LASTEXITCODE -ne 0) "uninstall removes only the ToolRackHost Run value"
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $env:LOCALAPPDATA "ToolRack"))) "uninstall removes local Host state"
} finally { Remove-Item -Recurse -Force $demo }
Exit-Test
