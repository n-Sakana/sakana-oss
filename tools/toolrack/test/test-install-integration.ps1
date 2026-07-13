# test/test-install-integration.ps1 -- real HKCU roundtrip with a temporary demo tool
. (Join-Path $PSScriptRoot "_assert.ps1")
$root = Split-Path $PSScriptRoot -Parent
$menuPath = Join-Path $root "menu.json"
$demo = Join-Path $root "tool\zz-demo"
New-Item -ItemType Directory -Force $demo | Out-Null
[System.IO.File]::WriteAllText((Join-Path $demo "main.ps1"), "param([string]`$Target)`nexit 0`n")
[System.IO.File]::WriteAllText((Join-Path $demo "tool.json"),
    '{"schema":1,"id":"zz-demo","name":"ZZ Demo","on":["folder"],"run":{"type":"powershell","entry":"main.ps1"},"variants":[{"label":"A","args":["-X","1"]},{"label":"B","args":[]}]}')
# A control character in a menu label must be isolated before .reg generation.
$badDir = Join-Path $root "tool\zz-bad"
New-Item -ItemType Directory -Force $badDir | Out-Null
[System.IO.File]::WriteAllText((Join-Path $badDir "main.ps1"), "exit 0`n")
[System.IO.File]::WriteAllText((Join-Path $badDir "tool.json"),
    '{"schema":1,"id":"zz-bad","name":"bad\nname","on":["folder"],"run":{"type":"powershell","entry":"main.ps1"}}')

try {
    $out = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root "common\install.ps1") 2>&1
    $outText = @($out | ForEach-Object { [string]$_ })
    Assert-Contains $outText "*OK*zz-demo*" "install reports OK for valid tool"
    Assert-Contains $outText "*SKIP*zz-bad*" "install rejects control characters without blocking valid tools"
    Assert-Contains $outText "*zz-demo*default category*" "unlisted valid tool uses default category"
    $demoKey = "HKCU\Software\Classes\Directory\shell\ToolRack.general.p2\shell\zz-demo"
    & reg query ($demoKey + "\shell\v0\command") | Out-Null
    Assert-True ($LASTEXITCODE -eq 0) "variant v0 registered in HKCU"
    & reg query "HKCU\Software\Classes\Directory\shell\ToolRack.general.p2\shell\zz-bad" 2>$null | Out-Null
    Assert-True ($LASTEXITCODE -ne 0) "invalid tool not registered"

    $hadMenu = Test-Path -LiteralPath $menuPath -PathType Leaf
    if ($hadMenu) { $menuBytes = [System.IO.File]::ReadAllBytes($menuPath) }
    [System.IO.File]::WriteAllText($menuPath,
        '{"schema":99,"default_category":"general","categories":[]}')
    $badOut = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root "common\install.ps1") 2>&1
    $badCode = $LASTEXITCODE
    Assert-True ($badCode -ne 0) "invalid menu.json makes install fail"
    Assert-Contains @($badOut | ForEach-Object { [string]$_ }) "*registry was not changed*" "failure states preservation guarantee"
    & reg query $demoKey | Out-Null
    Assert-True ($LASTEXITCODE -eq 0) "invalid menu.json preserves previous registry"
} finally {
    if ($hadMenu) { [System.IO.File]::WriteAllBytes($menuPath, $menuBytes) }
    else { Remove-Item -LiteralPath $menuPath -Force -ErrorAction SilentlyContinue }
    Remove-Item -Recurse -Force $demo, $badDir
    # re-run install without the demo -> its entries must vanish (idempotent wipe & rebuild)
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root "common\install.ps1") | Out-Null
}
& reg query "HKCU\Software\Classes\Directory\shell\ToolRack.general.p2\shell\zz-demo" 2>$null | Out-Null
Assert-True ($LASTEXITCODE -ne 0) "re-install without zz-demo removes its menu entry"
Exit-Test
