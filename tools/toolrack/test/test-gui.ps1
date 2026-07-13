# test/test-gui.ps1 -- gui path: detached, hidden, marker file appears
. (Join-Path $PSScriptRoot "_assert.ps1")
$root = Split-Path $PSScriptRoot -Parent
$launch = Join-Path $root "common\launch.ps1"
$vbs = Join-Path $root "common\silent.vbs"
$env:TOOLRACK_NOPAUSE = "1"
$fx = Join-Path $env:TEMP ("toolrack_gt_" + [guid]::NewGuid().ToString("N"))
$gdir = Join-Path $fx "gui-tool"
New-Item -ItemType Directory -Force $gdir | Out-Null
try {
$mark = Join-Path $fx "gui-mark.txt"
@'
param([string]$Target)
[System.IO.File]::WriteAllText($env:TOOLRACK_TEST_MARK, "gui-ran T=$Target")
'@ | Set-Content -LiteralPath (Join-Path $gdir "main.ps1") -Encoding Ascii
'{"schema":1,"id":"gui-tool","name":"GuiTool","on":["background"],"run":{"type":"powershell","entry":"main.ps1","window":"gui"}}' |
    Set-Content -LiteralPath (Join-Path $gdir "tool.json") -Encoding Ascii
$env:TOOLRACK_TEST_MARK = $mark

function Wait-Marker {
    param([int]$Sec = 10)
    for ($i = 0; $i -lt ($Sec * 4); $i++) {
        if (Test-Path -LiteralPath $mark) { return $true }
        Start-Sleep -Milliseconds 250
    }
    return $false
}

# route 1: launch.ps1 directly (console host, gui window mode -> detached)
$sw = [System.Diagnostics.Stopwatch]::StartNew()
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $launch -Tool $gdir -Variant -1 -Target "C:\g t"
$sw.Stop()
Assert-True ($LASTEXITCODE -eq 0) "gui launch returns 0 immediately"
Assert-True (Wait-Marker) "gui tool ran (marker written)"
Assert-True ([System.IO.File]::ReadAllText($mark) -eq "gui-ran T=C:\g t") "target passed intact"

# route 2: through silent.vbs (the real registry path)
Remove-Item $mark
if (Test-Path -LiteralPath $vbs) {
    & wscript.exe $vbs $gdir -1 "C:\g t"
    Assert-True (Wait-Marker) "silent.vbs route also runs the tool"
} else {
    Assert-True $false "silent.vbs route also runs the tool"
}

# route 3: hidden mode is detached without claiming the tool has a GUI
Remove-Item $mark -ErrorAction SilentlyContinue
'{"schema":1,"id":"gui-tool","name":"HiddenTool","on":["background"],"run":{"type":"powershell","entry":"main.ps1","window":"hidden"}}' |
    Set-Content -LiteralPath (Join-Path $gdir "tool.json") -Encoding Ascii
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $launch -Tool $gdir -Variant -1 -Target "C:\hidden"
$hiddenCode = $LASTEXITCODE
Assert-True ($hiddenCode -eq 0) "hidden launch returns 0"
if ($hiddenCode -eq 0) {
    Assert-True (Wait-Marker) "hidden tool ran without a console route"
} else {
    Assert-True $false "hidden tool ran without a console route"
}

# drive-root target keeps its trailing backslash through the real VBS route
Remove-Item $mark -ErrorAction SilentlyContinue
& wscript.exe $vbs $gdir -1 "C:\"
Assert-True (Wait-Marker) "drive-root target runs through silent.vbs"
if (Test-Path -LiteralPath $mark) {
    $driveRootText = [System.IO.File]::ReadAllText($mark)
    Assert-True ($driveRootText -eq "gui-ran T=C:\") ("drive-root target remains C:\ (got '" + $driveRootText + "')")
}

# stale hidden registry route + runtime console manifest must not wait on an invisible ReadKey
'{"schema":1,"id":"gui-tool","name":"ConsoleNow","on":["background"],"run":{"type":"powershell","entry":"main.ps1"}}' |
    Set-Content -LiteralPath (Join-Path $gdir "tool.json") -Encoding Ascii
$env:TOOLRACK_NOPAUSE = ""
$mismatchArgs = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -Gui -Tool "{1}" -Variant -1 -Target "C:\mismatch"' -f $launch, $gdir
$mismatch = Start-Process -FilePath "powershell.exe" -ArgumentList $mismatchArgs -WindowStyle Hidden -PassThru
$mismatchFinished = $mismatch.WaitForExit(3000)
if (-not $mismatchFinished) { Stop-Process -Id $mismatch.Id -Force -ErrorAction SilentlyContinue }
Assert-True $mismatchFinished "stale hidden route does not leave an invisible keep-open process"
if ($mismatchFinished) { Assert-True ($mismatch.ExitCode -eq 0) "stale hidden route exits cleanly" }
$env:TOOLRACK_NOPAUSE = "1"

} finally {
    Remove-Item -LiteralPath $fx -Recurse -Force -ErrorAction SilentlyContinue
}
Exit-Test
