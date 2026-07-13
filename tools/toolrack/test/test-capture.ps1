# test/test-capture.ps1 -- capture contract, rendering model, results, and latency.
. (Join-Path $PSScriptRoot "_assert.ps1")
. (Join-Path $PSScriptRoot "..\common\install.ps1")

$root = Split-Path $PSScriptRoot -Parent
$tool = Join-Path $root "tool\capture"
$manifestPath = Join-Path $tool "tool.json"
$mainPath = Join-Path $tool "main.ps1"
$nativePath = Join-Path $tool "capture.cs"
$dllPath = Join-Path $tool "bin\ToolrackCapture.dll"
$buildPath = Join-Path $tool "bin\build.json"

$requiredAssets = @($manifestPath, $mainPath, $nativePath, $dllPath, $buildPath)
foreach ($path in $requiredAssets) {
    Assert-True (Test-Path -LiteralPath $path -PathType Leaf) ("capture asset exists: " + (Split-Path $path -Leaf))
}
if (@($requiredAssets | Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) }).Count -gt 0) { Exit-Test }

$read = Read-Manifest $tool
Assert-True $read.Ok "manifest parses"
if ($read.Ok) {
    $errors = @(Test-Manifest $read.Data "capture" $tool)
    Assert-True ($errors.Count -eq 0) "manifest satisfies toolrack contract"
    Assert-True ($read.Data.run.window -eq "gui") "capture uses gui window mode"
    Assert-True ($null -eq $read.Data.PSObject.Properties["variants"]) "capture is one direct menu command"
}

$menu = ConvertFrom-Json ([IO.File]::ReadAllText((Join-Path $root "menu.json")))
$general = @($menu.categories | Where-Object { $_.id -eq "general" })[0]
Assert-True (@($general.tools) -contains "capture") "capture is in the general menu"

foreach ($path in @($mainPath, $nativePath)) {
    $bytes = [IO.File]::ReadAllBytes($path)
    Assert-True (@($bytes | Where-Object { $_ -gt 127 }).Count -eq 0) ((Split-Path $path -Leaf) + " is ASCII")
}

$build = ConvertFrom-Json ([IO.File]::ReadAllText($buildPath, (New-Object Text.UTF8Encoding($false, $true))))
$sourceHash = (Get-FileHash -LiteralPath $nativePath -Algorithm SHA256).Hash.ToLowerInvariant()
Assert-True ([string]$build.sourceSha256 -ceq $sourceHash) "build.json source SHA-256 matches capture.cs"
$assembly = [Reflection.Assembly]::LoadFrom($dllPath)
$buildType = $assembly.GetType("ToolrackCapture.BuildInfo", $true)
$dllBuildId = [string]$buildType.GetField("Id").GetValue($null)
Assert-True ($dllBuildId -ceq [string]$build.buildId) "DLL public build ID matches build.json"

$mainText = [IO.File]::ReadAllText($mainPath)
Assert-True ($mainText -notmatch 'Add-Type\s+-TypeDefinition') "main.ps1 never compiles capture.cs at runtime"
Assert-True ($mainText -match 'Assembly\]::LoadFrom') "main.ps1 loads the prebuilt DLL"
Assert-True ($mainText -match 'sourceSha256') "main.ps1 validates source hash before DLL load"
$sourceText = [IO.File]::ReadAllText($nativePath)
Assert-True ($sourceText -notmatch 'Segoe Fluent|Segoe MDL2|FontFamily') "palette does not depend on icon fonts"
Assert-True ($sourceText -match 'GraphicsPath') "palette icons are vector paths"

Add-Type -AssemblyName System.Drawing
$paletteType = $assembly.GetType("ToolrackCapture.CapturePalette", $true)
$workArea = New-Object Drawing.Rectangle -ArgumentList -1920, -200, 1920, 1080
$cursor = New-Object Drawing.Point -ArgumentList -5, -190
$audit = $paletteType.GetMethod("CreateAudit").Invoke($null, @(1.0, $false, $workArea, $cursor, $true))
Assert-True ($audit.ButtonCount -eq 6) "palette has exactly six actions"
Assert-True ((@($audit.Columns) -join ",") -ceq "range,window") "palette columns are range and window"
Assert-True ((@($audit.RangeActions) -join ",") -ceq "image,path,text") "range column has image, path, text"
Assert-True ((@($audit.WindowActions) -join ",") -ceq "image,path,text") "window column has image, path, text"
Assert-True ($audit.InnerCardBorders -eq 0 -and $audit.DividerCount -eq 1) "single surface has no inner cards and one center divider"
Assert-True ($audit.IconsAreVector -and -not $audit.UsesIconFont) "all action icons are vector geometry"
Assert-True ($audit.LightTextContrast -ge 4.5 -and $audit.DarkTextContrast -ge 4.5) "light and dark text contrast is at least 4.5 to 1"
Assert-True ($audit.DwmMode -eq "solid-fallback") "documented DWM failure uses solid fallback"
Assert-True ($audit.ClosesOnDeactivate) "outside deactivate closes the palette"
Assert-True ($audit.Keyboard.TabCycles -and $audit.Keyboard.ArrowsNavigate -and $audit.Keyboard.EnterActivates -and $audit.Keyboard.EscapeCloses) "keyboard focus state machine handles Tab, arrows, Enter, and Escape"

foreach ($scale in @(1.0, 1.5, 2.0)) {
    $dpiAudit = $paletteType.GetMethod("CreateAudit").Invoke($null, @($scale, $true, $workArea, $cursor, $false))
    $inside = $true
    $overlap = $false
    $rectangles = @($dpiAudit.HitRectangles)
    for ($i = 0; $i -lt $rectangles.Count; $i++) {
        $item = $rectangles[$i]
        if ($item.X -lt 0 -or $item.Y -lt 0 -or ($item.X + $item.Width) -gt $dpiAudit.PanelWidth -or
            ($item.Y + $item.Height) -gt $dpiAudit.PanelHeight) { $inside = $false }
        for ($j = $i + 1; $j -lt $rectangles.Count; $j++) {
            $other = $rectangles[$j]
            $intersects = $item.X -lt ($other.X + $other.Width) -and ($item.X + $item.Width) -gt $other.X -and
                $item.Y -lt ($other.Y + $other.Height) -and ($item.Y + $item.Height) -gt $other.Y
            if ($intersects) { $overlap = $true }
        }
    }
    Assert-True ($inside -and -not $overlap) ("hit rectangles are disjoint and inside at scale " + $scale)
    $panel = $dpiAudit.PanelBounds
    Assert-True ($panel.Left -ge $workArea.Left -and $panel.Top -ge $workArea.Top -and
        $panel.Right -le $workArea.Right -and $panel.Bottom -le $workArea.Bottom) ("negative-monitor edge placement stays in work area at scale " + $scale)
}

function Write-Utf8Json {
    param([string]$Path, $Value)
    [IO.File]::WriteAllText($Path, (ConvertTo-Json $Value -Depth 10), (New-Object Text.UTF8Encoding($false)))
}

function Send-HostCommand {
    param([string]$PipeName, [string]$Command)
    $client = New-Object IO.Pipes.NamedPipeClientStream(".", $PipeName, [IO.Pipes.PipeDirection]::InOut)
    try {
        $client.Connect(2000)
        $writer = New-Object IO.StreamWriter($client, (New-Object Text.UTF8Encoding($false)), 1024, $true)
        $reader = New-Object IO.StreamReader($client, (New-Object Text.UTF8Encoding($false, $true)), $false, 1024, $true)
        try {
            $writer.AutoFlush = $true
            $writer.WriteLine($Command)
            return ConvertFrom-Json $reader.ReadLine()
        } finally { $writer.Dispose(); $reader.Dispose() }
    } finally { $client.Dispose() }
}

function Wait-HostReady {
    param([string]$PipeName)
    $deadline = [DateTime]::UtcNow.AddSeconds(10)
    do {
        try {
            $status = Send-HostCommand $PipeName "status"
            if ([bool]$status.ready) { return $status }
        } catch { }
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $deadline)
    return $null
}

$fx = Join-Path $env:TEMP ("toolrack_capture_" + [guid]::NewGuid().ToString("N"))
$hostProcess = $null
$oldPaintDir = $env:TOOLRACK_CAPTURE_PAINT_DIR
$oldAutoClose = $env:TOOLRACK_CAPTURE_AUTOCLOSE
$oldOffscreen = $env:TOOLRACK_CAPTURE_TEST_OFFSCREEN
New-Item -ItemType Directory -Force -Path $fx | Out-Null
try {
    $lightPreview = Join-Path $fx "capture-light.png"
    $darkPreview = Join-Path $fx "capture-dark.png"
    $uiAudit = Join-Path $fx "ui.json"
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $mainPath -Smoke -AuditPath $uiAudit `
        -PreviewLightPath $lightPreview -PreviewDarkPath $darkPreview
    Assert-True ($LASTEXITCODE -eq 0) "prebuilt palette builds cleanly under PowerShell 5.1"
    Assert-True ((Test-Path -LiteralPath $uiAudit -PathType Leaf) -and
        (Test-Path -LiteralPath $lightPreview -PathType Leaf) -and
        (Test-Path -LiteralPath $darkPreview -PathType Leaf)) "smoke writes audit and light/dark previews"
    foreach ($preview in @($lightPreview, $darkPreview)) {
        $image = [Drawing.Image]::FromFile($preview)
        try { Assert-True ($image.Width -gt 300 -and $image.Height -gt 180) ((Split-Path $preview -Leaf) + " has full palette dimensions") }
        finally { $image.Dispose() }
    }

    $resultAudit = Join-Path $fx "result.json"
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $mainPath -SelfTest -AuditPath $resultAudit -TestOutputDir $fx
    Assert-True ($LASTEXITCODE -eq 0) "capture result self-test succeeds"
    Assert-True (Test-Path -LiteralPath $resultAudit -PathType Leaf) "result audit written"
    if (Test-Path -LiteralPath $resultAudit -PathType Leaf) {
        $result = ConvertFrom-Json ([IO.File]::ReadAllText($resultAudit))
        Assert-True (Test-Path -LiteralPath $result.SavedPath -PathType Leaf) "path action writes a PNG"
        Assert-True ($result.Width -eq 64 -and $result.Height -eq 40) "saved PNG preserves capture dimensions"
        Assert-True ($result.ClipboardText -eq $result.SavedPath) "path action copies the absolute path"
        Assert-True ($result.ImageClipboardWidth -eq 64 -and $result.ImageClipboardHeight -eq 40) "image action copies bitmap pixels"
        Assert-True ($result.OcrAvailable) "Windows local OCR engine is available"
        Assert-True ($result.OcrText -match "TEST\s*123") "Windows local OCR reads a generated text image"
    }

    $badTool = Join-Path $fx "bad capture"
    Copy-Item -LiteralPath $tool -Destination $badTool -Recurse
    [IO.File]::AppendAllText((Join-Path $badTool "capture.cs"), "`n// changed`n", (New-Object Text.UTF8Encoding($false)))
    $oldEap = $ErrorActionPreference
    try {
        $ErrorActionPreference = "Continue"
        $badOutput = @(& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $badTool "main.ps1") -Smoke 2>&1)
        $badCode = $LASTEXITCODE
    } finally { $ErrorActionPreference = $oldEap }
    Assert-True ($badCode -ne 0 -and (($badOutput -join " ") -match 'source hash')) "source SHA mismatch fails before loading DLL"

    $paintDir = Join-Path $fx "paint ticks"
    $state = Join-Path $fx "host state"
    New-Item -ItemType Directory -Force -Path $paintDir, $state | Out-Null
    $namespace = "capture-" + [guid]::NewGuid().ToString("N")
    Write-Utf8Json (Join-Path $state "host.json") ([ordered]@{ schema = 1; version = "1"; root = $root; namespace = $namespace })
    Write-Utf8Json (Join-Path $state "bindings.resolved.json") ([ordered]@{
        schema = 1; root = $root; sourceConfigPath = (Join-Path $root "bindings.json"); sourceConfigSha256 = ("0" * 64)
        active = @([ordered]@{
            id = "capture-latency"; trigger = [ordered]@{ type = "hotkey"; key = "F23"; modifiers = @("ctrl", "alt", "shift") }
            invoke = [ordered]@{ tool = "capture"; action = "default" }; toolDir = $tool
        }); rejected = @()
    })
    $env:TOOLRACK_CAPTURE_PAINT_DIR = $paintDir
    $env:TOOLRACK_CAPTURE_AUTOCLOSE = "1"
    $env:TOOLRACK_CAPTURE_TEST_OFFSCREEN = "1"
    $hostArguments = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden", "-File",
        (Join-Path $root "common\host.ps1"), "-StateRoot", $state, "-TestMode")
    $hostProcess = Start-Process -FilePath powershell.exe -ArgumentList (Join-ProcessArguments $hostArguments) -WindowStyle Hidden -PassThru
    $sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $pipeName = "ToolRackHost-" + $sid + "-" + $namespace
    $ready = Wait-HostReady $pipeName
    Assert-True ($null -ne $ready) "latency Host becomes ready"
    if ($null -ne $ready) {
        $durations = New-Object "System.Collections.Generic.List[double]"
        for ($index = 0; $index -lt 20; $index++) {
            $before = @(Get-ChildItem -LiteralPath $paintDir -Filter "*.tick" -File).Count
            $startTick = [Diagnostics.Stopwatch]::GetTimestamp()
            $queued = Send-HostCommand $pipeName "test-activate|capture|default"
            $deadline = [DateTime]::UtcNow.AddSeconds(5)
            $file = $null
            do {
                $files = @(Get-ChildItem -LiteralPath $paintDir -Filter "*.tick" -File | Sort-Object LastWriteTimeUtc)
                if ($files.Count -gt $before) { $file = $files[-1]; break }
                Start-Sleep -Milliseconds 10
            } while ([DateTime]::UtcNow -lt $deadline)
            if ($null -eq $file) { break }
            $paintTick = [long]([IO.File]::ReadAllText($file.FullName).Trim())
            [void]$durations.Add((($paintTick - $startTick) * 1000.0) / [Diagnostics.Stopwatch]::Frequency)
        }
        Assert-True ($durations.Count -eq 20) "twenty Host palette paints complete"
        if ($durations.Count -eq 20) {
            $sorted = @($durations.ToArray() | Sort-Object)
            $p95 = $sorted[[Math]::Ceiling($sorted.Count * 0.95) - 1]
            Assert-True ($durations[0] -le 800.0) ("first Host palette paint is at most 800 ms (" + [Math]::Round($durations[0], 1) + " ms)")
            Assert-True ($p95 -le 500.0) ("same-session Host palette p95 is at most 500 ms (" + [Math]::Round($p95, 1) + " ms)")
        }
        [void](Send-HostCommand $pipeName "shutdown")
        Assert-True ($hostProcess.WaitForExit(5000)) "latency Host shuts down"
        $hostProcess = $null
    }
} finally {
    if ($null -ne $hostProcess -and -not $hostProcess.HasExited) { try { $hostProcess.Kill() } catch { } }
    $env:TOOLRACK_CAPTURE_PAINT_DIR = $oldPaintDir
    $env:TOOLRACK_CAPTURE_AUTOCLOSE = $oldAutoClose
    $env:TOOLRACK_CAPTURE_TEST_OFFSCREEN = $oldOffscreen
    Remove-Item -LiteralPath $fx -Recurse -Force -ErrorAction SilentlyContinue
}

Exit-Test
