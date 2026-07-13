# test/test-host-hotkeys.ps1 -- deterministic and real API hotkey tests.
. (Join-Path $PSScriptRoot "_assert.ps1")

$root = Split-Path $PSScriptRoot -Parent
$hostScript = Join-Path $root "common\host.ps1"
$controlScript = Join-Path $root "common\host-control.ps1"
$sourceDir = Join-Path $root "common\host"
$managerSource = Join-Path $sourceDir "HotkeyManager.cs"
$harnessSource = Join-Path $PSScriptRoot "fixture\host-hotkeys\HotkeyHarness.cs"
Assert-True (Test-Path -LiteralPath $managerSource -PathType Leaf) "HotkeyManager source exists"
Assert-True (Test-Path -LiteralPath $harnessSource -PathType Leaf) "hotkey harness exists"
if (-not (Test-Path -LiteralPath $managerSource -PathType Leaf) -or -not (Test-Path -LiteralPath $harnessSource -PathType Leaf)) { Exit-Test }
foreach ($path in @($managerSource, $harnessSource)) {
    Assert-True (@([IO.File]::ReadAllBytes($path) | Where-Object { $_ -gt 127 }).Count -eq 0) ("hotkey source is ASCII: " + (Split-Path $path -Leaf))
}

$sources = @(
    (Join-Path $sourceDir "BindingModel.cs"),
    (Join-Path $sourceDir "NativeMethods.cs"),
    (Join-Path $sourceDir "HostPipe.cs"),
    $managerSource,
    (Join-Path $sourceDir "MouseGestureHook.cs"),
    (Join-Path $sourceDir "ActivationWorker.cs"),
    (Join-Path $sourceDir "HostApplication.cs"),
    $harnessSource
)
try {
    Add-Type -Path $sources -ReferencedAssemblies @(
        "System.dll", "System.Core.dll", "System.Windows.Forms.dll", "System.Drawing.dll",
        "System.Runtime.Serialization.dll", "System.Xml.dll",
        [Management.Automation.PSObject].Assembly.Location
    ) -ErrorAction Stop
} catch {
    Assert-True $false ("hotkey harness compiles: " + $_.Exception.Message)
    Exit-Test
}
$harness = [ToolRackTests.HotkeyHarness]::Run()
Assert-True $harness.InitialRegistered "global bindings register"
Assert-True $harness.DuplicateRejected "duplicate normalized trigger is isolated"
Assert-True $harness.FailureIsolated "RegisterHotKey failure leaves other binding active"
Assert-True $harness.ReloadUnregistered "reload unregisters old hotkeys"
Assert-True $harness.NoRepeat "MOD_NOREPEAT is always set"
Assert-True $harness.DispatchQueued "WM_HOTKEY dispatch reaches activation sink once"
Assert-True $harness.DisposeUnregistered "dispose unregisters every hotkey"
Assert-True $harness.StableIds "registration IDs are stable across reorder"

$combinedSource = @($sources | ForEach-Object { [IO.File]::ReadAllText($_) }) -join "`n"
Assert-True ($combinedSource -notmatch 'SetWinEventHook|GetForegroundWindow|WH_KEYBOARD_LL') "hotkey implementation has no foreground tracker or keyboard hook"

function Write-Utf8Json {
    param([string]$Path, $Value)
    [IO.File]::WriteAllText($Path, (ConvertTo-Json $Value -Depth 10), (New-Object Text.UTF8Encoding($false)))
}

function Get-Status {
    param([string]$StateRoot)
    $output = @(& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $controlScript -Status -StateRoot $StateRoot 2>&1)
    if ($LASTEXITCODE -ne 0) { return $null }
    $line = @($output | Where-Object { ([string]$_).StartsWith("{") } | Select-Object -Last 1)
    if ($line.Count -ne 1) { return $null }
    try { return ConvertFrom-Json ([string]$line[0]) } catch { return $null }
}

$fx = Join-Path $env:TEMP ("toolrack_hotkey_smoke_" + [guid]::NewGuid().ToString("N"))
$hostProcess = $null
$registered = $false
try {
    New-Item -ItemType Directory -Force -Path $fx | Out-Null
    foreach ($number in (24..13)) {
        if ($number -eq 12) { continue }
        $state = Join-Path $fx ("state-" + $number)
        New-Item -ItemType Directory -Force -Path $state | Out-Null
        $nameSpace = "hotkey-" + [guid]::NewGuid().ToString("N")
        Write-Utf8Json (Join-Path $state "host.json") ([ordered]@{ schema = 1; version = "1"; root = $root; namespace = $nameSpace })
        Write-Utf8Json (Join-Path $state "bindings.resolved.json") ([ordered]@{
            schema = 1
            root = $root
            sourceConfigPath = (Join-Path $root "bindings.json")
            sourceConfigSha256 = ("0" * 64)
            active = @([ordered]@{
                id = "smoke-hotkey"
                trigger = [ordered]@{ type = "hotkey"; key = ("F" + $number); modifiers = @("ctrl", "alt", "shift") }
                invoke = [ordered]@{ tool = "capture"; action = "default" }
                toolDir = (Join-Path $root "tool\capture")
            })
            rejected = @()
        })
        $hostProcess = Start-Process -FilePath powershell.exe -ArgumentList @(
            "-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden", "-File", $hostScript, "-StateRoot", $state
        ) -WindowStyle Hidden -PassThru
        $deadline = [DateTime]::UtcNow.AddSeconds(8)
        $status = $null
        do {
            $status = Get-Status $state
            if ($null -ne $status -and [bool]$status.ready) { break }
            Start-Sleep -Milliseconds 100
        } while ([DateTime]::UtcNow -lt $deadline)
        if ($null -ne $status -and [int]$status.activeHotkeys -eq 1) {
            $registered = $true
        }
        if ($null -ne $status) {
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $controlScript -Shutdown -StateRoot $state | Out-Null
        }
        if (-not $hostProcess.WaitForExit(5000)) { $hostProcess.Kill() }
        $hostProcess = $null
        if ($registered) { break }
    }
    Assert-True $registered "real RegisterHotKey smoke registers an F13-F24 chord"
} finally {
    if ($null -ne $hostProcess -and -not $hostProcess.HasExited) {
        try { $hostProcess.Kill() } catch {}
    }
    Remove-Item -LiteralPath $fx -Recurse -Force -ErrorAction SilentlyContinue
}

Exit-Test
