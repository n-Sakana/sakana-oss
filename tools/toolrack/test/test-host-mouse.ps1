# test/test-host-mouse.ps1 -- Ctrl+RightClick state machine and hook lifecycle.
. (Join-Path $PSScriptRoot "_assert.ps1")

$root = Split-Path $PSScriptRoot -Parent
$sourceDir = Join-Path $root "common\host"
$mouseSource = Join-Path $sourceDir "MouseGestureHook.cs"
$harnessSource = Join-Path $PSScriptRoot "fixture\host-mouse\MouseHarness.cs"
$hostScript = Join-Path $root "common\host.ps1"
$controlScript = Join-Path $root "common\host-control.ps1"
Assert-True (Test-Path -LiteralPath $mouseSource -PathType Leaf) "MouseGestureHook source exists"
Assert-True (Test-Path -LiteralPath $harnessSource -PathType Leaf) "mouse harness exists"
if (-not (Test-Path -LiteralPath $mouseSource -PathType Leaf) -or -not (Test-Path -LiteralPath $harnessSource -PathType Leaf)) { Exit-Test }

$sources = @(
    (Join-Path $sourceDir "BindingModel.cs"),
    (Join-Path $sourceDir "NativeMethods.cs"),
    (Join-Path $sourceDir "HostPipe.cs"),
    (Join-Path $sourceDir "HotkeyManager.cs"),
    $mouseSource,
    (Join-Path $sourceDir "ActivationWorker.cs"),
    (Join-Path $sourceDir "HostApplication.cs"),
    (Join-Path $PSScriptRoot "fixture\host-hotkeys\HotkeyHarness.cs"),
    $harnessSource
)
Add-Type -Path $sources -ReferencedAssemblies @(
    "System.dll", "System.Core.dll", "System.Windows.Forms.dll", "System.Drawing.dll",
    "System.Runtime.Serialization.dll", "System.Xml.dll",
    [Management.Automation.PSObject].Assembly.Location
)
$result = [ToolRackTests.MouseHarness]::Run()
Assert-True $result.PlainPass "plain right down/up pass"
Assert-True $result.ExactActivatesOnce "Ctrl+RightClick swallows and activates once"
Assert-True $result.ReleasedModifierStillSwallows "right up is swallowed after Ctrl is released"
Assert-True $result.ExtraModifierPasses "extra modifier does not match exact chord"
Assert-True $result.UpOnlyPasses "right up without armed down passes"
Assert-True $result.DoubleClickNoSecondActivation "double click does not activate twice"
Assert-True $result.InjectedPasses "injected input passes"
Assert-True $result.NegativeCodePasses "negative hook code passes"
Assert-True $result.DisabledPasses "disabled state passes all input"
Assert-True $result.ReloadDisposesOldHook "reload disposes old hook"
Assert-True $result.ShutdownDisposesHook "shutdown disposes hook"
Assert-True ([double]$result.P99Milliseconds -lt 5.0) ("state decision p99 below 5 ms (" + $result.P99Milliseconds + " ms)")

$mouseText = [IO.File]::ReadAllText($mouseSource)
Assert-True ($mouseText -match 'CallNextHookEx') "hook chain forwarding is implemented"
$callbackMatch = [regex]::Match($mouseText, '(?s)HookCallback\(.*?\n\s*\}(?=\n\s*private)')
Assert-True $callbackMatch.Success "hook callback body is found"
if ($callbackMatch.Success) {
    Assert-True ($callbackMatch.Value -notmatch 'File\.|Process\.Start|Json|log\(') "hook callback performs no I/O, launch, JSON, or logging"
}

function Write-Utf8Json {
    param([string]$Path, $Value)
    [IO.File]::WriteAllText($Path, (ConvertTo-Json $Value -Depth 10), (New-Object Text.UTF8Encoding($false)))
}

$fx = Join-Path $env:TEMP ("toolrack_mouse_smoke_" + [guid]::NewGuid().ToString("N"))
$hostProcess = $null
try {
    New-Item -ItemType Directory -Force -Path $fx | Out-Null
    $state = Join-Path $fx "state"
    New-Item -ItemType Directory -Force -Path $state | Out-Null
    Write-Utf8Json (Join-Path $state "host.json") ([ordered]@{ schema = 1; version = "1"; root = $root; namespace = ("mouse-" + [guid]::NewGuid().ToString("N")) })
    Write-Utf8Json (Join-Path $state "bindings.resolved.json") ([ordered]@{
        schema = 1
        root = $root
        sourceConfigPath = (Join-Path $root "bindings.json")
        sourceConfigSha256 = ("0" * 64)
        active = @([ordered]@{
            id = "capture-mouse"
            trigger = [ordered]@{ type = "mouse"; button = "right"; modifiers = @("ctrl") }
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
        $output = @(& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $controlScript -Status -StateRoot $state 2>&1)
        if ($LASTEXITCODE -eq 0) {
            $line = @($output | Where-Object { ([string]$_).StartsWith("{") } | Select-Object -Last 1)
            if ($line.Count -eq 1) { try { $status = ConvertFrom-Json ([string]$line[0]) } catch {} }
        }
        if ($null -ne $status -and [bool]$status.ready) { break }
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $deadline)
    Assert-True ($null -ne $status -and [bool]$status.mouseHookActive -and [int]$status.activeMouseBindings -eq 1) "real WH_MOUSE_LL hook installs without synthetic input"
    if ($null -ne $status) {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $controlScript -Shutdown -StateRoot $state | Out-Null
    }
    Assert-True ($hostProcess.WaitForExit(5000)) "Host shutdown unhooks real mouse hook"
    $hostProcess = $null
} finally {
    if ($null -ne $hostProcess -and -not $hostProcess.HasExited) {
        try { $hostProcess.Kill() } catch {}
    }
    Remove-Item -LiteralPath $fx -Recurse -Force -ErrorAction SilentlyContinue
}

Exit-Test
