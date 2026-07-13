# test/test-host-prereq.ps1 -- preflight for the PowerShell global host probe.
. (Join-Path $PSScriptRoot "_assert.ps1")

$root = Split-Path $PSScriptRoot -Parent
$probe = Join-Path $PSScriptRoot "probe-global-host.ps1"
$fixture = Join-Path $PSScriptRoot "fixture\host-latency"
$required = @(
    $probe,
    (Join-Path $fixture "HostProbe.cs"),
    (Join-Path $fixture "LatencyPalette.cs"),
    (Join-Path $fixture "main.ps1"),
    (Join-Path $fixture "tool.json"),
    (Join-Path $fixture "start-probe.vbs")
)

foreach ($path in $required) {
    Assert-True (Test-Path -LiteralPath $path -PathType Leaf) ("prerequisite asset exists: " + (Split-Path $path -Leaf))
}

if (@($required | Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) }).Count -gt 0) {
    Exit-Test
}

foreach ($path in $required) {
    $bytes = [IO.File]::ReadAllBytes($path)
    Assert-True (@($bytes | Where-Object { $_ -gt 127 }).Count -eq 0) ("asset is ASCII: " + (Split-Path $path -Leaf))
}

$major = & powershell.exe -NoProfile -ExecutionPolicy Bypass -Command '$PSVersionTable.PSVersion.Major'
Assert-True ($LASTEXITCODE -eq 0 -and [int]$major -eq 5) "child process is Windows PowerShell 5"

$state = Join-Path $env:TEMP ("toolrack_host_prereq_" + [guid]::NewGuid().ToString("N"))
try {
    New-Item -ItemType Directory -Force -Path $state | Out-Null
    $output = @(& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $probe -SelfTest -StateRoot $state 2>&1)
    $exitCode = $LASTEXITCODE
    Assert-True ($exitCode -eq 0) "probe self-test exits zero"

    $jsonLine = @($output | Where-Object { ([string]$_).Trim().StartsWith("{") } | Select-Object -Last 1)
    Assert-True ($jsonLine.Count -eq 1) "probe self-test returns JSON"
    if ($jsonLine.Count -eq 1) {
        $result = $null
        try { $result = ConvertFrom-Json ([string]$jsonLine[0]) }
        catch { $result = $null }
        Assert-True ($null -ne $result) "probe JSON parses"
        if ($null -ne $result) {
            Assert-True ([bool]$result.AddType) "Add-Type compiles and loads C# 5"
            Assert-True ([bool]$result.HookInstalled) "WH_MOUSE_LL installs in the probe"
            Assert-True ([bool]$result.HotkeyAvailable) "RegisterHotKey is callable"
            Assert-True ([bool]$result.MessageLoop) "message loop executes"
            Assert-True ([bool]$result.Cleanup) "hook and callback clean up"
            Assert-True ($null -ne $result.DwmAvailable) "DWM availability is reported as data"
            Assert-True ([bool]$result.RunspaceAvailable) "persistent worker runspace executes requests"
        }
    }
} finally {
    Remove-Item -LiteralPath $state -Recurse -Force -ErrorAction SilentlyContinue
}

Assert-True (-not (Test-Path -LiteralPath $state)) "prerequisite fixture is removed"

$singleState = Join-Path $env:TEMP ("toolrack_host_single_" + [guid]::NewGuid().ToString("N"))
$first = $null
$second = $null
try {
    New-Item -ItemType Directory -Force -Path $singleState | Out-Null
    $arguments = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden", "-File", $probe, "-Run", "-StateRoot", $singleState)
    $first = Start-Process -FilePath powershell.exe -ArgumentList $arguments -WindowStyle Hidden -PassThru
    $heartbeat = Join-Path $singleState "heartbeat.json"
    $deadline = [DateTime]::UtcNow.AddSeconds(5)
    while (-not (Test-Path -LiteralPath $heartbeat -PathType Leaf) -and [DateTime]::UtcNow -lt $deadline) {
        Start-Sleep -Milliseconds 50
    }
    Assert-True (Test-Path -LiteralPath $heartbeat -PathType Leaf) "first probe instance becomes ready"

    $second = Start-Process -FilePath powershell.exe -ArgumentList $arguments -WindowStyle Hidden -PassThru
    $secondExited = $second.WaitForExit(3000)
    Assert-True $secondExited "duplicate probe instance exits promptly"
    if ($secondExited) { Assert-True ($second.ExitCode -eq 3) "duplicate probe reports already running" }

    [IO.File]::WriteAllText((Join-Path $singleState "stop.request"), "stop")
    $firstExited = $first.WaitForExit(5000)
    Assert-True $firstExited "primary probe stops cleanly"
} finally {
    foreach ($process in @($first, $second)) {
        if ($null -ne $process -and -not $process.HasExited) {
            try { $process.Kill() } catch {}
        }
    }
    Remove-Item -LiteralPath $singleState -Recurse -Force -ErrorAction SilentlyContinue
}

Exit-Test
