# test/test-host-runtime.ps1 -- PowerShell Host, bootstrap, and control pipe.
. (Join-Path $PSScriptRoot "_assert.ps1")

$root = Split-Path $PSScriptRoot -Parent
$hostScript = Join-Path $root "common\host.ps1"
$bootstrapScript = Join-Path $root "common\host-bootstrap.ps1"
$startVbs = Join-Path $root "common\host-start.vbs"
$controlScript = Join-Path $root "common\host-control.ps1"
$sourceDir = Join-Path $root "common\host"
$required = @(
    $hostScript,
    $bootstrapScript,
    $startVbs,
    $controlScript,
    (Join-Path $sourceDir "BindingModel.cs"),
    (Join-Path $sourceDir "NativeMethods.cs"),
    (Join-Path $sourceDir "HostPipe.cs"),
    (Join-Path $sourceDir "ActivationWorker.cs"),
    (Join-Path $sourceDir "HostApplication.cs")
)
foreach ($path in $required) {
    Assert-True (Test-Path -LiteralPath $path -PathType Leaf) ("Host asset exists: " + (Split-Path $path -Leaf))
}
if (@($required | Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) }).Count -gt 0) { Exit-Test }

foreach ($path in $required) {
    $bytes = [IO.File]::ReadAllBytes($path)
    Assert-True (@($bytes | Where-Object { $_ -gt 127 }).Count -eq 0) ("Host asset is ASCII: " + (Split-Path $path -Leaf))
}

function Write-Utf8Json {
    param([string]$Path, $Value)
    [IO.File]::WriteAllText($Path, (ConvertTo-Json $Value -Depth 8), (New-Object Text.UTF8Encoding($false)))
}

function New-HostState {
    param([string]$Path, [string]$RepositoryRoot, [string]$Namespace, [switch]$BadConfig)
    New-Item -ItemType Directory -Force -Path $Path | Out-Null
    Write-Utf8Json (Join-Path $Path "host.json") ([ordered]@{
        schema = 1
        version = "1"
        root = [IO.Path]::GetFullPath($RepositoryRoot)
        namespace = $Namespace
    })
    if ($BadConfig) {
        [IO.File]::WriteAllText((Join-Path $Path "bindings.resolved.json"), "{")
    } else {
        Write-Utf8Json (Join-Path $Path "bindings.resolved.json") ([ordered]@{
            schema = 1
            root = [IO.Path]::GetFullPath($RepositoryRoot)
            sourceConfigPath = (Join-Path $RepositoryRoot "bindings.json")
            sourceConfigSha256 = ("0" * 64)
            active = @()
            rejected = @()
        })
    }
}

function Invoke-ChildCapture {
    param([string[]]$Arguments)
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName = "powershell.exe"
    $info.Arguments = ($Arguments | ForEach-Object {
        if ($_ -match '[\s"]') { '"' + ($_ -replace '"', '\"') + '"' } else { $_ }
    }) -join " "
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $process = [Diagnostics.Process]::Start($info)
    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    $process.WaitForExit()
    return @{ ExitCode = $process.ExitCode; StdOut = $stdout; StdErr = $stderr }
}

function Start-HostProcess {
    param([string]$StateRoot, [switch]$BadSource)
    $items = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden", "-File", $hostScript, "-StateRoot", $StateRoot)
    if ($BadSource) { $items += "-TestBadSource" }
    return Start-Process -FilePath powershell.exe -ArgumentList $items -WindowStyle Hidden -PassThru
}

function Get-HostStatus {
    param([string]$StateRoot)
    $output = @(& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $controlScript -Status -StateRoot $StateRoot 2>&1)
    if ($LASTEXITCODE -ne 0) { return $null }
    $line = @($output | Where-Object { ([string]$_).Trim().StartsWith("{") } | Select-Object -Last 1)
    if ($line.Count -ne 1) { return $null }
    try { return ConvertFrom-Json ([string]$line[0]) } catch { return $null }
}

function Wait-HostReady {
    param([string]$StateRoot)
    $deadline = [DateTime]::UtcNow.AddSeconds(8)
    do {
        $status = Get-HostStatus $StateRoot
        if ($null -ne $status -and [bool]$status.ready) { return $status }
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $deadline)
    return $null
}

$fx = Join-Path $env:TEMP ("toolrack_host_runtime_" + [guid]::NewGuid().ToString("N"))
$hostProcess = $null
$vbsHostStatus = $null
$stallJob = $null
try {
    New-Item -ItemType Directory -Force -Path $fx | Out-Null
    $namespace = "test-" + [guid]::NewGuid().ToString("N")
    $state = Join-Path $fx "state"
    New-HostState $state $root $namespace

    $selfTest = Invoke-ChildCapture @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $hostScript, "-SelfTest", "-StateRoot", $state)
    Assert-True ($selfTest.ExitCode -eq 0) "Host SelfTest exits zero"
    $selfJson = $null
    try { $selfJson = ConvertFrom-Json $selfTest.StdOut.Trim() } catch {}
    Assert-True ($null -ne $selfJson -and [bool]$selfJson.ok) "Host SelfTest returns JSON"
    if ($null -ne $selfJson) {
        Assert-True ([string]$selfJson.root -eq [IO.Path]::GetFullPath($root)) "SelfTest reports root"
        Assert-True ([string]$selfJson.version -eq "1") "SelfTest reports version"
        Assert-True ([bool]$selfJson.pipeCurrentUserOnly) "SelfTest reports current-user pipe ACL"
    }

    $badState = Join-Path $fx "bad-state"
    New-HostState $badState $root ("bad-" + [guid]::NewGuid().ToString("N")) -BadConfig
    $badResult = Invoke-ChildCapture @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $hostScript, "-SelfTest", "-StateRoot", $badState)
    Assert-True ($badResult.ExitCode -eq 1 -and $badResult.StdErr -match 'bindings') "malformed normalized config fails cleanly"

    $sourceFailure = Invoke-ChildCapture @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $hostScript, "-SelfTest", "-StateRoot", $state, "-TestBadSource")
    Assert-True ($sourceFailure.ExitCode -eq 1 -and $sourceFailure.StdErr -match 'Add-Type') "Add-Type failure seam is clean"

    $hostProcess = Start-HostProcess $state
    $status = Wait-HostReady $state
    Assert-True ($null -ne $status) "Host becomes ready"
    if ($null -ne $status) {
        Assert-True ([int]$status.pid -eq $hostProcess.Id) "status reports Host PID"
        Assert-True ([string]$status.root -eq [IO.Path]::GetFullPath($root)) "status reports root"
        Assert-True ([string]$status.version -eq "1") "status reports version"
        Assert-True ([int]$status.activeBindings -eq 0 -and [int]$status.rejectedBindings -eq 0) "status reports binding counts"
        Assert-True ([bool]$status.pipeCurrentUserOnly) "runtime pipe ACL is current-user only"

        $stallReady = Join-Path $fx "stall-ready.txt"
        $stallJob = Start-Job -ArgumentList ([string]$status.pipeName), $stallReady -ScriptBlock {
            param($name, $readyPath)
            $client = New-Object IO.Pipes.NamedPipeClientStream(".", $name, [IO.Pipes.PipeDirection]::InOut)
            try {
                $client.Connect(3000)
                [IO.File]::WriteAllText($readyPath, "ready")
                Start-Sleep -Seconds 6
            } finally {
                $client.Dispose()
            }
        }
        $stallDeadline = [DateTime]::UtcNow.AddSeconds(5)
        while (-not (Test-Path -LiteralPath $stallReady) -and [DateTime]::UtcNow -lt $stallDeadline) {
            Start-Sleep -Milliseconds 50
        }
        $stallWatch = [Diagnostics.Stopwatch]::StartNew()
        $statusAfterStall = Get-HostStatus $state
        $stallWatch.Stop()
        Assert-True ($null -ne $statusAfterStall -and $stallWatch.Elapsed.TotalSeconds -lt 4.5) "stalled pipe client is timed out and control recovers"
        Stop-Job $stallJob -ErrorAction SilentlyContinue
        Remove-Job $stallJob -Force -ErrorAction SilentlyContinue
        $stallJob = $null
    }

    $second = Start-HostProcess $state
    $secondExited = $second.WaitForExit(5000)
    Assert-True ($secondExited -and $second.ExitCode -eq 2) "second Host exits with code 2"
    Assert-True (-not $hostProcess.HasExited) "second Host does not disturb primary"

    $reloadOutput = @(& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $controlScript -Reload -StateRoot $state 2>&1)
    Assert-True ($LASTEXITCODE -eq 0 -and (@($reloadOutput | Where-Object { [string]$_ -match '"ok":true' }).Count -gt 0)) "reload control command responds"

    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $controlScript -Shutdown -StateRoot $state | Out-Null
    Assert-True ($LASTEXITCODE -eq 0) "shutdown control command succeeds"
    Assert-True ($hostProcess.WaitForExit(5000)) "Host exits after shutdown"
    $after = Get-HostStatus $state
    Assert-True ($null -eq $after) "pipe is gone after shutdown"
    $hostProcess = $null

    $unicodeLeaf = ([char]0x65E5) + ([char]0x672C) + " repo"
    $repoCopy = Join-Path $fx $unicodeLeaf
    New-Item -ItemType Directory -Force -Path (Join-Path $repoCopy "common") | Out-Null
    Copy-Item -LiteralPath $hostScript -Destination (Join-Path $repoCopy "common\host.ps1")
    Copy-Item -LiteralPath $sourceDir -Destination (Join-Path $repoCopy "common\host") -Recurse
    Copy-Item -LiteralPath (Join-Path $root "common\launch-core.ps1") -Destination (Join-Path $repoCopy "common\launch-core.ps1")
    Copy-Item -LiteralPath (Join-Path $root "common\resolve-host-config.ps1") -Destination (Join-Path $repoCopy "common\resolve-host-config.ps1")

    $localRoot = Join-Path $fx "local root"
    $bootstrapDir = Join-Path $localRoot "bootstrap\v1"
    $generation = [guid]::NewGuid().ToString("N")
    $vbsState = Join-Path $localRoot ("state\" + $generation)
    New-Item -ItemType Directory -Force -Path $bootstrapDir | Out-Null
    New-HostState $vbsState $repoCopy ("vbs-" + [guid]::NewGuid().ToString("N"))
    [IO.File]::WriteAllText((Join-Path $vbsState "root.txt"), ([IO.Path]::GetFullPath($repoCopy) + "\"), (New-Object Text.UTF8Encoding($false)))
    [IO.File]::WriteAllText((Join-Path $localRoot "active.txt"), $generation, (New-Object Text.UTF8Encoding($false)))
    Copy-Item -LiteralPath $bootstrapScript -Destination (Join-Path $bootstrapDir "host-bootstrap.ps1")
    Copy-Item -LiteralPath $startVbs -Destination (Join-Path $bootstrapDir "start-host.vbs")

    & cscript.exe //nologo (Join-Path $bootstrapDir "start-host.vbs") | Out-Null
    Assert-True ($LASTEXITCODE -eq 0) "VBS bootstrap parses and starts"
    $vbsHostStatus = Wait-HostReady $vbsState
    Assert-True ($null -ne $vbsHostStatus) "VBS starts Host through Unicode root"
    if ($null -ne $vbsHostStatus) {
        Assert-True ([string]$vbsHostStatus.root -eq [IO.Path]::GetFullPath($repoCopy)) "bootstrap preserves Unicode root"
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $controlScript -Shutdown -StateRoot $vbsState | Out-Null
        $vbsPid = [int]$vbsHostStatus.pid
        $deadline = [DateTime]::UtcNow.AddSeconds(5)
        while ($null -ne (Get-Process -Id $vbsPid -ErrorAction SilentlyContinue) -and [DateTime]::UtcNow -lt $deadline) {
            Start-Sleep -Milliseconds 100
        }
        Assert-True ($null -eq (Get-Process -Id $vbsPid -ErrorAction SilentlyContinue)) "VBS-started Host shuts down"
        $vbsHostStatus = $null
    }

    $vbsText = [IO.File]::ReadAllText($startVbs)
    Assert-True ($vbsText -match 'WindowStyle Hidden' -and $vbsText -match 'sh\.Run cmd, 0, False') "VBS requests two layers of hidden startup"
} finally {
    if ($null -ne $stallJob) {
        Stop-Job $stallJob -ErrorAction SilentlyContinue
        Remove-Job $stallJob -Force -ErrorAction SilentlyContinue
    }
    if ($null -ne $hostProcess -and -not $hostProcess.HasExited) {
        try { $hostProcess.Kill() } catch {}
    }
    if ($null -ne $vbsHostStatus) {
        try { Stop-Process -Id ([int]$vbsHostStatus.pid) -Force -ErrorAction SilentlyContinue } catch {}
    }
    Remove-Item -LiteralPath $fx -Recurse -Force -ErrorAction SilentlyContinue
}

Exit-Test
