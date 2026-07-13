# test/probe-global-host.ps1 -- isolated workplace gate for the planned host.
[CmdletBinding()]
param(
    [switch]$SelfTest,
    [switch]$Install,
    [switch]$Run,
    [switch]$Verify,
    [switch]$Latency,
    [switch]$Remove,
    [string]$StateRoot = (Join-Path $env:LOCALAPPDATA "ToolRackHostProbe"),
    [string]$RunValueName = "ToolRackHostProbe"
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

function ConvertTo-WindowsArgument {
    param([string]$Item)
    if ($null -eq $Item -or $Item.Length -eq 0) { return '""' }
    if ($Item -notmatch '[\s"]') { return $Item }

    $builder = New-Object System.Text.StringBuilder
    [void]$builder.Append('"')
    $backslashes = 0
    foreach ($char in $Item.ToCharArray()) {
        if ($char -eq '\') {
            $backslashes++
        } elseif ($char -eq '"') {
            [void]$builder.Append('\' * (($backslashes * 2) + 1))
            [void]$builder.Append('"')
            $backslashes = 0
        } else {
            if ($backslashes -gt 0) {
                [void]$builder.Append('\' * $backslashes)
                $backslashes = 0
            }
            [void]$builder.Append($char)
        }
    }
    if ($backslashes -gt 0) { [void]$builder.Append('\' * ($backslashes * 2)) }
    [void]$builder.Append('"')
    return $builder.ToString()
}

function ConvertTo-ArgString {
    param([string[]]$Items)
    return (@($Items | ForEach-Object { ConvertTo-WindowsArgument $_ }) -join " ")
}

function Get-PowerShellPath {
    $path = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Windows PowerShell 5.1 was not found: $path"
    }
    return $path
}

function Get-FixturePath {
    param([string]$Name)
    $repoPath = Join-Path $PSScriptRoot ("fixture\host-latency\" + $Name)
    if (Test-Path -LiteralPath $repoPath -PathType Leaf) { return $repoPath }
    $localPath = Join-Path $StateRoot $Name
    if (Test-Path -LiteralPath $localPath -PathType Leaf) { return $localPath }
    throw "probe fixture not found: $Name"
}

function Import-HostProbeType {
    if ($null -ne ("ToolRackProbe.HostProbe" -as [type])) { return }
    $source = Get-FixturePath "HostProbe.cs"
    Add-Type -Path $source -ReferencedAssemblies @(
        "System.dll",
        "System.Core.dll",
        "System.Windows.Forms.dll",
        "System.Drawing.dll"
    )
}

function Write-JsonLine {
    param($Value)
    Write-Output (ConvertTo-Json $Value -Compress)
}

function Get-Heartbeat {
    $path = Join-Path $StateRoot "heartbeat.json"
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return $null }
    try { return ConvertFrom-Json ([IO.File]::ReadAllText($path)) }
    catch { return $null }
}

function Test-PersistentRunspace {
    $runspace = $null
    $pipeline = $null
    try {
        $runspace = [RunspaceFactory]::CreateRunspace()
        $runspace.ThreadOptions = [Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
        $runspace.ApartmentState = [Threading.ApartmentState]::MTA
        $runspace.Open()
        $pipeline = [PowerShell]::Create()
        $pipeline.Runspace = $runspace
        [void]$pipeline.AddScript('param([int]$Value) $Value + 1').AddArgument(41)
        $output = @($pipeline.Invoke())
        return (-not $pipeline.HadErrors -and $output.Count -eq 1 -and [int]$output[0] -eq 42)
    } catch {
        return $false
    } finally {
        if ($null -ne $pipeline) { $pipeline.Dispose() }
        if ($null -ne $runspace) { $runspace.Dispose() }
    }
}

function Test-RunValueName {
    if ($RunValueName -notmatch '^ToolRackHostProbe[A-Za-z0-9-]*$') {
        throw "unsafe probe Run value name: $RunValueName"
    }
}

function Get-RunRegistryPath {
    return "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
}

function Invoke-SelfTest {
    New-Item -ItemType Directory -Force -Path $StateRoot | Out-Null
    $work = Join-Path $StateRoot ("selftest_" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Force -Path $work | Out-Null
    try {
        Import-HostProbeType
        $probeResult = [ToolRackProbe.HostProbe]::Run($work, 300)
        $runspaceAvailable = Test-PersistentRunspace
        $value = [ordered]@{
            AddType = [bool]$probeResult.AddType
            HookInstalled = [bool]$probeResult.HookInstalled
            HotkeyAvailable = [bool]$probeResult.HotkeyAvailable
            MessageLoop = [bool]$probeResult.MessageLoop
            Cleanup = [bool]$probeResult.Cleanup
            DwmAvailable = [bool]$probeResult.DwmAvailable
            RunspaceAvailable = $runspaceAvailable
            ProcessId = [int]$probeResult.ProcessId
        }
        Write-JsonLine $value
        if (-not $value.AddType -or -not $value.HookInstalled -or -not $value.HotkeyAvailable -or
            -not $value.MessageLoop -or -not $value.Cleanup -or -not $value.RunspaceAvailable) {
            exit 2
        }
    } finally {
        Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Invoke-Install {
    Test-RunValueName
    $fullState = [IO.Path]::GetFullPath($StateRoot)
    $rootPath = [IO.Path]::GetPathRoot($fullState)
    if ($fullState.TrimEnd('\') -eq $rootPath.TrimEnd('\')) {
        throw "refusing to use a drive root as probe state"
    }
    New-Item -ItemType Directory -Force -Path $fullState | Out-Null
    [IO.File]::WriteAllText((Join-Path $fullState ".toolrack-host-probe"), "v1", (New-Object Text.UTF8Encoding($false)))

    $assets = @(
        @($PSCommandPath, (Join-Path $fullState "probe-global-host.ps1")),
        @((Get-FixturePath "HostProbe.cs"), (Join-Path $fullState "HostProbe.cs")),
        @((Get-FixturePath "start-probe.vbs"), (Join-Path $fullState "start-probe.vbs"))
    )
    foreach ($asset in $assets) {
        $source = [IO.Path]::GetFullPath([string]$asset[0])
        $destination = [IO.Path]::GetFullPath([string]$asset[1])
        if ($source -ne $destination) {
            Copy-Item -LiteralPath $source -Destination $destination -Force
        }
    }
    Remove-Item -LiteralPath (Join-Path $fullState "stop.request") -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath (Join-Path $fullState "heartbeat.json") -Force -ErrorAction SilentlyContinue

    $vbsPath = Join-Path $fullState "start-probe.vbs"
    $command = "wscript.exe " + (ConvertTo-WindowsArgument $vbsPath)
    if ($command.Length -ge 260) { throw "probe Run command is 260 characters or longer" }
    $registryPath = Get-RunRegistryPath
    New-Item -Path $registryPath -Force | Out-Null
    New-ItemProperty -Path $registryPath -Name $RunValueName -Value $command -PropertyType String -Force | Out-Null

    Write-JsonLine ([ordered]@{
        Installed = $true
        StateRoot = $fullState
        RunValueName = $RunValueName
        RunCommandLength = $command.Length
        Next = "sign out, sign in, then run -Verify and -Latency"
    })
}

function Invoke-Run {
    New-Item -ItemType Directory -Force -Path $StateRoot | Out-Null
    Remove-Item -LiteralPath (Join-Path $StateRoot "stop.request") -Force -ErrorAction SilentlyContinue
    Import-HostProbeType
    $probeResult = [ToolRackProbe.HostProbe]::Run($StateRoot, 0)
    if ([bool]$probeResult.AlreadyRunning) { exit 3 }
}

function Invoke-Verify {
    Test-RunValueName
    $errors = New-Object System.Collections.Generic.List[string]
    $expected = @(".toolrack-host-probe", "probe-global-host.ps1", "HostProbe.cs", "start-probe.vbs")
    foreach ($name in $expected) {
        if (-not (Test-Path -LiteralPath (Join-Path $StateRoot $name) -PathType Leaf)) {
            [void]$errors.Add("missing or quarantined asset: $name")
        }
    }

    $heartbeat = Get-Heartbeat
    $ageSeconds = $null
    $processAlive = $false
    if ($null -eq $heartbeat) {
        [void]$errors.Add("heartbeat.json is missing or invalid")
    } else {
        $ageSeconds = ([DateTime]::UtcNow - (New-Object DateTime([long]$heartbeat.UtcTicks, [DateTimeKind]::Utc))).TotalSeconds
        if (-not [bool]$heartbeat.Ready) { [void]$errors.Add("probe is not ready") }
        if (-not [bool]$heartbeat.HookInstalled) { [void]$errors.Add("WH_MOUSE_LL was not installed") }
        if (-not [bool]$heartbeat.HotkeyAvailable) { [void]$errors.Add("RegisterHotKey was not available") }
        if (-not [bool]$heartbeat.MessageLoop) { [void]$errors.Add("message loop did not tick") }
        if ($ageSeconds -gt 5) { [void]$errors.Add("heartbeat is stale") }
        $processAlive = $null -ne (Get-Process -Id ([int]$heartbeat.ProcessId) -ErrorAction SilentlyContinue)
        if (-not $processAlive) { [void]$errors.Add("probe PowerShell process is not running") }
    }

    $registryPath = Get-RunRegistryPath
    $runCommand = $null
    try { $runCommand = (Get-ItemProperty -Path $registryPath -Name $RunValueName -ErrorAction Stop).$RunValueName }
    catch { [void]$errors.Add("probe HKCU Run value is missing") }

    $value = [ordered]@{
        Passed = ($errors.Count -eq 0)
        Errors = @($errors)
        AssetsPresent = (@($expected | Where-Object { Test-Path -LiteralPath (Join-Path $StateRoot $_) -PathType Leaf }).Count -eq $expected.Count)
        ProcessAlive = $processAlive
        HeartbeatAgeSeconds = $ageSeconds
        RunRegistered = ($null -ne $runCommand)
    }
    Write-JsonLine $value
    if ($errors.Count -gt 0) { exit 2 }
}

function Stop-ProbeProcess {
    $heartbeat = Get-Heartbeat
    if ($null -eq $heartbeat) { return }
    $processId = [int]$heartbeat.ProcessId
    $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
    if ($null -eq $process) { return }

    [IO.File]::WriteAllText((Join-Path $StateRoot "stop.request"), "stop", (New-Object Text.UTF8Encoding($false)))
    $deadline = [DateTime]::UtcNow.AddSeconds(5)
    while ([DateTime]::UtcNow -lt $deadline -and $null -ne (Get-Process -Id $processId -ErrorAction SilentlyContinue)) {
        Start-Sleep -Milliseconds 100
    }
    if ($null -eq (Get-Process -Id $processId -ErrorAction SilentlyContinue)) { return }

    $safeToStop = $false
    try {
        $nativeProcess = Get-CimInstance Win32_Process -Filter ("ProcessId = " + $processId) -ErrorAction Stop
        if ($null -ne $nativeProcess -and [string]$nativeProcess.CommandLine -like "*probe-global-host.ps1*" -and
            [string]$nativeProcess.CommandLine -like "*-Run*") {
            $safeToStop = $true
        }
    } catch {
        $safeToStop = $false
    }
    if (-not $safeToStop) {
        throw "probe did not stop and its process identity could not be verified; state was preserved"
    }
    Stop-Process -Id $processId -Force -ErrorAction Stop
}

function Invoke-Remove {
    Test-RunValueName
    $registryPath = Get-RunRegistryPath
    Remove-ItemProperty -Path $registryPath -Name $RunValueName -ErrorAction SilentlyContinue
    if (-not (Test-Path -LiteralPath $StateRoot)) {
        Write-JsonLine ([ordered]@{ Removed = $true; StateRoot = [IO.Path]::GetFullPath($StateRoot) })
        return
    }
    $marker = Join-Path $StateRoot ".toolrack-host-probe"
    if (-not (Test-Path -LiteralPath $marker -PathType Leaf) -or [IO.File]::ReadAllText($marker).Trim() -ne "v1") {
        throw "refusing to delete unowned probe state: $StateRoot"
    }
    Stop-ProbeProcess
    Remove-Item -LiteralPath $StateRoot -Recurse -Force
    Write-JsonLine ([ordered]@{ Removed = $true; StateRoot = [IO.Path]::GetFullPath($StateRoot) })
}

function New-LatencyWorker {
    $runspace = [RunspaceFactory]::CreateRunspace()
    $runspace.ThreadOptions = [Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $runspace.ApartmentState = [Threading.ApartmentState]::MTA
    $runspace.Open()
    $pipeline = [PowerShell]::Create()
    try {
        $pipeline.Runspace = $runspace
        $workerSource = @'
function global:ConvertTo-ProbeWindowsArgument {
    param([string]$Item)
    if ($null -eq $Item -or $Item.Length -eq 0) { return '""' }
    if ($Item -notmatch '[\s"]') { return $Item }
    $builder = New-Object System.Text.StringBuilder
    [void]$builder.Append('"')
    $backslashes = 0
    foreach ($char in $Item.ToCharArray()) {
        if ($char -eq '\') { $backslashes++ }
        elseif ($char -eq '"') {
            [void]$builder.Append('\' * (($backslashes * 2) + 1))
            [void]$builder.Append('"')
            $backslashes = 0
        } else {
            if ($backslashes -gt 0) {
                [void]$builder.Append('\' * $backslashes)
                $backslashes = 0
            }
            [void]$builder.Append($char)
        }
    }
    if ($backslashes -gt 0) { [void]$builder.Append('\' * ($backslashes * 2)) }
    [void]$builder.Append('"')
    return $builder.ToString()
}

function global:Invoke-ProbeLaunch {
    param([string]$ToolDir, [string]$PowerShellPath)
    try {
        $toolFull = [IO.Path]::GetFullPath($ToolDir)
        $manifestPath = Join-Path $toolFull "tool.json"
        $manifest = ConvertFrom-Json ([IO.File]::ReadAllText($manifestPath))
        if ($manifest.schema -ne 1) { throw "unsupported schema" }
        if ([string]$manifest.run.type -ne "powershell") { throw "fixture must be powershell" }
        if ([string]$manifest.run.window -ne "gui") { throw "fixture must be gui" }
        $entry = [string]$manifest.run.entry
        if ([string]::IsNullOrWhiteSpace($entry) -or [IO.Path]::IsPathRooted($entry)) { throw "invalid entry" }
        $entryFull = [IO.Path]::GetFullPath((Join-Path $toolFull $entry))
        $prefix = $toolFull.TrimEnd('\') + "\"
        if (-not $entryFull.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { throw "entry escaped tool folder" }
        if (-not (Test-Path -LiteralPath $entryFull -PathType Leaf)) { throw "entry not found" }

        $items = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden", "-File", $entryFull, "-Target", "")
        $arguments = (@($items | ForEach-Object { ConvertTo-ProbeWindowsArgument $_ }) -join " ")
        $info = New-Object Diagnostics.ProcessStartInfo
        $info.FileName = $PowerShellPath
        $info.Arguments = $arguments
        $info.UseShellExecute = $false
        $info.CreateNoWindow = $true
        $info.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
        $process = [Diagnostics.Process]::Start($info)
        return [pscustomobject]@{ Ok = $true; ProcessId = $process.Id; Error = "" }
    } catch {
        return [pscustomobject]@{ Ok = $false; ProcessId = 0; Error = $_.Exception.Message }
    }
}
'@
        [void]$pipeline.AddScript($workerSource)
        [void]$pipeline.Invoke()
        if ($pipeline.HadErrors) { throw "worker runspace initialization failed" }
        return $runspace
    } catch {
        $runspace.Dispose()
        throw
    } finally {
        $pipeline.Dispose()
    }
}

function Invoke-LatencyWorker {
    param($Runspace, [string]$ToolDir, [string]$PowerShellPath)
    $pipeline = [PowerShell]::Create()
    try {
        $pipeline.Runspace = $Runspace
        [void]$pipeline.AddCommand("Invoke-ProbeLaunch").AddParameter("ToolDir", $ToolDir).AddParameter("PowerShellPath", $PowerShellPath)
        $output = @($pipeline.Invoke())
        if ($pipeline.HadErrors -or $output.Count -ne 1) { throw "worker runspace invocation failed" }
        return $output[0]
    } finally {
        $pipeline.Dispose()
    }
}

function Invoke-Latency {
    $repositoryRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
    if (-not (Test-Path -LiteralPath (Join-Path $repositoryRoot "common\launch.ps1") -PathType Leaf)) {
        throw "run -Latency from the repository copy of probe-global-host.ps1"
    }
    $fixture = Join-Path $PSScriptRoot "fixture\host-latency"
    $work = Join-Path $env:TEMP ("toolrack_probe_latency_" + [guid]::NewGuid().ToString("N"))
    $tool = Join-Path $work "tool"
    New-Item -ItemType Directory -Force -Path $tool | Out-Null
    $oldMarker = $env:TOOLRACK_LATENCY_MARKER
    $samples = New-Object System.Collections.Generic.List[double]
    $worker = $null
    try {
        Copy-Item -LiteralPath (Join-Path $fixture "main.ps1") -Destination (Join-Path $tool "main.ps1")
        Copy-Item -LiteralPath (Join-Path $fixture "tool.json") -Destination (Join-Path $tool "tool.json")
        $dll = Join-Path $tool "LatencyPalette.dll"
        Add-Type -Path (Join-Path $fixture "LatencyPalette.cs") `
            -ReferencedAssemblies @("System.dll", "System.Windows.Forms.dll", "System.Drawing.dll") `
            -OutputAssembly $dll -OutputType Library

        $powerShellPath = Get-PowerShellPath
        $worker = New-LatencyWorker
        for ($index = 0; $index -lt 20; $index++) {
            $marker = Join-Path $work ("paint_" + $index + ".txt")
            $env:TOOLRACK_LATENCY_MARKER = $marker
            $started = [DateTime]::UtcNow.Ticks
            $launchResult = Invoke-LatencyWorker $worker $tool $powerShellPath
            if (-not [bool]$launchResult.Ok) { throw "worker failed at sample $index`: $($launchResult.Error)" }

            $deadline = [DateTime]::UtcNow.AddSeconds(10)
            while (-not (Test-Path -LiteralPath $marker -PathType Leaf) -and [DateTime]::UtcNow -lt $deadline) {
                Start-Sleep -Milliseconds 10
            }
            if (-not (Test-Path -LiteralPath $marker -PathType Leaf)) { throw "first paint timed out at sample $index" }
            $painted = [long]([IO.File]::ReadAllText($marker).Trim())
            [void]$samples.Add((New-Object TimeSpan($painted - $started)).TotalMilliseconds)
            $toolProcess = Get-Process -Id ([int]$launchResult.ProcessId) -ErrorAction SilentlyContinue
            if ($null -ne $toolProcess -and -not $toolProcess.WaitForExit(5000)) {
                try { $toolProcess.Kill() } catch {}
                throw "fixture tool did not exit at sample $index"
            }
        }

        $sorted = @($samples | Sort-Object)
        $p95Index = [Math]::Ceiling($sorted.Count * 0.95) - 1
        if ($p95Index -lt 0) { $p95Index = 0 }
        $first = [double]$samples[0]
        $p95 = [double]$sorted[$p95Index]
        $passed = ($first -le 800.0 -and $p95 -le 500.0)
        Write-JsonLine ([ordered]@{
            Passed = $passed
            Route = "resident-runspace-one-stage"
            SampleCount = $samples.Count
            ColdFirstMs = [Math]::Round($first, 2)
            P95Ms = [Math]::Round($p95, 2)
            MinimumMs = [Math]::Round([double]$sorted[0], 2)
            MaximumMs = [Math]::Round([double]$sorted[$sorted.Count - 1], 2)
            SamplesMs = @($samples | ForEach-Object { [Math]::Round([double]$_, 2) })
        })
        if (-not $passed) { exit 2 }
    } finally {
        if ($null -ne $worker) { $worker.Dispose() }
        $env:TOOLRACK_LATENCY_MARKER = $oldMarker
        Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
    }
}

$modes = @($SelfTest, $Install, $Run, $Verify, $Latency, $Remove)
$modeCount = @($modes | Where-Object { $_ }).Count
if ($modeCount -ne 1) {
    Write-Error "specify exactly one mode: -SelfTest, -Install, -Run, -Verify, -Latency, or -Remove"
    exit 1
}

try {
    if ($SelfTest) { Invoke-SelfTest }
    elseif ($Install) { Invoke-Install }
    elseif ($Run) { Invoke-Run }
    elseif ($Verify) { Invoke-Verify }
    elseif ($Latency) { Invoke-Latency }
    elseif ($Remove) { Invoke-Remove }
} catch {
    Write-Error $_.Exception.Message
    exit 1
}
