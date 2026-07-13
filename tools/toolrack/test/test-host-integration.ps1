# test/test-host-integration.ps1 -- activation worker and safe reload integration.
. (Join-Path $PSScriptRoot "_assert.ps1")

$root = Split-Path $PSScriptRoot -Parent
$hostScript = Join-Path $root "common\host.ps1"
$controlScript = Join-Path $root "common\host-control.ps1"
$workerSource = Join-Path $root "common\host\ActivationWorker.cs"
$hostSource = Join-Path $root "common\host\HostApplication.cs"
$pipeSource = Join-Path $root "common\host\HostPipe.cs"
$launchCore = Join-Path $root "common\launch-core.ps1"

$required = @($workerSource, $hostSource, $pipeSource, $launchCore)
foreach ($path in $required) {
    Assert-True (Test-Path -LiteralPath $path -PathType Leaf) ("Host integration asset exists: " + (Split-Path $path -Leaf))
}
if (@($required | Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) }).Count -gt 0) { Exit-Test }

foreach ($path in $required) {
    Assert-True (@([IO.File]::ReadAllBytes($path) | Where-Object { $_ -gt 127 }).Count -eq 0) ("Host integration asset is ASCII: " + (Split-Path $path -Leaf))
}
$workerText = [IO.File]::ReadAllText($workerSource)
Assert-True ($workerText -notmatch 'launch\.ps1') "worker does not start the Explorer wrapper"
Assert-True ($workerText -notmatch 'tool\.json|ConvertFrom-Json') "Host C# does not parse raw tool manifests"
Assert-True ($workerText -match 'System\.Management\.Automation') "worker uses the in-process Windows PowerShell runspace"
. $launchCore

function Write-Utf8Text {
    param([string]$Path, [string]$Text)
    $parent = Split-Path $Path -Parent
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        New-Item -ItemType Directory -Force -Path $parent | Out-Null
    }
    [IO.File]::WriteAllText($Path, $Text, (New-Object Text.UTF8Encoding($false)))
}

function Write-Utf8Json {
    param([string]$Path, $Value)
    Write-Utf8Text $Path (ConvertTo-Json $Value -Depth 12)
}

function Write-FixtureBindings {
    param([string]$Path, [string]$Action)
    Write-Utf8Json $Path ([ordered]@{
        schema = 1
        bindings = @(
            [ordered]@{
                id = "marker-hotkey"
                trigger = [ordered]@{ type = "hotkey"; key = "F13"; modifiers = @("ctrl", "alt", "shift") }
                invoke = [ordered]@{ tool = "marker"; action = $Action }
            },
            [ordered]@{
                id = "marker-slow-hotkey"
                trigger = [ordered]@{ type = "hotkey"; key = "F15"; modifiers = @("ctrl", "alt", "shift") }
                invoke = [ordered]@{ tool = "marker"; action = "slow" }
            },
            [ordered]@{
                id = "rejected-hotkey"
                trigger = [ordered]@{ type = "hotkey"; key = "F14"; modifiers = @("ctrl", "alt", "shift") }
                invoke = [ordered]@{ tool = "missing"; action = "default" }
            }
        )
    })
}

function Get-PipeName {
    param([string]$Namespace)
    $sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    return "ToolRackHost-" + $sid + "-" + $Namespace
}

function Send-HostCommand {
    param([string]$PipeName, [string]$Command, [int]$TimeoutMilliseconds = 2000)
    $client = New-Object IO.Pipes.NamedPipeClientStream(".", $PipeName, [IO.Pipes.PipeDirection]::InOut)
    try {
        $client.Connect($TimeoutMilliseconds)
        $writer = New-Object IO.StreamWriter($client, (New-Object Text.UTF8Encoding($false)), 1024, $true)
        $reader = New-Object IO.StreamReader($client, (New-Object Text.UTF8Encoding($false, $true)), $false, 1024, $true)
        try {
            $writer.AutoFlush = $true
            $writer.WriteLine($Command)
            $line = $reader.ReadLine()
        } finally {
            $writer.Dispose()
            $reader.Dispose()
        }
    } finally {
        $client.Dispose()
    }
    return ConvertFrom-Json $line
}

function Wait-HostStatus {
    param([string]$PipeName, [scriptblock]$Predicate, [int]$Seconds = 10)
    $deadline = [DateTime]::UtcNow.AddSeconds($Seconds)
    do {
        try {
            $status = Send-HostCommand $PipeName "status" 300
            if (& $Predicate $status) { return $status }
        } catch { }
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $deadline)
    return $null
}

function Get-MarkerFiles {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { return @() }
    return @(Get-ChildItem -LiteralPath $Path -Filter "*.json" -File)
}

function Wait-MarkerCount {
    param([string]$Path, [int]$Count, [int]$Seconds = 10)
    $deadline = [DateTime]::UtcNow.AddSeconds($Seconds)
    do {
        $files = @(Get-MarkerFiles $Path)
        if ($files.Count -ge $Count) { return $files }
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $deadline)
    return @(Get-MarkerFiles $Path)
}

$fx = Join-Path $env:TEMP ("toolrack host bridge '" + [char]0x65E5 + "' " + [guid]::NewGuid().ToString("N"))
$hostProcess = $null
$productionHost = $null
$markerPids = New-Object "System.Collections.Generic.HashSet[int]"
try {
    $fixtureRoot = Join-Path $fx "repo root"
    $fixtureCommon = Join-Path $fixtureRoot "common"
    $fixtureTool = Join-Path $fixtureRoot "tool\marker"
    $markerDir = Join-Path $fx "marker output"
    $state = Join-Path $fx "state"
    New-Item -ItemType Directory -Force -Path $fixtureCommon, $fixtureTool, $markerDir, $state | Out-Null

    foreach ($name in @("launch-core.ps1", "resolve-host-config.ps1", "bindings.ps1", "install.ps1")) {
        Copy-Item -LiteralPath (Join-Path $root ("common\" + $name)) -Destination (Join-Path $fixtureCommon $name)
    }

    $markerScript = @'
param(
    [AllowEmptyString()][string]$Target = "",
    [Parameter(Mandatory = $true)][string]$OutputDir,
    [Parameter(Mandatory = $true)][string]$Mode,
    [Parameter(Mandatory = $true)][string]$Value,
    [int]$DelayMs = 0
)
$parent = Get-WmiObject Win32_Process -Filter ("ProcessId=" + $PID)
$record = [ordered]@{
    pid = $PID
    parentPid = [int]$parent.ParentProcessId
    target = $Target
    mode = $Mode
    value = $Value
}
$path = Join-Path $OutputDir (([guid]::NewGuid().ToString("N")) + ".json")
[IO.File]::WriteAllText($path, (ConvertTo-Json $record -Compress), (New-Object Text.UTF8Encoding($false)))
if ($DelayMs -gt 0) { Start-Sleep -Milliseconds $DelayMs }
'@
    Write-Utf8Text (Join-Path $fixtureTool "main.ps1") $markerScript
    $manifest = [ordered]@{
        schema = 1
        id = "marker"
        name = "Marker"
        on = @("background")
        run = [ordered]@{ type = "powershell"; entry = "main.ps1"; window = "gui" }
        variants = @(
            [ordered]@{ id = "fast"; label = "Fast"; args = @("-OutputDir", $markerDir, "-Mode", "fast", "-Value", "argument with spaces") },
            [ordered]@{ id = "slow"; label = "Slow"; args = @("-OutputDir", $markerDir, "-Mode", "slow", "-Value", "argument with spaces", "-DelayMs", "1500") }
        )
    }
    $manifestPath = Join-Path $fixtureTool "tool.json"
    Write-Utf8Json $manifestPath $manifest
    $goodManifestText = [IO.File]::ReadAllText($manifestPath)

    $bindingsPath = Join-Path $fixtureRoot "bindings.json"
    $resolvedPath = Join-Path $state "bindings.resolved.json"
    Write-Utf8Text $bindingsPath "{"
    $resolverArguments = @(
        "-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden", "-File",
        (Join-Path $fixtureCommon "resolve-host-config.ps1"), "-Root", $fixtureRoot,
        "-BindingsPath", $bindingsPath, "-OutputPath", $resolvedPath
    )
    $resolverProcess = Start-Process -FilePath powershell.exe -ArgumentList (ConvertTo-ArgString $resolverArguments) -WindowStyle Hidden -PassThru
    Start-Sleep -Milliseconds 100
    Write-FixtureBindings $bindingsPath "fast"
    $resolverExited = $resolverProcess.WaitForExit(10000)
    Assert-True ($resolverExited -and $resolverProcess.ExitCode -eq 0 -and (Test-Path -LiteralPath $resolvedPath -PathType Leaf)) "resolver retries a transient partial binding write"

    $namespace = "bridge-" + [guid]::NewGuid().ToString("N")
    Write-Utf8Json (Join-Path $state "host.json") ([ordered]@{
        schema = 1
        version = "1"
        root = [IO.Path]::GetFullPath($fixtureRoot)
        namespace = $namespace
    })
    $pipeName = Get-PipeName $namespace
    $hostErrorPath = Join-Path $fx "host.stderr.txt"
    $hostArguments = @(
        "-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden", "-File", $hostScript,
        "-StateRoot", $state, "-TestMode"
    )
    $hostProcess = Start-Process -FilePath powershell.exe -ArgumentList (ConvertTo-ArgString $hostArguments) `
        -WindowStyle Hidden -RedirectStandardError $hostErrorPath -PassThru

    $status = Wait-HostStatus $pipeName { param($s) [bool]$s.ready }
    Assert-True ($null -ne $status) "test Host becomes ready"
    if ($null -eq $status) {
        if (Test-Path -LiteralPath $hostErrorPath -PathType Leaf) { Write-Host ([IO.File]::ReadAllText($hostErrorPath)) }
        throw "test Host did not become ready"
    }
    Assert-True ([int]$status.activeBindings -eq 2 -and [int]$status.rejectedBindings -eq 1) "Host reads normalized active and rejected bindings"

    $initialCount = @(Get-MarkerFiles $markerDir).Count
    $activation = Send-HostCommand $pipeName "test-activate|marker|fast"
    Assert-True ([bool]$activation.ok -and [bool]$activation.queued) "test activation queues tool and action"
    $files = @(Wait-MarkerCount $markerDir ($initialCount + 1))
    Assert-True ($files.Count -eq ($initialCount + 1)) "activation launches exactly one marker process"
    $first = ConvertFrom-Json ([IO.File]::ReadAllText(($files | Sort-Object LastWriteTimeUtc | Select-Object -Last 1).FullName))
    [void]$markerPids.Add([int]$first.pid)
    Assert-True ([string]$first.target -eq "") "Host route passes an empty Target"
    Assert-True ([string]$first.mode -eq "fast" -and [string]$first.value -eq "argument with spaces") "stable action arguments survive quoting"
    Assert-True ([int]$first.parentPid -eq $hostProcess.Id) "tool process is the direct child of Host"

    $unknown = Send-HostCommand $pipeName "test-activate|missing|default"
    Assert-True (-not [bool]$unknown.ok -and -not [bool]$unknown.queued) "unknown or rejected targets do not launch"
    Start-Sleep -Milliseconds 300
    Assert-True (@(Get-MarkerFiles $markerDir).Count -eq ($initialCount + 1)) "rejected activation creates no process"

    $beforeSlow = @(Get-MarkerFiles $markerDir).Count
    $slow = Send-HostCommand $pipeName "test-activate|marker|slow"
    $slowFiles = @(Wait-MarkerCount $markerDir ($beforeSlow + 1))
    $slowRecord = ConvertFrom-Json ([IO.File]::ReadAllText(($slowFiles | Sort-Object LastWriteTimeUtc | Select-Object -Last 1).FullName))
    [void]$markerPids.Add([int]$slowRecord.pid)
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $responsive = Send-HostCommand $pipeName "status"
    $watch.Stop()
    $slowProcess = Get-Process -Id ([int]$slowRecord.pid) -ErrorAction SilentlyContinue
    Assert-True ([bool]$slow.queued -and $watch.ElapsedMilliseconds -lt 500) "GUI tool launch does not block the Host message loop"
    Assert-True ($null -ne $slowProcess -and $slowProcess.MainWindowHandle -eq 0) "GUI tool process has no console window"

    $failureBefore = [int]$responsive.activationFailed
    Write-Utf8Text $manifestPath "{"
    [void](Send-HostCommand $pipeName "test-activate|marker|fast")
    $failedStatus = Wait-HostStatus $pipeName { param($s) [int]$s.activationFailed -gt $failureBefore }
    Assert-True ($null -ne $failedStatus -and [int]$failedStatus.notificationCount -gt 0) "launch preflight failure reaches Host status and notification seam"
    Write-Utf8Text $manifestPath $goodManifestText

    $countBeforeThrow = @(Get-MarkerFiles $markerDir).Count
    [void](Send-HostCommand $pipeName "test-throw-next")
    [void](Send-HostCommand $pipeName "test-activate|marker|fast")
    [void](Send-HostCommand $pipeName "test-activate|marker|fast")
    $afterThrow = @(Wait-MarkerCount $markerDir ($countBeforeThrow + 1))
    Assert-True ($afterThrow.Count -ge ($countBeforeThrow + 1)) "request exception does not poison the next activation"
    $throwStatus = Send-HostCommand $pipeName "status"
    $generationBefore = [int]$throwStatus.workerGeneration
    $countBeforeBreak = @(Get-MarkerFiles $markerDir).Count
    [void](Send-HostCommand $pipeName "test-break-worker")
    [void](Send-HostCommand $pipeName "test-activate|marker|fast")
    $afterBreak = @(Wait-MarkerCount $markerDir ($countBeforeBreak + 1))
    $recovered = Wait-HostStatus $pipeName { param($s) [int]$s.workerGeneration -gt $generationBefore }
    Assert-True ($afterBreak.Count -ge ($countBeforeBreak + 1) -and $null -ne $recovered) "closed worker runspace is recreated and launch core reloads"

    $reloadBefore = [int]$recovered.reloadSucceeded
    Write-FixtureBindings $bindingsPath "slow"
    $reloaded = Wait-HostStatus $pipeName { param($s) [int]$s.reloadSucceeded -gt $reloadBefore } 12
    Assert-True ($null -ne $reloaded) "valid bindings change reloads after debounce"
    $resolvedHash = (Get-FileHash -LiteralPath $resolvedPath -Algorithm SHA256).Hash

    $reloadFailureBefore = [int]$reloaded.reloadFailed
    Write-Utf8Text $bindingsPath "{"
    $badReload = Wait-HostStatus $pipeName { param($s) [int]$s.reloadFailed -gt $reloadFailureBefore } 12
    $hashAfterBad = (Get-FileHash -LiteralPath $resolvedPath -Algorithm SHA256).Hash
    Assert-True ($null -ne $badReload -and $hashAfterBad -eq $resolvedHash) "malformed reload preserves last-known-good normalized config"

    $failureCountBeforePartial = [int]$badReload.reloadFailed
    $successBeforePartial = [int]$badReload.reloadSucceeded
    Write-Utf8Text $bindingsPath "{"
    Start-Sleep -Milliseconds 100
    Write-FixtureBindings $bindingsPath "fast"
    $partialRecovered = Wait-HostStatus $pipeName { param($s) [int]$s.reloadSucceeded -gt $successBeforePartial } 12
    Assert-True ($null -ne $partialRecovered -and [int]$partialRecovered.reloadFailed -eq $failureCountBeforePartial) "partial write inside debounce window does not discard config"

    [void](Send-HostCommand $pipeName "test-pause-worker")
    $rejected = 0
    foreach ($index in 1..24) {
        $queued = Send-HostCommand $pipeName "test-activate|marker|fast"
        if (-not [bool]$queued.queued) { $rejected++ }
    }
    $queueWatch = [Diagnostics.Stopwatch]::StartNew()
    $queueStatus = Send-HostCommand $pipeName "status"
    $queueWatch.Stop()
    Assert-True ($rejected -gt 0 -and [int]$queueStatus.activationRejected -gt 0) "bounded activation queue rejects overflow"
    Assert-True ($queueWatch.ElapsedMilliseconds -lt 500) "queue overflow does not block Host control messages"
    [void](Send-HostCommand $pipeName "test-resume-worker")

    [void](Send-HostCommand $pipeName "shutdown")
    Assert-True ($hostProcess.WaitForExit(5000)) "test Host shuts down after draining or cancelling worker"
    $hostProcess = $null

    $productionArguments = @(
        "-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden", "-File", $hostScript, "-StateRoot", $state
    )
    $productionHost = Start-Process -FilePath powershell.exe -ArgumentList (ConvertTo-ArgString $productionArguments) -WindowStyle Hidden -PassThru
    $productionStatus = Wait-HostStatus $pipeName { param($s) [bool]$s.ready }
    Assert-True ($null -ne $productionStatus) "production-mode Host becomes ready"
    $productionActivation = Send-HostCommand $pipeName "test-activate|marker|fast"
    Assert-True (-not [bool]$productionActivation.ok) "production control pipe rejects arbitrary activation"
    [void](Send-HostCommand $pipeName "shutdown")
    Assert-True ($productionHost.WaitForExit(5000)) "production Host shuts down"
    $productionHost = $null
} finally {
    if ($null -ne $hostProcess -and -not $hostProcess.HasExited) { try { $hostProcess.Kill() } catch { } }
    if ($null -ne $productionHost -and -not $productionHost.HasExited) { try { $productionHost.Kill() } catch { } }
    foreach ($pidValue in $markerPids) {
        Stop-Process -Id $pidValue -Force -ErrorAction SilentlyContinue
    }
    Remove-Item -LiteralPath $fx -Recurse -Force -ErrorAction SilentlyContinue
}

Exit-Test
