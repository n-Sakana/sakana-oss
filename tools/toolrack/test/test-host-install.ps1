# test/test-host-install.ps1 -- transactional Host install in isolated real HKCU keys.
. (Join-Path $PSScriptRoot "_assert.ps1")

$root = Split-Path $PSScriptRoot -Parent
$installScript = Join-Path $root "common\install.ps1"
$uninstallScript = Join-Path $root "common\uninstall.ps1"
. $installScript
$ErrorActionPreference = "Continue"

Assert-True ($null -ne (Get-Command New-InstallContext -ErrorAction SilentlyContinue)) "install exposes an isolated context seam"
Assert-True ($null -ne (Get-Command Invoke-Install -ErrorAction SilentlyContinue)) "install exposes transactional entry point"
if ($null -eq (Get-Command New-InstallContext -ErrorAction SilentlyContinue)) { Exit-Test }

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

function Get-RegistryDefault {
    param([string]$Path, [string]$Name)
    $oldEap = $ErrorActionPreference
    try {
        $ErrorActionPreference = "Continue"
        $query = @(& reg.exe query $Path /v $Name 2>$null)
        $queryCode = $LASTEXITCODE
    } finally { $ErrorActionPreference = $oldEap }
    if ($queryCode -ne 0) { return $null }
    $line = @($query | Where-Object { [string]$_ -match ("\s" + [regex]::Escape($Name) + "\s+REG_") } | Select-Object -Last 1)
    if ($line.Count -ne 1) { return $null }
    return (([string]$line[0]) -replace '^.*?REG_\w+\s+', '')
}

function Get-HostStatusForState {
    param([string]$RepositoryRoot, [string]$StateRoot)
    $control = Join-Path $RepositoryRoot "common\host-control.ps1"
    $output = @(& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $control -Status -StateRoot $StateRoot 2>&1)
    if ($LASTEXITCODE -ne 0) { return $null }
    $line = @($output | Where-Object { ([string]$_).Trim().StartsWith("{") } | Select-Object -Last 1)
    if ($line.Count -ne 1) { return $null }
    try { return ConvertFrom-Json ([string]$line[0]) } catch { return $null }
}

function Get-ActiveState {
    param([string]$LocalRoot)
    $active = Join-Path $LocalRoot "active.txt"
    if (-not (Test-Path -LiteralPath $active -PathType Leaf)) { return $null }
    $generation = ([IO.File]::ReadAllText($active, (New-Object Text.UTF8Encoding($false, $true)))).Trim()
    if ($generation -cnotmatch '^[a-f0-9]{32}$') { return $null }
    return Join-Path $LocalRoot ("state\" + $generation)
}

function Wait-Ready {
    param([string]$RepositoryRoot, [string]$StateRoot, [int]$Seconds = 8)
    $deadline = [DateTime]::UtcNow.AddSeconds($Seconds)
    do {
        $status = Get-HostStatusForState $RepositoryRoot $StateRoot
        if ($null -ne $status -and [bool]$status.ready) { return $status }
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $deadline)
    return $null
}

function Copy-Context {
    param($Context, [hashtable]$Overrides)
    $copy = @{}
    foreach ($key in $Context.Keys) { $copy[$key] = $Context[$key] }
    foreach ($key in $Overrides.Keys) { $copy[$key] = $Overrides[$key] }
    return $copy
}

$token = [guid]::NewGuid().ToString("N")
$fx = Join-Path $env:TEMP ("toolrack install '" + [char]0x65E5 + "' " + $token)
$fixtureRoot = Join-Path $fx "repository root"
$localRoot = Join-Path $fx "local app data\ToolRack"
$registryParent = "HKCU\Software\ToolRackTests\" + $token
$registryParentFull = "HKEY_CURRENT_USER\Software\ToolRackTests\" + $token
$runPath = $registryParentFull + "\Run"
$runPathShort = $registryParent + "\Run"
$runValueName = "ToolRackHost-" + $token.Substring(0, 8)
$namespace = "install-" + $token
$hostStatus = $null

try {
    New-Item -ItemType Directory -Force -Path $fixtureRoot | Out-Null
    Copy-Item -LiteralPath (Join-Path $root "common") -Destination (Join-Path $fixtureRoot "common") -Recurse
    $toolDir = Join-Path $fixtureRoot "tool\marker"
    New-Item -ItemType Directory -Force -Path $toolDir | Out-Null
    Write-Utf8Text (Join-Path $toolDir "main.ps1") "param([AllowEmptyString()][string]`$Target='')`nexit 0`n"
    Write-Utf8Json (Join-Path $toolDir "tool.json") ([ordered]@{
        schema = 1
        id = "marker"
        name = "Marker"
        on = @("background")
        run = [ordered]@{ type = "powershell"; entry = "main.ps1"; window = "hidden" }
    })
    Write-Utf8Json (Join-Path $fixtureRoot "menu.json") ([ordered]@{
        schema = 1
        default_category = "general"
        categories = @([ordered]@{ id = "general"; label = "Fixture Tools"; tools = @("marker") })
    })
    Write-Utf8Json (Join-Path $fixtureRoot "bindings.json") ([ordered]@{
        schema = 1
        bindings = @([ordered]@{
            id = "marker-hotkey"
            trigger = [ordered]@{ type = "hotkey"; key = "F16"; modifiers = @("ctrl", "alt", "shift") }
            invoke = [ordered]@{ tool = "marker"; action = "default" }
        })
    })
    New-Item -ItemType Directory -Force -Path (Join-Path $fixtureRoot "output") | Out-Null
    Write-Utf8Text (Join-Path $fixtureRoot "output\keep.txt") "keep"

    $menuContexts = @(
        @{ on = "file"; base = $registryParentFull + "\file\shell" },
        @{ on = "folder"; base = $registryParentFull + "\folder\shell" },
        @{ on = "background"; base = $registryParentFull + "\background\shell" }
    )
    $context = New-InstallContext -Root ($fixtureRoot + "\") -LocalRoot $localRoot -HostNamespace $namespace `
        -RunRegistryPath $runPath -RunValueName $runValueName -MenuContexts $menuContexts

    & reg.exe add $runPathShort /v $runValueName /t REG_SZ /d "old-run-command" /f | Out-Null
    $oldMenu = $registryParent + "\background\shell\ToolRack"
    & reg.exe add $oldMenu /v MUIVerb /t REG_SZ /d "Old Fixture Menu" /f | Out-Null

    $badBindings = Join-Path $fx "bad-bindings.json"
    Write-Utf8Text $badBindings "{"
    $preflight = Invoke-Install -InstallContext $context -BindingPath $badBindings
    Assert-True (-not [bool]$preflight.Ok) "invalid preflight fails"
    Assert-True ((Get-RegistryDefault $runPathShort $runValueName) -eq "old-run-command") "preflight failure preserves Run value"
    & reg.exe query $oldMenu /v MUIVerb | Out-Null
    Assert-True ($LASTEXITCODE -eq 0 -and -not (Test-Path -LiteralPath $localRoot)) "preflight failure preserves menu and local state"

    $selfTestContext = Copy-Context $context @{ TestFailSelfTest = $true }
    $selfTestFailure = Invoke-Install -InstallContext $selfTestContext
    Assert-True (-not [bool]$selfTestFailure.Ok) "Add-Type SelfTest seam fails install"
    Assert-True ((Get-RegistryDefault $runPathShort $runValueName) -eq "old-run-command" -and -not (Test-Path -LiteralPath $localRoot)) "SelfTest failure has no external mutation"

    $installed = Invoke-Install -InstallContext $context
    Assert-True ([bool]$installed.Ok) "valid transactional install succeeds"
    $activeState = Get-ActiveState $localRoot
    Assert-True ($null -ne $activeState) "install writes an active generation"
    foreach ($relative in @("bootstrap\v1\start-host.vbs", "bootstrap\v1\host-bootstrap.ps1", "active.txt")) {
        Assert-True (Test-Path -LiteralPath (Join-Path $localRoot $relative) -PathType Leaf) ("install writes " + $relative)
    }
    foreach ($relative in @("root.txt", "host.json", "bindings.resolved.json")) {
        Assert-True (Test-Path -LiteralPath (Join-Path $activeState $relative) -PathType Leaf) ("generation writes " + $relative)
    }
    $rootText = ([IO.File]::ReadAllText((Join-Path $activeState "root.txt"), (New-Object Text.UTF8Encoding($false, $true)))).Trim()
    Assert-True ($rootText -eq [IO.Path]::GetFullPath($fixtureRoot)) "root.txt preserves a spaced root and normalizes trailing slash"
    $hostStatus = Wait-Ready $fixtureRoot $activeState
    Assert-True ($null -ne $hostStatus) "installed Host reports ready"
    $runCommand = Get-RegistryDefault $runPathShort $runValueName
    Assert-True ($null -ne $runCommand -and $runCommand.Length -lt 260) "Run command stays below 260 characters"
    Assert-True ($runCommand -notlike ("*" + $fixtureRoot + "*") -and $runCommand -like "*start-host.vbs*") "Run command uses only the local bootstrap"
    $newMenu = $registryParent + "\background\shell\ToolRack\shell\marker\command"
    & reg.exe query $newMenu | Out-Null
    Assert-True ($LASTEXITCODE -eq 0) "valid install writes isolated Explorer menu"

    $firstPid = [int]$hostStatus.pid
    $firstGeneration = Split-Path $activeState -Leaf
    $reinstalled = Invoke-Install -InstallContext $context
    Assert-True ([bool]$reinstalled.Ok) "reinstall succeeds"
    $secondState = Get-ActiveState $localRoot
    $secondStatus = Wait-Ready $fixtureRoot $secondState
    Assert-True ($null -ne $secondStatus -and [int]$secondStatus.pid -ne $firstPid) "reinstall replaces the old Host"
    Assert-True ($null -eq (Get-Process -Id $firstPid -ErrorAction SilentlyContinue)) "old Host process exits after reinstall"
    Assert-True ((Split-Path $secondState -Leaf) -ne $firstGeneration) "reinstall activates a new generation"

    $stateCountBeforeReadyFailure = @(Get-ChildItem -LiteralPath (Join-Path $localRoot "state") -Directory).Count
    $runBeforeReadyFailure = Get-RegistryDefault $runPathShort $runValueName
    $menuBeforeReadyFailure = (Get-RegistryDefault $oldMenu "MUIVerb")
    $readyFailureContext = Copy-Context $context @{ TestFailReady = $true }
    $readyFailure = Invoke-Install -InstallContext $readyFailureContext
    Assert-True (-not [bool]$readyFailure.Ok) "ready timeout seam fails install"
    $stateAfterReadyFailure = Get-ActiveState $localRoot
    $readyRollbackStatus = Wait-Ready $fixtureRoot $stateAfterReadyFailure
    Assert-True ($null -ne $readyRollbackStatus) "ready failure restores the previous Host"
    Assert-True ((Get-RegistryDefault $runPathShort $runValueName) -eq $runBeforeReadyFailure) "ready failure does not change Run"
    Assert-True (@(Get-ChildItem -LiteralPath (Join-Path $localRoot "state") -Directory).Count -eq $stateCountBeforeReadyFailure) "ready failure restores local generations"

    $activeBeforeRegistryFailure = Split-Path $stateAfterReadyFailure -Leaf
    $runBeforeRegistryFailure = Get-RegistryDefault $runPathShort $runValueName
    $registryFailureContext = Copy-Context $context @{ TestFailRegistry = $true }
    $registryFailure = Invoke-Install -InstallContext $registryFailureContext
    Assert-True (-not [bool]$registryFailure.Ok) "registry import failure seam fails install"
    $stateAfterRegistryFailure = Get-ActiveState $localRoot
    $registryRollbackStatus = Wait-Ready $fixtureRoot $stateAfterRegistryFailure
    Assert-True ($null -ne $registryRollbackStatus) "registry failure restores the previous Host"
    Assert-True ((Split-Path $stateAfterRegistryFailure -Leaf) -eq $activeBeforeRegistryFailure) "registry failure restores active.txt"
    Assert-True ((Get-RegistryDefault $runPathShort $runValueName) -eq $runBeforeRegistryFailure) "registry failure restores Run value"
    & reg.exe query $newMenu | Out-Null
    Assert-True ($LASTEXITCODE -eq 0) "registry failure restores menu tree"

    . $uninstallScript
    $ErrorActionPreference = "Continue"
    $uninstalled = Invoke-Uninstall -InstallContext $context
    Assert-True ([bool]$uninstalled.Ok) "transactional uninstall succeeds"
    Assert-True (-not (Test-Path -LiteralPath $localRoot)) "uninstall removes local Host state"
    Assert-True ($null -eq (Get-RegistryDefault $runPathShort $runValueName)) "uninstall removes only its Run value"
    & reg.exe query ($registryParent + "\background\shell\ToolRack") 2>$null | Out-Null
    Assert-True ($LASTEXITCODE -ne 0) "uninstall removes isolated menu roots"
    foreach ($relative in @("bindings.json", "tool\marker\tool.json", "output\keep.txt")) {
        Assert-True (Test-Path -LiteralPath (Join-Path $fixtureRoot $relative) -PathType Leaf) ("uninstall preserves " + $relative)
    }

    $brokenState = Join-Path $localRoot ("state\" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Force -Path $brokenState | Out-Null
    Write-Utf8Text (Join-Path $localRoot "active.txt") (Split-Path $brokenState -Leaf)
    Write-Utf8Json (Join-Path $brokenState "host.json") ([ordered]@{ schema = 1; version = "1"; root = $fixtureRoot; namespace = $namespace })
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $brokenUninstall = Invoke-Uninstall -InstallContext $context
    $timer.Stop()
    Assert-True ([bool]$brokenUninstall.Ok -and $timer.ElapsedMilliseconds -lt 3000) "uninstall does not hang on a stale or broken Host state"
} finally {
    $activeState = Get-ActiveState $localRoot
    if ($null -ne $activeState -and (Test-Path -LiteralPath $activeState)) {
        try {
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $fixtureRoot "common\host-control.ps1") -Shutdown -StateRoot $activeState | Out-Null
        } catch { }
    }
    & reg.exe delete $registryParent /f 2>$null | Out-Null
    Remove-Item -LiteralPath $fx -Recurse -Force -ErrorAction SilentlyContinue
}

Exit-Test
