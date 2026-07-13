# common/resolve-host-config.ps1 -- raw manifests/bindings to normalized Host JSON.
param(
    [string]$Root = (Split-Path $PSScriptRoot -Parent),
    [string]$BindingsPath,
    [Parameter(Mandatory = $true)][string]$OutputPath
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$requestedRoot = [IO.Path]::GetFullPath($Root)
if ([string]::IsNullOrWhiteSpace($BindingsPath)) {
    $BindingsPath = Join-Path $requestedRoot "bindings.json"
}
$requestedBindingsPath = [IO.Path]::GetFullPath($BindingsPath)

. (Join-Path $PSScriptRoot "install.ps1")
. (Join-Path $PSScriptRoot "bindings.ps1")

function Get-SourceSha256 {
    param([string]$Path)
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $hash = $sha.ComputeHash([IO.File]::ReadAllBytes($Path))
        return ([BitConverter]::ToString($hash)).Replace("-", "").ToLowerInvariant()
    } finally {
        $sha.Dispose()
    }
}

function Write-AtomicUtf8Json {
    param([string]$Path, $Value)
    $fullPath = [IO.Path]::GetFullPath($Path)
    $parent = Split-Path $fullPath -Parent
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        New-Item -ItemType Directory -Force -Path $parent | Out-Null
    }
    $temporary = Join-Path $parent ((Split-Path $fullPath -Leaf) + ".tmp." + [guid]::NewGuid().ToString("N"))
    try {
        $json = ConvertTo-Json $Value -Depth 10
        [IO.File]::WriteAllText($temporary, $json, (New-Object Text.UTF8Encoding($false)))
        if (Test-Path -LiteralPath $fullPath -PathType Leaf) {
            [IO.File]::Replace($temporary, $fullPath, $null)
        } else {
            [IO.File]::Move($temporary, $fullPath)
        }
    } finally {
        Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue
    }
}

$lastFailure = $null
foreach ($retryDelay in @(0, 200, 400, 800)) {
    if ($retryDelay -gt 0) { Start-Sleep -Milliseconds $retryDelay }
    try {
        $config = Read-BindingsConfig -Path $requestedBindingsPath
        if (-not $config.Ok) { throw ($config.Errors -join "`n") }

        $tools = New-Object "System.Collections.Generic.List[object]"
        $toolRoot = Join-Path $requestedRoot "tool"
        if (Test-Path -LiteralPath $toolRoot -PathType Container) {
            foreach ($directory in @(Get-ChildItem -LiteralPath $toolRoot -Directory | Sort-Object Name)) {
                $read = Read-Manifest $directory.FullName
                if (-not $read.Ok) { continue }
                $manifestErrors = @(Test-Manifest $read.Data $directory.Name $directory.FullName)
                if ($manifestErrors.Count -gt 0) { continue }
                [void]$tools.Add(@{
                    Id = [string](Get-Prop $read.Data "id")
                    SupportsActions = [bool]$read.SupportsActions
                    ActionIds = @($read.ActionIds)
                    Dir = $directory.FullName
                })
            }
        }
        $resolved = Resolve-Bindings -Config $config.Data -Tools $tools.ToArray()
        if (-not $resolved.Ok) { throw ($resolved.Errors -join "`n") }

        $activeOutput = @($resolved.Active | ForEach-Object {
            $trigger = [ordered]@{ type = [string]$_.Trigger.Type; modifiers = @($_.Trigger.Modifiers) }
            if ($_.Trigger.Type -eq "hotkey") { $trigger["key"] = [string]$_.Trigger.Key }
            else { $trigger["button"] = [string]$_.Trigger.Button }
            [ordered]@{
                id = [string]$_.Id
                trigger = $trigger
                invoke = [ordered]@{ tool = [string]$_.Invoke.Tool; action = [string]$_.Invoke.Action }
                toolDir = [string]$_.ToolDir
            }
        })
        $rejectedOutput = @($resolved.Rejected | ForEach-Object {
            [ordered]@{ id = [string]$_.Id; reason = [string]$_.Reason }
        })
        $output = [ordered]@{
            schema = 1
            root = $requestedRoot
            sourceConfigPath = $requestedBindingsPath
            sourceConfigSha256 = Get-SourceSha256 $requestedBindingsPath
            active = $activeOutput
            rejected = $rejectedOutput
        }
        Write-AtomicUtf8Json -Path $OutputPath -Value $output
        exit 0
    } catch {
        $lastFailure = $_.Exception
    }
}
Write-Error $lastFailure.Message
exit 1
