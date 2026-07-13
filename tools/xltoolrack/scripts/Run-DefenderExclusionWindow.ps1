<#
.SYNOPSIS
  Opens one time-bounded Defender exclusion window for the phase-2 experiment.

.DESCRIPTION
  This script must run elevated. It accepts no exclusion-path override: the
  only path it can add or remove is %TEMP%\xltoolrack. It records Defender's
  exclusion list before, during, and after the experiment. The exclusion is
  removed in finally when the caller creates measurement.done or when the
  timeout expires.

  The non-elevated coordinator waits for exclusion.ready, performs the
  measurement, and then creates measurement.done. Always retain the files in
  CoordinationDirectory as experiment evidence.
#>
[CmdletBinding()]
param(
    [string]$CoordinationDirectory = (Join-Path $env:TEMP 'xltoolrack-defender-coordination'),

    [ValidateRange(1, 60)]
    [int]$TimeoutMinutes = 30
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [Security.Principal.WindowsPrincipal]$identity
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Run-DefenderExclusionWindow.ps1 must run as administrator.'
}

$expectedPath = [IO.Path]::GetFullPath((Join-Path $env:TEMP 'xltoolrack')).TrimEnd('\')
if (-not (Test-Path -LiteralPath $expectedPath -PathType Container)) {
    New-Item -ItemType Directory -Path $expectedPath -Force | Out-Null
}
$actualPath = [IO.Path]::GetFullPath($expectedPath).TrimEnd('\')
if ($actualPath -ine $expectedPath -or
    [IO.Path]::GetFileName($actualPath) -ine 'xltoolrack' -or
    [IO.Path]::GetDirectoryName($actualPath).TrimEnd('\') -ine
        [IO.Path]::GetFullPath($env:TEMP).TrimEnd('\')) {
    throw "Refusing unexpected Defender exclusion path: $actualPath"
}

if (-not (Test-Path -LiteralPath $CoordinationDirectory)) {
    New-Item -ItemType Directory -Path $CoordinationDirectory -Force | Out-Null
}
$CoordinationDirectory = [IO.Path]::GetFullPath($CoordinationDirectory)
$readyPath = Join-Path $CoordinationDirectory 'exclusion.ready'
$donePath = Join-Path $CoordinationDirectory 'measurement.done'
$removedPath = Join-Path $CoordinationDirectory 'exclusion.removed'
$failurePath = Join-Path $CoordinationDirectory 'exclusion.failure.txt'

Remove-Item -LiteralPath $readyPath,$donePath,$removedPath,$failurePath -Force -ErrorAction SilentlyContinue

function Write-PreferenceSnapshot {
    param([Parameter(Mandatory)][string]$Name)
    $preference = Get-MpPreference
    [pscustomobject]@{
        CapturedAt = (Get-Date).ToString('o')
        ExclusionPath = @($preference.ExclusionPath)
    } | ConvertTo-Json -Depth 4 |
        Out-File -LiteralPath (Join-Path $CoordinationDirectory "$Name.json") -Encoding UTF8
}

function Test-ExactExclusion {
    $paths = @((Get-MpPreference).ExclusionPath)
    return @($paths | Where-Object {
        -not [string]::IsNullOrWhiteSpace([string]$_) -and
        [IO.Path]::GetFullPath([string]$_).TrimEnd('\') -ieq $actualPath
    }).Count -gt 0
}

$added = $false
try {
    Write-PreferenceSnapshot 'before'
    if (Test-ExactExclusion) {
        throw "$actualPath was already excluded; a clean before/after experiment is not possible."
    }

    Add-MpPreference -ExclusionPath $actualPath
    $added = $true
    if (-not (Test-ExactExclusion)) {
        throw "Defender did not report the requested exclusion: $actualPath"
    }
    Write-PreferenceSnapshot 'excluded'
    "path=$actualPath`nready_at=$((Get-Date).ToString('o'))" |
        Out-File -LiteralPath $readyPath -Encoding ASCII

    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    while (-not (Test-Path -LiteralPath $donePath) -and (Get-Date) -lt $deadline) {
        Start-Sleep -Seconds 1
    }
    if (-not (Test-Path -LiteralPath $donePath)) {
        throw "Timed out after $TimeoutMinutes minutes waiting for measurement.done."
    }
} catch {
    $_ | Out-String | Out-File -LiteralPath $failurePath -Encoding UTF8
    throw
} finally {
    if ($added) {
        Remove-MpPreference -ExclusionPath $actualPath -ErrorAction Continue
    }
    Write-PreferenceSnapshot 'final'
    if (Test-ExactExclusion) {
        "Removal verification failed for $actualPath" |
            Out-File -LiteralPath $failurePath -Append -Encoding UTF8
    } else {
        "path=$actualPath`nremoved_at=$((Get-Date).ToString('o'))" |
            Out-File -LiteralPath $removedPath -Encoding ASCII
    }
}
