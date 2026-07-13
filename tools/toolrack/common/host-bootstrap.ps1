# common/host-bootstrap.ps1 -- local bootstrap to the repository Host.
[CmdletBinding()]
param([string]$StateDirectory = "")
Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

function Read-StrictUtf8Text {
    param([string]$Path)
    $encoding = New-Object Text.UTF8Encoding($false, $true)
    return [IO.File]::ReadAllText($Path, $encoding)
}

try {
    $localRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\.."))
    $stateRootBase = [IO.Path]::GetFullPath((Join-Path $localRoot "state"))
    if ([string]::IsNullOrWhiteSpace($StateDirectory)) {
        $generation = (Read-StrictUtf8Text (Join-Path $localRoot "active.txt")).Trim()
        if ($generation -cnotmatch '^[a-f0-9]{32}$') { throw "active.txt generation is invalid" }
        $StateDirectory = Join-Path $stateRootBase $generation
    }
    $state = [IO.Path]::GetFullPath($StateDirectory)
    $expectedParent = [IO.Path]::GetFullPath((Split-Path $state -Parent))
    if ($expectedParent.TrimEnd("\") -ine $stateRootBase.TrimEnd("\")) {
        throw "state directory must be directly under the local state folder"
    }
    $leaf = Split-Path $state -Leaf
    if ($leaf -cnotmatch '^[a-f0-9]{32}$') { throw "state generation is invalid" }

    $repositoryRoot = (Read-StrictUtf8Text (Join-Path $state "root.txt")).Trim()
    $repositoryRoot = [IO.Path]::GetFullPath($repositoryRoot)
    $hostScript = Join-Path $repositoryRoot "common\host.ps1"
    if (-not (Test-Path -LiteralPath $hostScript -PathType Leaf)) { throw "repository Host not found: $hostScript" }
    & $hostScript -StateRoot $state
    exit $LASTEXITCODE
} catch {
    Write-Error $_.Exception.Message
    exit 1
}
