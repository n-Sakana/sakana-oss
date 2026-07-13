# common/host-control.ps1 -- status, reload, and shutdown over the user pipe.
[CmdletBinding(DefaultParameterSetName = "Status")]
param(
    [Parameter(Mandatory = $true, ParameterSetName = "Status")][switch]$Status,
    [Parameter(Mandatory = $true, ParameterSetName = "Reload")][switch]$Reload,
    [Parameter(Mandatory = $true, ParameterSetName = "Shutdown")][switch]$Shutdown,
    [string]$StateRoot = "",
    [int]$TimeoutMilliseconds = 2000
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

function Read-StrictUtf8Text {
    param([string]$Path)
    return [IO.File]::ReadAllText($Path, (New-Object Text.UTF8Encoding($false, $true)))
}

function Resolve-StateRoot {
    param([string]$Requested)
    if (-not [string]::IsNullOrWhiteSpace($Requested)) { return [IO.Path]::GetFullPath($Requested) }
    $localRoot = Join-Path $env:LOCALAPPDATA "ToolRack"
    $generation = (Read-StrictUtf8Text (Join-Path $localRoot "active.txt")).Trim()
    if ($generation -cnotmatch '^[a-f0-9]{32}$') { throw "active.txt generation is invalid" }
    return [IO.Path]::GetFullPath((Join-Path $localRoot ("state\" + $generation)))
}

try {
    $state = Resolve-StateRoot $StateRoot
    $hostData = ConvertFrom-Json (Read-StrictUtf8Text (Join-Path $state "host.json"))
    $nameSpace = [string]$hostData.namespace
    if ($nameSpace -notmatch '^[A-Za-z0-9-]{1,80}$') { throw "host namespace is invalid" }
    $sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $pipeName = "ToolRackHost-" + $sid + "-" + $nameSpace
    $command = $PSCmdlet.ParameterSetName.ToLowerInvariant()

    $client = New-Object IO.Pipes.NamedPipeClientStream(".", $pipeName, [IO.Pipes.PipeDirection]::InOut)
    try {
        $client.Connect($TimeoutMilliseconds)
        $writer = New-Object IO.StreamWriter($client, (New-Object Text.UTF8Encoding($false)), 1024, $true)
        $reader = New-Object IO.StreamReader($client, (New-Object Text.UTF8Encoding($false, $true)), $false, 1024, $true)
        try {
            $writer.AutoFlush = $true
            $writer.WriteLine($command)
            $response = $reader.ReadLine()
        } finally {
            $writer.Dispose()
            $reader.Dispose()
        }
    } finally {
        $client.Dispose()
    }
    if ([string]::IsNullOrWhiteSpace($response)) { throw "Host returned an empty response" }
    $parsed = ConvertFrom-Json $response
    Write-Output $response
    if (-not [bool]$parsed.ok) { exit 1 }
    exit 0
} catch {
    Write-Error $_.Exception.Message
    exit 1
}
