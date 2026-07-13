param([string]$Target = "")
Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$dll = Join-Path $PSScriptRoot "LatencyPalette.dll"
[void][Reflection.Assembly]::LoadFrom($dll)
exit [ToolRackProbe.LatencyPalette]::Run()
