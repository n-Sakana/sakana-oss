# test/verify-registry.ps1 -- after install.bat: assert all expected menu keys exist
. (Join-Path $PSScriptRoot "_assert.ps1")
$checks = @(
    "HKCU\Software\Classes\Directory\Background\shell\ToolRack\shell\capture\command",
    "HKCU\Software\Classes\Directory\shell\ToolRack\shell\tree\shell\v0\command",
    "HKCU\Software\Classes\Directory\Background\shell\ToolRack\shell\tree",
    "HKCU\Software\Classes\*\shell\ToolRack\shell\timestamp\shell\v1\command",
    "HKCU\Software\Classes\Directory\shell\ToolRack\shell\timestamp",
    "HKCU\Software\Classes\Directory\Background\shell\ToolRack\shell\keysend\shell\v1",
    "HKCU\Software\Classes\Directory\Background\shell\ToolRack\shell\timer\shell\v3\command",
    "HKCU\Software\Classes\Directory\Background\shell\ToolRack.ai\shell\clip\shell\v0\command",
    "HKCU\Software\Classes\*\shell\ToolRack.ai\shell\md-mirror\shell\v0\command",
    "HKCU\Software\Classes\Directory\shell\ToolRack.ai\shell\md-mirror\shell\v2\command",
    "HKCU\Software\Classes\Directory\Background\shell\ToolRack.ai\shell\md-mirror\shell\v3\command",
    "HKCU\Software\Classes\Directory\shell\ToolRack.ai.p2\shell\md-patch\shell\v0\command",
    "HKCU\Software\Classes\Directory\Background\shell\ToolRack.ai\shell\md-patch\shell\v1\command",
    "HKCU\Software\Classes\*\shell\ToolRack.ai\shell\md-extract\shell\v0\command",
    "HKCU\Software\Classes\Directory\shell\ToolRack.ai\shell\md-extract\shell\v2\command",
    "HKCU\Software\Classes\*\shell\ToolRack\shell\vba-devkit\shell\v0\command",
    "HKCU\Software\Classes\Directory\shell\ToolRack\shell\vba-devkit\shell\v2\command",
    "HKCU\Software\Classes\*\shell\ToolRack\shell\vba-devkit\shell\v5\command",
    "HKCU\Software\Classes\Directory\Background\shell\ToolRack\shell\transcribe\shell\v0\command"
)
foreach ($k in $checks) {
    & reg query $k | Out-Null
    Assert-True ($LASTEXITCODE -eq 0) ("registered: " + $k)
}
$cmd = (& reg query "HKCU\Software\Classes\Directory\Background\shell\ToolRack\shell\timer\shell\v0\command" /ve) -join " "
Assert-True ($cmd -like "*wscript.exe*silent.vbs*") "timer v0 routes through silent.vbs"
$cmd2 = (& reg query "HKCU\Software\Classes\Directory\shell\ToolRack\shell\tree\shell\v0\command" /ve) -join " "
Assert-True ($cmd2 -like "*launch.ps1*-Variant 0*") "tree v0 routes through launch.ps1"
$cmd3 = (& reg query "HKCU\Software\Classes\*\shell\ToolRack\shell\timestamp\shell\v0\command" /ve) -join " "
Assert-True ($cmd3 -like "*wscript.exe*silent.vbs*") "timestamp v0 routes through silent.vbs"
$cmd4 = (& reg query "HKCU\Software\Classes\Directory\shell\ToolRack.ai\shell\md-mirror\shell\v2\command" /ve) -join " "
Assert-True ($cmd4 -like "*launch.ps1*-Variant 2*") "md-mirror clipboard restore routes through variant 2"
$cmd5 = (& reg query "HKCU\Software\Classes\Directory\Background\shell\ToolRack.ai\shell\md-patch\shell\v0\command" /ve) -join " "
Assert-True ($cmd5 -like "*launch.ps1*-Variant 0*") "md-patch clipboard apply routes through variant 0"
$cmd6 = (& reg query "HKCU\Software\Classes\*\shell\ToolRack.ai\shell\md-extract\shell\v2\command" /ve) -join " "
Assert-True ($cmd6 -like "*launch.ps1*-Variant 2*") "md-extract Custom routes through variant 2"
$cmd7 = (& reg query "HKCU\Software\Classes\*\shell\ToolRack\shell\vba-devkit\shell\v0\command" /ve) -join " "
Assert-True ($cmd7 -like "*launch.ps1*-Variant 0*") "vba-devkit Analyze routes through variant 0"
$cmd8 = (& reg query "HKCU\Software\Classes\*\shell\ToolRack\shell\vba-devkit\shell\v5\command" /ve) -join " "
Assert-True ($cmd8 -like "*launch.ps1*-Variant 5*") "vba-devkit Unlock routes through variant 5"
$cmd9 = (& reg query "HKCU\Software\Classes\Directory\Background\shell\ToolRack\shell\transcribe\shell\v0\command" /ve) -join " "
Assert-True ($cmd9 -like "*wscript.exe*silent.vbs*transcribe* 0 *") "transcribe Start routes through silent.vbs variant 0"
$captureCommand = (& reg query "HKCU\Software\Classes\Directory\Background\shell\ToolRack\shell\capture\command" /ve) -join " "
Assert-True ($captureCommand -like "*wscript.exe*silent.vbs*capture* -1 *") "capture routes directly through silent.vbs"
$transcribeLabel = (& reg query "HKCU\Software\Classes\Directory\Background\shell\ToolRack\shell\transcribe" /v MUIVerb) -join " "
Assert-True ($transcribeLabel -like "*Transcribe*") "transcribe menu label is English"
$runCommand = (& reg query "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" /v ToolRackHost) -join " "
Assert-True ($LASTEXITCODE -eq 0 -and $runCommand -like "*start-host.vbs*") "ToolRackHost autostart uses the local bootstrap"
Assert-True ($runCommand -notlike "*C:\repos\pub\toolrack*") "ToolRackHost autostart does not embed the repository path"
$hostOutput = @(& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path (Split-Path $PSScriptRoot -Parent) "common\host-control.ps1") -Status 2>&1)
$hostLine = @($hostOutput | Where-Object { ([string]$_).Trim().StartsWith("{") } | Select-Object -Last 1)
$hostStatus = $null
if ($hostLine.Count -eq 1) { try { $hostStatus = ConvertFrom-Json ([string]$hostLine[0]) } catch {} }
Assert-True ($LASTEXITCODE -eq 0 -and $null -ne $hostStatus -and [bool]$hostStatus.ready) "ToolRackHost reports ready"
$aiLabel = (& reg query "HKCU\Software\Classes\Directory\shell\ToolRack.ai" /v MUIVerb) -join " "
Assert-True ($aiLabel -like "*Tool Rack AI*") "AI category has visible label"
$aiPage2Label = (& reg query "HKCU\Software\Classes\Directory\shell\ToolRack.ai.p2" /v MUIVerb) -join " "
Assert-True ($aiPage2Label -like "*Tool Rack AI 2*") "overflow page has visible number"
Exit-Test
