# test/test-ui.ps1 -- ui.ps1 log/output helpers land in toolrack\log and \output
. (Join-Path $PSScriptRoot "_assert.ps1")
$root = Split-Path $PSScriptRoot -Parent
$ui = Join-Path $root "common\ui.ps1"
Assert-True (Test-Path -LiteralPath $ui -PathType Leaf) "ui.ps1 exists"
if (-not (Test-Path -LiteralPath $ui -PathType Leaf)) { Exit-Test }
. $ui
$log = Start-Log -Tool "uitest"
Assert-True ($log.StartsWith((Join-Path $root "log"))) "log file under toolrack\log"
Assert-True (Test-Path -LiteralPath $log) "log file created"
Write-Log "hello"
Assert-True ((Get-Content -LiteralPath $log -Raw) -like "*hello*") "Write-Log appends"
$out = Get-OutputPath -Tool "uitest" -Label 'a/b:c' -Ext "md"
Assert-True ($out.StartsWith((Join-Path $root "output"))) "output path under toolrack\output"
Assert-True ((Split-Path $out -Leaf) -like "uitest_a_b_c_*.md") "label sanitized into filename"
Assert-True ($null -ne (Get-Command Open-Folder -ErrorAction SilentlyContinue)) "Open-Folder helper exists"
Assert-True ($null -ne (Get-Command Show-ErrorDialog -ErrorAction SilentlyContinue)) "Show-ErrorDialog helper exists"
Remove-Item -LiteralPath $log -Force
Exit-Test
