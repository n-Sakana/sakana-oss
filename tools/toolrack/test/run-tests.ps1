# test/run-tests.ps1 -- run every test-*.ps1 under PS 5.1, report summary
$here = $PSScriptRoot
$fail = 0
foreach ($t in (Get-ChildItem -LiteralPath $here -Filter "test-*.ps1" | Sort-Object Name)) {
    Write-Host ""
    Write-Host ("== {0} ==" -f $t.Name) -ForegroundColor Cyan
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $t.FullName
    if ($LASTEXITCODE -ne 0) { $fail++ }
}
Write-Host ""
if ($fail -gt 0) { Write-Host ("{0} test file(s) failed" -f $fail) -ForegroundColor Red; exit 1 }
Write-Host "ALL TEST FILES PASSED" -ForegroundColor Green
