# test/_assert.ps1 -- tiny assertion helpers (dot-source from each test)
$script:Failed = 0
function Assert-True {
    param([bool]$Cond, [string]$Msg)
    if ($Cond) { Write-Host ("  ok   {0}" -f $Msg) -ForegroundColor Green }
    else { $script:Failed++; Write-Host ("  FAIL {0}" -f $Msg) -ForegroundColor Red }
}
function Assert-Contains {
    param([string[]]$Lines, [string]$Pattern, [string]$Msg)
    $hit = @($Lines | Where-Object { $_ -like $Pattern })
    Assert-True ($hit.Count -gt 0) $Msg
}
function Exit-Test {
    Write-Host ""
    if ($script:Failed -gt 0) { Write-Host ("{0} failure(s)" -f $script:Failed) -ForegroundColor Red; exit 1 }
    Write-Host "all passed" -ForegroundColor Green; exit 0
}
