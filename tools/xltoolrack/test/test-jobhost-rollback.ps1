. (Join-Path $PSScriptRoot '_harness.ps1')

$dir = New-TestDirectory
$excel = $null
$workbook = $null
$target = $null
$oldTestMode = $env:XLTOOLRACK_TEST
try {
    $env:XLTOOLRACK_TEST = '1'
    $addin = Build-Addin $dir -Format all
    $opened = Open-TestAddin $addin
    $excel = $opened.Excel
    $workbook = $opened.Workbook
    $target = $excel.Workbooks.Add()
    $target.Activate()
    $bookName = [string]$workbook.Name

    $setFail = QM $bookName 'JobTest_SetFailStage'
    $tryStart = QM $bookName 'JobTest_TryStart'
    $running = QM $bookName 'JobTest_RunningCount'
    $setLimit = QM $bookName 'JobTest_SetGlobalLimit'
    $run = QM $bookName 'Infra_Run'
    $outcome = QM $bookName 'Infra_LastOutcome'
    $logPath = QM $bookName 'Infra_LastLogPath'

    $null = (Invoke-XlRun $excel $setFail 'after_create')
    $failedStart = [string](Invoke-XlRun $excel $tryStart 'probe_worker')
    Assert-True ($failedStart -like 'ERROR:*injected start failure*') 'failure after BE CreateObject was reported'
    Assert-Equal 0 ([int](Invoke-XlRun $excel $running)) 'failed StartJob released reservation and COM state'
    $null = (Invoke-XlRun $excel $setFail '')

    $null = (Invoke-XlRun $excel $setLimit 2)
    $null = (Invoke-XlRun $excel $run 'probe_batch')
    Assert-True (([string](Invoke-XlRun $excel $outcome)) -like 'ERROR:*') 'batch Run failed when third job exceeded global cap'
    $path = [string](Invoke-XlRun $excel $logPath)
    $content = if (Test-Path -LiteralPath $path) { Get-Content -LiteralPath $path -Raw } else { '' }
    Assert-True ($content -match 'global limit reached') 'batch failure retained the cap cause in the harness log'
    Start-Sleep -Seconds 3
    Assert-Equal 0 ([int](Invoke-XlRun $excel $running)) 'harness rolled back the two jobs started before batch failure'
} catch {
    Write-Host "  expected-red: $($_.Exception.Message)" -ForegroundColor Yellow
    Assert-True $false 'partial-start rollback diagnostics are available'
} finally {
    if ($null -ne $excel -and $null -ne $workbook) {
        try { $null = $excel.Run((QM ([string]$workbook.Name) 'JobTest_StopAll')) } catch {}
    }
    if ($null -ne $target) {
        try { $target.Close($false) } catch {}
        Release-Com $target
    }
    Close-TestExcel $excel $workbook
    $env:XLTOOLRACK_TEST = $oldTestMode
    Remove-TestDirectory $dir
}

Exit-Test

