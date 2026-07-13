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
    $tryStart = QM $bookName 'JobTest_TryStart'
    $setLimit = QM $bookName 'JobTest_SetGlobalLimit'
    $count = QM $bookName 'JobTest_RunningCount'
    $stop = QM $bookName 'JobTest_StopAll'

    $null = (Invoke-XlRun $excel $setLimit 10)
    Assert-True (([string](Invoke-XlRun $excel $tryStart 'probe_capped')) -like 'OK:*') 'first per-tool job accepted'
    Assert-True (([string](Invoke-XlRun $excel $tryStart 'probe_capped')) -like 'OK:*') 'second per-tool job accepted'
    $third = [string](Invoke-XlRun $excel $tryStart 'probe_capped')
    Assert-True ($third -like 'ERROR:*tool limit*') 'third per-tool job rejected without popup'
    Assert-Equal 2 ([int](Invoke-XlRun $excel $count)) 'per-tool rejection did not reserve a partial job'
    $null = (Invoke-XlRun $excel $stop)
    Start-Sleep -Seconds 2

    $null = (Invoke-XlRun $excel $setLimit 2)
    Assert-True (([string](Invoke-XlRun $excel $tryStart 'probe_worker')) -like 'OK:*') 'first global job accepted'
    Assert-True (([string](Invoke-XlRun $excel $tryStart 'probe_worker')) -like 'OK:*') 'second global job accepted'
    $globalThird = [string](Invoke-XlRun $excel $tryStart 'probe_worker')
    Assert-True ($globalThird -like 'ERROR:*global limit*') 'job above global limit rejected'
    Assert-Equal 2 ([int](Invoke-XlRun $excel $count)) 'global rejection did not reserve a partial job'
    $null = (Invoke-XlRun $excel $stop)
} catch {
    Write-Host "  expected-red: $($_.Exception.Message)" -ForegroundColor Yellow
    Assert-True $false 'resource-cap JobHost path is available'
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
