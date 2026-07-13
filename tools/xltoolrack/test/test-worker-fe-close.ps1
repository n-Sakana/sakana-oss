. (Join-Path $PSScriptRoot '_harness.ps1')

$dir = New-TestDirectory
$excel = $null
$workbook = $null
$target = $null
$donePath = ''
$workerCopy = ''
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
    $id = ([string](Invoke-XlRun $excel (QM $bookName 'JobTest_StartProbeWorkers') 1))
    $donePath = [string](Invoke-XlRun $excel (QM $bookName 'JobTest_DonePathFor') $id)
    $workerCopy = [string](Invoke-XlRun $excel (QM $bookName 'JobTest_WorkerCopyFor') $id)
    $null = (Invoke-XlRun $excel (QM $bookName 'JobTest_DetachForFeClose'))

    $workbook.Close($false)
    Release-Com $workbook
    $workbook = $null
    $target.Close($false)
    Release-Com $target
    $target = $null
    $excel.Quit()
    Release-Com $excel
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    $deadline = (Get-Date).AddSeconds(10)
    while ((Get-Date) -lt $deadline -and -not (Test-Path -LiteralPath $donePath)) {
        Start-Sleep -Milliseconds 250
    }
    Assert-True (Test-Path -LiteralPath $donePath) 'worker self-quit after FE disappeared'
    if (Test-Path -LiteralPath $donePath) {
        Assert-True ((Get-Content -LiteralPath $donePath -Raw) -match 'reason=fe_unavailable') 'done flag records fe_unavailable'
    }
    Start-Sleep -Seconds 1
    if (-not [string]::IsNullOrWhiteSpace($workerCopy)) {
        Remove-Item -LiteralPath $workerCopy -Force -ErrorAction SilentlyContinue
        Assert-True (-not (Test-Path -LiteralPath $workerCopy)) 'FE-disappearance test removed its released worker copy'
    }
    if (-not [string]::IsNullOrWhiteSpace($donePath)) {
        Remove-Item -LiteralPath $donePath -Force -ErrorAction SilentlyContinue
    }
} catch {
    Write-Host "  expected-red: $($_.Exception.Message)" -ForegroundColor Yellow
    Assert-True $false 'FE-disappearance self-quit path is available'
} finally {
    if ($null -ne $target) {
        try { $target.Close($false) } catch {}
        Release-Com $target
    }
    Close-TestExcel $excel $workbook
    $env:XLTOOLRACK_TEST = $oldTestMode
    Remove-TestDirectory $dir
}

Exit-Test
