. (Join-Path $PSScriptRoot '_harness.ps1')

$dir = New-TestDirectory
$excel = $null
$workbook = $null
$target = $null
$oldTestMode = $env:XLTOOLRACK_TEST
try {
    $env:XLTOOLRACK_TEST = '1'
    $addin = Build-Addin $dir -Format all
    $worker = Join-Path $dir 'xltoolrack-worker.xlsm'
    Assert-True (Test-Path -LiteralPath $worker) 'worker workbook was built'
    if (-not (Test-Path -LiteralPath $worker)) { throw 'worker workbook missing' }

    $opened = Open-TestAddin $addin
    $excel = $opened.Excel
    $workbook = $opened.Workbook
    $target = $excel.Workbooks.Add()
    $target.Activate()
    $bookName = [string]$workbook.Name
    $start = QM $bookName 'JobTest_StartProbeWorkers'
    $snapshot = QM $bookName 'JobTest_VersionSnapshot'
    $stop = QM $bookName 'JobTest_StopAll'

    $ids = ([string](Invoke-XlRun $excel $start 3)) -split ','
    Assert-Equal 3 $ids.Count 'three independent job ids returned'
    $deadline = (Get-Date).AddSeconds(9)
    $versions = @(0,0,0)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 250
        $parts = ([string](Invoke-XlRun $excel $snapshot ($ids -join ','))) -split ','
        if ($parts.Count -eq 3) {
            $versions = @($parts | ForEach-Object { [int]$_ })
            if (($versions | Where-Object { $_ -lt 4 }).Count -eq 0) { break }
        }
    }
    for ($i = 0; $i -lt 3; $i++) {
        Assert-True ($versions[$i] -ge 4) "job $($ids[$i]) advanced independently (version=$($versions[$i]))"
    }
    $null = (Invoke-XlRun $excel $stop)
} catch {
    Write-Host "  expected-red: $($_.Exception.Message)" -ForegroundColor Yellow
    Assert-True $false 'N-concurrent JobHost path is available'
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

