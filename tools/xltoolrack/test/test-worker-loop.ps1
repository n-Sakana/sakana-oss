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

    $startError = QM $bookName 'JobTest_StartErrorWorker'
    $stateFor = QM $bookName 'JobTest_StateFor'
    $errorFor = QM $bookName 'JobTest_ErrorFor'
    $doneFor = QM $bookName 'JobTest_DonePathFor'
    $logFor = QM $bookName 'JobTest_LogPathFor'
    $running = QM $bookName 'JobTest_RunningCount'
    $errorId = [string](Invoke-XlRun $excel $startError)
    $errorDone = [string](Invoke-XlRun $excel $doneFor $errorId)
    $errorLog = [string](Invoke-XlRun $excel $logFor $errorId)
    $deadline = (Get-Date).AddSeconds(8)
    $errorState = ''
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 250
        $errorState = [string](Invoke-XlRun $excel $stateFor $errorId)
        if ($errorState -eq 'failed' -and (Test-Path -LiteralPath $errorDone)) { break }
    }
    Assert-Equal 'failed' $errorState 'BE exception reached FE terminal state'
    Assert-True (([string](Invoke-XlRun $excel $errorFor $errorId)) -like '5:*probe worker failure*') 'BE error number and description reached FE'
    Assert-True (Test-Path -LiteralPath $errorDone) 'failed BE self-quit created done flag'
    $errorLogText = if (Test-Path -LiteralPath $errorLog) { Get-Content -LiteralPath $errorLog -Raw } else { '' }
    Assert-True ($errorLogText -match 'FAIL Err 5: probe worker failure') 'BE exception was logged'
    Start-Sleep -Seconds 1
    Assert-Equal 0 ([int](Invoke-XlRun $excel $running)) 'failed worker released its active slot'

    $start = QM $bookName 'JobTest_StartProbeWorkers'
    $snapshot = QM $bookName 'JobTest_VersionSnapshot'
    $donePathsMacro = QM $bookName 'JobTest_DonePaths'
    $stop = QM $bookName 'JobTest_StopAll'
    $ids = ([string](Invoke-XlRun $excel $start 3)) -split ','
    $donePaths = ([string](Invoke-XlRun $excel $donePathsMacro ($ids -join ','))) -split '\|'
    $observed = @(@(), @(), @())
    $last = @(0,0,0)
    $deadline = (Get-Date).AddSeconds(30)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 250
        $parts = ([string](Invoke-XlRun $excel $snapshot ($ids -join ','))) -split ','
        if ($parts.Count -ne 3) { continue }
        for ($i = 0; $i -lt 3; $i++) {
            $version = [int]$parts[$i]
            if ($version -gt $last[$i]) {
                $observed[$i] += ,(Get-Date)
                $last[$i] = $version
            }
        }
    }

    for ($i = 0; $i -lt 3; $i++) {
        Assert-True ($observed[$i].Count -ge 20) "job $($ids[$i]) kept an approximately one-second cadence (samples=$($observed[$i].Count))"
        $maxGap = 0.0
        for ($j = 1; $j -lt $observed[$i].Count; $j++) {
            $gap = ($observed[$i][$j] - $observed[$i][$j-1]).TotalSeconds
            if ($gap -gt $maxGap) { $maxGap = $gap }
        }
        Assert-True ($maxGap -lt 5.0) "job $($ids[$i]) had no five-second stall (maxGap=$([math]::Round($maxGap,2))s)"
    }

    $null = (Invoke-XlRun $excel $stop)
    $deadline = (Get-Date).AddSeconds(8)
    while ((Get-Date) -lt $deadline) {
        if (($donePaths | Where-Object { -not (Test-Path -LiteralPath $_) }).Count -eq 0) { break }
        Start-Sleep -Milliseconds 250
    }
    foreach ($path in $donePaths) {
        Assert-True (Test-Path -LiteralPath $path) "stop caused worker self-quit ($path)"
        if (Test-Path -LiteralPath $path) {
            Assert-True ((Get-Content -LiteralPath $path -Raw) -match 'reason=stop_requested') 'done flag records stop_requested'
        }
    }
    Start-Sleep -Seconds 2
    $afterStop = @(([string](Invoke-XlRun $excel $snapshot ($ids -join ','))) -split ',' | ForEach-Object { [int]$_ })
    Start-Sleep -Seconds 2
    $afterStop2 = @(([string](Invoke-XlRun $excel $snapshot ($ids -join ','))) -split ',' | ForEach-Object { [int]$_ })
    Assert-True (($afterStop2 -join ',') -eq ($afterStop -join ',')) 'versions stopped after stop request'
} catch {
    Write-Host "  expected-red: $($_.Exception.Message)" -ForegroundColor Yellow
    Assert-True $false 'worker failure and cadence diagnostics are available'
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
