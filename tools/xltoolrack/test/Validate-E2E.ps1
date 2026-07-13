. (Join-Path $PSScriptRoot '_harness.ps1')

$dir = New-TestDirectory
$excel = $null
$workbook = $null
$target = $null
$oldTestMode = $env:XLTOOLRACK_TEST
$oldLogDir = $env:XLTOOLRACK_LOG_DIR
try {
    $env:XLTOOLRACK_TEST = '1'
    $env:XLTOOLRACK_LOG_DIR = Join-Path $dir 'logs'
    $addin = Build-Addin $dir -Format all
    $opened = Open-TestAddin $addin
    $excel = $opened.Excel
    $workbook = $opened.Workbook
    $target = $excel.Workbooks.Add()
    $target.Activate()
    $bookName = [string]$workbook.Name
    $run = QM $bookName 'Infra_Run'
    $running = QM $bookName 'JobTest_RunningCount'
    $tryStart = QM $bookName 'JobTest_TryStart'
    $stop = QM $bookName 'JobTest_StopAll'

    foreach ($tool in @('stopwatch','pi_race','life')) {
        $target.Activate()
        $null = (Invoke-XlRun $excel $run $tool)
        Assert-Equal 'OK' ([string](Invoke-XlRun $excel (QM $bookName 'Infra_LastOutcome'))) "$tool FE Run completed"
    }
    Assert-Equal 7 ([int](Invoke-XlRun $excel $running)) 'three samples started seven independent BE jobs'

    $stopwatchSheet = $target.Worksheets.Item('xtr_stopwatch')
    $piSheet = $target.Worksheets.Item('xtr_pi_race')
    $lifeSheet = $target.Worksheets.Item('xtr_life')
    try {
        $deadline = (Get-Date).AddSeconds(10)
        $ready = $false
        while ((Get-Date) -lt $deadline) {
            Start-Sleep -Milliseconds 250
            $stopValues = @()
            $piVersions = @()
            for ($i = 3; $i -le 5; $i++) {
                $cell = $stopwatchSheet.Range("B$i")
                try { $stopValues += [int]$cell.Value2 } finally { Release-Com $cell }
                $cell = $piSheet.Range("D$i")
                try { $piVersions += [int]$cell.Value2 } finally { Release-Com $cell }
            }
            $generationCell = $lifeSheet.Range('W3')
            try { $lifeGeneration = [int]$generationCell.Value2 } finally { Release-Com $generationCell }
            $ready = (($stopValues | Where-Object { $_ -lt 3 }).Count -eq 0 -and
                      ($piVersions | Where-Object { $_ -lt 3 }).Count -eq 0 -and
                      $lifeGeneration -ge 2)
            if ($ready) { break }
        }
        Assert-True $ready "all seven jobs advanced concurrently (stop=[$($stopValues -join ',')] pi=[$($piVersions -join ',')] life=$lifeGeneration)"

        $sideCell = $lifeSheet.Range('Y3')
        try {
            $logCountBefore = @(Get-ChildItem -LiteralPath $env:XLTOOLRACK_LOG_DIR -File -ErrorAction SilentlyContinue).Count
            $maxInputMs = 0
            $inputsOk = $true
            for ($probe = 1; $probe -le 30; $probe++) {
                $inputTimer = [Diagnostics.Stopwatch]::StartNew()
                $expectedInput = "responsive-$probe"
                $sideCell.Value2 = $expectedInput
                $actualInput = [string]$sideCell.Value2
                $inputTimer.Stop()
                if ($inputTimer.ElapsedMilliseconds -gt $maxInputMs) { $maxInputMs = $inputTimer.ElapsedMilliseconds }
                if ($actualInput -ne $expectedInput) { $inputsOk = $false }
                Start-Sleep -Milliseconds 100
            }
            $logCountAfter = @(Get-ChildItem -LiteralPath $env:XLTOOLRACK_LOG_DIR -File -ErrorAction SilentlyContinue).Count
            Assert-True $inputsOk '30 consecutive user cell inputs remained available while all samples ran'
            Assert-True ($maxInputMs -lt 750) "FE cell round-trip stayed below 750ms (max=${maxInputMs}ms)"
            Assert-Equal $logCountBefore $logCountAfter 'result ticks did not create per-second log files'
        } finally { Release-Com $sideCell }
    } finally {
        Release-Com $lifeSheet
        Release-Com $piSheet
        Release-Com $stopwatchSheet
    }

    foreach ($tool in @('stopwatch','pi_race','life')) {
        $logs = @(Get-ChildItem -LiteralPath $env:XLTOOLRACK_LOG_DIR -File -Filter "$tool`_*.log" -ErrorAction SilentlyContinue)
        Assert-True ($logs.Count -ge 1) "$tool mandatory harness log exists"
    }

    for ($i = 1; $i -le 3; $i++) {
        Assert-True (([string](Invoke-XlRun $excel $tryStart 'probe_worker')) -like 'OK:*') "extra job $i filled remaining global capacity"
    }
    $rejected = [string](Invoke-XlRun $excel $tryStart 'probe_capped')
    Assert-True ($rejected -like 'ERROR:*global limit*') 'eleventh job was rejected at the FE-global limit of ten'
    Assert-Equal 10 ([int](Invoke-XlRun $excel $running)) 'limit rejection left exactly ten active jobs'

    $null = (Invoke-XlRun $excel $stop)
    Start-Sleep -Seconds 3
    Assert-Equal 0 ([int](Invoke-XlRun $excel $running)) 'StopAll released every sample and probe job'
} catch {
    Write-Host "  FATAL: $($_.Exception.Message)" -ForegroundColor Red
    Assert-True $false 'full platform E2E completed'
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
    $env:XLTOOLRACK_LOG_DIR = $oldLogDir
    Remove-TestDirectory $dir
}

Exit-Test
