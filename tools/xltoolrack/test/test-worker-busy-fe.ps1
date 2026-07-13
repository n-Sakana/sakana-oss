. (Join-Path $PSScriptRoot '_harness.ps1')

# Workers must survive an FE that is alive but temporarily unreachable
# (cell edit mode, open dialog, another macro). Reproduced here by holding
# the FE STA busy with Application.Wait so incoming COM calls are rejected
# with the same busy-class errors (0x80010001 / 0x8001010A / 0x800AC472).

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
    $startProbes = QM $bookName 'JobTest_StartProbeWorkers'
    $snapshot = QM $bookName 'JobTest_VersionSnapshot'
    $stateFor = QM $bookName 'JobTest_StateFor'
    $doneFor = QM $bookName 'JobTest_DonePathFor'
    $stop = QM $bookName 'JobTest_StopAll'
    $running = QM $bookName 'JobTest_RunningCount'

    $swStart = [string](Invoke-XlRun $excel $tryStart 'stopwatch')
    Assert-True ($swStart -like 'OK:*') 'stopwatch job started'
    $swId = $swStart.Substring(3)
    $pwId = [string](Invoke-XlRun $excel $startProbes 1)
    $ids = "$swId,$pwId"
    $swDone = [string](Invoke-XlRun $excel $doneFor $swId)
    $pwDone = [string](Invoke-XlRun $excel $doneFor $pwId)

    Start-Sleep -Seconds 6
    $v0 = ([string](Invoke-XlRun $excel $snapshot $ids)) -split ','
    Start-Sleep -Seconds 3
    $v1 = ([string](Invoke-XlRun $excel $snapshot $ids)) -split ','
    Assert-True (([int]$v1[0] -gt [int]$v0[0]) -and ([int]$v1[1] -gt [int]$v0[1])) "baseline versions advance ($($v0 -join ',') -> $($v1 -join ','))"

    # FE alive but unreachable for 22 seconds.
    $null = $excel.Wait([DateTime]::Now.AddSeconds(22))

    # Recovery time: workers sit out their 10s backoff, then push again.
    Start-Sleep -Seconds 15

    Assert-True (-not (Test-Path -LiteralPath $swDone)) 'stopwatch worker survived the busy window (no done flag)'
    Assert-True (-not (Test-Path -LiteralPath $pwDone)) 'probe worker survived the busy window (no done flag)'

    $v2 = ([string](Invoke-XlRun $excel $snapshot $ids)) -split ','
    Start-Sleep -Seconds 4
    $v3 = ([string](Invoke-XlRun $excel $snapshot $ids)) -split ','
    Assert-True ([int]$v3[0] -gt [int]$v2[0]) "stopwatch versions resumed advancing ($($v2[0]) -> $($v3[0]))"
    Assert-True ([int]$v3[1] -gt [int]$v2[1]) "probe versions resumed advancing ($($v2[1]) -> $($v3[1]))"
    Assert-Equal 'running' ([string](Invoke-XlRun $excel $stateFor $swId)) 'stopwatch still reported running'
    Assert-Equal 'running' ([string](Invoke-XlRun $excel $stateFor $pwId)) 'probe still reported running'

    # Clean stop still works after the busy window, and the sweep removes
    # finished jobs (no ghost "running" entries, no leaked done flags).
    $null = (Invoke-XlRun $excel $stop)
    $deadline = (Get-Date).AddSeconds(15)
    $swState = 'x'
    $pwState = 'x'
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 250
        $null = [int](Invoke-XlRun $excel $running)
        $swState = [string](Invoke-XlRun $excel $stateFor $swId)
        $pwState = [string](Invoke-XlRun $excel $stateFor $pwId)
        if ($swState -eq '' -and $pwState -eq '') { break }
    }
    Assert-Equal '' $swState 'stopwatch job swept after stop'
    Assert-Equal '' $pwState 'probe job swept after stop'
    Assert-Equal 0 ([int](Invoke-XlRun $excel $running)) 'no running jobs after stop'
    Assert-True (-not (Test-Path -LiteralPath $swDone)) 'stopwatch done flag cleaned up'
    Assert-True (-not (Test-Path -LiteralPath $pwDone)) 'probe done flag cleaned up'
} catch {
    Write-Host "  expected-red: $($_.Exception.Message)" -ForegroundColor Yellow
    Assert-True $false 'busy-FE survival harness is available'
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
