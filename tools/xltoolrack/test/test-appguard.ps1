. (Join-Path $PSScriptRoot '_harness.ps1')

$dir = New-TestDirectory
$excel = $null
$workbook = $null
$target = $null
try {
    $addin = Build-Addin $dir
    $opened = Open-TestAddin $addin
    $excel = $opened.Excel
    $workbook = $opened.Workbook
    $target = $excel.Workbooks.Add()
    $excel.ScreenUpdating = $true
    $excel.DisplayAlerts = $true
    $excel.Calculation = -4105

    $result = $excel.Run((QM ([string]$workbook.Name) 'AppGuard_SelfTest'))
    Assert-Equal 'OK' ([string]$result) 'normal, nested, and error scopes restore saved values'
    Assert-True ([bool]$excel.ScreenUpdating) 'ScreenUpdating restored at FE level'
    Assert-True ([bool]$excel.DisplayAlerts) 'DisplayAlerts restored at FE level'
    Assert-Equal -4105 ([int]$excel.Calculation) 'Calculation restored at FE level'

    # Scheduling and ticks must leave Excel's contextual cursor untouched.
    # Pump_Tick yields the UI message queue but never assigns Cursor itself.
    $excel.Cursor = 3 # xlIBeam
    $pumpStart = QM ([string]$workbook.Name) 'JobPump.Pump_Start'
    $pumpStop = QM ([string]$workbook.Name) 'JobPump.Pump_Stop'
    $null = Invoke-XlRun $excel $pumpStart
    $null = Invoke-XlRun $excel $pumpStart
    Assert-Equal 3 ([int]$excel.Cursor) 'pump start preserves the native cursor mode'
    $null = Invoke-XlRun $excel $pumpStop
    Assert-Equal 3 ([int]$excel.Cursor) 'pump stop leaves the native cursor mode unchanged'

    $null = Invoke-XlRun $excel $pumpStart
    Start-Sleep -Seconds 2
    Assert-Equal 3 ([int]$excel.Cursor) 'scheduled idle tick leaves the native cursor mode untouched'
    $excel.Cursor = -4143 # xlDefault
} catch {
    Write-Host "  expected-red: $($_.Exception.Message)" -ForegroundColor Yellow
    Assert-True $false 'AppGuard self-test macro is available'
} finally {
    if ($null -ne $target) {
        try { $target.Close($false) } catch {}
        Release-Com $target
    }
    Close-TestExcel $excel $workbook
    Remove-TestDirectory $dir
}

Exit-Test
