. (Join-Path $PSScriptRoot '_harness.ps1')

$dir = New-TestDirectory
$excel = $null
$workbook = $null
$target = $null
$oldLogDir = $env:XLTOOLRACK_LOG_DIR
$oldTestMode = $env:XLTOOLRACK_TEST
try {
    $env:XLTOOLRACK_LOG_DIR = Join-Path $dir 'logs'
    $env:XLTOOLRACK_TEST = '1'
    $addin = Build-Addin $dir
    $opened = Open-TestAddin $addin
    $excel = $opened.Excel
    $workbook = $opened.Workbook
    $target = $excel.Workbooks.Add()
    $target.Activate()

    $excel.ScreenUpdating = $true
    $excel.DisplayAlerts = $true
    $excel.Calculation = -4105
    $excel.StatusBar = 'outer-status'

    $runMacro = QM ([string]$workbook.Name) 'Infra_Run'
    $outcomeMacro = QM ([string]$workbook.Name) 'Infra_LastOutcome'
    $logMacro = QM ([string]$workbook.Name) 'Infra_LastLogPath'

    $null = (Invoke-XlRun $excel $runMacro 'probe_ok')
    Assert-Equal 'OK' ([string](Invoke-XlRun $excel $outcomeMacro)) 'normal tool completed inside harness'
    $okLog = [string](Invoke-XlRun $excel $logMacro)
    Assert-True (Test-Path -LiteralPath $okLog) 'normal run log exists'
    $okText = if (Test-Path -LiteralPath $okLog) { Get-Content -LiteralPath $okLog -Raw } else { '' }
    Assert-True ($okText -match 'probe ok') 'ctx log reached Logger'
    Assert-Equal 'outer-status' ([string]$excel.StatusBar) 'Status restored after normal run'

    $null = (Invoke-XlRun $excel $runMacro 'probe_state')
    Assert-True ([bool]$excel.ScreenUpdating) 'ScreenUpdating restored after tool mutation'
    Assert-True ([bool]$excel.DisplayAlerts) 'DisplayAlerts restored after tool mutation'
    Assert-Equal -4105 ([int]$excel.Calculation) 'Calculation restored after tool mutation'

    $null = (Invoke-XlRun $excel $runMacro 'probe_error')
    Assert-True (([string](Invoke-XlRun $excel $outcomeMacro)) -like 'ERROR:*') 'tool error was caught and exposed'
    $errorLog = [string](Invoke-XlRun $excel $logMacro)
    $errorText = if (Test-Path -LiteralPath $errorLog) { Get-Content -LiteralPath $errorLog -Raw } else { '' }
    Assert-True ($errorText -match 'FAIL Err 5: probe failure') 'tool error was logged'
    Assert-True ([bool]$excel.ScreenUpdating) 'ScreenUpdating restored after error'
    Assert-True ([bool]$excel.DisplayAlerts) 'DisplayAlerts restored after error'
    Assert-Equal -4105 ([int]$excel.Calculation) 'Calculation restored after error'
    Assert-Equal 'outer-status' ([string]$excel.StatusBar) 'Status restored after error'
} catch {
    Write-Host "  expected-red: $($_.Exception.Message)" -ForegroundColor Yellow
    Assert-True $false 'mandatory harness macros are available'
} finally {
    if ($null -ne $excel) { try { $excel.StatusBar = $false } catch {} }
    if ($null -ne $target) {
        try { $target.Close($false) } catch {}
        Release-Com $target
    }
    Close-TestExcel $excel $workbook
    $env:XLTOOLRACK_LOG_DIR = $oldLogDir
    $env:XLTOOLRACK_TEST = $oldTestMode
    Remove-TestDirectory $dir
}

Exit-Test
