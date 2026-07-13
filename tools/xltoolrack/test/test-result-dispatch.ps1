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

    $excel.ScreenUpdating = $true
    $excel.DisplayAlerts = $true
    $excel.Calculation = -4105
    $excel.EnableEvents = $true
    $excel.Interactive = $true
    $excel.Cursor = -4143

    $null = $excel.Run((QM ([string]$workbook.Name) 'Infra_Run'), 'probe_result_state')
    $cell = $target.Worksheets.Item(1).Range('A1')
    try {
        $deadline = (Get-Date).AddSeconds(8)
        while ((Get-Date) -lt $deadline -and [int]$cell.Value2 -lt 1) {
            Start-Sleep -Milliseconds 100
        }
        Assert-True ([int]$cell.Value2 -ge 1) 'asynchronous result callback ran'
    } finally { Release-Com $cell }

    Assert-True ([bool]$excel.ScreenUpdating) 'result callback restored ScreenUpdating'
    Assert-True ([bool]$excel.DisplayAlerts) 'result callback restored DisplayAlerts'
    Assert-Equal -4105 ([int]$excel.Calculation) 'result callback restored Calculation without steady-state toggles'
    Assert-True ([bool]$excel.EnableEvents) 'result callback restored EnableEvents'
    Assert-True ([bool]$excel.Interactive) 'result callback restored Interactive'
    Assert-Equal -4143 ([int]$excel.Cursor) 'result callback restored Excel native cursor mode'
} finally {
    if ($null -ne $excel -and $null -ne $workbook) {
        try { $null = $excel.Run((QM ([string]$workbook.Name) 'JobTest_StopAll')) } catch {}
        Start-Sleep -Seconds 2
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
