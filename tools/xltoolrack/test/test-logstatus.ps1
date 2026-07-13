. (Join-Path $PSScriptRoot '_harness.ps1')

$dir = New-TestDirectory
$excel = $null
$workbook = $null
$target = $null
$oldLogDir = $env:XLTOOLRACK_LOG_DIR
try {
    $env:XLTOOLRACK_LOG_DIR = Join-Path $dir 'logs'
    $addin = Build-Addin $dir
    $opened = Open-TestAddin $addin
    $excel = $opened.Excel
    $workbook = $opened.Workbook
    $target = $excel.Workbooks.Add()
    $excel.StatusBar = 'outer-status'

    $result = [string]$excel.Run((QM ([string]$workbook.Name) 'LogStatus_SelfTest'))
    $parts = $result -split '\|', 3
    Assert-Equal 3 $parts.Count 'self-test returned log/status fields'
    if ($parts.Count -eq 3) {
        $logPath = $parts[0]
        Assert-True (Test-Path -LiteralPath $logPath) 'Logger created a run file'
        $content = if (Test-Path -LiteralPath $logPath) { Get-Content -LiteralPath $logPath -Raw } else { '' }
        Assert-True ($content -match 'BEGIN tool=selftest') 'Logger recorded begin'
        Assert-True ($content -match 'hello from selftest') 'Logger recorded line'
        Assert-True ($content -match 'DONE') 'Logger recorded completion'
        Assert-Equal 'probe-status' $parts[1] 'Status.Show updated Application.StatusBar'
        Assert-Equal 'outer-status' $parts[2] 'Status.Clear restored previous StatusBar'
    }
} catch {
    Write-Host "  expected-red: $($_.Exception.Message)" -ForegroundColor Yellow
    Assert-True $false 'Logger and Status self-test macro is available'
} finally {
    if ($null -ne $excel) { try { $excel.StatusBar = $false } catch {} }
    if ($null -ne $target) {
        try { $target.Close($false) } catch {}
        Release-Com $target
    }
    Close-TestExcel $excel $workbook
    $env:XLTOOLRACK_LOG_DIR = $oldLogDir
    Remove-TestDirectory $dir
}

Exit-Test
