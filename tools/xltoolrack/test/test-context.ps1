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
    $target.Activate()
    $result = [string]$excel.Run((QM ([string]$workbook.Name) 'Context_SelfTest'))
    Assert-Equal 'OK' $result 'InfraContext delegates log, status, target, jobs, and rollback'
} catch {
    Write-Host "  expected-red: $($_.Exception.Message)" -ForegroundColor Yellow
    Assert-True $false 'InfraContext self-test macro is available'
} finally {
    if ($null -ne $target) {
        try { $target.Close($false) } catch {}
        Release-Com $target
    }
    Close-TestExcel $excel $workbook
    Remove-TestDirectory $dir
}

Exit-Test
