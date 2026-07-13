param(
    [switch]$Uninstall,
    [switch]$NoLaunch
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

function Release-ComObject {
    param($Object)
    if ($null -eq $Object) { return }
    try {
        if ([Runtime.InteropServices.Marshal]::IsComObject($Object)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
        }
    } catch {}
}

function Same-Path {
    param([string]$Left, [string]$Right)
    if ([string]::IsNullOrWhiteSpace($Left) -or [string]::IsNullOrWhiteSpace($Right)) { return $false }
    return [IO.Path]::GetFullPath($Left).TrimEnd('\') -ieq [IO.Path]::GetFullPath($Right).TrimEnd('\')
}

$root = Split-Path $PSScriptRoot -Parent
$addinPath = Join-Path $root 'dist\xltoolrack.xlam'
$workerPath = Join-Path $root 'dist\xltoolrack-worker.xlsm'
$buildScript = Join-Path $root 'scripts\Build-Addin.ps1'

if (-not $Uninstall -and ((-not (Test-Path -LiteralPath $addinPath)) -or (-not (Test-Path -LiteralPath $workerPath)))) {
    Write-Host 'Building xltoolrack...' -ForegroundColor Cyan
    & $buildScript -OutputFormat all
}

if (-not $Uninstall) {
    if (-not (Test-Path -LiteralPath $addinPath)) { throw "Add-in was not built: $addinPath" }
    if (-not (Test-Path -LiteralPath $workerPath)) { throw "Worker was not built: $workerPath" }
}

$excel = $null
$bootstrapWorkbook = $null
$addIns = $null
$selected = $null
$excelExe = $null
$launchWorkbookPath = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excelExe = Join-Path ([string]$excel.Path) 'EXCEL.EXE'
    # AddIns.Add fails in a zero-workbook automation instance on Excel 16.
    $bootstrapWorkbook = $excel.Workbooks.Add()
    if (-not $Uninstall -and -not $NoLaunch) {
        $launchWorkbookPath = Join-Path $env:TEMP 'xltoolrack-start.xlsx'
        try {
            if (Test-Path -LiteralPath $launchWorkbookPath) {
                Remove-Item -LiteralPath $launchWorkbookPath -Force
            }
        } catch {
            $launchWorkbookPath = Join-Path $env:TEMP ('xltoolrack-start-' + $PID + '.xlsx')
        }
        $bootstrapWorkbook.SaveAs($launchWorkbookPath, 51)
    }
    $addIns = $excel.AddIns

    for ($index = 1; $index -le $addIns.Count; $index++) {
        $candidate = $null
        try {
            $candidate = $addIns.Item($index)
            if (Same-Path ([string]$candidate.FullName) $addinPath) {
                $selected = $candidate
                $candidate = $null
                break
            }
        } finally {
            Release-ComObject $candidate
        }
    }

    if ($Uninstall) {
        if ($null -ne $selected) { $selected.Installed = $false }
    } else {
        if ($null -eq $selected) { $selected = $addIns.Add($addinPath, $false) }
        $selected.Installed = $true
        if (-not [bool]$selected.Installed) { throw 'Excel did not enable the xltoolrack add-in.' }
    }
} finally {
    Release-ComObject $selected
    Release-ComObject $addIns
    if ($null -ne $bootstrapWorkbook) {
        try { $bootstrapWorkbook.Close($false) } catch {}
        Release-ComObject $bootstrapWorkbook
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        # Excel.Quit has already ended the COM server. FinalReleaseComObject on
        # the Application RCW can block indefinitely on some Office builds.
        $excel = $null
    }
}

if ($Uninstall) {
    Write-Host 'xltoolrack was removed from Excel Add-ins.' -ForegroundColor Green
    exit 0
}

Write-Host 'xltoolrack is installed. Use its Ribbon tab to launch the three tools.' -ForegroundColor Green
if (-not $NoLaunch) {
    if (-not (Test-Path -LiteralPath $excelExe)) { throw "Excel executable was not found: $excelExe" }
    Start-Process -FilePath $excelExe -ArgumentList @('/x', ('"' + $launchWorkbookPath + '"'))
}
