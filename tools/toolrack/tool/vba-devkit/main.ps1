# VBA DevKit -- toolrack dispatcher for the vba-devkit tool set.
param(
    [string]$Target = "$PWD",
    [ValidateSet("analyze", "extract", "diff", "sanitize", "unlock")][string]$Mode = "analyze",
    [switch]$Settings,
    [string]$Second = ""
)

$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "..\..\common\ui.ps1")

$libDir = Join-Path $PSScriptRoot "lib"

function Invoke-VbaLib {
    param([string]$Script, [hashtable]$Parameters)
    $path = Join-Path $libDir $Script
    try {
        $global:LASTEXITCODE = 0
        $libOutput = @(& $path @Parameters)
        $libExitCode = $global:LASTEXITCODE
        if ($libOutput.Count -gt 0) { $libOutput | Out-Host }
        return $libExitCode
    } catch {
        Write-Err $_.Exception.Message
        return 1
    }
}

function Get-SecondFile {
    param([string]$FirstFile)
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    try {
        $dialog.Title = "VBA Diff: choose the second workbook"
        $dialog.Filter = "Excel with VBA (*.xlsm;*.xlam;*.xls)|*.xlsm;*.xlam;*.xls|All files (*.*)|*.*"
        $dialog.InitialDirectory = Split-Path $FirstFile -Parent
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            return $dialog.FileName
        }
        return ""
    } finally {
        $dialog.Dispose()
    }
}

if ($Mode -eq "analyze" -and $Settings) {
    exit (Invoke-VbaLib "Analyze.ps1" @{})
}

if (-not (Test-Path -LiteralPath $Target)) {
    Write-Err "Path not found: $Target"
    exit 1
}
$isFolder = Test-Path -LiteralPath $Target -PathType Container

switch ($Mode) {
    "analyze" {
        exit (Invoke-VbaLib "Analyze.ps1" @{ Paths = @($Target) })
    }
    "extract" {
        exit (Invoke-VbaLib "Extract.ps1" @{ Paths = @($Target) })
    }
    "sanitize" {
        exit (Invoke-VbaLib "Sanitize.ps1" @{ Path = @($Target) })
    }
    "unlock" {
        if ($isFolder) {
            Write-Err "Unlock needs a single Excel file, not a folder."
            exit 1
        }
        Write-Warn "Use Unlock only for files you are authorized to modify."
        exit (Invoke-VbaLib "Unlock.ps1" @{ FilePath = $Target })
    }
    "diff" {
        if ($isFolder) {
            Write-Err "Diff needs a single Excel file, not a folder."
            exit 1
        }
        $fileB = $Second
        if (-not $fileB) { $fileB = Get-SecondFile $Target }
        if (-not $fileB) {
            Write-Warn "Diff cancelled."
            exit 0
        }
        if (-not (Test-Path -LiteralPath $fileB -PathType Leaf)) {
            Write-Err "Second file not found: $fileB"
            exit 1
        }
        exit (Invoke-VbaLib "Diff.ps1" @{ FileA = $Target; FileB = $fileB })
    }
}
