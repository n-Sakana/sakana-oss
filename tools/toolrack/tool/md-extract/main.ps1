# MD Extract -- extract text and Office/PDF content into one markdown bundle.
[CmdletBinding()]
param(
    [string]$Target = "$PWD",
    [int]$TimeoutSeconds = 120,
    [int]$RetryTimeoutSeconds = 300,
    [ValidateSet("Ask", "Retry", "Skip")][string]$DeferredAction = "Ask",
    [switch]$Custom,
    [switch]$Worker,
    [string]$WorkerRequestPath = "",
    [switch]$NoOpen
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"
$workerRequestArgument = $WorkerRequestPath
$deferredActionArgument = $DeferredAction
if (-not $Worker) { . (Join-Path $PSScriptRoot "..\..\common\ui.ps1") }
. (Join-Path $PSScriptRoot "engine.ps1")
$script:ExtractEntryPath = $PSCommandPath
$script:WorkerRequestPath = $workerRequestArgument
$script:DeferredAction = $deferredActionArgument
$script:ExitCode = 0

function Invoke-MdExtractMain {
    param([string[]]$InputPaths)
    if ($null -eq $InputPaths -or $InputPaths.Count -eq 0) { throw "No input paths were provided." }
    if (-not (Test-Path -LiteralPath $InputPaths[0])) { throw "Path not found: $($InputPaths[0])" }

    $label = Split-Path $InputPaths[0] -Leaf
    $outputFullPath = Get-OutputPath -Tool "extract-md" -Label $label -Ext "md"
    $logFullPath = Start-Log -Tool "extract-md"
    Initialize-RunLog -Path $logFullPath

    Start-UI -Title "MD Extract"
    Write-Step ("Target: {0}" -f $InputPaths[0])
    Write-RunLog -Text "RUN_START"
    Write-RunLog -Text ("CONFIG timeout={0}s retry_timeout={1}s deferred_action={2}" -f $TimeoutSeconds, $RetryTimeoutSeconds, $script:DeferredAction)
    Write-RunLog -Text ("OUTPUT_PATH {0}" -f $outputFullPath)
    Write-RunLog -Text ("LOG_PATH {0}" -f $logFullPath)
    Write-RunLogBlock -Title "INPUT_PATHS" -Lines $InputPaths

    Write-RunLog -Text "SCAN_START"
    $fileTree = Build-FileTree -Paths $InputPaths
    $files = @(Collect-InputFiles -Paths $InputPaths)
    Write-RunLog -Text ("SCAN_END files={0} messages={1}" -f $files.Count, $script:ScanMessages.Count)
    if ($script:ScanMessages.Count -gt 0) { Write-RunLogBlock -Title "SCAN_MESSAGES" -Lines $script:ScanMessages.ToArray() }
    if ($files.Count -eq 0) { throw "No files were found." }
    $fileLines = @($files | ForEach-Object { "{0}`t{1}`t{2}" -f (Get-TypeCode -Path $_.FullName), $_.Length, $_.FullName })
    Write-RunLogBlock -Title "FILE_LIST" -Lines $fileLines

    Write-Step ("Extracting {0} file(s)" -f $files.Count)
    Write-RunLog -Text "EXTRACT_START"
    $results = Invoke-MainExtractionPass -Files $files -TimeoutSeconds $TimeoutSeconds
    Write-RunLog -Text "EXTRACT_END"
    Write-RunLog -Text "WRITE_START"
    Write-BundleOutput -InputPaths $InputPaths -FileTree $fileTree -Results $results.ToArray() -OutputPath $outputFullPath -TimeoutSeconds $TimeoutSeconds -RetryTimeoutSeconds $RetryTimeoutSeconds
    Write-RunLog -Text ("WRITE_END output={0}" -f $outputFullPath)

    $retryChoice = Get-DeferredRetryChoice -Results $results.ToArray()
    if ($retryChoice -eq "Retry") {
        Write-Step "Retrying deferred files"
        Invoke-DeferredRetryPass -Results $results -TimeoutSeconds $RetryTimeoutSeconds
        Write-RunLog -Text "WRITE_RETRY_START"
        Write-BundleOutput -InputPaths $InputPaths -FileTree $fileTree -Results $results.ToArray() -OutputPath $outputFullPath -TimeoutSeconds $TimeoutSeconds -RetryTimeoutSeconds $RetryTimeoutSeconds
        Write-RunLog -Text ("WRITE_RETRY_END output={0}" -f $outputFullPath)
    } elseif (@($results | Where-Object { $_.Status -eq "DEFERRED_TIMEOUT" }).Count -gt 0) {
        Write-RunLog -Text "DEFERRED_RETRY_SKIPPED"
    }

    $counts = Get-ResultCounts -Results $results.ToArray()
    Write-RunLog -Text ("RUN_SUMMARY ok={0} failed={1} deferred={2} unsupported={3} total={4}" -f $counts.OK, $counts.Failed, $counts.Deferred, $counts.Unsupported, $counts.Total)
    Write-RunLog -Text "RUN_END"
    Write-Host ""
    Write-Host ("Summary: OK={0} Failed={1} Deferred={2} Unsupported={3} Total={4}" -f $counts.OK, $counts.Failed, $counts.Deferred, $counts.Unsupported, $counts.Total) -ForegroundColor Gray
    Stop-UI -OutPath $outputFullPath
    if ($counts.Failed -gt 0 -or $counts.Deferred -gt 0 -or $script:ScanMessages.Count -gt 0) { $script:ExitCode = 1 }
    if (-not $NoOpen) { Open-Result $outputFullPath }
}

try {
    if ($Worker) {
        Invoke-WorkerMode
    } else {
        if ($Custom -and $env:TOOLRACK_NOPAUSE -ne "1") {
            Write-Host ""
            Write-Host "MD Extract custom settings" -ForegroundColor Cyan
            $TimeoutSeconds = Read-Int -Prompt "Timeout seconds" -Default $TimeoutSeconds
            $RetryTimeoutSeconds = Read-Int -Prompt "Retry timeout seconds" -Default $RetryTimeoutSeconds
            $DeferredAction = Read-Choice -Prompt "Deferred action" -Options @("Ask", "Retry", "Skip") -DefaultIndex 2
        }
        if ($TimeoutSeconds -lt 1) { throw "TimeoutSeconds must be at least 1." }
        if ($RetryTimeoutSeconds -lt 1) { throw "RetryTimeoutSeconds must be at least 1." }
        if ($env:TOOLRACK_NOPAUSE -eq "1" -and $DeferredAction -eq "Ask") { $DeferredAction = "Skip" }
        $script:DeferredAction = $DeferredAction
        Invoke-MdExtractMain -InputPaths @($Target)
    }
} catch {
    if (-not $Worker) {
        Write-RunLog -Text ("RUN_ERROR message={0}" -f $_.Exception.Message)
        Write-Section -Text "Error"
        Write-FailLine -Text $_.Exception.Message
        if (-not [string]::IsNullOrWhiteSpace($script:RunLogPath)) {
            Write-Host ""
            Write-Host "Log:" -ForegroundColor Gray
            Write-Host $script:RunLogPath -ForegroundColor Cyan
        }
    } else {
        [Console]::Error.WriteLine($_.Exception.ToString())
    }
    $script:ExitCode = 1
} finally {
    Stop-OfficeApplications
}

exit $script:ExitCode
