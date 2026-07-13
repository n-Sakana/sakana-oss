<#
.SYNOPSIS
  Aggregates phase-2 desktop measurement sets into the directive's table.

.DESCRIPTION
  Reads valid attempt directories produced by Run-ControlledMeasurementSet,
  concatenates their raw pump samples and reconstructed WAIT intervals, and
  writes per-configuration CSV/JSON summaries. Invalid attempts are retained
  in the rejection column but never contribute to distributions.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string[]]$SetDirectories,

    [string]$OutputDirectory
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

function Read-SummaryField {
    param([string]$Path, [string]$Name)
    $escaped = [regex]::Escape($Name)
    $line = Get-Content -LiteralPath $Path |
        Where-Object { $_ -match ('^' + $escaped + '\s*:\s*(.*)$') } |
        Select-Object -First 1
    if ($null -eq $line) { return $null }
    return [regex]::Match($line, ('^' + $escaped + '\s*:\s*(.*)$')).Groups[1].Value.Trim()
}

function Convert-ToDouble {
    param([object]$Value)
    $result = 0.0
    $text = [string]$Value
    if ([double]::TryParse(
            $text,
            [Globalization.NumberStyles]::Any,
            [Globalization.CultureInfo]::InvariantCulture,
            [ref]$result)) {
        return $result
    }
    return [double]$text
}

function Get-NearestRankPercentile {
    param([double[]]$Values, [double]$Percentile)
    if ($Values.Count -eq 0) { return 0.0 }
    $sorted = @($Values | Sort-Object)
    $index = [Math]::Max(0, [Math]::Ceiling($Percentile * $sorted.Count) - 1)
    return [double]$sorted[$index]
}

function Get-WaitDurations {
    param([string]$TransitionPath, [int]$MeasurementSeconds)
    $durations = New-Object Collections.Generic.List[double]
    $rows = @(Import-Csv -LiteralPath $TransitionPath)
    $previousClass = ''
    $waitOnsetMs = $null
    foreach ($row in $rows) {
        $elapsedMs = Convert-ToDouble $row.ElapsedMs
        $class = [string]$row.Class
        if ($class -eq 'WAIT' -and $previousClass -ne 'WAIT') {
            $waitOnsetMs = $elapsedMs
        } elseif ($class -ne 'WAIT' -and $previousClass -eq 'WAIT' -and
                  $null -ne $waitOnsetMs) {
            $durations.Add($elapsedMs - [double]$waitOnsetMs)
            $waitOnsetMs = $null
        }
        $previousClass = $class
    }
    if ($previousClass -eq 'WAIT' -and $null -ne $waitOnsetMs) {
        $durations.Add($MeasurementSeconds * 1000.0 - [double]$waitOnsetMs)
    }
    return [double[]]$durations.ToArray()
}

$setRows = New-Object Collections.Generic.List[object]
$windowRows = New-Object Collections.Generic.List[object]

foreach ($requestedSet in $SetDirectories) {
    $setDirectory = [IO.Path]::GetFullPath($requestedSet)
    if (-not (Test-Path -LiteralPath $setDirectory -PathType Container)) {
        throw "Measurement set directory was not found: $setDirectory"
    }

    $tickValues = New-Object Collections.Generic.List[double]
    $aggregateValues = New-Object Collections.Generic.List[double]
    $aggregateExistsValues = New-Object Collections.Generic.List[double]
    $aggregateOpenValues = New-Object Collections.Generic.List[double]
    $aggregateLineInputValues = New-Object Collections.Generic.List[double]
    $aggregateDecodeValues = New-Object Collections.Generic.List[double]
    $aggregateCloseValues = New-Object Collections.Generic.List[double]
    $waitValues = New-Object Collections.Generic.List[double]
    $rejections = New-Object Collections.Generic.List[string]
    $validWindows = 0
    $waitOnsets = 0
    $validMinutes = 0.0
    $attemptDirectories = @(Get-ChildItem -LiteralPath $setDirectory -Directory |
        Where-Object { $_.Name -like 'attempt-*' } |
        Sort-Object Name)

    foreach ($attemptDirectory in $attemptDirectories) {
        $desktopSummary = Get-ChildItem -LiteralPath $attemptDirectory.FullName `
            -Filter 'desktop-summary_*.txt' -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTimeUtc |
            Select-Object -Last 1
        $attemptMetadataPath = Join-Path $attemptDirectory.FullName 'attempt.json'
        $attemptMetadata = if (Test-Path -LiteralPath $attemptMetadataPath) {
            Get-Content -Raw -LiteralPath $attemptMetadataPath | ConvertFrom-Json
        } else { $null }

        if ($null -eq $desktopSummary) {
            $reason = if ($null -ne $attemptMetadata -and $attemptMetadata.Error) {
                [string]$attemptMetadata.Error
            } else { 'desktop summary missing' }
            $rejections.Add("$($attemptDirectory.Name): $reason")
            continue
        }

        $valid = ((Read-SummaryField $desktopSummary.FullName 'Window valid') -eq 'True')
        $seconds = [int](Read-SummaryField $desktopSummary.FullName 'Measurement seconds')
        $pumpCount = [int](Read-SummaryField $desktopSummary.FullName 'Pump samples')
        $attemptWaitOnsets = [int](Read-SummaryField $desktopSummary.FullName 'WAIT onset count')
        $jobCheck = Read-SummaryField $desktopSummary.FullName 'Active job check passed'
        if (-not $valid) {
            $reason = if ($null -ne $attemptMetadata -and $attemptMetadata.Error) {
                [string]$attemptMetadata.Error
            } else { 'window validity checks failed' }
            $rejections.Add("$($attemptDirectory.Name): $reason")
            $windowRows.Add([PSCustomObject]@{
                Configuration = Split-Path $setDirectory -Leaf
                Attempt = $attemptDirectory.Name
                Valid = $false
                Seconds = $seconds
                PumpSamples = $pumpCount
                WaitOnsets = $attemptWaitOnsets
                ActiveJobCheck = $jobCheck
                Rejection = $reason
            })
            continue
        }

        $pumpPath = Get-ChildItem -LiteralPath $attemptDirectory.FullName `
            -Filter 'pump-profile_*.csv' | Select-Object -Last 1 -ExpandProperty FullName
        $transitionPath = Get-ChildItem -LiteralPath $attemptDirectory.FullName `
            -Filter 'cursor-transitions_*.csv' | Select-Object -Last 1 -ExpandProperty FullName
        foreach ($pumpRow in (Import-Csv -LiteralPath $pumpPath)) {
            $tickValues.Add((Convert-ToDouble $pumpRow.ElapsedMs))
            $aggregateValues.Add((Convert-ToDouble $pumpRow.AggregateReadMs))
            if ($pumpRow.PSObject.Properties.Name -contains 'AggregateExistsMs') {
                $aggregateExistsValues.Add((Convert-ToDouble $pumpRow.AggregateExistsMs))
                $aggregateOpenValues.Add((Convert-ToDouble $pumpRow.AggregateOpenMs))
                $aggregateLineInputValues.Add((Convert-ToDouble $pumpRow.AggregateLineInputMs))
                $aggregateDecodeValues.Add((Convert-ToDouble $pumpRow.AggregateDecodeMs))
                if ($pumpRow.PSObject.Properties.Name -contains 'AggregateCloseMs') {
                    $aggregateCloseValues.Add((Convert-ToDouble $pumpRow.AggregateCloseMs))
                }
            }
        }
        foreach ($duration in (Get-WaitDurations $transitionPath $seconds)) {
            $waitValues.Add($duration)
        }
        $validWindows++
        $waitOnsets += $attemptWaitOnsets
        $validMinutes += $seconds / 60.0
        $windowRows.Add([PSCustomObject]@{
            Configuration = Split-Path $setDirectory -Leaf
            Attempt = $attemptDirectory.Name
            Valid = $true
            Seconds = $seconds
            PumpSamples = $pumpCount
            WaitOnsets = $attemptWaitOnsets
            ActiveJobCheck = $jobCheck
            Rejection = ''
        })
    }

    $setRows.Add([PSCustomObject][ordered]@{
        Configuration = Split-Path $setDirectory -Leaf
        Attempts = $attemptDirectories.Count
        ValidWindows = $validWindows
        RejectedWindows = $attemptDirectories.Count - $validWindows
        RejectionReasons = $rejections -join ' | '
        TickCount = $tickValues.Count
        TickP50Ms = [Math]::Round((Get-NearestRankPercentile $tickValues.ToArray() 0.50), 3)
        TickP95Ms = [Math]::Round((Get-NearestRankPercentile $tickValues.ToArray() 0.95), 3)
        AggregateP50Ms = [Math]::Round((Get-NearestRankPercentile $aggregateValues.ToArray() 0.50), 3)
        AggregateP95Ms = [Math]::Round((Get-NearestRankPercentile $aggregateValues.ToArray() 0.95), 3)
        AggregateExistsP95Ms = [Math]::Round((Get-NearestRankPercentile $aggregateExistsValues.ToArray() 0.95), 3)
        AggregateOpenP95Ms = [Math]::Round((Get-NearestRankPercentile $aggregateOpenValues.ToArray() 0.95), 3)
        AggregateLineInputP95Ms = [Math]::Round((Get-NearestRankPercentile $aggregateLineInputValues.ToArray() 0.95), 3)
        AggregateDecodeP95Ms = [Math]::Round((Get-NearestRankPercentile $aggregateDecodeValues.ToArray() 0.95), 3)
        AggregateCloseP95Ms = [Math]::Round((Get-NearestRankPercentile $aggregateCloseValues.ToArray() 0.95), 3)
        WaitOnsets = $waitOnsets
        WaitP50Ms = [Math]::Round((Get-NearestRankPercentile $waitValues.ToArray() 0.50), 3)
        WaitP95Ms = [Math]::Round((Get-NearestRankPercentile $waitValues.ToArray() 0.95), 3)
        WaitMaxMs = [Math]::Round((Get-NearestRankPercentile $waitValues.ToArray() 1.00), 3)
        WaitOnsetsPerMinute = if ($validMinutes -gt 0) {
            [Math]::Round($waitOnsets / $validMinutes, 2)
        } else { 0.0 }
    })
}

if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Split-Path ([IO.Path]::GetFullPath($SetDirectories[0])) -Parent
}
if (-not (Test-Path -LiteralPath $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
}
$setArray = [object[]]$setRows.ToArray()
$windowArray = [object[]]$windowRows.ToArray()
$setArray | Export-Csv -LiteralPath (Join-Path $OutputDirectory 'phase2-summary.csv') `
    -NoTypeInformation -Encoding UTF8
$windowArray | Export-Csv -LiteralPath (Join-Path $OutputDirectory 'phase2-windows.csv') `
    -NoTypeInformation -Encoding UTF8
$setArray | ConvertTo-Json -Depth 4 |
    Out-File -LiteralPath (Join-Path $OutputDirectory 'phase2-summary.json') -Encoding UTF8
$setArray
