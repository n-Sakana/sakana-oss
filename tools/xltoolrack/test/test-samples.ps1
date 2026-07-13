param([ValidateSet('stopwatch','pi','life','all')][string]$Sample = 'stopwatch')

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
    $bookName = [string]$workbook.Name
    $run = QM $bookName 'Infra_Run'

    if ($Sample -in @('stopwatch','all')) {
        $null = (Invoke-XlRun $excel (QM $bookName 'JobTest_ResetPumpProfile'))
        $null = (Invoke-XlRun $excel $run 'stopwatch')
        $sheet = $target.Worksheets.Item('xtr_stopwatch')
        try {
            $deadline = (Get-Date).AddSeconds(9)
            $values = @(0,0,0)
            while ((Get-Date) -lt $deadline) {
                Start-Sleep -Milliseconds 250
                for ($i = 0; $i -lt 3; $i++) {
                    $range = $sheet.Range("B$($i + 3)")
                    try { $values[$i] = [int]$range.Value2 } finally { Release-Com $range }
                }
                if (($values | Where-Object { $_ -lt 4 }).Count -eq 0) { break }
            }
            for ($i = 0; $i -lt 3; $i++) {
                Assert-True ($values[$i] -ge 4) "stopwatch $($i + 1) advanced independently (elapsed=$($values[$i]))"
            }

            $profileText = [string](Invoke-XlRun $excel (QM $bookName 'JobTest_PumpProfile'))
            $profileRows = @($profileText | ConvertFrom-Csv)
            Assert-True ($profileRows.Count -gt 0) 'tick profiler captured in-memory samples'
            $freshProfileRows = @($profileRows | Where-Object { [int]$_.FreshRecords -gt 0 })
            foreach ($profileRow in $freshProfileRows) {
                Assert-True ([double]$profileRow.ElapsedMs -ge 0) 'profile recorded non-negative tick duration'
                Assert-True ([double]$profileRow.AggregateReadMs -ge 0) 'profile recorded non-negative aggregate read duration'
                Assert-True ([double]$profileRow.AggregateExistsMs -ge 0) 'profile recorded non-negative aggregate exists duration'
                Assert-True ([double]$profileRow.AggregateOpenMs -ge 0) 'profile recorded non-negative aggregate open duration'
                Assert-True ([double]$profileRow.AggregateLineInputMs -ge 0) 'profile recorded non-negative aggregate line input duration'
                Assert-True ([double]$profileRow.AggregateDecodeMs -ge 0) 'profile recorded non-negative aggregate decode duration'
                Assert-True ([double]$profileRow.AggregateCloseMs -ge 0) 'profile recorded non-negative aggregate close duration'
                Assert-True ([int]$profileRow.FlushTools -eq 1) 'stopwatch results coalesced to one tool flush per tick'
            }
        } finally { Release-Com $sheet }
    }

    if ($Sample -in @('pi','all')) {
        $null = (Invoke-XlRun $excel $run 'pi_race')
        $sheet = $target.Worksheets.Item('xtr_pi_race')
        try {
            $deadline = (Get-Date).AddSeconds(10)
            $iterations = @(0,0,0)
            $versions = @(0,0,0)
            while ((Get-Date) -lt $deadline) {
                Start-Sleep -Milliseconds 250
                for ($i = 0; $i -lt 3; $i++) {
                    $iterationCell = $sheet.Range("B$($i + 3)")
                    $versionCell = $sheet.Range("D$($i + 3)")
                    try {
                        $iterations[$i] = [double]$iterationCell.Value2
                        $versions[$i] = [int]$versionCell.Value2
                    } finally {
                        Release-Com $iterationCell
                        Release-Com $versionCell
                    }
                }
                if (($versions | Where-Object { $_ -lt 3 }).Count -eq 0) { break }
            }
            for ($i = 0; $i -lt 3; $i++) {
                Assert-True ($iterations[$i] -gt 10000) "pi worker $($i + 1) performed sustained numeric work (iterations=$([int]$iterations[$i]))"
                Assert-True ($versions[$i] -ge 3) "pi worker $($i + 1) pushed progress independently (version=$($versions[$i]))"
            }
        } finally { Release-Com $sheet }
    }

    if ($Sample -in @('life','all')) {
        $null = (Invoke-XlRun $excel $run 'life')
        $sheet = $target.Worksheets.Item('xtr_life')
        try {
            $generationCell = $sheet.Range('W3')
            try {
                $deadline = (Get-Date).AddSeconds(8)
                $generation = 0
                while ((Get-Date) -lt $deadline) {
                    Start-Sleep -Milliseconds 250
                    $generation = [int]$generationCell.Value2
                    if ($generation -ge 2) { break }
                }
                Assert-True ($generation -ge 2) "life advanced generations (generation=$generation)"

                $horizontal = @('J12','K12','L12')
                $vertical = @('K11','K12','K13')
                $expected = if (($generation % 2) -eq 0) { $horizontal } else { $vertical }
                foreach ($address in $expected) {
                    $cell = $sheet.Range($address)
                    try { Assert-Equal 1 ([int]$cell.Value2) "blinker cell $address matches generation parity" } finally { Release-Com $cell }
                }

                $block = $sheet.Range('B3:C4')
                try { $block.Value2 = 1 } finally { Release-Com $block }
                $editedAt = $generation
                $deadline = (Get-Date).AddSeconds(5)
                while ((Get-Date) -lt $deadline) {
                    Start-Sleep -Milliseconds 250
                    $generation = [int]$generationCell.Value2
                    if ($generation -gt $editedAt) { break }
                }
                foreach ($address in @('B3','C3','B4','C4')) {
                    $cell = $sheet.Range($address)
                    try { Assert-Equal 1 ([int]$cell.Value2) "edited stable block survived next generation at $address" } finally { Release-Com $cell }
                }
            } finally { Release-Com $generationCell }
        } finally { Release-Com $sheet }
    }

    $null = (Invoke-XlRun $excel (QM $bookName 'JobTest_StopAll'))
} catch {
    Write-Host "  expected-red: $($_.Exception.Message)" -ForegroundColor Yellow
    Assert-True $false "$Sample sample is available"
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
