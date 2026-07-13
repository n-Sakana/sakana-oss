# test/test-vba-devkit.ps1 -- vba-devkit port tests
. (Join-Path $PSScriptRoot "_assert.ps1")
. (Join-Path $PSScriptRoot "..\common\install.ps1")
$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$dir = Join-Path $root "tool\vba-devkit"
$main = Join-Path $dir "main.ps1"
$fixtureDir = Join-Path $root "test\fixtures\vba-devkit"
$fixtureA = Join-Path $fixtureDir "fixtureA.xlsm"
$fixtureB = Join-Path $fixtureDir "fixtureB.xlsm"
$protectedXls = Join-Path $fixtureDir "protected.xls"
$outputRoot = Join-Path $root "output"
$logRoot = Join-Path $root "log"

function Get-VbaOutputDirectories {
    return @(Get-ChildItem -LiteralPath $outputRoot -Directory -Filter "vba-devkit_*" -ErrorAction SilentlyContinue)
}

function Get-NewVbaOutputDirectories {
    param([string[]]$BeforePaths)
    return @(Get-VbaOutputDirectories | Where-Object { $BeforePaths -notcontains $_.FullName })
}

$outputsBeforeTest = @(Get-VbaOutputDirectories | ForEach-Object { $_.FullName })
$logsBeforeTest = @(Get-ChildItem -LiteralPath $logRoot -File -Filter "vba-devkit_*.log" -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName })
$scratch = Join-Path $env:TEMP ("toolrack_vba_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $scratch -Force | Out-Null
$env:TOOLRACK_NOPAUSE = "1"

try {

# The proposed manifest must pass the real toolrack validator.
$manifest = Read-Manifest $dir
Assert-True $manifest.Ok "vba-devkit tool.json parses"
if ($manifest.Ok) {
    $violations = @(Test-Manifest $manifest.Data "vba-devkit" $dir)
    Assert-True ($violations.Count -eq 0) ("vba-devkit manifest valid (" + ($violations -join "; ") + ")")
    $variants = @($manifest.Data.variants)
    Assert-True ($variants.Count -eq 6) "vba-devkit has six variants"
    $labels = @($variants | ForEach-Object { $_.label })
    Assert-True (($labels -join "|") -eq "Analyze|Analyze (settings...)|Extract|Diff...|Sanitize|Unlock") "vba-devkit variant order is stable"
    $argumentSets = @($variants | ForEach-Object { @($_.args) -join " " })
    Assert-True (($argumentSets -join "|") -eq "-Mode analyze|-Mode analyze -Settings|-Mode extract|-Mode diff|-Mode sanitize|-Mode unlock") "vba-devkit variant arguments map to dispatcher modes"
} else {
    Assert-True $false ("vba-devkit manifest readable (" + ($manifest.Errors -join "; ") + ")")
}

# Extract must run through the dispatcher and write into the shared output root.
$beforeExtract = @(Get-VbaOutputDirectories | ForEach-Object { $_.FullName })
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main -Target $fixtureA -Mode extract | Out-Null
Assert-True ($LASTEXITCODE -eq 0) "vba-devkit Extract exits 0"
$extractOutputs = @(Get-NewVbaOutputDirectories $beforeExtract | Where-Object { $_.Name -like "vba-devkit_extract_*" })
Assert-True ($extractOutputs.Count -eq 1) "vba-devkit Extract creates one shared output directory"
if ($extractOutputs.Count -eq 1) {
    $combined = @(Get-ChildItem -LiteralPath $extractOutputs[0].FullName -Recurse -File -Filter "*_combined.txt")
    Assert-True ($combined.Count -eq 1) "vba-devkit Extract writes one combined source"
    if ($combined.Count -eq 1) {
        $combinedText = [IO.File]::ReadAllText($combined[0].FullName)
        Assert-True ($combinedText -like "*DangerModule*") "vba-devkit Extract includes the known fixture module"
    }
}

# Diff uses the non-dialog -Second seam and must detect a known changed and added module.
$beforeDiff = @(Get-VbaOutputDirectories | ForEach-Object { $_.FullName })
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main -Target $fixtureA -Mode diff -Second $fixtureB | Out-Null
Assert-True ($LASTEXITCODE -eq 0) "vba-devkit Diff exits 0"
$diffOutputs = @(Get-NewVbaOutputDirectories $beforeDiff | Where-Object { $_.Name -like "vba-devkit_diff_*" })
Assert-True ($diffOutputs.Count -eq 1) "vba-devkit Diff creates one shared output directory"
if ($diffOutputs.Count -eq 1) {
    $diffTextPath = Join-Path $diffOutputs[0].FullName "diff.txt"
    Assert-True (Test-Path -LiteralPath $diffTextPath -PathType Leaf) "vba-devkit Diff writes diff.txt"
    if (Test-Path -LiteralPath $diffTextPath -PathType Leaf) {
        $diffText = [IO.File]::ReadAllText($diffTextPath)
        Assert-True ($diffText -like "*Modified modules:*DataProcessor.bas*") "vba-devkit Diff reports the known modified module"
        Assert-True ($diffText -like "*Added modules:*AddedInB.bas*") "vba-devkit Diff reports the known added module"
        Assert-True ($diffText -like "*Unchanged: 4 module(s)*") "vba-devkit Diff reports unchanged modules"
    }
}

# Module integration: repeated import is safe, output names do not collide, and logs use the shared root.
$modulePath = Join-Path $dir "lib\VBAToolkit.psm1"
Import-Module $modulePath -Force -DisableNameChecking
Import-Module $modulePath -Force -DisableNameChecking
Assert-True ($null -ne ("VbaToolkitNative" -as [type])) "vba-devkit native helper survives repeated module import"
$unitOutput1 = New-VbaOutputDir $fixtureA "unit"
$unitOutput2 = New-VbaOutputDir $fixtureA "unit"
Assert-True ($unitOutput1.StartsWith($outputRoot, [StringComparison]::OrdinalIgnoreCase)) "vba-devkit output helper uses shared output root"
Assert-True ($unitOutput1 -ne $unitOutput2) "vba-devkit output helper avoids same-run collisions"
Write-VbaLog "Unit" $fixtureA "unit-test marker"
$unitLog = Get-VbaLogPath
Assert-True ($unitLog.StartsWith($logRoot, [StringComparison]::OrdinalIgnoreCase)) "vba-devkit log helper uses shared log root"
Assert-True ([IO.File]::ReadAllText($unitLog).Contains("unit-test marker")) "vba-devkit log helper writes the message"

# Analyze must emit every documented report type and find the known dangerous module.
$beforeAnalyze = @(Get-VbaOutputDirectories | ForEach-Object { $_.FullName })
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main -Target $fixtureA -Mode analyze | Out-Null
Assert-True ($LASTEXITCODE -eq 0) "vba-devkit Analyze exits 0"
$analyzeOutputs = @(Get-NewVbaOutputDirectories $beforeAnalyze | Where-Object { $_.Name -like "vba-devkit_analyze_*" })
Assert-True ($analyzeOutputs.Count -eq 1) "vba-devkit Analyze creates one shared output directory"
if ($analyzeOutputs.Count -eq 1) {
    $analyzeCsv = Join-Path $analyzeOutputs[0].FullName "analyze.csv"
    $hitsCsv = Join-Path $analyzeOutputs[0].FullName "hits.csv"
    Assert-True (Test-Path -LiteralPath $analyzeCsv -PathType Leaf) "vba-devkit Analyze writes analyze.csv"
    Assert-True (Test-Path -LiteralPath $hitsCsv -PathType Leaf) "vba-devkit Analyze writes hits.csv"
    $analyzeHtml = @(Get-ChildItem -LiteralPath $analyzeOutputs[0].FullName -Recurse -File -Filter "*_analyze.html")
    $analyzeText = @(Get-ChildItem -LiteralPath $analyzeOutputs[0].FullName -Recurse -File -Filter "*_analyze.txt")
    Assert-True ($analyzeHtml.Count -eq 1) "vba-devkit Analyze writes one HTML report"
    Assert-True ($analyzeText.Count -eq 1) "vba-devkit Analyze writes one text report"
    if (Test-Path -LiteralPath $analyzeCsv -PathType Leaf) {
        $rows = @(Import-Csv -LiteralPath $analyzeCsv)
        Assert-True ($rows.Count -eq 1) "vba-devkit Analyze writes one summary row"
        if ($rows.Count -eq 1) {
            Assert-True ([int]$rows[0].EdrHits -gt 0) "vba-devkit Analyze detects the known EDR statements"
        }
    }
    if (Test-Path -LiteralPath $hitsCsv -PathType Leaf) {
        $hitText = [IO.File]::ReadAllText($hitsCsv)
        Assert-True ($hitText -like "*DangerModule*") "vba-devkit Analyze identifies the known dangerous module"
    }
}

# Sanitize must alter only its copied workbook and leave the original bytes untouched.
$fixtureAHashBefore = (Get-FileHash -LiteralPath $fixtureA -Algorithm SHA256).Hash
$beforeSanitize = @(Get-VbaOutputDirectories | ForEach-Object { $_.FullName })
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main -Target $fixtureA -Mode sanitize | Out-Null
Assert-True ($LASTEXITCODE -eq 0) "vba-devkit Sanitize exits 0"
$sanitizeOutputs = @(Get-NewVbaOutputDirectories $beforeSanitize | Where-Object { $_.Name -like "vba-devkit_sanitize_*" })
Assert-True ($sanitizeOutputs.Count -eq 1) "vba-devkit Sanitize creates one shared output directory"
Assert-True ((Get-FileHash -LiteralPath $fixtureA -Algorithm SHA256).Hash -eq $fixtureAHashBefore) "vba-devkit Sanitize leaves the source bytes unchanged"
if ($sanitizeOutputs.Count -eq 1) {
    $sanitizedFile = Join-Path $sanitizeOutputs[0].FullName "fixtureA_sanitized.xlsm"
    $sanitizeCsv = Join-Path $sanitizeOutputs[0].FullName "sanitize.csv"
    Assert-True (Test-Path -LiteralPath $sanitizedFile -PathType Leaf) "vba-devkit Sanitize writes a copied workbook"
    Assert-True (Test-Path -LiteralPath $sanitizeCsv -PathType Leaf) "vba-devkit Sanitize writes sanitize.csv"
    Assert-True (Test-Path -LiteralPath (Join-Path $sanitizeOutputs[0].FullName "fixtureA_sanitized.html") -PathType Leaf) "vba-devkit Sanitize writes an HTML report"
    if (Test-Path -LiteralPath $sanitizedFile -PathType Leaf) {
        $sanitizedProject = Get-AllModuleCode $sanitizedFile -StripAttributes
        $sanitizedCode = @($sanitizedProject.Modules.Keys | ForEach-Object { $sanitizedProject.Modules[$_].Code }) -join "`n"
        Assert-True ($sanitizedCode -like "*sanitized: *****") "vba-devkit Sanitize writes the fixed sanitized marker"
        Assert-True ($sanitizedCode -notlike "*cmd.exe /c echo TOOLRACK_TEST*") "vba-devkit Sanitize removes the known dangerous command"
    }
    if (Test-Path -LiteralPath $sanitizeCsv -PathType Leaf) {
        $sanitizeRows = @(Import-Csv -LiteralPath $sanitizeCsv)
        Assert-True ($sanitizeRows.Count -eq 1 -and $sanitizeRows[0].Status -eq "sanitized") "vba-devkit Sanitize records sanitized status"
        if ($sanitizeRows.Count -eq 1) {
            Assert-True ([int]$sanitizeRows[0].ChangedLines -gt 0) "vba-devkit Sanitize records changed lines"
        }
    }
}

# Unlock .xls uses the pure PowerShell DPB patch and must not alter the input fixture.
$protectedHashBefore = (Get-FileHash -LiteralPath $protectedXls -Algorithm SHA256).Hash
$beforeUnlock = @(Get-VbaOutputDirectories | ForEach-Object { $_.FullName })
$unlockConsole = (& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main -Target $protectedXls -Mode unlock) -join "`n"
$unlockExit = $LASTEXITCODE
Assert-True ($unlockExit -eq 0) "vba-devkit Unlock .xls exits 0"
Assert-True ($unlockConsole -like "*authorized to modify*") "vba-devkit Unlock displays the authorization warning"
$unlockOutputs = @(Get-NewVbaOutputDirectories $beforeUnlock | Where-Object { $_.Name -like "vba-devkit_unlock_*" })
Assert-True ($unlockOutputs.Count -eq 1) "vba-devkit Unlock creates one shared output directory"
Assert-True ((Get-FileHash -LiteralPath $protectedXls -Algorithm SHA256).Hash -eq $protectedHashBefore) "vba-devkit Unlock leaves the source bytes unchanged"
if ($unlockOutputs.Count -eq 1) {
    $unlockedFile = Join-Path $unlockOutputs[0].FullName "protected.xls"
    Assert-True (Test-Path -LiteralPath $unlockedFile -PathType Leaf) "vba-devkit Unlock writes a copied workbook"
    if (Test-Path -LiteralPath $unlockedFile -PathType Leaf) {
        $unlockedAscii = [Text.Encoding]::ASCII.GetString([IO.File]::ReadAllBytes($unlockedFile))
        Assert-True ($unlockedAscii.Contains("DPx=")) "vba-devkit Unlock patches DPB to DPx in the copy"
        $beforeNoOp = @(Get-VbaOutputDirectories | ForEach-Object { $_.FullName })
        $noOpConsole = (& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main -Target $unlockedFile -Mode unlock) -join "`n"
        Assert-True ($LASTEXITCODE -eq 0) "vba-devkit Unlock no-protection result exits 0"
        Assert-True ($noOpConsole -like "*No VBA password hash*") "vba-devkit Unlock reports an already-unprotected copy"
        Assert-True (@(Get-NewVbaOutputDirectories $beforeNoOp | Where-Object { $_.Name -like "vba-devkit_unlock_*" }).Count -eq 1) "vba-devkit Unlock no-op remains isolated in a run directory"
    }
}

# Error contract: unsupported/empty inputs fail, while partial batch outputs are retained.
$emptyDir = Join-Path $scratch "empty"
New-Item -ItemType Directory -Path $emptyDir -Force | Out-Null
foreach ($modeName in @("analyze", "extract", "sanitize")) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main -Target $emptyDir -Mode $modeName | Out-Null
    Assert-True ($LASTEXITCODE -eq 1) ("vba-devkit " + $modeName + " empty input exits 1")
}
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main -Target $emptyDir -Mode diff -Second $fixtureB | Out-Null
Assert-True ($LASTEXITCODE -eq 1) "vba-devkit Diff folder input exits 1"
$plainFile = Join-Path $scratch "plain.txt"
[IO.File]::WriteAllText($plainFile, "not an Excel workbook")
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main -Target $plainFile -Mode unlock | Out-Null
Assert-True ($LASTEXITCODE -eq 1) "vba-devkit Unlock unsupported input exits 1"
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main -Target (Join-Path $scratch "missing.xlsm") -Mode extract | Out-Null
Assert-True ($LASTEXITCODE -eq 1) "vba-devkit missing target exits 1"
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main -Target $fixtureA -Mode diff -Second (Join-Path $scratch "missing-second.xlsm") | Out-Null
Assert-True ($LASTEXITCODE -eq 1) "vba-devkit Diff missing second file exits 1"

$partialDir = Join-Path $scratch "partial"
New-Item -ItemType Directory -Path $partialDir -Force | Out-Null
Copy-Item -LiteralPath $fixtureA -Destination (Join-Path $partialDir "good.xlsm")
[IO.File]::WriteAllBytes((Join-Path $partialDir "broken.xlsm"), ([Text.Encoding]::ASCII.GetBytes("not a zip workbook")))

$beforePartialAnalyze = @(Get-VbaOutputDirectories | ForEach-Object { $_.FullName })
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main -Target $partialDir -Mode analyze | Out-Null
Assert-True ($LASTEXITCODE -eq 1) "vba-devkit Analyze partial batch exits 1"
$partialAnalyzeOutputs = @(Get-NewVbaOutputDirectories $beforePartialAnalyze | Where-Object { $_.Name -like "vba-devkit_analyze_*" })
Assert-True ($partialAnalyzeOutputs.Count -eq 1) "vba-devkit Analyze partial batch retains one run directory"
if ($partialAnalyzeOutputs.Count -eq 1) {
    Assert-True (@(Get-ChildItem -LiteralPath $partialAnalyzeOutputs[0].FullName -Recurse -File -Filter "*_analyze.html").Count -eq 1) "vba-devkit Analyze partial batch retains successful report"
    $partialAnalyzeRows = @(Import-Csv -LiteralPath (Join-Path $partialAnalyzeOutputs[0].FullName "analyze.csv"))
    Assert-True ($partialAnalyzeRows.Count -eq 2) "vba-devkit Analyze partial batch records both files"
}

$beforePartialSanitize = @(Get-VbaOutputDirectories | ForEach-Object { $_.FullName })
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main -Target $partialDir -Mode sanitize | Out-Null
Assert-True ($LASTEXITCODE -eq 1) "vba-devkit Sanitize partial batch exits 1"
$partialSanitizeOutputs = @(Get-NewVbaOutputDirectories $beforePartialSanitize | Where-Object { $_.Name -like "vba-devkit_sanitize_*" })
Assert-True ($partialSanitizeOutputs.Count -eq 1) "vba-devkit Sanitize partial batch retains one run directory"
if ($partialSanitizeOutputs.Count -eq 1) {
    $partialSanitizeRows = @(Import-Csv -LiteralPath (Join-Path $partialSanitizeOutputs[0].FullName "sanitize.csv"))
    Assert-True (@($partialSanitizeRows | Where-Object { $_.Status -eq "sanitized" }).Count -eq 1) "vba-devkit Sanitize partial batch retains successful result"
    Assert-True (@($partialSanitizeRows | Where-Object { $_.Status -eq "error" }).Count -eq 1) "vba-devkit Sanitize partial batch records failed result"
}

Assert-True (-not (Test-Path -LiteralPath (Join-Path $dir "output"))) "vba-devkit does not write a private output directory"
Assert-True (-not (Test-Path -LiteralPath (Join-Path $dir "vba-toolkit.log"))) "vba-devkit does not write the legacy log file"

$asciiFiles = @(
    (Join-Path $dir "main.ps1"),
    (Join-Path $dir "tool.json"),
    (Join-Path $root "test\test-vba-devkit.ps1")
)
foreach ($asciiFile in $asciiFiles) {
    $asciiBytes = [IO.File]::ReadAllBytes($asciiFile)
    Assert-True (@($asciiBytes | Where-Object { $_ -gt 127 }).Count -eq 0) ("ASCII source: " + (Split-Path $asciiFile -Leaf))
}

Remove-Module VBAToolkit -Force -ErrorAction SilentlyContinue

} finally {
    $env:TOOLRACK_NOPAUSE = ""
    foreach ($path in @(Get-VbaOutputDirectories | ForEach-Object { $_.FullName } | Where-Object { $outputsBeforeTest -notcontains $_ })) {
        Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction SilentlyContinue
    }
    foreach ($file in @(Get-ChildItem -LiteralPath $logRoot -File -Filter "vba-devkit_*.log" -ErrorAction SilentlyContinue | Where-Object { $logsBeforeTest -notcontains $_.FullName })) {
        Remove-Item -LiteralPath $file.FullName -Force -ErrorAction SilentlyContinue
    }
    Remove-Item -LiteralPath $scratch -Recurse -Force -ErrorAction SilentlyContinue
}

Exit-Test
