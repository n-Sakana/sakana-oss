# Create repeatable manual fixtures for the three AI chat tools.
# PowerShell 5.1, ASCII source only.
param(
    [string]$Destination = "",
    [switch]$IncludeLegacyOffice,
    [switch]$NoOpen
)

$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
if (-not $Destination) {
    $Destination = Join-Path $root "output\manual-test-fixtures"
}
$destinationFull = [IO.Path]::GetFullPath($Destination).TrimEnd('\')
if (-not $destinationFull -or $destinationFull -eq [IO.Path]::GetPathRoot($destinationFull)) {
    throw "Refusing to use a filesystem root as the fixture destination."
}
$sentinel = Join-Path $destinationFull ".toolrack-manual-fixture-root"
if (Test-Path -LiteralPath $destinationFull) {
    if (-not (Test-Path -LiteralPath $sentinel -PathType Leaf)) {
        throw "Destination exists but is not a generated fixture folder: $destinationFull"
    }
    Remove-Item -LiteralPath $destinationFull -Recurse -Force
}
New-Item -ItemType Directory -Path $destinationFull -Force | Out-Null
[IO.File]::WriteAllText($sentinel, "toolrack manual fixtures`r`n", (New-Object Text.UTF8Encoding($false)))

$utf8 = New-Object Text.UTF8Encoding($false, $true)
$utf8Bom = New-Object Text.UTF8Encoding($true, $true)
$utf16 = New-Object Text.UnicodeEncoding($false, $true, $true)
$cp932 = [Text.Encoding]::GetEncoding(
    932,
    [Text.EncoderFallback]::ExceptionFallback,
    [Text.DecoderFallback]::ExceptionFallback
)
Add-Type -AssemblyName System.Drawing

function Join-FixtureBytes {
    param([byte[]]$First, [byte[]]$Second)
    $result = New-Object byte[] ($First.Length + $Second.Length)
    [Array]::Copy($First, 0, $result, 0, $First.Length)
    [Array]::Copy($Second, 0, $result, $First.Length, $Second.Length)
    return ,$result
}

function Write-FixtureBytes {
    param([string]$Path, [byte[]]$Bytes)
    $parent = Split-Path $Path -Parent
    if (-not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    [IO.File]::WriteAllBytes($Path, $Bytes)
}

function Write-FixtureText {
    param([string]$Path, [string]$Text, [Text.Encoding]$Encoding)
    Write-FixtureBytes $Path $Encoding.GetBytes($Text)
}

function Write-FixtureTextWithBom {
    param([string]$Path, [string]$Text, [Text.Encoding]$Encoding)
    Write-FixtureBytes $Path (Join-FixtureBytes $Encoding.GetPreamble() $Encoding.GetBytes($Text))
}

function Add-DocumentLine {
    param([Text.StringBuilder]$Builder, [string]$Line)
    [void]$Builder.Append($Line + "`n")
}

function ConvertTo-FixtureBase64Lines {
    param([byte[]]$Bytes)
    $encoded = [Convert]::ToBase64String($Bytes)
    $lines = New-Object System.Collections.Generic.List[string]
    for ($offset = 0; $offset -lt $encoded.Length; $offset += 76) {
        $length = [Math]::Min(76, $encoded.Length - $offset)
        $lines.Add($encoded.Substring($offset, $length))
    }
    return $lines.ToArray()
}

function New-FixturePng {
    param([string]$Path, [Drawing.Color]$Background, [Drawing.Color]$Accent)
    $parent = Split-Path $Path -Parent
    if (-not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    $bitmap = New-Object Drawing.Bitmap 64,64
    $graphics = [Drawing.Graphics]::FromImage($bitmap)
    $brush = New-Object Drawing.SolidBrush $Accent
    try {
        $graphics.Clear($Background)
        $graphics.FillEllipse($brush, 12, 12, 40, 40)
        $bitmap.Save($Path, [Drawing.Imaging.ImageFormat]::Png)
    } finally {
        $brush.Dispose()
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

function New-OpenXmlFixture {
    param([string]$Path, [hashtable]$Entries)
    Add-Type -AssemblyName System.IO.Compression
    $parent = Split-Path $Path -Parent
    if (-not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    $stream = [IO.File]::Open($Path, [IO.FileMode]::Create, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    $archive = New-Object IO.Compression.ZipArchive($stream, [IO.Compression.ZipArchiveMode]::Create, $false)
    try {
        foreach ($name in @($Entries.Keys | Sort-Object)) {
            $entry = $archive.CreateEntry($name)
            $writer = New-Object IO.StreamWriter($entry.Open(), $utf8)
            try { $writer.Write([string]$Entries[$name]) } finally { $writer.Dispose() }
        }
    } finally {
        $archive.Dispose()
        $stream.Dispose()
    }
}

function New-SimplePdfFixture {
    param([string]$Path, [string]$Text)
    $safeText = $Text.Replace("\", "\\").Replace("(", "\(").Replace(")", "\)")
    $content = "BT /F1 18 Tf 72 720 Td ($safeText) Tj ET`n"
    $objects = @(
        "<< /Type /Catalog /Pages 2 0 R >>",
        "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
        ("<< /Length " + $utf8.GetByteCount($content) + " >>`nstream`n" + $content + "endstream"),
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"
    )
    $stream = New-Object IO.MemoryStream
    $offsets = New-Object System.Collections.Generic.List[long]
    $write = {
        param([string]$Value)
        $bytes = [Text.Encoding]::ASCII.GetBytes($Value)
        $stream.Write($bytes, 0, $bytes.Length)
    }
    try {
        & $write "%PDF-1.4`n"
        for ($number = 1; $number -le $objects.Count; $number++) {
            $offsets.Add($stream.Position)
            & $write ("$number 0 obj`n" + $objects[$number - 1] + "`nendobj`n")
        }
        $xref = $stream.Position
        & $write ("xref`n0 " + ($objects.Count + 1) + "`n")
        & $write "0000000000 65535 f `n"
        foreach ($offset in $offsets) {
            & $write ($offset.ToString("0000000000") + " 00000 n `n")
        }
        & $write ("trailer`n<< /Size " + ($objects.Count + 1) + " /Root 1 0 R >>`n")
        & $write ("startxref`n" + $xref + "`n%%EOF`n")
        Write-FixtureBytes $Path $stream.ToArray()
    } finally {
        $stream.Dispose()
    }
}

function New-LegacyOfficeFixture {
    param([string]$Kind, [string]$Path, [string]$Text)
    $processName = if ($Kind -in @("word", "pdf")) { "WINWORD" } elseif ($Kind -eq "excel") { "EXCEL" } else { "POWERPNT" }
    $before = @(Get-Process -Name $processName -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $job = Start-Job -ArgumentList $Kind, $Path, $Text -ScriptBlock {
        param($fixtureKind, $fixturePath, $fixtureText)
        $app = $null
        $document = $null
        $workbook = $null
        $presentation = $null
        try {
            if ($fixtureKind -eq "word" -or $fixtureKind -eq "pdf") {
                $app = New-Object -ComObject Word.Application
                $app.Visible = $false
                $app.DisplayAlerts = 0
                $document = $app.Documents.Add()
                $document.Content.Text = $fixtureText
                if ($fixtureKind -eq "word") { $document.SaveAs2($fixturePath, 0) }
                else { $document.ExportAsFixedFormat($fixturePath, 17) }
                $document.Close($false)
                $document = $null
                $app.Quit()
                $app = $null
            } elseif ($fixtureKind -eq "excel") {
                $app = New-Object -ComObject Excel.Application
                $app.Visible = $false
                $app.DisplayAlerts = $false
                $workbook = $app.Workbooks.Add()
                $workbook.Worksheets.Item(1).Cells.Item(1, 1).Value2 = $fixtureText
                $workbook.SaveAs($fixturePath, 56)
                $workbook.Close($false)
                $workbook = $null
                $app.Quit()
                $app = $null
            } else {
                $app = New-Object -ComObject PowerPoint.Application
                $presentation = $app.Presentations.Add()
                $slide = $presentation.Slides.Add(1, 12)
                $shape = $slide.Shapes.AddTextbox(1, 10, 10, 500, 80)
                $shape.TextFrame.TextRange.Text = $fixtureText
                $presentation.SaveAs($fixturePath, 1)
                $presentation.Close()
                $presentation = $null
                $app.Quit()
                $app = $null
            }
        } finally {
            if ($null -ne $document) { try { $document.Close($false) } catch { } }
            if ($null -ne $workbook) { try { $workbook.Close($false) } catch { } }
            if ($null -ne $presentation) { try { $presentation.Close() } catch { } }
            if ($null -ne $app) { try { $app.Quit() } catch { } }
        }
    }
    $finished = Wait-Job $job -Timeout 45
    try {
        if ($null -eq $finished) {
            Stop-Job $job -ErrorAction SilentlyContinue
            return $false
        }
        Receive-Job $job -ErrorAction Stop | Out-Null
        return (Test-Path -LiteralPath $Path -PathType Leaf)
    } catch {
        return $false
    } finally {
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        $after = @(Get-Process -Name $processName -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id })
        foreach ($process in $after) {
            Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
        }
    }
}

function Test-FixtureBytesEqual {
    param([byte[]]$Left, [byte[]]$Right)
    if ($Left.Length -ne $Right.Length) { return $false }
    for ($index = 0; $index -lt $Left.Length; $index++) {
        if ($Left[$index] -ne $Right[$index]) { return $false }
    }
    return $true
}

function Assert-Fixture {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Fixture validation failed: $Message" }
}

# MD Mirror source tree.
$mirrorSource = Join-Path $destinationFull "mirror-source"
New-Item -ItemType Directory -Path (Join-Path $mirrorSource "empty-dir") -Force | Out-Null
Write-FixtureText (Join-Path $mirrorSource "src\app.ps1") "param()`r`nWrite-Output 'mirror fixture'`r`n" $utf8
Write-FixtureText (Join-Path $mirrorSource "src\calc.py") "def add(left, right):`n    return left + right`n" $utf8
Write-FixtureTextWithBom (Join-Path $mirrorSource "docs\utf8-bom.txt") "UTF-8 BOM`r`nsecond line`r`n" $utf8Bom
Write-FixtureTextWithBom (Join-Path $mirrorSource "docs\utf16le.txt") "UTF-16 LE`r`nsecond line`r`n" $utf16
$jpText = ([char]0x65E5).ToString() + [char]0x672C + [char]0x8A9E + " CP932`r`n"
Write-FixtureText (Join-Path $mirrorSource "docs\cp932.txt") $jpText $cp932
Write-FixtureText (Join-Path $mirrorSource "no-final-newline.txt") "no final newline" $utf8
Write-FixtureBytes (Join-Path $mirrorSource "empty.txt") ([byte[]]@())
Write-FixtureText (Join-Path $mirrorSource "mixed-eol.txt") "one`r`ntwo`nthree`r`n" $utf8
$markerText = @(
    "This file deliberately contains old-looking separators.",
    "===== FILE: fake.txt =====",
    "----- OLD CONTENT: 1 LINES -----",
    '```powershell',
    "Write-Output 'ordinary file content'",
    '```'
) -join "`r`n"
Write-FixtureText (Join-Path $mirrorSource "marker-looking-text.txt") ($markerText + "`r`n") $utf8
Write-FixtureText (Join-Path $mirrorSource "deep\a\b\c\deep.txt") "deep hierarchy`n" $utf8
$jpName = ([char]0x65E5).ToString() + [char]0x672C + [char]0x8A9E + ".txt"
Write-FixtureText (Join-Path $mirrorSource $jpName) ($jpText.Replace("CP932", "UTF-8")) $utf8
New-FixturePng (Join-Path $mirrorSource "assets\logo.png") ([Drawing.Color]::CornflowerBlue) ([Drawing.Color]::Gold)

. (Join-Path $root "tool\md-mirror\format.ps1")
$mirrorPackage = New-MirrorPackage -Target $mirrorSource
$readyMirror = Join-Path $destinationFull "ready-mirror.md"
[IO.File]::WriteAllText($readyMirror, $mirrorPackage.Document, $utf8)
$parsedMirror = Read-MirrorDocument -Text $mirrorPackage.Document
Assert-Fixture ($parsedMirror.Entries.Count -eq $mirrorPackage.Entries.Count) "ready-mirror.md parses"
$validationRestore = Join-Path $destinationFull ".validation-restore"
$restoreResult = Restore-MirrorDocument -Mirror $parsedMirror -Destination $validationRestore
try {
    foreach ($file in @(Get-ChildItem -LiteralPath $mirrorSource -File -Recurse -Force)) {
        $relative = $file.FullName.Substring($mirrorSource.Length).TrimStart('\')
        $restoredPath = Join-Path $validationRestore $relative
        Assert-Fixture (Test-Path -LiteralPath $restoredPath -PathType Leaf) ("mirror restore contains " + $relative)
        if ($relative -ieq "mixed-eol.txt") {
            $expectedMixed = $utf8.GetBytes("one`r`ntwo`r`nthree`r`n")
            Assert-Fixture (Test-FixtureBytesEqual $expectedMixed ([IO.File]::ReadAllBytes($restoredPath))) "mixed EOL is normalized"
        } else {
            Assert-Fixture (Test-FixtureBytesEqual ([IO.File]::ReadAllBytes($file.FullName)) ([IO.File]::ReadAllBytes($restoredPath))) ("mirror byte round trip: " + $relative)
        }
    }
    Assert-Fixture (Test-Path -LiteralPath (Join-Path $validationRestore "empty-dir") -PathType Container) "empty directory restores"
} finally {
    Remove-Item -LiteralPath $validationRestore -Recurse -Force -ErrorAction SilentlyContinue
}

# MD Patch target and documents.
$patchTarget = Join-Path $destinationFull "patch-target"
$patches = Join-Path $destinationFull "patches"
New-Item -ItemType Directory -Path $patches -Force | Out-Null
$appPath = Join-Path $patchTarget "src\app.ps1"
$obsoletePath = Join-Path $patchTarget "obsolete.txt"
$assetPath = Join-Path $patchTarget "assets\logo.png"
Write-FixtureText $appPath ('param()' + "`r`n" + '$name = "world"' + "`r`n" + 'Write-Output ("Hello, " + $name)' + "`r`n") $utf8
Write-FixtureText $obsoletePath "delete me`r`n" $utf8
New-FixturePng $assetPath ([Drawing.Color]::DarkSlateBlue) ([Drawing.Color]::White)
$replacementPng = Join-Path $destinationFull ".replacement.png"
New-FixturePng $replacementPng ([Drawing.Color]::SeaGreen) ([Drawing.Color]::Orange)

. (Join-Path $root "tool\md-patch\format.ps1")
. (Join-Path $root "tool\md-patch\engine.ps1")
$appHash = Get-PatchSha256 ([IO.File]::ReadAllBytes($appPath))
$obsoleteHash = Get-PatchSha256 ([IO.File]::ReadAllBytes($obsoletePath))
$assetHash = Get-PatchSha256 ([IO.File]::ReadAllBytes($assetPath))
$replacementBytes = [IO.File]::ReadAllBytes($replacementPng)
$replacementHash = Get-PatchSha256 $replacementBytes
$base64Lines = @(ConvertTo-FixtureBase64Lines $replacementBytes)

$patchBuilder = New-Object Text.StringBuilder
Add-DocumentLine $patchBuilder "# MD Patch v1"
Add-DocumentLine $patchBuilder "Operation Count: 5"
Add-DocumentLine $patchBuilder ""
Add-DocumentLine $patchBuilder "===== MODIFY FILE: src/app.ps1 ====="
Add-DocumentLine $patchBuilder ("Expected SHA-256: " + $appHash)
Add-DocumentLine $patchBuilder "Change Count: 1"
Add-DocumentLine $patchBuilder "----- CHANGE 1 -----"
Add-DocumentLine $patchBuilder "Old Start Line: 2"
Add-DocumentLine $patchBuilder "Old Line Count: 1"
Add-DocumentLine $patchBuilder "New Start Line: 2"
Add-DocumentLine $patchBuilder "New Line Count: 1"
Add-DocumentLine $patchBuilder "----- OLD CONTENT: 1 LINES -----"
Add-DocumentLine $patchBuilder '$name = "world"'
Add-DocumentLine $patchBuilder "----- NEW CONTENT: 1 LINES -----"
Add-DocumentLine $patchBuilder '$name = "toolrack"'
Add-DocumentLine $patchBuilder "----- END CHANGE 1 -----"
Add-DocumentLine $patchBuilder "===== END MODIFY FILE: src/app.ps1 ====="
Add-DocumentLine $patchBuilder ""
Add-DocumentLine $patchBuilder "===== CREATE TEXT FILE: created.txt ====="
Add-DocumentLine $patchBuilder "Encoding: utf-8"
Add-DocumentLine $patchBuilder "EOL: crlf"
Add-DocumentLine $patchBuilder "Final Newline: yes"
Add-DocumentLine $patchBuilder "New Line Count: 1"
Add-DocumentLine $patchBuilder "----- NEW CONTENT: 1 LINES -----"
Add-DocumentLine $patchBuilder "Created by the MD Patch manual fixture."
Add-DocumentLine $patchBuilder "===== END CREATE TEXT FILE: created.txt ====="
Add-DocumentLine $patchBuilder ""
Add-DocumentLine $patchBuilder "===== DELETE FILE: obsolete.txt ====="
Add-DocumentLine $patchBuilder ("Expected SHA-256: " + $obsoleteHash)
Add-DocumentLine $patchBuilder "===== END DELETE FILE: obsolete.txt ====="
Add-DocumentLine $patchBuilder ""
Add-DocumentLine $patchBuilder "===== CREATE DIRECTORY: empty-created/ ====="
Add-DocumentLine $patchBuilder "===== END CREATE DIRECTORY: empty-created/ ====="
Add-DocumentLine $patchBuilder ""
Add-DocumentLine $patchBuilder "===== REPLACE BINARY FILE: assets/logo.png ====="
Add-DocumentLine $patchBuilder ("Expected SHA-256: " + $assetHash)
Add-DocumentLine $patchBuilder ("Bytes: " + $replacementBytes.Length)
Add-DocumentLine $patchBuilder ("SHA-256: " + $replacementHash)
Add-DocumentLine $patchBuilder ("Base64 Line Count: " + $base64Lines.Count)
Add-DocumentLine $patchBuilder ("----- BASE64 CONTENT: " + $base64Lines.Count + " LINES -----")
foreach ($line in $base64Lines) { Add-DocumentLine $patchBuilder $line }
Add-DocumentLine $patchBuilder "===== END REPLACE BINARY FILE: assets/logo.png ====="
Add-DocumentLine $patchBuilder ""
Add-DocumentLine $patchBuilder "===== END PATCH ====="
$applyPatchText = $patchBuilder.ToString()
[IO.File]::WriteAllText((Join-Path $patches "apply-demo.md"), $applyPatchText, $utf8)

$mismatchBuilder = New-Object Text.StringBuilder
Add-DocumentLine $mismatchBuilder "# MD Patch v1"
Add-DocumentLine $mismatchBuilder "Operation Count: 1"
Add-DocumentLine $mismatchBuilder ""
Add-DocumentLine $mismatchBuilder "===== MODIFY FILE: src/app.ps1 ====="
Add-DocumentLine $mismatchBuilder ("Expected SHA-256: " + $appHash)
Add-DocumentLine $mismatchBuilder "Change Count: 1"
Add-DocumentLine $mismatchBuilder "----- CHANGE 1 -----"
Add-DocumentLine $mismatchBuilder "Old Start Line: 2"
Add-DocumentLine $mismatchBuilder "Old Line Count: 1"
Add-DocumentLine $mismatchBuilder "New Start Line: 2"
Add-DocumentLine $mismatchBuilder "New Line Count: 1"
Add-DocumentLine $mismatchBuilder "----- OLD CONTENT: 1 LINES -----"
Add-DocumentLine $mismatchBuilder '$name = "not-the-current-value"'
Add-DocumentLine $mismatchBuilder "----- NEW CONTENT: 1 LINES -----"
Add-DocumentLine $mismatchBuilder '$name = "must-not-be-applied"'
Add-DocumentLine $mismatchBuilder "----- END CHANGE 1 -----"
Add-DocumentLine $mismatchBuilder "===== END MODIFY FILE: src/app.ps1 ====="
Add-DocumentLine $mismatchBuilder ""
Add-DocumentLine $mismatchBuilder "===== END PATCH ====="
$mismatchPatchText = $mismatchBuilder.ToString()
[IO.File]::WriteAllText((Join-Path $patches "reject-old-content-mismatch.md"), $mismatchPatchText, $utf8)

$hashBuilder = New-Object Text.StringBuilder
Add-DocumentLine $hashBuilder "# MD Patch v1"
Add-DocumentLine $hashBuilder "Operation Count: 1"
Add-DocumentLine $hashBuilder ""
Add-DocumentLine $hashBuilder "===== DELETE FILE: obsolete.txt ====="
Add-DocumentLine $hashBuilder ("Expected SHA-256: " + ("0" * 64))
Add-DocumentLine $hashBuilder "===== END DELETE FILE: obsolete.txt ====="
Add-DocumentLine $hashBuilder ""
Add-DocumentLine $hashBuilder "===== END PATCH ====="
$hashPatchText = $hashBuilder.ToString()
[IO.File]::WriteAllText((Join-Path $patches "reject-hash-mismatch.md"), $hashPatchText, $utf8)

$attackBuilder = New-Object Text.StringBuilder
Add-DocumentLine $attackBuilder "# MD Patch v1"
Add-DocumentLine $attackBuilder "Operation Count: 1"
Add-DocumentLine $attackBuilder ""
Add-DocumentLine $attackBuilder "===== CREATE TEXT FILE: ../outside.txt ====="
Add-DocumentLine $attackBuilder "Encoding: utf-8"
Add-DocumentLine $attackBuilder "EOL: none"
Add-DocumentLine $attackBuilder "Final Newline: no"
Add-DocumentLine $attackBuilder "New Line Count: 1"
Add-DocumentLine $attackBuilder "----- NEW CONTENT: 1 LINES -----"
Add-DocumentLine $attackBuilder "This must never be written."
Add-DocumentLine $attackBuilder "===== END CREATE TEXT FILE: ../outside.txt ====="
Add-DocumentLine $attackBuilder ""
Add-DocumentLine $attackBuilder "===== END PATCH ====="
$attackPatchText = $attackBuilder.ToString()
[IO.File]::WriteAllText((Join-Path $patches "reject-path-attack.md"), $attackPatchText, $utf8)
Remove-Item -LiteralPath $replacementPng -Force

$parsedPatch = Read-PatchDocument $applyPatchText
$preparedPatch = Prepare-PatchApplication $parsedPatch $patchTarget
Assert-Fixture ($preparedPatch.Actions.Count -eq 5) "apply-demo.md passes preflight"
$rejected = $false
try { [void](Prepare-PatchApplication (Read-PatchDocument $mismatchPatchText) $patchTarget) } catch { $rejected = $true }
Assert-Fixture $rejected "old-content mismatch is rejected"
$rejected = $false
try { [void](Prepare-PatchApplication (Read-PatchDocument $hashPatchText) $patchTarget) } catch { $rejected = $true }
Assert-Fixture $rejected "hash mismatch is rejected"
$rejected = $false
try { [void](Read-PatchDocument $attackPatchText) } catch { $rejected = $true }
Assert-Fixture $rejected "path attack is rejected"

# MD Extract input tree.
$extractSource = Join-Path $destinationFull "extract-source"
Write-FixtureText (Join-Path $extractSource "utf8.txt") "UTF-8 extract fixture`nsecond line`n" $utf8
Write-FixtureTextWithBom (Join-Path $extractSource "utf8-bom.txt") "UTF-8 BOM extract fixture`r`n" $utf8Bom
Write-FixtureTextWithBom (Join-Path $extractSource "utf16le.txt") "UTF-16 LE extract fixture`r`n" $utf16
Write-FixtureText (Join-Path $extractSource "cp932.txt") $jpText $cp932
Write-FixtureText (Join-Path $extractSource "no-final-newline.txt") "extract without final newline" $utf8
Write-FixtureBytes (Join-Path $extractSource "empty.txt") ([byte[]]@())
Write-FixtureText (Join-Path $extractSource "src\sample.py") "print('extract fixture')`n" $utf8
Write-FixtureText (Join-Path $extractSource "deep\a\b\deep.txt") "deep extract fixture`n" $utf8
Write-FixtureText (Join-Path $extractSource ".git\hidden.txt") "this must be excluded`n" $utf8
Write-FixtureBytes (Join-Path $extractSource "unsupported.bin") ([byte[]](1,2,3,4,5))
New-FixturePng (Join-Path $extractSource "screenshot.png") ([Drawing.Color]::SteelBlue) ([Drawing.Color]::White)
New-OpenXmlFixture (Join-Path $extractSource "sample.docx") @{
    "word/document.xml" = '<w:document xmlns:w="urn:w"><w:body><w:p><w:r><w:t>DOCX MANUAL FIXTURE</w:t></w:r></w:p></w:body></w:document>'
}
New-OpenXmlFixture (Join-Path $extractSource "sample.xlsx") @{
    "xl/worksheets/sheet1.xml" = '<worksheet><sheetData><row><c r="A1" t="inlineStr"><is><t>XLSX MANUAL FIXTURE</t></is></c></row></sheetData></worksheet>'
}
New-OpenXmlFixture (Join-Path $extractSource "sample.pptx") @{
    "ppt/slides/slide1.xml" = '<p:sld xmlns:p="urn:p" xmlns:a="urn:a"><p:cSld><a:p><a:r><a:t>PPTX MANUAL FIXTURE</a:t></a:r></a:p></p:cSld></p:sld>'
    "ppt/notesSlides/notesSlide1.xml" = '<p:notes xmlns:p="urn:p" xmlns:a="urn:a"><a:p><a:r><a:t>NOTES MANUAL FIXTURE</a:t></a:r></a:p></p:notes>'
}
New-SimplePdfFixture (Join-Path $extractSource "sample.pdf") "PDF MANUAL FIXTURE"

$legacyStatus = New-Object System.Collections.Generic.List[string]
if ($IncludeLegacyOffice) {
    if ($null -ne [type]::GetTypeFromProgID("Word.Application")) {
        $made = New-LegacyOfficeFixture "word" (Join-Path $extractSource "legacy.doc") "WORD LEGACY MANUAL FIXTURE"
        $legacyStatus.Add(("legacy.doc: " + $(if ($made) { "created" } else { "skipped (creation failed)" })))
        $made = New-LegacyOfficeFixture "pdf" (Join-Path $extractSource "office-created.pdf") "PDF OFFICE MANUAL FIXTURE"
        $legacyStatus.Add(("office-created.pdf: " + $(if ($made) { "created" } else { "skipped (creation failed)" })))
    } else {
        $legacyStatus.Add("legacy.doc and office-created.pdf: skipped (Word is not installed)")
    }
    if ($null -ne [type]::GetTypeFromProgID("Excel.Application")) {
        $made = New-LegacyOfficeFixture "excel" (Join-Path $extractSource "legacy.xls") "EXCEL LEGACY MANUAL FIXTURE"
        $legacyStatus.Add(("legacy.xls: " + $(if ($made) { "created" } else { "skipped (creation failed)" })))
    } else {
        $legacyStatus.Add("legacy.xls: skipped (Excel is not installed)")
    }
    if ($null -ne [type]::GetTypeFromProgID("PowerPoint.Application")) {
        $made = New-LegacyOfficeFixture "powerpoint" (Join-Path $extractSource "legacy.ppt") "POWERPOINT LEGACY MANUAL FIXTURE"
        $legacyStatus.Add(("legacy.ppt: " + $(if ($made) { "created" } else { "skipped (creation failed)" })))
    } else {
        $legacyStatus.Add("legacy.ppt: skipped (PowerPoint is not installed)")
    }
} else {
    $legacyStatus.Add("Legacy Office fixtures were not requested.")
}

$readme = New-Object Text.StringBuilder
Add-DocumentLine $readme "TOOLRACK MANUAL TEST FIXTURES"
Add-DocumentLine $readme ""
Add-DocumentLine $readme "Run test/create-manual-fixtures.ps1 again to reset every fixture."
Add-DocumentLine $readme ""
Add-DocumentLine $readme "mirror-source/"
Add-DocumentLine $readme "  Right-click this folder and run MD Mirror > Create."
Add-DocumentLine $readme "ready-mirror.md"
Add-DocumentLine $readme "  Right-click this file and run MD Mirror > Restore from MD File."
Add-DocumentLine $readme "patch-target/ and patches/"
Add-DocumentLine $readme "  Apply patches/apply-demo.md to patch-target. Re-run this generator before another attempt."
Add-DocumentLine $readme "extract-source/"
Add-DocumentLine $readme "  Right-click this folder and run MD Extract > Default."
Add-DocumentLine $readme ""
Add-DocumentLine $readme "Legacy Office fixture status:"
foreach ($status in $legacyStatus) { Add-DocumentLine $readme ("  " + $status) }
[IO.File]::WriteAllText((Join-Path $destinationFull "README.txt"), $readme.ToString(), $utf8)

Write-Host "Manual fixtures created and validated:" -ForegroundColor Green
Write-Host ("  " + $destinationFull)
Write-Host ("  Mirror entries: " + $mirrorPackage.Entries.Count)
Write-Host ("  Patch operations: " + $preparedPatch.Actions.Count)
Write-Host ("  Extract files: " + @(Get-ChildItem -LiteralPath $extractSource -File -Recurse -Force).Count)

if (-not $NoOpen) {
    Start-Process -FilePath explorer.exe -ArgumentList ('"' + $destinationFull + '"')
}
