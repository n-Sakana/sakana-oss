# test/test-md-mirror.ps1 -- MD Mirror format, encoding, binary, and restore tests
. (Join-Path $PSScriptRoot "_assert.ps1")
$root = Split-Path $PSScriptRoot -Parent
$format = Join-Path $root "tool\md-mirror\format.ps1"
$ErrorActionPreference = "Stop"

Assert-True (Test-Path -LiteralPath $format -PathType Leaf) "md-mirror format.ps1 exists"
if (-not (Test-Path -LiteralPath $format -PathType Leaf)) { Exit-Test }
. $format

$fx = Join-Path $env:TEMP ("toolrack_mm_" + [guid]::NewGuid().ToString("N"))
$src = Join-Path $fx "source"
$dst = Join-Path $fx "restored"
New-Item -ItemType Directory -Force $src | Out-Null

function Write-Bytes {
    param([string]$Path, [byte[]]$Bytes)
    $parent = Split-Path $Path -Parent
    if (-not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Force $parent | Out-Null
    }
    [System.IO.File]::WriteAllBytes($Path, $Bytes)
}

function Join-ByteArrays {
    param([byte[]]$First, [byte[]]$Second)
    $result = New-Object byte[] ($First.Length + $Second.Length)
    [Array]::Copy($First, 0, $result, 0, $First.Length)
    [Array]::Copy($Second, 0, $result, $First.Length, $Second.Length)
    return $result
}

function Test-BytesEqual {
    param([string]$Left, [string]$Right)
    if (-not (Test-Path -LiteralPath $Left -PathType Leaf)) { return $false }
    if (-not (Test-Path -LiteralPath $Right -PathType Leaf)) { return $false }
    $a = [System.IO.File]::ReadAllBytes($Left)
    $b = [System.IO.File]::ReadAllBytes($Right)
    if ($a.Length -ne $b.Length) { return $false }
    for ($i = 0; $i -lt $a.Length; $i++) {
        if ($a[$i] -ne $b[$i]) { return $false }
    }
    return $true
}

function Test-Throws {
    param([scriptblock]$Action)
    try { & $Action; return $false } catch { return $true }
}

try {
    $utf8 = New-Object System.Text.UTF8Encoding($false, $true)
    $utf8Bom = New-Object System.Text.UTF8Encoding($true, $true)
    $utf16 = New-Object System.Text.UnicodeEncoding($false, $true, $true)
    $cp932 = [System.Text.Encoding]::GetEncoding(
        932,
        [System.Text.EncoderFallback]::ExceptionFallback,
        [System.Text.DecoderFallback]::ExceptionFallback
    )

    Write-Bytes (Join-Path $src "utf8-lf.txt") ($utf8.GetBytes("alpha`nbeta`n"))
    Write-Bytes (Join-Path $src "utf8-no-final.txt") ($utf8.GetBytes("no final"))
    Write-Bytes (Join-Path $src "empty.txt") ([byte[]]@())
    Write-Bytes (Join-Path $src "marker.txt") ($utf8.GetBytes("<<<MDMIRROR-DATA>>>`n===== END ENTRY =====`n"))
    Write-Bytes (Join-Path $src "mixed.txt") ($utf8.GetBytes("one`r`ntwo`nthree`r`n"))

    $bom8 = $utf8Bom.GetPreamble()
    Write-Bytes (Join-Path $src "utf8-bom.txt") (Join-ByteArrays $bom8 ($utf8Bom.GetBytes("bom8`r`n")))
    $bom16 = $utf16.GetPreamble()
    Write-Bytes (Join-Path $src "utf16le.txt") (Join-ByteArrays $bom16 ($utf16.GetBytes("wide`r`ntext")))

    $jpText = ([char]0x7D4C).ToString() + [char]0x8CBB + [char]0x30E1 + [char]0x30E2
    $jpName = ([char]0x65E5).ToString() + [char]0x672C + [char]0x8A9E + ".txt"
    Write-Bytes (Join-Path $src $jpName) ($cp932.GetBytes($jpText + "`r`n"))

    New-Item -ItemType Directory -Force (Join-Path $src "empty-dir") | Out-Null
    Write-Bytes (Join-Path $src "a\b\c\d\deep.txt") ($utf8.GetBytes("deep"))

    $png = [byte[]](137,80,78,71,13,10,26,10,0,1,2,3,4,5)
    Write-Bytes (Join-Path $src "image.png") $png
    $video = [byte[]](0,0,0,24,102,116,121,112,105,115,111,109,0,0,0,0)
    Write-Bytes (Join-Path $src "movie.mp4") $video

    $pkg = New-MirrorPackage -Target $src -BinaryFileLimitBytes 1024 -BinaryTotalLimitBytes 2048
    Assert-True ($pkg.Document.StartsWith("# MD Mirror v1`n")) "generated document has the exact version header"
    Assert-True ($pkg.Document -notlike "*```*") "transport does not use code fences"
    Assert-True ($pkg.Document -like "*<<<MDMIRROR-DATA>>>*") "old fixed-marker text remains editable text"
    Assert-True (@($pkg.Warnings | Where-Object { $_ -like "*mixed.txt*normalized*" }).Count -eq 1) "mixed EOL normalization is reported"
    Assert-True (@($pkg.Warnings | Where-Object { $_ -like "*movie.mp4*not embedded*" }).Count -eq 1) "unsupported binary is reported"

    $mirror = Read-MirrorDocument -Text $pkg.Document
    Assert-True ($mirror.Entries.Count -ge 13) "all fixture entries are represented"
    $imageEntry = @($mirror.Entries | Where-Object { $_.Path -eq "image.png" })[0]
    $movieEntry = @($mirror.Entries | Where-Object { $_.Path -eq "movie.mp4" })[0]
    $mixedEntry = @($mirror.Entries | Where-Object { $_.Path -eq "mixed.txt" })[0]
    Assert-True ($imageEntry.Kind -eq "binary" -and $imageEntry.Embedded) "supported image is embedded"
    Assert-True ($movieEntry.Kind -eq "binary" -and -not $movieEntry.Embedded) "video stays inventory-only"
    Assert-True ($mixedEntry.OriginalSha256 -cne $mixedEntry.RestoreSha256) "mixed EOL entry records original and normalized hashes separately"
    Assert-True ($mixedEntry.OriginalSha256 -ceq (Get-MirrorSha256 ([IO.File]::ReadAllBytes((Join-Path $src "mixed.txt")))) ) "Original SHA identifies the current file for MD Patch"

    $result = Restore-MirrorDocument -Mirror $mirror -Destination $dst
    Assert-True ($result.RestoredFiles -ge 10) "restore reports written files"
    Assert-True (Test-Path -LiteralPath (Join-Path $dst "empty-dir") -PathType Container) "empty folder restored"
    Assert-True (Test-Path -LiteralPath (Join-Path $dst "a\b\c\d\deep.txt")) "deep hierarchy restored"
    Assert-True (Test-Path -LiteralPath (Join-Path $dst $jpName)) "Japanese filename restored"
    Assert-True (Test-BytesEqual (Join-Path $src "utf8-lf.txt") (Join-Path $dst "utf8-lf.txt")) "UTF-8 LF bytes round-trip"
    Assert-True (Test-BytesEqual (Join-Path $src "utf8-no-final.txt") (Join-Path $dst "utf8-no-final.txt")) "no-final-newline bytes round-trip"
    Assert-True (Test-BytesEqual (Join-Path $src "utf8-bom.txt") (Join-Path $dst "utf8-bom.txt")) "UTF-8 BOM bytes round-trip"
    Assert-True (Test-BytesEqual (Join-Path $src "utf16le.txt") (Join-Path $dst "utf16le.txt")) "UTF-16LE bytes round-trip"
    Assert-True (Test-BytesEqual (Join-Path $src $jpName) (Join-Path $dst $jpName)) "CP932 bytes round-trip"
    Assert-True (Test-BytesEqual (Join-Path $src "marker.txt") (Join-Path $dst "marker.txt")) "marker-like source bytes round-trip"
    Assert-True (Test-BytesEqual (Join-Path $src "empty.txt") (Join-Path $dst "empty.txt")) "empty file round-trip"
    Assert-True (Test-BytesEqual (Join-Path $src "a\b\c\d\deep.txt") (Join-Path $dst "a\b\c\d\deep.txt")) "single-line no-EOL bytes round-trip"
    Assert-True (Test-BytesEqual (Join-Path $src "image.png") (Join-Path $dst "image.png")) "embedded PNG bytes round-trip"
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $dst "movie.mp4"))) "inventory-only video is not restored"
    $mixedExpected = Join-Path $fx "mixed-expected.txt"
    Write-Bytes $mixedExpected ($utf8.GetBytes("one`r`ntwo`r`nthree`r`n"))
    Assert-True (Test-BytesEqual $mixedExpected (Join-Path $dst "mixed.txt")) "mixed EOL file restores with first EOL"

    $badHeader = $pkg.Document.Replace("# MD Mirror v1", "# md mirror v1")
    Assert-True (Test-Throws { Read-MirrorDocument -Text $badHeader | Out-Null }) "header spelling is strict"
    $badPath = $pkg.Document.Replace("Path: utf8-lf.txt", "Path: ../escape.txt")
    Assert-True (Test-Throws { Read-MirrorDocument -Text $badPath | Out-Null }) "parent-path attack is rejected"
    $badReserved = $pkg.Document.Replace("Path: utf8-lf.txt", "Path: CON.txt")
    Assert-True (Test-Throws { Read-MirrorDocument -Text $badReserved | Out-Null }) "Windows reserved name is rejected"

    $smallRoot = Join-Path $fx "binary-limits"
    New-Item -ItemType Directory -Force $smallRoot | Out-Null
    Write-Bytes (Join-Path $smallRoot "a.png") ([byte[]](137,80,78,71,13,10,26,10,1,1,1,1))
    Write-Bytes (Join-Path $smallRoot "b.png") ([byte[]](137,80,78,71,13,10,26,10,2,2,2,2))
    $limitPkg = New-MirrorPackage -Target $smallRoot -BinaryFileLimitBytes 64 -BinaryTotalLimitBytes 12
    $limitMirror = Read-MirrorDocument -Text $limitPkg.Document
    Assert-True (@($limitMirror.Entries | Where-Object { $_.Kind -eq "binary" -and $_.Embedded }).Count -eq 1) "binary total limit is enforced"
    Assert-True (@($limitPkg.Warnings | Where-Object { $_ -like "*total limit*" }).Count -eq 1) "binary total-limit warning is emitted"

    Write-Bytes (Join-Path $smallRoot "large.png") ([byte[]](137,80,78,71,13,10,26,10,3,3,3,3,3,3,3,3,3))
    $fileLimitPkg = New-MirrorPackage -Target $smallRoot -BinaryFileLimitBytes 16 -BinaryTotalLimitBytes 100
    $fileLimitMirror = Read-MirrorDocument -Text $fileLimitPkg.Document
    $largeEntry = @($fileLimitMirror.Entries | Where-Object { $_.Path -eq "large.png" })[0]
    Assert-True (-not $largeEntry.Embedded -and $largeEntry.Reason -eq "file-limit") "binary per-file limit is enforced"

    $inventoryRoot = Join-Path $fx "large-inventory"
    New-Item -ItemType Directory -Force $inventoryRoot | Out-Null
    $largeVideoPath = Join-Path $inventoryRoot "large.mp4"
    $largeVideo = [IO.File]::Open($largeVideoPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    try {
        $largeVideo.Write($video, 0, $video.Length)
        $largeVideo.SetLength(8MB)
    } finally {
        $largeVideo.Dispose()
    }
    $inventoryPackage = New-MirrorPackage -Target $inventoryRoot -BinaryFileLimitBytes 1024 -BinaryTotalLimitBytes 2048
    $inventoryMirror = Read-MirrorDocument -Text $inventoryPackage.Document
    $inventoryEntry = @($inventoryMirror.Entries | Where-Object { $_.Path -eq "large.mp4" })[0]
    $inventoryDigest = Get-MirrorFileDigest $largeVideoPath
    Assert-True (-not $inventoryEntry.Embedded -and $inventoryEntry.Bytes -eq 8MB) "large inventory-only binary records its size"
    Assert-True ($inventoryEntry.Sha256 -ceq $inventoryDigest.Sha256) "large inventory-only binary is hashed by the streaming path"
    $formatText = [IO.File]::ReadAllText($format)
    Assert-True ($formatText -match 'Get-MirrorFilePrefix' -and $formatText -match 'Get-MirrorFileDigest') "binary inventory uses prefix classification and streaming digest helpers"

    $restoreCandidate = Get-MirrorRestoreDestination -OutputRoot (Join-Path $fx "fixed-output")
    Assert-True ((Split-Path $restoreCandidate -Leaf) -like "md-mirror_restore_*") "restore output uses the fixed timestamp name"

    $asciiFiles = @(Get-ChildItem (Join-Path $root "tool\md-mirror") -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in @(".ps1", ".cs", ".json") })
    foreach ($file in $asciiFiles) {
        $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
        Assert-True (@($bytes | Where-Object { $_ -gt 127 }).Count -eq 0) ("ASCII source: " + $file.Name)
    }
} finally {
    Remove-Item -LiteralPath $fx -Recurse -Force -ErrorAction SilentlyContinue
}

Exit-Test
