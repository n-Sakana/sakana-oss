# test/test-md-patch.ps1 -- strict patch format, preflight, backup, and clipboard tests
. (Join-Path $PSScriptRoot "_assert.ps1")
$ErrorActionPreference = "Stop"
$root = Split-Path $PSScriptRoot -Parent
$toolDir = Join-Path $root "tool\md-patch"
$format = Join-Path $toolDir "format.ps1"
$engine = Join-Path $toolDir "engine.ps1"
$main = Join-Path $toolDir "main.ps1"

Assert-True (Test-Path -LiteralPath $format -PathType Leaf) "md-patch format.ps1 exists"
Assert-True (Test-Path -LiteralPath $engine -PathType Leaf) "md-patch engine.ps1 exists"
Assert-True (Test-Path -LiteralPath $main -PathType Leaf) "md-patch main.ps1 exists"
if (-not (Test-Path -LiteralPath $format) -or -not (Test-Path -LiteralPath $engine) -or -not (Test-Path -LiteralPath $main)) { Exit-Test }
. $format
. $engine

$fx = Join-Path $env:TEMP ("toolrack_mp_" + [guid]::NewGuid().ToString("N"))
$target = Join-Path $fx "target"
$backupRoot = Join-Path $fx "backups"
New-Item -ItemType Directory -Force $target, $backupRoot | Out-Null

function Write-TestBytes {
    param([string]$Path, [byte[]]$Bytes)
    $parent = Split-Path $Path -Parent
    if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Force $parent | Out-Null }
    [System.IO.File]::WriteAllBytes($Path, $Bytes)
}

function Test-TestBytesEqual {
    param([string]$Left, [string]$Right)
    if (-not (Test-Path -LiteralPath $Left) -or -not (Test-Path -LiteralPath $Right)) { return $false }
    $a = [System.IO.File]::ReadAllBytes($Left)
    $b = [System.IO.File]::ReadAllBytes($Right)
    if ($a.Length -ne $b.Length) { return $false }
    for ($i = 0; $i -lt $a.Length; $i++) { if ($a[$i] -ne $b[$i]) { return $false } }
    return $true
}

function Test-PatchThrows {
    param([scriptblock]$Action)
    try { & $Action; return $false } catch { return $true }
}

function Add-TestLine {
    param([System.Text.StringBuilder]$Builder, [string]$Line)
    [void]$Builder.Append($Line + "`n")
}

function New-MainPatchText {
    param([string]$CodeHash, [string]$MemoHash, [string]$ObsoleteHash, [string]$AssetHash, [byte[]]$Png)
    $base64 = [Convert]::ToBase64String($Png)
    $b = New-Object System.Text.StringBuilder
    Add-TestLine $b "# MD Patch v1"
    Add-TestLine $b "Operation Count: 6"
    Add-TestLine $b ""
    Add-TestLine $b "===== MODIFY FILE: code.txt ====="
    Add-TestLine $b ("Expected SHA-256: " + $CodeHash)
    Add-TestLine $b "Change Count: 2"
    Add-TestLine $b "----- CHANGE 1 -----"
    Add-TestLine $b "Old Start Line: 2"
    Add-TestLine $b "Old Line Count: 1"
    Add-TestLine $b "New Start Line: 2"
    Add-TestLine $b "New Line Count: 2"
    Add-TestLine $b "----- OLD CONTENT: 1 LINES -----"
    Add-TestLine $b "old"
    Add-TestLine $b "----- NEW CONTENT: 2 LINES -----"
    Add-TestLine $b "new one"
    Add-TestLine $b "new two"
    Add-TestLine $b "----- END CHANGE 1 -----"
    Add-TestLine $b "----- CHANGE 2 -----"
    Add-TestLine $b "Old Start Line: 3"
    Add-TestLine $b "Old Line Count: 1"
    Add-TestLine $b "New Start Line: 4"
    Add-TestLine $b "New Line Count: 1"
    Add-TestLine $b "----- OLD CONTENT: 1 LINES -----"
    Add-TestLine $b "----- NEW CONTENT: 1 LINES -----"
    Add-TestLine $b "----- NEW CONTENT: 1 LINES -----"
    Add-TestLine $b "safe content"
    Add-TestLine $b "----- END CHANGE 2 -----"
    Add-TestLine $b "===== END MODIFY FILE: code.txt ====="
    Add-TestLine $b ""
    Add-TestLine $b "===== MODIFY FILE: memo.txt ====="
    Add-TestLine $b ("Expected SHA-256: " + $MemoHash)
    Add-TestLine $b "Change Count: 1"
    Add-TestLine $b "----- CHANGE 1 -----"
    Add-TestLine $b "Old Start Line: 1"
    Add-TestLine $b "Old Line Count: 1"
    Add-TestLine $b "New Start Line: 1"
    Add-TestLine $b "New Line Count: 1"
    Add-TestLine $b "----- OLD CONTENT: 1 LINES -----"
    Add-TestLine $b (([char]0x53E4).ToString() + [char]0x3044)
    Add-TestLine $b "----- NEW CONTENT: 1 LINES -----"
    Add-TestLine $b (([char]0x65B0).ToString() + [char]0x3057 + [char]0x3044)
    Add-TestLine $b "----- END CHANGE 1 -----"
    Add-TestLine $b "===== END MODIFY FILE: memo.txt ====="
    Add-TestLine $b ""
    Add-TestLine $b "===== CREATE TEXT FILE: new/deep.txt ====="
    Add-TestLine $b "Encoding: utf-8"
    Add-TestLine $b "EOL: crlf"
    Add-TestLine $b "Final Newline: no"
    Add-TestLine $b "New Line Count: 2"
    Add-TestLine $b "----- NEW CONTENT: 2 LINES -----"
    Add-TestLine $b "created"
    Add-TestLine $b "file"
    Add-TestLine $b "===== END CREATE TEXT FILE: new/deep.txt ====="
    Add-TestLine $b ""
    Add-TestLine $b "===== CREATE DIRECTORY: empty/ ====="
    Add-TestLine $b "===== END CREATE DIRECTORY: empty/ ====="
    Add-TestLine $b ""
    Add-TestLine $b "===== DELETE FILE: obsolete.txt ====="
    Add-TestLine $b ("Expected SHA-256: " + $ObsoleteHash)
    Add-TestLine $b "===== END DELETE FILE: obsolete.txt ====="
    Add-TestLine $b ""
    Add-TestLine $b "===== REPLACE BINARY FILE: asset.txt ====="
    Add-TestLine $b ("Expected SHA-256: " + $AssetHash)
    Add-TestLine $b "Bytes: -"
    Add-TestLine $b "SHA-256: -"
    Add-TestLine $b "Base64 Line Count: 1"
    Add-TestLine $b "----- BASE64 CONTENT: 1 LINES -----"
    Add-TestLine $b $base64
    Add-TestLine $b "===== END REPLACE BINARY FILE: asset.txt ====="
    Add-TestLine $b ""
    Add-TestLine $b "===== END PATCH ====="
    return $b.ToString()
}

function New-SingleModifyPatch {
    param([string]$Path, [string]$Hash, [string]$Old, [string]$New)
    $b = New-Object System.Text.StringBuilder
    Add-TestLine $b "# MD Patch v1"
    Add-TestLine $b "Operation Count: 1"
    Add-TestLine $b ""
    Add-TestLine $b ("===== MODIFY FILE: " + $Path + " =====")
    Add-TestLine $b ("Expected SHA-256: " + $Hash)
    Add-TestLine $b "Change Count: 1"
    Add-TestLine $b "----- CHANGE 1 -----"
    Add-TestLine $b "Old Start Line: 1"
    Add-TestLine $b "Old Line Count: 1"
    Add-TestLine $b "New Start Line: 1"
    Add-TestLine $b "New Line Count: 1"
    Add-TestLine $b "----- OLD CONTENT: 1 LINES -----"
    Add-TestLine $b $Old
    Add-TestLine $b "----- NEW CONTENT: 1 LINES -----"
    Add-TestLine $b $New
    Add-TestLine $b "----- END CHANGE 1 -----"
    Add-TestLine $b ("===== END MODIFY FILE: " + $Path + " =====")
    Add-TestLine $b ""
    Add-TestLine $b "===== END PATCH ====="
    return $b.ToString()
}

function New-SingleSimplePatch {
    param([string]$Start, [string]$End)
    $b = New-Object System.Text.StringBuilder
    Add-TestLine $b "# MD Patch v1"
    Add-TestLine $b "Operation Count: 1"
    Add-TestLine $b ""
    Add-TestLine $b $Start
    Add-TestLine $b $End
    Add-TestLine $b ""
    Add-TestLine $b "===== END PATCH ====="
    return $b.ToString()
}

function New-SingleCreateTextPatch {
    param([string]$Path, [string]$Content)
    $b = New-Object System.Text.StringBuilder
    Add-TestLine $b "# MD Patch v1"
    Add-TestLine $b "Operation Count: 1"
    Add-TestLine $b ""
    Add-TestLine $b ("===== CREATE TEXT FILE: " + $Path + " =====")
    Add-TestLine $b "Encoding: utf-8"
    Add-TestLine $b "EOL: none"
    Add-TestLine $b "Final Newline: no"
    Add-TestLine $b "New Line Count: 1"
    Add-TestLine $b "----- NEW CONTENT: 1 LINES -----"
    Add-TestLine $b $Content
    Add-TestLine $b ("===== END CREATE TEXT FILE: " + $Path + " =====")
    Add-TestLine $b ""
    Add-TestLine $b "===== END PATCH ====="
    return $b.ToString()
}

try {
    $utf8Bom = New-Object System.Text.UTF8Encoding($true, $true)
    $utf8 = New-Object System.Text.UTF8Encoding($false, $true)
    $cp932 = [System.Text.Encoding]::GetEncoding(932, [Text.EncoderFallback]::ExceptionFallback, [Text.DecoderFallback]::ExceptionFallback)

    $codePath = Join-Path $target "code.txt"
    $memoPath = Join-Path $target "memo.txt"
    $obsoletePath = Join-Path $target "obsolete.txt"
    $assetPath = Join-Path $target "asset.txt"
    Write-TestBytes $codePath (Join-PatchByteArrays ($utf8Bom.GetPreamble()) ($utf8Bom.GetBytes("first`r`nold`r`n----- NEW CONTENT: 1 LINES -----`r`nlast`r`n")))
    Write-TestBytes $memoPath ($cp932.GetBytes((([char]0x53E4).ToString() + [char]0x3044) + "`r`n"))
    Write-TestBytes $obsoletePath ($utf8.GetBytes("remove me`n"))
    Write-TestBytes $assetPath ($utf8.GetBytes("was text`n"))

    $codeOriginal = Join-Path $fx "code-original.bin"; [IO.File]::Copy($codePath, $codeOriginal)
    $memoOriginal = Join-Path $fx "memo-original.bin"; [IO.File]::Copy($memoPath, $memoOriginal)
    $obsoleteOriginal = Join-Path $fx "obsolete-original.bin"; [IO.File]::Copy($obsoletePath, $obsoleteOriginal)
    $assetOriginal = Join-Path $fx "asset-original.bin"; [IO.File]::Copy($assetPath, $assetOriginal)
    $png = [byte[]](137,80,78,71,13,10,26,10,7,8,9)

    $patchText = New-MainPatchText (Get-PatchFileSha256 $codePath) (Get-PatchFileSha256 $memoPath) (Get-PatchFileSha256 $obsoletePath) (Get-PatchFileSha256 $assetPath) $png
    Assert-True ($patchText -notlike "*```*") "patch transport does not use code fences"
    $patch = Read-PatchDocument -Text $patchText
    Assert-True ($patch.Operations.Count -eq 6) "strict parser reads all operations"
    $modify = @($patch.Operations | Where-Object { $_.Type -eq "modify" -and $_.Path -eq "code.txt" })[0]
    Assert-True ($modify.Changes.Count -eq 2) "multiple changes in one file are parsed"
    Assert-True ($modify.Changes[1].OldLines[0] -eq "----- NEW CONTENT: 1 LINES -----") "declared counts protect structural-looking source lines"

    $prepared = Prepare-PatchApplication -Patch $patch -TargetRoot $target
    Assert-True ($prepared.Actions.Count -eq 6) "all operations pass preflight"
    $result = Apply-PreparedPatch -Prepared $prepared -BackupRoot $backupRoot
    Assert-True ($result.Applied -eq 6) "all operations applied"
    Assert-True (([IO.File]::ReadAllText($codePath, $utf8Bom)) -eq "first`r`nnew one`r`nnew two`r`nsafe content`r`nlast`r`n") "line-numbered changes applied with BOM and CRLF preserved"
    Assert-True (([IO.File]::ReadAllText($memoPath, $cp932)) -eq (([char]0x65B0).ToString() + [char]0x3057 + [char]0x3044 + "`r`n")) "CP932 change stays CP932"
    Assert-True ([IO.File]::ReadAllText((Join-Path $target "new\deep.txt"), $utf8) -eq "created`r`nfile") "new text file created with declared EOL and final-newline state"
    Assert-True (Test-Path -LiteralPath (Join-Path $target "empty") -PathType Container) "empty directory created"
    Assert-True (-not (Test-Path -LiteralPath $obsoletePath)) "file deletion applied"
    Assert-True ((Get-PatchFileSha256 $assetPath) -eq (Get-PatchSha256 $png)) "text-to-binary replacement applied"
    Assert-True (Test-TestBytesEqual $codeOriginal (Join-Path $result.BackupPath "data\code.txt")) "modified file backup is byte-identical"
    Assert-True (Test-TestBytesEqual $memoOriginal (Join-Path $result.BackupPath "data\memo.txt")) "CP932 backup is byte-identical"
    Assert-True (Test-TestBytesEqual $obsoleteOriginal (Join-Path $result.BackupPath "data\obsolete.txt")) "deleted file backup is byte-identical"
    Assert-True (Test-TestBytesEqual $assetOriginal (Join-Path $result.BackupPath "data\asset.txt")) "binary-replaced file backup is byte-identical"

    $mismatchText = New-SingleModifyPatch "code.txt" (Get-PatchFileSha256 $codePath) "wrong old line" "x"
    $mismatchPatch = Read-PatchDocument $mismatchText
    $codeAfter = [IO.File]::ReadAllBytes($codePath)
    Assert-True (Test-PatchThrows { Prepare-PatchApplication $mismatchPatch $target | Out-Null }) "OLD-content mismatch aborts preflight"
    Assert-True (Test-PatchByteArraysEqual $codeAfter ([IO.File]::ReadAllBytes($codePath))) "mismatch writes nothing"

    $badPathText = $mismatchText.Replace("code.txt", "../escape.txt")
    Assert-True (Test-PatchThrows { Read-PatchDocument $badPathText | Out-Null }) "relative-path attack is rejected"
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $fx "escape.txt"))) "relative-path attack creates nothing"

    $readonly = Join-Path $target "readonly.txt"
    Write-TestBytes $readonly ($utf8.GetBytes("old"))
    $readPatch = Read-PatchDocument (New-SingleModifyPatch "readonly.txt" (Get-PatchFileSha256 $readonly) "old" "new")
    [IO.File]::SetAttributes($readonly, [IO.FileAttributes]::ReadOnly)
    try {
        Assert-True (Test-PatchThrows { Prepare-PatchApplication $readPatch $target | Out-Null }) "read-only target is rejected before writes"
    } finally {
        [IO.File]::SetAttributes($readonly, [IO.FileAttributes]::Normal)
    }
    Assert-True ([IO.File]::ReadAllText($readonly, $utf8) -eq "old") "read-only rejection leaves content unchanged"

    $locked = Join-Path $target "locked.txt"
    Write-TestBytes $locked ($utf8.GetBytes("locked"))
    $lockedPatch = Read-PatchDocument (New-SingleModifyPatch "locked.txt" (Get-PatchFileSha256 $locked) "locked" "changed")
    $lockStream = New-Object IO.FileStream($locked, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    try {
        Assert-True (Test-PatchThrows { Prepare-PatchApplication $lockedPatch $target | Out-Null }) "locked target is rejected before writes"
    } finally {
        $lockStream.Dispose()
    }
    Assert-True ([IO.File]::ReadAllText($locked, $utf8) -eq "locked") "locked-file rejection leaves content unchanged"

    $collisionFile = Join-Path $target "dir-collision"
    Write-TestBytes $collisionFile ($utf8.GetBytes("file"))
    $dirCollisionText = New-SingleSimplePatch "===== CREATE DIRECTORY: dir-collision/ =====" "===== END CREATE DIRECTORY: dir-collision/ ====="
    $dirCollisionPatch = Read-PatchDocument $dirCollisionText
    Assert-True (Test-PatchThrows { Prepare-PatchApplication $dirCollisionPatch $target | Out-Null }) "directory entry colliding with an existing file is rejected"

    $longName = ("a" * 245) + ".txt"
    $longPatch = Read-PatchDocument (New-SingleCreateTextPatch $longName "x")
    Assert-True (Test-PatchThrows { Prepare-PatchApplication $longPatch $target | Out-Null }) "path beyond the PowerShell 5.1 limit is rejected before writes"

    $outside = Join-Path $fx "junction-outside"
    $junction = Join-Path $target "junction"
    New-Item -ItemType Directory $outside | Out-Null
    $junctionMade = $false
    try {
        New-Item -ItemType Junction -Path $junction -Target $outside -ErrorAction Stop | Out-Null
        $junctionMade = $true
    } catch {
        Write-Host "  skip: junction creation unavailable" -ForegroundColor Yellow
    }
    if ($junctionMade) {
        $junctionPatch = Read-PatchDocument (New-SingleCreateTextPatch "junction/escape.txt" "x")
        Assert-True (Test-PatchThrows { Prepare-PatchApplication $junctionPatch $target | Out-Null }) "junction traversal is rejected"
        Assert-True (-not (Test-Path -LiteralPath (Join-Path $outside "escape.txt"))) "junction rejection writes nothing outside the target"
        Remove-Item -LiteralPath $junction -Force
    }

    $deleteDir = Join-Path $target "delete-empty"
    New-Item -ItemType Directory $deleteDir | Out-Null
    $deleteDirText = New-SingleSimplePatch "===== DELETE DIRECTORY: delete-empty/ =====" "===== END DELETE DIRECTORY: delete-empty/ ====="
    $deleteDirPatch = Read-PatchDocument $deleteDirText
    $deletePrepared = Prepare-PatchApplication $deleteDirPatch $target
    $deleteResult = Apply-PreparedPatch $deletePrepared $backupRoot
    Assert-True (-not (Test-Path -LiteralPath $deleteDir)) "empty directory deletion is applied"
    Assert-True (Test-Path -LiteralPath (Join-Path $deleteResult.BackupPath "data\delete-empty") -PathType Container) "deleted empty directory is represented in the backup"

    $unencodable = ([char]::ConvertFromUtf32(0x1F600))
    $badEncodingText = New-SingleModifyPatch "memo.txt" (Get-PatchFileSha256 $memoPath) (([char]0x65B0).ToString() + [char]0x3057 + [char]0x3044) $unencodable
    $badEncodingPatch = Read-PatchDocument $badEncodingText
    $memoBefore = [IO.File]::ReadAllBytes($memoPath)
    Assert-True (Test-PatchThrows { Prepare-PatchApplication $badEncodingPatch $target | Out-Null }) "unencodable CP932 change aborts the entire patch"
    Assert-True (Test-PatchByteArraysEqual $memoBefore ([IO.File]::ReadAllBytes($memoPath))) "encoding failure writes nothing"

    $rollbackTarget = Join-Path $fx "rollback"
    New-Item -ItemType Directory $rollbackTarget | Out-Null
    Write-TestBytes (Join-Path $rollbackTarget "one.txt") ($utf8.GetBytes("one"))
    Write-TestBytes (Join-Path $rollbackTarget "two.txt") ($utf8.GetBytes("two"))
    $p1 = New-SingleModifyPatch "one.txt" (Get-PatchFileSha256 (Join-Path $rollbackTarget "one.txt")) "one" "ONE"
    $p2 = New-SingleModifyPatch "two.txt" (Get-PatchFileSha256 (Join-Path $rollbackTarget "two.txt")) "two" "TWO"
    $combined = $p1.Replace("Operation Count: 1", "Operation Count: 2").Replace("`n===== END PATCH =====`n", "`n")
    $secondBody = $p2.Substring($p2.IndexOf("===== MODIFY FILE:"))
    $combined += $secondBody
    $rollbackPatch = Read-PatchDocument $combined
    $rollbackPrepared = Prepare-PatchApplication $rollbackPatch $rollbackTarget
    Assert-True (Test-PatchThrows { Apply-PreparedPatch $rollbackPrepared $backupRoot -FailAfterAction 1 | Out-Null }) "forced mid-apply failure is surfaced"
    Assert-True ([IO.File]::ReadAllText((Join-Path $rollbackTarget "one.txt"), $utf8) -eq "one") "rollback restores the first modified file"
    Assert-True ([IO.File]::ReadAllText((Join-Path $rollbackTarget "two.txt"), $utf8) -eq "two") "rollback preserves the untouched file"

    $newRollbackTarget = Join-Path $fx "rollback-new"
    New-Item -ItemType Directory $newRollbackTarget | Out-Null
    Write-TestBytes (Join-Path $newRollbackTarget "existing.txt") ($utf8.GetBytes("old"))
    $createPart = New-SingleCreateTextPatch "new/tree/file.txt" "created"
    $modifyPart = New-SingleModifyPatch "existing.txt" (Get-PatchFileSha256 (Join-Path $newRollbackTarget "existing.txt")) "old" "new"
    $newRollbackText = $createPart.Replace("Operation Count: 1", "Operation Count: 2").Replace("`n===== END PATCH =====`n", "`n")
    $newRollbackText += $modifyPart.Substring($modifyPart.IndexOf("===== MODIFY FILE:"))
    $newRollbackPatch = Read-PatchDocument $newRollbackText
    $newRollbackPrepared = Prepare-PatchApplication $newRollbackPatch $newRollbackTarget
    Assert-True (Test-PatchThrows { Apply-PreparedPatch $newRollbackPrepared $backupRoot -FailAfterAction 1 | Out-Null }) "forced failure after nested create is surfaced"
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $newRollbackTarget "new"))) "rollback removes implicitly created parent directories"
    Assert-True ([IO.File]::ReadAllText((Join-Path $newRollbackTarget "existing.txt"), $utf8) -eq "old") "nested-create rollback preserves other files"

    $clipTarget = Join-Path $fx "clipboard-target"
    New-Item -ItemType Directory $clipTarget | Out-Null
    $clipPatchBuilder = New-Object System.Text.StringBuilder
    Add-TestLine $clipPatchBuilder "# MD Patch v1"
    Add-TestLine $clipPatchBuilder "Operation Count: 1"
    Add-TestLine $clipPatchBuilder ""
    Add-TestLine $clipPatchBuilder "===== CREATE TEXT FILE: clipboard.txt ====="
    Add-TestLine $clipPatchBuilder "Encoding: utf-8"
    Add-TestLine $clipPatchBuilder "EOL: none"
    Add-TestLine $clipPatchBuilder "Final Newline: no"
    Add-TestLine $clipPatchBuilder "New Line Count: 1"
    Add-TestLine $clipPatchBuilder "----- NEW CONTENT: 1 LINES -----"
    Add-TestLine $clipPatchBuilder "clipboard path"
    Add-TestLine $clipPatchBuilder "===== END CREATE TEXT FILE: clipboard.txt ====="
    Add-TestLine $clipPatchBuilder ""
    Add-TestLine $clipPatchBuilder "===== END PATCH ====="
    $oldClipboard = $null
    try { $oldClipboard = Get-Clipboard -Raw -Format Text -ErrorAction SilentlyContinue } catch { }
    $beforeBackups = @(Get-ChildItem (Join-Path $root "output") -Directory -Filter "backup_*" -ErrorAction SilentlyContinue)
    try {
        Set-Clipboard -Value $clipPatchBuilder.ToString()
        $env:TOOLRACK_NOPAUSE = "1"
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $main -Target $clipTarget -Source clipboard -NoOpen
        Assert-True ($LASTEXITCODE -eq 0) "real clipboard route exits successfully"
        Assert-True ([IO.File]::ReadAllText((Join-Path $clipTarget "clipboard.txt"), $utf8) -eq "clipboard path") "real Get-Clipboard -Raw route applies the patch"
    } finally {
        if ($null -ne $oldClipboard) {
            Set-Clipboard -Value $oldClipboard
        } else {
            Add-Type -AssemblyName System.Windows.Forms
            [Windows.Forms.Clipboard]::Clear()
        }
        $env:TOOLRACK_NOPAUSE = ""
        $afterBackups = @(Get-ChildItem (Join-Path $root "output") -Directory -Filter "backup_*" -ErrorAction SilentlyContinue)
        foreach ($dir in @($afterBackups | Where-Object { $beforeBackups.Name -notcontains $_.Name })) {
            Remove-Item -LiteralPath $dir.FullName -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    $asciiFiles = @(Get-ChildItem $toolDir -File | Where-Object { $_.Extension -in @(".ps1", ".cs", ".json") })
    foreach ($file in $asciiFiles) {
        $bytes = [IO.File]::ReadAllBytes($file.FullName)
        Assert-True (@($bytes | Where-Object { $_ -gt 127 }).Count -eq 0) ("ASCII source: " + $file.Name)
    }
} finally {
    if (Test-Path -LiteralPath $fx) {
        Get-ChildItem -LiteralPath $fx -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
            if (-not $_.PSIsContainer) { $_.Attributes = [IO.FileAttributes]::Normal }
        }
        Remove-Item -LiteralPath $fx -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Exit-Test
