# MD Patch preflight and transactional apply engine. PowerShell 5.1, ASCII only.

$script:PatchMaxPath = 259

function New-PatchUtf8Encoding {
    param([bool]$Bom = $false)
    return (New-Object Text.UTF8Encoding($Bom, $true))
}

function New-PatchUtf16Encoding {
    return (New-Object Text.UnicodeEncoding($false, $true, $true))
}

function New-PatchCp932Encoding {
    return [Text.Encoding]::GetEncoding(932, [Text.EncoderFallback]::ExceptionFallback, [Text.DecoderFallback]::ExceptionFallback)
}

function Join-PatchByteArrays {
    param([byte[]]$First, [byte[]]$Second)
    $result = New-Object byte[] ($First.Length + $Second.Length)
    [Array]::Copy($First, 0, $result, 0, $First.Length)
    [Array]::Copy($Second, 0, $result, $First.Length, $Second.Length)
    return ,$result
}

function Test-PatchByteArraysEqual {
    param([byte[]]$Left, [byte[]]$Right)
    if ($Left.Length -ne $Right.Length) { return $false }
    for ($i = 0; $i -lt $Left.Length; $i++) { if ($Left[$i] -ne $Right[$i]) { return $false } }
    return $true
}

function Get-PatchFileSha256 {
    param([string]$Path)
    return (Get-PatchSha256 ([IO.File]::ReadAllBytes($Path)))
}

function Test-PatchTextCharacters {
    param([string]$Text)
    foreach ($ch in $Text.ToCharArray()) {
        $n = [int]$ch
        if (($n -lt 32 -and $n -notin @(9, 10, 12, 13)) -or $n -eq 127) { return $false }
    }
    return $true
}

function Test-PatchPrefixBytes {
    param([byte[]]$Bytes, [byte[]]$Prefix)
    if ($Bytes.Length -lt $Prefix.Length) { return $false }
    for ($i = 0; $i -lt $Prefix.Length; $i++) { if ($Bytes[$i] -ne $Prefix[$i]) { return $false } }
    return $true
}

function ConvertFrom-PatchTextBytes {
    param([byte[]]$Bytes, [string]$Path)
    $candidates = New-Object System.Collections.Generic.List[object]
    if (Test-PatchPrefixBytes $Bytes ([byte[]](239,187,191))) {
        $payload = New-Object byte[] ($Bytes.Length - 3)
        if ($payload.Length -gt 0) { [Array]::Copy($Bytes, 3, $payload, 0, $payload.Length) }
        $candidates.Add([pscustomobject]@{ Name = "utf-8-bom"; Encoding = (New-PatchUtf8Encoding $true); Payload = $payload; Preamble = [byte[]](239,187,191) })
    } elseif (Test-PatchPrefixBytes $Bytes ([byte[]](255,254))) {
        $payload = New-Object byte[] ($Bytes.Length - 2)
        if ($payload.Length -gt 0) { [Array]::Copy($Bytes, 2, $payload, 0, $payload.Length) }
        $candidates.Add([pscustomobject]@{ Name = "utf-16le"; Encoding = (New-PatchUtf16Encoding); Payload = $payload; Preamble = [byte[]](255,254) })
    } else {
        $candidates.Add([pscustomobject]@{ Name = "utf-8"; Encoding = (New-PatchUtf8Encoding $false); Payload = $Bytes; Preamble = [byte[]]@() })
        $candidates.Add([pscustomobject]@{ Name = "cp932"; Encoding = (New-PatchCp932Encoding); Payload = $Bytes; Preamble = [byte[]]@() })
    }
    foreach ($candidate in $candidates) {
        try {
            $text = $candidate.Encoding.GetString($candidate.Payload)
            if (-not (Test-PatchTextCharacters $text)) { continue }
            $roundTrip = Join-PatchByteArrays $candidate.Preamble ($candidate.Encoding.GetBytes($text))
            if (-not (Test-PatchByteArraysEqual $Bytes $roundTrip)) { continue }
            $first = [regex]::Match($text, "`r`n|`r|`n")
            $eol = "none"
            if ($first.Success) { if ($first.Value -eq "`r`n") { $eol = "crlf" } else { $eol = "lf" } }
            $normalized = $text.Replace("`r`n", "`n").Replace("`r", "`n")
            $final = $normalized.EndsWith("`n")
            if ($final) { $body = $normalized.Substring(0, $normalized.Length - 1) } else { $body = $normalized }
            if ($normalized.Length -eq 0) {
                $lines = @()
            } elseif ($body.Length -eq 0) {
                $lines = if ($final) { @([string]"") } else { @() }
            } else {
                $lines = [regex]::Split($body, "`n")
            }
            return [pscustomobject]@{
                Encoding = $candidate.Name; Eol = $eol; FinalNewline = $final
                Lines = [object[]]$lines; OriginalBytes = $Bytes
            }
        } catch { }
    }
    throw "File is not valid supported text: $Path"
}

function ConvertTo-PatchTextBytes {
    param(
        [object[]]$Lines,
        [string]$Encoding,
        [string]$Eol,
        [bool]$FinalNewline
    )
    if (@("utf-8", "utf-8-bom", "utf-16le", "cp932") -cnotcontains $Encoding) { throw "Unsupported encoding: $Encoding" }
    if (@("lf", "crlf", "none") -cnotcontains $Eol) { throw "Unsupported EOL: $Eol" }
    if ($Lines.Count -eq 0 -and $FinalNewline) { throw "An empty file cannot have a final newline." }
    if ($Eol -eq "none" -and ($Lines.Count -gt 1 -or $FinalNewline)) { throw "EOL none cannot represent this content." }
    $text = $Lines -join "`n"
    if ($FinalNewline) { $text += "`n" }
    if ($Eol -eq "crlf") { $text = $text.Replace("`n", "`r`n") }
    switch ($Encoding) {
        "utf-8" { return ,((New-PatchUtf8Encoding $false).GetBytes($text)) }
        "utf-8-bom" {
            $enc = New-PatchUtf8Encoding $true
            return ,(Join-PatchByteArrays ($enc.GetPreamble()) ($enc.GetBytes($text)))
        }
        "utf-16le" {
            $enc = New-PatchUtf16Encoding
            return ,(Join-PatchByteArrays ($enc.GetPreamble()) ($enc.GetBytes($text)))
        }
        "cp932" { return ,((New-PatchCp932Encoding).GetBytes($text)) }
    }
}

function Get-PatchModifiedLines {
    param([object[]]$OriginalLines, [object[]]$Changes, [string]$Path)
    $output = New-Object System.Collections.Generic.List[string]
    [int]$cursor = 1
    [int]$lastStart = 0
    foreach ($change in $Changes) {
        if ($change.OldStartLine -le $lastStart) { throw "Changes are not in strictly increasing line order: $Path" }
        $lastStart = $change.OldStartLine
        if ($change.OldStartLine -lt $cursor -or $change.OldStartLine -gt ($OriginalLines.Count + 1)) { throw "Invalid or overlapping OLD range in $Path at line $($change.OldStartLine)." }
        if (($change.OldStartLine + $change.OldLineCount - 1) -gt $OriginalLines.Count) { throw "OLD range exceeds $Path at line $($change.OldStartLine)." }
        while ($cursor -lt $change.OldStartLine) {
            $output.Add([string]$OriginalLines[$cursor - 1])
            $cursor++
        }
        for ($j = 0; $j -lt $change.OldLineCount; $j++) {
            $actual = [string]$OriginalLines[$change.OldStartLine - 1 + $j]
            $expected = [string]$change.OldLines[$j]
            if ($actual -cne $expected) { throw "OLD content mismatch in $Path at line $($change.OldStartLine + $j)." }
        }
        $expectedNewStart = $output.Count + 1
        if ($change.NewStartLine -ne $expectedNewStart) { throw "New Start Line mismatch in ${Path}: expected $expectedNewStart, got $($change.NewStartLine)." }
        foreach ($line in $change.NewLines) { $output.Add([string]$line) }
        $cursor = $change.OldStartLine + $change.OldLineCount
    }
    while ($cursor -le $OriginalLines.Count) {
        $output.Add([string]$OriginalLines[$cursor - 1])
        $cursor++
    }
    return [pscustomobject]@{ Lines = $output.ToArray() }
}

function Test-PatchReparsePath {
    param([string]$Root, [string]$FullPath)
    $rootFull = [IO.Path]::GetFullPath($Root).TrimEnd('\')
    $relative = $FullPath.Substring($rootFull.Length).TrimStart('\')
    $current = $rootFull
    if (Test-Path -LiteralPath $current) {
        $rootItem = Get-Item -LiteralPath $current -Force -ErrorAction Stop
        if (($rootItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { return $true }
    }
    foreach ($part in $relative.Split('\')) {
        if (-not $part) { continue }
        $current = Join-Path $current $part
        if (-not (Test-Path -LiteralPath $current)) { break }
        $item = Get-Item -LiteralPath $current -Force -ErrorAction Stop
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { return $true }
    }
    return $false
}

function Get-PatchFullPath {
    param([string]$Root, [string]$RelativePath)
    $rootFull = [IO.Path]::GetFullPath($Root).TrimEnd('\')
    $relative = $RelativePath.TrimEnd('/').Replace('/', '\')
    $full = [IO.Path]::GetFullPath((Join-Path $rootFull $relative))
    $prefix = $rootFull + "\"
    if (-not $full.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { throw "Patch path escapes the target folder: $RelativePath" }
    if ($full.Length -gt $script:PatchMaxPath) { throw "Patch path is too long for PowerShell 5.1: $RelativePath" }
    if (Test-PatchReparsePath $rootFull $full) { throw "Patch path crosses a reparse point: $RelativePath" }
    return $full
}

function Test-PatchFileWritableAndUnlocked {
    param([string]$Path)
    $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
    if (($item.Attributes -band [IO.FileAttributes]::ReadOnly) -ne 0) { throw "File is read-only: $Path" }
    $stream = $null
    try {
        $stream = New-Object IO.FileStream($Path, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    } catch {
        throw "File is locked or not writable: $Path"
    } finally {
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

function Prepare-PatchApplication {
    param(
        [Parameter(Mandatory = $true)]$Patch,
        [Parameter(Mandatory = $true)][string]$TargetRoot
    )
    if (-not (Test-Path -LiteralPath $TargetRoot -PathType Container)) { throw "Target folder not found: $TargetRoot" }
    $rootFull = (Get-Item -LiteralPath $TargetRoot -Force -ErrorAction Stop).FullName.TrimEnd('\')
    if (Test-PatchReparsePath $rootFull $rootFull) { throw "Target folder is a reparse point." }
    $actions = New-Object System.Collections.Generic.List[object]
    $fileActionPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $deleteFiles = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $deleteDirectories = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)

    foreach ($operation in $Patch.Operations) {
        $full = Get-PatchFullPath $rootFull $operation.Path
        $relative = $operation.Path.TrimEnd('/').Replace('/', '\')
        $type = $operation.Type
        $existing = $false
        $originalBytes = $null
        $originalSha = ""
        $newBytes = $null
        $actionKind = ""

        if ($type -in @("modify", "replace-text", "replace-binary", "delete-file")) {
            if (-not (Test-Path -LiteralPath $full -PathType Leaf)) { throw "Existing file required for ${type}: $($operation.Path)" }
            Test-PatchFileWritableAndUnlocked $full
            $existing = $true
            $originalBytes = [IO.File]::ReadAllBytes($full)
            $originalSha = Get-PatchSha256 $originalBytes
            if ($originalSha -cne $operation.ExpectedSha256) { throw "Expected SHA-256 mismatch: $($operation.Path)" }
        }

        if ($type -eq "modify") {
            $state = ConvertFrom-PatchTextBytes $originalBytes $operation.Path
            $changed = Get-PatchModifiedLines $state.Lines $operation.Changes $operation.Path
            $eol = $state.Eol
            if ($eol -eq "none" -and $changed.Lines.Count -gt 1) { $eol = "crlf" }
            $final = $state.FinalNewline
            if ($changed.Lines.Count -eq 0) { $final = $false }
            $newBytes = ConvertTo-PatchTextBytes $changed.Lines $state.Encoding $eol $final
            $actionKind = "write"
        } elseif ($type -eq "replace-text") {
            $newBytes = ConvertTo-PatchTextBytes $operation.Lines $operation.Encoding $operation.Eol $operation.FinalNewline
            $actionKind = "write"
        } elseif ($type -eq "replace-binary") {
            $newBytes = $operation.Bytes
            $actionKind = "write"
        } elseif ($type -eq "delete-file") {
            $actionKind = "delete-file"
            [void]$deleteFiles.Add($relative)
        } elseif ($type -eq "create-text" -or $type -eq "create-binary") {
            if (Test-Path -LiteralPath $full) { throw "Create path already exists: $($operation.Path)" }
            if ($type -eq "create-text") { $newBytes = ConvertTo-PatchTextBytes $operation.Lines $operation.Encoding $operation.Eol $operation.FinalNewline }
            else { $newBytes = $operation.Bytes }
            $actionKind = "create-file"
        } elseif ($type -eq "create-directory") {
            if (Test-Path -LiteralPath $full) { throw "Create directory already exists: $($operation.Path)" }
            $actionKind = "create-directory"
        } elseif ($type -eq "delete-directory") {
            if (-not (Test-Path -LiteralPath $full -PathType Container)) { throw "Delete directory not found: $($operation.Path)" }
            $actionKind = "delete-directory"
            [void]$deleteDirectories.Add($relative)
        } else {
            throw "Unsupported prepared operation: $type"
        }

        if ($actionKind -in @("write", "create-file")) { [void]$fileActionPaths.Add($relative) }
        $actions.Add([pscustomobject]@{
            Type = $type; Kind = $actionKind; Path = $operation.Path; RelativePath = $relative
            FullPath = $full; Existing = $existing; OriginalBytes = $originalBytes
            OriginalSha256 = $originalSha; NewBytes = $newBytes
        })
    }

    foreach ($action in $actions) {
        $parent = Split-Path $action.FullPath -Parent
        while ($parent -and $parent.Length -gt $rootFull.Length) {
            if (Test-Path -LiteralPath $parent -PathType Leaf) { throw "A file blocks a child path: $($action.Path)" }
            $parent = Split-Path $parent -Parent
        }
        foreach ($filePath in $fileActionPaths) {
            if ($action.RelativePath.StartsWith($filePath + "\", [StringComparison]::OrdinalIgnoreCase)) {
                throw "A planned file blocks a child path: $($action.Path)"
            }
        }
    }

    foreach ($action in @($actions | Where-Object { $_.Kind -eq "delete-directory" })) {
        $items = @(Get-ChildItem -LiteralPath $action.FullPath -Recurse -Force -ErrorAction Stop)
        foreach ($item in $items) {
            if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "Delete directory contains a reparse point: $($action.Path)" }
            $relative = $item.FullName.Substring($rootFull.Length).TrimStart('\')
            if ($item.PSIsContainer) {
                if (-not $deleteDirectories.Contains($relative)) { throw "Directory is not empty after planned deletions: $($action.Path)" }
            } else {
                if (-not $deleteFiles.Contains($relative)) { throw "Directory is not empty after planned deletions: $($action.Path)" }
            }
        }
    }

    return [pscustomobject]@{ TargetRoot = $rootFull; Actions = $actions.ToArray() }
}

function Get-PatchBackupPath {
    param([string]$BackupRoot)
    if (-not (Test-Path -LiteralPath $BackupRoot)) { New-Item -ItemType Directory -Path $BackupRoot -Force -ErrorAction Stop | Out-Null }
    $stem = "backup_" + (Get-Date -Format "yyyyMMdd_HHmmss")
    $candidate = Join-Path $BackupRoot $stem
    $suffix = 1
    while (Test-Path -LiteralPath $candidate) {
        $candidate = Join-Path $BackupRoot ($stem + "_" + $suffix)
        $suffix++
    }
    return $candidate
}

function Restore-PatchActions {
    param([object[]]$Actions, [object[]]$CreatedDirectories = @())
    $errors = New-Object System.Collections.Generic.List[string]
    foreach ($action in $Actions) {
        try {
            if ($action.Existing -and $null -ne $action.OriginalBytes) {
                $parent = Split-Path $action.FullPath -Parent
                if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
                [IO.File]::WriteAllBytes($action.FullPath, $action.OriginalBytes)
            } elseif ($action.Kind -eq "create-file" -and (Test-Path -LiteralPath $action.FullPath -PathType Leaf)) {
                Remove-Item -LiteralPath $action.FullPath -Force
            }
        } catch { $errors.Add($_.Exception.Message) }
    }
    foreach ($action in @($Actions | Where-Object { $_.Kind -eq "delete-directory" } | Sort-Object { $_.FullPath.Length })) {
        try { if (-not (Test-Path -LiteralPath $action.FullPath)) { New-Item -ItemType Directory -Path $action.FullPath -Force | Out-Null } } catch { $errors.Add($_.Exception.Message) }
    }
    foreach ($directory in @($CreatedDirectories | Sort-Object { $_.Length } -Descending)) {
        try { if (Test-Path -LiteralPath $directory -PathType Container) { [IO.Directory]::Delete($directory, $false) } } catch { $errors.Add($_.Exception.Message) }
    }
    return $errors.ToArray()
}

function New-PatchDirectoryChain {
    param([string]$Directory, [string]$Root, [System.Collections.Generic.List[string]]$CreatedDirectories)
    $missing = New-Object System.Collections.Generic.List[string]
    $current = $Directory
    while ($current -and $current.Length -gt $Root.Length -and -not (Test-Path -LiteralPath $current)) {
        $missing.Add($current)
        $current = Split-Path $current -Parent
    }
    foreach ($path in @($missing.ToArray() | Sort-Object { $_.Length })) {
        New-Item -ItemType Directory -Path $path -ErrorAction Stop | Out-Null
        $CreatedDirectories.Add($path)
    }
}

function Apply-PreparedPatch {
    param(
        [Parameter(Mandatory = $true)]$Prepared,
        [Parameter(Mandatory = $true)][string]$BackupRoot,
        [int]$FailAfterAction = 0
    )
    foreach ($action in $Prepared.Actions) {
        if ($action.Existing) {
            if (-not (Test-Path -LiteralPath $action.FullPath -PathType Leaf)) { throw "Target changed after preflight: $($action.Path)" }
            if ((Get-PatchFileSha256 $action.FullPath) -cne $action.OriginalSha256) { throw "Target changed after preflight: $($action.Path)" }
        } elseif ($action.Kind -in @("create-file", "create-directory") -and (Test-Path -LiteralPath $action.FullPath)) {
            throw "Create target appeared after preflight: $($action.Path)"
        }
    }

    $backupPath = Get-PatchBackupPath $BackupRoot
    New-Item -ItemType Directory -Path $backupPath -Force -ErrorAction Stop | Out-Null
    $backupDataPath = Join-Path $backupPath "data"
    New-Item -ItemType Directory -Path $backupDataPath -Force -ErrorAction Stop | Out-Null
    $manifest = New-Object System.Collections.Generic.List[string]
    try {
        foreach ($action in $Prepared.Actions) {
            $manifest.Add($action.Type + "`t" + $action.Path)
            if ($action.Existing -and $null -ne $action.OriginalBytes) {
                $backupFile = Join-Path $backupDataPath $action.RelativePath
                $parent = Split-Path $backupFile -Parent
                if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force -ErrorAction Stop | Out-Null }
                [IO.File]::WriteAllBytes($backupFile, $action.OriginalBytes)
            } elseif ($action.Kind -eq "delete-directory") {
                New-Item -ItemType Directory -Path (Join-Path $backupDataPath $action.RelativePath) -Force -ErrorAction Stop | Out-Null
            }
        }
        [IO.File]::WriteAllLines((Join-Path $backupPath "backup-manifest.txt"), $manifest.ToArray(), (New-PatchUtf8Encoding $false))
    } catch {
        Remove-Item -LiteralPath $backupPath -Recurse -Force -ErrorAction SilentlyContinue
        throw "Backup failed before any target write: $($_.Exception.Message)"
    }

    [int]$applied = 0
    $tempFiles = New-Object System.Collections.Generic.List[string]
    $createdDirectories = New-Object System.Collections.Generic.List[string]
    try {
        foreach ($action in @($Prepared.Actions | Where-Object { $_.Kind -eq "create-directory" } | Sort-Object { $_.FullPath.Length })) {
            New-PatchDirectoryChain $action.FullPath $Prepared.TargetRoot $createdDirectories
            $applied++
            if ($FailAfterAction -gt 0 -and $applied -ge $FailAfterAction) { throw "Forced test failure after action $applied." }
        }
        foreach ($action in @($Prepared.Actions | Where-Object { $_.Kind -in @("write", "create-file") })) {
            $parent = Split-Path $action.FullPath -Parent
            if (-not (Test-Path -LiteralPath $parent)) { New-PatchDirectoryChain $parent $Prepared.TargetRoot $createdDirectories }
            $temp = Join-Path $parent (".toolrack-" + [guid]::NewGuid().ToString("N") + ".tmp")
            $tempFiles.Add($temp)
            [IO.File]::WriteAllBytes($temp, $action.NewBytes)
            if ($action.Existing) {
                $swap = Join-Path $parent (".toolrack-old-" + [guid]::NewGuid().ToString("N") + ".tmp")
                $tempFiles.Add($swap)
                [IO.File]::Replace($temp, $action.FullPath, $swap)
                Remove-Item -LiteralPath $swap -Force -ErrorAction Stop
                [void]$tempFiles.Remove($swap)
            } else {
                [IO.File]::Move($temp, $action.FullPath)
            }
            [void]$tempFiles.Remove($temp)
            $applied++
            if ($FailAfterAction -gt 0 -and $applied -ge $FailAfterAction) { throw "Forced test failure after action $applied." }
        }
        foreach ($action in @($Prepared.Actions | Where-Object { $_.Kind -eq "delete-file" })) {
            Remove-Item -LiteralPath $action.FullPath -Force -ErrorAction Stop
            $applied++
            if ($FailAfterAction -gt 0 -and $applied -ge $FailAfterAction) { throw "Forced test failure after action $applied." }
        }
        foreach ($action in @($Prepared.Actions | Where-Object { $_.Kind -eq "delete-directory" } | Sort-Object { $_.FullPath.Length } -Descending)) {
            [IO.Directory]::Delete($action.FullPath, $false)
            $applied++
            if ($FailAfterAction -gt 0 -and $applied -ge $FailAfterAction) { throw "Forced test failure after action $applied." }
        }
    } catch {
        foreach ($temp in $tempFiles) { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
        $applyMessage = $_.Exception.Message
        $rollbackErrors = @(Restore-PatchActions $Prepared.Actions $createdDirectories.ToArray())
        if ($rollbackErrors.Count -gt 0) { throw "Patch failed: $applyMessage; rollback errors: $($rollbackErrors -join '; ')" }
        throw "Patch failed and was rolled back: $applyMessage"
    }
    return [pscustomobject]@{ Applied = $applied; BackupPath = $backupPath }
}
