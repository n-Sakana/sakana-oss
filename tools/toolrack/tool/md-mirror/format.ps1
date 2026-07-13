# MD Mirror v1 format implementation. PowerShell 5.1, ASCII source only.

$script:MirrorHeader = "# MD Mirror v1"
$script:MirrorDefaultBinaryFileLimit = 5MB
$script:MirrorDefaultBinaryTotalLimit = 10MB
$script:MirrorMaxPath = 259
$script:MirrorBinaryExtensions = @(
    ".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".ico", ".avif",
    ".woff", ".woff2", ".ttf", ".otf", ".eot", ".wasm", ".cur", ".icns"
)
$script:MirrorExcludedDirectories = @(".git", "node_modules", ".venv", "__pycache__")

function New-MirrorUtf8Encoding {
    param([bool]$Bom = $false)
    return (New-Object System.Text.UTF8Encoding($Bom, $true))
}

function New-MirrorUtf16Encoding {
    return (New-Object System.Text.UnicodeEncoding($false, $true, $true))
}

function New-MirrorCp932Encoding {
    return [System.Text.Encoding]::GetEncoding(
        932,
        [System.Text.EncoderFallback]::ExceptionFallback,
        [System.Text.DecoderFallback]::ExceptionFallback
    )
}

function Join-MirrorByteArrays {
    param([byte[]]$First, [byte[]]$Second)
    $result = New-Object byte[] ($First.Length + $Second.Length)
    [Array]::Copy($First, 0, $result, 0, $First.Length)
    [Array]::Copy($Second, 0, $result, $First.Length, $Second.Length)
    return ,$result
}

function Test-MirrorBytesEqual {
    param([byte[]]$Left, [byte[]]$Right)
    if ($Left.Length -ne $Right.Length) { return $false }
    for ($i = 0; $i -lt $Left.Length; $i++) {
        if ($Left[$i] -ne $Right[$i]) { return $false }
    }
    return $true
}

function Get-MirrorSha256 {
    param([byte[]]$Bytes)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = $sha.ComputeHash($Bytes)
        return (($hash | ForEach-Object { $_.ToString("x2") }) -join "")
    } finally {
        $sha.Dispose()
    }
}

function Get-MirrorFileDigest {
    param([Parameter(Mandatory = $true)][string]$Path)
    $stream = [System.IO.File]::Open(
        $Path,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::Read,
        [System.IO.FileShare]::Read
    )
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = $sha.ComputeHash($stream)
        return [pscustomobject]@{
            Length = [long]$stream.Length
            Sha256 = (($hash | ForEach-Object { $_.ToString("x2") }) -join "")
        }
    } finally {
        $sha.Dispose()
        $stream.Dispose()
    }
}

function Get-MirrorFilePrefix {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [int]$MaximumBytes = 16
    )
    if ($MaximumBytes -lt 1) { return ,([byte[]]@()) }
    $stream = [System.IO.File]::Open(
        $Path,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::Read,
        [System.IO.FileShare]::Read
    )
    try {
        $buffer = New-Object byte[] $MaximumBytes
        $read = $stream.Read($buffer, 0, $buffer.Length)
        if ($read -eq $buffer.Length) { return ,$buffer }
        $prefix = New-Object byte[] $read
        if ($read -gt 0) { [Array]::Copy($buffer, 0, $prefix, 0, $read) }
        return ,$prefix
    } finally {
        $stream.Dispose()
    }
}

function Test-MirrorControlCharacters {
    param([string]$Text)
    foreach ($ch in $Text.ToCharArray()) {
        $n = [int]$ch
        if (($n -lt 32 -and $n -notin @(9, 10, 12, 13)) -or $n -eq 127) {
            return $false
        }
    }
    return $true
}

function Test-MirrorPrefixBytes {
    param([byte[]]$Bytes, [byte[]]$Prefix)
    if ($Bytes.Length -lt $Prefix.Length) { return $false }
    for ($i = 0; $i -lt $Prefix.Length; $i++) {
        if ($Bytes[$i] -ne $Prefix[$i]) { return $false }
    }
    return $true
}

function Test-MirrorKnownBinarySignature {
    param([byte[]]$Bytes)
    if ($Bytes.Length -ge 8 -and
        $Bytes[0] -eq 137 -and $Bytes[1] -eq 80 -and $Bytes[2] -eq 78 -and $Bytes[3] -eq 71 -and
        $Bytes[4] -eq 13 -and $Bytes[5] -eq 10 -and $Bytes[6] -eq 26 -and $Bytes[7] -eq 10) { return $true }
    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 255 -and $Bytes[1] -eq 216 -and $Bytes[2] -eq 255) { return $true }
    if ($Bytes.Length -ge 6) {
        $six = [System.Text.Encoding]::ASCII.GetString($Bytes, 0, 6)
        if ($six -eq "GIF87a" -or $six -eq "GIF89a") { return $true }
    }
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 66 -and $Bytes[1] -eq 77) { return $true }
    if ($Bytes.Length -ge 12) {
        $riff = [System.Text.Encoding]::ASCII.GetString($Bytes, 0, 4)
        $webp = [System.Text.Encoding]::ASCII.GetString($Bytes, 8, 4)
        if ($riff -eq "RIFF" -and $webp -eq "WEBP") { return $true }
    }
    if ($Bytes.Length -ge 4) {
        if ($Bytes[0] -eq 0 -and $Bytes[1] -eq 97 -and $Bytes[2] -eq 115 -and $Bytes[3] -eq 109) { return $true }
        if ($Bytes[0] -eq 80 -and $Bytes[1] -eq 75 -and $Bytes[2] -in @(3,5,7) -and $Bytes[3] -in @(4,6,8)) { return $true }
        if ($Bytes[0] -eq 37 -and $Bytes[1] -eq 80 -and $Bytes[2] -eq 68 -and $Bytes[3] -eq 70) { return $true }
        if ($Bytes[0] -eq 77 -and $Bytes[1] -eq 90) { return $true }
        if ($Bytes[0] -eq 127 -and $Bytes[1] -eq 69 -and $Bytes[2] -eq 76 -and $Bytes[3] -eq 70) { return $true }
        if ($Bytes[0] -eq 208 -and $Bytes[1] -eq 207 -and $Bytes[2] -eq 17 -and $Bytes[3] -eq 224) { return $true }
        if ($Bytes[0] -eq 255 -and $Bytes[1] -eq 254 -and $Bytes[2] -eq 0 -and $Bytes[3] -eq 0) { return $true }
        if ($Bytes[0] -eq 0 -and $Bytes[1] -eq 0 -and $Bytes[2] -eq 254 -and $Bytes[3] -eq 255) { return $true }
    }
    if ($Bytes.Length -ge 8 -and [System.Text.Encoding]::ASCII.GetString($Bytes, 4, 4) -eq "ftyp") { return $true }
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 31 -and $Bytes[1] -eq 139) { return $true }
    if ($Bytes.Length -ge 7 -and $Bytes[0] -eq 82 -and $Bytes[1] -eq 97 -and $Bytes[2] -eq 114 -and
        $Bytes[3] -eq 33 -and $Bytes[4] -eq 26 -and $Bytes[5] -eq 7) { return $true }
    if ($Bytes.Length -ge 6 -and $Bytes[0] -eq 55 -and $Bytes[1] -eq 122 -and $Bytes[2] -eq 188 -and
        $Bytes[3] -eq 175 -and $Bytes[4] -eq 39 -and $Bytes[5] -eq 28) { return $true }
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 254 -and $Bytes[1] -eq 255) { return $true }
    return $false
}

function Test-MirrorSupportedBinaryExtension {
    param([string]$Path)
    $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    return ($script:MirrorBinaryExtensions -contains $ext)
}

function ConvertFrom-MirrorBytesAsText {
    param([byte[]]$Bytes)

    $candidates = New-Object System.Collections.Generic.List[object]
    if (Test-MirrorPrefixBytes $Bytes ([byte[]](239,187,191))) {
        $payload = New-Object byte[] ($Bytes.Length - 3)
        if ($payload.Length -gt 0) { [Array]::Copy($Bytes, 3, $payload, 0, $payload.Length) }
        $candidates.Add([pscustomobject]@{ Name = "utf-8-bom"; Encoding = (New-MirrorUtf8Encoding $true); Payload = $payload; Preamble = [byte[]](239,187,191) })
    } elseif (Test-MirrorPrefixBytes $Bytes ([byte[]](255,254))) {
        $payload = New-Object byte[] ($Bytes.Length - 2)
        if ($payload.Length -gt 0) { [Array]::Copy($Bytes, 2, $payload, 0, $payload.Length) }
        $candidates.Add([pscustomobject]@{ Name = "utf-16le"; Encoding = (New-MirrorUtf16Encoding); Payload = $payload; Preamble = [byte[]](255,254) })
    } else {
        $candidates.Add([pscustomobject]@{ Name = "utf-8"; Encoding = (New-MirrorUtf8Encoding $false); Payload = $Bytes; Preamble = [byte[]]@() })
        $candidates.Add([pscustomobject]@{ Name = "cp932"; Encoding = (New-MirrorCp932Encoding); Payload = $Bytes; Preamble = [byte[]]@() })
    }

    foreach ($candidate in $candidates) {
        try {
            $text = $candidate.Encoding.GetString($candidate.Payload)
            if (-not (Test-MirrorControlCharacters $text)) { continue }
            $encoded = $candidate.Encoding.GetBytes($text)
            $roundTrip = Join-MirrorByteArrays $candidate.Preamble $encoded
            if (-not (Test-MirrorBytesEqual $Bytes $roundTrip)) { continue }
            return [pscustomobject]@{ Encoding = $candidate.Name; Text = $text }
        } catch { }
    }
    return $null
}

function Get-MirrorTextPayload {
    param([string]$Text, [string]$Encoding)

    $first = [regex]::Match($Text, "`r`n|`r|`n")
    $eol = "none"
    if ($first.Success) {
        if ($first.Value -eq "`r`n") { $eol = "crlf" } else { $eol = "lf" }
    }

    $hasCrLf = $Text.Contains("`r`n")
    $withoutCrLf = $Text.Replace("`r`n", "")
    $hasLf = $withoutCrLf.Contains("`n")
    $hasCr = $withoutCrLf.Contains("`r")
    $styleCount = 0
    if ($hasCrLf) { $styleCount++ }
    if ($hasLf) { $styleCount++ }
    if ($hasCr) { $styleCount++ }
    $mixed = ($styleCount -gt 1)

    $normalized = $Text.Replace("`r`n", "`n").Replace("`r", "`n")
    $final = $normalized.EndsWith("`n")
    $body = $normalized
    if ($final) { $body = $body.Substring(0, $body.Length - 1) }
    if ($eol -eq "none" -and $final) { throw "Invalid final-newline state." }

    $bytes = ConvertTo-MirrorTextBytes -Body $body -Encoding $Encoding -Eol $eol -FinalNewline $final
    return [pscustomobject]@{
        Body = $body
        Eol = $eol
        FinalNewline = $final
        Mixed = $mixed
        Bytes = $bytes
    }
}

function ConvertTo-MirrorTextBytes {
    param(
        [string]$Body,
        [string]$Encoding,
        [string]$Eol,
        [bool]$FinalNewline
    )
    if (@("utf-8", "utf-8-bom", "utf-16le", "cp932") -notcontains $Encoding) {
        throw "Unsupported encoding: $Encoding"
    }
    if (@("lf", "crlf", "none") -notcontains $Eol) { throw "Unsupported EOL: $Eol" }
    if ($Body.Contains("`r")) { throw "Text body contains a CR character."
    }
    if ($Eol -eq "none") {
        if ($Body.Contains("`n") -or $FinalNewline) { throw "EOL none requires one non-terminated line."
        }
    }

    $text = $Body
    if ($FinalNewline) { $text += "`n" }
    if ($Eol -eq "crlf") { $text = $text.Replace("`n", "`r`n") }

    switch ($Encoding) {
        "utf-8" { return ,((New-MirrorUtf8Encoding $false).GetBytes($text)) }
        "utf-8-bom" {
            $enc = New-MirrorUtf8Encoding $true
            return ,(Join-MirrorByteArrays ($enc.GetPreamble()) ($enc.GetBytes($text)))
        }
        "utf-16le" {
            $enc = New-MirrorUtf16Encoding
            return ,(Join-MirrorByteArrays ($enc.GetPreamble()) ($enc.GetBytes($text)))
        }
        "cp932" { return ,((New-MirrorCp932Encoding).GetBytes($text)) }
    }
}

function Test-MirrorSingleLineValue {
    param([string]$Value, [string]$Name)
    if ([string]::IsNullOrWhiteSpace($Value)) { throw "$Name must not be empty." }
    foreach ($ch in $Value.ToCharArray()) {
        if ([int]$ch -lt 32 -or [int]$ch -eq 127) { throw "$Name contains a control character." }
    }
}

function Test-MirrorRelativePath {
    param([string]$Path, [string]$Kind = "")
    Test-MirrorSingleLineValue $Path "Path"
    if ($Path.Contains("\")) { throw "Path must use forward slashes: $Path" }
    if ($Path.StartsWith("/") -or [System.IO.Path]::IsPathRooted($Path)) { throw "Path must be relative: $Path" }
    if ($Path.Contains(":")) { throw "Path contains an invalid colon: $Path" }
    if ($Kind -eq "dir") {
        if (-not $Path.EndsWith("/")) { throw "Directory path must end with '/': $Path" }
        $check = $Path.Substring(0, $Path.Length - 1)
    } else {
        if ($Path.EndsWith("/")) { throw "File path must not end with '/': $Path" }
        $check = $Path
    }
    if ([string]::IsNullOrEmpty($check)) { throw "Path must not be empty." }
    $segments = $check.Split([char]'/', [System.StringSplitOptions]::None)
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($segment in $segments) {
        if ([string]::IsNullOrEmpty($segment) -or $segment -eq "." -or $segment -eq "..") {
            throw "Path contains an invalid segment: $Path"
        }
        if ($segment.EndsWith(".") -or $segment.EndsWith(" ")) { throw "Path segment ends with a dot or space: $Path" }
        if ($segment.IndexOfAny($invalid) -ge 0) { throw "Path contains an invalid character: $Path" }
        if ($segment -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?$') {
            throw "Path uses a Windows reserved name: $Path"
        }
    }
    return $true
}

function Get-MirrorSourceItems {
    param([string]$Target, [System.Collections.Generic.List[string]]$Warnings)
    if (-not (Test-Path -LiteralPath $Target)) { throw "Path not found: $Target" }
    $item = Get-Item -LiteralPath $Target -Force -ErrorAction Stop
    if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "Reparse-point targets are not supported: $Target"
    }
    $items = New-Object System.Collections.Generic.List[object]
    if (-not $item.PSIsContainer) {
        $items.Add([pscustomobject]@{ Item = $item; RelativePath = $item.Name; IsDirectory = $false })
        return $items.ToArray()
    }

    $excluded = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($name in $script:MirrorExcludedDirectories) { [void]$excluded.Add($name) }
    $walk = $null
    $walk = {
        param([System.IO.DirectoryInfo]$Directory, [string]$Prefix)
        $children = @(Get-ChildItem -LiteralPath $Directory.FullName -Force -ErrorAction Stop | Sort-Object Name)
        foreach ($child in $children) {
            $relative = if ($Prefix) { $Prefix + "/" + $child.Name } else { $child.Name }
            if (($child.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                $Warnings.Add("${relative}: reparse point skipped")
                continue
            }
            if ($child.PSIsContainer) {
                if ($excluded.Contains($child.Name)) { continue }
                $items.Add([pscustomobject]@{ Item = $child; RelativePath = $relative + "/"; IsDirectory = $true })
                & $walk $child $relative
            } else {
                $items.Add([pscustomobject]@{ Item = $child; RelativePath = $relative; IsDirectory = $false })
            }
        }
    }
    & $walk $item ""
    return $items.ToArray()
}

function New-MirrorBoundary {
    param([object[]]$Entries, [string]$Instructions, [string]$SourceName)
    while ($true) {
        $candidate = "MDM-" + [guid]::NewGuid().ToString("N").ToUpperInvariant()
        $collision = $Instructions.Contains($candidate) -or $SourceName.Contains($candidate)
        if (-not $collision) {
            foreach ($entry in $Entries) {
                if ($entry.Path.Contains($candidate) -or ($entry.PSObject.Properties.Name -contains "Body" -and $entry.Body.Contains($candidate))) {
                    $collision = $true
                    break
                }
            }
        }
        if (-not $collision) { return $candidate }
    }
}

function ConvertTo-MirrorBase64Lines {
    param([byte[]]$Bytes)
    $base64 = [Convert]::ToBase64String($Bytes)
    $parts = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $base64.Length; $i += 76) {
        $take = [Math]::Min(76, $base64.Length - $i)
        $parts.Add($base64.Substring($i, $take))
    }
    return ($parts -join "`n")
}

function New-MirrorPackage {
    param(
        [Parameter(Mandatory = $true)][string]$Target,
        [long]$BinaryFileLimitBytes = $script:MirrorDefaultBinaryFileLimit,
        [long]$BinaryTotalLimitBytes = $script:MirrorDefaultBinaryTotalLimit
    )
    if ($BinaryFileLimitBytes -lt 0 -or $BinaryTotalLimitBytes -lt 0) { throw "Binary limits must not be negative." }
    $targetItem = Get-Item -LiteralPath $Target -Force -ErrorAction Stop
    $sourceName = $targetItem.Name
    Test-MirrorSingleLineValue $sourceName "Source Name"
    $sourceType = if ($targetItem.PSIsContainer) { "folder" } else { "file" }
    $warnings = New-Object System.Collections.Generic.List[string]
    $sourceItems = @(Get-MirrorSourceItems -Target $Target -Warnings $warnings)
    $sourceItems = @($sourceItems | Sort-Object @{ Expression = { $_.RelativePath.ToLowerInvariant() } }, RelativePath)
    $entries = New-Object System.Collections.Generic.List[object]
    [long]$embeddedTotal = 0

    foreach ($source in $sourceItems) {
        $path = $source.RelativePath
        if ($source.Item.FullName.Length -gt $script:MirrorMaxPath) { throw "Path is too long for PowerShell 5.1: $path" }
        if ($source.IsDirectory) {
            [void](Test-MirrorRelativePath $path "dir")
            $entries.Add([pscustomobject]@{ Path = $path; Kind = "dir" })
            continue
        }
        [void](Test-MirrorRelativePath $path "file")
        [long]$fileLength = $source.Item.Length
        $supportedBinary = Test-MirrorSupportedBinaryExtension $path
        $knownBinary = Test-MirrorKnownBinarySignature (Get-MirrorFilePrefix $source.Item.FullName)
        $bytes = $null
        $decoded = $null
        if (-not $supportedBinary -and -not $knownBinary) {
            $bytes = [System.IO.File]::ReadAllBytes($source.Item.FullName)
            if ($bytes.Length -ne $fileLength) { throw "File changed while it was being read: $path" }
            $decoded = ConvertFrom-MirrorBytesAsText $bytes
        }

        if ($null -ne $decoded) {
            $payload = Get-MirrorTextPayload -Text $decoded.Text -Encoding $decoded.Encoding
            if ($payload.Mixed) { $warnings.Add("${path}: mixed line endings normalized to $($payload.Eol)") }
            $entries.Add([pscustomobject]@{
                Path = $path
                Kind = "text"
                Encoding = $decoded.Encoding
                Eol = $payload.Eol
                FinalNewline = $payload.FinalNewline
                OriginalBytes = [long]$bytes.Length
                OriginalSha256 = Get-MirrorSha256 $bytes
                RestoreBytes = [long]$payload.Bytes.Length
                RestoreSha256 = Get-MirrorSha256 $payload.Bytes
                Body = $payload.Body
            })
            continue
        }

        $embedded = $false
        $reason = "unsupported"
        if (-not $supportedBinary) {
            $warnings.Add("${path}: binary type not embedded")
        } elseif ($fileLength -gt $BinaryFileLimitBytes) {
            $reason = "file-limit"
            $warnings.Add("${path}: binary file limit exceeded; not embedded")
        } elseif (($embeddedTotal + $fileLength) -gt $BinaryTotalLimitBytes) {
            $reason = "total-limit"
            $warnings.Add("${path}: binary total limit exceeded; not embedded")
        } else {
            $embedded = $true
            $reason = ""
            $bytes = [System.IO.File]::ReadAllBytes($source.Item.FullName)
            if ($bytes.Length -ne $fileLength) { throw "File changed while it was being read: $path" }
            $embeddedTotal += $fileLength
        }
        $digest = if ($null -ne $bytes) {
            [pscustomobject]@{ Length = [long]$bytes.Length; Sha256 = Get-MirrorSha256 $bytes }
        } else {
            Get-MirrorFileDigest $source.Item.FullName
        }
        if ($digest.Length -ne $fileLength) { throw "File changed while it was being read: $path" }
        $entries.Add([pscustomobject]@{
            Path = $path
            Kind = "binary"
            Bytes = [long]$digest.Length
            Sha256 = $digest.Sha256
            Embedded = $embedded
            Reason = $reason
            DataBytes = if ($embedded) { $bytes } else { $null }
        })
    }

    $instructionsPath = Join-Path $PSScriptRoot "instructions.txt"
    $instructions = ""
    if (Test-Path -LiteralPath $instructionsPath -PathType Leaf) {
        $instructionBytes = [System.IO.File]::ReadAllBytes($instructionsPath)
        $instructions = (New-MirrorUtf8Encoding $false).GetString($instructionBytes)
        $instructions = $instructions.Replace("`r`n", "`n").Replace("`r", "`n").TrimEnd("`n")
    }
    $boundary = New-MirrorBoundary -Entries $entries.ToArray() -Instructions $instructions -SourceName $sourceName
    $builder = New-Object System.Text.StringBuilder
    [void]$builder.Append($script:MirrorHeader + "`n")
    [void]$builder.Append("Boundary: $boundary`n")
    [void]$builder.Append("Source Name: $sourceName`n")
    [void]$builder.Append("Source Type: $sourceType`n")
    [void]$builder.Append("Created At: " + (Get-Date -Format "yyyy-MM-ddTHH:mm:sszzz") + "`n")
    [void]$builder.Append("Entry Count: " + $entries.Count + "`n`n")
    [void]$builder.Append("===== INSTRUCTIONS $boundary =====`n")
    [void]$builder.Append($instructions)
    [void]$builder.Append("`n===== END INSTRUCTIONS $boundary =====`n`n")

    for ($i = 0; $i -lt $entries.Count; $i++) {
        $entry = $entries[$i]
        $number = $i + 1
        [void]$builder.Append("===== ENTRY $number $boundary =====`n")
        [void]$builder.Append("Path: " + $entry.Path + "`n")
        [void]$builder.Append("Kind: " + $entry.Kind + "`n")
        if ($entry.Kind -eq "text") {
            [void]$builder.Append("Encoding: " + $entry.Encoding + "`n")
            [void]$builder.Append("EOL: " + $entry.Eol + "`n")
            [void]$builder.Append("Final Newline: " + $(if ($entry.FinalNewline) { "yes" } else { "no" }) + "`n")
            [void]$builder.Append("Original Bytes: " + $entry.OriginalBytes + "`n")
            [void]$builder.Append("Original SHA-256: " + $entry.OriginalSha256 + "`n")
            [void]$builder.Append("Restore Bytes: " + $entry.RestoreBytes + "`n")
            [void]$builder.Append("Restore SHA-256: " + $entry.RestoreSha256 + "`n")
            [void]$builder.Append("===== TEXT $boundary =====`n")
            [void]$builder.Append($entry.Body)
            [void]$builder.Append("`n===== END ENTRY $number $boundary =====`n`n")
        } elseif ($entry.Kind -eq "binary") {
            [void]$builder.Append("Bytes: " + $entry.Bytes + "`n")
            [void]$builder.Append("SHA-256: " + $entry.Sha256 + "`n")
            [void]$builder.Append("Embedded: " + $(if ($entry.Embedded) { "yes" } else { "no" }) + "`n")
            if ($entry.Embedded) {
                [void]$builder.Append("===== BASE64 $boundary =====`n")
                [void]$builder.Append((ConvertTo-MirrorBase64Lines $entry.DataBytes))
                [void]$builder.Append("`n===== END ENTRY $number $boundary =====`n`n")
            } else {
                [void]$builder.Append("Reason: " + $entry.Reason + "`n")
                [void]$builder.Append("===== END ENTRY $number $boundary =====`n`n")
            }
        } else {
            [void]$builder.Append("===== END ENTRY $number $boundary =====`n`n")
        }
    }
    [void]$builder.Append("===== END MIRROR $boundary =====`n")
    return [pscustomobject]@{
        Document = $builder.ToString()
        Entries = $entries.ToArray()
        Warnings = $warnings.ToArray()
        SourceName = $sourceName
        SourceType = $sourceType
        EmbeddedBinaryBytes = $embeddedTotal
    }
}

function Read-MirrorRequiredLine {
    param([string[]]$Lines, [ref]$Index, [string]$Expected)
    if ($Index.Value -ge $Lines.Count) { throw "Line $($Index.Value + 1): expected '$Expected', reached end of input." }
    $actual = $Lines[$Index.Value]
    if ($actual -cne $Expected) { throw "Line $($Index.Value + 1): expected '$Expected', got '$actual'." }
    $Index.Value++
}

function Read-MirrorValueLine {
    param([string[]]$Lines, [ref]$Index, [string]$Prefix)
    if ($Index.Value -ge $Lines.Count) { throw "Line $($Index.Value + 1): expected '$Prefix', reached end of input." }
    $actual = $Lines[$Index.Value]
    if (-not $actual.StartsWith($Prefix, [System.StringComparison]::Ordinal)) {
        throw "Line $($Index.Value + 1): expected '$Prefix'."
    }
    $value = $actual.Substring($Prefix.Length)
    $Index.Value++
    return $value
}

function ConvertFrom-MirrorByteCount {
    param([string]$Value, [bool]$AllowDash)
    if ($AllowDash -and $Value -ceq "-") { return [long]-1 }
    [long]$number = 0
    if (-not [long]::TryParse($Value, [Globalization.NumberStyles]::None, [Globalization.CultureInfo]::InvariantCulture, [ref]$number) -or $number -lt 0) {
        throw "Invalid byte count: $Value"
    }
    return $number
}

function Test-MirrorShaValue {
    param([string]$Value, [bool]$AllowDash)
    if ($AllowDash -and $Value -ceq "-") { return $true }
    if ($Value -cnotmatch '^[0-9a-f]{64}$') { throw "Invalid SHA-256 value: $Value" }
    return $true
}

function Read-MirrorDocument {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text)
    if ($Text.Contains([char]0)) { throw "Input contains a NUL character." }
    $normalized = $Text.Replace("`r`n", "`n").Replace("`r", "`n")
    $lines = [regex]::Split($normalized, "`n")
    [int]$i = 0
    Read-MirrorRequiredLine $lines ([ref]$i) $script:MirrorHeader
    $boundary = Read-MirrorValueLine $lines ([ref]$i) "Boundary: "
    if ($boundary -cnotmatch '^MDM-[0-9A-F]{32}$') { throw "Line 2: invalid Boundary value." }
    $sourceName = Read-MirrorValueLine $lines ([ref]$i) "Source Name: "
    Test-MirrorSingleLineValue $sourceName "Source Name"
    $sourceType = Read-MirrorValueLine $lines ([ref]$i) "Source Type: "
    if (@("file", "folder", "generated") -cnotcontains $sourceType) { throw "Invalid Source Type: $sourceType" }
    $createdAt = Read-MirrorValueLine $lines ([ref]$i) "Created At: "
    if ($createdAt -cne "-") {
        [DateTimeOffset]$parsedDate = [DateTimeOffset]::MinValue
        if (-not [DateTimeOffset]::TryParseExact($createdAt, "yyyy-MM-ddTHH:mm:sszzz", [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsedDate)) {
            throw "Invalid Created At value: $createdAt"
        }
    }
    $entryCountText = Read-MirrorValueLine $lines ([ref]$i) "Entry Count: "
    [int]$entryCount = 0
    if (-not [int]::TryParse($entryCountText, [Globalization.NumberStyles]::None, [Globalization.CultureInfo]::InvariantCulture, [ref]$entryCount) -or $entryCount -lt 0) {
        throw "Invalid Entry Count: $entryCountText"
    }
    Read-MirrorRequiredLine $lines ([ref]$i) ""
    Read-MirrorRequiredLine $lines ([ref]$i) "===== INSTRUCTIONS $boundary ====="
    $instructionStart = $i
    $instructionEndMarker = "===== END INSTRUCTIONS $boundary ====="
    while ($i -lt $lines.Count -and $lines[$i] -cne $instructionEndMarker) { $i++ }
    if ($i -ge $lines.Count) { throw "Instructions end marker is missing." }
    $instructionLines = @()
    if ($i -gt $instructionStart) { $instructionLines = $lines[$instructionStart..($i - 1)] }
    $instructions = $instructionLines -join "`n"
    $i++
    Read-MirrorRequiredLine $lines ([ref]$i) ""

    $entries = New-Object System.Collections.Generic.List[object]
    $paths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    [long]$embeddedTotal = 0
    for ($number = 1; $number -le $entryCount; $number++) {
        Read-MirrorRequiredLine $lines ([ref]$i) "===== ENTRY $number $boundary ====="
        $path = Read-MirrorValueLine $lines ([ref]$i) "Path: "
        $kind = Read-MirrorValueLine $lines ([ref]$i) "Kind: "
        if (@("text", "binary", "dir") -cnotcontains $kind) { throw "Invalid Kind for ${path}: $kind" }
        [void](Test-MirrorRelativePath $path $kind)
        $canonicalPath = $path.TrimEnd('/')
        if (-not $paths.Add($canonicalPath)) { throw "Duplicate path: $path" }
        $endMarker = "===== END ENTRY $number $boundary ====="

        if ($kind -eq "text") {
            $encoding = Read-MirrorValueLine $lines ([ref]$i) "Encoding: "
            if (@("utf-8", "utf-8-bom", "utf-16le", "cp932") -cnotcontains $encoding) { throw "Invalid Encoding for ${path}: $encoding" }
            $eol = Read-MirrorValueLine $lines ([ref]$i) "EOL: "
            if (@("lf", "crlf", "none") -cnotcontains $eol) { throw "Invalid EOL for ${path}: $eol" }
            $finalText = Read-MirrorValueLine $lines ([ref]$i) "Final Newline: "
            if (@("yes", "no") -cnotcontains $finalText) { throw "Invalid Final Newline for ${path}: $finalText" }
            $final = ($finalText -ceq "yes")
            $originalByteText = Read-MirrorValueLine $lines ([ref]$i) "Original Bytes: "
            $originalByteCount = ConvertFrom-MirrorByteCount $originalByteText $true
            $originalSha = Read-MirrorValueLine $lines ([ref]$i) "Original SHA-256: "
            [void](Test-MirrorShaValue $originalSha $true)
            $restoreByteText = Read-MirrorValueLine $lines ([ref]$i) "Restore Bytes: "
            $restoreByteCount = ConvertFrom-MirrorByteCount $restoreByteText $true
            $restoreSha = Read-MirrorValueLine $lines ([ref]$i) "Restore SHA-256: "
            [void](Test-MirrorShaValue $restoreSha $true)
            Read-MirrorRequiredLine $lines ([ref]$i) "===== TEXT $boundary ====="
            $bodyStart = $i
            while ($i -lt $lines.Count -and $lines[$i] -cne $endMarker) { $i++ }
            if ($i -ge $lines.Count) { throw "End marker is missing for $path." }
            $bodyLines = @()
            if ($i -gt $bodyStart) { $bodyLines = $lines[$bodyStart..($i - 1)] }
            $body = $bodyLines -join "`n"
            if ($body.Contains($boundary)) { throw "Text body contains the document Boundary: $path" }
            $dataBytes = ConvertTo-MirrorTextBytes -Body $body -Encoding $encoding -Eol $eol -FinalNewline $final
            if ($restoreByteCount -ge 0 -and $dataBytes.Length -ne $restoreByteCount) { throw "Restore byte count mismatch for $path." }
            $actualSha = Get-MirrorSha256 $dataBytes
            if ($restoreSha -cne "-" -and $actualSha -cne $restoreSha) { throw "Restore SHA-256 mismatch for $path." }
            $entries.Add([pscustomobject]@{
                Path = $path; Kind = "text"; Encoding = $encoding; Eol = $eol
                FinalNewline = $final; OriginalBytes = $originalByteCount; OriginalSha256 = $originalSha
                RestoreBytes = [long]$dataBytes.Length; RestoreSha256 = $actualSha
                Body = $body; DataBytes = $dataBytes; Embedded = $true
            })
            $i++
        } elseif ($kind -eq "binary") {
            $byteText = Read-MirrorValueLine $lines ([ref]$i) "Bytes: "
            $byteCount = ConvertFrom-MirrorByteCount $byteText $false
            $sha = Read-MirrorValueLine $lines ([ref]$i) "SHA-256: "
            [void](Test-MirrorShaValue $sha $false)
            $embeddedText = Read-MirrorValueLine $lines ([ref]$i) "Embedded: "
            if (@("yes", "no") -cnotcontains $embeddedText) { throw "Invalid Embedded value for ${path}: $embeddedText" }
            $embedded = ($embeddedText -ceq "yes")
            $reason = ""
            $dataBytes = $null
            if ($embedded) {
                Read-MirrorRequiredLine $lines ([ref]$i) "===== BASE64 $boundary ====="
                $base64Start = $i
                while ($i -lt $lines.Count -and $lines[$i] -cne $endMarker) { $i++ }
                if ($i -ge $lines.Count) { throw "End marker is missing for $path." }
                $base64Lines = @()
                if ($i -gt $base64Start) { $base64Lines = $lines[$base64Start..($i - 1)] }
                foreach ($line in $base64Lines) {
                    if ($line -cnotmatch '^[A-Za-z0-9+/]*={0,2}$' -or $line.Length -gt 76) { throw "Invalid Base64 line for $path." }
                }
                try { $dataBytes = [Convert]::FromBase64String(($base64Lines -join "")) } catch { throw "Invalid Base64 data for $path." }
                if ($dataBytes.Length -gt $script:MirrorDefaultBinaryFileLimit) { throw "Embedded binary exceeds the file limit: $path" }
                $embeddedTotal += $dataBytes.Length
                if ($embeddedTotal -gt $script:MirrorDefaultBinaryTotalLimit) { throw "Embedded binaries exceed the total limit." }
                if ($dataBytes.Length -ne $byteCount) { throw "Byte count mismatch for $path." }
                if ((Get-MirrorSha256 $dataBytes) -cne $sha) { throw "SHA-256 mismatch for $path." }
                $i++
            } else {
                $reason = Read-MirrorValueLine $lines ([ref]$i) "Reason: "
                if (@("unsupported", "file-limit", "total-limit") -cnotcontains $reason) { throw "Invalid binary skip reason for ${path}: $reason" }
                Read-MirrorRequiredLine $lines ([ref]$i) $endMarker
            }
            $entries.Add([pscustomobject]@{
                Path = $path; Kind = "binary"; Bytes = $byteCount; Sha256 = $sha
                Embedded = $embedded; Reason = $reason; DataBytes = $dataBytes
            })
        } else {
            Read-MirrorRequiredLine $lines ([ref]$i) $endMarker
            $entries.Add([pscustomobject]@{ Path = $path; Kind = "dir"; Embedded = $false })
        }
        Read-MirrorRequiredLine $lines ([ref]$i) ""
    }
    Read-MirrorRequiredLine $lines ([ref]$i) "===== END MIRROR $boundary ====="
    if ($i -lt $lines.Count) {
        if ($i -ne ($lines.Count - 1) -or $lines[$i] -cne "") { throw "Line $($i + 1): unexpected content after the mirror." }
        $i++
    }
    if ($i -ne $lines.Count) { throw "Unexpected trailing content." }

    return [pscustomobject]@{
        Version = 1; Boundary = $boundary; SourceName = $sourceName; SourceType = $sourceType
        CreatedAt = $createdAt; Instructions = $instructions; Entries = $entries.ToArray()
    }
}

function Test-MirrorExistingReparseAncestor {
    param([string]$Path)
    $current = $Path
    while ($current) {
        if (Test-Path -LiteralPath $current) {
            $item = Get-Item -LiteralPath $current -Force -ErrorAction Stop
            if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { return $true }
        }
        $parent = Split-Path $current -Parent
        if (-not $parent -or $parent -eq $current) { break }
        $current = $parent
    }
    return $false
}

function Restore-MirrorDocument {
    param(
        [Parameter(Mandatory = $true)]$Mirror,
        [Parameter(Mandatory = $true)][string]$Destination
    )
    $destinationFull = [System.IO.Path]::GetFullPath($Destination).TrimEnd('\')
    if (Test-Path -LiteralPath $destinationFull) { throw "Restore destination already exists: $destinationFull" }
    if (Test-MirrorExistingReparseAncestor (Split-Path $destinationFull -Parent)) { throw "Restore destination has a reparse-point ancestor." }
    $prefix = $destinationFull + "\"
    $filePaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $dirPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $planned = New-Object System.Collections.Generic.List[object]
    $skipped = New-Object System.Collections.Generic.List[string]

    foreach ($entry in $Mirror.Entries) {
        [void](Test-MirrorRelativePath $entry.Path $entry.Kind)
        $relative = $entry.Path.TrimEnd('/').Replace('/', '\')
        $full = [System.IO.Path]::GetFullPath((Join-Path $destinationFull $relative))
        if (-not $full.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) { throw "Restore path escapes the destination: $($entry.Path)" }
        if ($full.Length -gt $script:MirrorMaxPath) { throw "Restore path is too long for PowerShell 5.1: $($entry.Path)" }
        if ($entry.Kind -eq "dir") {
            if (-not $dirPaths.Add($relative)) { throw "Duplicate restore directory: $($entry.Path)" }
            $planned.Add([pscustomobject]@{ Kind = "dir"; FullPath = $full; RelativePath = $relative; Bytes = $null })
        } elseif ($entry.Kind -eq "binary" -and -not $entry.Embedded) {
            $skipped.Add($entry.Path)
        } else {
            if (-not $filePaths.Add($relative)) { throw "Duplicate restore file: $($entry.Path)" }
            $bytes = $entry.DataBytes
            if ($null -eq $bytes) { throw "Restorable entry has no data: $($entry.Path)" }
            $planned.Add([pscustomobject]@{ Kind = "file"; FullPath = $full; RelativePath = $relative; Bytes = $bytes })
        }
    }
    foreach ($path in $filePaths) {
        $parts = $path.Split('\')
        $parent = ""
        for ($j = 0; $j -lt ($parts.Length - 1); $j++) {
            $parent = if ($parent) { $parent + "\" + $parts[$j] } else { $parts[$j] }
            if ($filePaths.Contains($parent)) { throw "A file blocks a child path: $path" }
        }
        if ($dirPaths.Contains($path)) { throw "File/directory collision: $path" }
    }

    try {
        New-Item -ItemType Directory -Path $destinationFull -Force -ErrorAction Stop | Out-Null
        foreach ($item in @($planned | Where-Object { $_.Kind -eq "dir" } | Sort-Object { $_.FullPath.Length })) {
            New-Item -ItemType Directory -Path $item.FullPath -Force -ErrorAction Stop | Out-Null
        }
        foreach ($item in @($planned | Where-Object { $_.Kind -eq "file" })) {
            $parent = Split-Path $item.FullPath -Parent
            if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force -ErrorAction Stop | Out-Null }
            [System.IO.File]::WriteAllBytes($item.FullPath, $item.Bytes)
        }
    } catch {
        if (Test-Path -LiteralPath $destinationFull) { Remove-Item -LiteralPath $destinationFull -Recurse -Force -ErrorAction SilentlyContinue }
        throw
    }
    return [pscustomobject]@{
        Destination = $destinationFull
        RestoredFiles = @($planned | Where-Object { $_.Kind -eq "file" }).Count
        RestoredDirectories = @($planned | Where-Object { $_.Kind -eq "dir" }).Count
        Skipped = $skipped.ToArray()
    }
}

function Get-MirrorRestoreDestination {
    param([Parameter(Mandatory = $true)][string]$OutputRoot)
    if (-not (Test-Path -LiteralPath $OutputRoot)) { New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null }
    $stem = "md-mirror_restore_" + (Get-Date -Format "yyyyMMdd_HHmmss")
    $candidate = Join-Path $OutputRoot $stem
    $suffix = 1
    while (Test-Path -LiteralPath $candidate) {
        $candidate = Join-Path $OutputRoot ($stem + "_" + $suffix)
        $suffix++
    }
    return $candidate
}
