# Capture -- range/window capture with image, path, and local OCR results.
param(
    [string]$Target = "",
    [switch]$Smoke,
    [switch]$SelfTest,
    [string]$AuditPath = "",
    [string]$TestOutputDir = "",
    [string]$PreviewLightPath = "",
    [string]$PreviewDarkPath = ""
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$script:AsTaskMethod = $null

function Get-UiText {
    param([string]$Entities)
    return [Net.WebUtility]::HtmlDecode($Entities)
}

function Write-JsonFile {
    param([string]$Path, $Value)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    $parent = Split-Path $Path -Parent
    if ($parent -ne "" -and -not (Test-Path -LiteralPath $parent -PathType Container)) {
        New-Item -ItemType Directory -Force -Path $parent | Out-Null
    }
    [IO.File]::WriteAllText($Path, (ConvertTo-Json $Value -Depth 10), (New-Object Text.UTF8Encoding($false)))
}

function Get-Sha256Text {
    param([string]$Path)
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $hash = $sha.ComputeHash([IO.File]::ReadAllBytes($Path))
        return ([BitConverter]::ToString($hash)).Replace("-", "").ToLowerInvariant()
    } finally { $sha.Dispose() }
}

function Initialize-CaptureRuntime {
    Add-Type -AssemblyName System.Drawing, System.Windows.Forms
    $sourcePath = Join-Path $PSScriptRoot "capture.cs"
    $dllPath = Join-Path $PSScriptRoot "bin\ToolrackCapture.dll"
    $buildPath = Join-Path $PSScriptRoot "bin\build.json"
    foreach ($path in @($sourcePath, $dllPath, $buildPath)) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Capture runtime asset not found: $path" }
    }
    try {
        $build = ConvertFrom-Json ([IO.File]::ReadAllText($buildPath, (New-Object Text.UTF8Encoding($false, $true))))
    } catch { throw ("Capture build metadata is invalid: " + $_.Exception.Message) }
    if ($build.schema -isnot [int] -or [int]$build.schema -ne 1 -or
        [string]$build.sourceSha256 -cnotmatch '^[a-f0-9]{64}$' -or
        [string]::IsNullOrWhiteSpace([string]$build.buildId)) {
        throw "Capture build metadata is invalid"
    }
    $actualHash = Get-Sha256Text $sourcePath
    if ($actualHash -cne [string]$build.sourceSha256) {
        throw "Capture source hash does not match the prebuilt DLL metadata"
    }
    $assembly = [Reflection.Assembly]::LoadFrom($dllPath)
    $buildType = $assembly.GetType("ToolrackCapture.BuildInfo", $true)
    $dllBuildId = [string]$buildType.GetField("Id").GetValue($null)
    if ($dllBuildId -cne [string]$build.buildId) { throw "Capture DLL build ID does not match build metadata" }
    [ToolrackCapture.Native]::EnableDpiAwareness()
}

function Initialize-OcrRuntime {
    Add-Type -AssemblyName System.Runtime.WindowsRuntime
    [void][Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]
    [void][Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType=WindowsRuntime]
    [void][Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType=WindowsRuntime]
    [void][Windows.Graphics.Imaging.SoftwareBitmap, Windows.Foundation, ContentType=WindowsRuntime]
    [void][Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime]
    [void][Windows.Media.Ocr.OcrResult, Windows.Foundation, ContentType=WindowsRuntime]
    if ($null -eq $script:AsTaskMethod) {
        $script:AsTaskMethod = @(
            [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
                $_.Name -eq "AsTask" -and $_.IsGenericMethod -and $_.GetParameters().Count -eq 1
            }
        )[0]
    }
}

function Get-WinRtResult {
    param($Operation, [type]$ResultType)
    $method = $script:AsTaskMethod.MakeGenericMethod(@($ResultType))
    $task = $method.Invoke($null, @($Operation))
    $task.Wait()
    return $task.Result
}

function Get-OcrEngine {
    Initialize-OcrRuntime
    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
    if ($null -eq $engine) { throw "Windows OCR is unavailable. Install a Windows language OCR feature and try again." }
    return $engine
}

function Test-OcrAvailable {
    try { return ($null -ne (Get-OcrEngine)) } catch { return $false }
}

function Get-OcrText {
    param([Drawing.Bitmap]$Bitmap)
    $engine = Get-OcrEngine
    $temp = Join-Path $env:TEMP ("toolrack_capture_ocr_" + [guid]::NewGuid().ToString("N") + ".png")
    $stream = $null
    $softwareBitmap = $null
    try {
        $Bitmap.Save($temp, [Drawing.Imaging.ImageFormat]::Png)
        $storageFile = Get-WinRtResult ([Windows.Storage.StorageFile]::GetFileFromPathAsync($temp)) ([Windows.Storage.StorageFile])
        $stream = Get-WinRtResult ($storageFile.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
        $decoder = Get-WinRtResult ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
        $softwareBitmap = Get-WinRtResult ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
        $result = Get-WinRtResult ($engine.RecognizeAsync($softwareBitmap)) ([Windows.Media.Ocr.OcrResult])
        return ([string]$result.Text).Trim()
    } finally {
        if ($null -ne $softwareBitmap) { $softwareBitmap.Dispose() }
        if ($null -ne $stream) { $stream.Dispose() }
        Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue
    }
}

function Set-ClipboardImage {
    param([Drawing.Bitmap]$Bitmap)
    $lastError = $null
    for ($index = 0; $index -lt 6; $index++) {
        try { [Windows.Forms.Clipboard]::SetImage($Bitmap); return }
        catch { $lastError = $_.Exception; Start-Sleep -Milliseconds 40 }
    }
    throw $lastError
}

function Set-ClipboardText {
    param([string]$Text)
    $lastError = $null
    for ($index = 0; $index -lt 6; $index++) {
        try { [Windows.Forms.Clipboard]::SetText($Text); return }
        catch { $lastError = $_.Exception; Start-Sleep -Milliseconds 40 }
    }
    throw $lastError
}

function Get-CaptureOutputPath {
    param([string]$OutputDirectory = "")
    if ($OutputDirectory -eq "") {
        $repositoryRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
        $OutputDirectory = Join-Path $repositoryRoot "output"
    }
    if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
        New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
    }
    $stem = "capture_" + (Get-Date -Format "yyyyMMdd_HHmmss_fff")
    $path = Join-Path $OutputDirectory ($stem + ".png")
    $suffix = 1
    while (Test-Path -LiteralPath $path) {
        $path = Join-Path $OutputDirectory ($stem + "_" + $suffix + ".png")
        $suffix++
    }
    return [IO.Path]::GetFullPath($path)
}

function Invoke-CaptureResult {
    param(
        [ValidateSet("image", "path", "text")][string]$Action,
        [Drawing.Bitmap]$Bitmap,
        [string]$OutputDirectory = ""
    )
    if ($Action -eq "image") {
        Set-ClipboardImage $Bitmap
        return @{ Value = ""; Message = Get-UiText "&#x753B;&#x50CF;&#x3092;&#x30B3;&#x30D4;&#x30FC;&#x3057;&#x307E;&#x3057;&#x305F;" }
    }
    if ($Action -eq "path") {
        $path = Get-CaptureOutputPath $OutputDirectory
        $Bitmap.Save($path, [Drawing.Imaging.ImageFormat]::Png)
        Set-ClipboardText $path
        return @{ Value = $path; Message = Get-UiText "&#x4FDD;&#x5B58;&#x3057;&#x3066;&#x30D1;&#x30B9;&#x3092;&#x30B3;&#x30D4;&#x30FC;&#x3057;&#x307E;&#x3057;&#x305F;" }
    }
    $text = Get-OcrText $Bitmap
    if ($text -eq "") { throw (Get-UiText "&#x30C6;&#x30AD;&#x30B9;&#x30C8;&#x3092;&#x691C;&#x51FA;&#x3067;&#x304D;&#x307E;&#x305B;&#x3093;&#x3067;&#x3057;&#x305F;") }
    Set-ClipboardText $text
    return @{ Value = $text; Message = Get-UiText "&#x30C6;&#x30AD;&#x30B9;&#x30C8;&#x3092;&#x30B3;&#x30D4;&#x30FC;&#x3057;&#x307E;&#x3057;&#x305F;" }
}

function Show-CaptureError {
    param([string]$Message)
    if ($Smoke.IsPresent -or $SelfTest.IsPresent -or $env:TOOLRACK_NOPAUSE -eq "1") {
        [Console]::Error.WriteLine($Message)
        return
    }
    [void][Windows.Forms.MessageBox]::Show($Message, "Capture", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Error)
}

function Invoke-CaptureSelfTest {
    $bitmap = New-Object Drawing.Bitmap -ArgumentList 64, 40
    try {
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        try {
            $graphics.Clear([Drawing.Color]::FromArgb(38, 92, 180))
            $graphics.FillRectangle([Drawing.Brushes]::White, 8, 8, 20, 12)
        } finally { $graphics.Dispose() }

        [void](Invoke-CaptureResult -Action "image" -Bitmap $bitmap -OutputDirectory $TestOutputDir)
        $clipboardImage = [Windows.Forms.Clipboard]::GetImage()
        $clipboardWidth = 0
        $clipboardHeight = 0
        if ($null -ne $clipboardImage) {
            $clipboardWidth = $clipboardImage.Width
            $clipboardHeight = $clipboardImage.Height
            $clipboardImage.Dispose()
        }
        $pathResult = Invoke-CaptureResult -Action "path" -Bitmap $bitmap -OutputDirectory $TestOutputDir
        $saved = [Drawing.Image]::FromFile([string]$pathResult.Value)
        try { $width = $saved.Width; $height = $saved.Height } finally { $saved.Dispose() }

        $ocrBitmap = New-Object Drawing.Bitmap -ArgumentList 520, 150
        try {
            $ocrGraphics = [Drawing.Graphics]::FromImage($ocrBitmap)
            $ocrFont = New-Object Drawing.Font -ArgumentList "Arial", 48, ([Drawing.FontStyle]::Bold)
            try {
                $ocrGraphics.Clear([Drawing.Color]::White)
                $ocrGraphics.DrawString("TEST 123", $ocrFont, [Drawing.Brushes]::Black, 18, 34)
            } finally { $ocrFont.Dispose(); $ocrGraphics.Dispose() }
            $ocrText = Get-OcrText $ocrBitmap
        } finally { $ocrBitmap.Dispose() }

        Write-JsonFile $AuditPath ([ordered]@{
            SavedPath = [string]$pathResult.Value; Width = $width; Height = $height
            ClipboardText = [Windows.Forms.Clipboard]::GetText()
            ImageClipboardWidth = $clipboardWidth; ImageClipboardHeight = $clipboardHeight
            OcrAvailable = Test-OcrAvailable; OcrText = $ocrText
        })
    } finally { $bitmap.Dispose() }
}

try {
    Initialize-CaptureRuntime
    if ($SelfTest.IsPresent) { Invoke-CaptureSelfTest; exit 0 }
    if ($Smoke.IsPresent -or $env:TOOLRACK_SMOKE -eq "1") {
        $cursor = [Windows.Forms.Cursor]::Position
        $work = [Windows.Forms.Screen]::FromPoint($cursor).WorkingArea
        $audit = [ToolrackCapture.CapturePalette]::CreateAudit(1.0, $false, $work, $cursor, $false)
        Write-JsonFile $AuditPath $audit
        if ($PreviewLightPath -ne "") { [ToolrackCapture.CapturePalette]::RenderPreview($PreviewLightPath, $false, 1.0) }
        if ($PreviewDarkPath -ne "") { [ToolrackCapture.CapturePalette]::RenderPreview($PreviewDarkPath, $true, 1.0) }
        exit 0
    }

    $choice = [ToolrackCapture.CapturePalette]::ShowPalette()
    if ([string]::IsNullOrWhiteSpace($choice)) { exit 0 }
    $parts = @($choice -split ":", 2)
    $mode = $parts[0]
    $action = $parts[1]
    if ($mode -eq "range") { $rectangle = [ToolrackCapture.RangeSelector]::SelectRectangle() }
    else { $rectangle = [ToolrackCapture.WindowSelector]::SelectRectangle() }
    if ($rectangle.Width -lt 1 -or $rectangle.Height -lt 1) { exit 0 }

    Start-Sleep -Milliseconds 90
    $bitmap = [ToolrackCapture.Native]::CaptureRectangle($rectangle)
    try { $result = Invoke-CaptureResult -Action $action -Bitmap $bitmap }
    finally { $bitmap.Dispose() }
    [ToolrackCapture.CaptureToast]::Show([string]$result.Message)
    exit 0
} catch {
    Show-CaptureError $_.Exception.Message
    exit 1
}
