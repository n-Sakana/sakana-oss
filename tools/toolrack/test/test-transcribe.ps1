# test/test-transcribe.ps1 -- transcribe assets and behavior under PowerShell 5.1.
. (Join-Path $PSScriptRoot "_assert.ps1")

$root = Split-Path $PSScriptRoot -Parent
$tool = Join-Path $root "tool\transcribe"
$bin = Join-Path $tool "bin"
$model = Join-Path $tool "model"
$fixture = Join-Path $root "test\fixtures\transcribe"

foreach ($name in @(
    "sherpa-onnx.dll",
    "sherpa-onnx-c-api.dll",
    "onnxruntime.dll",
    "NAudio.dll"
)) {
    Assert-True (Test-Path -LiteralPath (Join-Path $bin $name) -PathType Leaf) ("bin has " + $name)
}

$modelFiles = @(
    "encoder-epoch-99-avg-1.int8.onnx.part01",
    "encoder-epoch-99-avg-1.int8.onnx.part02",
    "decoder-epoch-99-avg-1.int8.onnx",
    "joiner-epoch-99-avg-1.int8.onnx",
    "tokens.txt",
    "silero_vad.onnx"
)
foreach ($name in $modelFiles) {
    Assert-True (Test-Path -LiteralPath (Join-Path $model $name) -PathType Leaf) ("model has " + $name)
}

$limit = 100MB
$parts = @(Get-ChildItem -LiteralPath $model -Filter "*.part*" -File -ErrorAction SilentlyContinue)
Assert-True ($parts.Count -eq 2) "encoder is stored as two regular-Git parts"
Assert-True (@($parts | Where-Object { $_.Length -ge $limit }).Count -eq 0) "each model part is below 100 MiB"

Assert-True (Test-Path -LiteralPath (Join-Path $fixture "sample2.wav") -PathType Leaf) "Japanese WAV fixture is present"
Assert-True (Test-Path -LiteralPath (Join-Path $fixture "sample2.expected.txt") -PathType Leaf) "expected transcript data is present"
Assert-True (Test-Path -LiteralPath (Join-Path $tool "THIRD-PARTY-NOTICES.md") -PathType Leaf) "third-party notices are present"
Assert-True (Test-Path -LiteralPath (Join-Path $tool "licenses\Apache-2.0.txt") -PathType Leaf) "Apache license is present"
Assert-True (Test-Path -LiteralPath (Join-Path $tool "licenses\MIT.txt") -PathType Leaf) "MIT license is present"

$assetHashes = @{
    "bin\NAudio.dll" = "BC4BACC3B8B28D898F1671B79F216CCA439F95EB60CD32D3E3ECAFBECAC42780"
    "bin\sherpa-onnx.dll" = "B6D9A12D659C742A3D9E4D72204186B50AFBD6C56A0F36893F2DC0DCE627245F"
    "bin\sherpa-onnx-c-api.dll" = "614878147C05121AEB1514EC4FB3E48B89751591532ECA9208235B9AB868306A"
    "bin\onnxruntime.dll" = "DAA77083A45BF525DA0DDE9E87F85D8EB146F58F9C9AA7124CA84545E1C0F148"
    "model\encoder-epoch-99-avg-1.int8.onnx.part01" = "48895C41020DA39B020128252B9053152E0C5BB6CA555C49DFD5756811223517"
    "model\encoder-epoch-99-avg-1.int8.onnx.part02" = "9B06F691E3505A5EE5760E57EB559E40A5C95CCF9FB08881717DD5EB74EC4F87"
    "model\decoder-epoch-99-avg-1.int8.onnx" = "8F0BFF94D38797B03B762634ED03211A8E303D06CC4603CDD0CF4199D6EB1485"
    "model\joiner-epoch-99-avg-1.int8.onnx" = "49CC7EA1D3D35A40A27442DB5E89996DA64BF0E683A903DCE76E99E57A12E4DE"
    "model\tokens.txt" = "2C3AC659818A48A0C04010E0593BBC4D7C8A24A054340B01131499C05FD52DEF"
    "model\silero_vad.onnx" = "9E2449E1087496D8D4CABA907F23E0BD3F78D91FA552479BB9C23AC09CBB1FD6"
}
foreach ($relativePath in $assetHashes.Keys) {
    $actual = (Get-FileHash -Algorithm SHA256 -LiteralPath (Join-Path $tool $relativePath)).Hash
    Assert-True ($actual -eq $assetHashes[$relativePath]) ("asset hash: " + $relativePath)
}

$enginePath = Join-Path $tool "engine.cs"
Assert-True (Test-Path -LiteralPath $enginePath -PathType Leaf) "engine.cs exists"
if (Test-Path -LiteralPath $enginePath -PathType Leaf) {
    $env:PATH = $bin + ";" + $env:PATH
    Add-Type -Path (Join-Path $bin "sherpa-onnx.dll")
    Add-Type -Path (Join-Path $bin "NAudio.dll")
    $engineSource = [IO.File]::ReadAllText($enginePath)
    $engineReferences = @(
        (Join-Path $bin "sherpa-onnx.dll"),
        (Join-Path $bin "NAudio.dll")
    )
    Add-Type -TypeDefinition $engineSource -ReferencedAssemblies $engineReferences

    $encoderPath = Join-Path $model "encoder-epoch-99-avg-1.int8.onnx"
    if (Test-Path -LiteralPath $encoderPath) {
        Remove-Item -LiteralPath $encoderPath -Force
    }
    $prepared = [TranscriberEngine]::EnsureModel($model)
    Assert-True ($prepared -eq $encoderPath) "split encoder reconstructs to the expected path"
    Assert-True (Test-Path -LiteralPath $encoderPath -PathType Leaf) "reconstructed encoder exists"
    Assert-True ((Get-Item -LiteralPath $encoderPath).Length -eq 154670139) "reconstructed encoder length matches"
    $encoderHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $encoderPath).Hash
    Assert-True ($encoderHash -eq "2C7BD08A8A99F9DDD0D9E458456577B1F6279214E51426F114F9ECED44C54E1D") "reconstructed encoder hash matches"

    $oldTime = [datetime]::SpecifyKind([datetime]"2001-01-01T00:00:00", [DateTimeKind]::Utc)
    [IO.File]::SetLastWriteTimeUtc($encoderPath, $oldTime)
    [void][TranscriberEngine]::EnsureModel($model)
    Assert-True ([IO.File]::GetLastWriteTimeUtc($encoderPath) -eq $oldTime) "valid reconstructed encoder is reused"

    $wav = Join-Path $fixture "sample2.wav"
    $expected = [IO.File]::ReadAllText((Join-Path $fixture "sample2.expected.txt")).Trim()
    $segments = [TranscriberEngine]::DecodeFileSegments($model, $wav)
    Assert-True ($segments.Count -ge 1) "VAD produces at least one speech segment"
    $actualText = (($segments -join "").Trim())
    Assert-True ($actualText -eq $expected) ("VAD segments recognize the expected Japanese text (got: " + $actualText + ")")

    $convertMethod = [TranscriberEngine].GetMethod("ConvertPcm16")
    $testStartMethod = [TranscriberEngine].GetMethod("StartTestSession")
    $testFeedMethod = [TranscriberEngine].GetMethod("FeedPcm16ForTest")
    $waitMethod = [TranscriberEngine].GetMethod("WaitForStop")
    $queueApiReady = $null -ne $convertMethod -and $null -ne $testStartMethod -and
        $null -ne $testFeedMethod -and $null -ne $waitMethod
    Assert-True $queueApiReady "engine exposes gain and queued-worker test seams"

    if ($queueApiReady) {
        $pcm = New-Object byte[] 8
        [Array]::Copy([BitConverter]::GetBytes([int16]1000), 0, $pcm, 0, 2)
        [Array]::Copy([BitConverter]::GetBytes([int16]-1000), 0, $pcm, 2, 2)
        [Array]::Copy([BitConverter]::GetBytes([int16]20000), 0, $pcm, 4, 2)
        [Array]::Copy([BitConverter]::GetBytes([int16]-20000), 0, $pcm, 6, 2)
        [single]$level1 = 0
        [single]$level2 = 0
        $samples1 = [TranscriberEngine]::ConvertPcm16($pcm, $pcm.Length, 1.0, [ref]$level1)
        $samples2 = [TranscriberEngine]::ConvertPcm16($pcm, $pcm.Length, 2.0, [ref]$level2)
        Assert-True ([Math]::Abs($samples2[0] - (2000.0 / 32768.0)) -lt 0.00001) "software gain doubles PCM samples"
        Assert-True ($samples2[2] -eq 1.0 -and $samples2[3] -eq -1.0) "software gain clips to the valid range"
        Assert-True ($level2 -gt $level1 -and $level2 -le 1.0) "dBFS meter rises with gain and stays normalized"

        $wavBytes = [IO.File]::ReadAllBytes($wav)
        $chunk = 12
        $dataOffset = -1
        $dataLength = 0
        while ($chunk + 8 -le $wavBytes.Length) {
            $chunkId = [Text.Encoding]::ASCII.GetString($wavBytes, $chunk, 4)
            $chunkLength = [BitConverter]::ToInt32($wavBytes, $chunk + 4)
            if ($chunkId -eq "data") {
                $dataOffset = $chunk + 8
                $dataLength = $chunkLength
                break
            }
            $chunk += 8 + $chunkLength + ($chunkLength -band 1)
        }
        Assert-True ($dataOffset -ge 0) "WAV fixture has a PCM data chunk"

        $queued = New-Object TranscriberEngine($model)
        try {
            $queued.Gain = 1.0
            $queued.StartTestSession()
            for ($offset = 0; $offset -lt $dataLength; $offset += 3200) {
                $count = [Math]::Min(3200, $dataLength - $offset)
                $buffer = New-Object byte[] $count
                [Array]::Copy($wavBytes, $dataOffset + $offset, $buffer, 0, $count)
                $queued.FeedPcm16ForTest($buffer, $count)
            }
            $queued.Stop()
            Assert-True ($queued.WaitForStop(30000)) "queued worker stops and flushes within the timeout"
            $queuedText = New-Object System.Collections.Generic.List[string]
            [string]$textItem = ""
            while ($queued.TryGetText([ref]$textItem)) {
                $queuedText.Add($textItem)
                $textItem = ""
            }
            [string]$errorItem = ""
            $hasError = $queued.TryGetError([ref]$errorItem)
            Assert-True (-not $hasError) ("queued worker has no background error: " + $errorItem)
            Assert-True ((($queuedText -join "").Trim()) -eq $expected) "queued PCM path recognizes and flushes the full fixture"
            Assert-True (-not $queued.IsRunning) "queued worker reports stopped state"

            $queued.StartTestSession()
            $queued.Stop()
            Assert-True ($queued.WaitForStop(5000)) "queued worker can start and stop a second session"
        } finally {
            $queued.Dispose()
        }
    }

    Assert-True ([NAudio.Wave.WaveInEvent]::DeviceCount -ge 0) "NAudio microphone API is available"
}

$mainPath = Join-Path $tool "main.ps1"
Assert-True (Test-Path -LiteralPath $mainPath -PathType Leaf) "main.ps1 exists"
if (Test-Path -LiteralPath $mainPath -PathType Leaf) {
    $mainBytes = [IO.File]::ReadAllBytes($mainPath)
    $hasBom = $mainBytes.Length -ge 3 -and $mainBytes[0] -eq 0xEF -and
        $mainBytes[1] -eq 0xBB -and $mainBytes[2] -eq 0xBF
    Assert-True $hasBom "main.ps1 has a UTF-8 BOM for PowerShell 5.1"

    $oldNoPause = $env:TOOLRACK_NOPAUSE
    $env:TOOLRACK_NOPAUSE = "1"
    try {
        $smokeOutput = @(& powershell.exe -NoProfile -Sta -ExecutionPolicy Bypass -File $mainPath -Smoke 2>&1)
        $smokeCode = $LASTEXITCODE
    } finally {
        $env:TOOLRACK_NOPAUSE = $oldNoPause
    }
    Assert-True ($smokeCode -eq 0) ("main.ps1 GUI smoke exits 0: " + ($smokeOutput -join " | "))
    Assert-Contains $smokeOutput "*SMOKE_OK*" "main.ps1 GUI smoke reaches the ready state"
    Assert-Contains $smokeOutput "*CLEAR_OK*" "Clear empties the transcript without stopping the GUI"
}

$manifestPath = Join-Path $tool "tool.json"
Assert-True (Test-Path -LiteralPath $manifestPath -PathType Leaf) "tool.json exists"
if (Test-Path -LiteralPath $manifestPath -PathType Leaf) {
    . (Join-Path $root "common\install.ps1")
    $manifestResult = Read-Manifest $tool
    Assert-True $manifestResult.Ok "transcribe tool.json parses"
    $manifestErrors = @(Test-Manifest $manifestResult.Data "transcribe" $tool)
    Assert-True ($manifestErrors.Count -eq 0) ("transcribe manifest is valid: " + ($manifestErrors -join "; "))
    Assert-True ($manifestResult.Data.name -eq "Transcribe") "context-menu name is English"

    $manifest = $manifestResult.Data
    $run = $manifest.run
    $variants = @($manifest.variants)
    Assert-True ($manifest.name -eq "Transcribe") "manifest has the English menu name"
    Assert-True (@($manifest.on).Count -eq 1 -and @($manifest.on)[0] -eq "background") "manifest targets folder background only"
    Assert-True ($run.type -eq "powershell") "manifest uses the PowerShell launcher"
    Assert-True ($run.entry -eq "main.ps1") "manifest points to main.ps1"
    Assert-True ($run.window -eq "gui") "manifest uses GUI window mode"
    Assert-True ($variants.Count -eq 1) "manifest has one Start variant"
    Assert-True ($variants[0].label -eq "Start (Japanese)") "manifest variant label is stable"
    Assert-True (@($variants[0].args).Count -eq 0) "manifest variant has no extra arguments"
}

Exit-Test
