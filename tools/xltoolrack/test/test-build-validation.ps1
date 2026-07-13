. (Join-Path $PSScriptRoot '_harness.ps1')

Add-Type -AssemblyName System.IO.Compression.FileSystem
$root = Split-Path $PSScriptRoot -Parent
$dir = New-TestDirectory

function New-Fixture {
    param([string]$Name)
    $fixture = Join-Path $dir $Name
    New-Item -ItemType Directory -Force -Path $fixture | Out-Null
    Copy-Item -LiteralPath (Join-Path $root 'scripts') -Destination $fixture -Recurse
    Copy-Item -LiteralPath (Join-Path $root 'src') -Destination $fixture -Recurse
    return $fixture
}

function Write-AsciiTool {
    param([string]$Project, [string]$Name, [string]$Body)
    $path = Join-Path $Project ("src\tools\$Name.bas")
    [IO.File]::WriteAllText($path, $Body, [Text.Encoding]::ASCII)
}

function Invoke-FixtureBuild {
    param([string]$Project)
    $out = Join-Path $Project 'out'
    & (Join-Path $Project 'scripts\Build-Addin.ps1') -OutputFormat xlam -OutputDirectory $out | Out-Null
    return (Join-Path $out 'xltoolrack.xlam')
}

function Expect-BuildFailure {
    param([string]$Project, [string]$Pattern, [string]$Message)
    $caught = ''
    try { $null = Invoke-FixtureBuild $Project } catch { $caught = $_.Exception.Message }
    Assert-True ($caught -match $Pattern) "$Message (error=$caught)"
}

try {
    $valid = New-Fixture 'valid'
    Write-AsciiTool $valid 'sample_valid' @'
Attribute VB_Name = "sample_valid"
Option Explicit
Option Private Module
 '@name Sample Valid
 '@ribbon Sample Valid
 '@group Validation
 '@maxjobs 1
Public Sub Run(ByVal ctx As InfraContext)
    ctx.LogMessage "valid"
End Sub
Public Sub OnResult(ByVal ctx As InfraContext, ByVal jobId As String, ByVal version As Long)
End Sub
Public Sub OnFlush(ByVal ctx As InfraContext)
End Sub
'@
    $xlam = Invoke-FixtureBuild $valid
    Assert-True (Test-Path -LiteralPath $xlam) 'compliant tool built successfully'
    $zip = [IO.Compression.ZipFile]::OpenRead($xlam)
    try {
        $entry = $zip.GetEntry('customUI/customUI14.xml')
        Assert-True ($null -ne $entry) 'generated xlam contains customUI14.xml'
        if ($null -ne $entry) {
            $reader = New-Object IO.StreamReader($entry.Open())
            try { $xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
            Assert-True ($xml -match 'id="sample_valid"') 'dynamic ribbon contains compliant tool id'
            Assert-True ($xml -match 'onAction="Infra_RunTool"') 'dynamic ribbon uses mandatory dispatcher'
        }
    } finally { $zip.Dispose() }

    $missingRun = New-Fixture 'missing_run'
    Write-AsciiTool $missingRun 'bad_missing_run' @'
Attribute VB_Name = "bad_missing_run"
Option Explicit
Option Private Module
 '@name Missing Run
 '@ribbon Missing Run
 '@group Validation
 '@maxjobs 1
Private Sub Helper()
End Sub
'@
    Expect-BuildFailure $missingRun 'Run.*signature|missing.*Run' 'tool without Run was rejected'

    $badSignature = New-Fixture 'bad_signature'
    Write-AsciiTool $badSignature 'bad_signature' @'
Attribute VB_Name = "bad_signature"
Option Explicit
Option Private Module
 '@name Bad Signature
 '@ribbon Bad Signature
 '@group Validation
 '@maxjobs 1
Public Sub Run(ByRef ctx As InfraContext)
End Sub
'@
    Expect-BuildFailure $badSignature 'Run.*signature' 'wrong Run signature was rejected'

    $badFlushSignature = New-Fixture 'bad_flush_signature'
    Write-AsciiTool $badFlushSignature 'bad_flush_signature' @'
Attribute VB_Name = "bad_flush_signature"
Option Explicit
Option Private Module
 '@name Bad Flush Signature
 '@ribbon Bad Flush Signature
 '@group Validation
 '@maxjobs 1
Public Sub Run(ByVal ctx As InfraContext)
End Sub
Public Sub OnResult(ByVal ctx As InfraContext, ByVal jobId As String, ByVal version As Long)
End Sub
Public Sub OnFlush(ByRef ctx As InfraContext)
End Sub
'@
    Expect-BuildFailure $badFlushSignature 'OnFlush.*signature' 'wrong OnFlush signature was rejected'

    $orphanFlush = New-Fixture 'orphan_flush'
    Write-AsciiTool $orphanFlush 'orphan_flush' @'
Attribute VB_Name = "orphan_flush"
Option Explicit
Option Private Module
 '@name Orphan Flush
 '@ribbon Orphan Flush
 '@group Validation
 '@maxjobs 1
Public Sub Run(ByVal ctx As InfraContext)
End Sub
Public Sub OnFlush(ByVal ctx As InfraContext)
End Sub
'@
    Expect-BuildFailure $orphanFlush 'OnFlush requires OnResult' 'OnFlush without OnResult was rejected'

    $autoOpen = New-Fixture 'auto_open'
    Write-AsciiTool $autoOpen 'bad_auto' @'
Attribute VB_Name = "bad_auto"
Option Explicit
Option Private Module
 '@name Bad Auto
 '@ribbon Bad Auto
 '@group Validation
 '@maxjobs 1
Public Sub Run(ByVal ctx As InfraContext)
End Sub
Public Sub Auto_Open()
End Sub
'@
    Expect-BuildFailure $autoOpen 'Auto_Open|auto hook' 'Auto_Open bypass was rejected'

    $ribbon = New-Fixture 'ribbon_callback'
    Write-AsciiTool $ribbon 'bad_ribbon' @'
Attribute VB_Name = "bad_ribbon"
Option Explicit
Option Private Module
 '@name Bad Ribbon
 '@ribbon Bad Ribbon
 '@group Validation
 '@maxjobs 1
Public Sub Run(ByVal ctx As InfraContext)
End Sub
Public Sub DirectRibbon(ByVal control As IRibbonControl)
End Sub
'@
    Expect-BuildFailure $ribbon 'IRibbonControl|ribbon callback' 'direct ribbon callback was rejected'

    $onTime = New-Fixture 'ontime'
    Write-AsciiTool $onTime 'bad_ontime' @'
Attribute VB_Name = "bad_ontime"
Option Explicit
Option Private Module
 '@name Bad OnTime
 '@ribbon Bad OnTime
 '@group Validation
 '@maxjobs 1
Public Sub Run(ByVal ctx As InfraContext)
    Application.OnTime Now, "bad_ontime.Run"
End Sub
'@
    Expect-BuildFailure $onTime 'OnTime' 'tool-owned OnTime scheduling was rejected'

    $cursor = New-Fixture 'cursor'
    Write-AsciiTool $cursor 'bad_cursor' @'
Attribute VB_Name = "bad_cursor"
Option Explicit
Option Private Module
 '@name Bad Cursor
 '@ribbon Bad Cursor
 '@group Validation
 '@maxjobs 1
Public Sub Run(ByVal ctx As InfraContext)
    Application.Cursor = xlWait
End Sub
'@
    Expect-BuildFailure $cursor 'Application.Cursor' 'cursor overrides were rejected'

    $nonAscii = New-Fixture 'non_ascii'
    $nonAsciiPath = Join-Path $nonAscii 'src\tools\bad_non_ascii.bas'
    [IO.File]::WriteAllText($nonAsciiPath, "Attribute VB_Name = `"bad_non_ascii`"`r`nOption Explicit`r`n' Japanese: test-char-あ`r`n", [Text.Encoding]::UTF8)
    Expect-BuildFailure $nonAscii 'Non-ASCII' 'non-ASCII VBA source was rejected'
} finally {
    Remove-TestDirectory $dir
}

Exit-Test
