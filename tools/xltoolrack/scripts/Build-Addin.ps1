param(
    [ValidateSet('all', 'xlam', 'xlsm')]
    [string]$OutputFormat = 'all',
    [string]$OutputDirectory = ''
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir
$srcDir = Join-Path $projectDir 'src'
$toolsDir = Join-Path $srcDir 'tools'
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $distDir = Join-Path $projectDir 'dist'
} elseif ([IO.Path]::IsPathRooted($OutputDirectory)) {
    $distDir = [IO.Path]::GetFullPath($OutputDirectory)
} else {
    $distDir = [IO.Path]::GetFullPath((Join-Path $projectDir $OutputDirectory))
}

function Release-ComObject {
    param([object]$ComObject)
    if ($null -eq $ComObject) { return }
    try {
        if ([Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
        }
    } catch {}
}

function Get-VbaFiles {
    param([string]$Root)
    if (-not (Test-Path -LiteralPath $Root)) { return @() }
    return @(Get-ChildItem -LiteralPath $Root -Recurse -File |
        Where-Object { $_.Extension.ToLowerInvariant() -in @('.bas', '.cls', '.frm') } |
        Sort-Object FullName)
}

function Assert-AsciiAndRuntimePolicy {
    $files = Get-VbaFiles $srcDir
    foreach ($file in $files) {
        $bytes = [IO.File]::ReadAllBytes($file.FullName)
        if ($bytes | Where-Object { $_ -gt 127 } | Select-Object -First 1) {
            throw "Non-ASCII VBA source: $($file.FullName)"
        }
        $text = [IO.File]::ReadAllText($file.FullName, [Text.Encoding]::ASCII)
        if ($text -match '(?im)^\s*(?:Public\s+|Private\s+)?Declare\b') {
            throw "Forbidden Declare in $($file.FullName)"
        }
        if ($text -match '(?i)\bShell\b|winmgmts|WbemScripting|SWbem|taskkill') {
            throw "Forbidden Shell/WMI construct in $($file.FullName)"
        }
        if ($text -match '(?i)Application\s*\.\s*Cursor') {
            throw "Forbidden Application.Cursor in $($file.FullName); leave Excel contextual cursors unmanaged"
        }
        if ($file.Name -notin @('WorkerBridge.bas', 'JobPump.bas') -and $text -match '(?i)Application\s*\.\s*OnTime') {
            throw "Application.OnTime is reserved for the infrastructure (worker bootstrap and FE pump): $($file.FullName)"
        }
    }
}

function Get-MetaValue {
    param([string]$Text, [string]$Name, [bool]$Required = $true)
    $match = [regex]::Match($Text, "(?im)^\s*'@$([regex]::Escape($Name))\s+(.+?)\s*$")
    if (-not $match.Success) {
        if ($Required) { throw "missing @$Name metadata" }
        return ''
    }
    return $match.Groups[1].Value.Trim()
}

function Normalize-VbaText {
    param([string]$Text)
    return [regex]::Replace($Text, '_\s*\r?\n\s*', ' ')
}

function Get-ToolDefinitions {
    $definitions = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    $toolFiles = @(Get-ChildItem -LiteralPath $toolsDir -Recurse -File -Filter '*.bas' | Sort-Object FullName)
    foreach ($file in $toolFiles) {
        $text = [IO.File]::ReadAllText($file.FullName, [Text.Encoding]::ASCII)
        $normalized = Normalize-VbaText $text
        $nameMatch = [regex]::Match($text, '(?im)^Attribute\s+VB_Name\s*=\s*"([^"]+)"\s*$')
        if (-not $nameMatch.Success) { throw "$($file.FullName): missing Attribute VB_Name" }
        $module = $nameMatch.Groups[1].Value
        if ($module -notmatch '^[A-Za-z][A-Za-z0-9_]{0,30}$') {
            throw "$($file.FullName): invalid VBA module/tool id '$module'"
        }
        if ($seen.ContainsKey($module.ToLowerInvariant())) { throw "duplicate tool id: $module" }
        $seen[$module.ToLowerInvariant()] = $true

        try {
            $displayName = Get-MetaValue $text 'name'
            $ribbon = Get-MetaValue $text 'ribbon'
            $group = Get-MetaValue $text 'group'
            $maxText = Get-MetaValue $text 'maxjobs'
            $internalText = Get-MetaValue $text 'internal' $false
        } catch {
            throw "$($file.FullName): $($_.Exception.Message)"
        }
        $maxJobs = 0
        if (-not [int]::TryParse($maxText, [ref]$maxJobs) -or $maxJobs -lt 1 -or $maxJobs -gt 10) {
            throw "$($file.FullName): @maxjobs must be an integer from 1 to 10"
        }
        if ($normalized -notmatch '(?im)^\s*Option\s+Private\s+Module\s*$') {
            throw "$($file.FullName): Option Private Module is required"
        }
        $runPattern = '(?im)^\s*Public\s+Sub\s+Run\s*\(\s*ByVal\s+ctx\s+As\s+InfraContext\s*\)\s*$'
        if ($normalized -notmatch $runPattern) {
            throw "$($file.FullName): missing exact Run signature: Public Sub Run(ByVal ctx As InfraContext)"
        }
        $hasWorkerDeclaration = $normalized -match '(?im)^\s*Public\s+Sub\s+Worker\s*\('
        $workerPattern = '(?im)^\s*Public\s+Sub\s+Worker\s*\(\s*ByVal\s+job\s+As\s+InfraJob\s*\)\s*$'
        if ($hasWorkerDeclaration -and $normalized -notmatch $workerPattern) {
            throw "$($file.FullName): invalid Worker signature"
        }
        $hasResultDeclaration = $normalized -match '(?im)^\s*Public\s+Sub\s+OnResult\s*\('
        $resultPattern = '(?im)^\s*Public\s+Sub\s+OnResult\s*\(\s*ByVal\s+ctx\s+As\s+InfraContext\s*,\s*ByVal\s+jobId\s+As\s+String\s*,\s*ByVal\s+version\s+As\s+Long\s*\)\s*$'
        if ($hasResultDeclaration -and $normalized -notmatch $resultPattern) {
            throw "$($file.FullName): invalid OnResult signature"
        }
        $hasFlushDeclaration = $normalized -match '(?im)^\s*Public\s+Sub\s+OnFlush\s*\('
        $flushPattern = '(?im)^\s*Public\s+Sub\s+OnFlush\s*\(\s*ByVal\s+ctx\s+As\s+InfraContext\s*\)\s*$'
        if ($hasFlushDeclaration -and $normalized -notmatch $flushPattern) {
            throw "$($file.FullName): invalid OnFlush signature"
        }
        if ($hasFlushDeclaration -and -not $hasResultDeclaration) {
            throw "$($file.FullName): OnFlush requires OnResult"
        }
        if ($normalized -match '(?im)^\s*(?:Public\s+|Private\s+)?Sub\s+(Auto_Open|Auto_Close|Auto_Activate|Auto_Deactivate|Workbook_Open|Workbook_BeforeClose)\b') {
            throw "$($file.FullName): forbidden auto hook $($matches[1])"
        }
        if ($normalized -match '(?is)Public\s+Sub\s+[A-Za-z][A-Za-z0-9_]*\s*\([^\)]*IRibbonControl') {
            throw "$($file.FullName): direct IRibbonControl ribbon callback is forbidden"
        }
        $publicProcedures = [regex]::Matches($normalized, '(?im)^\s*Public\s+(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z][A-Za-z0-9_]*)')
        foreach ($procedure in $publicProcedures) {
            if ($procedure.Groups[1].Value -notin @('Run', 'Worker', 'OnResult', 'OnFlush')) {
                throw "$($file.FullName): unsupported public procedure $($procedure.Groups[1].Value)"
            }
        }
        if ($text -match '(?i)Application\s*\.\s*(ScreenUpdating|DisplayAlerts|Calculation)\s*=|\bMsgBox\b|\bOpen\s+.+\s+For\s+(?:Input|Output|Append|Binary)') {
            Write-Warning "$($file.Name): direct application state, MsgBox, or raw file IO detected"
        }

        $definitions.Add([pscustomobject]@{
            Module = $module
            Name = $displayName
            Ribbon = $ribbon
            Group = $group
            MaxJobs = $maxJobs
            Internal = ($internalText -match '^(?i:true|1|yes)$')
            HasWorker = [bool]$hasWorkerDeclaration
            HasResult = [bool]$hasResultDeclaration
            HasFlush = [bool]$hasFlushDeclaration
            Path = $file.FullName
        })
    }
    return $definitions.ToArray()
}

function Get-HostModuleFiles {
    return @(Get-VbaFiles $srcDir | Where-Object { $_.Name -ne 'ToolRegistry.bas' })
}

function Get-WorkerModuleFiles {
    $common = Join-Path $srcDir 'common'
    $items = @()
    $items += Get-VbaFiles $common | Where-Object { $_.Name -notin @('AppEventHandler.cls', 'ToolRegistry.bas') }
    $items += Get-VbaFiles $toolsDir
    return @($items | Sort-Object FullName)
}

function Extract-VbaCode {
    param([string]$Path)
    $lines = Get-Content -LiteralPath $Path -Encoding ASCII
    $codeLines = New-Object System.Collections.Generic.List[string]
    $inHeader = $true
    foreach ($line in $lines) {
        if ($inHeader) {
            if ($line -match '^Attribute VB_Exposed') { $inHeader = $false }
            continue
        }
        $codeLines.Add($line)
    }
    return ($codeLines -join "`r`n")
}

function Set-CodeModuleText {
    param([object]$CodeModule, [string]$Code)
    if ($CodeModule.CountOfLines -gt 0) { $CodeModule.DeleteLines(1, $CodeModule.CountOfLines) }
    if (-not [string]::IsNullOrWhiteSpace($Code)) { $CodeModule.AddFromString($Code) }
}

function Import-Modules {
    param([object]$Project, [object[]]$Files)
    foreach ($file in $Files) {
        $component = $null
        $codeModule = $null
        try {
            switch ($file.Extension.ToLowerInvariant()) {
                '.bas' { $component = $Project.VBComponents.Import($file.FullName) }
                '.cls' {
                    $component = $Project.VBComponents.Add(2)
                    $component.Name = [IO.Path]::GetFileNameWithoutExtension($file.Name)
                    $codeModule = $component.CodeModule
                    Set-CodeModuleText $codeModule (Extract-VbaCode $file.FullName)
                }
                '.frm' { $component = $Project.VBComponents.Import($file.FullName) }
            }
        } finally {
            Release-ComObject $codeModule
            Release-ComObject $component
        }
    }
}

function Add-GeneratedModule {
    param([object]$Project, [string]$Name, [string]$Code)
    $component = $null
    $codeModule = $null
    try {
        $component = $Project.VBComponents.Add(1)
        $component.Name = $Name
        $codeModule = $component.CodeModule
        Set-CodeModuleText $codeModule $Code
    } finally {
        Release-ComObject $codeModule
        Release-ComObject $component
    }
}

function New-RegistryCode {
    param([object[]]$Tools)
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('Option Explicit')
    $lines.Add('Option Private Module')
    $lines.Add('')
    $lines.Add('Public Sub Registry_Run(ByVal toolId As String, ByVal ctx As InfraContext)')
    $lines.Add('    Select Case LCase$(toolId)')
    foreach ($tool in $Tools) {
        $lines.Add("        Case `"$($tool.Module.ToLowerInvariant())`": $($tool.Module).Run ctx")
    }
    $lines.Add('        Case Else: Err.Raise vbObjectError + 2510, "ToolRegistry.Registry_Run", "Unknown tool: " & toolId')
    $lines.Add('    End Select')
    $lines.Add('End Sub')
    $lines.Add('')
    $lines.Add('Public Sub Registry_Worker(ByVal toolId As String, ByVal job As InfraJob)')
    $lines.Add('    Select Case LCase$(toolId)')
    foreach ($tool in ($Tools | Where-Object HasWorker)) {
        $lines.Add("        Case `"$($tool.Module.ToLowerInvariant())`": $($tool.Module).Worker job")
    }
    $lines.Add('        Case Else: Err.Raise vbObjectError + 2511, "ToolRegistry.Registry_Worker", "Unknown worker tool: " & toolId')
    $lines.Add('    End Select')
    $lines.Add('End Sub')
    $lines.Add('')
    $lines.Add('Public Function Registry_MaxJobs(ByVal toolId As String) As Long')
    $lines.Add('    Select Case LCase$(toolId)')
    foreach ($tool in $Tools) {
        $lines.Add("        Case `"$($tool.Module.ToLowerInvariant())`": Registry_MaxJobs = $($tool.MaxJobs)")
    }
    $lines.Add('        Case Else: Registry_MaxJobs = 3')
    $lines.Add('    End Select')
    $lines.Add('End Function')
    $lines.Add('')
    $lines.Add('Public Function Registry_HasResult(ByVal toolId As String) As Boolean')
    $lines.Add('    Select Case LCase$(toolId)')
    foreach ($tool in ($Tools | Where-Object HasResult)) {
        $lines.Add("        Case `"$($tool.Module.ToLowerInvariant())`": Registry_HasResult = True")
    }
    $lines.Add('    End Select')
    $lines.Add('End Function')
    $lines.Add('')
    $lines.Add('Public Sub Registry_OnResult(ByVal toolId As String, ByVal ctx As InfraContext, ByVal jobId As String, ByVal versionValue As Long)')
    $lines.Add('    Select Case LCase$(toolId)')
    foreach ($tool in ($Tools | Where-Object HasResult)) {
        $lines.Add("        Case `"$($tool.Module.ToLowerInvariant())`": $($tool.Module).OnResult ctx, jobId, versionValue")
    }
    $lines.Add('        Case Else: Err.Raise vbObjectError + 2512, "ToolRegistry.Registry_OnResult", "Tool has no result handler: " & toolId')
    $lines.Add('    End Select')
    $lines.Add('End Sub')
    $lines.Add('')
    $lines.Add('Public Function Registry_HasFlush(ByVal toolId As String) As Boolean')
    $lines.Add('    Select Case LCase$(toolId)')
    foreach ($tool in ($Tools | Where-Object HasFlush)) {
        $lines.Add("        Case `"$($tool.Module.ToLowerInvariant())`": Registry_HasFlush = True")
    }
    $lines.Add('    End Select')
    $lines.Add('End Function')
    $lines.Add('')
    $lines.Add('Public Sub Registry_OnFlush(ByVal toolId As String, ByVal ctx As InfraContext)')
    $lines.Add('    Select Case LCase$(toolId)')
    foreach ($tool in ($Tools | Where-Object HasFlush)) {
        $lines.Add("        Case `"$($tool.Module.ToLowerInvariant())`": $($tool.Module).OnFlush ctx")
    }
    $lines.Add('        Case Else: Err.Raise vbObjectError + 2513, "ToolRegistry.Registry_OnFlush", "Tool has no flush handler: " & toolId')
    $lines.Add('    End Select')
    $lines.Add('End Sub')
    return ($lines -join "`r`n")
}

function Escape-Xml {
    param([string]$Value)
    return [Security.SecurityElement]::Escape($Value)
}

function New-CustomUiXml {
    param([object[]]$Tools)
    $visible = @($Tools | Where-Object { -not $_.Internal })
    $builder = New-Object Text.StringBuilder
    [void]$builder.AppendLine('<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" xmlns:pub="http://n-sakana/pub">')
    [void]$builder.AppendLine('  <ribbon><tabs><tab idQ="pub:XltoolsTab" label="xltoolrack">')
    $groups = @($visible | Group-Object Group)
    $groupIndex = 0
    foreach ($group in $groups) {
        $groupIndex++
        [void]$builder.AppendLine("    <group id=`"xtrGroup$groupIndex`" label=`"$(Escape-Xml $group.Name)`">")
        foreach ($tool in $group.Group) {
            [void]$builder.AppendLine("      <button id=`"$($tool.Module)`" label=`"$(Escape-Xml $tool.Ribbon)`" size=`"large`" imageMso=`"MacroPlay`" onAction=`"Infra_RunTool`" />")
        }
        [void]$builder.AppendLine('    </group>')
    }
    [void]$builder.AppendLine('  </tab></tabs></ribbon>')
    [void]$builder.AppendLine('</customUI>')
    return $builder.ToString()
}

function Set-ThisWorkbookCode {
    param([object]$Project, [string]$Code)
    $component = $null
    $codeModule = $null
    try {
        $component = $Project.VBComponents.Item('ThisWorkbook')
        $codeModule = $component.CodeModule
        Set-CodeModuleText $codeModule $Code
    } finally {
        Release-ComObject $codeModule
        Release-ComObject $component
    }
}

function Get-HostWorkbookCode {
    param([bool]$IsAddin)
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('Option Explicit')
    $lines.Add('Private Sub Workbook_Open()')
    $lines.Add('    HostMain.InitAddin')
    $lines.Add('End Sub')
    if ($IsAddin) {
        $lines.Add('Private Sub Workbook_AddinInstall()')
        $lines.Add('    HostMain.InitAddin')
        $lines.Add('End Sub')
        $lines.Add('Private Sub Workbook_AddinUninstall()')
        $lines.Add('    HostMain.Shutdown')
        $lines.Add('End Sub')
    }
    $lines.Add('Private Sub Workbook_BeforeClose(Cancel As Boolean)')
    $lines.Add('    On Error Resume Next')
    $lines.Add('    HostMain.Shutdown')
    $lines.Add('    Me.Saved = True')
    $lines.Add('    On Error GoTo 0')
    $lines.Add('End Sub')
    return ($lines -join "`r`n")
}

function Get-WorkerWorkbookCode {
    return @'
Option Explicit
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    WorkerBridge.WorkerBeforeClose
    Me.Saved = True
    On Error GoTo 0
End Sub
'@
}

function New-BuiltWorkbook {
    param([object]$Excel, [string]$OutputPath, [object[]]$ModuleFiles,
          [ValidateSet('host-xlam', 'host-xlsm', 'worker')][string]$Kind,
          [string]$RegistryCode)
    if (Test-Path -LiteralPath $OutputPath) { Remove-Item -LiteralPath $OutputPath -Force }
    $workbook = $null
    $project = $null
    $sheet = $null
    try {
        $workbook = $Excel.Workbooks.Add()
        $sheet = $workbook.Worksheets.Item(1)
        $sheet.Name = if ($Kind -eq 'worker') { '_worker' } else { '_host' }
        $project = $workbook.VBProject
        try { $null = $project.VBComponents.Count } catch { throw 'Enable Trust access to the VBA project object model in Excel.' }
        Import-Modules $project $ModuleFiles
        Add-GeneratedModule $project 'ToolRegistry' $RegistryCode
        switch ($Kind) {
            'host-xlam' {
                Set-ThisWorkbookCode $project (Get-HostWorkbookCode $true)
                $workbook.IsAddin = $true
                $workbook.SaveAs($OutputPath, 55)
            }
            'host-xlsm' {
                Set-ThisWorkbookCode $project (Get-HostWorkbookCode $false)
                $workbook.SaveAs($OutputPath, 52)
            }
            'worker' {
                Set-ThisWorkbookCode $project (Get-WorkerWorkbookCode)
                $workbook.SaveAs($OutputPath, 52)
            }
        }
    } finally {
        Release-ComObject $sheet
        Release-ComObject $project
        if ($null -ne $workbook) { try { $workbook.Close($false) } catch {}; Release-ComObject $workbook }
    }
}

function Add-CustomUi {
    param([string]$WorkbookPath, [string]$Xml)
    $tempDir = Join-Path $env:TEMP ('xltoolrack_ribbon_' + [Guid]::NewGuid().ToString('N'))
    $zipPath = $WorkbookPath + '.zip'
    try {
        Copy-Item -LiteralPath $WorkbookPath -Destination $zipPath -Force
        Expand-Archive -LiteralPath $zipPath -DestinationPath $tempDir -Force
        Remove-Item -LiteralPath $zipPath -Force
        $customUiDir = Join-Path $tempDir 'customUI'
        New-Item -ItemType Directory -Path $customUiDir -Force | Out-Null
        [IO.File]::WriteAllText((Join-Path $customUiDir 'customUI14.xml'), $Xml, (New-Object Text.UTF8Encoding($false)))
        $contentTypesPath = Join-Path $tempDir '[Content_Types].xml'
        $contentTypes = [xml](Get-Content -LiteralPath $contentTypesPath -Raw)
        $ns = $contentTypes.DocumentElement.NamespaceURI
        if (-not ($contentTypes.Types.Override | Where-Object PartName -eq '/customUI/customUI14.xml')) {
            $node = $contentTypes.CreateElement('Override', $ns)
            $node.SetAttribute('PartName', '/customUI/customUI14.xml')
            $node.SetAttribute('ContentType', 'application/xml')
            [void]$contentTypes.DocumentElement.AppendChild($node)
            $contentTypes.Save($contentTypesPath)
        }
        $relationshipsPath = Join-Path $tempDir '_rels\.rels'
        $relationships = [xml](Get-Content -LiteralPath $relationshipsPath -Raw)
        $relNs = $relationships.DocumentElement.NamespaceURI
        if (-not ($relationships.Relationships.Relationship | Where-Object Target -eq 'customUI/customUI14.xml')) {
            $node = $relationships.CreateElement('Relationship', $relNs)
            $node.SetAttribute('Id', 'rIdXltoolrackCustomUI')
            $node.SetAttribute('Type', 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility')
            $node.SetAttribute('Target', 'customUI/customUI14.xml')
            [void]$relationships.DocumentElement.AppendChild($node)
            $relationships.Save($relationshipsPath)
        }
        Remove-Item -LiteralPath $WorkbookPath -Force
        New-NormalizedZip $tempDir $zipPath
        Move-Item -LiteralPath $zipPath -Destination $WorkbookPath -Force
    } finally {
        if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue }
        if (Test-Path -LiteralPath $tempDir) { Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

function New-NormalizedZip {
    param([string]$SourceDirectory, [string]$ZipPath)
    if (Test-Path -LiteralPath $ZipPath) { Remove-Item -LiteralPath $ZipPath -Force }
    $basePath = [IO.Path]::GetFullPath($SourceDirectory).TrimEnd('\') + '\'
    $stream = $null
    $archive = $null
    try {
        $stream = [IO.File]::Open($ZipPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $archive = New-Object IO.Compression.ZipArchive -ArgumentList $stream, ([IO.Compression.ZipArchiveMode]::Create), $false
        foreach ($file in (Get-ChildItem -LiteralPath $SourceDirectory -Recurse -File)) {
            $relative = $file.FullName.Substring($basePath.Length).Replace('\', '/')
            $entry = $archive.CreateEntry($relative, [IO.Compression.CompressionLevel]::Optimal)
            $input = $null
            $output = $null
            try {
                $input = $file.OpenRead()
                $output = $entry.Open()
                $input.CopyTo($output)
            } finally {
                if ($null -ne $output) { $output.Dispose() }
                if ($null -ne $input) { $input.Dispose() }
            }
        }
    } finally {
        if ($null -ne $archive) { $archive.Dispose() }
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

Assert-AsciiAndRuntimePolicy
$tools = Get-ToolDefinitions
$registryCode = New-RegistryCode $tools
$customUiXml = New-CustomUiXml $tools
$hostModules = Get-HostModuleFiles
$workerModules = Get-WorkerModuleFiles
New-Item -ItemType Directory -Path $distDir -Force | Out-Null

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    if ($OutputFormat -in @('all', 'xlam')) {
        $xlamPath = Join-Path $distDir 'xltoolrack.xlam'
        New-BuiltWorkbook $excel $xlamPath $hostModules 'host-xlam' $registryCode
        Add-CustomUi $xlamPath $customUiXml
        Write-Host "Built host add-in: $xlamPath" -ForegroundColor Green
    }
    if ($OutputFormat -in @('all', 'xlsm')) {
        $xlsmPath = Join-Path $distDir 'xltoolrack-test.xlsm'
        New-BuiltWorkbook $excel $xlsmPath $hostModules 'host-xlsm' $registryCode
        Write-Host "Built test host: $xlsmPath" -ForegroundColor Green
    }
    if (Test-Path -LiteralPath (Join-Path $srcDir 'common\JobWorker.cls')) {
        $workerPath = Join-Path $distDir 'xltoolrack-worker.xlsm'
        New-BuiltWorkbook $excel $workerPath $workerModules 'worker' $registryCode
        Write-Host "Built worker: $workerPath" -ForegroundColor Green
    }
} finally {
    if ($null -ne $excel) { try { $excel.Quit() } catch {}; Release-ComObject $excel }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
