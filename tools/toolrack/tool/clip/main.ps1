# Clip -- copy a preset text snippet to the clipboard (config-driven, hidden window).
# Snippets live in snippets.json (key -> label + text). The variant passes only -Key,
# so arbitrary text (quotes, spaces, newlines) stays in the config, never in the registry.
param(
    [string]$Target = "$PWD",
    [string]$Key = ""
)
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

function Show-ClipError {
    param([string]$Message)
    [void][System.Windows.Forms.MessageBox]::Show($Message, "Clip", "OK", "Error")
    exit 1
}

function Show-ClipBalloon {
    param([string]$Text)
    try {
        $ni = New-Object System.Windows.Forms.NotifyIcon
        $ni.Icon = [System.Drawing.SystemIcons]::Information
        $ni.Visible = $true
        $ni.ShowBalloonTip(1200, "Copied to clipboard", $Text, [System.Windows.Forms.ToolTipIcon]::Info)
        Start-Sleep -Milliseconds 1400
        $ni.Visible = $false
        $ni.Dispose()
    } catch { }
}

$cfgPath = Join-Path $PSScriptRoot "snippets.json"
if (-not (Test-Path -LiteralPath $cfgPath)) { Show-ClipError "Config not found: $cfgPath" }

$cfg = $null
try { $cfg = ConvertFrom-Json ([System.IO.File]::ReadAllText($cfgPath)) }
catch { Show-ClipError ("snippets.json is not valid JSON: " + $_.Exception.Message) }

$list = @()
if ($cfg -and $cfg.PSObject.Properties["snippets"]) { $list = @($cfg.snippets) }
$hit = $list | Where-Object { $_.key -eq $Key } | Select-Object -First 1
if (-not $hit) { Show-ClipError ("Snippet key not found: '" + $Key + "'") }

$text = [string]$hit.text
try { Set-Clipboard -Value $text -ErrorAction Stop }
catch { Show-ClipError ("Clipboard write failed: " + $_.Exception.Message) }

Show-ClipBalloon -Text ([string]$hit.label)
exit 0
