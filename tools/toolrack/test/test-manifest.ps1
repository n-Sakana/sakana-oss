# test/test-manifest.ps1 -- unit tests for manifest validation
. (Join-Path $PSScriptRoot "_assert.ps1")
. (Join-Path $PSScriptRoot "..\common\install.ps1")   # dot-source: functions only

$fx = Join-Path $env:TEMP ("toolrack_fx_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force $fx | Out-Null
try {

function New-Fixture {
    param([string]$Name, [string]$Json, [string[]]$ExtraFiles = @())
    $d = Join-Path $fx $Name
    New-Item -ItemType Directory -Force $d | Out-Null
    [System.IO.File]::WriteAllText((Join-Path $d "tool.json"), $Json)
    foreach ($f in $ExtraFiles) { [System.IO.File]::WriteAllText((Join-Path $d $f), "# stub") }
    return $d
}

# -- valid minimal
$d = New-Fixture "demo" '{"schema":1,"id":"demo","name":"Demo","on":["folder"],"run":{"type":"powershell","entry":"main.ps1"}}' @("main.ps1")
$r = Read-Manifest $d
Assert-True $r.Ok "valid: parses"
$e = @(Test-Manifest $r.Data "demo" $d)
Assert-True ($e.Count -eq 0) "valid: no violations"

# -- valid hidden/no-console tool
$d = New-Fixture "hidden" '{"schema":1,"id":"hidden","name":"Hidden","on":["file"],"run":{"type":"powershell","entry":"main.ps1","window":"hidden"}}' @("main.ps1")
$r = Read-Manifest $d
$e = @(Test-Manifest $r.Data "hidden" $d)
Assert-True ($e.Count -eq 0) "valid: hidden window mode"

# -- violations (one per rule)
$cases = @(
    @{ n="badschema"; j='{"schema":9,"id":"badschema","name":"x","on":["file"],"run":{"type":"powershell","entry":"main.ps1"}}'; f=@("main.ps1"); pat="schema*" },
    @{ n="strschema"; j='{"schema":"1","id":"strschema","name":"x","on":["file"],"run":{"type":"powershell","entry":"main.ps1"}}'; f=@("main.ps1"); pat="schema*" },
    @{ n="badid";     j='{"schema":1,"id":"other","name":"x","on":["file"],"run":{"type":"powershell","entry":"main.ps1"}}'; f=@("main.ps1"); pat="id*" },
    @{ n="noname";    j='{"schema":1,"id":"noname","name":"","on":["file"],"run":{"type":"powershell","entry":"main.ps1"}}'; f=@("main.ps1"); pat="name*" },
    @{ n="ctrlname";  j='{"schema":1,"id":"ctrlname","name":"bad\nname","on":["file"],"run":{"type":"powershell","entry":"main.ps1"}}'; f=@("main.ps1"); pat="name*" },
    @{ n="badon";     j='{"schema":1,"id":"badon","name":"x","on":["desktop"],"run":{"type":"powershell","entry":"main.ps1"}}'; f=@("main.ps1"); pat="on*" },
    @{ n="scalaron";  j='{"schema":1,"id":"scalaron","name":"x","on":"file","run":{"type":"powershell","entry":"main.ps1"}}'; f=@("main.ps1"); pat="on*" },
    @{ n="badtype";   j='{"schema":1,"id":"badtype","name":"x","on":["file"],"run":{"type":"bash","entry":"main.sh"}}'; f=@("main.sh"); pat="run.type*" },
    @{ n="noentry";   j='{"schema":1,"id":"noentry","name":"x","on":["file"],"run":{"type":"powershell","entry":"main.ps1"}}'; f=@(); pat="run.entry*" },
    @{ n="escentry";  j='{"schema":1,"id":"escentry","name":"x","on":["file"],"run":{"type":"powershell","entry":"..\\..\\evil.ps1"}}'; f=@(); pat="run.entry*" },
    @{ n="badwin";    j='{"schema":1,"id":"badwin","name":"x","on":["file"],"run":{"type":"powershell","entry":"main.ps1","window":"tray"}}'; f=@("main.ps1"); pat="run.window*" },
    @{ n="reqps";     j='{"schema":1,"id":"reqps","name":"x","on":["file"],"run":{"type":"powershell","entry":"main.ps1","requires":["pandas"]}}'; f=@("main.ps1"); pat="run.requires*" },
    @{ n="badreq";    j='{"schema":1,"id":"badreq","name":"x","on":["file"],"run":{"type":"python","entry":"main.py","requires":["bad-name"]}}'; f=@("main.py"); pat="run.requires*" },
    @{ n="scalarreq"; j='{"schema":1,"id":"scalarreq","name":"x","on":["file"],"run":{"type":"python","entry":"main.py","requires":"json"}}'; f=@("main.py"); pat="run.requires*" },
    @{ n="dupvar";    j='{"schema":1,"id":"dupvar","name":"x","on":["file"],"run":{"type":"powershell","entry":"main.ps1"},"variants":[{"label":"A","args":[]},{"label":"A","args":[]}]}'; f=@("main.ps1"); pat="variants*" },
    @{ n="ctrllabel"; j='{"schema":1,"id":"ctrllabel","name":"x","on":["file"],"run":{"type":"powershell","entry":"main.ps1"},"variants":[{"label":"bad\u0001label","args":[]}]}'; f=@("main.ps1"); pat="variants*" },
    @{ n="scalarvar"; j='{"schema":1,"id":"scalarvar","name":"x","on":["file"],"run":{"type":"powershell","entry":"main.ps1"},"variants":{"label":"A","args":[]}}'; f=@("main.ps1"); pat="variants*" },
    @{ n="scalararg"; j='{"schema":1,"id":"scalararg","name":"x","on":["file"],"run":{"type":"powershell","entry":"main.ps1"},"variants":[{"label":"A","args":"-X"}]}'; f=@("main.ps1"); pat="variants*" },
    @{ n="emptyarg";  j='{"schema":1,"id":"emptyarg","name":"x","on":["file"],"run":{"type":"powershell","entry":"main.ps1"},"variants":[{"label":"A","args":[""]}]}'; f=@("main.ps1"); pat="variants*" },
    @{ n="quotearg";  j='{"schema":1,"id":"quotearg","name":"x","on":["file"],"run":{"type":"powershell","entry":"main.ps1"},"variants":[{"label":"A","args":["-X","a\"b"]}]}'; f=@("main.ps1"); pat="variants*" }
)
foreach ($c in $cases) {
    $d = New-Fixture $c.n $c.j $c.f
    $r = Read-Manifest $d
    $e = @()
    if ($r.Ok) { $e = @(Test-Manifest $r.Data $c.n $d) }
    else { $e = @($r.Errors) }
    Assert-True ($e.Count -gt 0) ("violation detected: " + $c.n)
    Assert-Contains $e $c.pat ("message mentions rule: " + $c.n)
}

# -- broken json / missing tool.json
$d = New-Fixture "brokenjson" '{"schema":1,'
$r = Read-Manifest $d
Assert-True (-not $r.Ok) "broken json rejected"
$d2 = Join-Path $fx "nomanifest"; New-Item -ItemType Directory -Force $d2 | Out-Null
$r2 = Read-Manifest $d2
Assert-True (-not $r2.Ok) "missing tool.json rejected"

} finally {
    Remove-Item -LiteralPath $fx -Recurse -Force -ErrorAction SilentlyContinue
}
Exit-Test
