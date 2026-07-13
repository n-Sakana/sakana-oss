$ErrorActionPreference = 'Stop'
$testDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$cases = @(
    @{ File = 'test-appguard.ps1'; Args = @() },
    @{ File = 'test-logstatus.ps1'; Args = @() },
    @{ File = 'test-context.ps1'; Args = @() },
    @{ File = 'test-harness.ps1'; Args = @() },
    @{ File = 'test-result-dispatch.ps1'; Args = @() },
    @{ File = 'test-build-validation.ps1'; Args = @() },
    @{ File = 'test-ribbon-ui.ps1'; Args = @() },
    @{ File = 'test-jobhost-concurrency.ps1'; Args = @() },
    @{ File = 'test-jobhost-caps.ps1'; Args = @() },
    @{ File = 'test-jobhost-rollback.ps1'; Args = @() },
    @{ File = 'test-worker-loop.ps1'; Args = @() },
    @{ File = 'test-worker-fe-close.ps1'; Args = @() },
    @{ File = 'test-worker-busy-fe.ps1'; Args = @() },
    @{ File = 'test-samples.ps1'; Args = @('-Sample','all') },
    @{ File = 'Validate-E2E.ps1'; Args = @() }
)

foreach ($case in $cases) {
    Write-Host "=== $($case.File) ===" -ForegroundColor Cyan
    & $powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $testDir $case.File) @($case.Args)
    if ($LASTEXITCODE -ne 0) { throw "$($case.File) failed with exit code $LASTEXITCODE" }
}

Write-Host 'ALL COM HARNESSES PASSED' -ForegroundColor Green
