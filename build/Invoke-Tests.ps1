#Requires -Version 5.1
<#
.SYNOPSIS
    Run every tool's Pester suite, each in its own process.

.DESCRIPTION
    openslot uses Pester 3.x (ships with Windows PowerShell) and
    outlook-signature uses Pester 5.x+. The two major versions cannot be
    loaded in the same session, so each suite runs in a separate child
    process with the Pester version it expects.

    Exit code is 0 only if every suite passes.

.PARAMETER Tool
    Limit the run to one tool ('openslot' or 'outlook-signature').
    Omit to run all.
#>
[CmdletBinding()]
param(
    [ValidateSet('openslot', 'outlook-signature')]
    [string] $Tool
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$failed = @()

function Invoke-Suite {
    param(
        [string] $Name,
        [string] $Exe,
        [string] $Body
    )
    Write-Host ''
    Write-Host "=== $Name ===" -ForegroundColor Cyan

    $exePath = (Get-Command $Exe -ErrorAction SilentlyContinue).Source
    if (-not $exePath) {
        Write-Warning "$Exe not found; skipping $Name"
        return $true
    }

    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("ot-tests-{0}.ps1" -f [guid]::NewGuid().ToString('N'))
    Set-Content -LiteralPath $tmp -Value $Body -Encoding UTF8
    try {
        # Pipe through Out-Host so the child's stdout goes straight to the
        # console instead of being captured by the caller's expression.
        & $exePath -NoProfile -NonInteractive -File $tmp | Out-Host
        return ($LASTEXITCODE -eq 0)
    }
    finally {
        Remove-Item -LiteralPath $tmp -ErrorAction SilentlyContinue
    }
}

# --- openslot: Pester 3.x via Windows PowerShell ---------------------------
if (-not $Tool -or $Tool -eq 'openslot') {
    $dir = Join-Path $repoRoot 'tools/openslot/tests'
    $body = @"
`$m = Get-Module -ListAvailable Pester |
    Where-Object { `$_.Version.Major -eq 3 } | Sort-Object Version | Select-Object -Last 1
if (`$m) { Import-Module `$m.Path } else { Import-Module Pester }
`$r = Invoke-Pester -Path '$dir' -PassThru
Write-Host ('Passed: ' + `$r.PassedCount + '  Failed: ' + `$r.FailedCount)
exit `$r.FailedCount
"@
    if (-not (Invoke-Suite -Name 'openslot (Pester 3.x)' -Exe 'powershell' -Body $body)) {
        $failed += 'openslot'
    }
}

# --- outlook-signature: Pester 5.x+ ---------------------------------------
if (-not $Tool -or $Tool -eq 'outlook-signature') {
    $dir = Join-Path $repoRoot 'tools/outlook-signature/tests'
    $body = @"
Import-Module Pester -MinimumVersion 5.0
`$cfg = New-PesterConfiguration
`$cfg.Run.Path = '$dir'
`$cfg.Run.PassThru = `$true
`$cfg.Output.Verbosity = 'Detailed'
`$r = Invoke-Pester -Configuration `$cfg
Write-Host ('Passed: ' + `$r.PassedCount + '  Failed: ' + `$r.FailedCount)
exit `$r.FailedCount
"@
    $exe = if (Get-Command pwsh -ErrorAction SilentlyContinue) { 'pwsh' } else { 'powershell' }
    if (-not (Invoke-Suite -Name 'outlook-signature (Pester 5.x+)' -Exe $exe -Body $body)) {
        $failed += 'outlook-signature'
    }
}

Write-Host ''
if ($failed.Count -gt 0) {
    Write-Host ("FAILED: {0}" -f ($failed -join ', ')) -ForegroundColor Red
    exit 1
}
Write-Host 'All suites passed.' -ForegroundColor Green
exit 0
