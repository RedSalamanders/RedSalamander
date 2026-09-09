<#
.SYNOPSIS
Builds and verifies the locked Ghostty private Terminal runtime.

.DESCRIPTION
Invokes the lock-driven Ghostty lifecycle module. Product mode consumes the canonical tracked lock and alone may publish .build\TerminalEngine\Product plus its generated pairing header. Candidate mode requires an explicit lock and output below .build\TerminalEngineUpgrade and cannot write Product or a pairing header.

.PARAMETER Platform
Selects the x64 or ARM64 runtime architecture.

.PARAMETER Force
Rebuilds the runtime even when the selected output and its complete identity pass validation.

.PARAMETER Offline
Forbids downloads and requires all locked source and toolchain inputs to exist in the selected lifecycle cache.

.PARAMETER CandidateLockPath
Enables candidate mode with an explicit lock file below .build\TerminalEngineUpgrade. Candidate mode never falls back to the canonical product lock.

.PARAMETER CandidateOutputRoot
Requires the candidate output root below .build\TerminalEngineUpgrade. It must be supplied together with CandidateLockPath.

.PARAMETER PassThru
Returns the parsed terminal-runtime-identity object instead of only writing a status message.

.OUTPUTS
With -PassThru, a PSCustomObject containing the verified runtime identity. Otherwise no pipeline output.

.NOTES
Prerequisites: PowerShell 7, Git, the canonical schema, the selected lock, locked source/toolchain caches or network access when not offline, and the reviewed overlay/notices. Side effects: Product mode may update .build\TerminalEngine; candidate mode may update only .build\TerminalEngineUpgrade. Native children are kill-on-close contained. An exact independently owned output blocks replacement and is never killed.

The delegated lifecycle preserves runtime identity schemaVersion = 5 and pairingContract = 'generated-raw-sha256-v1'. GhosttyRuntimeLifecycle.psm1 owns Get-TerminalRuntimeIdentityHeaderText, $generatedIdentityHeader publication in Product mode, and $expectedLicenseHashes verification.

.EXAMPLE
.\Tools\Build-GhosttyTerminalRuntime.ps1 -Platform x64 -Offline -PassThru

.EXAMPLE
.\Tools\Build-GhosttyTerminalRuntime.ps1 -Platform x64 -CandidateLockPath .build\TerminalEngineUpgrade\candidate.lock.json -CandidateOutputRoot .build\TerminalEngineUpgrade\Candidate\x64 -Force -PassThru
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet('x64', 'ARM64')]
    [string]$Platform,

    [switch]$Force,

    [switch]$Offline,

    [string]$CandidateLockPath,

    [string]$CandidateOutputRoot,

    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$ghosttyLifecycleModule = Join-Path $repoRoot 'Tools\TerminalEngine\GhosttyRuntimeLifecycle.psm1'
Import-Module $ghosttyLifecycleModule -Force -ErrorAction Stop

Invoke-RSGhosttyRuntimeBuild `
    -RepoRoot $repoRoot `
    -Platform $Platform `
    -Force:$Force `
    -Offline:$Offline `
    -PassThru:$PassThru `
    -CandidateLockPath $CandidateLockPath `
    -CandidateOutputRoot $CandidateOutputRoot
