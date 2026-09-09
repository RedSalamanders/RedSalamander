<#
.SYNOPSIS
Freezes an exact Ghostty runtime candidate into bounded review evidence.

.DESCRIPTION
Validates an explicit pristine checkout, two byte-identical official archive downloads, an exact Candidate status report, official GitHub commit verification, ancestry from the canonical product lock, the complete lib-vt path-filtered change inventory, public C declarations/enums, every loader-required export, and security-relevant primary commits. Publishes canonical schema-validated review and descriptor files below the exact candidate .build root without changing Product or tracked state.

.PARAMETER CandidateCommit
The exact lowercase full Ghostty commit to freeze.

.PARAMETER SourcePath
A pristine detached checkout of the official Ghostty repository at CandidateCommit, below the exact candidate .build root.

.PARAMETER ArchivePath
The first official GitHub source archive download for CandidateCommit, below the exact candidate .build root.

.PARAMETER RepeatArchivePath
A separately downloaded copy of the same official archive. Its SHA-256 must exactly match ArchivePath.

.PARAMETER ObservationReportPath
A canonical Candidate report produced by Get-GhosttyTerminalRuntimeStatus.ps1 for the exact commit/tree.

.PARAMETER VerificationFixturePath
Injects a bounded GitHub commit API fixture for deterministic tests. Production review omits it and queries the official API.

.PARAMETER GeneratedUtc
Overrides the evidence timestamp for deterministic tests. Production callers omit it.

.OUTPUTS
A PSCustomObject containing the authenticated candidate identity, review counts, and canonical report paths.

.NOTES
Prerequisites: PowerShell 7.4 or newer, Git, an exact official checkout/archive pair, the canonical Ghostty lock and schemas, and network access to the official GitHub commit API unless a test fixture is supplied. Side effects are confined to creating or replacing candidate-review.json and candidate-descriptor.json below .build/TerminalEngineUpgrade/<commit>. The command never edits Product, the canonical lock, source files, branches, issues, or pull requests.

.EXAMPLE
.\Tools\New-GhosttyTerminalRuntimeCandidateReview.ps1 -CandidateCommit 9c3ec931d64561a8407dde7ac984ce156ae91539 -SourcePath .build\TerminalEngineUpgrade\9c3ec931d64561a8407dde7ac984ce156ae91539\source-review -ArchivePath .build\TerminalEngineUpgrade\9c3ec931d64561a8407dde7ac984ce156ae91539\ghostty-9c3ec931d64561a8407dde7ac984ce156ae91539.tar.gz -RepeatArchivePath .build\TerminalEngineUpgrade\9c3ec931d64561a8407dde7ac984ce156ae91539\ghostty-9c3ec931d64561a8407dde7ac984ce156ae91539-repeat.tar.gz -ObservationReportPath .build\TerminalEngineStatus\candidate-freeze\status.json
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$CandidateCommit,
    [Parameter(Mandatory = $true)][string]$SourcePath,
    [Parameter(Mandatory = $true)][string]$ArchivePath,
    [Parameter(Mandatory = $true)][string]$RepeatArchivePath,
    [Parameter(Mandatory = $true)][string]$ObservationReportPath,
    [string]$VerificationFixturePath,
    [DateTime]$GeneratedUtc = [DateTime]::UtcNow
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
Import-Module (Join-Path $repoRoot 'Tools\TerminalEngine\GhosttyCandidateReview.psm1') -Force -ErrorAction Stop

New-RSGhosttyRuntimeCandidateReview `
    -RepoRoot $repoRoot `
    -CandidateCommit $CandidateCommit `
    -SourcePath $SourcePath `
    -ArchivePath $ArchivePath `
    -RepeatArchivePath $RepeatArchivePath `
    -ObservationReportPath $ObservationReportPath `
    -VerificationFixturePath $VerificationFixturePath `
    -GeneratedUtc $GeneratedUtc
