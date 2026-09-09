<#
.SYNOPSIS
Validates and records one rederived Ghostty candidate overlay.

.DESCRIPTION
Proves that one normalized candidate patch applies to a pristine exact Ghostty commit, exactly reproduces an independently applied source diff, stays within review-size and concern limits, maps every current overlay concern, and allocates every newly added private C/Zig enum value after the complete candidate public ranges. Publishes a strict overlay review and updates only the untracked candidate descriptor below the exact candidate .build root.

.PARAMETER CandidateCommit
The exact lowercase full Ghostty candidate commit.

.PARAMETER PristineSourcePath
A clean exact candidate checkout below .build/TerminalEngineUpgrade/<commit>.

.PARAMETER AppliedSourcePath
An exact candidate checkout containing only the proposed overlay diff.

.PARAMETER PatchPath
The single normalized candidate overlay patch below the exact candidate root.

.PARAMETER DescriptorPath
The candidate-descriptor.json produced by the candidate-freeze workflow.

.PARAMETER ConcernReviewPath
A reviewer-authored JSON file mapping every current concern to one of at most five ordered candidate concerns.

.PARAMETER GeneratedUtc
Overrides the evidence timestamp for deterministic tests. Production callers omit it.

.OUTPUTS
A PSCustomObject containing patch/review/descriptor identities and measured overlay size.

.NOTES
Prerequisites: PowerShell 7.4 or newer, Git, the exact candidate freeze descriptor, pristine and applied candidate worktrees, and the single candidate patch. Side effects are confined to candidate-descriptor.p4.json, candidate-overlay-review.json, and candidate-descriptor.json below .build/TerminalEngineUpgrade/<commit>. Product, the canonical lock, and tracked patch remain unchanged.

.EXAMPLE
.\Tools\New-GhosttyTerminalRuntimeCandidateOverlayReview.ps1 -CandidateCommit 9c3ec931d64561a8407dde7ac984ce156ae91539 -PristineSourcePath .build\TerminalEngineUpgrade\9c3ec931d64561a8407dde7ac984ce156ae91539\source-review -AppliedSourcePath .build\TerminalEngineUpgrade\9c3ec931d64561a8407dde7ac984ce156ae91539\overlay-source -PatchPath .build\TerminalEngineUpgrade\9c3ec931d64561a8407dde7ac984ce156ae91539\Ghostty-WindowsPrivateRuntime.patch -DescriptorPath .build\TerminalEngineUpgrade\9c3ec931d64561a8407dde7ac984ce156ae91539\candidate-descriptor.json -ConcernReviewPath .build\TerminalEngineUpgrade\9c3ec931d64561a8407dde7ac984ce156ae91539\candidate-overlay-concerns.json
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$CandidateCommit,
    [Parameter(Mandatory = $true)][string]$PristineSourcePath,
    [Parameter(Mandatory = $true)][string]$AppliedSourcePath,
    [Parameter(Mandatory = $true)][string]$PatchPath,
    [Parameter(Mandatory = $true)][string]$DescriptorPath,
    [Parameter(Mandatory = $true)][string]$ConcernReviewPath,
    [DateTime]$GeneratedUtc = [DateTime]::UtcNow
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
Import-Module (Join-Path $repoRoot 'Tools\TerminalEngine\GhosttyCandidateOverlay.psm1') -Force -ErrorAction Stop

New-RSGhosttyRuntimeCandidateOverlayReview `
    -RepoRoot $repoRoot `
    -CandidateCommit $CandidateCommit `
    -PristineSourcePath $PristineSourcePath `
    -AppliedSourcePath $AppliedSourcePath `
    -PatchPath $PatchPath `
    -DescriptorPath $DescriptorPath `
    -ConcernReviewPath $ConcernReviewPath `
    -GeneratedUtc $GeneratedUtc
