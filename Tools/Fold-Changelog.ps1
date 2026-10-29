<#
.SYNOPSIS
Folds pending Changes fragments into CHANGELOG.md, even when only one is pending.
.DESCRIPTION
Validates all input before writing, folds newest filenames first under Unreleased,
then removes only folded files. Preserves UTF-8 BOM and changelog line endings.
Fragments are folded on the default branch after merge, normally by Tools/Submit-ChangelogFold.ps1;
a feature change keeps its fragment. An interrupted cleanup can be retried without duplicating
already published entries.
.PARAMETER Root
Repository root; defaults to the repository containing this command.
.PARAMETER Check
Read-only check of a finished fold: fails if any fragment remains or the changelog is malformed.
.OUTPUTS
Console summary. Replaces CHANGELOG.md and removes folded files unless Check or WhatIf is used.
.NOTES
Requires PowerShell 7. No Git, network, build, staging, commits or pushes.
Invalid input or I/O failure raises a terminating error/nonzero process exit.
Rejects reparse paths inside the repository. A checkout-local exclusive lease serializes folds.
Primary consumers: the fold submission and integration owners.
.EXAMPLE
./Tools/Fold-Changelog.ps1 -WhatIf
.EXAMPLE
./Tools/Fold-Changelog.ps1
.EXAMPLE
./Tools/Fold-Changelog.ps1 -Check
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Low')]
param([string] $Root = (Split-Path $PSScriptRoot -Parent), [switch] $Check)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'Modules/Documentation/Changelog.psm1')
$Root = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Root)
$problems = Get-ChangelogProblems $Root -RequireFolded:$Check
if ($problems.Count) { throw "Changelog validation failed:`n$($problems -join "`n")" }
if ($Check) { Write-Host 'Changelog is valid and fully folded.'; return }
$fragments = Get-ChangelogFragments $Root
if (-not $fragments.Count) { Write-Host 'No changelog fragments to fold.'; return }
if ($PSCmdlet.ShouldProcess($Root, "Fold $($fragments.Count) changelog fragment(s) and remove their source files")) {
    $count = Invoke-ChangelogFold $Root
    Write-Host "Folded $count changelog fragment(s) into CHANGELOG.md, newest first."
}
