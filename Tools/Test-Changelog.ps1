<#
.SYNOPSIS
Validates the shared Changes workflow without writing files.
.DESCRIPTION
Reports fragment filename/content and changelog-heading problems. Editing permits valid
pending fragments; RequireFolded enforces the PrePush/final-handoff gate.
.PARAMETER Root
Repository root; defaults to the repository containing this command.
.PARAMETER RequireFolded
Also rejects pending fragments, even one. Equivalent to Fold-Changelog.ps1 -Check.
.OUTPUTS
Console summary; no persistent files. Invalid input raises a terminating error/nonzero process exit.
.NOTES
Requires PowerShell 7. Read-only; no Git, network, build or desktop access.
Primary consumers: repository validation, tooling tests and final-handoff checks.
.EXAMPLE
./Tools/Test-Changelog.ps1
.EXAMPLE
./Tools/Test-Changelog.ps1 -RequireFolded
#>
[CmdletBinding()]
param([string] $Root = (Split-Path $PSScriptRoot -Parent), [switch] $RequireFolded)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'Modules/Documentation/Changelog.psm1')
$Root = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Root)
$problems = Get-ChangelogProblems $Root -RequireFolded:$RequireFolded
if ($problems.Count) { throw "Changelog validation failed:`n$($problems -join "`n")" }
$fragments = Get-ChangelogFragments $Root
Write-Host "Changelog valid: $($fragments.Count) pending fragment(s)."
