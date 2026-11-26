<#
.SYNOPSIS
Folds pending Changes fragments into CHANGELOG.md, even when only one is pending.
.DESCRIPTION
Validates all input before writing, folds newest filenames first under Unreleased,
then removes only folded files. Preserves UTF-8 BOM and changelog line endings.
Fragments are folded on the default branch after merge, normally by Tools/Submit-ChangelogFold.ps1;
a feature change keeps its fragment. An interrupted cleanup can be retried without duplicating
already published entries; skipped fragment names are reported as warnings. Two pending fragments
with the same new text are refused, since one entry would stand for both changes.
Engine shim: runs Tools/Fold-Changelog.ps1 of the shared CI engine that .ci/project.json pins, from .ci-engine,
and refuses with one message when that checkout is missing, at another commit or changed. The engine's
Specs/ProjectCommands.md is the contract.
.PARAMETER Root
Repository root; defaults to the repository containing this command.
.PARAMETER Check
Read-only check of a finished fold: fails if any fragment remains or the changelog is malformed.
.OUTPUTS
Console summary. Replaces CHANGELOG.md and removes folded files unless Check or WhatIf is used.
.NOTES
Requires PowerShell 7, Git to verify the engine, and the pinned engine at .ci-engine. No network, build, staging,
commits or pushes. Invalid input or I/O failure raises a terminating error/nonzero process exit.
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
# Engine shim (Specs/ProjectCommands.md in the engine); the same lines in every shim.
# The product root holds .ci/project.json: the shim's own folder, or the one above it for a shim under Tools/.
$project = if ([IO.File]::Exists((Join-Path $PSScriptRoot '.ci/project.json'))) { $PSScriptRoot } else { Split-Path $PSScriptRoot -Parent }
$binding = Join-Path $project '.ci-engine/Tools/Modules/Ci.EngineBinding.psm1'
if (-not [IO.File]::Exists($binding)) {
    throw "$($MyInvocation.MyCommand.Name) runs the shared CI engine that .ci/project.json pins from .ci-engine, and $project/.ci-engine is missing or predates engine shims. Check out the pinned commit (engine.ref in .ci/project.json): git -C .ci-engine fetch origin <engine.ref>; git -C .ci-engine checkout --detach <engine.ref> (first time: git clone https://github.com/RedSalamanders/.github.git .ci-engine)."
}
Import-Module $binding -Force
$command = Resolve-CiPinnedEngineCommand -ProjectRoot $project -Name $MyInvocation.MyCommand.Name
$PSBoundParameters['Root'] = $Root
& $command @PSBoundParameters
