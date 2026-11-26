<#
.SYNOPSIS
Generates one release heading from the existing folded Unreleased notes.
.DESCRIPTION
Requires an already fully folded changelog, preserves previous release history and atomically publishes the exact
version/date transition. The result is reviewed as a distinct release change; feature changes never run this command.
Engine shim: runs Tools/Release-Changelog.ps1 of the shared CI engine that .ci/project.json pins, from .ci-engine,
and refuses with one message when that checkout is missing, at another commit or changed. The engine's
Specs/ProjectCommands.md is the contract.
.PARAMETER Version
Explicit semantic version, such as 1.2.3.
.PARAMETER Date
Explicit release date in yyyy-MM-dd.
.PARAMETER Root
Repository root; defaults to the repository containing this command.
.OUTPUTS
CHANGELOG.md only, unless WhatIf is specified. No staging, commit, network, build or desktop action.
.NOTES
Requires PowerShell 7, Git to verify the engine, and the pinned engine at .ci-engine.
.EXAMPLE
./Tools/Release-Changelog.ps1 -Version 1.2.3 -Date 2026-11-01 -WhatIf
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Low')]
param([Parameter(Mandatory)][string] $Version, [Parameter(Mandatory)][string] $Date, [string] $Root = (Split-Path $PSScriptRoot -Parent))
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
