<#
.SYNOPSIS
Validates the shared Changes workflow without writing files.
.DESCRIPTION
Reports fragment filename/content and changelog-heading problems; valid pending fragments are expected.
Base also compares the checkout with its merge base: fragments may be added, and CHANGELOG.md may
change only by folding fragments that base already holds (the fold on the default branch), or by the exact
generated release transition of Release-Changelog.ps1 after every fragment is folded.
RequireFolded rejects any pending fragment, for checking a finished fold.
Engine shim: runs Tools/Test-Changelog.ps1 of the shared CI engine that .ci/project.json pins, from .ci-engine,
and refuses with one message when that checkout is missing, at another commit or changed. The engine's
Specs/ProjectCommands.md is the contract.
.PARAMETER Root
Repository root; defaults to the repository containing this command.
.PARAMETER Base
Ref whose merge base with HEAD the checkout is compared with: origin/main before a push; in CI, HEAD^1 for a pull
request's merge commit and the replaced tip (github.event.before) for a push.
.PARAMETER RequireFolded
Also rejects pending fragments, even one. Equivalent to Fold-Changelog.ps1 -Check.
.OUTPUTS
Console summary; no persistent files. Invalid input raises a terminating error/nonzero process exit.
.NOTES
Requires PowerShell 7, Git, and the pinned engine at .ci-engine. Read-only; no network, build or desktop access.
Primary consumers: repository validation, PrePush, CI, tooling tests and fold checks.
.EXAMPLE
./Tools/Test-Changelog.ps1
.EXAMPLE
./Tools/Test-Changelog.ps1 -Base origin/main
#>
[CmdletBinding()]
param([string] $Root = (Split-Path $PSScriptRoot -Parent), [string] $Base = '', [switch] $RequireFolded)
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
