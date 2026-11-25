<#
.SYNOPSIS
Runs the shared Changes behavior fixtures without a native build or external test framework.
.DESCRIPTION
Creates uniquely owned fixture directories beneath an existing caller-owned scratch root.
Never reads or changes the application's changelog. Removes only its own direct fixture children.
Engine shim: runs Tools/Test-ChangelogWorkflow.ps1 of the shared CI engine that .ci/project.json pins, from
.ci-engine, for this product, which also proves that this product's changelog commands are the engine's shims and
that its PrePush runner validates the changelog first. Refuses with one message when that checkout is missing, at
another commit or changed. The engine's Specs/ProjectCommands.md is the contract.
.PARAMETER ScratchRoot
Existing test scratch parent supplied by the repository's test harness.
.OUTPUTS
Console case results; throws if any assertion fails. Fixture files are removed on completion.
.NOTES
Requires PowerShell 7, Git and the pinned engine at .ci-engine. Its local fixture repositories use a fixed identity
and none of the person's hooks or signing. No network, native code, application processes or desktop input.
.EXAMPLE
./Tools/Test-ChangelogWorkflow.ps1 -ScratchRoot <existing-test-scratch-directory>
#>
[CmdletBinding()]
param([Parameter(Mandatory)][string] $ScratchRoot)
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
$PSBoundParameters['ProjectRoot'] = $project
& $command @PSBoundParameters
