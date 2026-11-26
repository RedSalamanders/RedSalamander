<#
.SYNOPSIS
Folds the Changes fragments merged to the default branch on a fold branch and opens its pull request.
.DESCRIPTION
Feature changes only add fragments; they are folded here, after they reach the default branch. Fetches that
branch, folds it in a temporary worktree outside this checkout, verifies the result is an exact fold, commits it,
pushes changelog/fold-<commit> and opens its pull request with gh. This checkout's branch, index and files never
change. Nothing is pushed when the branch has no fragments. A fold branch already pushed for the same commit is
reused only after its single parent and exact fold are verified, and an open fold pull request of this repository
(a fork's is ignored) is reported instead of opening a second, conflicting one. The pull request takes the
repository's normal review and merge; merge it before folding again.
Engine shim: runs Tools/Submit-ChangelogFold.ps1 of the shared CI engine that .ci/project.json pins, from
.ci-engine, and refuses with one message when that checkout is missing, at another commit or changed. The
engine's Specs/ProjectCommands.md is the contract.
.PARAMETER Root
Repository root; defaults to the repository containing this command.
.PARAMETER Remote
Remote that holds the default branch; origin by default.
.PARAMETER Branch
Default branch to fold; main by default.
.PARAMETER Repository
GitHub repository (owner/name) of the pull request; read from the remote's URL by default.
.PARAMETER PushOnly
Pushes the fold branch without opening its pull request, for a remote gh cannot reach.
.PARAMETER WorktreeParent
Existing folder that receives the temporary fold worktree; the user's temporary folder by default. Test harnesses
pass a folder inside their own sandbox.
.OUTPUTS
Console summary naming the fold branch and pull request. WhatIf folds and commits only in the temporary
worktree, prints that commit, and neither pushes nor opens a pull request, also for a fold branch pushed earlier.
.NOTES
Requires PowerShell 7 and Git, the pinned engine at .ci-engine, and an authenticated gh unless PushOnly. Uses the
network for git fetch/push and the GitHub API. Commits with the repository's own Git identity, signing and hooks.
.EXAMPLE
./Tools/Submit-ChangelogFold.ps1 -WhatIf
.EXAMPLE
./Tools/Submit-ChangelogFold.ps1
#>
[CmdletBinding(SupportsShouldProcess)]
param([string] $Root = (Split-Path $PSScriptRoot -Parent), [string] $Remote = 'origin', [string] $Branch = 'main',
    [string] $Repository = '', [switch] $PushOnly, [string] $WorktreeParent = [IO.Path]::GetTempPath())
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
