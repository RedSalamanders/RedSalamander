<#
.SYNOPSIS
    Runs raw cloc reporting for the directories represented by a Visual Studio solution.

.DESCRIPTION
    Preserves the supported InvokeSolutionCloc.ps1 path used for hand-run solution
    reports. The command forwards to Measure-SourceLines.ps1 in Solution scope so
    solution parsing and minimal-directory discovery have one maintained
    implementation while this established command line remains available.

.PARAMETER SolutionPath
    Solution to measure. Defaults to RedSalamander.sln at the repository root.

.PARAMETER ClocPath
    cloc executable name or full path. Defaults to cloc.

.PARAMETER ClocArgs
    Additional arguments passed directly to cloc. Unbound trailing arguments are
    collected here, preserving calls such as InvokeSolutionCloc.ps1 --by-file --xml.

.OUTPUTS
    Raw text emitted by cloc. The wrapper defines no stable pipeline object schema.

.NOTES
    Prerequisites: Git, a readable Visual Studio solution, and cloc (installable
    with winget install AlDanial.Cloc). Side effects: launches cloc as a read-only
    child process and does not write repository files. Primary consumers:
    developers running ad hoc solution-scoped line reports by hand. Resolution,
    parsing, and cloc failures produce a nonzero exit; cloc's exit status is
    preserved by the canonical command.

.EXAMPLE
    .\Tools\InvokeSolutionCloc.ps1

    Prints the standard raw cloc summary for RedSalamander.sln directories.

.EXAMPLE
    .\Tools\InvokeSolutionCloc.ps1 --by-file --xml

    Passes raw report-format arguments through to cloc.
#>

[CmdletBinding(PositionalBinding = $false)]
param(
    [string]$SolutionPath = (Join-Path $PSScriptRoot '..\RedSalamander.sln'),

    [string]$ClocPath = 'cloc',

    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ClocArgs = @()
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$canonicalCommand = Join-Path $PSScriptRoot 'Measure-SourceLines.ps1'
& $canonicalCommand `
    -Scope Solution `
    -SolutionPath $SolutionPath `
    -ClocPath $ClocPath `
    -ClocArgs $ClocArgs

if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}
