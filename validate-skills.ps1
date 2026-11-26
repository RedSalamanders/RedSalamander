<#
.SYNOPSIS
Validates scoped agent guidance and skill packages with the pinned CI engine.
.OUTPUTS
Guidance console summary. Validation failures are terminating errors.
.NOTES
Requires PowerShell 7 and the exact clean .ci-engine checkout. Read-only.
.EXAMPLE
./validate-skills.ps1
#>
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$binding = Join-Path $PSScriptRoot '.ci-engine/Tools/Modules/Ci.EngineBinding.psm1'
if (-not [IO.File]::Exists($binding)) {
    throw 'validate-skills.ps1 requires the clean .ci-engine checkout at engine.ref in .ci/project.json.'
}
Import-Module $binding -Scope Local
$guidanceModule = Resolve-CiPinnedEngineModule -ProjectRoot $PSScriptRoot -Name AgentGuidance.psm1
Import-Module $guidanceModule -Scope Local
$result = Get-CiAgentGuidanceResult -Root $PSScriptRoot -Kind Product

# Specs/README.md excludes the human-managed note from the WIP index.
# Its name is discovered by G12, which does not read the note.
$privateNoteDiagnostic = 'Specs/Plans/WIP/Notes_Scratch.md: is not linked from Specs/Plans/WIP/README.md, which names each plan exactly once (G12)'
$problems = @($result.Problems | Where-Object { $_ -cne $privateNoteDiagnostic })
if ($problems.Count) {
    throw "Agent guidance validation failed:$([Environment]::NewLine)$($problems -join [Environment]::NewLine)"
}
Write-Host "Agent guidance valid: $($result.Skills) skills, $($result.References) references."
