<#
.SYNOPSIS
Observes the canonical Ghostty runtime lock without mutating tracked state.

.DESCRIPTION
Validates the active Ghostty lock and emits a schema-validated canonical status report below .build. Locked mode is network-free. Online mode observes the official default branch and OSV commit advisories. Candidate mode resolves only the explicit full commit and refuses substitution with branch head.

.PARAMETER Mode
Selects network-free Locked validation, Online default-branch/advisory observation, or exact Candidate observation.

.PARAMETER CandidateCommit
Requires one lowercase full 40-character Ghostty commit in Candidate mode. It is rejected in other modes.

.PARAMETER OutputPath
Selects the report path. Relative paths are repository-relative and every path must remain below .build. The default is .build/TerminalEngineStatus/<timestamp>/status.json.

.PARAMETER NetworkFixturePath
Injects a deterministic bounded network-response fixture for focused tests. Locked mode rejects it. Production workflows omit it.

.PARAMETER AdvisoryOnly
In Online mode, queries OSV for the locked commit without observing the default branch. The daily scheduled workflow uses this narrower read-only mode.

.PARAMETER GeneratedUtc
Overrides the report timestamp for deterministic tests. Production callers omit it.

.PARAMETER ExpectedSourceCommit
Requires the checked-out repository HEAD to equal this full commit. Release preflight uses it to prevent report replay across source revisions.

.PARAMETER PassThru
Returns the in-memory report with its reportPath property after canonical publication.

.OUTPUTS
With -PassThru, a PSCustomObject status report. Otherwise no pipeline output; the canonical report path and result are written to the host.

.NOTES
Prerequisites: PowerShell 7.4 or newer, Git, the canonical Ghostty lock/schema, and network access for Online/Candidate unless a test fixture is supplied. Side effects are limited to creating/replacing the selected report below .build; no lock, source, branch, issue, pull request, or tracked file is changed. Primary consumers are maintainers, the read-only status workflow, and release preflight.

.EXAMPLE
.\Tools\Get-GhosttyTerminalRuntimeStatus.ps1 -Mode Locked

.EXAMPLE
.\Tools\Get-GhosttyTerminalRuntimeStatus.ps1 -Mode Candidate -CandidateCommit 9c3ec931d64561a8407dde7ac984ce156ae91539 -PassThru
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Locked', 'Online', 'Candidate')]
    [string]$Mode,

    [string]$CandidateCommit,

    [string]$OutputPath,

    [string]$NetworkFixturePath,

    [switch]$AdvisoryOnly,

    [DateTime]$GeneratedUtc = [DateTime]::UtcNow,

    [string]$ExpectedSourceCommit,

    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$modulePath = Join-Path $repoRoot 'Tools\TerminalEngine\GhosttyRuntimeStatus.psm1'
Import-Module $modulePath -Force -ErrorAction Stop

$report = New-RSGhosttyRuntimeStatusReport `
    -RepoRoot $repoRoot `
    -Mode $Mode `
    -CandidateCommit $CandidateCommit `
    -NetworkFixturePath $NetworkFixturePath `
    -AdvisoryOnly:$AdvisoryOnly `
    -GeneratedUtc $GeneratedUtc `
    -ExpectedSourceCommit $ExpectedSourceCommit
$reportPath = Write-RSGhosttyRuntimeStatusReport -RepoRoot $repoRoot -Report $report -OutputPath $OutputPath

if ($PassThru) {
    [pscustomobject]$result = $report
    $result | Add-Member -NotePropertyName reportPath -NotePropertyValue $reportPath
    $result
}
else {
    Write-Host "Ghostty runtime status: mode=$Mode result=$($report.result) report=$reportPath"
}
