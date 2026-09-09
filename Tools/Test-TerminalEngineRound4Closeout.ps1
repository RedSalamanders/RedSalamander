<#
.SYNOPSIS
Reauthenticates the sealed Terminal Round-4 no-winner closeout.

.DESCRIPTION
Historical integrity command that verifies exact workflow, policy, universe, candidate, and runtime-dependency identities, re-evaluates the frozen governance closure, and writes fresh disposable evidence for comparison. It must not be used as a current product runtime test.

.PARAMETER RepoRoot
Specifies the repository whose sealed Round-4 inputs are authenticated.

.PARAMETER EvidenceRoot
Specifies a fresh nonexistent directory below .build for reproduced closeout evidence.

.OUTPUTS
Writes status text and historical evidence beneath EvidenceRoot; it does not return supported pipeline objects.

.NOTES
Prerequisites: unchanged sealed Round-4 inputs, PowerShell 7, and the historical evaluator dependencies. Side effects: creates the exact EvidenceRoot and launches historical evaluator child processes. The root must not exist before invocation. Exit is nonzero on any identity, closure, cohort, or reproduction mismatch. Primary consumers: the manual Round-4 closeout workflow and explicit HistoricalTerminal archival profile; default CI uses only the narrow historical integrity contract.

.EXAMPLE
.\Tools\Test-TerminalEngineRound4Closeout.ps1 -EvidenceRoot .\.build\terminal-round4-recheck\evidence
#>
[CmdletBinding()]
param(
    [string]$RepoRoot = [IO.Path]::GetFullPath((
            Join-Path $PSScriptRoot '..')),

    [string]$EvidenceRoot = [IO.Path]::GetFullPath((
            Join-Path $PSScriptRoot (
                '..\.build\terminal-engine-round4-closeout\evidence')))
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$expectedIdentities = [ordered]@{
    ExecutableWorkflow =
        'e6f15b6fef9c0d367e12dc7f43cb769106dc982934658f501120728599f32d1c'
    NeutralityPolicy =
        '7f9eee4b83576357a9407c39681c5d3f6a0d25a0d13bd30b3fb6885286863d77'
    Universe =
        '7621ab8562cb4b1d530d24355c67d90c905781a5d4f6bbc3364699ac059cb740'
    CandidateReview =
        'ba2942c7abe8f94b6f50be22ddb78570eaade5f60a152a8b0cca48dcea91542d'
    CandidateAssessment =
        'adf5ae89c3c90fd6018652f9ebd88aa809048a849ba29f49b04a4fde6d20343e'
    CandidateSet =
        'fb240d3ef838daea0740417fafd2d7d44b9536295f16c607e45fe57e64d11f4d'
    RuntimeDependencies =
        '89e93ff6c15c513897a7cb04ecb052ab9b311213cf5a5fb1988976d925efa149'
}

function Get-RSRound4FileSha256 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not [IO.File]::Exists($Path)) {
        throw [IO.FileNotFoundException]::new(
            "Required Round-4 closeout file is missing: $Path",
            $Path)
    }
    return (Get-FileHash -Algorithm SHA256 -LiteralPath $Path).
        Hash.ToLowerInvariant()
}

function Invoke-RSRound4ChildProcess {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    $start = [Diagnostics.ProcessStartInfo]::new()
    $start.FileName = (Get-Command pwsh -ErrorAction Stop).Source
    $start.UseShellExecute = $false
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    foreach ($argument in $Arguments) {
        [void]$start.ArgumentList.Add($argument)
    }
    $process = [Diagnostics.Process]::Start($start)
    $stdoutTask = $process.StandardOutput.ReadToEndAsync()
    $stderrTask = $process.StandardError.ReadToEndAsync()
    $process.WaitForExit()
    return [pscustomobject]@{
        ExitCode = $process.ExitCode
        StandardOutput = $stdoutTask.GetAwaiter().GetResult()
        StandardError = $stderrTask.GetAwaiter().GetResult()
    }
}

$RepoRoot = [IO.Path]::GetFullPath($RepoRoot)
$EvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
if ([IO.Directory]::Exists($EvidenceRoot) -or
    [IO.File]::Exists($EvidenceRoot)) {
    throw [IO.IOException]::new(
        'The Round-4 closeout evidence root must not exist before evaluation.')
}

$identityPaths = [ordered]@{
    ExecutableWorkflow = Join-Path $RepoRoot (
        '.github\workflows\terminal-engine-gate0.yml')
    NeutralityPolicy = Join-Path $RepoRoot (
        'Specs\Terminal\TerminalEngineCollectionNeutrality.v1.json')
    Universe = Join-Path $RepoRoot (
        'Specs\Terminal\TerminalEngineCandidateUniverse.v4.json')
    CandidateReview = Join-Path $RepoRoot (
        'Specs\Terminal\TerminalEngineCandidateReview.v1.json')
    CandidateAssessment = Join-Path $RepoRoot (
        'Specs\Terminal\TerminalEngineCandidateAssessment.v1.json')
    CandidateSet = Join-Path $RepoRoot (
        'Specs\Terminal\TerminalEngineCandidateSet.v1.json')
    RuntimeDependencies = Join-Path $RepoRoot 'RuntimeDependencies.props'
}
foreach ($name in $expectedIdentities.Keys) {
    $actual = Get-RSRound4FileSha256 -Path $identityPaths[$name]
    if ($actual -cne $expectedIdentities[$name]) {
        throw [IO.InvalidDataException]::new(
            "Round-4 closeout identity '$name' changed.")
    }
}

Import-Module (Join-Path $RepoRoot 'Tools\TerminalEngineEvaluator.psm1') -Force
$closure = Read-RSCandidateGovernanceClosure -RepoRoot $RepoRoot
if ([string]$closure.UniverseSha256 -cne $expectedIdentities.Universe -or
    [string]$closure.CandidateReviewSha256 -cne
        $expectedIdentities.CandidateReview -or
    [string]$closure.CandidateAssessmentSha256 -cne
        $expectedIdentities.CandidateAssessment -or
    [string]$closure.CandidateSetSha256 -cne
        $expectedIdentities.CandidateSet -or
    [string]$closure.UniverseOpenUtc -cne '2026-07-31T01:53:33Z' -or
    [string]$closure.GovernanceFreezeUtc -cne '2026-07-31T02:29:00Z') {
    throw [IO.InvalidDataException]::new(
        'The authenticated Round-4 governance closure is not exact.')
}

$candidateIds = [string[]]@(
    $closure.CandidateSet.candidates |
        ForEach-Object { [string]$_.candidateId } |
        Sort-Object -CaseSensitive)
$collectCandidateCount = @(
    $closure.CandidateSet.candidates |
        Where-Object {
            [string]$_.scoredOrControl -ceq 'scored' -and
            [string]$_.disposition -ceq 'collect'
        }).Count
if ($candidateIds.Count -ne 6 -or $collectCandidateCount -ne 0) {
    throw [IO.InvalidDataException]::new(
        'Round 4 is no longer the exact six-record, zero-collect cohort.')
}

$terminalEvidenceRootsBefore = [string[]]@(
    Get-ChildItem `
        -LiteralPath (Join-Path $RepoRoot 'Specs\TestRuns') `
        -Directory `
        -Recurse `
        -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -ceq 'TerminalEngine' } |
        ForEach-Object { $_.FullName } |
        Sort-Object -CaseSensitive)
if ($terminalEvidenceRootsBefore.Count -ne 0) {
    throw [IO.InvalidDataException]::new(
        'Round 4 must begin with no archived TerminalEngine evidence root.')
}

$forbiddenProductPaths = [ordered]@{
    ProductTerminalDirectory = Join-Path $RepoRoot 'Plugins\Terminal'
    ProductTerminalAbi = Join-Path $RepoRoot 'Common\PlugInterfaces\Terminal.h'
    WinnerLock = Join-Path $RepoRoot 'External\terminal-engine.lock.json'
    WinnerBuildRoot = Join-Path $RepoRoot '.build\TerminalEngine'
}
foreach ($name in $forbiddenProductPaths.Keys) {
    $path = $forbiddenProductPaths[$name]
    if ([IO.Directory]::Exists($path) -or [IO.File]::Exists($path)) {
        throw [IO.InvalidDataException]::new(
            "Round-4 closeout found forbidden winner/product state '$name'.")
    }
}

$evaluateArguments = [Collections.Generic.List[string]]::new()
foreach ($argument in @(
        '-NoLogo',
        '-NoProfile',
        '-File',
        (Join-Path $RepoRoot 'Tools\Evaluate-TerminalEngines.ps1'),
        '-Mode',
        'Collect',
        '-Candidate',
        'All',
        '-Platform',
        'x64',
        '-Configuration',
        'Release',
        '-Bootstrap',
        '-Clean',
        '-Offline',
        '-EvidenceRoot',
        $EvidenceRoot,
        '-PassThruManifest')) {
    $evaluateArguments.Add($argument)
}
$evaluation = Invoke-RSRound4ChildProcess `
    -Arguments $evaluateArguments.ToArray()
if ($evaluation.ExitCode -ne 21 -or
    -not [string]::IsNullOrEmpty($evaluation.StandardOutput) -or
    $evaluation.StandardError -cnotmatch 'static-only evidence') {
    throw [IO.InvalidDataException]::new(
        'Canonical Round-4 Collect did not produce the exact pre-effect static-only rejection.')
}

$providerOutput = @()
$providerRejectedStaticCohort = $false
try {
    $providerOutput = @(
        & (Join-Path $RepoRoot (
                'Tools\TerminalEngine\Invoke-TerminalEngineCollectionProvider.ps1')) `
            -Phase Bootstrap `
            -CandidateId $candidateIds `
            -Platform x64 `
            -Configuration Release `
            -RepoRoot $RepoRoot
    )
} catch [IO.InvalidDataException] {
    if ($_.Exception.Message -cmatch
        'static-only cohort cannot attest a lane') {
        $providerRejectedStaticCohort = $true
    }
    else {
        throw
    }
}
if (-not $providerRejectedStaticCohort -or $providerOutput.Count -ne 0) {
    throw [IO.InvalidDataException]::new(
        'The direct provider did not reject the complete static-only cohort before effects.')
}

if ([IO.Directory]::Exists($EvidenceRoot) -or
    [IO.File]::Exists($EvidenceRoot)) {
    throw [IO.InvalidDataException]::new(
        'Round-4 closeout validation mutated its evidence root.')
}
$terminalEvidenceRootsAfter = [string[]]@(
    Get-ChildItem `
        -LiteralPath (Join-Path $RepoRoot 'Specs\TestRuns') `
        -Directory `
        -Recurse `
        -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -ceq 'TerminalEngine' } |
        ForEach-Object { $_.FullName } |
        Sort-Object -CaseSensitive)
if ($terminalEvidenceRootsAfter.Count -ne 0) {
    throw [IO.InvalidDataException]::new(
        'Round 4 must end with no archived TerminalEngine evidence root.')
}
foreach ($name in $forbiddenProductPaths.Keys) {
    $path = $forbiddenProductPaths[$name]
    if ([IO.Directory]::Exists($path) -or [IO.File]::Exists($path)) {
        throw [IO.InvalidDataException]::new(
            "Round-4 closeout created forbidden winner/product state '$name'.")
    }
}

[pscustomobject]@{
    Status = 'valid-no-winner-product-blocked'
    UniverseSha256 = [string]$closure.UniverseSha256
    CandidateReviewSha256 = [string]$closure.CandidateReviewSha256
    CandidateAssessmentSha256 = [string]$closure.CandidateAssessmentSha256
    CandidateSetSha256 = [string]$closure.CandidateSetSha256
    CandidateRecordCount = $candidateIds.Count
    CollectCandidateCount = $collectCandidateCount
    CollectExitCode = $evaluation.ExitCode
    CollectStandardOutputLength = $evaluation.StandardOutput.Length
    ProviderRejectedStaticCohort = $providerRejectedStaticCohort
    EvidenceRootAbsentBefore = $true
    EvidenceRootAbsentAfter = $true
    ArchivedTerminalEngineRootsAbsentBefore = $true
    ArchivedTerminalEngineRootsAbsentAfter = $true
    WinnerLockAbsent = $true
    ProductTerminalDirectoryAbsent = $true
    ProductTerminalAbiAbsent = $true
    WinnerBuildRootAbsent = $true
    RuntimeDependencyManifestFrozen = $true
}
