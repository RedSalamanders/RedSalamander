<#
.SYNOPSIS
Reproduces the sealed Gate-0/Round-4 Terminal candidate evaluator.

.DESCRIPTION
Historical compatibility entrypoint that either collects the frozen candidate cohort or finalizes its evidence using TerminalEngineEvaluator.psm1. It is retained for archival reproduction and is not the current Ghostty product runtime workflow.

.PARAMETER Mode
Selects the frozen evaluator operation accepted by Assert-RSTerminalEngineInvocation.

.PARAMETER Candidate
Specifies the historical candidate selection accepted by Collect mode.

.PARAMETER Platform
Specifies the historical collection platform.

.PARAMETER Configuration
Specifies the historical collection configuration.

.PARAMETER Bootstrap
Permits the frozen bootstrap phase when the evaluator contract requires it.

.PARAMETER Clean
Requests a clean historical candidate collection.

.PARAMETER Offline
Requests the frozen offline collection policy.

.PARAMETER EvidenceRoot
Selects the historical evidence output root.

.PARAMETER PassThruManifest
Returns the evaluator manifest path.

.PARAMETER EvidenceManifest
Specifies the exact historical evidence manifests consumed by finalization.

.PARAMETER RequireAllEvidence
Requires the complete frozen evidence cohort.

.OUTPUTS
With -PassThruManifest, a string containing the manifest path. Otherwise no pipeline output. Top-level exit codes are 0 for success, 20 for complete no-winner, 21 for incomplete evidence, 64 for parameter errors, and 70 for internal errors.

.NOTES
Prerequisites: the sealed Terminal candidate corpus and its historical toolchain/runner requirements. Side effects: Collect may bootstrap/build external candidates and write evidence beneath the selected root. Primary consumers: the frozen Gate-0 workflow and explicit HistoricalTerminal archival profile. Do not add this command to current product CI.

.EXAMPLE
.\Tools\Evaluate-TerminalEngines.ps1 -Mode Finalize -EvidenceManifest .\.build\TerminalEvidence\*.json -RequireAllEvidence -PassThruManifest
#>
[CmdletBinding()]
param(
    [string]$Mode,

    [string[]]$Candidate,

    [string]$Platform,

    [string]$Configuration,

    [switch]$Bootstrap,

    [switch]$Clean,

    [switch]$Offline,

    [string]$EvidenceRoot,

    [switch]$PassThruManifest,

    [string[]]$EvidenceManifest,

    [switch]$RequireAllEvidence
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$exitSuccess = 0
$exitCompleteNoWinner = 20
$exitIncompleteEvidence = 21
$exitParameterError = 64
$exitInternalError = 70
$script:EvaluatorExitCode = $exitSuccess
$script:IsTopLevelFileInvocation = $false
$processArguments = [Environment]::GetCommandLineArgs()
for ($argumentIndex = 0; $argumentIndex + 1 -lt $processArguments.Count; ++$argumentIndex) {
    if ($processArguments[$argumentIndex] -ieq '-File' -or
        $processArguments[$argumentIndex] -ieq '-f') {
        try {
            $invokedFile = [IO.Path]::GetFullPath($processArguments[$argumentIndex + 1])
            $script:IsTopLevelFileInvocation = [string]::Equals(
                $invokedFile,
                [IO.Path]::GetFullPath($PSCommandPath),
                [StringComparison]::OrdinalIgnoreCase)
        } catch [ArgumentException] {
            $script:IsTopLevelFileInvocation = $false
        }
        break
    }
}

function Complete-RSEvaluatorInvocation {
    param(
        [Parameter(Mandatory = $true)]
        [int]$ExitCode,

        [string]$Message = ''
    )

    $script:EvaluatorExitCode = $ExitCode
    $global:LASTEXITCODE = $ExitCode
    if (-not [string]::IsNullOrEmpty($Message)) {
        [Console]::Error.WriteLine($Message)
    }
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$modulePath = Join-Path $PSScriptRoot 'TerminalEngineEvaluator.psm1'

try {
    Import-Module $modulePath -Force
    $null = Assert-RSTerminalEngineInvocation -BoundParameters $PSBoundParameters

    if ($Mode -ceq 'Collect') {
        $result = Invoke-RSTerminalEngineCollection `
            -RepoRoot $repoRoot `
            -Candidate $Candidate `
            -Platform $Platform `
            -Configuration $Configuration `
            -EvidenceRoot $EvidenceRoot
    } else {
        $result = Invoke-RSTerminalEngineFinalization `
            -RepoRoot $repoRoot `
            -EvidenceManifest $EvidenceManifest `
            -EvidenceRoot $EvidenceRoot `
            -RequireAllEvidence:$RequireAllEvidence
    }

    $resultExitCode = [int]$result.ExitCode
    if ($resultExitCode -ne $exitSuccess -and $resultExitCode -ne $exitCompleteNoWinner) {
        throw [InvalidOperationException]::new("Evaluator returned unsupported exit code $resultExitCode.")
    }
    Complete-RSEvaluatorInvocation -ExitCode $resultExitCode
    if ($PassThruManifest) {
        [string]$result.ManifestPath
    }
} catch [ArgumentException] {
    Complete-RSEvaluatorInvocation `
        -ExitCode $exitParameterError `
        -Message ("TerminalEngine evaluator parameter error: {0}" -f $_.Exception.Message)
} catch [IO.InvalidDataException] {
    Complete-RSEvaluatorInvocation `
        -ExitCode $exitIncompleteEvidence `
        -Message ("TerminalEngine evidence is incomplete or invalid: {0}" -f $_.Exception.Message)
} catch [IO.FileNotFoundException] {
    Complete-RSEvaluatorInvocation `
        -ExitCode $exitIncompleteEvidence `
        -Message ("TerminalEngine evidence is incomplete or invalid: {0}" -f $_.Exception.Message)
} catch {
    Complete-RSEvaluatorInvocation `
        -ExitCode $exitInternalError `
        -Message ("TerminalEngine evaluator internal error: {0}" -f $_.Exception.Message)
}

if ($script:IsTopLevelFileInvocation) {
    exit $script:EvaluatorExitCode
}
