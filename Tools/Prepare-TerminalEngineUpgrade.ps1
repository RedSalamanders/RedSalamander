<#
.SYNOPSIS
Collects or finalizes lock-bound evidence for a reviewed Terminal-engine upgrade.

.DESCRIPTION
Compatibility entrypoint for the generic Terminal upgrade protocol. Collect builds one approved candidate lane and emits evidence; Finalize compares the reviewed five-lane evidence set with the current lock and writes a rollback-bound upgrade comparison. It does not choose a candidate or replace the current product lock automatically.

.PARAMETER Mode
Selects Collect for one evidence lane or Finalize for the complete comparison.

.PARAMETER TargetPin
Specifies the exact proposed upstream source pin.

.PARAMETER TargetSourceSha256
Specifies the approved lowercase SHA-256 for the proposed source input.

.PARAMETER ExpectedCurrentLockSha256
Binds the operation to the exact current lock identity.

.PARAMETER EngineLock
Specifies the current Terminal-engine lock.

.PARAMETER Platform
Selects the Collect architecture. It is not used as a Finalize lane selector.

.PARAMETER Configuration
Selects the Collect configuration.

.PARAMETER Bootstrap
Permits the reviewed candidate bootstrap step during Collect.

.PARAMETER Clean
Rebuilds the candidate lane from a clean output directory.

.PARAMETER Offline
Requires verified network isolation for the candidate build.

.PARAMETER ExclusiveDisposableRunner
Attests disposable-runner ownership; required for offline collection.

.PARAMETER CandidateLockFile
Specifies the exact reviewed candidate lock.

.PARAMETER EvidenceManifest
Specifies the exact lane manifests consumed by Finalize.

.PARAMETER UpgradeRoot
Selects the safe repository build root for upgrade evidence and comparisons.

.PARAMETER RequireAllEvidence
Requires the complete approved lane set during Finalize.

.PARAMETER PassThruManifest
Returns the generated lane manifest or upgrade-comparison path.

.OUTPUTS
With -PassThruManifest, a string containing the generated artifact path. Otherwise no pipeline output. Evidence is written below UpgradeRoot.

.NOTES
Prerequisites: reviewed current/candidate locks, approved hashes, required toolchains, and the Terminal upgrade record described in Specs/Terminal/TerminalEngineDecision.md. Side effects: Collect builds external code and writes evidence; offline mode applies host-wide isolation on an exclusive disposable runner. Exit is nonzero for identity, lane, evidence, rollback, or policy mismatch. Primary consumer: maintainer-operated Terminal runtime upgrades, not ordinary product builds.

.EXAMPLE
.\Tools\Prepare-TerminalEngineUpgrade.ps1 -Mode Finalize -TargetPin <pin> -TargetSourceSha256 <sha256> -ExpectedCurrentLockSha256 <sha256> -CandidateLockFile .\.build\candidate-lock.json -EvidenceManifest .\.build\lane1.json,.\.build\lane2.json -UpgradeRoot .\.build\TerminalEngineUpgrade -RequireAllEvidence -PassThruManifest
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet('Collect', 'Finalize')]
    [string]$Mode,

    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$TargetPin,

    [Parameter(Mandatory)]
    [ValidatePattern('^[0-9a-f]{64}$')]
    [string]$TargetSourceSha256,

    [Parameter(Mandatory)]
    [ValidatePattern('^[0-9a-f]{64}$')]
    [string]$ExpectedCurrentLockSha256,

    [string]$EngineLock = 'External\terminal-engine.lock.json',

    [ValidateSet('x64', 'ARM64')]
    [string]$Platform = '',

    [ValidateSet('Debug', 'Release', 'ASan Debug')]
    [string]$Configuration = '',

    [switch]$Bootstrap,

    [switch]$Clean,

    [switch]$Offline,

    [switch]$ExclusiveDisposableRunner,

    [string]$CandidateLockFile = '',

    [string[]]$EvidenceManifest = @(),

    [Parameter(Mandatory)]
    [string]$UpgradeRoot,

    [switch]$RequireAllEvidence,

    [switch]$PassThruManifest
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$lifecycleModule = Join-Path $repoRoot 'Tools\TerminalEngine\TerminalEngineLifecycle.psm1'
Import-Module $lifecycleModule -Force -ErrorAction Stop

function Get-RSUtf8Sha256 {
    param([Parameter(Mandatory)][string]$Value)

    $algorithm = [Security.Cryptography.SHA256]::Create()
    try {
        return [Convert]::ToHexString(
            $algorithm.ComputeHash([Text.UTF8Encoding]::new($false).GetBytes($Value))).
            ToLowerInvariant()
    }
    finally {
        $algorithm.Dispose()
    }
}

function Get-RSUpgradeComparisonSnapshot {
    param(
        [Parameter(Mandatory)]
        [object]$Lock
    )

    $sourceInputId = [string](Get-RSObjectProperty `
            -Object $Lock.build `
            -Name sourceInputId `
            -Required)
    if (-not $Lock.inputById.ContainsKey($sourceInputId)) {
        throw "build.sourceInputId '$sourceInputId' does not identify a locked input."
    }
    $toolInputs = @($Lock.inputs | Where-Object {
            [string](Get-RSObjectProperty -Object $_ -Name category -Default '') -eq
            'build-tool'
        } | ForEach-Object {
            [ordered]@{ id = $_.id; sha256 = $_.sha256 }
        })
    $qualification = Get-RSObjectProperty -Object $Lock -Name qualification -Required
    foreach ($name in @(
            'security',
            'conformance',
            'fuzz',
            'performance',
            'memory',
            'packageSize')) {
        if ($null -eq $qualification.PSObject.Properties[$name]) {
            throw "qualification.$name is required for an upgrade comparison."
        }
    }
    return [ordered]@{
        source = Get-RSCanonicalObjectSha256 -InputObject ([ordered]@{
                pin = [string]$Lock.engine.pin
                sourceInputId = $sourceInputId
                sourceSha256 = [string]$Lock.inputById[$sourceInputId].sha256
                adaptation = $Lock.sourceAdaptation
                derivedInputs = @($Lock.derivedInputs)
            })
        submodules = Get-RSCanonicalObjectSha256 -InputObject @(
            $Lock.dependencies | ForEach-Object {
                [ordered]@{ id = $_.id; pin = $_.pin }
            })
        tools = Get-RSCanonicalObjectSha256 -InputObject ([ordered]@{
                cachedInputs = $toolInputs
                installed = @($Lock.installedToolchains)
            })
        dependencySbom = Get-RSCanonicalObjectSha256 -InputObject @($Lock.dependencies)
        runtimeClosure = Get-RSCanonicalObjectSha256 -InputObject @($Lock.runtimeClosure)
        licensesAndNotices = Get-RSCanonicalObjectSha256 -InputObject @(
            $Lock.dependencies | ForEach-Object {
                [ordered]@{ id = $_.id; license = $_.license }
            })
        abi = Get-RSCanonicalObjectSha256 -InputObject $Lock.abi
        capabilities = Get-RSCanonicalObjectSha256 -InputObject $Lock.abi.capabilities
        security = Get-RSCanonicalObjectSha256 -InputObject $qualification.security
        conformance = Get-RSCanonicalObjectSha256 -InputObject $qualification.conformance
        fuzz = Get-RSCanonicalObjectSha256 -InputObject $qualification.fuzz
        performance = Get-RSCanonicalObjectSha256 -InputObject $qualification.performance
        memory = Get-RSCanonicalObjectSha256 -InputObject $qualification.memory
        packageSize = Get-RSCanonicalObjectSha256 -InputObject $qualification.packageSize
    }
}

function Get-RSUpgradeManifestReference {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $resolved = [IO.Path]::GetFullPath($Path)
    if (-not (Test-RSPathIsWithin `
            -Path $resolved `
            -Root (Join-Path $repoRoot '.build')) -or
        -not (Test-Path -LiteralPath $resolved -PathType Leaf)) {
        throw 'Upgrade build and verification manifests must be exact files below .build.'
    }
    return [ordered]@{
        path = [IO.Path]::GetRelativePath($repoRoot, $resolved).Replace('\', '/')
        sha256 = Get-RSSha256 -Path $resolved
    }
}

function Resolve-RSUpgradeManifestReference {
    param(
        [Parameter(Mandatory)]
        [object]$Reference,

        [Parameter(Mandatory)]
        [string]$Description
    )

    $pathValue = [string](Get-RSObjectProperty `
            -Object $Reference `
            -Name path `
            -Required)
    if ([string]::IsNullOrWhiteSpace($pathValue) -or
        [IO.Path]::IsPathFullyQualified($pathValue)) {
        throw "$Description path must be one repository-relative file below .build."
    }
    $resolved = Resolve-RSTerminalPathInRoot `
        -Root $repoRoot `
        -Path $pathValue `
        -MustExist
    if (-not (Test-RSPathIsWithin `
            -Path $resolved `
            -Root (Join-Path $repoRoot '.build'))) {
        throw "$Description path must resolve below .build."
    }
    $canonicalRelative = [IO.Path]::GetRelativePath(
        $repoRoot,
        $resolved).Replace('\', '/')
    if ($canonicalRelative -cne $pathValue.Replace('\', '/')) {
        throw "$Description path is not canonical."
    }
    $expectedSha256 = [string](Get-RSObjectProperty `
            -Object $Reference `
            -Name sha256 `
            -Required)
    if ($expectedSha256 -cnotmatch '^[0-9a-f]{64}$') {
        throw "$Description.sha256 must be exactly 64 lowercase hexadecimal characters."
    }
    if ((Get-RSSha256 -Path $resolved) -cne $expectedSha256) {
        throw "$Description digest differs from the collected evidence."
    }
    $manifest = Get-Content -LiteralPath $resolved -Raw |
        ConvertFrom-Json -Depth 100 -ErrorAction Stop
    return [pscustomobject]@{
        path = $resolved
        relativePath = $canonicalRelative
        sha256 = $expectedSha256
        manifest = $manifest
    }
}

function Assert-RSUpgradeOfflineAttestation {
    param(
        [Parameter(Mandatory)]
        [object]$Manifest,

        [Parameter(Mandatory)]
        [string]$Description
    )

    $offline = Get-RSObjectProperty `
        -Object $Manifest `
        -Name offline `
        -Required
    if (-not [bool](Get-RSObjectProperty `
            -Object $offline `
            -Name requested `
            -Required) -or
        [string](Get-RSObjectProperty `
            -Object $offline `
            -Name providerStatus `
            -Required) -cne 'pass' -or
        -not [bool](Get-RSObjectProperty `
            -Object $offline `
            -Name isolationVerified `
            -Required)) {
        throw "$Description lacks runner-enforced offline attestation."
    }
}

function Test-RSUpgradeLaneManifestPair {
    param(
        [Parameter(Mandatory)]
        [object]$BuildReference,

        [Parameter(Mandatory)]
        [object]$VerifyReference,

        [Parameter(Mandatory)]
        [string]$Role,

        [Parameter(Mandatory)]
        [string]$Platform,

        [Parameter(Mandatory)]
        [string]$Configuration,

        [Parameter(Mandatory)]
        [string]$LockSha256,

        [Parameter(Mandatory)]
        [string]$Pin
    )

    $lane = "$Platform/$Configuration"
    $build = Resolve-RSUpgradeManifestReference `
        -Reference $BuildReference `
        -Description "$Role $lane build manifest"
    $verify = Resolve-RSUpgradeManifestReference `
        -Reference $VerifyReference `
        -Description "$Role $lane verify manifest"
    if ([string]::Equals(
            $build.path,
            $verify.path,
            [StringComparison]::OrdinalIgnoreCase)) {
        throw "$Role $lane build and verify manifests must be distinct files."
    }

    $buildManifest = $build.manifest
    if ([string](Get-RSObjectProperty `
            -Object $buildManifest `
            -Name formatId `
            -Required) -cne 'red-salamander-terminal-engine-build' -or
        [int](Get-RSObjectProperty `
            -Object $buildManifest `
            -Name schemaVersion `
            -Required) -ne 1 -or
        [string]$buildManifest.lock.sha256 -cne $LockSha256 -or
        [string]$buildManifest.lock.pin -cne $Pin -or
        [string]$buildManifest.platform -cne $Platform -or
        [string]$buildManifest.configuration -cne $Configuration -or
        -not [bool](Get-RSObjectProperty `
            -Object $buildManifest `
            -Name clean `
            -Required) -or
        @(Get-RSObjectProperty `
            -Object $buildManifest `
            -Name outputs `
            -Required).Count -eq 0) {
        throw "$Role $lane build manifest is not bound to the exact clean lane and lock."
    }
    Assert-RSUpgradeOfflineAttestation `
        -Manifest $buildManifest `
        -Description "$Role $lane build manifest"

    $verifyManifest = $verify.manifest
    if ([string](Get-RSObjectProperty `
            -Object $verifyManifest `
            -Name formatId `
            -Required) -cne
            'red-salamander-terminal-engine-verification' -or
        [int](Get-RSObjectProperty `
            -Object $verifyManifest `
            -Name schemaVersion `
            -Required) -ne 1 -or
        [string]$verifyManifest.lock.sha256 -cne $LockSha256 -or
        [string]$verifyManifest.lock.pin -cne $Pin -or
        [string]$verifyManifest.platform -cne $Platform -or
        [string]$verifyManifest.configuration -cne $Configuration -or
        [string]$verifyManifest.buildManifestSha256 -cne $build.sha256 -or
        -not [bool](Get-RSObjectProperty `
            -Object $verifyManifest `
            -Name runSmoke `
            -Required) -or
        [string](Get-RSObjectProperty `
            -Object $verifyManifest `
            -Name status `
            -Required) -cne 'pass' -or
        @(Get-RSObjectProperty `
            -Object $verifyManifest `
            -Name outputs `
            -Required).Count -eq 0) {
        throw "$Role $lane verify manifest is not bound to the exact build, lane, and lock."
    }
    Assert-RSUpgradeOfflineAttestation `
        -Manifest $verifyManifest `
        -Description "$Role $lane verify manifest"

    $expectedNative = if ($Platform -ceq 'x64') { 'X64' } else { 'Arm64' }
    if ([string]$verifyManifest.native.process -cne $expectedNative -or
        [string]$verifyManifest.native.os -cne $expectedNative) {
        throw "$Role $lane verify manifest is not native."
    }
    $smokeRecords = @(Get-RSObjectProperty `
            -Object $verifyManifest `
            -Name smoke `
            -Required)
    if ($smokeRecords.Count -eq 0) {
        throw "$Role $lane verify manifest contains no smoke proof."
    }
    foreach ($smoke in $smokeRecords) {
        if (-not [bool](Get-RSObjectProperty `
                -Object $smoke `
                -Name verified `
                -Required) -or
            [int](Get-RSObjectProperty `
                -Object $smoke `
                -Name exitCode `
                -Required) -ne 0) {
            throw "$Role $lane verify manifest contains a failing smoke result."
        }
    }
    return [ordered]@{
        platform = $Platform
        configuration = $Configuration
        buildManifest = [ordered]@{
            path = $build.relativePath
            sha256 = $build.sha256
        }
        verifyManifest = [ordered]@{
            path = $verify.relativePath
            sha256 = $verify.sha256
        }
    }
}

$currentLock = Read-RSTerminalEngineLock -LockFile $EngineLock -ValidateInputs
if ($currentLock.lockSha256 -cne $ExpectedCurrentLockSha256) {
    throw 'The current terminal-engine lock digest differs from ExpectedCurrentLockSha256.'
}
$resolvedUpgradeRoot = Assert-RSTerminalSafeBuildRoot `
    -RepoRoot $repoRoot `
    -BuildRoot $UpgradeRoot
$targetPinDigest = Get-RSUtf8Sha256 -Value $TargetPin
$candidateRoot = Join-Path (
    Join-Path $resolvedUpgradeRoot $ExpectedCurrentLockSha256) $targetPinDigest
[void](Assert-RSTerminalSafeBuildRoot -RepoRoot $repoRoot -BuildRoot $candidateRoot)

if ($Mode -eq 'Collect') {
    if ([string]::IsNullOrWhiteSpace($Platform) -or
        [string]::IsNullOrWhiteSpace($Configuration)) {
        throw 'Collect requires explicit Platform and Configuration.'
    }
    if (-not $Bootstrap -or -not $Clean -or -not $Offline) {
        throw 'Collect requires explicit -Bootstrap -Clean -Offline.'
    }
    if (-not $ExclusiveDisposableRunner) {
        throw 'Offline upgrade collection requires -ExclusiveDisposableRunner.'
    }
    if ($EvidenceManifest.Count -ne 0 -or $RequireAllEvidence) {
        throw 'Collect does not accept finalization evidence parameters.'
    }
    if ([string]::IsNullOrWhiteSpace($CandidateLockFile)) {
        throw 'Collect requires an exact pre-reviewed -CandidateLockFile under .build.'
    }

    $candidateLock = Read-RSTerminalEngineLock `
        -LockFile $CandidateLockFile `
        -ValidateInputs
    $dotBuild = Join-Path $repoRoot '.build'
    if (-not (Test-RSPathIsWithin -Path $candidateLock.lockPath -Root $dotBuild)) {
        throw 'CandidateLockFile must be isolated below the repository .build directory.'
    }
    if ([string]$candidateLock.engine.pin -cne $TargetPin) {
        throw 'Candidate lock engine.pin differs from TargetPin.'
    }
    $sourceInputId = [string](Get-RSObjectProperty `
            -Object $candidateLock.build `
            -Name sourceInputId `
            -Required)
    if (-not $candidateLock.inputById.ContainsKey($sourceInputId) -or
        [string]$candidateLock.inputById[$sourceInputId].sha256 -cne
        $TargetSourceSha256) {
        throw 'Candidate lock source input differs from TargetSourceSha256.'
    }
    $upgradeContract = Get-RSObjectProperty `
        -Object $candidateLock `
        -Name upgrade `
        -Required
    if ([string](Get-RSObjectProperty `
            -Object $upgradeContract `
            -Name basedOnLockSha256 `
            -Required) -cne $ExpectedCurrentLockSha256 -or
        [string](Get-RSObjectProperty `
            -Object $upgradeContract `
            -Name targetSourceSha256 `
            -Required) -cne $TargetSourceSha256) {
        throw 'Candidate lock upgrade binding differs from the requested current lock or source.'
    }
    if ([bool](Get-RSObjectProperty `
            -Object $upgradeContract `
            -Name persistedStateSchemaChanged `
            -Required)) {
        throw 'An engine-only candidate may not change persisted terminal state.'
    }
    [void](Get-RSUpgradeComparisonSnapshot -Lock $candidateLock)

    [void](New-Item -ItemType Directory -Path $candidateRoot -Force)
    $isolatedCandidateLock = Join-Path $candidateRoot 'candidate-terminal-engine.lock.json'
    if (-not [string]::Equals(
            $candidateLock.lockPath,
            $isolatedCandidateLock,
            [StringComparison]::OrdinalIgnoreCase)) {
        [IO.File]::Copy($candidateLock.lockPath, $isolatedCandidateLock, $true)
    }
    if ((Get-RSSha256 -Path $isolatedCandidateLock) -cne $candidateLock.lockSha256) {
        throw 'Isolated candidate lock copy changed identity.'
    }

    $native = Get-RSTerminalNativeArchitecture
    $expectedNative = if ($Platform -eq 'x64') { 'X64' } else { 'Arm64' }
    if ([string]$native.process -cne $expectedNative -or
        [string]$native.os -cne $expectedNative) {
        throw "Upgrade collection requires a native $Platform process and OS."
    }

    $candidateOutputRoot = Join-Path $candidateRoot 'candidate-build'
    $buildArguments = @{
        Platform = $Platform
        Configuration = $Configuration
        LockFile = $isolatedCandidateLock
        OutputRoot = $candidateOutputRoot
        Clean = $true
        Offline = $true
        ExclusiveDisposableRunner = $true
        PassThruManifest = $true
    }
    $candidateBuildManifest = & (Join-Path $repoRoot 'Tools\Build-TerminalEngine.ps1') `
        @buildArguments
    $verifyArguments = @{
        Platform = $Platform
        Configuration = $Configuration
        LockFile = $isolatedCandidateLock
        OutputRoot = $candidateOutputRoot
        RunSmoke = $true
        Offline = $true
        ExclusiveDisposableRunner = $true
        PassThruManifest = $true
    }
    $candidateVerifyManifest = & (Join-Path $repoRoot 'Tools\Verify-TerminalEngine.ps1') `
        @verifyArguments

    $rollbackOutputRoot = Join-Path $candidateRoot 'rollback-build'
    $rollbackBuildArguments = @{
        Platform = $Platform
        Configuration = $Configuration
        LockFile = $currentLock.lockPath
        OutputRoot = $rollbackOutputRoot
        Clean = $true
        Offline = $true
        ExclusiveDisposableRunner = $true
        PassThruManifest = $true
    }
    $rollbackBuildManifest = & (Join-Path $repoRoot 'Tools\Build-TerminalEngine.ps1') `
        @rollbackBuildArguments
    $rollbackVerifyArguments = @{
        Platform = $Platform
        Configuration = $Configuration
        LockFile = $currentLock.lockPath
        OutputRoot = $rollbackOutputRoot
        RunSmoke = $true
        Offline = $true
        ExclusiveDisposableRunner = $true
        PassThruManifest = $true
    }
    $rollbackVerifyManifest = & (Join-Path $repoRoot 'Tools\Verify-TerminalEngine.ps1') `
        @rollbackVerifyArguments

    foreach ($path in @(
            $candidateBuildManifest,
            $candidateVerifyManifest,
            $rollbackBuildManifest,
            $rollbackVerifyManifest)) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw 'An upgrade build or verification step did not return an exact manifest.'
        }
    }

    $evidenceDirectory = Join-Path (
        Join-Path (
            Join-Path $candidateRoot 'evidence') $Platform) $Configuration
    [void](New-Item -ItemType Directory -Path $evidenceDirectory -Force)
    $manifestPath = Join-Path $evidenceDirectory 'upgrade-collection.json'
    $manifest = [ordered]@{
        formatId = 'red-salamander-terminal-engine-upgrade-collection'
        schemaVersion = 2
        createdUtc = [DateTime]::UtcNow.ToString(
            'o',
            [Globalization.CultureInfo]::InvariantCulture)
        currentLockSha256 = $ExpectedCurrentLockSha256
        candidateLockPath = [IO.Path]::GetRelativePath(
            $repoRoot,
            $isolatedCandidateLock).Replace('\', '/')
        candidateLockSha256 = $candidateLock.lockSha256
        targetPin = $TargetPin
        targetPinDigest = $targetPinDigest
        targetSourceSha256 = $TargetSourceSha256
        platform = $Platform
        configuration = $Configuration
        native = $native
        policy = [ordered]@{
            bootstrap = $true
            clean = $true
            offline = $true
            networkIsolation = 'runner-enforced'
        }
        candidate = [ordered]@{
            buildManifest = Get-RSUpgradeManifestReference `
                -Path $candidateBuildManifest
            verifyManifest = Get-RSUpgradeManifestReference `
                -Path $candidateVerifyManifest
            snapshot = Get-RSUpgradeComparisonSnapshot -Lock $candidateLock
        }
        rollback = [ordered]@{
            previousLockSha256 = $currentLock.lockSha256
            buildManifest = Get-RSUpgradeManifestReference `
                -Path $rollbackBuildManifest
            verifyManifest = Get-RSUpgradeManifestReference `
                -Path $rollbackVerifyManifest
            persistedStateSchemaChanged = $false
            status = 'pass'
        }
        status = 'pass'
    }
    Write-RSTerminalJsonFile -Path $manifestPath -Value $manifest
    if ($PassThruManifest) {
        $manifestPath
    }
    else {
        Write-Host "Upgrade collection passed: $Platform/$Configuration"
        Write-Host "Manifest: $manifestPath"
    }
    return
}

if (-not [string]::IsNullOrWhiteSpace($Platform) -or
    -not [string]::IsNullOrWhiteSpace($Configuration) -or
    $Bootstrap -or $Clean -or $Offline -or $ExclusiveDisposableRunner -or
    -not [string]::IsNullOrWhiteSpace($CandidateLockFile)) {
    throw 'Finalize does not accept collection-only parameters.'
}
if (-not $RequireAllEvidence) {
    throw 'Finalize requires -RequireAllEvidence.'
}
if ($EvidenceManifest.Count -ne 5) {
    throw 'Finalize requires exactly five explicitly named evidence manifests.'
}

$records = [Collections.Generic.List[object]]::new()
$manifestIdentities = [Collections.Generic.List[object]]::new()
$candidateLaneProofs = [Collections.Generic.List[object]]::new()
$rollbackLaneProofs = [Collections.Generic.List[object]]::new()
foreach ($pathValue in $EvidenceManifest) {
    $path = if ([IO.Path]::IsPathFullyQualified($pathValue)) {
        [IO.Path]::GetFullPath($pathValue)
    }
    else {
        [IO.Path]::GetFullPath((Join-Path $repoRoot $pathValue))
    }
    if (-not (Test-RSPathIsWithin -Path $path -Root (Join-Path $repoRoot '.build')) -or
        -not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw 'Every upgrade evidence manifest must be an exact file below .build.'
    }
    $record = Get-Content -LiteralPath $path -Raw |
        ConvertFrom-Json -Depth 100 -ErrorAction Stop
    if ([string]$record.formatId -cne
        'red-salamander-terminal-engine-upgrade-collection' -or
        [int]$record.schemaVersion -ne 2 -or
        [string]$record.status -cne 'pass') {
        throw 'Upgrade evidence record is invalid or did not pass.'
    }
    foreach ($binding in @(
            @{ Actual = [string]$record.currentLockSha256; Expected = $ExpectedCurrentLockSha256 },
            @{ Actual = [string]$record.targetPin; Expected = $TargetPin },
            @{ Actual = [string]$record.targetPinDigest; Expected = $targetPinDigest },
            @{ Actual = [string]$record.targetSourceSha256; Expected = $TargetSourceSha256 })) {
        if ($binding.Actual -cne $binding.Expected) {
            throw 'Upgrade evidence binding differs from the finalization request.'
        }
    }
    if (-not $record.policy.bootstrap -or
        -not $record.policy.clean -or
        -not $record.policy.offline -or
        [string]$record.policy.networkIsolation -cne 'runner-enforced' -or
        [string]$record.rollback.previousLockSha256 -cne
            $ExpectedCurrentLockSha256 -or
        [bool]$record.rollback.persistedStateSchemaChanged -or
        [string]$record.rollback.status -cne 'pass') {
        throw 'Upgrade evidence lacks mandatory clean/offline/rollback proof.'
    }
    $recordPlatform = [string]$record.platform
    $recordConfiguration = [string]$record.configuration
    if ($recordPlatform -cnotin @('x64', 'ARM64') -or
        $recordConfiguration -cnotin @(
            'Debug',
            'Release',
            'ASan Debug')) {
        throw 'Upgrade evidence names an unsupported lane.'
    }
    $expectedNative = if ($recordPlatform -ceq 'x64') { 'X64' } else {
        'Arm64'
    }
    if ([string]$record.native.process -cne $expectedNative -or
        [string]$record.native.os -cne $expectedNative) {
        throw "Upgrade evidence is not native for '$($record.platform)'."
    }
    $candidateLaneProofs.Add((Test-RSUpgradeLaneManifestPair `
            -BuildReference $record.candidate.buildManifest `
            -VerifyReference $record.candidate.verifyManifest `
            -Role 'candidate' `
            -Platform $recordPlatform `
            -Configuration $recordConfiguration `
            -LockSha256 ([string]$record.candidateLockSha256) `
            -Pin $TargetPin))
    $rollbackLaneProofs.Add((Test-RSUpgradeLaneManifestPair `
            -BuildReference $record.rollback.buildManifest `
            -VerifyReference $record.rollback.verifyManifest `
            -Role 'rollback' `
            -Platform $recordPlatform `
            -Configuration $recordConfiguration `
            -LockSha256 $ExpectedCurrentLockSha256 `
            -Pin ([string]$currentLock.engine.pin)))
    $records.Add($record)
    $manifestIdentities.Add([ordered]@{
            path = [IO.Path]::GetRelativePath($repoRoot, $path).Replace('\', '/')
            sha256 = Get-RSSha256 -Path $path
        })
}

$tupleKeys = @($records | ForEach-Object {
        "$($_.platform)/$($_.configuration)"
    } | Sort-Object -CaseSensitive)
$expectedTuples = @(
    'ARM64/Debug',
    'ARM64/Release',
    'x64/ASan Debug',
    'x64/Debug',
    'x64/Release')
if (($tupleKeys -join "`n") -cne ($expectedTuples -join "`n")) {
    throw 'Finalize requires all five native shipping lanes: x64 Debug/Release/ASan Debug and ARM64 Debug/Release.'
}
$candidateLockHashes = @($records | ForEach-Object {
        [string]$_.candidateLockSha256
    } | Sort-Object -Unique -CaseSensitive)
if ($candidateLockHashes.Count -ne 1) {
    throw 'Upgrade evidence records do not bind one candidate lock.'
}
$candidateLockPaths = @($records | ForEach-Object {
        [string]$_.candidateLockPath
    } | Sort-Object -Unique -CaseSensitive)
if ($candidateLockPaths.Count -ne 1) {
    throw 'Upgrade evidence records do not bind one candidate lock path.'
}
$snapshotHashes = @($records | ForEach-Object {
        Get-RSCanonicalObjectSha256 -InputObject $_.candidate.snapshot
    } | Sort-Object -Unique -CaseSensitive)
if ($snapshotHashes.Count -ne 1) {
    throw 'Upgrade evidence records do not bind one comparison snapshot.'
}

$candidateLockRelative = $candidateLockPaths[0]
$candidateLockSource = Resolve-RSTerminalPathInRoot `
    -Root $repoRoot `
    -Path $candidateLockRelative `
    -MustExist
if ((Get-RSSha256 -Path $candidateLockSource) -cne $candidateLockHashes[0]) {
    throw 'Candidate lock file differs from the collection evidence.'
}
$candidateLock = Read-RSTerminalEngineLock `
    -LockFile $candidateLockSource `
    -ValidateInputs
if ([string]$candidateLock.engine.pin -cne $TargetPin) {
    throw 'Finalized candidate lock engine.pin differs from TargetPin.'
}
$candidateSourceInputId = [string](Get-RSObjectProperty `
        -Object $candidateLock.build `
        -Name sourceInputId `
        -Required)
if (-not $candidateLock.inputById.ContainsKey($candidateSourceInputId) -or
    [string]$candidateLock.inputById[$candidateSourceInputId].sha256 -cne
        $TargetSourceSha256) {
    throw 'Finalized candidate lock source input differs from TargetSourceSha256.'
}
$candidateUpgrade = Get-RSObjectProperty `
    -Object $candidateLock `
    -Name upgrade `
    -Required
if ([string](Get-RSObjectProperty `
        -Object $candidateUpgrade `
        -Name basedOnLockSha256 `
        -Required) -cne $ExpectedCurrentLockSha256 -or
    [string](Get-RSObjectProperty `
        -Object $candidateUpgrade `
        -Name targetSourceSha256 `
        -Required) -cne $TargetSourceSha256 -or
    [bool](Get-RSObjectProperty `
        -Object $candidateUpgrade `
        -Name persistedStateSchemaChanged `
        -Required)) {
    throw 'Finalized candidate lock upgrade binding is invalid.'
}
$currentSnapshot = Get-RSUpgradeComparisonSnapshot -Lock $currentLock
$candidateSnapshot = Get-RSUpgradeComparisonSnapshot -Lock $candidateLock
$recordSnapshotSha256 = Get-RSCanonicalObjectSha256 `
    -InputObject $records[0].candidate.snapshot
$candidateSnapshotSha256 = Get-RSCanonicalObjectSha256 -InputObject $candidateSnapshot
if ($recordSnapshotSha256 -cne $candidateSnapshotSha256) {
    throw "Upgrade evidence comparison snapshot differs from the candidate lock (record=$recordSnapshotSha256 candidate=$candidateSnapshotSha256)."
}

[void](New-Item -ItemType Directory -Path $candidateRoot -Force)
$candidateLockOutput = Join-Path $candidateRoot 'candidate-terminal-engine.lock.json'
if (-not [string]::Equals(
        $candidateLockSource,
        $candidateLockOutput,
        [StringComparison]::OrdinalIgnoreCase)) {
    [IO.File]::Copy($candidateLockSource, $candidateLockOutput, $true)
}
if ((Get-RSSha256 -Path $candidateLockOutput) -cne $candidateLockHashes[0]) {
    throw 'Finalized candidate lock copy changed identity.'
}

$derivedChanges = @(
    foreach ($property in $candidateSnapshot.Keys) {
        if ([string]$currentSnapshot[$property] -cne [string]$candidateSnapshot[$property]) {
            $property
        }
    })
$declaredChanges = @(Get-RSObjectProperty `
        -Object $candidateLock.upgrade `
        -Name requiredChanges `
        -Required)
$declaredChangeSet = [Collections.Generic.HashSet[string]]::new(
    [StringComparer]::Ordinal)
foreach ($changeValue in $declaredChanges) {
    $change = [string]$changeValue
    if ([string]::IsNullOrWhiteSpace($change) -or
        -not $candidateSnapshot.Contains($change) -or
        -not $declaredChangeSet.Add($change)) {
        throw "Candidate upgrade.requiredChanges contains an empty, duplicate, or unknown comparison field '$change'."
    }
}
$requiredChanges = @($derivedChanges | Sort-Object -CaseSensitive)
$declaredChanges = @($declaredChanges | Sort-Object -CaseSensitive)
if (($requiredChanges -join "`n") -cne ($declaredChanges -join "`n")) {
    throw (
        "Candidate upgrade.requiredChanges must exactly equal the derived changed fields (declared={0}; derived={1})." -f
        ($declaredChanges -join ','),
        ($requiredChanges -join ','))
}
$comparisonPath = Join-Path $candidateRoot 'upgrade-comparison.json'
$comparison = [ordered]@{
    formatId = 'red-salamander-terminal-engine-upgrade-comparison'
    schemaVersion = 1
    createdUtc = [DateTime]::UtcNow.ToString(
        'o',
        [Globalization.CultureInfo]::InvariantCulture)
    inputs = @($manifestIdentities | Sort-Object -Stable -CaseSensitive -Property path)
    currentLockSha256 = $currentLock.lockSha256
    candidateLockPath = [IO.Path]::GetRelativePath(
        $repoRoot,
        $candidateLockOutput).Replace('\', '/')
    candidateLockSha256 = $candidateLockHashes[0]
    targetPin = $TargetPin
    targetPinDigest = $targetPinDigest
    targetSourceSha256 = $TargetSourceSha256
    current = $currentSnapshot
    candidate = $candidateSnapshot
    changedComparisonFields = $requiredChanges
    requiredRepositoryChangeCategories = $declaredChanges
    laneEvidence = @($candidateLaneProofs |
        Sort-Object -Stable -CaseSensitive -Property platform, configuration)
    rollback = [ordered]@{
        previousLockSha256 = $currentLock.lockSha256
        lanes = @($rollbackLaneProofs |
            Sort-Object -Stable -CaseSensitive -Property platform, configuration)
        everyLaneReconstructed = $rollbackLaneProofs.Count -eq 5
        persistedStateSchemaChanged = $false
        status = 'pass'
    }
    reviewerApprovalRequired = 2
    status = 'pass'
}
Write-RSTerminalJsonFile -Path $comparisonPath -Value $comparison

if ($PassThruManifest) {
    $comparisonPath
}
else {
    Write-Host 'Terminal engine upgrade comparison passed.'
    Write-Host "Manifest: $comparisonPath"
}
