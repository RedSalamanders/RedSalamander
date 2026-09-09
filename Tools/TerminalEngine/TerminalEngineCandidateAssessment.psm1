Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-RSAssessmentProperty {
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [object]$Object,

        [Parameter(Mandatory)]
        [string]$Name
    )

    if ($null -eq $Object) {
        throw [IO.InvalidDataException]::new("Assessment object is null while reading '$Name'.")
    }
    if ($Object -is [Collections.IDictionary]) {
        foreach ($entry in $Object.GetEnumerator()) {
            if ([string]::Equals(
                    [string]$entry.Key,
                    $Name,
                    [StringComparison]::Ordinal)) {
                return $entry.Value
            }
        }
    }
    else {
        foreach ($property in $Object.PSObject.Properties) {
            if ([string]::Equals(
                    $property.Name,
                    $Name,
                    [StringComparison]::Ordinal)) {
                return $property.Value
            }
        }
    }
    throw [IO.InvalidDataException]::new("Assessment property '$Name' is missing.")
}

function Get-RSAssessmentInteger {
    param(
        [Parameter(Mandatory)]
        [object]$Object,

        [Parameter(Mandatory)]
        [string]$Name
    )

    $value = Get-RSAssessmentProperty -Object $Object -Name $Name
    if ($value -isnot [byte] -and
        $value -isnot [sbyte] -and
        $value -isnot [int16] -and
        $value -isnot [uint16] -and
        $value -isnot [int32] -and
        $value -isnot [uint32] -and
        $value -isnot [int64] -and
        $value -isnot [uint64]) {
        throw [IO.InvalidDataException]::new("Assessment property '$Name' must be an integer.")
    }
    return [int64]$value
}

function Get-RSAssessmentBoolean {
    param(
        [Parameter(Mandatory)]
        [object]$Object,

        [Parameter(Mandatory)]
        [string]$Name
    )

    $value = Get-RSAssessmentProperty -Object $Object -Name $Name
    if ($value -isnot [bool]) {
        throw [IO.InvalidDataException]::new("Assessment property '$Name' must be Boolean.")
    }
    return [bool]$value
}

function Assert-RSAssessmentDescriptorBuildTruth {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Candidate
    )

    $ledger = Get-RSAssessmentProperty -Object $Candidate -Name rawLedger
    $patch = Get-RSAssessmentProperty -Object $ledger -Name patch
    $build = Get-RSAssessmentProperty -Object $ledger -Name build
    $buildKind = [string](Get-RSAssessmentProperty `
            -Object $patch `
            -Name buildKind)
    $allowedBuildKinds = @(
        'provider-generic-cmake-msvc-v1',
        'provider-generic-zig-msvc-v1',
        'provider-generic-cargo-msvc-v1',
        'candidate-wrapper-v1')
    if ($allowedBuildKinds -cnotcontains $buildKind) {
        throw [IO.InvalidDataException]::new(
            "Candidate assessment patch proof has unsupported build kind '$buildKind'.")
    }
    $derivedMaintainedBuildWrapper =
        $buildKind -ceq 'candidate-wrapper-v1'
    $proofMaintainedBuildWrapper = Get-RSAssessmentBoolean `
        -Object $patch `
        -Name maintainedBuildWrapper
    $ledgerMaintainedBuildWrapper = Get-RSAssessmentBoolean `
        -Object $build `
        -Name maintainedBuildWrapper
    if ($proofMaintainedBuildWrapper -ne
            $derivedMaintainedBuildWrapper -or
        $ledgerMaintainedBuildWrapper -ne
            $derivedMaintainedBuildWrapper) {
        throw [IO.InvalidDataException]::new(
            "Candidate assessment maintainedBuildWrapper must equal the value mechanically derived from descriptor build kind '$buildKind'.")
    }
    return [pscustomobject]@{
        buildKind = $buildKind
        maintainedBuildWrapper = $derivedMaintainedBuildWrapper
    }
}

function New-RSAssessmentRating {
    param(
        [Parameter(Mandatory)]
        [string]$CriterionId,

        [Parameter(Mandatory)]
        [ValidateRange(0, 5)]
        [int]$Rating,

        [Parameter(Mandatory)]
        [string]$EvidenceSha256
    )

    return [pscustomobject]@{
        criterionId = $CriterionId
        rating = $Rating
        anchorId = "$($CriterionId.ToLowerInvariant())-r$Rating"
        evidenceSha256 = $EvidenceSha256
    }
}

function Get-RSCandidateLedgerSha256 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Candidate,

        [string]$TerminalEvidenceModule = (
            Join-Path $PSScriptRoot '..\TerminalEvidence.psm1')
    )

    Import-Module $TerminalEvidenceModule -Force -ErrorAction Stop
    $identity = [ordered]@{
        candidateId = Get-RSAssessmentProperty -Object $Candidate -Name candidateId
        seedCandidateId = Get-RSAssessmentProperty -Object $Candidate -Name seedCandidateId
        pin = Get-RSAssessmentProperty -Object $Candidate -Name pin
        pinSha256 = Get-RSAssessmentProperty -Object $Candidate -Name pinSha256
        sourceArchiveSha256 = Get-RSAssessmentProperty `
            -Object $Candidate `
            -Name sourceArchiveSha256
        rawLedger = Get-RSAssessmentProperty -Object $Candidate -Name rawLedger
    }
    return Get-RSJcsSha256 -InputObject $identity
}

function Get-RSCandidateAdmission {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Candidate,

        [Parameter(Mandatory)]
        [object]$CandidateUniverse
    )

    $ledger = Get-RSAssessmentProperty -Object $Candidate -Name rawLedger
    $patch = Get-RSAssessmentProperty -Object $ledger -Name patch
    $maintenance = Get-RSAssessmentProperty -Object $ledger -Name maintenance
    $build = Get-RSAssessmentProperty -Object $ledger -Name build
    [void](Assert-RSAssessmentDescriptorBuildTruth -Candidate $Candidate)
    $integration = Get-RSAssessmentProperty -Object $ledger -Name integration
    $policy = Get-RSAssessmentProperty `
        -Object $CandidateUniverse `
        -Name adaptationPolicy
    $caps = Get-RSAssessmentProperty -Object $policy -Name hardCaps

    $lines = Get-RSAssessmentInteger -Object $patch -Name changedMaintainedLines
    $files = Get-RSAssessmentInteger -Object $patch -Name changedMaintainedFiles
    $binaryFiles = Get-RSAssessmentInteger -Object $patch -Name binaryFiles
    $concerns = @(
        Get-RSAssessmentProperty -Object $patch -Name logicalConcerns)
    $sortedConcerns = [string[]]@($concerns | ForEach-Object { [string]$_ })
    [Array]::Sort($sortedConcerns, [StringComparer]::Ordinal)
    for ($index = 1; $index -lt $sortedConcerns.Count; ++$index) {
        if ($sortedConcerns[$index - 1] -ceq $sortedConcerns[$index]) {
            throw [IO.InvalidDataException]::new(
                "Candidate assessment repeats logical concern '$($sortedConcerns[$index])'.")
        }
    }

    $newToolchains = Get-RSAssessmentInteger `
        -Object $build `
        -Name additionalToolchainFamilies
    $runtimeLeaves = Get-RSAssessmentInteger `
        -Object $integration `
        -Name nonSystemRuntimeDependencies
    $adapterClass = [string](Get-RSAssessmentProperty `
            -Object $maintenance `
            -Name adapterClass)
    $adapterFloor = switch ($adapterClass) {
        'same-language-or-existing-c-abi' { 3 }
        'new-cross-language-ffi' { 5 }
        default {
            throw [IO.InvalidDataException]::new(
                "Unsupported assessment adapter class '$adapterClass'.")
        }
    }
    $initialAdapterDays = Get-RSAssessmentInteger `
        -Object $maintenance `
        -Name initialAdapterDays
    if ($initialAdapterDays -lt $adapterFloor) {
        throw [IO.InvalidDataException]::new(
            "initialAdapterDays is below the frozen $adapterFloor-day floor.")
    }

    $declaredPatchDays = Get-RSAssessmentInteger `
        -Object $maintenance `
        -Name declaredInitialPatchDays
    $patchFloor = [Math]::Max(
        $declaredPatchDays,
        [Math]::Max(
            2 * $concerns.Count,
            [int][Math]::Ceiling([decimal]$lines / [decimal]256)))
    $initialPatchDays = Get-RSAssessmentInteger `
        -Object $maintenance `
        -Name initialPatchDays
    if ($initialPatchDays -ne $patchFloor) {
        throw [IO.InvalidDataException]::new(
            "initialPatchDays must equal the mechanically derived $patchFloor-day floor.")
    }

    $perUpgradeFloor =
        1 +
        [int][Math]::Ceiling([decimal]$files / [decimal]8) +
        $newToolchains +
        $runtimeLeaves
    $perUpgradeDays = Get-RSAssessmentInteger `
        -Object $maintenance `
        -Name perUpgradeDays
    if ($perUpgradeDays -lt $perUpgradeFloor) {
        throw [IO.InvalidDataException]::new(
            "perUpgradeDays is below the mechanically derived $perUpgradeFloor-day floor.")
    }

    $touchesSecuritySurface = Get-RSAssessmentBoolean `
        -Object $maintenance `
        -Name patchTouchesParserDecoderAccounting
    $securityFloor = if ($touchesSecuritySurface) { 2 } else { 1 }
    $securityDays = Get-RSAssessmentInteger `
        -Object $maintenance `
        -Name annualSecurityResponseDays
    if ($securityDays -lt $securityFloor) {
        throw [IO.InvalidDataException]::new(
            "annualSecurityResponseDays is below the frozen $securityFloor-day floor.")
    }
    $noticeDays = Get-RSAssessmentInteger `
        -Object $maintenance `
        -Name releaseNoticeDays
    if ($noticeDays -lt 1) {
        throw [IO.InvalidDataException]::new(
            'releaseNoticeDays is below the frozen one-day floor.')
    }
    $ownerDays = [int][Math]::Ceiling(
        ([decimal]$initialAdapterDays +
            [decimal]$initialPatchDays +
            [decimal]9 * [decimal]$perUpgradeDays +
            [decimal]3 * [decimal]$securityDays +
            [decimal]$noticeDays) *
        [decimal]1.20)
    if ((Get-RSAssessmentInteger -Object $maintenance -Name ownerDays) -ne $ownerDays) {
        throw [IO.InvalidDataException]::new(
            "ownerDays must equal the mechanically derived value $ownerDays.")
    }

    $failures = [Collections.Generic.List[string]]::new()
    if ($ownerDays -gt (Get-RSAssessmentInteger -Object $caps -Name ownerDays)) {
        $failures.Add('owner-days-cap')
    }
    if ($lines -gt (Get-RSAssessmentInteger -Object $caps -Name changedMaintainedLines)) {
        $failures.Add('changed-lines-cap')
    }
    if ($files -gt (Get-RSAssessmentInteger -Object $caps -Name changedMaintainedFiles)) {
        $failures.Add('changed-files-cap')
    }
    if ($concerns.Count -gt (Get-RSAssessmentInteger -Object $caps -Name logicalConcerns)) {
        $failures.Add('logical-concerns-cap')
    }
    if ($newToolchains -gt
        (Get-RSAssessmentInteger -Object $caps -Name newToolchainFamilies)) {
        $failures.Add('new-toolchain-cap')
    }
    if ($runtimeLeaves -gt
        (Get-RSAssessmentInteger -Object $caps -Name newNonSystemRuntimeLeaves)) {
        $failures.Add('runtime-leaf-cap')
    }
    if ($binaryFiles -ne 0) {
        $failures.Add('binary-patch-forbidden')
    }

    return [pscustomobject]@{
        admitted = $failures.Count -eq 0
        failures = $failures.ToArray()
        ownerDays = $ownerDays
        initialPatchDays = $initialPatchDays
        perUpgradeDays = $perUpgradeDays
        changedMaintainedLines = $lines
        changedMaintainedFiles = $files
        logicalConcerns = $concerns.Count
        newToolchainFamilies = $newToolchains
        nonSystemRuntimeDependencies = $runtimeLeaves
    }
}

function Get-RSCandidateStaticRatings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Candidate,

        [Parameter(Mandatory)]
        [object]$CandidateUniverse,

        [string]$TerminalEvidenceModule = (
            Join-Path $PSScriptRoot '..\TerminalEvidence.psm1')
    )

    Import-Module $TerminalEvidenceModule -Force -ErrorAction Stop
    $ledger = Get-RSAssessmentProperty -Object $Candidate -Name rawLedger
    $admission = Get-RSCandidateAdmission `
        -Candidate $Candidate `
        -CandidateUniverse $CandidateUniverse
    if (-not $admission.admitted) {
        throw [IO.InvalidDataException]::new(
            "Candidate assessment exceeds adaptation caps: $($admission.failures -join ',').")
    }

    $ratings = [Collections.Generic.List[object]]::new()

    $boundary = Get-RSAssessmentProperty -Object $ledger -Name boundary
    $boundaryKind = [string](Get-RSAssessmentProperty -Object $boundary -Name boundaryKind)
    $crossed = Get-RSAssessmentInteger -Object $boundary -Name crossedAbiElements
    $identity = Get-RSAssessmentBoolean -Object $boundary -Name generatedIdentityComplete
    $immutable = Get-RSAssessmentBoolean `
        -Object $boundary `
        -Name immutableSnapshotTransaction
    $lifetime = Get-RSAssessmentBoolean `
        -Object $boundary `
        -Name borrowedLifetimeExplicit
    $copyAdapter = [string](Get-RSAssessmentProperty `
            -Object $boundary `
            -Name ownedCopyAdapter)
    $publicConcerns = Get-RSAssessmentInteger `
        -Object $boundary `
        -Name maintainedPublicSurfaceConcerns
    $privateTypePatch = Get-RSAssessmentBoolean `
        -Object $boundary `
        -Name maintainedPrivateTypePatch
    $recurringInternalAdaptation = Get-RSAssessmentBoolean `
        -Object $boundary `
        -Name recurringInternalAdaptation
    $extensiveGlue = Get-RSAssessmentBoolean `
        -Object $boundary `
        -Name extensiveMaintainedGlue
    $embeddable = Get-RSAssessmentBoolean -Object $boundary -Name embeddable
    $c1 = if (-not $embeddable) {
        0
    }
    elseif ($boundaryKind -in @('versioned-stable-c-abi', 'fully-source-vendored') -and
        $crossed -le 24 -and $identity -and $immutable -and
        -not $privateTypePatch) {
        5
    }
    elseif ($boundaryKind -in @(
            'versioned-stable-c-abi',
            'public-c-abi',
            'fully-source-vendored') -and
        $crossed -le 64 -and $identity -and $lifetime -and
        $copyAdapter -in @('none', 'thin') -and $publicConcerns -le 1) {
        4
    }
    elseif ($boundaryKind -cne 'undocumented' -and
        $crossed -le 128 -and $identity -and $immutable -and
        ($publicConcerns -ge 2 -or $copyAdapter -ceq 'conservative')) {
        3
    }
    elseif ($boundaryKind -cne 'undocumented' -and
        ($crossed -gt 128 -or $recurringInternalAdaptation -or
            $copyAdapter -ceq 'broad')) {
        2
    }
    elseif ($boundaryKind -ceq 'undocumented' -or $extensiveGlue) {
        1
    }
    else {
        throw [IO.InvalidDataException]::new(
            'Candidate boundary facts do not match any frozen C1 anchor.')
    }
    $ratings.Add((New-RSAssessmentRating `
            -CriterionId C1 `
            -Rating $c1 `
            -EvidenceSha256 (Get-RSJcsSha256 -InputObject $boundary)))

    $functional = Get-RSAssessmentProperty -Object $ledger -Name functional
    $assertions = Get-RSAssessmentInteger `
        -Object $functional `
        -Name upstreamAssertionCount
    $differences = Get-RSAssessmentInteger `
        -Object $functional `
        -Name compatibilityDifferences
    $normalizations = Get-RSAssessmentInteger `
        -Object $functional `
        -Name requiredNormalizationBehaviors
    $c2 = if ($assertions -ge 1000 -and $differences -eq 0 -and $normalizations -eq 0) {
        5
    }
    elseif ($assertions -ge 250 -and $differences -le 2 -and $normalizations -eq 0) {
        4
    }
    elseif ($assertions -lt 250 -or
        ($normalizations -ge 1 -and $normalizations -le 2)) {
        3
    }
    else {
        2
    }
    $ratings.Add((New-RSAssessmentRating `
            -CriterionId C2 `
            -Rating $c2 `
            -EvidenceSha256 (Get-RSJcsSha256 -InputObject $functional)))

    $resources = Get-RSAssessmentProperty -Object $ledger -Name resources
    $typed = Get-RSAssessmentBoolean `
        -Object $resources `
        -Name allLimitsTypedAtOwningSites
    $hardBounded = Get-RSAssessmentBoolean `
        -Object $resources `
        -Name allHardLimitsPrebounded
    $resourceLifetime = Get-RSAssessmentBoolean `
        -Object $resources `
        -Name borrowedLifetimesExplicit
    $reentrancy = Get-RSAssessmentBoolean `
        -Object $resources `
        -Name callbackReentrancyExplicit
    $charges = Get-RSAssessmentInteger `
        -Object $resources `
        -Name conservativeChargeCount
    $resourceConcerns = Get-RSAssessmentInteger `
        -Object $resources `
        -Name maintainedResourceConcerns
    $enclosingCap = Get-RSAssessmentBoolean `
        -Object $resources `
        -Name nonAdversarialTransientUsesEnclosingCap
    $capacityReduced = Get-RSAssessmentBoolean `
        -Object $resources `
        -Name measuredCapacityReduced
    $mostPrebounded = Get-RSAssessmentBoolean `
        -Object $resources `
        -Name mostAllocationsPrebounded
    $coarseCaps = Get-RSAssessmentBoolean `
        -Object $resources `
        -Name coarseAllocatorCaps
    $defensiveSerialization = Get-RSAssessmentBoolean `
        -Object $resources `
        -Name requiresSubstantialDefensiveSerialization
    $mandatoryUnbounded = Get-RSAssessmentBoolean `
        -Object $resources `
        -Name anyMandatoryAllocationUnboundedBeforeWork
    $lifetimeTeardownProved = Get-RSAssessmentBoolean `
        -Object $resources `
        -Name lifetimeTeardownSafetyProved
    $remainingResourceIssues = $charges + $resourceConcerns
    $c3 = if ($mandatoryUnbounded -or -not $lifetimeTeardownProved) {
        0
    }
    elseif ($typed -and $hardBounded -and $resourceLifetime -and $reentrancy -and
        $charges -eq 0 -and $resourceConcerns -eq 0) {
        5
    }
    elseif ($hardBounded -and $resourceLifetime -and $reentrancy -and
        $remainingResourceIssues -le 1) {
        4
    }
    elseif ($hardBounded -and $remainingResourceIssues -ge 2 -and
        $capacityReduced) {
        3
    }
    elseif ($mostPrebounded -and $enclosingCap) {
        2
    }
    elseif ($coarseCaps -and $defensiveSerialization) {
        1
    }
    else {
        throw [IO.InvalidDataException]::new(
            'Candidate resource facts do not match any frozen C3 anchor.')
    }
    $ratings.Add((New-RSAssessmentRating `
            -CriterionId C3 `
            -Rating $c3 `
            -EvidenceSha256 (Get-RSJcsSha256 -InputObject $resources)))

    $build = Get-RSAssessmentProperty -Object $ledger -Name build
    $exactBuild =
        (Get-RSAssessmentBoolean -Object $build -Name exactRawReproducible) -and
        (Get-RSAssessmentBoolean -Object $build -Name offline) -and
        (Get-RSAssessmentBoolean -Object $build -Name x64) -and
        (Get-RSAssessmentBoolean -Object $build -Name arm64) -and
        (Get-RSAssessmentBoolean -Object $build -Name x64AsanConsumer)
    if (-not $exactBuild) {
        throw [IO.InvalidDataException]::new(
            'A complete candidate assessment lacks a mandatory exact build fact.')
    }
    $toolchains = Get-RSAssessmentInteger `
        -Object $build `
        -Name additionalToolchainFamilies
    $wrapper = Get-RSAssessmentBoolean -Object $build -Name maintainedBuildWrapper
    $repositoryMsvc = Get-RSAssessmentBoolean -Object $build -Name repositoryMsvcOnly
    if (($repositoryMsvc -and $toolchains -ne 0) -or
        (-not $repositoryMsvc -and $toolchains -eq 0)) {
        throw [IO.InvalidDataException]::new(
            'repositoryMsvcOnly and additionalToolchainFamilies are inconsistent.')
    }
    $c4 = if ($repositoryMsvc -and $toolchains -eq 0 -and -not $wrapper) {
        5
    }
    elseif (-not $wrapper -and $toolchains -eq 1) {
        4
    }
    else {
        3
    }
    $ratings.Add((New-RSAssessmentRating `
            -CriterionId C4 `
            -Rating $c4 `
            -EvidenceSha256 (Get-RSJcsSha256 -InputObject $build)))

    $kitty = Get-RSAssessmentProperty -Object $ledger -Name kitty
    $kittyPublic = Get-RSAssessmentBoolean `
        -Object $kitty `
        -Name publicImmutableRepresentation
    $kittyConcerns = Get-RSAssessmentInteger `
        -Object $kitty `
        -Name maintainedPreallocationOrExposureConcerns
    $kittyCopies = Get-RSAssessmentInteger -Object $kitty -Name conservativeCopies
    $pngCallbacks = Get-RSAssessmentInteger -Object $kitty -Name embedderPngCallbacks
    $parserWork = Get-RSAssessmentBoolean `
        -Object $kitty `
        -Name substantialAuthoritativeParserWork
    $preDecodeAccounting = Get-RSAssessmentBoolean `
        -Object $kitty `
        -Name preDecodeTypedAccounting
    $staticStateRepresentable = Get-RSAssessmentBoolean `
        -Object $kitty `
        -Name requiredStaticStateRepresentable
    $basicDirectOnly = Get-RSAssessmentBoolean `
        -Object $kitty `
        -Name basicDirectImagesOnly
    if (-not $preDecodeAccounting -and $kittyConcerns -eq 0) {
        throw [IO.InvalidDataException]::new(
            'Missing pre-decode Kitty accounting requires a maintained concern.')
    }
    $c5 = if ($kittyPublic -and $preDecodeAccounting -and $kittyConcerns -eq 0 -and
        $kittyCopies -eq 0 -and $pngCallbacks -eq 0) {
        5
    }
    elseif ($kittyPublic -and $preDecodeAccounting -and $kittyConcerns -le 1 -and
        $kittyCopies -eq 0 -and $pngCallbacks -le 1) {
        4
    }
    elseif ($preDecodeAccounting -and -not $parserWork -and
        ($kittyConcerns -ge 2 -or $kittyCopies -ge 1)) {
        3
    }
    elseif ($staticStateRepresentable -and $parserWork) {
        2
    }
    elseif ($basicDirectOnly) {
        1
    }
    else {
        0
    }
    $ratings.Add((New-RSAssessmentRating `
            -CriterionId C5 `
            -Rating $c5 `
            -EvidenceSha256 (Get-RSJcsSha256 -InputObject $kitty)))

    $integration = Get-RSAssessmentProperty -Object $ledger -Name integration
    $modelKind = [string](Get-RSAssessmentProperty -Object $integration -Name modelKind)
    $ownsUi = Get-RSAssessmentBoolean -Object $integration -Name candidateOwnsUi
    $rendererReady = Get-RSAssessmentBoolean `
        -Object $integration `
        -Name rendererReadyState
    $repositoryRuntime = Get-RSAssessmentBoolean `
        -Object $integration `
        -Name repositoryRuntime
    $runtimeDependencies = Get-RSAssessmentInteger `
        -Object $integration `
        -Name nonSystemRuntimeDependencies
    $foreignLoop = Get-RSAssessmentBoolean -Object $integration -Name foreignEventLoop
    $substantialUiAdaptation = Get-RSAssessmentBoolean `
        -Object $integration `
        -Name substantialShapingAccessibilityAdaptation
    $removedUi = Get-RSAssessmentBoolean `
        -Object $integration `
        -Name candidateUiAssumptionsRemoved
    $c6 = if (-not $ownsUi -and $rendererReady -and
        $modelKind -ceq 'native-c-cpp' -and $repositoryRuntime -and
        $runtimeDependencies -eq 0 -and -not $foreignLoop) {
        5
    }
    elseif (-not $ownsUi -and $rendererReady -and
        $modelKind -in @('source-static-foreign', 'private-runtime-foreign') -and
        $runtimeDependencies -le 1 -and -not $foreignLoop) {
        4
    }
    elseif (-not $ownsUi -and $rendererReady -and $substantialUiAdaptation) {
        3
    }
    elseif (-not $ownsUi -and $removedUi) {
        2
    }
    elseif ($modelKind -ceq 'foreign-control') {
        1
    }
    else {
        0
    }
    $ratings.Add((New-RSAssessmentRating `
            -CriterionId C6 `
            -Rating $c6 `
            -EvidenceSha256 (Get-RSJcsSha256 -InputObject $integration)))

    $effects = Get-RSAssessmentProperty -Object $ledger -Name effects
    $trustsOutput = Get-RSAssessmentBoolean `
        -Object $effects `
        -Name requiresTrustingTerminalOutput
    $changesPolicy = Get-RSAssessmentBoolean `
        -Object $effects `
        -Name requiresShellPolicyChange
    $typedEffects = Get-RSAssessmentBoolean `
        -Object $effects `
        -Name typedPromptInputOutput
    $pwdEffects = Get-RSAssessmentBoolean -Object $effects -Name pwdTitleHyperlink
    $ordering = Get-RSAssessmentBoolean -Object $effects -Name writeResponseOrdering
    $sideChannel = Get-RSAssessmentBoolean `
        -Object $effects `
        -Name authenticatedSideChannelComposes
    $idleSideChannel = Get-RSAssessmentBoolean `
        -Object $effects `
        -Name idleTruthSideChannelOnly
    $copiedCallbacks = Get-RSAssessmentBoolean `
        -Object $effects `
        -Name boundedCopiedCallbacks
    $missingSemantics = Get-RSAssessmentInteger `
        -Object $effects `
        -Name optionalSemanticClassesAbsent
    $titlePwdOnly = Get-RSAssessmentBoolean -Object $effects -Name titlePwdOnly
    $c7 = if ($trustsOutput -or $changesPolicy -or -not $sideChannel) {
        0
    }
    elseif ($typedEffects -and $pwdEffects -and $ordering -and -not $idleSideChannel) {
        5
    }
    elseif ($pwdEffects -and $ordering -and $idleSideChannel) {
        4
    }
    elseif ($copiedCallbacks -and $missingSemantics -le 1) {
        3
    }
    elseif ($titlePwdOnly) {
        2
    }
    else {
        1
    }
    $ratings.Add((New-RSAssessmentRating `
            -CriterionId C7 `
            -Rating $c7 `
            -EvidenceSha256 (Get-RSJcsSha256 -InputObject $effects)))

    $maintenance = Get-RSAssessmentProperty -Object $ledger -Name maintenance
    $ownerDays = $admission.ownerDays
    $c8 = if ($ownerDays -le 15) {
        5
    }
    elseif ($ownerDays -le 30) {
        4
    }
    elseif ($ownerDays -le 60) {
        3
    }
    elseif ($ownerDays -le 90) {
        2
    }
    elseif ($ownerDays -le 150) {
        1
    }
    else {
        0
    }
    $c8Evidence = [ordered]@{
        patch = Get-RSAssessmentProperty -Object $ledger -Name patch
        maintenance = $maintenance
    }
    $ratings.Add((New-RSAssessmentRating `
            -CriterionId C8 `
            -Rating $c8 `
            -EvidenceSha256 (Get-RSJcsSha256 -InputObject $c8Evidence)))

    $legal = Get-RSAssessmentProperty -Object $ledger -Name legal
    $closure = Get-RSAssessmentBoolean `
        -Object $legal `
        -Name distributableClosureComplete
    $permissive = Get-RSAssessmentBoolean -Object $legal -Name permissiveClosure
    $notices = Get-RSAssessmentBoolean -Object $legal -Name automatedNotices
    $provenance = Get-RSAssessmentBoolean -Object $legal -Name verifiedProvenance
    $manualNotice = [string](Get-RSAssessmentProperty `
            -Object $legal `
            -Name manualNoticeWork)
    $copyleft = Get-RSAssessmentBoolean `
        -Object $legal `
        -Name copyleftReviewComplexity
    $c10 = if (-not $closure -or -not $provenance -or $manualNotice -ceq 'unresolved') {
        0
    }
    elseif ($permissive -and $notices -and $runtimeDependencies -eq 0) {
        5
    }
    elseif ($notices -and $runtimeDependencies -le 2 -and
        $manualNotice -in @('none', 'bounded')) {
        4
    }
    elseif ($runtimeDependencies -le 5 -and $manualNotice -in @('none', 'bounded')) {
        3
    }
    elseif ($runtimeDependencies -le 10 -or $manualNotice -ceq 'recurring') {
        2
    }
    elseif ($copyleft) {
        1
    }
    else {
        1
    }
    $c10Evidence = [ordered]@{
        legal = $legal
        nonSystemRuntimeDependencies = $runtimeDependencies
    }
    $ratings.Add((New-RSAssessmentRating `
            -CriterionId C10 `
            -Rating $c10 `
            -EvidenceSha256 (Get-RSJcsSha256 -InputObject $c10Evidence)))

    $testability = Get-RSAssessmentProperty -Object $ledger -Name testability
    $publicFuzz = Get-RSAssessmentBoolean `
        -Object $testability `
        -Name publicDeterministicFuzzSeam
    $fuzzSeam = Get-RSAssessmentBoolean `
        -Object $testability `
        -Name deterministicFuzzSeam
    $parserSeam = Get-RSAssessmentBoolean `
        -Object $testability `
        -Name deterministicParserSeam
    $unitSeam = Get-RSAssessmentBoolean `
        -Object $testability `
        -Name deterministicUnitSeam
    $testAssertions = Get-RSAssessmentInteger `
        -Object $testability `
        -Name upstreamAssertionCount
    $windows = Get-RSAssessmentBoolean -Object $testability -Name windowsExecution
    $allocationInjection = Get-RSAssessmentBoolean `
        -Object $testability `
        -Name allocationFailureInjection
    $diagnostics = [string](Get-RSAssessmentProperty `
            -Object $testability `
            -Name diagnostics)
    if ($publicFuzz -and -not $fuzzSeam) {
        throw [IO.InvalidDataException]::new(
            'A public deterministic fuzz seam must also be a deterministic fuzz seam.')
    }
    $c11 = if ($publicFuzz -and $testAssertions -ge 1000 -and $windows -and
        $allocationInjection -and $diagnostics -ceq 'bounded-structured') {
        5
    }
    elseif ($fuzzSeam -and $testAssertions -ge 250 -and
        $windows -and $diagnostics -in @('bounded-structured', 'actionable')) {
        4
    }
    elseif ($parserSeam -and $testAssertions -ge 100 -and $windows -and
        $diagnostics -in @(
            'bounded-structured',
            'actionable',
            'harness-supplied')) {
        3
    }
    elseif ($unitSeam -and ($testAssertions -lt 100 -or -not $windows)) {
        2
    }
    elseif (-not $unitSeam -and $diagnostics -ceq 'integration-only') {
        1
    }
    elseif (-not $unitSeam -and -not $parserSeam -and -not $fuzzSeam) {
        0
    }
    else {
        throw [IO.InvalidDataException]::new(
            'Candidate testability facts do not match any frozen C11 anchor.')
    }
    $ratings.Add((New-RSAssessmentRating `
            -CriterionId C11 `
            -Rating $c11 `
            -EvidenceSha256 (Get-RSJcsSha256 -InputObject $testability)))

    return $ratings.ToArray()
}

function Get-RSMeasurementMedian {
    param(
        [Parameter(Mandatory)]
        [object]$Measurement
    )

    $samples = @(
        Get-RSAssessmentProperty -Object $Measurement -Name samples)
    if ($samples.Count -eq 0) {
        throw [IO.InvalidDataException]::new('Performance measurement has no samples.')
    }
    $values = [decimal[]]@(
        foreach ($sample in $samples) {
            if ($sample -isnot [string] -or
                [string]$sample -cnotmatch '^(?:0|[1-9][0-9]{0,19})$') {
                throw [IO.InvalidDataException]::new(
                    'Performance sample is not one canonical uint64 decimal string.')
            }
            [decimal]::Parse(
                [string]$sample,
                [Globalization.NumberStyles]::None,
                [Globalization.CultureInfo]::InvariantCulture)
        })
    [Array]::Sort($values)
    $middle = [Math]::Floor($values.Count / 2)
    if (($values.Count % 2) -eq 1) {
        return $values[$middle]
    }
    return ($values[$middle - 1] + $values[$middle]) / [decimal]2
}

function Get-RSCandidatePerformanceRating {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Measurements,

        [string]$TerminalEvidenceModule = (
            Join-Path $PSScriptRoot '..\TerminalEvidence.psm1')
    )

    Import-Module $TerminalEvidenceModule -Force -ErrorAction Stop
    $required = [ordered]@{
        'create-destroy-us' = $null
        'parse-8mib-us' = $null
        'dirty-snapshot-us' = $null
        'eight-session-retained-bytes' = $null
        'added-release-package-bytes' = $null
    }
    foreach ($measurement in $Measurements) {
        $metricId = [string](Get-RSAssessmentProperty `
                -Object $measurement `
                -Name metricId)
        if ($required.Contains($metricId)) {
            if ($null -ne $required[$metricId]) {
                throw [IO.InvalidDataException]::new(
                    "Duplicate performance measurement '$metricId'.")
            }
            $required[$metricId] = $measurement
        }
    }
    foreach ($metricId in $required.Keys) {
        if ($null -eq $required[$metricId]) {
            throw [IO.InvalidDataException]::new(
                "Required performance measurement '$metricId' is missing.")
        }
    }
    $medians = [ordered]@{}
    foreach ($metricId in $required.Keys) {
        $medians[$metricId] = Get-RSMeasurementMedian -Measurement $required[$metricId]
    }

    $rating = 1
    foreach ($budget in @(
            [pscustomobject]@{
                rating = 5
                create = 1000
                parse = 32000
                dirty = 500
                retained = 67108864
                package = 4194304
            },
            [pscustomobject]@{
                rating = 4
                create = 2500
                parse = 64000
                dirty = 1000
                retained = 100663296
                package = 8388608
            },
            [pscustomobject]@{
                rating = 3
                create = 5000
                parse = 125000
                dirty = 2500
                retained = 167772160
                package = 16777216
            },
            [pscustomobject]@{
                rating = 2
                create = 10000
                parse = 250000
                dirty = 5000
                retained = 268435456
                package = 33554432
            })) {
        if ($medians['create-destroy-us'] -le $budget.create -and
            $medians['parse-8mib-us'] -le $budget.parse -and
            $medians['dirty-snapshot-us'] -le $budget.dirty -and
            $medians['eight-session-retained-bytes'] -le $budget.retained -and
            $medians['added-release-package-bytes'] -le $budget.package) {
            $rating = $budget.rating
            break
        }
    }
    $evidence = [ordered]@{
        aggregate = 'exact-median'
        medians = [ordered]@{
            'create-destroy-us' = $medians['create-destroy-us'].ToString(
                '0.############################',
                [Globalization.CultureInfo]::InvariantCulture)
            'parse-8mib-us' = $medians['parse-8mib-us'].ToString(
                '0.############################',
                [Globalization.CultureInfo]::InvariantCulture)
            'dirty-snapshot-us' = $medians['dirty-snapshot-us'].ToString(
                '0.############################',
                [Globalization.CultureInfo]::InvariantCulture)
            'eight-session-retained-bytes' =
                $medians['eight-session-retained-bytes'].ToString(
                    '0.############################',
                    [Globalization.CultureInfo]::InvariantCulture)
            'added-release-package-bytes' =
                $medians['added-release-package-bytes'].ToString(
                    '0.############################',
                    [Globalization.CultureInfo]::InvariantCulture)
        }
    }
    return New-RSAssessmentRating `
        -CriterionId C9 `
        -Rating $rating `
        -EvidenceSha256 (Get-RSJcsSha256 -InputObject $evidence)
}

function New-RSCandidateAssessmentManifest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RepoRoot,

        [Parameter(Mandatory)]
        [object]$Candidate,

        [Parameter(Mandatory)]
        [ValidatePattern('^[a-z0-9][a-z0-9._-]{0,63}$')]
        [string]$AuthorId,

        [Parameter(Mandatory)]
        [string]$FrozenUtc,

        [Parameter(Mandatory)]
        [object[]]$Approval
    )

    $resolvedRepoRoot = [IO.Path]::GetFullPath($RepoRoot)
    $universePath = @(
        @(
            'Specs\Terminal\TerminalEngineCandidateUniverse.v3.json',
            'Specs\Terminal\TerminalEngineCandidateUniverse.v2.json',
            'Specs\Terminal\TerminalEngineCandidateUniverse.v1.json') |
            ForEach-Object { Join-Path $resolvedRepoRoot $_ } |
            Where-Object { [IO.File]::Exists($_) } |
            Select-Object -First 1)
    if ($universePath.Count -ne 1) {
        throw [IO.FileNotFoundException]::new(
            'No frozen TerminalEngine candidate universe is present.')
    }
    $universePath = [string]$universePath[0]
    $anchorsPath = Join-Path $resolvedRepoRoot (
        'Specs\Terminal\TerminalEngineScoreAnchors.v1.json')
    $schemaPath = Join-Path $resolvedRepoRoot (
        'Specs\Terminal\TerminalEngineCandidateAssessment.schema.json')
    $evaluatorModulePath = Join-Path $resolvedRepoRoot (
        'Tools\TerminalEngineEvaluator.psm1')
    foreach ($path in @(
            $universePath,
            $anchorsPath,
            $schemaPath,
            $evaluatorModulePath)) {
        if (-not [IO.File]::Exists($path)) {
            throw [IO.FileNotFoundException]::new(
                'Candidate-assessment generation input is missing.',
                $path)
        }
    }

    Import-Module $evaluatorModulePath -Force -ErrorAction Stop
    $universe = Read-RSCurrentCandidateUniverse -RepoRoot $resolvedRepoRoot
    $candidateId = [string](Get-RSAssessmentProperty `
            -Object $Candidate `
            -Name candidateId)
    $seedCandidateId = [string](Get-RSAssessmentProperty `
            -Object $Candidate `
            -Name seedCandidateId)
    foreach ($identifier in @($candidateId, $seedCandidateId)) {
        if ($identifier -cnotmatch '^[a-z0-9][a-z0-9._-]{0,63}$') {
            throw [IO.InvalidDataException]::new(
                "Candidate-assessment identifier '$identifier' is invalid.")
        }
    }
    $seedMatches = @(
        @($universe['seedCandidates']) |
            Where-Object {
                [string]$_['candidateId'] -ceq $seedCandidateId
            })
    if ($seedMatches.Count -ne 1 -or
        [string]$seedMatches[0]['adaptationPermission'] -cne
            'allowed-existing-subsystem' -or
        [string]$seedMatches[0]['allowedNominationId'] -cne $candidateId) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$candidateId' is not the frozen nomination for '$seedCandidateId'.")
    }
    foreach ($name in @('pin', 'pinSha256', 'sourceArchiveSha256')) {
        if ([string](Get-RSAssessmentProperty -Object $Candidate -Name $name) -cne
            [string]$seedMatches[0][$name]) {
            throw [IO.InvalidDataException]::new(
                "Candidate '$candidateId' $name differs from its frozen seed.")
        }
    }

    $glueAuthorIds = [string[]]@(
        Get-RSAssessmentProperty -Object $Candidate -Name glueAuthorIds)
    if ($glueAuthorIds.Count -eq 0) {
        throw [IO.InvalidDataException]::new(
            'Candidate assessment requires at least one glue author.')
    }
    $orderedGlueAuthorIds = [string[]]@($glueAuthorIds)
    [Array]::Sort($orderedGlueAuthorIds, [StringComparer]::Ordinal)
    for ($index = 0; $index -lt $orderedGlueAuthorIds.Count; ++$index) {
        if ($orderedGlueAuthorIds[$index] -cnotmatch
            '^[a-z0-9][a-z0-9._-]{0,63}$' -or
            ($index -gt 0 -and
                $orderedGlueAuthorIds[$index - 1] -ceq
                    $orderedGlueAuthorIds[$index])) {
            throw [IO.InvalidDataException]::new(
                'Candidate glue authors must be unique bounded identifiers.')
        }
    }

    $candidateForDigest = [ordered]@{
        candidateId = $candidateId
        seedCandidateId = $seedCandidateId
        pin = [string](Get-RSAssessmentProperty -Object $Candidate -Name pin)
        pinSha256 = [string](Get-RSAssessmentProperty `
                -Object $Candidate `
                -Name pinSha256)
        sourceArchiveSha256 = [string](Get-RSAssessmentProperty `
                -Object $Candidate `
                -Name sourceArchiveSha256)
        rawLedger = Get-RSAssessmentProperty -Object $Candidate -Name rawLedger
    }
    $ledgerSha256 = Get-RSCandidateLedgerSha256 `
        -Candidate $candidateForDigest
    $admission = Get-RSCandidateAdmission `
        -Candidate $candidateForDigest `
        -CandidateUniverse $universe
    if (-not [bool]$admission.admitted) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$candidateId' exceeds the frozen adaptation caps.")
    }
    $null = Get-RSCandidateStaticRatings `
        -Candidate $candidateForDigest `
        -CandidateUniverse $universe

    $requiredApprovals = [int]$universe['adaptationPolicy'][
        'patchMetricAlgorithm']['requiredIndependentApprovals']
    if ($Approval.Count -ne $requiredApprovals) {
        throw [IO.InvalidDataException]::new(
            "Candidate '$candidateId' requires exactly $requiredApprovals approvals.")
    }
    $approvalById =
        [Collections.Generic.SortedDictionary[string, object]]::new(
            [StringComparer]::Ordinal)
    $reviewerIds =
        [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($record in $Approval) {
        $properties = [string[]]@(
            if ($record -is [Collections.IDictionary]) {
                $record.Keys | ForEach-Object { [string]$_ }
            }
            else {
                $record.PSObject.Properties.Name
            })
        [Array]::Sort($properties, [StringComparer]::Ordinal)
        if (($properties -join ',') -cne
            'approvalId,decision,ledgerSha256,reviewedUtc,reviewerId,reviewerRole') {
            throw [IO.InvalidDataException]::new(
                'Candidate-assessment approval has a non-closed shape.')
        }
        $approvalId = [string](Get-RSAssessmentProperty `
                -Object $record `
                -Name approvalId)
        $reviewerId = [string](Get-RSAssessmentProperty `
                -Object $record `
                -Name reviewerId)
        if ($approvalId -cnotmatch '^[a-z0-9][a-z0-9._-]{0,63}$' -or
            $reviewerId -cnotmatch '^[a-z0-9][a-z0-9._-]{0,63}$' -or
            $approvalById.ContainsKey($approvalId) -or
            -not $reviewerIds.Add($reviewerId)) {
            throw [IO.InvalidDataException]::new(
                'Candidate-assessment approval ids and reviewers must be unique bounded identifiers.')
        }
        $approvalById[$approvalId] = $record
        if ($reviewerId -ceq $AuthorId -or
            $orderedGlueAuthorIds -ccontains $reviewerId -or
            [string](Get-RSAssessmentProperty `
                -Object $record `
                -Name reviewerRole) -cne
                'independent-candidate-assessor' -or
            [string](Get-RSAssessmentProperty `
                -Object $record `
                -Name decision) -cne 'approved' -or
            [string](Get-RSAssessmentProperty `
                -Object $record `
                -Name ledgerSha256) -cne $ledgerSha256) {
            throw [IO.InvalidDataException]::new(
                "Candidate-assessment approval '$approvalId' is not an independent approval of the exact ledger.")
        }
    }
    $orderedApprovals = [Collections.Generic.List[object]]::new()
    foreach ($record in $approvalById.Values) {
        $orderedApprovals.Add($record)
    }

    $assessmentCandidate = [ordered]@{
        candidateId = $candidateId
        seedCandidateId = $seedCandidateId
        pin = $candidateForDigest.pin
        pinSha256 = $candidateForDigest.pinSha256
        sourceArchiveSha256 = $candidateForDigest.sourceArchiveSha256
        rawLedger = $candidateForDigest.rawLedger
        ledgerSha256 = $ledgerSha256
        glueAuthorIds = $orderedGlueAuthorIds
        approvals = [object[]]$orderedApprovals.ToArray()
    }
    Import-Module (
        Join-Path $resolvedRepoRoot 'Tools\TerminalEvidence.psm1') `
        -Force `
        -ErrorAction Stop
    $manifest = [ordered]@{
        formatId = 'red-salamander-terminal-engine-candidate-assessment'
        version = 1
        frozenUtc = $FrozenUtc
        authorId = $AuthorId
        candidateUniverseSha256 = (
            Get-FileHash -LiteralPath $universePath -Algorithm SHA256
        ).Hash.ToLowerInvariant()
        scoreAnchorsSha256 = (
            Get-FileHash -LiteralPath $anchorsPath -Algorithm SHA256
        ).Hash.ToLowerInvariant()
        candidates = [object[]]@($assessmentCandidate)
    }
    $json = [Text.Encoding]::UTF8.GetString(
        (ConvertTo-RSJcsUtf8Bytes -InputObject $manifest))
    if (-not (Test-Json -Json $json -SchemaFile $schemaPath -ErrorAction Stop)) {
        throw [IO.InvalidDataException]::new(
            'Generated candidate assessment does not satisfy its closed schema.')
    }
    return $manifest
}

Export-ModuleMember -Function `
    Assert-RSAssessmentDescriptorBuildTruth, `
    Get-RSCandidateLedgerSha256, `
    Get-RSCandidateAdmission, `
    Get-RSCandidateStaticRatings, `
    Get-RSCandidatePerformanceRating, `
    New-RSCandidateAssessmentManifest
