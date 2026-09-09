Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$modulePath = Join-Path $repoRoot (
    'Tools\TerminalEngine\TerminalEngineCandidateAssessment.psm1')
Import-Module $modulePath -Force

function New-RSTestAssessmentCandidate {
    $hash = '1' * 64
    return [ordered]@{
        candidateId = 'synthetic-ffi'
        seedCandidateId = 'synthetic-seed'
        pin = 'synthetic-pin+rs-v1'
        pinSha256 = '2' * 64
        sourceArchiveSha256 = '3' * 64
        rawLedger = [ordered]@{
            patch = [ordered]@{
                upstreamTree = '4' * 40
                patchRelativePath = 'External/TerminalEngine/CandidatePatches/synthetic-ffi.patch'
                patchSha256 = '5' * 64
                patchSizeBytes = 100000
                resultTree = '6' * 40
                changedMaintainedLines = 2400
                changedMaintainedFiles = 8
                binaryFiles = 0
                logicalConcerns = @(
                    'abi-snapshot-identity',
                    'build-reproducibility-packaging',
                    'parser-side-effect-bounds',
                    'teardown-concurrency',
                    'typed-resource-accounting-admission')
                concernMappingRelativePath = (
                    'External/TerminalEngine/CandidatePatches/synthetic-ffi.concerns.json')
                concernMappingSha256 = '7' * 64
                collectionDescriptorRelativePath = (
                    'red-salamander/collection-descriptor.v1.json')
                collectionDescriptorSha256 = '8' * 64
                buildKind = 'candidate-wrapper-v1'
                maintainedBuildWrapper = $true
            }
            boundary = [ordered]@{
                boundaryKind = 'public-c-abi'
                crossedAbiElements = 20
                generatedIdentityComplete = $true
                immutableSnapshotTransaction = $true
                borrowedLifetimeExplicit = $true
                ownedCopyAdapter = 'thin'
                maintainedPublicSurfaceConcerns = 1
                maintainedPrivateTypePatch = $false
                recurringInternalAdaptation = $true
                extensiveMaintainedGlue = $false
                embeddable = $true
            }
            functional = [ordered]@{
                upstreamAssertionCount = 2000
                compatibilityDifferences = 0
                requiredNormalizationBehaviors = 0
            }
            resources = [ordered]@{
                allLimitsTypedAtOwningSites = $true
                allHardLimitsPrebounded = $true
                borrowedLifetimesExplicit = $true
                callbackReentrancyExplicit = $true
                conservativeChargeCount = 0
                maintainedResourceConcerns = 1
                nonAdversarialTransientUsesEnclosingCap = $false
                measuredCapacityReduced = $false
                mostAllocationsPrebounded = $true
                coarseAllocatorCaps = $false
                requiresSubstantialDefensiveSerialization = $false
                anyMandatoryAllocationUnboundedBeforeWork = $false
                lifetimeTeardownSafetyProved = $true
            }
            build = [ordered]@{
                repositoryMsvcOnly = $false
                additionalToolchainFamilies = 1
                exactRawReproducible = $true
                maintainedBuildWrapper = $true
                offline = $true
                x64 = $true
                arm64 = $true
                x64AsanConsumer = $true
            }
            kitty = [ordered]@{
                publicImmutableRepresentation = $true
                maintainedPreallocationOrExposureConcerns = 1
                conservativeCopies = 0
                embedderPngCallbacks = 0
                substantialAuthoritativeParserWork = $false
                preDecodeTypedAccounting = $true
                requiredStaticStateRepresentable = $true
                basicDirectImagesOnly = $false
            }
            integration = [ordered]@{
                modelKind = 'source-static-foreign'
                candidateOwnsUi = $false
                rendererReadyState = $true
                repositoryRuntime = $false
                nonSystemRuntimeDependencies = 0
                foreignEventLoop = $false
                substantialShapingAccessibilityAdaptation = $false
                candidateUiAssumptionsRemoved = $false
            }
            effects = [ordered]@{
                typedPromptInputOutput = $true
                pwdTitleHyperlink = $true
                writeResponseOrdering = $true
                authenticatedSideChannelComposes = $true
                idleTruthSideChannelOnly = $true
                boundedCopiedCallbacks = $true
                optionalSemanticClassesAbsent = 0
                titlePwdOnly = $false
                requiresTrustingTerminalOutput = $false
                requiresShellPolicyChange = $false
            }
            maintenance = [ordered]@{
                initialAdapterDays = 5
                adapterClass = 'new-cross-language-ffi'
                declaredInitialPatchDays = 10
                initialPatchDays = 10
                perUpgradeDays = 3
                annualSecurityResponseDays = 2
                releaseNoticeDays = 1
                ownerDays = 59
                patchTouchesParserDecoderAccounting = $true
            }
            legal = [ordered]@{
                distributableClosureComplete = $true
                permissiveClosure = $true
                automatedNotices = $true
                verifiedProvenance = $true
                manualNoticeWork = 'none'
                copyleftReviewComplexity = $false
            }
            testability = [ordered]@{
                publicDeterministicFuzzSeam = $true
                deterministicFuzzSeam = $true
                deterministicParserSeam = $true
                deterministicUnitSeam = $true
                upstreamAssertionCount = 2000
                windowsExecution = $true
                allocationFailureInjection = $true
                diagnostics = 'bounded-structured'
            }
        }
        ledgerSha256 = $hash
        glueAuthorIds = @('candidate-author')
        approvals = @()
    }
}

function New-RSTestMeasurement {
    param(
        [Parameter(Mandatory)]
        [string]$MetricId,

        [Parameter(Mandatory)]
        [string[]]$Samples
    )

    return [ordered]@{
        metricId = $MetricId
        unit = if ($MetricId.EndsWith('-bytes')) { 'bytes' } else { 'microseconds' }
        warmupCount = 0
        samples = $Samples
    }
}

Describe 'TerminalEngine reviewed candidate assessment derivation' {
    It 'keeps the frozen no-scored-candidate outcome representable' {
        $schema = Get-Content -LiteralPath (
            Join-Path $repoRoot (
                'Specs\Terminal\TerminalEngineCandidateAssessment.schema.json')) -Raw |
            ConvertFrom-Json

        $schema.properties.candidates.minItems | Should Be 0
    }

    BeforeAll {
        $script:universe = Get-Content -LiteralPath (
            Join-Path $repoRoot 'Specs\Terminal\TerminalEngineCandidateUniverse.v1.json') -Raw |
            ConvertFrom-Json -AsHashtable -Depth 100 -NoEnumerate -DateKind String
    }

    It 'derives the effective eight-file synthetic FFI admission and every static anchor' {
        $candidate = New-RSTestAssessmentCandidate
        $admission = Get-RSCandidateAdmission `
            -Candidate $candidate `
            -CandidateUniverse $script:universe
        $ratings = @(Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $script:universe)

        $admission.admitted | Should Be $true
        $admission.ownerDays | Should Be 59
        $admission.perUpgradeDays | Should Be 3
        @($ratings | ForEach-Object { "$($_.criterionId):$($_.rating)" }) -join ',' |
            Should Be 'C1:4,C2:5,C3:4,C4:3,C5:4,C6:4,C7:4,C8:3,C10:5,C11:5'
        (Get-RSCandidateLedgerSha256 -Candidate $candidate) |
            Should Match '^[0-9a-f]{64}$'
    }

    It 'mechanically excludes a ninth synthetic FFI file through the owner-day formula' {
        $candidate = New-RSTestAssessmentCandidate
        $candidate.rawLedger.patch.changedMaintainedFiles = 9
        $candidate.rawLedger.maintenance.perUpgradeDays = 4
        $candidate.rawLedger.maintenance.ownerDays = 70
        $admission = Get-RSCandidateAdmission `
            -Candidate $candidate `
            -CandidateUniverse $script:universe

        $admission.admitted | Should Be $false
        ($admission.failures -ccontains 'owner-days-cap') | Should Be $true
        $admission.ownerDays | Should Be 70
    }

    It 'caps a reproducible candidate with a maintained build wrapper at C4 rating 3' {
        $candidate = New-RSTestAssessmentCandidate
        $rating = @(Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $script:universe |
                Where-Object criterionId -CEQ 'C4')

        $rating.Count | Should Be 1
        $rating[0].rating | Should Be 3
        $rating[0].anchorId | Should Be 'c4-r3'
    }

    It 'reserves C4 ratings 4 and 5 for candidates without a maintained wrapper' {
        $candidate = New-RSTestAssessmentCandidate
        $candidate.rawLedger.patch.buildKind =
            'provider-generic-zig-msvc-v1'
        $candidate.rawLedger.patch.maintainedBuildWrapper = $false
        $candidate.rawLedger.build.maintainedBuildWrapper = $false
        $rating = @(Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $script:universe |
                Where-Object criterionId -CEQ 'C4')
        $rating[0].rating | Should Be 4

        $candidate.rawLedger.build.repositoryMsvcOnly = $true
        $candidate.rawLedger.build.additionalToolchainFamilies = 0
        $rating = @(Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $script:universe |
                Where-Object criterionId -CEQ 'C4')
        $rating[0].rating | Should Be 5
    }

    It 'rejects a candidate-wrapper descriptor that claims no maintained wrapper' {
        $candidate = New-RSTestAssessmentCandidate
        $candidate.rawLedger.build.maintainedBuildWrapper = $false

        $caught = $null
        try {
            Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $script:universe | Out-Null
        }
        catch {
            $caught = $_.Exception
        }
        $caught | Should Not BeNullOrEmpty
        $caught.Message | Should Match (
            'maintainedBuildWrapper must equal.*candidate-wrapper-v1')
    }

    It 'requires every C1 conjunct and does not promote a conservative boundary' {
        $candidate = New-RSTestAssessmentCandidate
        $candidate.rawLedger.boundary.boundaryKind = 'versioned-stable-c-abi'
        $candidate.rawLedger.boundary.maintainedPublicSurfaceConcerns = 0
        $rating = @(Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $script:universe |
                Where-Object criterionId -CEQ 'C1')
        $rating[0].rating | Should Be 5

        $candidate.rawLedger.boundary.maintainedPrivateTypePatch = $true
        $rating = @(Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $script:universe |
                Where-Object criterionId -CEQ 'C1')
        $rating[0].rating | Should Be 4

        $candidate.rawLedger.boundary.boundaryKind = 'source-boundary'
        $candidate.rawLedger.boundary.maintainedPrivateTypePatch = $false
        $candidate.rawLedger.boundary.ownedCopyAdapter = 'conservative'
        $candidate.rawLedger.boundary.maintainedPublicSurfaceConcerns = 1
        $rating = @(Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $script:universe |
                Where-Object criterionId -CEQ 'C1')
        $rating[0].rating | Should Be 3

        $candidate.rawLedger.boundary.ownedCopyAdapter = 'thin'
        $rating = @(Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $script:universe |
                Where-Object criterionId -CEQ 'C1')
        $rating[0].rating | Should Be 2
    }

    It 'does not use C2 rating 3 to forgive more than two compatibility differences' {
        $candidate = New-RSTestAssessmentCandidate
        $candidate.rawLedger.functional.compatibilityDifferences = 3
        $rating = @(Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $script:universe |
                Where-Object criterionId -CEQ 'C2')

        $rating[0].rating | Should Be 2
        $rating[0].anchorId | Should Be 'c2-r2'
    }

    It 'matches C3 issue counts, capacity loss, enclosing caps, and fatal uncertainty exactly' {
        $candidate = New-RSTestAssessmentCandidate
        $candidate.rawLedger.resources.conservativeChargeCount = 1
        $candidate.rawLedger.resources.maintainedResourceConcerns = 1
        $candidate.rawLedger.resources.measuredCapacityReduced = $true
        $rating = @(Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $script:universe |
                Where-Object criterionId -CEQ 'C3')
        $rating[0].rating | Should Be 3

        $candidate.rawLedger.resources.measuredCapacityReduced = $false
        {
            Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $script:universe
        } | Should Throw 'Candidate resource facts do not match any frozen C3 anchor.'

        $candidate = New-RSTestAssessmentCandidate
        $candidate.rawLedger.resources.allHardLimitsPrebounded = $false
        $candidate.rawLedger.resources.mostAllocationsPrebounded = $true
        $candidate.rawLedger.resources.nonAdversarialTransientUsesEnclosingCap = $true
        $rating = @(Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $script:universe |
                Where-Object criterionId -CEQ 'C3')
        $rating[0].rating | Should Be 2

        $candidate.rawLedger.resources.mostAllocationsPrebounded = $false
        $candidate.rawLedger.resources.nonAdversarialTransientUsesEnclosingCap = $false
        $candidate.rawLedger.resources.coarseAllocatorCaps = $true
        $candidate.rawLedger.resources.requiresSubstantialDefensiveSerialization = $true
        $rating = @(Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $script:universe |
                Where-Object criterionId -CEQ 'C3')
        $rating[0].rating | Should Be 1

        $candidate.rawLedger.resources.anyMandatoryAllocationUnboundedBeforeWork = $true
        $rating = @(Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $script:universe |
                Where-Object criterionId -CEQ 'C3')
        $rating[0].rating | Should Be 0
    }

    It 'rejects an impossible C4 zero-toolchain non-repository build fact' {
        $candidate = New-RSTestAssessmentCandidate
        $candidate.rawLedger.build.additionalToolchainFamilies = 0

        {
            Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $script:universe
        } | Should Throw (
            'repositoryMsvcOnly and additionalToolchainFamilies are inconsistent.')
    }

    It 'requires pre-decode C5 accounting and sends conservative copies to rating 3' {
        $candidate = New-RSTestAssessmentCandidate
        $candidate.rawLedger.kitty.maintainedPreallocationOrExposureConcerns = 0
        $rating = @(Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $script:universe |
                Where-Object criterionId -CEQ 'C5')
        $rating[0].rating | Should Be 5

        $candidate.rawLedger.kitty.preDecodeTypedAccounting = $false
        {
            Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $script:universe
        } | Should Throw (
            'Missing pre-decode Kitty accounting requires a maintained concern.')

        $candidate.rawLedger.kitty.maintainedPreallocationOrExposureConcerns = 1
        $rating = @(Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $script:universe |
                Where-Object criterionId -CEQ 'C5')
        $rating[0].rating | Should Be 0

        $candidate.rawLedger.kitty.preDecodeTypedAccounting = $true
        $candidate.rawLedger.kitty.maintainedPreallocationOrExposureConcerns = 2
        $rating = @(Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $script:universe |
                Where-Object criterionId -CEQ 'C5')
        $rating[0].rating | Should Be 3

        $candidate.rawLedger.kitty.maintainedPreallocationOrExposureConcerns = 1
        $candidate.rawLedger.kitty.conservativeCopies = 1
        $rating = @(Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $script:universe |
                Where-Object criterionId -CEQ 'C5')
        $rating[0].rating | Should Be 3
    }

    It 'requires automated notices for C10 rating 4' {
        $candidate = New-RSTestAssessmentCandidate
        $candidate.rawLedger.legal.automatedNotices = $false
        $rating = @(Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $script:universe |
                Where-Object criterionId -CEQ 'C10')

        $rating[0].rating | Should Be 3
        $rating[0].anchorId | Should Be 'c10-r3'
    }

    It 'distinguishes fuzz, parser, unit, Windows, and diagnostic C11 seams' {
        $candidate = New-RSTestAssessmentCandidate
        $candidate.rawLedger.testability.publicDeterministicFuzzSeam = $false
        $candidate.rawLedger.testability.deterministicFuzzSeam = $false
        $candidate.rawLedger.testability.diagnostics = 'actionable'
        $rating = @(Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $script:universe |
                Where-Object criterionId -CEQ 'C11')
        $rating[0].rating | Should Be 3

        $candidate.rawLedger.testability.deterministicFuzzSeam = $true
        $rating = @(Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $script:universe |
                Where-Object criterionId -CEQ 'C11')
        $rating[0].rating | Should Be 4

        $candidate.rawLedger.testability.deterministicFuzzSeam = $false
        $candidate.rawLedger.testability.windowsExecution = $false
        $rating = @(Get-RSCandidateStaticRatings `
                -Candidate $candidate `
                -CandidateUniverse $script:universe |
                Where-Object criterionId -CEQ 'C11')
        $rating[0].rating | Should Be 2
    }

    It 'uses exact even-sample medians and derives C9 without a provider category' {
        $measurements = @(
            (New-RSTestMeasurement `
                -MetricId 'create-destroy-us' `
                -Samples @('999', '1001'))
            (New-RSTestMeasurement `
                -MetricId 'parse-8mib-us' `
                -Samples @('31999', '32001'))
            (New-RSTestMeasurement `
                -MetricId 'dirty-snapshot-us' `
                -Samples @('499', '501'))
            (New-RSTestMeasurement `
                -MetricId 'eight-session-retained-bytes' `
                -Samples @('67108864'))
            (New-RSTestMeasurement `
                -MetricId 'added-release-package-bytes' `
                -Samples @('4194304'))
        )
        $rating = Get-RSCandidatePerformanceRating -Measurements $measurements

        $rating.criterionId | Should Be 'C9'
        $rating.rating | Should Be 5
        $rating.anchorId | Should Be 'c9-r5'
        $rating.evidenceSha256 | Should Match '^[0-9a-f]{64}$'
    }
}
