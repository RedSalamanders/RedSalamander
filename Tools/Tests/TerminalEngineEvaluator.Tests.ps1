Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$evaluatorModulePath = Join-Path $repoRoot 'Tools\TerminalEngineEvaluator.psm1'
$evidenceModulePath = Join-Path $repoRoot 'Tools\TerminalEvidence.psm1'
. (Join-Path $repoRoot 'Tools\TestRunPlan.ps1')
Import-Module $evidenceModulePath -Force -Global
Import-Module $evaluatorModulePath -Force
$evaluatorModule = Get-Module -Name TerminalEngineEvaluator
$round4FrozenInputLeaves = [string[]]@(
    & $evaluatorModule {
        [string[]]@($script:Round4FrozenInputLeaves)
    })
$neutralityModule = Import-Module `
    (Join-Path $repoRoot (
        'Tools\TerminalEngine\TerminalEngineCollectionNeutrality.psm1')) `
    -Force `
    -PassThru
$neutralitySurfaceLeaves = [string[]]@(
    & $neutralityModule {
        Get-RSTerminalEngineCollectionNeutralitySurface
    })
Remove-Module $neutralityModule -Force

function Assert-RSThrowsType {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Action,

        [Parameter(Mandatory = $true)]
        [type]$ExceptionType,

        [string]$MessageFragment = ''
    )

    $caught = $null
    try {
        & $Action | Out-Null
    } catch {
        $caught = $_.Exception
    }
    if ($null -eq $caught) {
        throw "Expected $($ExceptionType.FullName), but the action did not throw."
    }
    if (-not $ExceptionType.IsAssignableFrom($caught.GetType())) {
        throw "Expected $($ExceptionType.FullName), got $($caught.GetType().FullName): $($caught.Message)"
    }
    if (-not [string]::IsNullOrEmpty($MessageFragment) -and
        -not $caught.Message.Contains(
            $MessageFragment,
            [StringComparison]::Ordinal)) {
        throw (
            "Expected exception message containing '$MessageFragment', " +
            "got '$($caught.Message)'.")
    }
}

function Get-RSTestSha256 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value
    )

    $bytes = [Text.UTF8Encoding]::new($false, $true).GetBytes($Value)
    $sha256 = [Security.Cryptography.SHA256]::Create()
    try {
        $digest = $sha256.ComputeHash($bytes)
    } finally {
        $sha256.Dispose()
    }
    return ([BitConverter]::ToString($digest) -replace '-', '').ToLowerInvariant()
}

function Get-RSTestFrozenCandidateSet {
    param(
        [string]$CandidateReviewSha256 = (Get-RSTestSha256 -Value 'candidate-review'),

        [string]$CandidateAssessmentSha256 = (
            Get-RSTestSha256 -Value 'candidate-assessment'),

        [string]$CandidateUniverseSha256 = (
            (Get-FileHash `
                -LiteralPath (Join-Path $repoRoot 'Specs\Terminal\TerminalEngineCandidateUniverse.v1.json') `
                -Algorithm SHA256).Hash.ToLowerInvariant())
    )

    $commonBootstrapInputs = [object[]]@(
        [ordered]@{
            inputId = 'source-alpha'
            inputKind = 'source-archive'
            version = 'alpha-pin'
            fileName = 'alpha.tar.gz'
            sha256 = Get-RSTestSha256 -Value 'alpha-source'
            retrievalUrl = 'https://example.invalid/alpha.tar.gz'
            consumers = @('alpha')
        },
        [ordered]@{
            inputId = 'source-beta'
            inputKind = 'source-archive'
            version = 'beta-pin'
            fileName = 'beta.tar.gz'
            sha256 = Get-RSTestSha256 -Value 'beta-source'
            retrievalUrl = 'https://example.invalid/beta.tar.gz'
            consumers = @('beta')
        },
        [ordered]@{
            inputId = 'source-failed'
            inputKind = 'source-archive'
            version = 'failed-pin'
            fileName = 'failed.tar.gz'
            sha256 = Get-RSTestSha256 -Value 'failed-source'
            retrievalUrl = 'https://example.invalid/failed.tar.gz'
            consumers = @('failed')
        })
    $arm64Inputs = [object[]]@(
        $commonBootstrapInputs +
        [ordered]@{
            inputId = 'toolchain-msvc'
            inputKind = 'toolchain-input'
            version = 'fixture-arm64'
            fileName = 'msvc-arm64.zip'
            sha256 = Get-RSTestSha256 -Value 'arm64-toolchain'
            retrievalUrl = 'https://example.invalid/msvc-arm64.zip'
            consumers = @('alpha', 'beta', 'failed')
        })
    $x64Inputs = [object[]]@(
        $commonBootstrapInputs +
        [ordered]@{
            inputId = 'toolchain-msvc'
            inputKind = 'toolchain-input'
            version = 'fixture-x64'
            fileName = 'msvc-x64.zip'
            sha256 = Get-RSTestSha256 -Value 'x64-toolchain'
            retrievalUrl = 'https://example.invalid/msvc-x64.zip'
            consumers = @('alpha', 'beta', 'failed')
        })
    $x64Bootstrap = Get-RSJcsSha256 -InputObject $x64Inputs
    $arm64Bootstrap = Get-RSJcsSha256 -InputObject $arm64Inputs
    return [ordered]@{
        formatId = 'red-salamander-terminal-engine-candidate-set'
        version = 1
        frozenUtc = '2026-07-30T18:00:03Z'
        authorId = 'fixture-candidate-set-author'
        candidateUniverseSha256 = $CandidateUniverseSha256
        candidateReviewSha256 = $CandidateReviewSha256
        candidateAssessmentSha256 = $CandidateAssessmentSha256
        nominationOutcomes = [object[]]@(
            [ordered]@{
                seedCandidateId = 'alpha'
                pin = 'alpha-pin'
                pinSha256 = Get-RSTestSha256 -Value 'alpha-pin'
                sourceArchiveSha256 = Get-RSTestSha256 -Value 'alpha-source'
                outcome = 'admitted'
                nominationId = 'alpha'
                ledgerSha256 = Get-RSTestSha256 -Value 'alpha-ledger'
            },
            [ordered]@{
                seedCandidateId = 'beta'
                pin = 'beta-pin'
                pinSha256 = Get-RSTestSha256 -Value 'beta-pin'
                sourceArchiveSha256 = Get-RSTestSha256 -Value 'beta-source'
                outcome = 'admitted'
                nominationId = 'beta'
                ledgerSha256 = Get-RSTestSha256 -Value 'beta-ledger'
            },
            [ordered]@{
                seedCandidateId = 'failed'
                pin = 'failed-pin'
                pinSha256 = Get-RSTestSha256 -Value 'failed-pin'
                sourceArchiveSha256 = Get-RSTestSha256 -Value 'failed-source'
                outcome = 'admitted'
                nominationId = 'failed'
                ledgerSha256 = Get-RSTestSha256 -Value 'failed-ledger'
            })
        bootstrapInputSets = [object[]]@(
            [ordered]@{
                platform = 'ARM64'
                inputs = $arm64Inputs
                sha256 = $arm64Bootstrap
            },
            [ordered]@{
                platform = 'x64'
                inputs = $x64Inputs
                sha256 = $x64Bootstrap
            }
        )
        candidates = [object[]]@(
            [ordered]@{
                candidateId = 'alpha'
                pin = 'alpha-pin'
                pinSha256 = Get-RSTestSha256 -Value 'alpha-pin'
                sourceArchiveSha256 = Get-RSTestSha256 -Value 'alpha-source'
                scoredOrControl = 'scored'
                disposition = 'collect'
            },
            [ordered]@{
                candidateId = 'beta'
                pin = 'beta-pin'
                pinSha256 = Get-RSTestSha256 -Value 'beta-pin'
                sourceArchiveSha256 = Get-RSTestSha256 -Value 'beta-source'
                scoredOrControl = 'scored'
                disposition = 'collect'
            },
            [ordered]@{
                candidateId = 'failed'
                pin = 'failed-pin'
                pinSha256 = Get-RSTestSha256 -Value 'failed-pin'
                sourceArchiveSha256 = Get-RSTestSha256 -Value 'failed-source'
                scoredOrControl = 'scored'
                disposition = 'collect'
            },
            [ordered]@{
                candidateId = 'web-control'
                pin = 'web-control-pin'
                pinSha256 = Get-RSTestSha256 -Value 'web-control-pin'
                scoredOrControl = 'control'
                constraintId = 'web-runtime-second-stack'
            }
        )
    }
}

function New-RSTestAssessmentCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [object]$FrozenCandidate,

        [ValidateSet('alpha', 'beta', 'failed')]
        [string]$Profile
    )

    $candidateId = [string]$FrozenCandidate.candidateId
    $patchText = "fixture-patch-$candidateId"
    $isAlpha = $Profile -ceq 'alpha'
    $declaredPatchDays = if ($isAlpha) { 12 } else { 2 }
    $ownerDays = if ($isAlpha) { 45 } else { 33 }
    $boundaryKind = if ($isAlpha) { 'versioned-stable-c-abi' } else { 'public-c-abi' }
    $publicConcerns = if ($isAlpha) { 0 } else { 1 }
    $candidate = [ordered]@{
        candidateId = $candidateId
        seedCandidateId = $candidateId
        pin = [string]$FrozenCandidate.pin
        pinSha256 = [string]$FrozenCandidate.pinSha256
        sourceArchiveSha256 = [string]$FrozenCandidate.sourceArchiveSha256
        rawLedger = [ordered]@{
            patch = [ordered]@{
                upstreamTree = '4' * 40
                patchRelativePath = "External/TerminalEngine/CandidatePatches/$candidateId.patch"
                patchSha256 = Get-RSTestSha256 -Value $patchText
                patchSizeBytes = [Text.UTF8Encoding]::new($false).GetByteCount($patchText)
                resultTree = '6' * 40
                changedMaintainedLines = 100
                changedMaintainedFiles = 1
                binaryFiles = 0
                logicalConcerns = @('abi-snapshot-identity')
                concernMappingRelativePath = (
                    "External/TerminalEngine/CandidatePatches/$candidateId.concerns.json")
                concernMappingSha256 = '0' * 64
                collectionDescriptorRelativePath = (
                    'red-salamander/collection-descriptor.v1.json')
                collectionDescriptorSha256 = '8' * 64
                buildKind = 'provider-generic-cmake-msvc-v1'
                maintainedBuildWrapper = $false
            }
            boundary = [ordered]@{
                boundaryKind = $boundaryKind
                crossedAbiElements = 20
                generatedIdentityComplete = $true
                immutableSnapshotTransaction = $true
                borrowedLifetimeExplicit = $true
                ownedCopyAdapter = if ($isAlpha) { 'none' } else { 'thin' }
                maintainedPublicSurfaceConcerns = $publicConcerns
                maintainedPrivateTypePatch = $false
                recurringInternalAdaptation = $false
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
                maintainedResourceConcerns = 0
                nonAdversarialTransientUsesEnclosingCap = $false
                measuredCapacityReduced = $false
                mostAllocationsPrebounded = $true
                coarseAllocatorCaps = $false
                requiresSubstantialDefensiveSerialization = $false
                anyMandatoryAllocationUnboundedBeforeWork = $false
                lifetimeTeardownSafetyProved = $true
            }
            build = [ordered]@{
                repositoryMsvcOnly = $true
                additionalToolchainFamilies = 0
                exactRawReproducible = $true
                maintainedBuildWrapper = $false
                offline = $true
                x64 = $true
                arm64 = $true
                x64AsanConsumer = $true
            }
            kitty = [ordered]@{
                publicImmutableRepresentation = $true
                maintainedPreallocationOrExposureConcerns = 0
                conservativeCopies = 0
                embedderPngCallbacks = 0
                substantialAuthoritativeParserWork = $false
                preDecodeTypedAccounting = $true
                requiredStaticStateRepresentable = $true
                basicDirectImagesOnly = $false
            }
            integration = [ordered]@{
                modelKind = 'native-c-cpp'
                candidateOwnsUi = $false
                rendererReadyState = $true
                repositoryRuntime = $true
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
                idleTruthSideChannelOnly = $false
                boundedCopiedCallbacks = $true
                optionalSemanticClassesAbsent = 0
                titlePwdOnly = $false
                requiresTrustingTerminalOutput = $false
                requiresShellPolicyChange = $false
            }
            maintenance = [ordered]@{
                initialAdapterDays = 3
                adapterClass = 'same-language-or-existing-c-abi'
                declaredInitialPatchDays = $declaredPatchDays
                initialPatchDays = $declaredPatchDays
                perUpgradeDays = 2
                annualSecurityResponseDays = 1
                releaseNoticeDays = 1
                ownerDays = $ownerDays
                patchTouchesParserDecoderAccounting = $false
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
        ledgerSha256 = '0' * 64
        glueAuthorIds = @('fixture-glue')
        approvals = @()
    }
    Import-Module $evidenceModulePath -Force -Global
    $mappingPayload = [ordered]@{
        candidateId = $candidateId
        upstreamTree = [string]$candidate.rawLedger.patch.upstreamTree
        resultTree = [string]$candidate.rawLedger.patch.resultTree
        patchSha256 = [string]$candidate.rawLedger.patch.patchSha256
        pathConcerns = [object[]]@(
            [ordered]@{
                path = 'fixture.txt'
                concerns = @('abi-snapshot-identity')
            })
    }
    $candidate.rawLedger.patch.concernMappingSha256 = (
        Get-RSJcsSha256 -InputObject $mappingPayload)
    Import-Module (
        Join-Path $repoRoot (
            'Tools\TerminalEngine\TerminalEngineCandidateAssessment.psm1')) -Force
    $candidate.ledgerSha256 = Get-RSCandidateLedgerSha256 -Candidate $candidate
    $candidate.approvals = [object[]]@(
        [ordered]@{
            approvalId = "$candidateId-review-a"
            reviewerId = 'fixture-reviewer-a'
            reviewerRole = 'independent-candidate-assessor'
            ledgerSha256 = $candidate.ledgerSha256
            decision = 'approved'
            reviewedUtc = '2026-07-30T18:00:01Z'
        },
        [ordered]@{
            approvalId = "$candidateId-review-b"
            reviewerId = 'fixture-reviewer-b'
            reviewerRole = 'independent-candidate-assessor'
            ledgerSha256 = $candidate.ledgerSha256
            decision = 'approved'
            reviewedUtc = '2026-07-30T18:00:02Z'
        })
    return $candidate
}

function Get-RSTestCandidateSetSha256 {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$CandidateSet
    )

    Import-Module $evidenceModulePath -Force -Global
    return Get-RSJcsSha256 -InputObject $CandidateSet
}

function New-RSTestPassGates {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CandidateId,

        [Parameter(Mandatory = $true)]
        [string]$LaneKey
    )

    return [object[]]@(
        1..12 | ForEach-Object {
            [ordered]@{
                gateId = "G$_"
                status = 'pass'
                resultCategory = 'passed'
                evidenceSha256 = Get-RSTestSha256 -Value "$CandidateId|$LaneKey|G$_"
            }
        }
    )
}

function New-RSTestEarlyFailureGates {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CandidateId,

        [Parameter(Mandatory = $true)]
        [string]$LaneKey,

        [int]$FailedGateNumber = 7
    )

    return [object[]]@(
        1..12 | ForEach-Object {
            if ($_ -lt $FailedGateNumber) {
                [ordered]@{
                    gateId = "G$_"
                    status = 'pass'
                    resultCategory = 'passed'
                    evidenceSha256 = Get-RSTestSha256 -Value "$CandidateId|$LaneKey|G$_"
                }
            } elseif ($_ -eq $FailedGateNumber) {
                [ordered]@{
                    gateId = "G$_"
                    status = 'fail'
                    resultCategory = 'mandatory-failure'
                    evidenceSha256 = Get-RSTestSha256 -Value "$CandidateId|$LaneKey|G$_|failure"
                }
            } else {
                [ordered]@{
                    gateId = "G$_"
                    status = 'not-run'
                    resultCategory = "failed-G$FailedGateNumber"
                    evidenceSha256 = $null
                }
            }
        }
    )
}

function New-RSTestCompleteCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$FrozenCandidate,

        [Parameter(Mandatory = $true)]
        [string]$LaneKey,

        [AllowNull()]
        [object]$AssessmentCandidate = $null
    )

    $candidateId = [string]$FrozenCandidate.candidateId
    $ratingById = @{}
    if ($null -ne $AssessmentCandidate) {
        Import-Module (
            Join-Path $repoRoot (
                'Tools\TerminalEngine\TerminalEngineCandidateAssessment.psm1')) -Force
        $universe = Get-Content -LiteralPath (
            Join-Path $repoRoot (
                'Specs\Terminal\TerminalEngineCandidateUniverse.v1.json')) -Raw |
            ConvertFrom-Json -AsHashtable -Depth 100 -NoEnumerate -DateKind String
        foreach ($rating in @(
                Get-RSCandidateStaticRatings `
                    -Candidate $AssessmentCandidate `
                    -CandidateUniverse $universe)) {
            $ratingById[[string]$rating.criterionId] = $rating
        }
    }
    $scoreObjects = [object[]]@(
        @(1, 10, 11) + @(2..9) | ForEach-Object {
            $criterionId = "C$_"
            $rating = if ($ratingById.ContainsKey($criterionId)) {
                $ratingById[$criterionId]
            }
            else {
                [pscustomobject]@{
                    rating = 5
                    anchorId = "c$_-r5"
                    evidenceSha256 = Get-RSTestSha256 `
                        -Value "$candidateId|$LaneKey|score.c$_|r5"
                }
            }
            [ordered]@{
                evidenceId = "score.c$_"
                sha256 = [string]$rating.evidenceSha256
                sizeBytes = '1'
                resultCategory = [string]$rating.anchorId
            }
        }
    )
    $scoreById = [Collections.Generic.Dictionary[string, object]]::new(
        [StringComparer]::Ordinal)
    foreach ($record in $scoreObjects) {
        [void]$scoreById.Add([string]$record.evidenceId, $record)
    }
    foreach ($evidenceId in @(
            'abi.build-identity',
            'abi.fuzz-probe',
            'abi.layout-pair',
            'abi.source-dependencies')) {
        [void]$scoreById.Add(
            $evidenceId,
            [ordered]@{
                evidenceId = $evidenceId
                sha256 = Get-RSTestSha256 `
                    -Value "$candidateId|$LaneKey|$evidenceId|verified"
                sizeBytes = '1'
                resultCategory = 'verified'
            })
    }
    [void]$scoreById.Add(
        'score.c7-proof',
        [ordered]@{
            evidenceId = 'score.c7-proof'
            sha256 = Get-RSTestSha256 `
                -Value "$candidateId|$LaneKey|score.c7-proof|verified"
            sizeBytes = '1564'
            resultCategory = 'verified'
        })
    $scoreIds = [string[]]@($scoreById.Keys)
    [Array]::Sort($scoreIds, [StringComparer]::Ordinal)
    $scoreObjects = [object[]]@(
        foreach ($scoreId in $scoreIds) {
            $scoreById[$scoreId]
        }
    )
    $measurements = [object[]]@(
        [ordered]@{
            metricId = 'added-release-package-bytes'
            unit = 'bytes'
            warmupCount = 0
            samples = [object[]]@('4000000')
        },
        [ordered]@{
            metricId = 'create-destroy-us'
            unit = 'microseconds'
            warmupCount = 10
            samples = [object[]]@('500')
        },
        [ordered]@{
            metricId = 'crossed-abi-elements'
            unit = 'count'
            warmupCount = 0
            samples = [object[]]@('20')
        },
        [ordered]@{
            metricId = 'dirty-snapshot-us'
            unit = 'microseconds'
            warmupCount = 10
            samples = [object[]]@('250')
        },
        [ordered]@{
            metricId = 'eight-session-retained-bytes'
            unit = 'bytes'
            warmupCount = 1
            samples = [object[]]@('50000000')
        },
        [ordered]@{
            metricId = 'non-system-runtime-dependencies'
            unit = 'count'
            warmupCount = 0
            samples = [object[]]@('0')
        },
        [ordered]@{
            metricId = 'parse-8mib-us'
            unit = 'microseconds'
            warmupCount = 5
            samples = [object[]]@('20000')
        }
    )
    return [ordered]@{
        candidateId = $candidateId
        pin = [string]$FrozenCandidate.pin
        pinSha256 = [string]$FrozenCandidate.pinSha256
        outcome = 'complete'
        gates = New-RSTestPassGates -CandidateId $candidateId -LaneKey $LaneKey
        evidenceObjects = $scoreObjects
        measurements = $measurements
    }
}

function New-RSTestEarlyFailureCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$FrozenCandidate,

        [Parameter(Mandatory = $true)]
        [string]$LaneKey
    )

    $candidateId = [string]$FrozenCandidate.candidateId
    return [ordered]@{
        candidateId = $candidateId
        pin = [string]$FrozenCandidate.pin
        pinSha256 = [string]$FrozenCandidate.pinSha256
        outcome = 'early-failure'
        gates = New-RSTestEarlyFailureGates -CandidateId $candidateId -LaneKey $LaneKey
        earlyFailure = [ordered]@{
            firstFailedGate = 'G7'
            failureCode = 'kitty-zlib-missing'
            reproducerSha256 = Get-RSTestSha256 -Value "$candidateId|$LaneKey|reproducer"
            reproducerSizeBytes = '256'
        }
    }
}

function New-RSTestControlCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$FrozenCandidate,

        [Parameter(Mandatory = $true)]
        [string]$LaneKey
    )

    return [ordered]@{
        candidateId = [string]$FrozenCandidate.candidateId
        pin = [string]$FrozenCandidate.pin
        pinSha256 = [string]$FrozenCandidate.pinSha256
        outcome = 'constraint-control'
        controlFailure = [ordered]@{
            constraintId = [string]$FrozenCandidate.constraintId
            evidenceSha256 = Get-RSTestSha256 -Value "$($FrozenCandidate.candidateId)|$LaneKey|control"
        }
    }
}

function New-RSTestFixtureRepo {
    Import-Module $evidenceModulePath -Force -Global
    $scratch = New-RSTestSandboxScratchDirectory `
        -RepoRoot $repoRoot `
        -Harness 'terminal-engine-evaluator-pester' `
        -Case ([guid]::NewGuid().ToString('N'))
    $fixtureRepo = Join-Path $scratch 'repo'
    $null = New-Item -ItemType Directory -Path (Join-Path $fixtureRepo 'Tools') -Force
    $null = New-Item -ItemType Directory -Path (Join-Path $fixtureRepo 'Specs\Terminal') -Force

    $fixtureLeaves = [string[]]@(
        @(
            'Directory.Build.props',
            'Directory.Build.targets',
            'RuntimeDependencies.props',
            'Tools\Evaluate-TerminalEngines.ps1',
            'Tools\TerminalEngineEvaluator.psm1',
            'Tools\TerminalEvidence.psm1',
            'Tools\TerminalEngine\Invoke-RSIsolatedProcess.ps1',
            'Tools\TerminalEngine\TerminalEngineCollectionDescriptor.psm1',
            'Tools\TerminalEngine\TerminalEngineCollectionDescriptorContract.md',
            'Tools\TerminalEngine\TerminalEngineCollectionDriverContract.md',
            'Tools\TerminalEngine\TerminalEngineCollectionDriverInput.schema.json',
            'Tools\TerminalEngine\TerminalEngineCollectionDriverResult.schema.json',
            'Tools\TerminalEngine\TerminalEngineCollectionNeutrality.psm1',
            'Tools\TerminalEngine\Test-TerminalEngineCollectionNeutrality.ps1',
            'Tools\TerminalEngine\TerminalEnginePeIdentity.psm1',
            'Tools\TerminalEngine\TerminalEngineSourceReview.psm1',
            'Tools\TerminalEngine\TerminalEngineCandidateAssessment.psm1',
            'Tools\TerminalEngine\TerminalEngineC7Proof.psm1',
            'Tools\TerminalEngine\TerminalEngineC7ProofContract.md',
            'Tools\TerminalEngine\TerminalEngineGate0Adapter.h',
            'Tools\TerminalEngine\TerminalEngineGate0AdapterContract.md',
            'Tools\TerminalEngine\TerminalEngineGate0Harness.cpp',
            'Tools\TerminalEngine\TerminalEngineGate0Harness.vcxproj',
            'Tools\TerminalEngine\Invoke-TerminalEngineGate0Harness.ps1',
            'Tools\TerminalEngine\TerminalEngineGate0Layout.h',
            'Tools\TerminalEngine\TerminalEngineGate0Probe.cpp',
            'Tools\TerminalEngine\TerminalEngineGate0Probe.vcxproj',
            'Tools\TerminalEngine\Invoke-TerminalEngineGate0Probe.ps1',
            'Tools\TerminalEngine\TerminalEngineNotices.psm1',
            'Tools\TerminalEngine\Generate-TerminalEngineFixtures.ps1',
            'Tools\TerminalEngine\Generated\TerminalEngineFixtures.v1.h',
            'Specs\Terminal\TerminalEngineEvidence.schema.json',
            'Specs\Terminal\TerminalEngineCandidateUniverse.schema.json',
            'Specs\Terminal\TerminalEngineCandidateUniverseTransition.schema.json',
            'Specs\Terminal\TerminalEngineCandidateUniverseTransition.v3.schema.json',
            'Specs\Terminal\TerminalEngineCandidateUniverseTransition.v4.schema.json',
            'Specs\Terminal\TerminalEngineCandidateUniverse.v1.json',
            'Specs\Terminal\TerminalEngineCandidateSet.schema.json',
            'Specs\Terminal\TerminalEngineCandidateReview.schema.json',
            'Specs\Terminal\TerminalEngineCandidateAssessment.schema.json',
            'Specs\Terminal\TerminalEngineConcernMapping.schema.json',
            'Specs\Terminal\TerminalEngineCollectionDescriptor.schema.json',
            'Specs\Terminal\TerminalEngineCollectionNeutrality.schema.json',
            'Specs\Terminal\TerminalEngineCollectionNeutrality.v1.json',
            'Specs\Terminal\TerminalEngineC7Proof.schema.json',
            'Specs\Terminal\TerminalEngineC7ProofContract.v1.json',
            'Tools\TerminalEngine\Test-TerminalEngineCandidatePatch.ps1',
            'Specs\Terminal\TerminalEngineRequirements.v1.json',
            'Specs\Terminal\TerminalEngineScoreAnchors.v1.json',
            'Specs\Terminal\TerminalEngineFixtureCorpus.v1.json',
            'Specs\Terminal\TerminalEngineMeasurementProtocol.v1.json',
            'vcpkg.json') +
        $round4FrozenInputLeaves +
        $neutralitySurfaceLeaves +
        @(
            'Specs\Terminal\TerminalEngineCandidateUniverse.v2.json',
            'Specs\Terminal\TerminalEngineCandidateUniverse.v3.json'))
    $fixtureLeaves = [string[]]@(
        $fixtureLeaves |
            ForEach-Object { $_.Replace('/', '\') } |
            Sort-Object -Unique)
    foreach ($leaf in $fixtureLeaves) {
        $target = Join-Path $fixtureRepo $leaf
        $null = New-Item -ItemType Directory -Path ([IO.Path]::GetDirectoryName($target)) -Force
        Copy-Item -LiteralPath (Join-Path $repoRoot $leaf) -Destination $target
    }

    # This isolated synthetic-v1 fixture exercises evaluator scoring and provider
    # contracts, not universe-transition governance. Override only the copied
    # module; production named/current readers remain fail-closed.
    $fixtureEvaluatorPath = Join-Path $fixtureRepo (
        'Tools\TerminalEngineEvaluator.psm1')
    $fixtureUniverseOverride = @'

function Read-RSCurrentCandidateUniverseCore {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [switch]$IncludeRecordIdentity
    )
    $manifest = Read-RSCandidateUniverse -RepoRoot $RepoRoot
    if ($IncludeRecordIdentity) {
        $path = Join-Path $RepoRoot (
            'Specs\Terminal\TerminalEngineCandidateUniverse.v1.json')
        return [pscustomobject]@{
            Manifest = $manifest
            Sha256 = Get-RSFileSha256 -Path $path
            FrozenInputSha256ByRelativePath = @{
                'Specs/Terminal/TerminalEngineCollectionNeutrality.v1.json' =
                    Get-RSFileSha256 -Path (Join-Path $RepoRoot (
                        'Specs\Terminal\TerminalEngineCollectionNeutrality.v1.json'))
            }
        }
    }
    return $manifest
}

function Read-RSCurrentCandidateUniverse {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$RepoRoot)
    return Read-RSCurrentCandidateUniverseCore -RepoRoot $RepoRoot
}

function Get-RSCurrentCandidateUniverseFilePath {
    param([Parameter(Mandatory = $true)][string]$RepoRoot)
    return Join-Path $RepoRoot (
        'Specs\Terminal\TerminalEngineCandidateUniverse.v1.json')
}
'@
    [IO.File]::AppendAllText(
        $fixtureEvaluatorPath,
        $fixtureUniverseOverride,
        [Text.UTF8Encoding]::new($false))

    $universePath = Join-Path $fixtureRepo (
        'Specs\Terminal\TerminalEngineCandidateUniverse.v1.json')
    $universe = Get-Content -LiteralPath $universePath -Raw |
        ConvertFrom-Json -AsHashtable -Depth 100 -NoEnumerate -DateKind String
    $universe.authorId = 'fixture-universe-author'
    $universe.frozenUtc = '2026-07-30T17:59:58Z'
    $universe.seedCandidates = [object[]]@(
        [ordered]@{
            adaptationPermission = 'allowed-existing-subsystem'
            allowedNominationId = 'alpha'
            candidateId = 'alpha'
            kittySubsystem = 'present'
            languageFamily = 'cpp'
            pin = 'alpha-pin'
            pinSha256 = Get-RSTestSha256 -Value 'alpha-pin'
            repositoryUrl = 'https://example.invalid/alpha'
            retrievalUrl = 'https://example.invalid/alpha.tar.gz'
            selectionReason = 'Evaluator fixture admitted candidate alpha.'
            sourceArchiveSha256 = Get-RSTestSha256 -Value 'alpha-source'
        },
        [ordered]@{
            adaptationPermission = 'allowed-existing-subsystem'
            allowedNominationId = 'beta'
            candidateId = 'beta'
            kittySubsystem = 'present'
            languageFamily = 'cpp'
            pin = 'beta-pin'
            pinSha256 = Get-RSTestSha256 -Value 'beta-pin'
            repositoryUrl = 'https://example.invalid/beta'
            retrievalUrl = 'https://example.invalid/beta.tar.gz'
            selectionReason = 'Evaluator fixture admitted candidate beta.'
            sourceArchiveSha256 = Get-RSTestSha256 -Value 'beta-source'
        },
        [ordered]@{
            adaptationPermission = 'allowed-existing-subsystem'
            allowedNominationId = 'failed'
            candidateId = 'failed'
            kittySubsystem = 'present'
            languageFamily = 'cpp'
            pin = 'failed-pin'
            pinSha256 = Get-RSTestSha256 -Value 'failed-pin'
            repositoryUrl = 'https://example.invalid/failed'
            retrievalUrl = 'https://example.invalid/failed.tar.gz'
            selectionReason = 'Evaluator fixture admitted candidate that later fails a gate.'
            sourceArchiveSha256 = Get-RSTestSha256 -Value 'failed-source'
        })
    $universe.controls = [object[]]@(
        [ordered]@{
            candidateId = 'web-control'
            constraintId = 'web-runtime-second-stack'
            pin = 'web-control-pin'
            pinSha256 = Get-RSTestSha256 -Value 'web-control-pin'
        })
    $proposal = [ordered]@{}
    foreach ($entry in $universe.GetEnumerator()) {
        if ([string]$entry.Key -cnotin @('proposalSha256', 'approvals')) {
            $proposal[[string]$entry.Key] = $entry.Value
        }
    }
    $universe.proposalSha256 = Get-RSJcsSha256 -InputObject $proposal
    $universe.approvals = [object[]]@(
        [ordered]@{
            approvalId = 'fixture-universe-review-a'
            decision = 'approved'
            proposalSha256 = $universe.proposalSha256
            reviewedUtc = '2026-07-30T17:59:59Z'
            reviewerId = 'fixture-universe-reviewer-a'
            reviewerRole = 'independent-universe-reviewer'
        },
        [ordered]@{
            approvalId = 'fixture-universe-review-b'
            decision = 'approved'
            proposalSha256 = $universe.proposalSha256
            reviewedUtc = '2026-07-30T18:00:00Z'
            reviewerId = 'fixture-universe-reviewer-b'
            reviewerRole = 'independent-universe-reviewer'
        })
    [IO.File]::WriteAllBytes(
        $universePath,
        (ConvertTo-RSJcsUtf8Bytes -InputObject $universe))
    $universeSha = (Get-FileHash `
            -LiteralPath $universePath `
            -Algorithm SHA256).Hash.ToLowerInvariant()

    $candidateReviewPath = Join-Path $fixtureRepo 'Specs\Terminal\TerminalEngineCandidateReview.v1.json'
    $candidateReview = [ordered]@{
        formatId = 'red-salamander-terminal-engine-candidate-review'
        version = 1
        frozenUtc = '2026-07-30T18:00:03Z'
        authorId = 'fixture-author'
        reviewRule = 'No source-rejected candidates in this evaluator fixture.'
        candidates = [object[]]@(
            [ordered]@{
                candidateId = 'alpha'
                pin = 'alpha-pin'
                pinSha256 = Get-RSTestSha256 -Value 'alpha-pin'
                sourceArchiveSha256 = Get-RSTestSha256 -Value 'alpha-source'
                disposition = 'collect'
                glueAuthorIds = @()
                blockers = @()
            },
            [ordered]@{
                candidateId = 'beta'
                pin = 'beta-pin'
                pinSha256 = Get-RSTestSha256 -Value 'beta-pin'
                sourceArchiveSha256 = Get-RSTestSha256 -Value 'beta-source'
                disposition = 'collect'
                glueAuthorIds = @()
                blockers = @()
            },
            [ordered]@{
                candidateId = 'failed'
                pin = 'failed-pin'
                pinSha256 = Get-RSTestSha256 -Value 'failed-pin'
                sourceArchiveSha256 = Get-RSTestSha256 -Value 'failed-source'
                disposition = 'collect'
                glueAuthorIds = @()
                blockers = @()
            })
        approvals = @()
    }
    [IO.File]::WriteAllBytes(
        $candidateReviewPath,
        (ConvertTo-RSJcsUtf8Bytes -InputObject $candidateReview))
    $candidateReviewSha = (Get-FileHash `
            -LiteralPath $candidateReviewPath `
            -Algorithm SHA256).Hash.ToLowerInvariant()
    $candidateSetSeed = Get-RSTestFrozenCandidateSet `
        -CandidateReviewSha256 $candidateReviewSha `
        -CandidateUniverseSha256 $universeSha
    $assessmentCandidates = [object[]]@(
        (New-RSTestAssessmentCandidate `
                -FrozenCandidate $candidateSetSeed.candidates[0] `
                -Profile alpha),
        (New-RSTestAssessmentCandidate `
                -FrozenCandidate $candidateSetSeed.candidates[1] `
                -Profile beta),
        (New-RSTestAssessmentCandidate `
                -FrozenCandidate $candidateSetSeed.candidates[2] `
                -Profile failed))
    $patchDirectory = Join-Path $fixtureRepo (
        'External\TerminalEngine\CandidatePatches')
    [void](New-Item -ItemType Directory -Path $patchDirectory -Force)
    Import-Module $evidenceModulePath -Force -Global
    foreach ($assessmentCandidate in $assessmentCandidates) {
        [IO.File]::WriteAllText(
            (Join-Path $patchDirectory (
                "$($assessmentCandidate.candidateId).patch")),
            "fixture-patch-$($assessmentCandidate.candidateId)",
            [Text.UTF8Encoding]::new($false))
        $mappingPayload = [ordered]@{
            candidateId = [string]$assessmentCandidate.candidateId
            upstreamTree = [string]$assessmentCandidate.rawLedger.patch.upstreamTree
            resultTree = [string]$assessmentCandidate.rawLedger.patch.resultTree
            patchSha256 = [string]$assessmentCandidate.rawLedger.patch.patchSha256
            pathConcerns = [object[]]@(
                [ordered]@{
                    path = 'fixture.txt'
                    concerns = @('abi-snapshot-identity')
                })
        }
        $mapping = [ordered]@{
            formatId = 'red-salamander-terminal-engine-concern-mapping'
            version = 1
            candidateId = [string]$mappingPayload.candidateId
            upstreamTree = [string]$mappingPayload.upstreamTree
            resultTree = [string]$mappingPayload.resultTree
            patchSha256 = [string]$mappingPayload.patchSha256
            pathConcerns = $mappingPayload.pathConcerns
            mappingSha256 = Get-RSJcsSha256 -InputObject $mappingPayload
        }
        [IO.File]::WriteAllBytes(
            (Join-Path $patchDirectory (
                "$($assessmentCandidate.candidateId).concerns.json")),
            (ConvertTo-RSJcsUtf8Bytes -InputObject $mapping))
    }
    $candidateAssessmentPath = Join-Path $fixtureRepo (
        'Specs\Terminal\TerminalEngineCandidateAssessment.v1.json')
    $candidateAssessment = [ordered]@{
        formatId = 'red-salamander-terminal-engine-candidate-assessment'
        version = 1
        frozenUtc = '2026-07-30T18:00:03Z'
        authorId = 'fixture-assessment-author'
        candidateUniverseSha256 = (Get-FileHash `
                -LiteralPath (Join-Path $fixtureRepo (
                    'Specs\Terminal\TerminalEngineCandidateUniverse.v1.json')) `
                -Algorithm SHA256).Hash.ToLowerInvariant()
        scoreAnchorsSha256 = (Get-FileHash `
                -LiteralPath (Join-Path $fixtureRepo (
                    'Specs\Terminal\TerminalEngineScoreAnchors.v1.json')) `
                -Algorithm SHA256).Hash.ToLowerInvariant()
        candidates = $assessmentCandidates
    }
    Import-Module $evidenceModulePath -Force -Global
    [IO.File]::WriteAllBytes(
        $candidateAssessmentPath,
        (ConvertTo-RSJcsUtf8Bytes -InputObject $candidateAssessment))
    $candidateAssessmentSha = (Get-FileHash `
            -LiteralPath $candidateAssessmentPath `
            -Algorithm SHA256).Hash.ToLowerInvariant()
    $candidateSet = Get-RSTestFrozenCandidateSet `
        -CandidateReviewSha256 $candidateReviewSha `
        -CandidateAssessmentSha256 $candidateAssessmentSha `
        -CandidateUniverseSha256 $universeSha
    foreach ($outcome in $candidateSet.nominationOutcomes) {
        $assessmentCandidate = @(
            $assessmentCandidates | Where-Object {
                [string]$_.candidateId -ceq [string]$outcome.nominationId
            })[0]
        $outcome.ledgerSha256 = [string]$assessmentCandidate.ledgerSha256
    }
    $candidateSetPath = Join-Path $fixtureRepo 'Specs\Terminal\TerminalEngineCandidateSet.v1.json'
    [IO.File]::WriteAllBytes(
        $candidateSetPath,
        (ConvertTo-RSJcsUtf8Bytes -InputObject $candidateSet))
    [IO.File]::WriteAllText(
        (Join-Path $fixtureRepo '.gitignore'),
        "/evidence/`n",
        [Text.UTF8Encoding]::new($false))
    $providerPath = Join-Path $fixtureRepo 'Tools\TerminalEngine\Invoke-TerminalEngineCollectionProvider.ps1'
    $null = New-Item -ItemType Directory -Path ([IO.Path]::GetDirectoryName($providerPath)) -Force
    $providerSource = @'
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Phase,
    [Parameter(Mandatory = $true)][string[]]$CandidateId,
    [Parameter(Mandatory = $true)][string]$Platform,
    [Parameter(Mandatory = $true)][string]$Configuration,
    [Parameter(Mandatory = $true)][string]$RepoRoot
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrEmpty($env:RS_TERMINAL_EVALUATOR_TEST_PROVIDER_JSON)) {
    throw 'Test provider result JSON was not supplied.'
}
$envelope = Get-Content -LiteralPath $env:RS_TERMINAL_EVALUATOR_TEST_PROVIDER_JSON -Raw |
    ConvertFrom-Json -AsHashtable -Depth 100 -DateKind String
switch ($Phase) {
    'Bootstrap' { return $envelope.bootstrap }
    'Clean' { return [ordered]@{ clean = $true } }
    'Collect' { return $envelope.collect }
    default { throw "Unsupported test provider phase '$Phase'." }
}
'@
    [IO.File]::WriteAllText($providerPath, $providerSource, [Text.UTF8Encoding]::new($false))

    & git init -q $fixtureRepo
    if ($LASTEXITCODE -ne 0) { throw 'git init failed for evaluator fixture.' }
    & git -C $fixtureRepo config core.longpaths true
    & git -C $fixtureRepo config user.name 'Terminal Evaluator Test'
    & git -C $fixtureRepo config user.email 'terminal-evaluator-test@example.invalid'
    & git -C $fixtureRepo add -- .
    & git -C $fixtureRepo commit -q -m 'frozen terminal evaluator fixture'
    if ($LASTEXITCODE -ne 0) { throw 'git commit failed for evaluator fixture.' }

    Import-Module (Join-Path $fixtureRepo 'Tools\TerminalEvidence.psm1') -Force -Global
    $commit = [string](& git -C $fixtureRepo rev-parse HEAD)
    $machineHash = Get-RSTestMachineHash
    $schemaSha = (Get-FileHash `
            -LiteralPath (Join-Path $fixtureRepo 'Specs\Terminal\TerminalEngineEvidence.schema.json') `
            -Algorithm SHA256).Hash.ToLowerInvariant()
    $harnessSha = Get-RSTerminalHarnessSourceSha256 -RepoRoot $fixtureRepo
    $frozen = [ordered]@{
        requirementsSha256 = (Get-FileHash `
                -LiteralPath (Join-Path $fixtureRepo 'Specs\Terminal\TerminalEngineRequirements.v1.json') `
                -Algorithm SHA256).Hash.ToLowerInvariant()
        scoreAnchorsSha256 = (Get-FileHash `
                -LiteralPath (Join-Path $fixtureRepo 'Specs\Terminal\TerminalEngineScoreAnchors.v1.json') `
                -Algorithm SHA256).Hash.ToLowerInvariant()
        fixtureCorpusSha256 = (Get-FileHash `
                -LiteralPath (Join-Path $fixtureRepo 'Specs\Terminal\TerminalEngineFixtureCorpus.v1.json') `
                -Algorithm SHA256).Hash.ToLowerInvariant()
        measurementProtocolSha256 = (Get-FileHash `
                -LiteralPath (Join-Path $fixtureRepo 'Specs\Terminal\TerminalEngineMeasurementProtocol.v1.json') `
                -Algorithm SHA256).Hash.ToLowerInvariant()
        candidateUniverseSha256 = (Get-FileHash `
                -LiteralPath (Join-Path $fixtureRepo 'Specs\Terminal\TerminalEngineCandidateUniverse.v1.json') `
                -Algorithm SHA256).Hash.ToLowerInvariant()
        candidateSetSha256 = Get-RSTestCandidateSetSha256 -CandidateSet $candidateSet
        candidateReviewSha256 = $candidateReviewSha
        candidateAssessmentSha256 = $candidateAssessmentSha
    }
    $assessmentById = @{}
    foreach ($candidate in $assessmentCandidates) {
        $assessmentById[[string]$candidate.candidateId] = $candidate
    }
    return [pscustomobject]@{
        RepoRoot = $fixtureRepo
        EvidenceRoot = Join-Path $fixtureRepo 'evidence'
        CandidateSet = $candidateSet
        Commit = $commit
        MachineHash = $machineHash
        SchemaSha256 = $schemaSha
        HarnessSha256 = $harnessSha
        Frozen = $frozen
        AssessmentById = $assessmentById
    }
}

function Write-RSTestCollection {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Fixture,

        [Parameter(Mandatory = $true)]
        [string]$Platform,

        [Parameter(Mandatory = $true)]
        [string]$Configuration,

        [Parameter(Mandatory = $true)]
        [string]$RunId,

        [switch]$AllEarlyFailure,

        [switch]$OmitBetaC11,

        [switch]$OmitBetaC7Proof,

        [AllowNull()]
        [string]$BetaC7ProofSha256 = $null,

        [AllowNull()]
        [string]$BetaC7ProofCategory = $null
    )

    $laneKey = "$Platform|$Configuration"
    if ($AllEarlyFailure) {
        $alpha = New-RSTestEarlyFailureCandidate `
            -FrozenCandidate $Fixture.CandidateSet.candidates[0] `
            -LaneKey $laneKey
        $beta = New-RSTestEarlyFailureCandidate `
            -FrozenCandidate $Fixture.CandidateSet.candidates[1] `
            -LaneKey $laneKey
    } else {
        $alpha = New-RSTestCompleteCandidate `
            -FrozenCandidate $Fixture.CandidateSet.candidates[0] `
            -LaneKey $laneKey `
            -AssessmentCandidate $Fixture.AssessmentById['alpha']
        $beta = New-RSTestCompleteCandidate `
            -FrozenCandidate $Fixture.CandidateSet.candidates[1] `
            -LaneKey $laneKey `
            -AssessmentCandidate $Fixture.AssessmentById['beta']
        if ($OmitBetaC11) {
            $beta.evidenceObjects = [object[]]@(
                $beta.evidenceObjects | Where-Object evidenceId -cne 'score.c11'
            )
        }
        if ($OmitBetaC7Proof) {
            $beta.evidenceObjects = [object[]]@(
                $beta.evidenceObjects |
                    Where-Object evidenceId -cne 'score.c7-proof'
            )
        }
        $betaC7Proof = @(
            $beta.evidenceObjects |
                Where-Object evidenceId -ceq 'score.c7-proof'
        )
        if ($PSBoundParameters.ContainsKey('BetaC7ProofSha256')) {
            $betaC7Proof[0].sha256 = $BetaC7ProofSha256
        }
        if ($PSBoundParameters.ContainsKey('BetaC7ProofCategory')) {
            $betaC7Proof[0].resultCategory = $BetaC7ProofCategory
        }
    }
    $failed = New-RSTestEarlyFailureCandidate `
        -FrozenCandidate $Fixture.CandidateSet.candidates[2] `
        -LaneKey $laneKey
    $control = New-RSTestControlCandidate `
        -FrozenCandidate $Fixture.CandidateSet.candidates[3] `
        -LaneKey $laneKey
    $nativeMachine = if ($Platform -ceq 'x64') { 'AMD64' } else { 'ARM64' }
    $bootstrapSha = @(
        $Fixture.CandidateSet.bootstrapInputSets |
            Where-Object platform -CEQ $Platform
    )[0].sha256
    $created = [DateTime]::ParseExact(
        $RunId,
        'yyyy-MM-dd_HHmmss',
        [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::AssumeUniversal)
    $manifest = [ordered]@{
        formatId = 'red-salamander-terminal-engine-evidence'
        schemaVersion = 1
        recordKind = 'collection'
        runId = $RunId
        createdUtc = $created.ToUniversalTime().ToString(
            'yyyy-MM-ddTHH:mm:ssZ',
            [Globalization.CultureInfo]::InvariantCulture)
        repositoryCommit = [string]$Fixture.Commit
        repositoryClean = $true
        machineHash = [string]$Fixture.MachineHash
        schemaSha256 = [string]$Fixture.SchemaSha256
        harness = [ordered]@{
            version = '1.0.0'
            sourceSha256 = [string]$Fixture.HarnessSha256
        }
        frozen = $Fixture.Frozen
        platform = $Platform
        configuration = $Configuration
        nativeAttestation = [ordered]@{
            processMachine = 'UNKNOWN'
            nativeMachine = $nativeMachine
            windowsBuild = '10.0.26200.0'
            isNative = $true
            minimumSupportedBuild = '22000.2600'
            minimumOsProof = 'pe-import-audit'
            minimumOsEvidenceSha256 = Get-RSTestSha256 -Value "$Platform|minimum-os"
        }
        toolchain = [ordered]@{
            compilerId = 'msvc'
            version = '19.50.10000'
            identitySha256 = Get-RSTestSha256 -Value "$Platform|$Configuration|toolchain"
        }
        collectionPolicy = [ordered]@{
            bootstrap = $true
            clean = $true
            offline = $true
            bootstrapInputSetSha256 = [string]$bootstrapSha
            networkIsolation = 'runner-enforced'
        }
        candidates = [object[]]@($alpha, $beta, $failed, $control)
    }
    $directory = Join-Path `
        $Fixture.EvidenceRoot `
        "$($Fixture.MachineHash)\TerminalEngine\$RunId"
    $null = New-Item -ItemType Directory -Path $directory -Force
    Import-Module (Join-Path $Fixture.RepoRoot 'Tools\TerminalEvidence.psm1') -Force -Global
    return Write-RSEvidenceRecord `
        -Manifest $manifest `
        -Directory $directory `
        -JsonLeaf 'terminal-engine-evidence.v1.json' `
        -CompanionLeaf 'terminal-engine-evidence.v1.sha256'
}

function Read-RSTestGovernanceRecord {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return [Text.UTF8Encoding]::new($false, $true).GetString(
        [IO.File]::ReadAllBytes($Path)) |
        ConvertFrom-Json `
            -AsHashtable `
            -Depth 100 `
            -NoEnumerate `
            -DateKind String
}

function Write-RSTestGovernanceRecord {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Value
    )

    Import-Module $evidenceModulePath -Force -Global
    [IO.File]::WriteAllBytes(
        $Path,
        (ConvertTo-RSJcsUtf8Bytes -InputObject $Value))
}

function Edit-RSTestAssessmentAndRebind {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FixtureRoot,

        [Parameter(Mandatory = $true)]
        [scriptblock]$Edit
    )

    $assessmentPath = Join-Path $FixtureRoot (
        'Specs\Terminal\TerminalEngineCandidateAssessment.v1.json')
    $assessment = Read-RSTestGovernanceRecord -Path $assessmentPath
    & $Edit $assessment
    Write-RSTestGovernanceRecord -Path $assessmentPath -Value $assessment
    $assessmentSha256 = (Get-FileHash `
            -LiteralPath $assessmentPath `
            -Algorithm SHA256).Hash.ToLowerInvariant()
    $setPath = Join-Path $FixtureRoot (
        'Specs\Terminal\TerminalEngineCandidateSet.v1.json')
    $candidateSet = Read-RSTestGovernanceRecord -Path $setPath
    $candidateSet['candidateAssessmentSha256'] = $assessmentSha256
    Write-RSTestGovernanceRecord -Path $setPath -Value $candidateSet
}

function Edit-RSTestReviewAndRebind {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FixtureRoot,

        [Parameter(Mandatory = $true)]
        [scriptblock]$Edit
    )

    $reviewPath = Join-Path $FixtureRoot (
        'Specs\Terminal\TerminalEngineCandidateReview.v1.json')
    $review = Read-RSTestGovernanceRecord -Path $reviewPath
    & $Edit $review
    Write-RSTestGovernanceRecord -Path $reviewPath -Value $review
    $reviewSha256 = (Get-FileHash `
            -LiteralPath $reviewPath `
            -Algorithm SHA256).Hash.ToLowerInvariant()
    $setPath = Join-Path $FixtureRoot (
        'Specs\Terminal\TerminalEngineCandidateSet.v1.json')
    $candidateSet = Read-RSTestGovernanceRecord -Path $setPath
    $candidateSet['candidateReviewSha256'] = $reviewSha256
    foreach ($candidate in $candidateSet['candidates']) {
        if ([string]$candidate['disposition'] -ceq 'source-rejected') {
            $candidate['sourceRejection']['candidateReviewSha256'] =
                $reviewSha256
        }
    }
    foreach ($outcome in $candidateSet['nominationOutcomes']) {
        if ([string]$outcome['outcome'] -ceq 'source-rejected') {
            $outcome['sourceRejection']['candidateReviewSha256'] =
                $reviewSha256
        }
    }
    Write-RSTestGovernanceRecord -Path $setPath -Value $candidateSet
}

function Set-RSTestSourceRejectedCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FixtureRoot
    )

    $universePath = Join-Path $FixtureRoot (
        'Specs\Terminal\TerminalEngineCandidateUniverse.v1.json')
    $universe = Read-RSTestGovernanceRecord -Path $universePath
    $seed = @(
        $universe['seedCandidates'] |
            Where-Object { [string]$_['candidateId'] -ceq 'failed' }
    )[0]
    $seed['adaptationPermission'] = 'forbidden-absent-subsystem'
    $seed['allowedNominationId'] = $null
    $seed['kittySubsystem'] = 'absent'
    $proposal = [ordered]@{}
    foreach ($entry in $universe.GetEnumerator()) {
        if ([string]$entry.Key -cnotin @('proposalSha256', 'approvals')) {
            $proposal[[string]$entry.Key] = $entry.Value
        }
    }
    $universe['proposalSha256'] = Get-RSJcsSha256 -InputObject $proposal
    foreach ($approval in $universe['approvals']) {
        $approval['proposalSha256'] = $universe['proposalSha256']
    }
    Write-RSTestGovernanceRecord -Path $universePath -Value $universe
    $universeSha256 = (Get-FileHash `
            -LiteralPath $universePath `
            -Algorithm SHA256).Hash.ToLowerInvariant()

    $reviewPath = Join-Path $FixtureRoot (
        'Specs\Terminal\TerminalEngineCandidateReview.v1.json')
    $review = Read-RSTestGovernanceRecord -Path $reviewPath
    $reviewCandidate = @(
        $review['candidates'] |
            Where-Object { [string]$_['candidateId'] -ceq 'failed' }
    )[0]
    $reviewCandidate['disposition'] = 'source-rejected'
    $reviewCandidate['archiveRelativePath'] = (
        '.build/terminal-engine-spike/bootstrap/codeload/failed.tar.gz')
    $reviewCandidate['glueAuthorIds'] = [object[]]@(
        'fixture-source-glue-author')
    $blocker = [ordered]@{
        blockerId = 'failed-kitty-source'
        gateId = 'G7'
        failureCode = 'kitty-subsystem-absent'
        claim = 'The frozen source lacks the required Kitty subsystem.'
        source = [ordered]@{
            repositoryUrl = 'https://example.invalid/failed'
            relativePath = 'src/terminal.c'
            fileSha256 = Get-RSTestSha256 -Value 'failed-source-file'
            lineStart = 1
            lineEnd = 1
            lineRangeSha256 = Get-RSTestSha256 -Value 'failed-source-line'
            permalink = 'https://example.invalid/failed/src/terminal.c#L1'
        }
    }
    $reviewCandidate['blockers'] = [object[]]@($blocker)
    $blockerSha256 = Get-RSJcsSha256 -InputObject ([ordered]@{
            candidateId = 'failed'
            pin = [string]$reviewCandidate['pin']
            pinSha256 = [string]$reviewCandidate['pinSha256']
            sourceArchiveSha256 = [string]$reviewCandidate[
                'sourceArchiveSha256']
            blocker = $blocker
        })
    $review['approvals'] = [object[]]@(
        [ordered]@{
            approvalId = 'failed-source-review-a'
            reviewerId = 'fixture-source-reviewer-a'
            reviewerRole = 'independent-source-reviewer'
            blockerSha256 = $blockerSha256
            decision = 'approved'
            reviewedUtc = '2026-07-30T18:00:01Z'
        },
        [ordered]@{
            approvalId = 'failed-source-review-b'
            reviewerId = 'fixture-source-reviewer-b'
            reviewerRole = 'independent-source-reviewer'
            blockerSha256 = $blockerSha256
            decision = 'approved'
            reviewedUtc = '2026-07-30T18:00:02Z'
        })
    Write-RSTestGovernanceRecord -Path $reviewPath -Value $review
    $reviewSha256 = (Get-FileHash `
            -LiteralPath $reviewPath `
            -Algorithm SHA256).Hash.ToLowerInvariant()

    $assessmentPath = Join-Path $FixtureRoot (
        'Specs\Terminal\TerminalEngineCandidateAssessment.v1.json')
    $assessment = Read-RSTestGovernanceRecord -Path $assessmentPath
    $assessment['candidateUniverseSha256'] = $universeSha256
    $assessment['candidates'] = [object[]]@(
        $assessment['candidates'] |
            Where-Object { [string]$_['candidateId'] -cne 'failed' })
    Write-RSTestGovernanceRecord -Path $assessmentPath -Value $assessment
    $assessmentSha256 = (Get-FileHash `
            -LiteralPath $assessmentPath `
            -Algorithm SHA256).Hash.ToLowerInvariant()

    $sourceRejection = [ordered]@{
        gateId = 'G7'
        failureCode = 'kitty-subsystem-absent'
        candidateReviewSha256 = $reviewSha256
        blockerId = 'failed-kitty-source'
        blockerSha256 = $blockerSha256
        sourceArchiveSha256 = [string]$reviewCandidate[
            'sourceArchiveSha256']
    }
    $setPath = Join-Path $FixtureRoot (
        'Specs\Terminal\TerminalEngineCandidateSet.v1.json')
    $candidateSet = Read-RSTestGovernanceRecord -Path $setPath
    $candidateSet['candidateUniverseSha256'] = $universeSha256
    $candidateSet['candidateReviewSha256'] = $reviewSha256
    $candidateSet['candidateAssessmentSha256'] = $assessmentSha256
    $frozenCandidate = @(
        $candidateSet['candidates'] |
            Where-Object { [string]$_['candidateId'] -ceq 'failed' }
    )[0]
    $frozenCandidate['disposition'] = 'source-rejected'
    $frozenCandidate['sourceRejection'] = $sourceRejection
    $outcomeIndex = -1
    for ($index = 0;
        $index -lt $candidateSet['nominationOutcomes'].Count;
        $index++) {
        if ([string]$candidateSet['nominationOutcomes'][$index][
                'seedCandidateId'] -ceq 'failed') {
            $outcomeIndex = $index
            break
        }
    }
    $candidateSet['nominationOutcomes'][$outcomeIndex] = [ordered]@{
        seedCandidateId = 'failed'
        pin = [string]$frozenCandidate['pin']
        pinSha256 = [string]$frozenCandidate['pinSha256']
        sourceArchiveSha256 = [string]$frozenCandidate[
            'sourceArchiveSha256']
        outcome = 'source-rejected'
        sourceRejection = $sourceRejection
    }
    Write-RSTestGovernanceRecord -Path $setPath -Value $candidateSet
}

function Set-RSTestAdaptationRejectedCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FixtureRoot
    )

    $reviewPath = Join-Path $FixtureRoot (
        'Specs\Terminal\TerminalEngineCandidateReview.v1.json')
    $review = Read-RSTestGovernanceRecord -Path $reviewPath
    $review['candidates'] = [object[]]@(
        $review['candidates'] |
            Where-Object { [string]$_['candidateId'] -cne 'alpha' })
    Write-RSTestGovernanceRecord -Path $reviewPath -Value $review
    $reviewSha256 = (Get-FileHash `
            -LiteralPath $reviewPath `
            -Algorithm SHA256).Hash.ToLowerInvariant()

    $assessmentPath = Join-Path $FixtureRoot (
        'Specs\Terminal\TerminalEngineCandidateAssessment.v1.json')
    $assessment = Read-RSTestGovernanceRecord -Path $assessmentPath
    $assessmentCandidate = @(
        $assessment['candidates'] |
            Where-Object { [string]$_['candidateId'] -ceq 'alpha' }
    )[0]
    $assessment['candidates'] = [object[]]@(
        $assessment['candidates'] |
            Where-Object { [string]$_['candidateId'] -cne 'alpha' })
    Write-RSTestGovernanceRecord -Path $assessmentPath -Value $assessment
    $assessmentSha256 = (Get-FileHash `
            -LiteralPath $assessmentPath `
            -Algorithm SHA256).Hash.ToLowerInvariant()

    $patch = $assessmentCandidate['rawLedger']['patch']
    $maintenance = $assessmentCandidate['rawLedger']['maintenance']
    $patchPath = Join-Path $FixtureRoot ([string]$patch['patchRelativePath'])
    $rejection = [ordered]@{
        upstreamTree = [string]$patch['upstreamTree']
        patchRelativePath = [string]$patch['patchRelativePath']
        patchSha256 = (Get-FileHash `
                -LiteralPath $patchPath `
                -Algorithm SHA256).Hash.ToLowerInvariant()
        patchSizeBytes = [int64](Get-Item -LiteralPath $patchPath).Length
        resultTree = [string]$patch['resultTree']
        changedMaintainedLines = 2501
        changedMaintainedFiles = 1
        binaryFiles = 0
        logicalConcerns = [object[]]@('abi-snapshot-identity')
        concernMappingRelativePath = [string]$patch[
            'concernMappingRelativePath']
        concernMappingSha256 = [string]$patch['concernMappingSha256']
        newToolchainFamilies = 0
        newNonSystemRuntimeLeaves = 0
        adapterClass = 'same-language-or-existing-c-abi'
        initialAdapterDays = 3
        declaredInitialPatchDays = 10
        initialPatchDays = 10
        perUpgradeDays = 2
        annualSecurityResponseDays = 1
        releaseNoticeDays = 1
        patchTouchesParserDecoderAccounting = $false
        ownerDays = 42
        failureCodes = [object[]]@('changed-lines-cap')
        glueAuthorIds = [object[]]@('fixture-adaptation-glue-author')
    }

    $setPath = Join-Path $FixtureRoot (
        'Specs\Terminal\TerminalEngineCandidateSet.v1.json')
    $candidateSet = Read-RSTestGovernanceRecord -Path $setPath
    $identity = [ordered]@{
        seedCandidateId = 'alpha'
        nominationId = 'alpha'
        pin = [string]$assessmentCandidate['pin']
        pinSha256 = [string]$assessmentCandidate['pinSha256']
        sourceArchiveSha256 = [string]$assessmentCandidate[
            'sourceArchiveSha256']
        adaptationRejection = $rejection
    }
    $evidenceSha256 = Get-RSJcsSha256 -InputObject $identity
    $rejection['evidenceSha256'] = $evidenceSha256
    $rejection['approvals'] = [object[]]@(
        [ordered]@{
            approvalId = 'alpha-adaptation-review-a'
            reviewerId = 'fixture-adaptation-reviewer-a'
            reviewerRole = 'independent-candidate-assessor'
            evidenceSha256 = $evidenceSha256
            decision = 'approved'
            reviewedUtc = '2026-07-30T18:00:01Z'
        },
        [ordered]@{
            approvalId = 'alpha-adaptation-review-b'
            reviewerId = 'fixture-adaptation-reviewer-b'
            reviewerRole = 'independent-candidate-assessor'
            evidenceSha256 = $evidenceSha256
            decision = 'approved'
            reviewedUtc = '2026-07-30T18:00:02Z'
        })
    $candidateSet['candidateReviewSha256'] = $reviewSha256
    $candidateSet['candidateAssessmentSha256'] = $assessmentSha256
    $candidateSet['candidates'] = [object[]]@(
        $candidateSet['candidates'] |
            Where-Object { [string]$_['candidateId'] -cne 'alpha' })
    foreach ($inputSet in $candidateSet['bootstrapInputSets']) {
        $inputs = [object[]]@(
            $inputSet['inputs'] |
                Where-Object { [string]$_['inputId'] -cne 'source-alpha' })
        foreach ($input in $inputs) {
            $input['consumers'] = [object[]]@(
                $input['consumers'] |
                    Where-Object { [string]$_ -cne 'alpha' })
        }
        $inputSet['inputs'] = $inputs
        $inputSet['sha256'] = Get-RSJcsSha256 -InputObject $inputs
    }
    $outcomeIndex = -1
    for ($index = 0;
        $index -lt $candidateSet['nominationOutcomes'].Count;
        $index++) {
        if ([string]$candidateSet['nominationOutcomes'][$index][
                'seedCandidateId'] -ceq 'alpha') {
            $outcomeIndex = $index
            break
        }
    }
    $candidateSet['nominationOutcomes'][$outcomeIndex] = [ordered]@{
        seedCandidateId = 'alpha'
        pin = [string]$assessmentCandidate['pin']
        pinSha256 = [string]$assessmentCandidate['pinSha256']
        sourceArchiveSha256 = [string]$assessmentCandidate[
            'sourceArchiveSha256']
        outcome = 'adaptation-rejected'
        nominationId = 'alpha'
        adaptationRejection = $rejection
    }
    Write-RSTestGovernanceRecord -Path $setPath -Value $candidateSet
}

Describe 'TerminalEngine evaluator invocation contract' {
    It 'accepts only explicit Collect phase switches and the three supported lanes' {
        $valid = [ordered]@{
            Mode = 'Collect'
            Candidate = [string[]]@('All')
            Platform = 'x64'
            Configuration = 'Release'
            Bootstrap = $true
            Clean = $true
            Offline = $true
            EvidenceRoot = 'evidence'
        }
        (Assert-RSTerminalEngineInvocation -BoundParameters $valid).Mode | Should Be 'Collect'

        foreach ($missing in @('Bootstrap', 'Clean', 'Offline')) {
            $invalid = [ordered]@{}
            foreach ($entry in $valid.GetEnumerator()) {
                if ($entry.Key -cne $missing) {
                    $invalid[$entry.Key] = $entry.Value
                }
            }
            Assert-RSThrowsType `
                -Action { Assert-RSTerminalEngineInvocation -BoundParameters $invalid } `
                -ExceptionType ([ArgumentException])
        }

        $falseOffline = [ordered]@{}
        foreach ($entry in $valid.GetEnumerator()) { $falseOffline[$entry.Key] = $entry.Value }
        $falseOffline['Offline'] = $false
        Assert-RSThrowsType `
            -Action { Assert-RSTerminalEngineInvocation -BoundParameters $falseOffline } `
            -ExceptionType ([ArgumentException])

        $partial = [ordered]@{}
        foreach ($entry in $valid.GetEnumerator()) {
            $partial[$entry.Key] = $entry.Value
        }
        $partial['Candidate'] = [string[]]@('alpha')
        Assert-RSThrowsType `
            -Action {
                Assert-RSTerminalEngineInvocation `
                    -BoundParameters $partial
            } `
            -ExceptionType ([ArgumentException]) `
            -MessageFragment 'partial collection manifests'

        $arm64Asan = [ordered]@{}
        foreach ($entry in $valid.GetEnumerator()) { $arm64Asan[$entry.Key] = $entry.Value }
        $arm64Asan['Platform'] = 'ARM64'
        $arm64Asan['Configuration'] = 'ASan Debug'
        Assert-RSThrowsType `
            -Action { Assert-RSTerminalEngineInvocation -BoundParameters $arm64Asan } `
            -ExceptionType ([ArgumentException])
    }

    It 'requires exactly three explicit manifests and Candidate All in Finalize' {
        $valid = [ordered]@{
            Mode = 'Finalize'
            Candidate = [string[]]@('All')
            EvidenceRoot = 'evidence'
            EvidenceManifest = [string[]]@('a', 'b', 'c')
            RequireAllEvidence = $true
        }
        (Assert-RSTerminalEngineInvocation -BoundParameters $valid).Mode | Should Be 'Finalize'

        $tooFew = [ordered]@{}
        foreach ($entry in $valid.GetEnumerator()) { $tooFew[$entry.Key] = $entry.Value }
        $tooFew['EvidenceManifest'] = [string[]]@('a', 'b')
        Assert-RSThrowsType `
            -Action { Assert-RSTerminalEngineInvocation -BoundParameters $tooFew } `
            -ExceptionType ([ArgumentException])

        $narrowed = [ordered]@{}
        foreach ($entry in $valid.GetEnumerator()) { $narrowed[$entry.Key] = $entry.Value }
        $narrowed['Candidate'] = [string[]]@('alpha')
        Assert-RSThrowsType `
            -Action { Assert-RSTerminalEngineInvocation -BoundParameters $narrowed } `
            -ExceptionType ([ArgumentException])

        $collectSwitch = [ordered]@{}
        foreach ($entry in $valid.GetEnumerator()) { $collectSwitch[$entry.Key] = $entry.Value }
        $collectSwitch['Offline'] = $true
        Assert-RSThrowsType `
            -Action { Assert-RSTerminalEngineInvocation -BoundParameters $collectSwitch } `
            -ExceptionType ([ArgumentException])
    }

    It 'returns the distinct parameter exit from a top-level pwsh -File invocation' {
        $pwshPath = (Get-Process -Id $PID).Path
        $output = @(
            & $pwshPath `
                -NoProfile `
                -File (Join-Path $repoRoot 'Tools\Evaluate-TerminalEngines.ps1') `
                -Mode Collect `
                -Candidate All `
                -Platform x64 `
                -Configuration Release `
                -EvidenceRoot (Join-Path $repoRoot '.build\terminal-evaluator-invalid') 2>&1
        )
        $LASTEXITCODE | Should Be 64
        ($output -join "`n") | Should Match 'parameter error'
    }
}

Describe 'TerminalEngine collection semantic validation' {
    BeforeAll {
        $script:semanticCandidateSet = Get-RSTestFrozenCandidateSet
        $script:semanticRequirements = Get-Content `
            -LiteralPath (Join-Path $repoRoot 'Specs\Terminal\TerminalEngineRequirements.v1.json') `
            -Raw |
            ConvertFrom-Json -AsHashtable -Depth 100
        $candidateSetSha = Get-RSTestCandidateSetSha256 -CandidateSet $script:semanticCandidateSet
        $lane = 'x64|Release'
        $script:semanticManifest = [ordered]@{
            recordKind = 'collection'
            platform = 'x64'
            configuration = 'Release'
            nativeAttestation = [ordered]@{
                processMachine = 'UNKNOWN'
                nativeMachine = 'AMD64'
                windowsBuild = '10.0.26200.0'
                isNative = $true
                minimumSupportedBuild = '22000.2600'
                minimumOsProof = 'pe-import-audit'
                minimumOsEvidenceSha256 = Get-RSTestSha256 -Value 'semantic-minimum-os'
            }
            collectionPolicy = [ordered]@{
                bootstrap = $true
                clean = $true
                offline = $true
                bootstrapInputSetSha256 = [string]$script:semanticCandidateSet.bootstrapInputSets[1].sha256
                networkIsolation = 'runner-enforced'
            }
            frozen = [ordered]@{ candidateSetSha256 = $candidateSetSha }
            candidates = [object[]]@(
                (New-RSTestCompleteCandidate `
                        -FrozenCandidate $script:semanticCandidateSet.candidates[0] `
                        -LaneKey $lane),
                (New-RSTestCompleteCandidate `
                        -FrozenCandidate $script:semanticCandidateSet.candidates[1] `
                        -LaneKey $lane),
                (New-RSTestEarlyFailureCandidate `
                        -FrozenCandidate $script:semanticCandidateSet.candidates[2] `
                        -LaneKey $lane),
                (New-RSTestControlCandidate `
                        -FrozenCandidate $script:semanticCandidateSet.candidates[3] `
                        -LaneKey $lane)
            )
        }

        It 'accepts a complete, early-failure, and control set with exact gate transitions' {
            {
                Assert-RSCollectionSemantics `
                    -Manifest $script:semanticManifest `
                    -CandidateSet $script:semanticCandidateSet `
                    -Requirements $script:semanticRequirements `
                    -ExpectedCandidateId ([string[]]@('alpha', 'beta', 'failed', 'web-control'))
            } | Should Not Throw
        }

        It 'rejects any complete candidate without provider ABI proofs' {
            $copy = $script:semanticManifest |
                ConvertTo-Json -Depth 100 |
                ConvertFrom-Json -AsHashtable -Depth 100
            $copy.candidates[1].evidenceObjects = [object[]]@(
                $copy.candidates[1].evidenceObjects |
                    Where-Object evidenceId -cne 'abi.build-identity')

            {
                Assert-RSCollectionSemantics `
                    -Manifest $copy `
                    -CandidateSet $script:semanticCandidateSet `
                    -Requirements $script:semanticRequirements `
                    -ExpectedCandidateId ([string[]]@(
                            'alpha',
                            'beta',
                            'failed',
                            'web-control'))
            } | Should Throw (
                "Complete candidate 'beta' lacks provider-owned 'abi.build-identity' proof.")
        }

        It 'accepts a separately frozen source rejection without a gate sequence' {
            $copy = $script:semanticManifest |
                ConvertTo-Json -Depth 100 |
                ConvertFrom-Json -AsHashtable -Depth 100
            $candidateSet = $script:semanticCandidateSet |
                ConvertTo-Json -Depth 100 |
                ConvertFrom-Json -AsHashtable -Depth 100 -DateKind String
            $sourceRejection = [ordered]@{
                gateId = 'G7'
                failureCode = 'kitty-zlib-missing'
                candidateReviewSha256 = [string]$candidateSet.candidateReviewSha256
                blockerId = 'failed-kitty-zlib'
                blockerSha256 = Get-RSTestSha256 -Value 'failed-kitty-zlib'
                sourceArchiveSha256 = Get-RSTestSha256 -Value 'failed-source'
            }
            $candidateSet.candidates[2].disposition = 'source-rejected'
            $candidateSet.candidates[2].sourceRejection = $sourceRejection
            $copy.frozen.candidateSetSha256 = Get-RSTestCandidateSetSha256 `
                -CandidateSet $candidateSet
            $copy.candidates[2] = [ordered]@{
                candidateId = [string]$candidateSet.candidates[2].candidateId
                pin = [string]$candidateSet.candidates[2].pin
                pinSha256 = [string]$candidateSet.candidates[2].pinSha256
                outcome = 'source-rejected'
                sourceRejection = $sourceRejection
            }

            {
                Assert-RSCollectionSemantics `
                    -Manifest $copy `
                    -CandidateSet $candidateSet `
                    -Requirements $script:semanticRequirements
            } | Should Not Throw
        }

        It 'rejects a plausible static-only collection manifest' {
            $manifest = $script:semanticManifest |
                ConvertTo-Json -Depth 100 |
                ConvertFrom-Json -AsHashtable -Depth 100
            $candidateSet = $script:semanticCandidateSet |
                ConvertTo-Json -Depth 100 |
                ConvertFrom-Json -AsHashtable -Depth 100 -DateKind String
            foreach ($index in 0..2) {
                $frozenCandidate = $candidateSet.candidates[$index]
                $candidateId = [string]$frozenCandidate.candidateId
                $sourceRejection = [ordered]@{
                    gateId = 'G7'
                    failureCode = 'kitty-zlib-missing'
                    candidateReviewSha256 = [string]$candidateSet.candidateReviewSha256
                    blockerId = "$candidateId-kitty-zlib"
                    blockerSha256 = Get-RSTestSha256 -Value "$candidateId-kitty-zlib"
                    sourceArchiveSha256 = [string]$frozenCandidate.sourceArchiveSha256
                }
                $frozenCandidate.disposition = 'source-rejected'
                $frozenCandidate.sourceRejection = $sourceRejection
                $manifest.candidates[$index] = [ordered]@{
                    candidateId = $candidateId
                    pin = [string]$frozenCandidate.pin
                    pinSha256 = [string]$frozenCandidate.pinSha256
                    outcome = 'source-rejected'
                    sourceRejection = $sourceRejection
                }
            }
            $manifest.frozen.candidateSetSha256 =
                Get-RSTestCandidateSetSha256 -CandidateSet $candidateSet

            Assert-RSThrowsType `
                -Action {
                    Assert-RSCollectionSemantics `
                        -Manifest $manifest `
                        -CandidateSet $candidateSet `
                        -Requirements $script:semanticRequirements
                } `
                -ExceptionType ([IO.InvalidDataException]) `
                -MessageFragment 'static-only evidence'
        }

        It 'rejects a gate-specific failure code mismatch' {
            $copy = $script:semanticManifest |
                ConvertTo-Json -Depth 100 |
                ConvertFrom-Json -AsHashtable -Depth 100
            $copy.candidates[2].earlyFailure.failureCode = 'offline-build-failed'
            Assert-RSThrowsType `
                -Action {
                    Assert-RSCollectionSemantics `
                        -Manifest $copy `
                        -CandidateSet $script:semanticCandidateSet `
                        -Requirements $script:semanticRequirements
                } `
                -ExceptionType ([IO.InvalidDataException])
        }

        It 'rejects emulation and an unfrozen bootstrap digest' {
            $emulated = $script:semanticManifest |
                ConvertTo-Json -Depth 100 |
                ConvertFrom-Json -AsHashtable -Depth 100
            $emulated.nativeAttestation.processMachine = 'AMD64'
            Assert-RSThrowsType `
                -Action {
                    Assert-RSCollectionSemantics `
                        -Manifest $emulated `
                        -CandidateSet $script:semanticCandidateSet `
                        -Requirements $script:semanticRequirements
                } `
                -ExceptionType ([IO.InvalidDataException])

            $unfrozen = $script:semanticManifest |
                ConvertTo-Json -Depth 100 |
                ConvertFrom-Json -AsHashtable -Depth 100
            $unfrozen.collectionPolicy.bootstrapInputSetSha256 = ('0' * 64)
            Assert-RSThrowsType `
                -Action {
                    Assert-RSCollectionSemantics `
                        -Manifest $unfrozen `
                        -CandidateSet $script:semanticCandidateSet `
                        -Requirements $script:semanticRequirements
                } `
                -ExceptionType ([IO.InvalidDataException])
        }

        It 'rejects a later gate that does not name the first failure' {
            $copy = $script:semanticManifest |
                ConvertTo-Json -Depth 100 |
                ConvertFrom-Json -AsHashtable -Depth 100
            $copy.candidates[2].gates[8].resultCategory = 'failed-G8'
            Assert-RSThrowsType `
                -Action {
                    Assert-RSCollectionSemantics `
                        -Manifest $copy `
                        -CandidateSet $script:semanticCandidateSet `
                        -Requirements $script:semanticRequirements
                } `
                -ExceptionType ([IO.InvalidDataException])
        }
    }
}

Describe 'TerminalEngine deterministic selection' {
    function New-RSTestScoredResult {
        param(
            [string]$CandidateId,
            [string]$Total,
            [string]$OwnerDays,
            [string]$Abi = '10',
            [string]$Runtime = '1',
            [string]$Bytes = '1000'
        )
        [ordered]@{
            candidateId = $CandidateId
            total = $Total
            tieTuple = [ordered]@{
                ownerDays = $OwnerDays
                crossedAbiElements = $Abi
                nonSystemRuntimeDependencies = $Runtime
                addedReleasePackageBytes = $Bytes
            }
        }
    }

    It 'lets the lower lexicographic tie tuple win inside the inclusive five-point cohort' {
        $selection = New-RSSelection `
            -ScoredResults ([object[]]@(
                    (New-RSTestScoredResult -CandidateId alpha -Total '100.00' -OwnerDays '30'),
                    (New-RSTestScoredResult -CandidateId beta -Total '96.40' -OwnerDays '20'))) `
            -Threshold ([decimal]80) `
            -CohortDelta ([decimal]5)
        $selection.status | Should Be 'winner'
        $selection.rawLeaderScore | Should Be '100.00'
        ($selection.cohort -join ',') | Should Be 'alpha,beta'
        $selection.winnerCandidateId | Should Be 'beta'
    }

    It 'distinguishes no scored candidate, below threshold, and tied first' {
        (New-RSSelection `
                -ScoredResults ([object[]]@()) `
                -Threshold ([decimal]80) `
                -CohortDelta ([decimal]5)).status |
            Should Be 'no-scored-candidate'
        (New-RSSelection `
                -ScoredResults ([object[]]@(
                        (New-RSTestScoredResult -CandidateId alpha -Total '79.99' -OwnerDays '1'))) `
                -Threshold ([decimal]80) `
                -CohortDelta ([decimal]5)).status |
            Should Be 'below-threshold'
        $tied = New-RSSelection `
            -ScoredResults ([object[]]@(
                    (New-RSTestScoredResult -CandidateId alpha -Total '100.00' -OwnerDays '10'),
                    (New-RSTestScoredResult -CandidateId beta -Total '100.00' -OwnerDays '10'))) `
            -Threshold ([decimal]80) `
            -CohortDelta ([decimal]5)
        $tied.status | Should Be 'tied-first-no-winner'
        $null -eq $tied.winnerCandidateId | Should Be $true
    }
}

Describe 'TerminalEngine three-manifest finalization' {
    BeforeEach {
        Mock `
            -CommandName Read-RSCurrentCandidateUniverseCore `
            -ModuleName TerminalEngineEvaluator `
            -MockWith {
                $manifest = Read-RSCandidateUniverse -RepoRoot $RepoRoot
                if ($IncludeRecordIdentity) {
                    $path = Join-Path $RepoRoot (
                        'Specs\Terminal\TerminalEngineCandidateUniverse.v1.json')
                    return [pscustomobject]@{
                        Manifest = $manifest
                        Sha256 = (Get-FileHash `
                                -LiteralPath $path `
                                -Algorithm SHA256).Hash.ToLowerInvariant()
                        FrozenInputSha256ByRelativePath = @{
                            'Specs/Terminal/TerminalEngineCollectionNeutrality.v1.json' =
                                (Get-FileHash `
                                    -LiteralPath (Join-Path $RepoRoot (
                                        'Specs\Terminal\TerminalEngineCollectionNeutrality.v1.json')) `
                                    -Algorithm SHA256).Hash.ToLowerInvariant()
                        }
                    }
                }
                return $manifest
            }
        Mock `
            -CommandName Read-RSCurrentCandidateUniverse `
            -ModuleName TerminalEngineEvaluator `
            -MockWith {
                return Read-RSCandidateUniverse -RepoRoot $RepoRoot
            }
        Mock `
            -CommandName Get-RSCurrentCandidateUniverseFilePath `
            -ModuleName TerminalEngineEvaluator `
            -MockWith {
                return Join-Path $RepoRoot (
                    'Specs\Terminal\TerminalEngineCandidateUniverse.v1.json')
            }
    }

    AfterEach {
        Get-Module -Name TerminalEngineEvaluator -All |
            Remove-Module -Force -ErrorAction Stop
        Import-Module $evaluatorModulePath -Force
        @(Get-Module -Name TerminalEngineEvaluator -All).Count |
            Should Be 1
    }

    It 'collects only through the real three-phase external provider result contract' {
        $fixture = New-RSTestFixtureRepo
        $laneKey = 'x64|Release'
        $providerEnvelope = [ordered]@{
            bootstrap = [ordered]@{
                bootstrap = $true
                bootstrapInputSetSha256 = [string]$fixture.CandidateSet.bootstrapInputSets[1].sha256
            }
            collect = [ordered]@{
                offline = $true
                networkIsolation = 'runner-enforced'
                nativeAttestation = [ordered]@{
                    processMachine = 'UNKNOWN'
                    nativeMachine = 'AMD64'
                    windowsBuild = '10.0.26200.0'
                    isNative = $true
                    minimumSupportedBuild = '22000.2600'
                    minimumOsProof = 'pe-import-audit'
                    minimumOsEvidenceSha256 = Get-RSTestSha256 -Value 'provider-minimum-os'
                }
                toolchain = [ordered]@{
                    compilerId = 'msvc'
                    version = '19.50.10000'
                    identitySha256 = Get-RSTestSha256 -Value 'provider-toolchain'
                }
                candidates = [object[]]@(
                    (New-RSTestCompleteCandidate `
                            -FrozenCandidate $fixture.CandidateSet.candidates[0] `
                            -LaneKey $laneKey),
                    (New-RSTestCompleteCandidate `
                            -FrozenCandidate $fixture.CandidateSet.candidates[1] `
                            -LaneKey $laneKey),
                    (New-RSTestEarlyFailureCandidate `
                            -FrozenCandidate $fixture.CandidateSet.candidates[2] `
                            -LaneKey $laneKey),
                    (New-RSTestControlCandidate `
                            -FrozenCandidate $fixture.CandidateSet.candidates[3] `
                            -LaneKey $laneKey)
                )
            }
        }
        $providerJsonPath = Join-Path $fixture.EvidenceRoot 'provider-result.json'
        $null = New-Item -ItemType Directory -Path $fixture.EvidenceRoot -Force
        [IO.File]::WriteAllText(
            $providerJsonPath,
            ($providerEnvelope | ConvertTo-Json -Depth 100),
            [Text.UTF8Encoding]::new($false))
        $previousProviderJson = $env:RS_TERMINAL_EVALUATOR_TEST_PROVIDER_JSON
        try {
            $env:RS_TERMINAL_EVALUATOR_TEST_PROVIDER_JSON = $providerJsonPath
            $output = @(
                & (Join-Path $fixture.RepoRoot 'Tools\Evaluate-TerminalEngines.ps1') `
                    -Mode Collect `
                    -Candidate All `
                    -Platform x64 `
                    -Configuration Release `
                    -Bootstrap `
                    -Clean `
                    -Offline `
                    -EvidenceRoot $fixture.EvidenceRoot `
                    -PassThruManifest
            )
        } finally {
            $env:RS_TERMINAL_EVALUATOR_TEST_PROVIDER_JSON = $previousProviderJson
        }
        $LASTEXITCODE | Should Be 0
        $output.Count | Should Be 1
        Import-Module $evaluatorModulePath -Force
        $record = Read-RSTerminalEvidenceManifest -Path ([string]$output[0]) -RepoRoot $fixture.RepoRoot
        $record.Manifest.recordKind | Should Be 'collection'
        $record.Manifest.collectionPolicy.networkIsolation | Should Be 'runner-enforced'
        ($record.Manifest.candidates.candidateId -join ',') |
            Should Be 'alpha,beta,failed,web-control'
    }

    It 'revalidates all references, scores, controls, failure records, and selects deterministically' {
        $fixture = New-RSTestFixtureRepo
        $arm64 = Write-RSTestCollection `
            -Fixture $fixture `
            -Platform ARM64 `
            -Configuration Release `
            -RunId '2026-07-30_170000'
        $x64Asan = Write-RSTestCollection `
            -Fixture $fixture `
            -Platform x64 `
            -Configuration 'ASan Debug' `
            -RunId '2026-07-30_170001'
        $x64Release = Write-RSTestCollection `
            -Fixture $fixture `
            -Platform x64 `
            -Configuration Release `
            -RunId '2026-07-30_170002'

        $output = @(
            & (Join-Path $fixture.RepoRoot 'Tools\Evaluate-TerminalEngines.ps1') `
                -Mode Finalize `
                -Candidate All `
                -EvidenceManifest ([string[]]@($x64Release, $arm64, $x64Asan)) `
                -EvidenceRoot $fixture.EvidenceRoot `
                -RequireAllEvidence `
                -PassThruManifest
        )
        $LASTEXITCODE | Should Be 0
        $output.Count | Should Be 1
        $resultManifest = [string]$output[0]

        Import-Module $evaluatorModulePath -Force
        $record = Read-RSTerminalEvidenceManifest `
            -Path $resultManifest `
            -RepoRoot $fixture.RepoRoot
        $manifest = $record.Manifest
        $manifest.selection.rawLeaderScore | Should Be '97.20'
        ($manifest.selection.cohort -join ',') | Should Be 'alpha,beta'
        $manifest.selection.winnerCandidateId | Should Be 'beta'

        $alpha = @($manifest.candidateResults | Where-Object candidateId -CEQ alpha)[0]
        $beta = @($manifest.candidateResults | Where-Object candidateId -CEQ beta)[0]
        $failed = @($manifest.candidateResults | Where-Object candidateId -CEQ failed)[0]
        $alpha.total | Should Be '97.20'
        $beta.total | Should Be '93.60'
        $beta.criteria[0].matchedAnchorId | Should Be 'c1-r4'
        $beta.criteria[0].contribution | Should Be '14.40'
        $beta.criteria[0].evidenceRefs.Count | Should Be 3
        $failed.outcome | Should Be 'mandatory-failed'
        $failed.failureReferences.Count | Should Be 3
        $manifest.controlResults.Count | Should Be 1
        $manifest.controlResults[0].failureReferences.Count | Should Be 3
        ($manifest.inputManifests | ForEach-Object { "$($_.platform)|$($_.configuration)" }) -join ',' |
            Should Be 'ARM64|Release,x64|ASan Debug,x64|Release'
    }

    It 'returns no-scored-candidate for a nonempty executable cohort that fails every lane' {
        $fixture = New-RSTestFixtureRepo
        $arm64 = Write-RSTestCollection `
            -Fixture $fixture `
            -Platform ARM64 `
            -Configuration Release `
            -RunId '2026-07-30_175000' `
            -AllEarlyFailure
        $x64Asan = Write-RSTestCollection `
            -Fixture $fixture `
            -Platform x64 `
            -Configuration 'ASan Debug' `
            -RunId '2026-07-30_175001' `
            -AllEarlyFailure
        $x64Release = Write-RSTestCollection `
            -Fixture $fixture `
            -Platform x64 `
            -Configuration Release `
            -RunId '2026-07-30_175002' `
            -AllEarlyFailure

        # Exercise this synthetic-v1 scoring fixture through its isolated copied
        # evaluator module. Production current-authority readers remain v4-only.
        Import-Module (Join-Path $fixture.RepoRoot (
                'Tools\TerminalEngineEvaluator.psm1')) -Force
        $result = Invoke-RSTerminalEngineFinalization `
            -RepoRoot $fixture.RepoRoot `
            -EvidenceManifest ([string[]]@(
                    $arm64,
                    $x64Asan,
                    $x64Release)) `
            -EvidenceRoot $fixture.EvidenceRoot `
            -RequireAllEvidence
        $result.ExitCode | Should Be 20
        $result.SelectionStatus | Should Be 'no-scored-candidate'

        $record = Read-RSTerminalEvidenceManifest `
            -Path $result.ManifestPath `
            -RepoRoot $fixture.RepoRoot
        $record.Manifest.selection.status | Should Be 'no-scored-candidate'
        @($record.Manifest.candidateResults).Count | Should Be 3
        @(
            $record.Manifest.candidateResults |
                Where-Object outcome -cne 'mandatory-failed'
        ).Count | Should Be 0
    }

    It 'rejects complete evidence without one frozen scoring anchor object per lane' {
        $fixture = New-RSTestFixtureRepo
        $arm64 = Write-RSTestCollection `
            -Fixture $fixture `
            -Platform ARM64 `
            -Configuration Release `
            -RunId '2026-07-30_180000'
        $x64Asan = Write-RSTestCollection `
            -Fixture $fixture `
            -Platform x64 `
            -Configuration 'ASan Debug' `
            -RunId '2026-07-30_180001'
        $x64Release = Write-RSTestCollection `
            -Fixture $fixture `
            -Platform x64 `
            -Configuration Release `
            -RunId '2026-07-30_180002' `
            -OmitBetaC11

        Import-Module $evaluatorModulePath -Force
        Assert-RSThrowsType `
            -Action {
                Invoke-RSTerminalEngineFinalization `
                    -RepoRoot $fixture.RepoRoot `
                    -EvidenceManifest ([string[]]@($arm64, $x64Asan, $x64Release)) `
                    -EvidenceRoot $fixture.EvidenceRoot `
                    -RequireAllEvidence
            } `
            -ExceptionType ([IO.InvalidDataException])
    }

    It 'rejects missing, replayed, or non-verified C7 rating-5 provider proofs' {
        $fixture = New-RSTestFixtureRepo

        $missingArm64 = Write-RSTestCollection `
            -Fixture $fixture `
            -Platform ARM64 `
            -Configuration Release `
            -RunId '2026-07-30_190000'
        $missingAsan = Write-RSTestCollection `
            -Fixture $fixture `
            -Platform x64 `
            -Configuration 'ASan Debug' `
            -RunId '2026-07-30_190001'
        $missingRelease = Write-RSTestCollection `
            -Fixture $fixture `
            -Platform x64 `
            -Configuration Release `
            -RunId '2026-07-30_190002' `
            -OmitBetaC7Proof
        Assert-RSThrowsType `
            -Action {
                Invoke-RSTerminalEngineFinalization `
                    -RepoRoot $fixture.RepoRoot `
                    -EvidenceManifest ([string[]]@(
                            $missingArm64,
                            $missingAsan,
                            $missingRelease)) `
                    -EvidenceRoot $fixture.EvidenceRoot `
                    -RequireAllEvidence
            } `
            -ExceptionType ([IO.InvalidDataException])

        $replayedSha256 = '9' * 64
        $replayArm64 = Write-RSTestCollection `
            -Fixture $fixture `
            -Platform ARM64 `
            -Configuration Release `
            -RunId '2026-07-30_190010' `
            -BetaC7ProofSha256 $replayedSha256
        $replayAsan = Write-RSTestCollection `
            -Fixture $fixture `
            -Platform x64 `
            -Configuration 'ASan Debug' `
            -RunId '2026-07-30_190011' `
            -BetaC7ProofSha256 $replayedSha256
        $replayRelease = Write-RSTestCollection `
            -Fixture $fixture `
            -Platform x64 `
            -Configuration Release `
            -RunId '2026-07-30_190012' `
            -BetaC7ProofSha256 $replayedSha256
        Assert-RSThrowsType `
            -Action {
                Invoke-RSTerminalEngineFinalization `
                    -RepoRoot $fixture.RepoRoot `
                    -EvidenceManifest ([string[]]@(
                            $replayArm64,
                            $replayAsan,
                            $replayRelease)) `
                    -EvidenceRoot $fixture.EvidenceRoot `
                    -RequireAllEvidence
            } `
            -ExceptionType ([IO.InvalidDataException])

        $statusArm64 = Write-RSTestCollection `
            -Fixture $fixture `
            -Platform ARM64 `
            -Configuration Release `
            -RunId '2026-07-30_190020'
        $statusAsan = Write-RSTestCollection `
            -Fixture $fixture `
            -Platform x64 `
            -Configuration 'ASan Debug' `
            -RunId '2026-07-30_190021'
        $statusRelease = Write-RSTestCollection `
            -Fixture $fixture `
            -Platform x64 `
            -Configuration Release `
            -RunId '2026-07-30_190022' `
            -BetaC7ProofCategory claimed
        Assert-RSThrowsType `
            -Action {
                Invoke-RSTerminalEngineFinalization `
                    -RepoRoot $fixture.RepoRoot `
                    -EvidenceManifest ([string[]]@(
                            $statusArm64,
                            $statusAsan,
                            $statusRelease)) `
                    -EvidenceRoot $fixture.EvidenceRoot `
                    -RequireAllEvidence
            } `
            -ExceptionType ([IO.InvalidDataException])
    }
}

Describe 'TerminalEngine candidate-governance chronology closure' {
    BeforeEach {
        Mock `
            -CommandName Read-RSCurrentCandidateUniverseCore `
            -ModuleName TerminalEngineEvaluator `
            -MockWith {
                $manifest = Read-RSCandidateUniverse -RepoRoot $RepoRoot
                if ($IncludeRecordIdentity) {
                    $path = Join-Path $RepoRoot (
                        'Specs\Terminal\TerminalEngineCandidateUniverse.v1.json')
                    return [pscustomobject]@{
                        Manifest = $manifest
                        Sha256 = (Get-FileHash `
                                -LiteralPath $path `
                                -Algorithm SHA256).Hash.ToLowerInvariant()
                        FrozenInputSha256ByRelativePath = @{
                            'Specs/Terminal/TerminalEngineCollectionNeutrality.v1.json' =
                                (Get-FileHash `
                                    -LiteralPath (Join-Path $RepoRoot (
                                        'Specs\Terminal\TerminalEngineCollectionNeutrality.v1.json')) `
                                    -Algorithm SHA256).Hash.ToLowerInvariant()
                        }
                    }
                }
                return $manifest
            }
        Mock `
            -CommandName Read-RSCurrentCandidateUniverse `
            -ModuleName TerminalEngineEvaluator `
            -MockWith {
                return Read-RSCandidateUniverse -RepoRoot $RepoRoot
            }
        Mock `
            -CommandName Get-RSCurrentCandidateUniverseFilePath `
            -ModuleName TerminalEngineEvaluator `
            -MockWith {
                return Join-Path $RepoRoot (
                    'Specs\Terminal\TerminalEngineCandidateUniverse.v1.json')
            }
    }

    It 'uses one exact UTC-second lexical contract in all four live schemas' {
        $expectedPattern =
            '^[0-9]{4}-(?:0[1-9]|1[0-2])-(?:0[1-9]|[12][0-9]|3[01])T(?:[01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]Z$'
        $reviewSchema = Read-RSTestGovernanceRecord -Path (
            Join-Path $repoRoot (
                'Specs\Terminal\TerminalEngineCandidateReview.schema.json'))
        $assessmentSchema = Read-RSTestGovernanceRecord -Path (
            Join-Path $repoRoot (
                'Specs\Terminal\TerminalEngineCandidateAssessment.schema.json'))
        $setSchema = Read-RSTestGovernanceRecord -Path (
            Join-Path $repoRoot (
                'Specs\Terminal\TerminalEngineCandidateSet.schema.json'))
        $v4Schema = Read-RSTestGovernanceRecord -Path (
            Join-Path $repoRoot (
                'Specs\Terminal\TerminalEngineCandidateUniverseTransition.v4.schema.json'))
        $reviewSchema.properties.frozenUtc.pattern | Should Be $expectedPattern
        $reviewSchema.'$defs'.approval.properties.reviewedUtc.pattern |
            Should Be $expectedPattern
        $assessmentSchema.properties.frozenUtc.pattern |
            Should Be $expectedPattern
        $assessmentSchema.'$defs'.approval.properties.reviewedUtc.pattern |
            Should Be $expectedPattern
        $setSchema.properties.frozenUtc.pattern | Should Be $expectedPattern
        $setSchema.'$defs'.adaptationApproval.properties.reviewedUtc.pattern |
            Should Be $expectedPattern
        $v4Schema.'$defs'.utcSecond.pattern | Should Be $expectedPattern
    }

    It 'accepts one stable complete closure and returns all bound identities' {
        $fixture = New-RSTestFixtureRepo
        $closure = Read-RSCandidateGovernanceClosure `
            -RepoRoot $fixture.RepoRoot
        $closure.UniverseOpenUtc | Should Be '2026-07-30T18:00:00Z'
        $closure.GovernanceFreezeUtc | Should Be '2026-07-30T18:00:03Z'
        $closure.UniverseSha256 | Should Be (
            Get-FileHash `
                -LiteralPath (Join-Path $fixture.RepoRoot (
                    'Specs\Terminal\TerminalEngineCandidateUniverse.v1.json')) `
                -Algorithm SHA256).Hash.ToLowerInvariant()
        $closure.CandidateSetSha256 | Should Be (
            Get-FileHash `
                -LiteralPath (Join-Path $fixture.RepoRoot (
                    'Specs\Terminal\TerminalEngineCandidateSet.v1.json')) `
                -Algorithm SHA256).Hash.ToLowerInvariant()
        $null -eq $closure.Review | Should Be $false
        $null -eq $closure.Assessment | Should Be $false
    }

    It 'rejects a valid governance closure swapped after provider context creation' {
        $fixture = New-RSTestFixtureRepo
        $closureA = Read-RSCandidateGovernanceClosure `
            -RepoRoot $fixture.RepoRoot
        Edit-RSTestAssessmentAndRebind `
            -FixtureRoot $fixture.RepoRoot `
            -Edit {
                param($assessment)
                $assessment['authorId'] =
                    'fixture-assessment-author-after-context'
            }

        Assert-RSThrowsType `
            -Action {
                Invoke-RSAuthenticatedCollectionNeutralityProof `
                    -RepoRoot $fixture.RepoRoot `
                    -ExpectedNeutralityPolicySha256 (
                        [string]$closureA.NeutralityPolicySha256) `
                    -ExpectedUniverseSha256 (
                        [string]$closureA.UniverseSha256) `
                    -ExpectedCandidateSetSha256 (
                        [string]$closureA.CandidateSetSha256) `
                    -ExpectedCandidateReviewSha256 (
                        [string]$closureA.CandidateReviewSha256) `
                    -ExpectedCandidateAssessmentSha256 (
                        [string]$closureA.CandidateAssessmentSha256)
            } `
            -ExceptionType ([IO.InvalidDataException]) `
            -MessageFragment (
                'CandidateSet identity differs from the authenticated provider context')
    }

    It 'keeps the public candidate-set reader closure-backed with no skip switch' {
        $command = Get-Command -Name Read-RSCandidateSet
        $command.Parameters.ContainsKey('SkipGovernanceClosure') |
            Should Be $false

        $fixture = New-RSTestFixtureRepo
        $candidateSet = Read-RSCandidateSet -RepoRoot $fixture.RepoRoot
        $candidateSet.formatId |
            Should Be 'red-salamander-terminal-engine-candidate-set'
    }

    It 'binds all four parsed governance records to their locked bytes' {
        $fixture = New-RSTestFixtureRepo
        $module = Get-Module -Name TerminalEngineEvaluator
        $candidateSetResult = & $module {
            param($fixtureRoot)
            Read-RSCandidateSetPreflight `
                -RepoRoot $fixtureRoot `
                -IncludeRecordIdentity
        } $fixture.RepoRoot
        $reviewResult = Read-RSCandidateReview `
            -RepoRoot $fixture.RepoRoot `
            -CandidateSet $candidateSetResult.Manifest
        $assessmentResult = Read-RSCandidateAssessment `
            -RepoRoot $fixture.RepoRoot `
            -CandidateSet $candidateSetResult.Manifest
        $universePath = Join-Path $fixture.RepoRoot (
            'Specs\Terminal\TerminalEngineCandidateUniverse.v1.json')
        $global:RSTestParsedIdentityActive = $true
        $global:RSTestParsedIdentityFixture = [pscustomobject]@{
            Universe = [pscustomobject]@{
                Manifest = Read-RSCandidateUniverse `
                    -RepoRoot $fixture.RepoRoot
                Sha256 = (Get-FileHash `
                        -LiteralPath $universePath `
                        -Algorithm SHA256).Hash.ToLowerInvariant()
                FrozenInputSha256ByRelativePath = @{
                    'Specs/Terminal/TerminalEngineCollectionNeutrality.v1.json' =
                        (Get-FileHash `
                            -LiteralPath (Join-Path $fixture.RepoRoot (
                                'Specs\Terminal\TerminalEngineCollectionNeutrality.v1.json')) `
                            -Algorithm SHA256).Hash.ToLowerInvariant()
                }
            }
            CandidateSet = $candidateSetResult
            CandidateReview = $reviewResult
            CandidateAssessment = $assessmentResult
        }
        Mock `
            -CommandName Read-RSCurrentCandidateUniverseCore `
            -ModuleName TerminalEngineEvaluator `
            -ParameterFilter {
                $global:RSTestParsedIdentityActive
            } `
            -MockWith {
                $identity = $global:RSTestParsedIdentityFixture.Universe
                return [pscustomobject]@{
                    Manifest = $identity.Manifest
                    Sha256 = if (
                        $global:RSTestParsedIdentityMismatchName -ceq
                            'Universe') {
                        'f' * 64
                    } else {
                        $identity.Sha256
                    }
                    FrozenInputSha256ByRelativePath =
                        $identity.FrozenInputSha256ByRelativePath
                }
            }
        Mock `
            -CommandName Read-RSCandidateSetPreflight `
            -ModuleName TerminalEngineEvaluator `
            -ParameterFilter {
                $global:RSTestParsedIdentityActive
            } `
            -MockWith {
                $identity = $global:RSTestParsedIdentityFixture.CandidateSet
                return [pscustomobject]@{
                    Manifest = $identity.Manifest
                    Sha256 = if (
                        $global:RSTestParsedIdentityMismatchName -ceq
                            'CandidateSet') {
                        'f' * 64
                    } else {
                        $identity.Sha256
                    }
                }
            }
        Mock `
            -CommandName Read-RSCandidateReview `
            -ModuleName TerminalEngineEvaluator `
            -ParameterFilter {
                $global:RSTestParsedIdentityActive
            } `
            -MockWith {
                $identity =
                    $global:RSTestParsedIdentityFixture.CandidateReview
                return [pscustomobject]@{
                    Manifest = $identity.Manifest
                    Sha256 = if (
                        $global:RSTestParsedIdentityMismatchName -ceq
                            'CandidateReview') {
                        'f' * 64
                    } else {
                        $identity.Sha256
                    }
                }
            }
        Mock `
            -CommandName Read-RSCandidateAssessment `
            -ModuleName TerminalEngineEvaluator `
            -ParameterFilter {
                $global:RSTestParsedIdentityActive
            } `
            -MockWith {
                $identity =
                    $global:RSTestParsedIdentityFixture.CandidateAssessment
                return [pscustomobject]@{
                    Manifest = $identity.Manifest
                    Sha256 = if (
                        $global:RSTestParsedIdentityMismatchName -ceq
                            'CandidateAssessment') {
                        'f' * 64
                    } else {
                        $identity.Sha256
                    }
                }
            }
        try {
            foreach ($name in @(
                    'Universe',
                    'CandidateSet',
                    'CandidateReview',
                    'CandidateAssessment')) {
                $global:RSTestParsedIdentityMismatchName = $name
                Assert-RSThrowsType `
                    -Action {
                        Read-RSCandidateGovernanceClosure `
                            -RepoRoot $fixture.RepoRoot
                    } `
                    -ExceptionType ([IO.InvalidDataException]) `
                    -MessageFragment (
                        "record '$name' parsed bytes differ from its locked identity")
            }
        } finally {
            $global:RSTestParsedIdentityActive = $false
            Remove-Variable `
                -Name RSTestParsedIdentityFixture `
                -Scope Global `
                -ErrorAction SilentlyContinue
            Remove-Variable `
                -Name RSTestParsedIdentityMismatchName `
                -Scope Global `
                -ErrorAction SilentlyContinue
        }
    }

    It 'rejects every representative static and dynamic input already opened for writing' {
        $fixture = New-RSTestFixtureRepo
        $representativeLeaves = [string[]]@(
            'Specs\Terminal\TerminalEngineCandidateAssessment.v1.json',
            'Specs\Terminal\TerminalEngineCandidateUniverseTransition.v4.schema.json',
            'Specs\Terminal\TerminalEngineRequirements.v1.json',
            'Specs\Terminal\TerminalEngineCandidateUniverse.v2.json',
            'Specs\Terminal\TerminalEngineCandidateUniverse.schema.json',
            'Tools\TerminalEngine\TerminalEngineArchive.psm1',
            'External\TerminalEngine\CandidatePatches\alpha.patch',
            'External\TerminalEngine\CandidatePatches\alpha.concerns.json')
        foreach ($leaf in $representativeLeaves) {
            $writer = [IO.FileStream]::new(
                (Join-Path $fixture.RepoRoot $leaf),
                [IO.FileMode]::Open,
                [IO.FileAccess]::Write,
                [IO.FileShare]::ReadWrite)
            try {
                $caught = $null
                try {
                    Read-RSCandidateGovernanceClosure `
                        -RepoRoot $fixture.RepoRoot
                } catch {
                    $caught = $_.Exception
                }
                $null -eq $caught | Should Be $false
                $caught.Message.Contains(
                    'Cannot acquire immutable read snapshot',
                    [StringComparison]::Ordinal) | Should Be $true
            } finally {
                $writer.Dispose()
            }
        }

        Set-RSTestAdaptationRejectedCandidate `
            -FixtureRoot $fixture.RepoRoot
        foreach ($leaf in [string[]]@(
                'External\TerminalEngine\CandidatePatches\alpha.patch',
                'External\TerminalEngine\CandidatePatches\alpha.concerns.json')) {
            $writer = [IO.FileStream]::new(
                (Join-Path $fixture.RepoRoot $leaf),
                [IO.FileMode]::Open,
                [IO.FileAccess]::Write,
                [IO.FileShare]::ReadWrite)
            try {
                $caught = $null
                try {
                    Read-RSCandidateGovernanceClosure `
                        -RepoRoot $fixture.RepoRoot
                } catch {
                    $caught = $_.Exception
                }
                $null -eq $caught | Should Be $false
                $caught.Message.Contains(
                    'Cannot acquire immutable read snapshot',
                    [StringComparison]::Ordinal) | Should Be $true
            } finally {
                $writer.Dispose()
            }
        }
    }

    It 'rejects an escaping dynamic path without opening the outside file' {
        $fixture = New-RSTestFixtureRepo
        $assessmentPath = Join-Path $fixture.RepoRoot (
            'Specs\Terminal\TerminalEngineCandidateAssessment.v1.json')
        $assessment = Read-RSTestGovernanceRecord -Path $assessmentPath
        $assessment.candidates[0].rawLedger.patch.patchRelativePath =
            '../outside.patch'
        Write-RSTestGovernanceRecord `
            -Path $assessmentPath `
            -Value $assessment
        $outsidePath = Join-Path (
            [IO.Path]::GetDirectoryName($fixture.RepoRoot)) 'outside.patch'
        [IO.File]::WriteAllText(
            $outsidePath,
            'outside',
            [Text.UTF8Encoding]::new($false))
        $exclusive = [IO.FileStream]::new(
            $outsidePath,
            [IO.FileMode]::Open,
            [IO.FileAccess]::ReadWrite,
            [IO.FileShare]::None)
        try {
            Assert-RSThrowsType `
                -Action {
                    Read-RSCandidateGovernanceClosure `
                        -RepoRoot $fixture.RepoRoot
                } `
                -ExceptionType ([IO.InvalidDataException]) `
                -MessageFragment (
                    "CandidateAssessment patch property 'patchRelativePath' has an invalid value")
        } finally {
            $exclusive.Dispose()
        }
    }

    It 'rejects a reparse directory before any governance path is followed' {
        $fixture = New-RSTestFixtureRepo
        $terminalEnginePath = Join-Path $fixture.RepoRoot (
            'Tools\TerminalEngine')
        $junctionTarget = Join-Path (
            [IO.Path]::GetDirectoryName($fixture.RepoRoot)) (
            'terminal-engine-junction-target')
        Move-Item `
            -LiteralPath $terminalEnginePath `
            -Destination $junctionTarget
        $null = New-Item `
            -ItemType Junction `
            -Path $terminalEnginePath `
            -Target $junctionTarget
        Assert-RSThrowsType `
            -Action {
                Read-RSCandidateGovernanceClosure `
                    -RepoRoot $fixture.RepoRoot
            } `
            -ExceptionType ([IO.InvalidDataException]) `
            -MessageFragment 'is a reparse point'
    }

    It 'fails closed before released-handle source-archive verification' {
        $fixture = New-RSTestFixtureRepo
        Mock `
            -CommandName Read-RSCurrentCandidateUniverse `
            -ModuleName TerminalEngineEvaluator `
            -MockWith {
                return [ordered]@{ version = 4 }
            }
        Assert-RSThrowsType `
            -Action {
                Read-RSCandidateReview `
                    -RepoRoot $fixture.RepoRoot `
                    -CandidateSet $fixture.CandidateSet `
                    -VerifySourceArchives
            } `
            -ExceptionType ([IO.InvalidDataException]) `
            -MessageFragment (
                'Round 4 does not authorize released-handle source-archive verification')
    }

    It 'rejects every noncanonical timestamp spelling and invalid calendar instant after rebinding' {
        $fixture = New-RSTestFixtureRepo
        $assessmentPath = Join-Path $fixture.RepoRoot (
            'Specs\Terminal\TerminalEngineCandidateAssessment.v1.json')
        $setPath = Join-Path $fixture.RepoRoot (
            'Specs\Terminal\TerminalEngineCandidateSet.v1.json')
        $originalAssessment = [IO.File]::ReadAllBytes($assessmentPath)
        $originalSet = [IO.File]::ReadAllBytes($setPath)
        $badValues = [string[]]@(
            '2026-07-30T18:00:03+00:00',
            '2026-07-30T18:00:03.000Z',
            '2026-07-30T18:00:03z',
            ' 2026-07-30T18:00:03Z',
            '2026-07-30T18:00:03Z ',
            '2026-02-30T18:00:03Z',
            '2026-07-30T18:00:60Z')

        foreach ($badValue in $badValues) {
            [IO.File]::WriteAllBytes($assessmentPath, $originalAssessment)
            [IO.File]::WriteAllBytes($setPath, $originalSet)
            Edit-RSTestAssessmentAndRebind `
                -FixtureRoot $fixture.RepoRoot `
                -Edit {
                    param($assessment)
                    $assessment['frozenUtc'] = $badValue
                }
            Assert-RSThrowsType `
                -Action {
                    Read-RSCandidateGovernanceClosure `
                        -RepoRoot $fixture.RepoRoot
                } `
                -ExceptionType ([IO.InvalidDataException])
        }

        foreach ($badValue in $badValues) {
            [IO.File]::WriteAllBytes($assessmentPath, $originalAssessment)
            [IO.File]::WriteAllBytes($setPath, $originalSet)
            Edit-RSTestAssessmentAndRebind `
                -FixtureRoot $fixture.RepoRoot `
                -Edit {
                    param($assessment)
                    $assessment['candidates'][0]['approvals'][0][
                        'reviewedUtc'] = $badValue
                }
            Assert-RSThrowsType `
                -Action {
                    Read-RSCandidateGovernanceClosure `
                        -RepoRoot $fixture.RepoRoot
                } `
                -ExceptionType ([IO.InvalidDataException])
        }
    }

    It 'rejects a fully rebound candidate approval at either open boundary' {
        foreach ($boundaryCase in @(
                [pscustomobject]@{
                    Value = '2026-07-30T18:00:00Z'
                    Message = 'strictly after'
                },
                [pscustomobject]@{
                    Value = '2026-07-30T18:00:03Z'
                    Message = 'strictly before'
                })) {
            $fixture = New-RSTestFixtureRepo
            Edit-RSTestAssessmentAndRebind `
                -FixtureRoot $fixture.RepoRoot `
                -Edit {
                    param($assessment)
                    $assessment['candidates'][0]['approvals'][0][
                        'reviewedUtc'] = $boundaryCase.Value
                }
            Assert-RSThrowsType `
                -Action {
                    Read-RSCandidateGovernanceClosure `
                        -RepoRoot $fixture.RepoRoot
                } `
                -ExceptionType ([IO.InvalidDataException]) `
                -MessageFragment $boundaryCase.Message
        }
    }

    It 'rejects fully rebound Review source approvals at both boundaries' {
        foreach ($instant in @(
                '2026-07-30T18:00:00Z',
                '2026-07-30T18:00:03Z')) {
            $fixture = New-RSTestFixtureRepo
            Set-RSTestSourceRejectedCandidate `
                -FixtureRoot $fixture.RepoRoot
            Edit-RSTestReviewAndRebind `
                -FixtureRoot $fixture.RepoRoot `
                -Edit {
                    param($review)
                    $review['approvals'][0]['reviewedUtc'] = $instant
                }
            Assert-RSThrowsType `
                -Action {
                    Read-RSCandidateGovernanceClosure `
                        -RepoRoot $fixture.RepoRoot
                } `
                -ExceptionType ([IO.InvalidDataException])
        }
    }

    It 'rejects fully rebound Set adaptation approvals at both boundaries' {
        foreach ($instant in @(
                '2026-07-30T18:00:00Z',
                '2026-07-30T18:00:03Z')) {
            $fixture = New-RSTestFixtureRepo
            Set-RSTestAdaptationRejectedCandidate `
                -FixtureRoot $fixture.RepoRoot
            $setPath = Join-Path $fixture.RepoRoot (
                'Specs\Terminal\TerminalEngineCandidateSet.v1.json')
            $candidateSet = Read-RSTestGovernanceRecord -Path $setPath
            $candidateSet['nominationOutcomes'][0][
                'adaptationRejection']['approvals'][0]['reviewedUtc'] =
                $instant
            Write-RSTestGovernanceRecord `
                -Path $setPath `
                -Value $candidateSet
            Assert-RSThrowsType `
                -Action {
                    Read-RSCandidateGovernanceClosure `
                        -RepoRoot $fixture.RepoRoot
                } `
                -ExceptionType ([IO.InvalidDataException])
        }
    }

    It 'rejects Review author, glue-author, and duplicate-reviewer approvals' {
        foreach ($case in @(
                [pscustomobject]@{
                    Kind = 'review-author'
                    ReviewerId = 'fixture-author'
                },
                [pscustomobject]@{
                    Kind = 'glue-author'
                    ReviewerId = 'fixture-source-glue-author'
                },
                [pscustomobject]@{
                    Kind = 'duplicate-reviewer'
                    ReviewerId = 'fixture-source-reviewer-a'
                })) {
            $fixture = New-RSTestFixtureRepo
            Set-RSTestSourceRejectedCandidate `
                -FixtureRoot $fixture.RepoRoot
            Edit-RSTestReviewAndRebind `
                -FixtureRoot $fixture.RepoRoot `
                -Edit {
                    param($review)
                    $review['approvals'][1]['reviewerId'] =
                        $case.ReviewerId
                }
            Assert-RSThrowsType `
                -Action {
                    Read-RSCandidateGovernanceClosure `
                        -RepoRoot $fixture.RepoRoot
                } `
                -ExceptionType ([IO.InvalidDataException])
        }
    }

    It 'rejects Set author, glue-author, and duplicate-reviewer adaptation approvals' {
        foreach ($case in @(
                [pscustomobject]@{
                    Kind = 'set-author'
                    ReviewerId = 'fixture-candidate-set-author'
                },
                [pscustomobject]@{
                    Kind = 'glue-author'
                    ReviewerId = 'fixture-adaptation-glue-author'
                },
                [pscustomobject]@{
                    Kind = 'duplicate-reviewer'
                    ReviewerId = 'fixture-adaptation-reviewer-a'
                })) {
            $fixture = New-RSTestFixtureRepo
            Set-RSTestAdaptationRejectedCandidate `
                -FixtureRoot $fixture.RepoRoot
            $setPath = Join-Path $fixture.RepoRoot (
                'Specs\Terminal\TerminalEngineCandidateSet.v1.json')
            $candidateSet = Read-RSTestGovernanceRecord -Path $setPath
            $candidateSet['nominationOutcomes'][0][
                'adaptationRejection']['approvals'][1]['reviewerId'] =
                $case.ReviewerId
            Write-RSTestGovernanceRecord `
                -Path $setPath `
                -Value $candidateSet
            Assert-RSThrowsType `
                -Action {
                    Read-RSCandidateGovernanceClosure `
                        -RepoRoot $fixture.RepoRoot
                } `
                -ExceptionType ([IO.InvalidDataException])
        }
    }

    It 'rejects Assessment author, glue-author, and duplicate-reviewer approvals' {
        foreach ($case in @(
                [pscustomobject]@{
                    Kind = 'assessment-author'
                    ReviewerId = 'fixture-assessment-author'
                },
                [pscustomobject]@{
                    Kind = 'glue-author'
                    ReviewerId = 'fixture-glue'
                },
                [pscustomobject]@{
                    Kind = 'duplicate-reviewer'
                    ReviewerId = 'fixture-reviewer-a'
                })) {
            $fixture = New-RSTestFixtureRepo
            Edit-RSTestAssessmentAndRebind `
                -FixtureRoot $fixture.RepoRoot `
                -Edit {
                    param($assessment)
                    $assessment['candidates'][0]['approvals'][1][
                        'reviewerId'] = $case.ReviewerId
                }
            Assert-RSThrowsType `
                -Action {
                    Read-RSCandidateGovernanceClosure `
                        -RepoRoot $fixture.RepoRoot
                } `
                -ExceptionType ([IO.InvalidDataException])
        }
    }

    It 'rejects independently rebound Review and Assessment freeze mismatches' {
        $reviewFixture = New-RSTestFixtureRepo
        Edit-RSTestReviewAndRebind `
            -FixtureRoot $reviewFixture.RepoRoot `
            -Edit {
                param($review)
                $review['frozenUtc'] = '2026-07-30T18:00:04Z'
            }
        Assert-RSThrowsType `
            -Action {
                Read-RSCandidateGovernanceClosure `
                    -RepoRoot $reviewFixture.RepoRoot
            } `
            -ExceptionType ([IO.InvalidDataException]) `
            -MessageFragment 'one exact governance freeze'

        $assessmentFixture = New-RSTestFixtureRepo
        Edit-RSTestAssessmentAndRebind `
            -FixtureRoot $assessmentFixture.RepoRoot `
            -Edit {
                param($assessment)
                $assessment['frozenUtc'] = '2026-07-30T18:00:04Z'
            }
        Assert-RSThrowsType `
            -Action {
                Read-RSCandidateGovernanceClosure `
                    -RepoRoot $assessmentFixture.RepoRoot
            } `
            -ExceptionType ([IO.InvalidDataException]) `
            -MessageFragment 'one exact governance freeze'
    }

    It 'rejects an empty fully rebound candidate assessment' {
        $fixture = New-RSTestFixtureRepo
        Edit-RSTestAssessmentAndRebind `
            -FixtureRoot $fixture.RepoRoot `
            -Edit {
                param($assessment)
                $assessment['candidates'] = [object[]]@()
            }
        Assert-RSThrowsType `
            -Action {
                Read-RSCandidateGovernanceClosure `
                    -RepoRoot $fixture.RepoRoot
            } `
            -ExceptionType ([IO.InvalidDataException])
    }
}

Describe 'TerminalEngine executable cohort entry invariant' {
    BeforeEach {
        $global:RSTerminalEvaluatorStaticOnlyCandidateSet = [ordered]@{
            candidates = [object[]]@(
                [ordered]@{
                    candidateId = 'static-rejection'
                    pin = 'static-pin'
                    pinSha256 = Get-RSTestSha256 -Value 'static-pin'
                    sourceArchiveSha256 = Get-RSTestSha256 -Value 'static-source'
                    scoredOrControl = 'scored'
                    disposition = 'source-rejected'
                },
                [ordered]@{
                    candidateId = 'static-control'
                    pin = 'control-pin'
                    pinSha256 = Get-RSTestSha256 -Value 'control-pin'
                    scoredOrControl = 'control'
                    constraintId = 'web-runtime-second-stack'
                })
        }
        Mock `
            -CommandName Read-RSCandidateSetPreflight `
            -ModuleName TerminalEngineEvaluator `
            -MockWith {
                return $global:RSTerminalEvaluatorStaticOnlyCandidateSet
            }
    }

    AfterEach {
        Remove-Variable `
            -Name RSTerminalEvaluatorStaticOnlyCandidateSet `
            -Scope Global `
            -ErrorAction SilentlyContinue
    }

    It 'rejects static-only collection before invoking a provider' {
        Assert-RSThrowsType `
            -Action {
                Invoke-RSTerminalEngineCollection `
                    -RepoRoot $repoRoot `
                    -Candidate All `
                    -Platform x64 `
                    -Configuration Release `
                    -EvidenceRoot $TestDrive
            } `
            -ExceptionType ([IO.InvalidDataException]) `
            -MessageFragment 'static-only evidence'
    }

    It 'rejects static-only finalization before accepting hand-authored lanes' {
        Assert-RSThrowsType `
            -Action {
                Invoke-RSTerminalEngineFinalization `
                    -RepoRoot $repoRoot `
                    -EvidenceManifest ([string[]]@(
                            'plausible-arm64.json',
                            'plausible-x64-asan.json',
                            'plausible-x64-release.json')) `
                    -EvidenceRoot $TestDrive `
                    -RequireAllEvidence
            } `
            -ExceptionType ([IO.InvalidDataException]) `
            -MessageFragment 'static-only evidence'
    }

    It 'rejects every static subset before review, assessment, or provider phases' {
        $global:RSTerminalEvaluatorStaticOnlyCandidateSet = [ordered]@{
            candidates = [object[]]@(
                [ordered]@{
                    candidateId = 'engine-alpha'
                    scoredOrControl = 'scored'
                    disposition = 'collect'
                },
                [ordered]@{
                    candidateId = 'source-static'
                    scoredOrControl = 'scored'
                    disposition = 'source-rejected'
                },
                [ordered]@{
                    candidateId = 'static-control'
                    scoredOrControl = 'control'
                })
        }
        Mock `
            -CommandName Read-RSCandidateReview `
            -ModuleName TerminalEngineEvaluator `
            -MockWith { throw 'candidate review must not be read' }
        Mock `
            -CommandName Read-RSCandidateAssessment `
            -ModuleName TerminalEngineEvaluator `
            -MockWith { throw 'candidate assessment must not be read' }
        Mock `
            -CommandName Invoke-RSCollectionProviderPhase `
            -ModuleName TerminalEngineEvaluator `
            -MockWith { throw 'provider phase must not run' }
        $cases = [object[]]@(
            [pscustomobject]@{
                CandidateId = [string[]]@('source-static')
            },
            [pscustomobject]@{
                CandidateId = [string[]]@('static-control')
            },
            [pscustomobject]@{
                CandidateId =
                    [string[]]@('source-static', 'static-control')
            })
        foreach ($case in $cases) {
            Assert-RSThrowsType `
                -Action {
                    Invoke-RSTerminalEngineCollection `
                        -RepoRoot $repoRoot `
                        -Candidate $case.CandidateId `
                        -Platform x64 `
                        -Configuration Release `
                        -EvidenceRoot $TestDrive
                } `
                -ExceptionType ([IO.InvalidDataException]) `
                -MessageFragment 'selected candidate cohort'
        }
        Assert-MockCalled `
            -CommandName Read-RSCandidateReview `
            -ModuleName TerminalEngineEvaluator `
            -Times 0
        Assert-MockCalled `
            -CommandName Read-RSCandidateAssessment `
            -ModuleName TerminalEngineEvaluator `
            -Times 0
        Assert-MockCalled `
            -CommandName Invoke-RSCollectionProviderPhase `
            -ModuleName TerminalEngineEvaluator `
            -Times 0
    }

    It 'rejects one executable from a multi-executable cohort before governance or effects' {
        $global:RSTerminalEvaluatorStaticOnlyCandidateSet = [ordered]@{
            candidates = [object[]]@(
                [ordered]@{
                    candidateId = 'engine-alpha'
                    scoredOrControl = 'scored'
                    disposition = 'collect'
                },
                [ordered]@{
                    candidateId = 'engine-beta'
                    scoredOrControl = 'scored'
                    disposition = 'collect'
                })
        }
        Mock `
            -CommandName Read-RSCandidateGovernanceClosure `
            -ModuleName TerminalEngineEvaluator `
            -MockWith { throw 'governance closure must not run' }
        Mock `
            -CommandName Read-RSCandidateReview `
            -ModuleName TerminalEngineEvaluator `
            -MockWith { throw 'candidate review must not run' }
        Mock `
            -CommandName Read-RSCandidateAssessment `
            -ModuleName TerminalEngineEvaluator `
            -MockWith { throw 'candidate assessment must not run' }
        Mock `
            -CommandName Get-RSRepositoryState `
            -ModuleName TerminalEngineEvaluator `
            -MockWith { throw 'repository inspection must not run' }
        Mock `
            -CommandName Invoke-RSCollectionProviderPhase `
            -ModuleName TerminalEngineEvaluator `
            -MockWith { throw 'provider phase must not run' }
        $evidenceRoot = Join-Path $TestDrive 'partial-executable-evidence'

        Assert-RSThrowsType `
            -Action {
                Invoke-RSTerminalEngineCollection `
                    -RepoRoot $repoRoot `
                    -Candidate engine-alpha `
                    -Platform x64 `
                    -Configuration Release `
                    -EvidenceRoot $evidenceRoot
            } `
            -ExceptionType ([ArgumentException]) `
            -MessageFragment 'Candidate All'
        Assert-MockCalled `
            -CommandName Read-RSCandidateGovernanceClosure `
            -ModuleName TerminalEngineEvaluator `
            -Times 0 `
            -Scope It
        Assert-MockCalled `
            -CommandName Read-RSCandidateReview `
            -ModuleName TerminalEngineEvaluator `
            -Times 0 `
            -Scope It
        Assert-MockCalled `
            -CommandName Read-RSCandidateAssessment `
            -ModuleName TerminalEngineEvaluator `
            -Times 0 `
            -Scope It
        Assert-MockCalled `
            -CommandName Get-RSRepositoryState `
            -ModuleName TerminalEngineEvaluator `
            -Times 0 `
            -Scope It
        Assert-MockCalled `
            -CommandName Invoke-RSCollectionProviderPhase `
            -ModuleName TerminalEngineEvaluator `
            -Times 0 `
            -Scope It
        [IO.Directory]::Exists($evidenceRoot) | Should Be $false
    }
}

Describe 'TerminalEngine current-universe authority downgrade invariant' {
    BeforeEach {
        Mock `
            -CommandName Read-RSCandidateGovernanceClosure `
            -ModuleName TerminalEngineEvaluator `
            -MockWith { throw 'governance closure must not run' }
        Mock `
            -CommandName Read-RSCandidateReview `
            -ModuleName TerminalEngineEvaluator `
            -MockWith { throw 'candidate review must not run' }
        Mock `
            -CommandName Read-RSCandidateAssessment `
            -ModuleName TerminalEngineEvaluator `
            -MockWith { throw 'candidate assessment must not run' }
        Mock `
            -CommandName Get-RSRepositoryState `
            -ModuleName TerminalEngineEvaluator `
            -MockWith { throw 'repository inspection must not run' }
        Mock `
            -CommandName Invoke-RSCollectionProviderPhase `
            -ModuleName TerminalEngineEvaluator `
            -MockWith { throw 'provider phase must not run' }
    }

    It 'rejects Collect when canonical round 4 is missing before effects' {
        $fixture = New-RSTestFixtureRepo
        foreach ($leaf in @(
                'Specs\Terminal\TerminalEngineCandidateUniverse.v2.json',
                'Specs\Terminal\TerminalEngineCandidateUniverse.v3.json')) {
            Copy-Item `
                -LiteralPath (Join-Path $repoRoot $leaf) `
                -Destination (Join-Path $fixture.RepoRoot $leaf)
        }
        $evidenceRoot = Join-Path $fixture.RepoRoot (
            'missing-v4-collect-evidence')
        Assert-RSThrowsType `
            -Action {
                Invoke-RSTerminalEngineCollection `
                    -RepoRoot $fixture.RepoRoot `
                    -Candidate All `
                    -Platform x64 `
                    -Configuration Release `
                    -EvidenceRoot $evidenceRoot
            } `
            -ExceptionType ([IO.InvalidDataException]) `
            -MessageFragment 'round-4 candidate universe is missing'
        Assert-MockCalled `
            -CommandName Read-RSCandidateGovernanceClosure `
            -ModuleName TerminalEngineEvaluator `
            -Times 0 `
            -Scope It
        Assert-MockCalled `
            -CommandName Read-RSCandidateReview `
            -ModuleName TerminalEngineEvaluator `
            -Times 0 `
            -Scope It
        Assert-MockCalled `
            -CommandName Read-RSCandidateAssessment `
            -ModuleName TerminalEngineEvaluator `
            -Times 0 `
            -Scope It
        Assert-MockCalled `
            -CommandName Get-RSRepositoryState `
            -ModuleName TerminalEngineEvaluator `
            -Times 0 `
            -Scope It
        Assert-MockCalled `
            -CommandName Invoke-RSCollectionProviderPhase `
            -ModuleName TerminalEngineEvaluator `
            -Times 0 `
            -Scope It
        [IO.Directory]::Exists($evidenceRoot) | Should Be $false
    }

    It 'rejects Finalize when canonical round 4 is missing before effects' {
        $fixture = New-RSTestFixtureRepo
        foreach ($leaf in @(
                'Specs\Terminal\TerminalEngineCandidateUniverse.v2.json',
                'Specs\Terminal\TerminalEngineCandidateUniverse.v3.json')) {
            Copy-Item `
                -LiteralPath (Join-Path $repoRoot $leaf) `
                -Destination (Join-Path $fixture.RepoRoot $leaf)
        }
        $evidenceRoot = Join-Path $fixture.RepoRoot (
            'missing-v4-finalize-evidence')
        Assert-RSThrowsType `
            -Action {
                Invoke-RSTerminalEngineFinalization `
                    -RepoRoot $fixture.RepoRoot `
                    -EvidenceManifest ([string[]]@(
                            'plausible-arm64.json',
                            'plausible-x64-asan.json',
                            'plausible-x64-release.json')) `
                    -EvidenceRoot $evidenceRoot `
                    -RequireAllEvidence
            } `
            -ExceptionType ([IO.InvalidDataException]) `
            -MessageFragment 'round-4 candidate universe is missing'
        Assert-MockCalled `
            -CommandName Read-RSCandidateGovernanceClosure `
            -ModuleName TerminalEngineEvaluator `
            -Times 0 `
            -Scope It
        Assert-MockCalled `
            -CommandName Read-RSCandidateReview `
            -ModuleName TerminalEngineEvaluator `
            -Times 0 `
            -Scope It
        Assert-MockCalled `
            -CommandName Read-RSCandidateAssessment `
            -ModuleName TerminalEngineEvaluator `
            -Times 0 `
            -Scope It
        Assert-MockCalled `
            -CommandName Get-RSRepositoryState `
            -ModuleName TerminalEngineEvaluator `
            -Times 0 `
            -Scope It
        [IO.Directory]::Exists($evidenceRoot) | Should Be $false
    }
}

Describe 'TerminalEngine exact frozen-module caller scope' {
    It 'does not add global exports for a previously private module path' {
        $privateModulePath = Join-Path $TestDrive 'RSTestPrivateExactModule.psm1'
        [IO.File]::WriteAllText(
            $privateModulePath,
            @'
function Invoke-RSTestPrivateExactModule {
    return 'authenticated-private-module'
}
Export-ModuleMember -Function Invoke-RSTestPrivateExactModule
'@,
            [Text.UTF8Encoding]::new($false))

        $exactModule = & $evaluatorModule {
            param($path)
            Import-RSExactModule -Path $path
        } $privateModulePath

        [IO.Path]::GetFullPath([string]$exactModule.Path) |
            Should Be ([IO.Path]::GetFullPath($privateModulePath))
        (& $exactModule { Invoke-RSTestPrivateExactModule }) |
            Should Be 'authenticated-private-module'
        $globalExecutionContext = (
            Get-Variable -Name ExecutionContext -Scope Global -ValueOnly)
        $globalCommands = [object[]]@(
            $globalExecutionContext.SessionState.InvokeCommand.GetCommands(
                'Invoke-RSTestPrivateExactModule',
                [Management.Automation.CommandTypes]::All,
                $false))
        $globalCommands.Count | Should Be 0
    }

    It 'preserves authenticated caller-global exports across a current-v4 read' -Skip:(
        Test-Path -LiteralPath (Join-Path $repoRoot 'Plugins\Terminal')) {
        $neutralityModulePath = Join-Path $repoRoot (
            'Tools\TerminalEngine\TerminalEngineCollectionNeutrality.psm1')
        Import-Module $evidenceModulePath -Force -Global
        Import-Module $neutralityModulePath -Force -Global
        $evidenceBefore = (
            Get-Command ConvertTo-RSJcsUtf8Bytes -ErrorAction Stop).Module
        $neutralityBefore = (
            Get-Command Get-RSTerminalEngineCollectionNeutralitySurface `
                -ErrorAction Stop).Module

        $currentUniverse = Read-RSCurrentCandidateUniverse -RepoRoot $repoRoot

        $currentUniverse.version | Should Be 4
        $evidenceAfter = (
            Get-Command ConvertTo-RSJcsUtf8Bytes -ErrorAction Stop).Module
        $neutralityAfter = (
            Get-Command Get-RSTerminalEngineCollectionNeutralitySurface `
                -ErrorAction Stop).Module
        [IO.Path]::GetFullPath([string]$evidenceAfter.Path) |
            Should Be ([IO.Path]::GetFullPath($evidenceModulePath))
        [IO.Path]::GetFullPath([string]$neutralityAfter.Path) |
            Should Be ([IO.Path]::GetFullPath($neutralityModulePath))
        [object]::ReferenceEquals($evidenceBefore, $evidenceAfter) |
            Should Be $false
        [object]::ReferenceEquals($neutralityBefore, $neutralityAfter) |
            Should Be $false
        @(Get-RSTerminalEngineCollectionNeutralitySurface).Count |
            Should Be 51
    }
}
