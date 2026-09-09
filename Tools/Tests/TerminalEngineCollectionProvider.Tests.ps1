Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:RepoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $script:RepoRoot 'Tools\TestRunPlan.ps1')
Import-Module (
    Join-Path $script:RepoRoot 'Tools\TerminalEvidence.psm1') -Force -Global
Import-Module (Join-Path $script:RepoRoot (
        'Tools\TerminalEngine\TerminalEngineCollectionNeutrality.psm1')) `
    -Force -Global
. (Join-Path $script:RepoRoot (
        'Tools\TerminalEngine\Invoke-TerminalEngineCollectionProvider.ps1')) `
    -RepoRoot $script:RepoRoot `
    -LoadOnly

function New-RSProviderTestScratch {
    return New-RSTestSandboxScratchDirectory `
        -RepoRoot $script:RepoRoot `
        -Harness 'terminal-engine-provider-pester' `
        -Case ([Guid]::NewGuid().ToString('N'))
}

function New-RSProviderTestDescriptorValue {
    param(
        [ValidateSet(
            'provider-generic-cmake-msvc-v1',
            'candidate-wrapper-v1')]
        [string]$BuildKind = 'candidate-wrapper-v1',

        [string]$CandidateId = 'engine-alpha',

        [object[]]$HeaderPackages = @(),

        [object[]]$LegalInputs = @(),

        [object[]]$PathBindings = @(),

        [object[]]$PrivateRuntimeLibraries = @()
    )

    if ($LegalInputs.Count -eq 0) {
        $LegalInputs = [object[]]@(
            [ordered]@{
                inputId = 'license-alpha'
                kind = 'license'
                destinationLeaf = 'license-alpha.txt'
                sha256 = '1' * 64
                sizeBytes = '1'
                source = [ordered]@{
                    sourceKind = 'candidate-source'
                    origin = 'upstream'
                    sourceRelativePath = 'license-alpha.txt'
                }
            })
    }
    $build = if ($BuildKind -ceq 'candidate-wrapper-v1') {
        [ordered]@{
            kind = $BuildKind
            driverRelativePath = 'tools/build-driver.ps1'
            pathBindings = [object[]]$PathBindings
            adapter = [ordered]@{
                dllLeaf = 'adapter-alpha.dll'
            }
            privateRuntimeLibraries =
                [object[]]$PrivateRuntimeLibraries
        }
    }
    else {
        [ordered]@{
            kind = $BuildKind
            buildTargets = [string[]]@('adapter-alpha')
            pathBindings = [object[]]$PathBindings
            adapter = [ordered]@{
                dllLeaf = 'adapter-alpha.dll'
                projectLeaf = 'adapter-alpha.vcxproj'
            }
            staticLibraries = [object[]]@()
            privateRuntimeLibraries =
                [object[]]$PrivateRuntimeLibraries
        }
    }
    return [ordered]@{
        formatId =
            'red-salamander-terminal-engine-collection-descriptor'
        version = 1
        candidateId = $CandidateId
        headerPackages = [object[]]$HeaderPackages
        legal = [ordered]@{
            closureId = 'closure-alpha'
            outputRelativePath = 'generated/notices.txt'
            dependencies = [object[]]@(
                [ordered]@{
                    dependencyId = 'dependency-alpha'
                    spdx = 'MIT'
                    inputs = [object[]]$LegalInputs
                })
        }
        build = $build
    }
}

function New-RSProviderTestDescriptor {
    param(
        [ValidateSet(
            'provider-generic-cmake-msvc-v1',
            'candidate-wrapper-v1')]
        [string]$BuildKind = 'candidate-wrapper-v1',

        [object[]]$HeaderPackages = @(),

        [object[]]$LegalInputs = @(),

        [object[]]$PathBindings = @(),

        [object[]]$PrivateRuntimeLibraries = @()
    )

    return [pscustomobject]@{
        Value = New-RSProviderTestDescriptorValue @PSBoundParameters
        Sha256 = '5' * 64
        BuildKind = $BuildKind
        MaintainedBuildWrapper =
            $BuildKind -ceq 'candidate-wrapper-v1'
    }
}

function New-RSProviderTestCandidate {
    return [ordered]@{
        candidateId = 'engine-alpha'
        pin = 'pin-alpha'
        pinSha256 = '1' * 64
        sourceArchiveSha256 = '2' * 64
        scoredOrControl = 'scored'
        disposition = 'collect'
    }
}

function New-RSProviderTestContext {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Scratch,

        [object[]]$Inputs = @()
    )

    $inputById =
        [Collections.Generic.Dictionary[string, object]]::new(
            [StringComparer]::Ordinal)
    foreach ($input in $Inputs) {
        $inputById.Add([string]$input['inputId'], $input)
    }
    $candidate = New-RSProviderTestCandidate
    $assessmentCandidateById =
        [Collections.Generic.Dictionary[string, object]]::new(
            [StringComparer]::Ordinal)
    $assessmentCandidateById.Add(
        'engine-alpha',
        [ordered]@{
            candidateId = 'engine-alpha'
            rawLedger = [ordered]@{
                patch = [ordered]@{ marker = 'authenticated-closure-patch' }
            }
        })
    return [pscustomobject]@{
        RepoRoot = $script:RepoRoot
        Platform = 'x64'
        Configuration = 'Release'
        Inputs = [object[]]$Inputs
        InputById = $inputById
        SelectedIds = [string[]]@('engine-alpha')
        CandidateById = @{}
        SelectedCandidates = [object[]]@($candidate)
        CandidateSet = [ordered]@{
            candidates = [object[]]@($candidate)
        }
        GovernanceClosure = [pscustomobject]@{
            AssessmentResult = [pscustomobject]@{
                CandidateById = $assessmentCandidateById
            }
        }
        NeutralityPolicy = [ordered]@{
            formatId =
                'red-salamander-terminal-engine-collection-neutrality'
            version = 1
        }
        NeutralityPolicySha256 = '6' * 64
        CandidateUniverseSha256 = 'a' * 64
        CandidateSetSha256 = 'b' * 64
        CandidateReviewSha256 = 'c' * 64
        CandidateAssessmentSha256 = 'd' * 64
        BootstrapInputSetSha256 = '7' * 64
        NeutralityBaselineProof = [ordered]@{
            algorithmId = 'terminal-engine-collection-neutrality-v1'
            surfaceSetSha256 = '1' * 64
            forbiddenLiteralSetSha256 = '2' * 64
            exceptionSetSha256 = '3' * 64
            findingCount = 0
            status = 'verified'
            proofSha256 = '4' * 64
        }
        NeutralityProof = [ordered]@{
            algorithmId = 'terminal-engine-collection-neutrality-v1'
            surfaceSetSha256 = '1' * 64
            forbiddenLiteralSetSha256 = '2' * 64
            exceptionSetSha256 = '3' * 64
            findingCount = 0
            status = 'verified'
            proofSha256 = '4' * 64
        }
        TestScratch = $Scratch
    }
}

function New-RSProviderPassingIsolation {
    return [pscustomobject]@{
        providerStatus = 'pass'
        isolationVerified = $true
        rootCommandExecuted = $true
        enforcement = 'runner-enforced'
        ruleSetupVerified = $true
        ruleRemovalVerified = $true
        programIdentityStable = $true
        scopeContract = [pscustomobject]@{
            exclusiveDisposableRunnerAcknowledged = $true
        }
        rootCommand = [pscustomobject]@{
            result = [pscustomobject]@{
                exitCode = 0
                timedOut = $false
                outputComplete = $true
            }
        }
        failures = [object[]]@()
    }
}

function Write-RSProviderTestJcs {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Value
    )

    [void][IO.Directory]::CreateDirectory(
        [IO.Path]::GetDirectoryName($Path))
    [IO.File]::WriteAllBytes(
        $Path,
        (ConvertTo-RSJcsUtf8Bytes -InputObject $Value))
}

function Test-RSProviderThrows {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Action
    )

    try {
        $null = & $Action
        return $false
    }
    catch {
        return $true
    }
}

function Get-RSProviderExceptionMessage {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Action
    )

    try {
        $null = & $Action
    }
    catch {
        return $_.Exception.Message
    }
    throw 'Expected the provider action to throw.'
}

function New-RSProviderSyntheticRepo {
    $scratch = New-RSProviderTestScratch
    $root = Join-Path $scratch 'repo'
    foreach ($relative in @(
            'Tools\TerminalEvidence.psm1',
            'Specs\Terminal\TerminalEngineCandidateUniverse.v4.json',
            'Specs\Terminal\TerminalEngineCandidateSet.v1.json',
            'Specs\Terminal\TerminalEngineCollectionNeutrality.v1.json')) {
        [void][IO.Directory]::CreateDirectory(
            [IO.Path]::GetDirectoryName((Join-Path $root $relative)))
    }
    Copy-Item `
        -LiteralPath (Join-Path `
            $script:RepoRoot `
            'Tools\TerminalEvidence.psm1') `
        -Destination (Join-Path $root 'Tools\TerminalEvidence.psm1')
    $universe = [ordered]@{
        formatId = 'synthetic-candidate-universe'
        version = 4
        candidates = [object[]]@()
    }
    $universePath = Join-Path $root (
        'Specs\Terminal\TerminalEngineCandidateUniverse.v4.json')
    Write-RSProviderTestJcs -Path $universePath -Value $universe
    $inputs = [object[]]@(
        [ordered]@{
            inputId = 'source-alpha'
            inputKind = 'source-archive'
            version = '1'
            fileName = 'source-alpha.tar.gz'
            sha256 = '8' * 64
            retrievalUrl = 'https://example.invalid/source-alpha.tar.gz'
            consumers = [string[]]@('engine-alpha')
        })
    $candidateSet = [ordered]@{
        formatId = 'red-salamander-terminal-engine-candidate-set'
        version = 1
        candidateUniverseSha256 =
            (Get-FileHash `
                -LiteralPath $universePath `
                -Algorithm SHA256).Hash.ToLowerInvariant()
        candidateAssessmentSha256 = '9' * 64
        candidateReviewSha256 = 'a' * 64
        bootstrapInputSets = [object[]]@(
            [ordered]@{
                platform = 'x64'
                inputs = $inputs
                sha256 = Get-RSJcsSha256 -InputObject $inputs
            })
        candidates = [object[]]@(
            [ordered]@{
                candidateId = 'engine-alpha'
                pin = 'engine-pin'
                pinSha256 = 'b' * 64
                scoredOrControl = 'scored'
                disposition = 'collect'
            })
    }
    Write-RSProviderTestJcs `
        -Path (Join-Path $root (
            'Specs\Terminal\TerminalEngineCandidateSet.v1.json')) `
        -Value $candidateSet
    $neutralityPolicyPath = Join-Path $root (
        'Specs\Terminal\TerminalEngineCollectionNeutrality.v1.json')
    $neutralityPolicy = [ordered]@{
        formatId = 'red-salamander-terminal-engine-collection-neutrality'
        version = 1
    }
    Write-RSProviderTestJcs `
        -Path $neutralityPolicyPath `
        -Value $neutralityPolicy
    return [pscustomobject]@{
        Root = $root
        UniversePath = $universePath
        Universe = $universe
        CandidateSet = $candidateSet
        NeutralityPolicy = $neutralityPolicy
        NeutralityPolicyPath = $neutralityPolicyPath
        NeutralityPolicySha256 = (
            Get-FileHash `
                -LiteralPath $neutralityPolicyPath `
                -Algorithm SHA256).Hash.ToLowerInvariant()
    }
}

Describe 'Terminal-engine collection provider' {
    It 'is syntactically valid and bound to the round-4 universe leaf' {
        $provider = Join-Path $script:RepoRoot (
            'Tools\TerminalEngine\Invoke-TerminalEngineCollectionProvider.ps1')
        $tokens = $null
        $errors = $null
        [void][Management.Automation.Language.Parser]::ParseFile(
            $provider,
            [ref]$tokens,
            [ref]$errors)
        @($errors).Count | Should Be 0
        $script:RSProviderCandidateUniverseLeaf |
            Should Be (
                'Specs\Terminal\TerminalEngineCandidateUniverse.v4.json')
    }

    It 'derives artifact and wrapper accounting only from the descriptor' {
        $generic = Get-RSProviderArtifactContract `
            -Descriptor (New-RSProviderTestDescriptor `
                -BuildKind provider-generic-cmake-msvc-v1)
        $generic.BuildKind |
            Should Be 'provider-generic-cmake-msvc-v1'
        $generic.MaintainedBuildWrapper | Should Be $false
        $generic.AdapterDllLeaf | Should Be 'adapter-alpha.dll'
        $generic.AdapterProjectLeaf | Should Be 'adapter-alpha.vcxproj'

        $wrapper = Get-RSProviderArtifactContract `
            -Descriptor (New-RSProviderTestDescriptor)
        $wrapper.BuildKind | Should Be 'candidate-wrapper-v1'
        $wrapper.MaintainedBuildWrapper | Should Be $true
        $wrapper.AdapterProjectLeaf | Should Be ''
    }

    It 'rejects descriptor and assessment build-accounting disagreement' {
        $descriptor = New-RSProviderTestDescriptor
        (Test-RSProviderThrows -Action {
            Assert-RSProviderDescriptorAssessmentTruth `
                -Descriptor $descriptor `
                -PatchRecord @{
                    collectionDescriptorSha256 = $descriptor.Sha256
                    buildKind = 'candidate-wrapper-v1'
                    maintainedBuildWrapper = $false
                } `
                -CandidateId engine-alpha
        }) | Should Be $true
    }

    It 'uses only the authenticated closure assessment after governance closes' {
        $context = New-RSProviderTestContext `
            -Scratch (New-RSProviderTestScratch)
        Mock Read-RSProviderCanonicalJson `
            -ParameterFilter {
                $Label -ceq 'Frozen candidate assessment'
            } `
            -MockWith {
                throw 'Assessment paths must not be reread after closure.'
            }
        Mock Get-RSProviderSha256 `
            -ParameterFilter {
                $LiteralPath.EndsWith(
                    'TerminalEngineCandidateAssessment.v1.json',
                    [StringComparison]::OrdinalIgnoreCase)
            } `
            -MockWith {
                throw 'Assessment paths must not be rehashed after closure.'
            }

        $assessment = Get-RSProviderAssessmentCandidate `
            -Context $context `
            -Candidate $context.SelectedCandidates[0]
        $assessment['rawLedger']['patch']['marker'] |
            Should Be 'authenticated-closure-patch'
        Assert-MockCalled `
            Read-RSProviderCanonicalJson `
            -Times 0 `
            -Scope It `
            -ParameterFilter {
                $Label -ceq 'Frozen candidate assessment'
            }
        Assert-MockCalled `
            Get-RSProviderSha256 `
            -Times 0 `
            -Scope It `
            -ParameterFilter {
                $LiteralPath.EndsWith(
                    'TerminalEngineCandidateAssessment.v1.json',
                    [StringComparison]::OrdinalIgnoreCase)
            }
    }

    It 'forbids a wrapper input for every provider-generic build kind' {
        $scratch = New-RSProviderTestScratch
        (Test-RSProviderThrows -Action {
            New-RSProviderWrapperInput `
                -Context (New-RSProviderTestContext -Scratch $scratch) `
                -Candidate (New-RSProviderTestCandidate) `
                -Descriptor (New-RSProviderTestDescriptor `
                    -BuildKind provider-generic-cmake-msvc-v1) `
                -PatchRecord @{} `
                -SourceRoot $scratch `
                -BuildRoot $scratch `
                -NoticeState @{} `
                -HostHeaders ([pscustomobject]@{
                    IncludeRoot = $scratch
                    TreeSha256 = '6' * 64
                }) `
                -PathBindings @()
        }) | Should Be $true
    }

    It 'round-trips a provider-created wrapper input through one synthetic driver result' {
        $scratch = New-RSProviderTestScratch
        $source = Join-Path $scratch 'source'
        $build = Join-Path $scratch 'build'
        [void][IO.Directory]::CreateDirectory(
            (Join-Path $source 'tools'))
        [void][IO.Directory]::CreateDirectory($build)
        $notice = Join-Path $source 'generated\notices.txt'
        [void][IO.Directory]::CreateDirectory(
            [IO.Path]::GetDirectoryName($notice))
        [IO.File]::WriteAllText(
            $notice,
            "synthetic license`n",
            [Text.UTF8Encoding]::new($false))
        $driver = Join-Path $source 'tools\build-driver.ps1'
        [IO.File]::WriteAllText(
            $driver,
            @'
param([string]$InputPath,[string]$ResultPath)
$inputRecord = Get-Content -LiteralPath $InputPath -Raw |
    ConvertFrom-Json
if ($inputRecord.formatId -cne
    'red-salamander-terminal-engine-wrapper-input') {
    throw 'wrong input'
}
$record = [ordered]@{
    failure = $null
    formatId = 'red-salamander-terminal-engine-wrapper-result'
    status = 'pass'
    version = 1
}
[IO.File]::WriteAllText(
    $ResultPath,
    ($record | ConvertTo-Json -Compress),
    [Text.UTF8Encoding]::new($false))
'@,
            [Text.UTF8Encoding]::new($false))
        $context = New-RSProviderTestContext -Scratch $scratch
        Mock Get-RSProviderMSBuildPath {
            (Get-Command pwsh).Source
        }
        $descriptor = New-RSProviderTestDescriptor
        $input = New-RSProviderWrapperInput `
            -Context $context `
            -Candidate (New-RSProviderTestCandidate) `
            -Descriptor $descriptor `
            -PatchRecord @{
                patchSha256 = '3' * 64
                resultTree = '4' * 40
            } `
            -SourceRoot $source `
            -BuildRoot $build `
            -NoticeState ([pscustomobject]@{
                OutputPath = $notice
            }) `
            -HostHeaders ([pscustomobject]@{
                IncludeRoot = $scratch
                TreeSha256 = '6' * 64
            }) `
            -PathBindings @()

        $input.Record.candidate.buildKind |
            Should Be 'candidate-wrapper-v1'
        $input.Record.paths.hostHeaderIncludeRoot |
            Should Be ([IO.Path]::GetFullPath($scratch))
        $input.Record.gate0.hostHeaderTreeSha256 |
            Should Be ('6' * 64)
        @($input.Record.Keys).Count | Should Be 8
        & $driver -InputPath $input.Path -ResultPath $input.ResultPath
        $result = Read-RSProviderWrapperResult `
            -ResultPath $input.ResultPath `
            -RepoRoot $script:RepoRoot `
            -BuildRoot $build
        $result.Status | Should Be 'pass'
        $result.Record.failure | Should BeNullOrEmpty
    }

    It 'rejects a wrapper result outside the exact JCS schema' {
        $scratch = New-RSProviderTestScratch
        $resultPath = Join-Path $scratch 'result.jcs.json'
        Write-RSProviderTestJcs `
            -Path $resultPath `
            -Value ([ordered]@{
                formatId =
                    'red-salamander-terminal-engine-wrapper-result'
                version = 1
                status = 'pass'
                failure = $null
                artifactClaim = 'forbidden'
            })
        (Test-RSProviderThrows -Action {
            Read-RSProviderWrapperResult `
                -ResultPath $resultPath `
                -RepoRoot $script:RepoRoot `
                -BuildRoot $scratch
        }) | Should Be $true
    }

    It 'accepts exactly the descriptor DLL closure' {
        $scratch = New-RSProviderTestScratch
        $private = [object[]]@(
            [ordered]@{
                fileName = 'runtime-alpha.dll'
                relativeBuildPath = 'runtime/runtime-alpha.dll'
            })
        $descriptor = New-RSProviderTestDescriptor `
            -PrivateRuntimeLibraries $private
        [void][IO.Directory]::CreateDirectory(
            (Join-Path $scratch 'runtime'))
        [IO.File]::WriteAllBytes(
            (Join-Path $scratch 'adapter-alpha.dll'),
            [byte[]]@(1, 2, 3))
        [IO.File]::WriteAllBytes(
            (Join-Path $scratch 'runtime\runtime-alpha.dll'),
            [byte[]]@(4, 5, 6))

        $closure = Assert-RSProviderArtifactClosure `
            -BuildRoot $scratch `
            -Descriptor $descriptor
        $closure.AdapterSha256 | Should Match '^[0-9a-f]{64}$'
        @($closure.PrivateRuntimePaths).Count | Should Be 1
    }

    It 'rejects missing, extra, and case-folding duplicate DLL artifacts' {
        $scratch = New-RSProviderTestScratch
        $descriptor = New-RSProviderTestDescriptor
        (Test-RSProviderThrows -Action {
            Assert-RSProviderArtifactClosure `
                -BuildRoot $scratch `
                -Descriptor $descriptor
        }) | Should Be $true

        [IO.File]::WriteAllBytes(
            (Join-Path $scratch 'adapter-alpha.dll'),
            [byte[]]@(1))
        [IO.File]::WriteAllBytes(
            (Join-Path $scratch 'extra-alpha.dll'),
            [byte[]]@(2))
        (Test-RSProviderThrows -Action {
            Assert-RSProviderArtifactClosure `
                -BuildRoot $scratch `
                -Descriptor $descriptor
        }) | Should Be $true

        [IO.File]::Delete((Join-Path $scratch 'extra-alpha.dll'))
        [void][IO.Directory]::CreateDirectory(
            (Join-Path $scratch 'duplicate'))
        [IO.File]::WriteAllBytes(
            (Join-Path $scratch 'duplicate\ADAPTER-ALPHA.DLL'),
            [byte[]]@(3))
        (Test-RSProviderThrows -Action {
            Assert-RSProviderArtifactClosure `
                -BuildRoot $scratch `
                -Descriptor $descriptor
        }) | Should Be $true
    }

    It 'materializes inline legal provenance from candidate source bytes' {
        $scratch = New-RSProviderTestScratch
        $source = Join-Path $scratch 'source'
        $bootstrap = Join-Path $scratch 'bootstrap'
        [void][IO.Directory]::CreateDirectory($source)
        [void][IO.Directory]::CreateDirectory($bootstrap)
        $license = Join-Path $source 'license-alpha.txt'
        [IO.File]::WriteAllText(
            $license,
            "license alpha`n",
            [Text.UTF8Encoding]::new($false))
        $item = Get-Item -LiteralPath $license
        $legalInputs = [object[]]@(
            [ordered]@{
                inputId = 'license-alpha'
                kind = 'license'
                destinationLeaf = 'license-alpha.txt'
                sha256 = (
                    Get-FileHash `
                        -LiteralPath $license `
                        -Algorithm SHA256).Hash.ToLowerInvariant()
                sizeBytes = $item.Length.ToString(
                    [Globalization.CultureInfo]::InvariantCulture)
                source = [ordered]@{
                    sourceKind = 'candidate-source'
                    origin = 'upstream'
                    sourceRelativePath = 'license-alpha.txt'
                }
            })
        $descriptor = New-RSProviderTestDescriptor `
            -LegalInputs $legalInputs
        $state = Publish-RSProviderDescriptorNotices `
            -Context (New-RSProviderTestContext -Scratch $scratch) `
            -Descriptor $descriptor `
            -SourceRoot $source `
            -BootstrapRoot $bootstrap `
            -CandidateId engine-alpha

        [IO.File]::Exists($state.OutputPath) | Should Be $true
        (Get-Content -LiteralPath $state.OutputPath -Raw) |
            Should Match 'license alpha'
        $state.ClosureSha256 | Should Match '^[0-9a-f]{64}$'
    }

    It 'fails closed when a cached frozen input has the wrong digest' {
        $scratch = New-RSProviderTestScratch
        $input = [ordered]@{
            inputId = 'input-alpha'
            fileName = 'input-alpha.bin'
            sha256 = '0' * 64
        }
        $path = Get-RSProviderInputPath `
            -BootstrapRoot $scratch `
            -Platform x64 `
            -InputRecord $input
        [void][IO.Directory]::CreateDirectory(
            [IO.Path]::GetDirectoryName($path))
        [IO.File]::WriteAllBytes($path, [byte[]]@(1, 2, 3))
        (Test-RSProviderThrows -Action {
            Get-RSProviderVerifiedInput `
                -BootstrapRoot $scratch `
                -Platform x64 `
                -InputRecord $input
        }) | Should Be $true
    }

    It 'materializes one frozen candidate-neutral host WIL header tree' {
        $scratch = New-RSProviderTestScratch
        $packageRoot = Join-Path $scratch 'package'
        $wilRoot = Join-Path $packageRoot 'include\wil'
        [void][IO.Directory]::CreateDirectory($wilRoot)
        [IO.File]::WriteAllText(
            (Join-Path $wilRoot 'resource.h'),
            "#pragma once`n",
            [Text.UTF8Encoding]::new($false))
        [IO.File]::WriteAllText(
            (Join-Path $wilRoot 'result.h'),
            "#pragma once`n",
            [Text.UTF8Encoding]::new($false))
        $archive = Join-Path $scratch 'host-wil-headers.zip'
        [IO.Compression.ZipFile]::CreateFromDirectory(
            $packageRoot,
            $archive)
        $input = [ordered]@{
            inputId = 'host-wil-headers'
            inputKind = 'dependency-input'
            version = 'fixture'
            fileName = 'host-wil-headers.zip'
            sha256 = (
                Get-FileHash -LiteralPath $archive -Algorithm SHA256
            ).Hash.ToLowerInvariant()
            retrievalUrl =
                'https://example.invalid/host-wil-headers.zip'
            consumers = [string[]]@('engine-alpha')
        }
        $context = New-RSProviderTestContext `
            -Scratch $scratch `
            -Inputs ([object[]]@($input))
        $bootstrap = Join-Path $scratch 'bootstrap'
        $cached = Get-RSProviderInputPath `
            -BootstrapRoot $bootstrap `
            -Platform x64 `
            -InputRecord $input
        [void][IO.Directory]::CreateDirectory(
            [IO.Path]::GetDirectoryName($cached))
        [IO.File]::Copy($archive, $cached)

        $headers = Resolve-RSProviderHostHeaders `
            -Context $context `
            -BootstrapRoot $bootstrap
        $headers.InputId | Should Be 'host-wil-headers'
        $headers.TreeSha256 | Should Match '^[0-9a-f]{64}$'
        [IO.File]::Exists(
            (Join-Path $headers.IncludeRoot 'wil\resource.h')) |
            Should Be $true
        $originalTreeSha256 = [string]$headers.TreeSha256
        $originalArchiveSha256 = [string]$input.sha256
        [IO.File]::WriteAllText(
            (Join-Path $headers.IncludeRoot 'wil\resource.h'),
            "// mutated between provider phases`n",
            [Text.UTF8Encoding]::new($false))
        (Test-RSProviderThrows -Action {
            Assert-RSProviderHostHeadersIdentity `
                -Context $context `
                -BootstrapRoot $bootstrap `
                -HostHeaders $headers
        }) | Should Be $true
        $headers = Resolve-RSProviderHostHeaders `
            -Context $context `
            -BootstrapRoot $bootstrap
        $headers.TreeSha256 | Should Be $originalTreeSha256
        (Get-FileHash -LiteralPath $archive -Algorithm SHA256).
            Hash.ToLowerInvariant() |
            Should Be $originalArchiveSha256
        [IO.File]::ReadAllText(
            (Join-Path $headers.IncludeRoot 'wil\resource.h')) |
            Should Be "#pragma once`n"
        (Get-Command ConvertTo-RSJcsUtf8Bytes -ErrorAction Stop).Name |
            Should Be 'ConvertTo-RSJcsUtf8Bytes'
        $postIdentityJcs = Join-Path $scratch 'post-host-identity.jcs.json'
        Write-RSProviderTestJcs `
            -Path $postIdentityJcs `
            -Value ([ordered]@{ status = 'caller-module-export-preserved' })
        [IO.File]::Exists($postIdentityJcs) | Should Be $true

        $input.consumers = [string[]]@('engine-beta')
        (Test-RSProviderThrows -Action {
            Resolve-RSProviderHostHeaders `
                -Context $context `
                -BootstrapRoot $bootstrap
        }) | Should Be $true
    }

    It 'rejects an input leaf that could escape its cache directory' {
        (Test-RSProviderThrows -Action {
            Get-RSProviderInputPath `
                -BootstrapRoot (New-RSProviderTestScratch) `
                -Platform x64 `
                -InputRecord @{
                    inputId = 'input-alpha'
                    fileName = '..\escape.bin'
                }
        }) | Should Be $true
    }

    It 'loads a candidate set only through its exact round-4 governance closure' {
        $fixture = New-RSProviderSyntheticRepo
        $script:ProviderFixtureClosure = [pscustomobject]@{
            CandidateSet = $fixture.CandidateSet
            Universe = $fixture.Universe
            CandidateSetSha256 = (
                Get-FileHash `
                    -LiteralPath (Join-Path $fixture.Root (
                        'Specs\Terminal\TerminalEngineCandidateSet.v1.json')) `
                    -Algorithm SHA256).Hash.ToLowerInvariant()
            UniverseSha256 = (
                Get-FileHash `
                    -LiteralPath $fixture.UniversePath `
                    -Algorithm SHA256).Hash.ToLowerInvariant()
            CandidateReviewSha256 = 'c' * 64
            CandidateAssessmentSha256 = 'd' * 64
            NeutralityPolicySha256 = $fixture.NeutralityPolicySha256
            AssessmentResult = [pscustomobject]@{
                CandidateById = (
                    [Collections.Generic.Dictionary[string, object]]::new(
                        [StringComparer]::Ordinal))
            }
        }
        Mock Read-RSProviderCandidateGovernanceClosure {
            return $script:ProviderFixtureClosure
        }
        $context = Get-RSProviderContext `
            -RepoRoot $fixture.Root `
            -CandidateId engine-alpha `
            -Platform x64 `
            -Configuration Release
        $context.SelectedIds | Should Be @('engine-alpha')
        $context.CandidateSet | Should Be $fixture.CandidateSet
        $context.CandidateUniverse | Should Be $fixture.Universe
        Assert-MockCalled `
            Read-RSProviderCandidateGovernanceClosure `
            -Times 1 `
            -Scope It

        [IO.File]::AppendAllText(
            $fixture.UniversePath,
            "`n",
            [Text.UTF8Encoding]::new($false))
        (Test-RSProviderThrows -Action {
            Get-RSProviderContext `
                -RepoRoot $fixture.Root `
                -CandidateId engine-alpha `
                -Platform x64 `
                -Configuration Release
        }) | Should Be $true
    }

    It 'rejects a swapped self-consistent neutrality policy before effects' {
        $fixture = New-RSProviderSyntheticRepo
        Mock Read-RSProviderCandidateGovernanceClosure {
            throw [IO.InvalidDataException]::new(
                'Frozen collection-neutrality policy differs from its authenticated frozen identity.')
        }
        Write-RSProviderTestJcs `
            -Path $fixture.NeutralityPolicyPath `
            -Value ([ordered]@{
                formatId =
                    'red-salamander-terminal-engine-collection-neutrality'
                version = 2
            })

        $message = Get-RSProviderExceptionMessage -Action {
            Get-RSProviderContext `
                -RepoRoot $fixture.Root `
                -CandidateId engine-alpha `
                -Platform x64 `
                -Configuration Release
        }
        $message | Should Match 'authenticated frozen identity'
        Assert-MockCalled `
            Read-RSProviderCandidateGovernanceClosure `
            -Times 1 `
            -Scope It
        [IO.Directory]::Exists((Join-Path $fixture.Root '.build')) |
            Should Be $false
    }

    It 'rejects a partial provider phase candidate set' {
        $fixture = New-RSProviderSyntheticRepo
        Mock Read-RSProviderCandidateGovernanceClosure {
            throw 'Governance closure must not run for a partial cohort.'
        }
        $fixture.CandidateSet.candidates = [object[]]@(
            $fixture.CandidateSet.candidates +
            [ordered]@{
                candidateId = 'control-beta'
                pin = 'control-beta-pin'
                pinSha256 = 'c' * 64
                scoredOrControl = 'control'
                constraintId = 'candidate-owned-control'
            })
        Write-RSProviderTestJcs `
            -Path (Join-Path $fixture.Root (
                'Specs\Terminal\TerminalEngineCandidateSet.v1.json')) `
            -Value $fixture.CandidateSet

        $message = Get-RSProviderExceptionMessage -Action {
            Get-RSProviderContext `
                -RepoRoot $fixture.Root `
                -CandidateId engine-alpha `
                -Platform x64 `
                -Configuration Release
        }
        $message | Should Match 'partial collection manifests'
        Assert-MockCalled `
            Read-RSProviderCandidateGovernanceClosure `
            -Times 0 `
            -Scope It
    }

    It 'rejects a static-only provider phase before governance closure or effects' {
        $fixture = New-RSProviderSyntheticRepo
        $candidate = $fixture.CandidateSet.candidates[0]
        $candidate['scoredOrControl'] = 'control'
        $candidate.Remove('disposition')
        $candidate['constraintId'] = 'candidate-owned-control'
        Write-RSProviderTestJcs `
            -Path (Join-Path $fixture.Root (
                'Specs\Terminal\TerminalEngineCandidateSet.v1.json')) `
            -Value $fixture.CandidateSet
        Mock Read-RSProviderCandidateGovernanceClosure {
            throw 'Governance closure must not run for a static-only cohort.'
        }
        $markerPath = Join-Path $fixture.Root (
            'Tools\preliminary-helper-executed.txt')
        [IO.File]::WriteAllText(
            (Join-Path $fixture.Root 'Tools\TerminalEvidence.psm1'),
            @'
[IO.File]::WriteAllText(
    (Join-Path $PSScriptRoot 'preliminary-helper-executed.txt'),
    'executed')
throw 'Unauthenticated helper module executed.'
'@,
            [Text.UTF8Encoding]::new($false))
        Remove-Module TerminalEvidence -Force -ErrorAction SilentlyContinue
        try {
            $message = Get-RSProviderExceptionMessage -Action {
                Get-RSProviderContext `
                    -RepoRoot $fixture.Root `
                    -CandidateId engine-alpha `
                    -Platform x64 `
                    -Configuration Release
            }
        }
        finally {
            Import-Module (
                Join-Path $script:RepoRoot 'Tools\TerminalEvidence.psm1') `
                -Force `
                -Global
        }
        $message | Should Match 'static-only'
        Assert-MockCalled `
            Read-RSProviderCandidateGovernanceClosure `
            -Times 0 `
            -Scope It
        [IO.Directory]::Exists((Join-Path $fixture.Root '.build')) |
            Should Be $false
        [IO.File]::Exists($markerPath) | Should Be $false
    }

    It 'rejects an invalid chronology closure before provider effects' {
        $fixture = New-RSProviderSyntheticRepo
        Mock Read-RSProviderCandidateGovernanceClosure {
            throw [IO.InvalidDataException]::new(
                'candidate approval equals the nomination-window boundary')
        }

        $message = Get-RSProviderExceptionMessage -Action {
            Get-RSProviderContext `
                -RepoRoot $fixture.Root `
                -CandidateId engine-alpha `
                -Platform x64 `
                -Configuration Release
        }
        $message | Should Match 'nomination-window boundary'
        Assert-MockCalled `
            Read-RSProviderCandidateGovernanceClosure `
            -Times 1 `
            -Scope It
        [IO.Directory]::Exists((Join-Path $fixture.Root '.build')) |
            Should Be $false
    }

    It 'rejects a missing round-4 universe before closure or provider effects' {
        $fixture = New-RSProviderSyntheticRepo
        Remove-Item -LiteralPath $fixture.UniversePath
        Mock Read-RSProviderCandidateGovernanceClosure {
            throw 'Governance closure must not run without round 4.'
        }

        $message = Get-RSProviderExceptionMessage -Action {
            Get-RSProviderContext `
                -RepoRoot $fixture.Root `
                -CandidateId engine-alpha `
                -Platform x64 `
                -Configuration Release
        }
        $message | Should Match 'Required file does not exist'
        Assert-MockCalled `
            Read-RSProviderCandidateGovernanceClosure `
            -Times 0 `
            -Scope It
        [IO.Directory]::Exists((Join-Path $fixture.Root '.build')) |
            Should Be $false
    }

    It 'cleans only selected lane roots and preserves neighboring state' {
        $scratch = New-RSProviderTestScratch
        $context = New-RSProviderTestContext -Scratch $scratch
        $context.RepoRoot = $scratch
        $selected = Get-RSProviderCandidateBuildRoot `
            -Context $context `
            -CandidateId engine-alpha
        $neighbor = Get-RSProviderCandidateBuildRoot `
            -Context $context `
            -CandidateId engine-beta
        [void][IO.Directory]::CreateDirectory($selected)
        [void][IO.Directory]::CreateDirectory($neighbor)
        [IO.File]::WriteAllText(
            (Join-Path $selected 'selected.txt'),
            'selected')
        [IO.File]::WriteAllText(
            (Join-Path $neighbor 'neighbor.txt'),
            'neighbor')

        $result = Invoke-RSProviderClean -Context $context
        $result.clean | Should Be $true
        [IO.Directory]::Exists($selected) | Should Be $false
        [IO.File]::Exists(
            (Join-Path $neighbor 'neighbor.txt')) | Should Be $true
    }

    It 'derives static source and control outcomes without executable code' {
        $context = New-RSProviderTestContext `
            -Scratch (New-RSProviderTestScratch)
        $control = New-RSProviderStaticCandidateResult `
            -Context $context `
            -Candidate @{
            candidateId = 'control-alpha'
            pin = 'pin-control'
            pinSha256 = '1' * 64
            scoredOrControl = 'control'
            constraintId = 'candidate-owned-control'
        }
        $control.outcome | Should Be 'constraint-control'
        @($control.Keys).Count | Should Be 5
        $control.controlFailure.evidenceSha256 |
            Should Match '^[0-9a-f]{64}$'

        $rejected = New-RSProviderStaticCandidateResult `
            -Context $context `
            -Candidate @{
            candidateId = 'source-alpha'
            pin = 'pin-source'
            pinSha256 = '2' * 64
            scoredOrControl = 'scored'
            disposition = 'source-rejected'
            sourceRejection = @{ gateId = 'G7' }
        }
        $rejected.outcome | Should Be 'source-rejected'
        @($rejected.Keys).Count | Should Be 5
    }

    It 'rejects every static-only Collect shape before mutation or execution' {
        Mock Invoke-RSProviderCollectCandidate {
            throw 'Collect candidate must not run for a static-only cohort.'
        }
        Mock Start-Process {
            throw 'A subprocess must not start for a static-only cohort.'
        }
        $cases = [object[]]@(
            [pscustomobject]@{
                Name = 'empty selection'
                Candidates = [object[]]@()
            },
            [pscustomobject]@{
                Name = 'source-rejected only'
                Candidates = [object[]]@(
                    [ordered]@{
                        candidateId = 'source-alpha'
                        scoredOrControl = 'scored'
                        disposition = 'source-rejected'
                    })
            },
            [pscustomobject]@{
                Name = 'controls only'
                Candidates = [object[]]@(
                    [ordered]@{
                        candidateId = 'control-alpha'
                        scoredOrControl = 'control'
                    })
            },
            [pscustomobject]@{
                Name = 'mixed static records'
                Candidates = [object[]]@(
                    [ordered]@{
                        candidateId = 'control-alpha'
                        scoredOrControl = 'control'
                    },
                    [ordered]@{
                        candidateId = 'source-alpha'
                        scoredOrControl = 'scored'
                        disposition = 'source-rejected'
                    })
            })
        foreach ($case in $cases) {
            $scratch = New-RSProviderTestScratch
            $sentinel = Join-Path $scratch 'sentinel.bin'
            [IO.File]::WriteAllBytes($sentinel, [byte[]](1, 2, 3, 4))
            $before = (Get-FileHash `
                    -LiteralPath $sentinel `
                    -Algorithm SHA256).Hash
            $context = New-RSProviderTestContext -Scratch $scratch
            $context.SelectedCandidates = $case.Candidates
            $message = Get-RSProviderExceptionMessage -Action {
                Invoke-RSProviderCollect -Context $context
            }
            $message | Should Be (
                'Collect requires at least one selected executable candidate; static records cannot attest a lane.')
            [IO.File]::Exists($sentinel) | Should Be $true
            (Get-FileHash `
                    -LiteralPath $sentinel `
                    -Algorithm SHA256).Hash | Should Be $before
            @(
                Get-ChildItem `
                    -LiteralPath $scratch `
                    -Force `
                    -Recurse `
                    -File).Count | Should Be 1
        }
        Assert-MockCalled Invoke-RSProviderCollectCandidate -Times 0
        Assert-MockCalled Start-Process -Times 0
    }

    It 'rejects executable collection until the next approved universe round' {
        $context = New-RSProviderTestContext `
            -Scratch (New-RSProviderTestScratch)
        $context.SelectedCandidates = [object[]]@(
            (New-RSProviderTestCandidate))
        (Test-RSProviderThrows -Action {
                Invoke-RSProviderCollect -Context $context
            }) | Should Be $true
    }

    It 'rejects an incomplete runner isolation attestation' {
        $attestation = New-RSProviderPassingIsolation
        $attestation.ruleRemovalVerified = $false
        (Test-RSProviderThrows -Action {
            Assert-RSProviderIsolationAttestation `
                -Attestation $attestation `
                -CandidateId engine-alpha
        }) | Should Be $true
    }

    It 'rejects a stale policy baseline before the dynamic runtime scan' {
        $scratch = New-RSProviderTestScratch
        $context = New-RSProviderTestContext -Scratch $scratch
        $context | Add-Member `
            -NotePropertyName CandidateUniverse `
            -NotePropertyValue @{}
        $context.CandidateSet = @{}
        Mock Invoke-RSProviderAuthenticatedNeutralityProof {
            throw 'stale synthetic baseline'
        }

        (Test-RSProviderThrows -Action {
                Invoke-RSProviderNeutralityProof -Context $context
            }) | Should Be $true
        Assert-MockCalled `
            Invoke-RSProviderAuthenticatedNeutralityProof `
            -Times 1 `
            -Scope It `
            -ParameterFilter {
                $RepoRoot -ceq $context.RepoRoot -and
                $ExpectedNeutralityPolicySha256 -ceq
                    $context.NeutralityPolicySha256 -and
                $ExpectedUniverseSha256 -ceq
                    $context.CandidateUniverseSha256 -and
                $ExpectedCandidateSetSha256 -ceq
                    $context.CandidateSetSha256 -and
                $ExpectedCandidateReviewSha256 -ceq
                    $context.CandidateReviewSha256 -and
                $ExpectedCandidateAssessmentSha256 -ceq
                    $context.CandidateAssessmentSha256
            }
    }

    It 'delegates runtime neutrality to the authenticated evaluator entrypoint' {
        $scratch = New-RSProviderTestScratch
        $context = New-RSProviderTestContext -Scratch $scratch
        $descriptor = [object[]]@(
            [ordered]@{ candidateId = 'engine-alpha' })
        $context.Inputs = [object[]]@(
            [ordered]@{ inputId = 'source-alpha' })
        Mock Invoke-RSProviderAuthenticatedNeutralityProof {
            return [pscustomobject]@{
                Baseline = [pscustomobject]@{ status = 'verified' }
                Runtime = [pscustomobject]@{ status = 'verified' }
            }
        }

        $proof = Invoke-RSProviderNeutralityProof `
            -Context $context `
            -Descriptor $descriptor
        $proof.Baseline.status | Should Be 'verified'
        $proof.Runtime.status | Should Be 'verified'
        Assert-MockCalled `
            Invoke-RSProviderAuthenticatedNeutralityProof `
            -Times 1 `
            -Scope It `
            -ParameterFilter {
                $RepoRoot -ceq $context.RepoRoot -and
                $ExpectedNeutralityPolicySha256 -ceq
                    $context.NeutralityPolicySha256 -and
                $ExpectedUniverseSha256 -ceq
                    $context.CandidateUniverseSha256 -and
                $ExpectedCandidateSetSha256 -ceq
                    $context.CandidateSetSha256 -and
                $ExpectedCandidateReviewSha256 -ceq
                    $context.CandidateReviewSha256 -and
                $ExpectedCandidateAssessmentSha256 -ceq
                    $context.CandidateAssessmentSha256 -and
                @($Descriptor).Count -eq 1 -and
                @($FrozenInput).Count -eq 1
            }
    }

    It 'retains one evaluator module binding across closure and neutrality phases' {
        $root = Join-Path $TestDrive 'retained-evaluator'
        $tools = Join-Path $root 'Tools'
        [void][IO.Directory]::CreateDirectory($tools)
        $evaluatorPath = Join-Path $tools 'TerminalEngineEvaluator.psm1'
        $markerPath = Join-Path $tools 'swapped-evaluator-executed.txt'
        [IO.File]::WriteAllText(
            $evaluatorPath,
            @'
function Read-RSCandidateGovernanceClosure {
    param([string]$RepoRoot)
    return [pscustomobject]@{ phase = 'closure'; root = $RepoRoot }
}
function Invoke-RSAuthenticatedCollectionNeutralityProof {
    param(
        [string]$RepoRoot,
        [string]$ExpectedNeutralityPolicySha256,
        [string]$ExpectedUniverseSha256,
        [string]$ExpectedCandidateSetSha256,
        [string]$ExpectedCandidateReviewSha256,
        [string]$ExpectedCandidateAssessmentSha256,
        [object[]]$Descriptor,
        [object[]]$FrozenInput)
    return [pscustomobject]@{
        Baseline = [pscustomobject]@{ status = 'verified' }
        Runtime = [pscustomobject]@{ status = 'verified' }
    }
}
Export-ModuleMember -Function @(
    'Read-RSCandidateGovernanceClosure',
    'Invoke-RSAuthenticatedCollectionNeutralityProof')
'@,
            [Text.UTF8Encoding]::new($false))

        $script:RSProviderEvaluatorModule = $null
        try {
            $bound = Get-RSProviderEvaluatorModule -RepoRoot $root
            $closure = & $bound {
                param($repoRoot)
                Read-RSCandidateGovernanceClosure -RepoRoot $repoRoot
            } $root
            $closure.phase | Should Be 'closure'
            [IO.File]::WriteAllText(
                $evaluatorPath,
                @'
[IO.File]::WriteAllText(
    (Join-Path $PSScriptRoot 'swapped-evaluator-executed.txt'),
    'executed')
throw 'Swapped evaluator executed.'
'@,
                [Text.UTF8Encoding]::new($false))

            $proof = Invoke-RSProviderAuthenticatedNeutralityProof `
                -RepoRoot $root `
                -ExpectedNeutralityPolicySha256 ('6' * 64) `
                -ExpectedUniverseSha256 ('a' * 64) `
                -ExpectedCandidateSetSha256 ('b' * 64) `
                -ExpectedCandidateReviewSha256 ('c' * 64) `
                -ExpectedCandidateAssessmentSha256 ('d' * 64)
            $proof.Runtime.status | Should Be 'verified'
            [object]::ReferenceEquals(
                $bound,
                $script:RSProviderEvaluatorModule) | Should Be $true
            [IO.File]::Exists($markerPath) | Should Be $false
        }
        finally {
            if ($null -ne $script:RSProviderEvaluatorModule) {
                Remove-Module `
                    -ModuleInfo $script:RSProviderEvaluatorModule `
                    -Force `
                    -ErrorAction SilentlyContinue
            }
            $script:RSProviderEvaluatorModule = $null
        }
    }
}
