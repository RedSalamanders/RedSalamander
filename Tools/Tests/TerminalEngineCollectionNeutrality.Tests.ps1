Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:RepoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$script:ModulePath = Join-Path $script:RepoRoot (
    'Tools\TerminalEngine\TerminalEngineCollectionNeutrality.psm1')
Import-Module $script:ModulePath -Force

function New-RSNeutralityFixture {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root
    )

    foreach ($relative in (
            Get-RSTerminalEngineCollectionNeutralitySurface)) {
        $path = Join-Path $Root ($relative.Replace('/', '\'))
        [void](New-Item -ItemType Directory -Path (
                Split-Path -Parent $path) -Force)
        [IO.File]::WriteAllText(
            $path,
            'synthetic-neutral-host-surface',
            [Text.UTF8Encoding]::new($false))
    }
    return $Root
}

function Test-RSNeutralityThrows {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Action
    )

    $threw = $false
    try {
        $null = & $Action
    }
    catch {
        $threw = $true
    }
    return $threw
}

Describe 'Terminal-engine collection neutrality scanner' {
    BeforeEach {
        $script:Fixture = New-RSNeutralityFixture -Root (
            Join-Path $TestDrive ([Guid]::NewGuid().ToString('N')))
        $script:Literal = 'vendoromega'
        $script:TargetRelative =
            'Tools/TerminalEngine/Invoke-TerminalEngineCollectionProvider.ps1'
        $script:Target = Join-Path $script:Fixture (
            $script:TargetRelative.Replace('/', '\'))
    }

    It 'includes the exact sorted root build and lane-orchestration inputs' {
        $surface = [string[]]@(
            Get-RSTerminalEngineCollectionNeutralitySurface)
        $surface | Should Be ([string[]]@(
                $surface | Sort-Object -CaseSensitive))
        foreach ($path in @(
                '.github/workflows/terminal-engine-gate0.yml',
                'Directory.Build.props',
                'Directory.Build.targets',
                'RuntimeDependencies.props',
                'vcpkg.json')) {
            ($surface -ccontains $path) | Should Be $true
        }
    }

    It 'returns a content-free proof for an exact synthetic host surface' {
        $result = Invoke-RSTerminalEngineCollectionNeutralityScan `
            -RepoRoot $script:Fixture `
            -AdditionalLiteral $script:Literal

        $result.status | Should Be 'verified'
        $result.algorithmId |
            Should Be 'terminal-engine-collection-neutrality-v1'
        $result.proofSha256 | Should Match '^[0-9a-f]{64}$'
        $result.forbiddenLiteralSetSha256 |
            Should Match '^[0-9a-f]{64}$'
    }

    It 'keeps full-cohort host-header and toolchain inputs neutral' {
        [IO.File]::WriteAllText(
            $script:Target,
            (
                "host-wil-headers host-wil-headers.zip " +
                "generic msvc toolchain`n"),
            [Text.UTF8Encoding]::new($false))
        $governance = [ordered]@{
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
        $hostInput = [ordered]@{
            inputId = 'host-wil-headers'
            inputKind = 'dependency-input'
            version = 'fixture'
            fileName = 'host-wil-headers.zip'
            sha256 = '0' * 64
            retrievalUrl = 'https://example.invalid/host-wil-headers.zip'
            consumers = [string[]]@('engine-alpha', 'engine-beta')
        }
        $toolchainInputId = 'toolchain-{0}' -f 'msvc'
        $toolchainFileName = '{0}.zip' -f 'msvc'
        $toolInput = [ordered]@{
            inputId = $toolchainInputId
            inputKind = 'toolchain-input'
            version = 'fixture'
            fileName = $toolchainFileName
            sha256 = '1' * 64
            retrievalUrl =
                'https://example.invalid/{0}' -f $toolchainFileName
            consumers = [string[]]@('engine-alpha', 'engine-beta')
        }
        $forbidden = Get-RSTerminalEngineCollectionForbiddenLiteral `
            -GovernanceObject @($governance) `
            -FrozenInput @($hostInput, $toolInput)
        ($forbidden -ccontains 'host-wil-headers') | Should Be $false
        ($forbidden -ccontains 'host-wil-headers.zip') | Should Be $false
        ($forbidden -ccontains $toolchainInputId) | Should Be $true
        ($forbidden -ccontains $toolchainFileName) | Should Be $true
        ($forbidden -ccontains 'msvc') | Should Be $false
        ($forbidden -ccontains 'engine-alpha') | Should Be $true
        ($forbidden -ccontains 'engine-beta') | Should Be $true

        $result = Invoke-RSTerminalEngineCollectionNeutralityScan `
            -RepoRoot $script:Fixture `
            -GovernanceObject @($governance) `
            -FrozenInput @($hostInput, $toolInput)
        $result.status | Should Be 'verified'
        $result.findingCount | Should Be 0

        $hostInput.consumers = [string[]]@('engine-alpha')
        $wrongCohort = Get-RSTerminalEngineCollectionForbiddenLiteral `
            -GovernanceObject @($governance) `
            -FrozenInput @($hostInput, $toolInput)
        ($wrongCohort -ccontains 'host-wil-headers') | Should Be $true
    }

    It 'still forbids an exact toolchain id and a candidate dependency or source input' {
        $governance = [ordered]@{
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
        $toolchainInputId = 'toolchain-{0}' -f 'msvc'
        $toolchainFileName = '{0}.zip' -f 'msvc'
        [IO.File]::WriteAllText(
            $script:Target,
            "$toolchainInputId`n",
            [Text.UTF8Encoding]::new($false))
        (Test-RSNeutralityThrows -Action {
            Invoke-RSTerminalEngineCollectionNeutralityScan `
                -RepoRoot $script:Fixture `
                -GovernanceObject @($governance) `
                -FrozenInput @([ordered]@{
                        inputId = $toolchainInputId
                        inputKind = 'toolchain-input'
                        version = 'fixture'
                        fileName = $toolchainFileName
                        sha256 = '1' * 64
                        retrievalUrl =
                            'https://example.invalid/{0}' -f
                                $toolchainFileName
                        consumers =
                            [string[]]@('engine-alpha', 'engine-beta')
                    })
        }) | Should Be $true

        $candidateInputs = [object[]]@(
            [ordered]@{
                inputId = 'sharedomega-support'
                inputKind = 'dependency-input'
                consumers = [string[]]@('engine-alpha', 'engine-beta')
            },
            [ordered]@{
                inputId = 'parseromega'
                inputKind = 'dependency-input'
                consumers = [string[]]@('engine-alpha')
            },
            [ordered]@{
                inputId = 'sourceomega'
                inputKind = 'source-archive'
                consumers = [string[]]@('engine-alpha', 'engine-beta')
            })
        foreach ($candidateInput in $candidateInputs) {
            $inputId = [string]$candidateInput.inputId
            $candidateRecord = [ordered]@{
                inputId = $inputId
                inputKind = [string]$candidateInput.inputKind
                version = 'fixture'
                fileName = "$inputId.zip"
                sha256 = '2' * 64
                retrievalUrl =
                    "https://example.invalid/$inputId.zip"
                consumers = [string[]]$candidateInput.consumers
            }
            [IO.File]::WriteAllText(
                $script:Target,
                "$inputId`n",
                [Text.UTF8Encoding]::new($false))
            $derived = Get-RSTerminalEngineCollectionForbiddenLiteral `
                -GovernanceObject @($governance) `
                -FrozenInput @($candidateRecord)
            ($derived -ccontains $inputId) | Should Be $true
            $rejected = Test-RSNeutralityThrows -Action {
                Invoke-RSTerminalEngineCollectionNeutralityScan `
                    -RepoRoot $script:Fixture `
                    -GovernanceObject @($governance) `
                    -FrozenInput @($candidateRecord)
            }
            if (-not $rejected) {
                throw "Neutrality scan accepted candidate input '$inputId'."
            }
        }
    }

    It 'accepts blank lines in a scanned host file' {
        [IO.File]::WriteAllText(
            $script:Target,
            "synthetic-neutral-host-surface`r`n`r`nsecond-neutral-line`r`n",
            [Text.UTF8Encoding]::new($false))

        $result = Invoke-RSTerminalEngineCollectionNeutralityScan `
            -RepoRoot $script:Fixture `
            -AdditionalLiteral $script:Literal

        $result.status | Should Be 'verified'
        $result.findingCount | Should Be 0
    }

    It 'rejects an embedded identifier and mixed case' {
        [IO.File]::WriteAllText(
            $script:Target,
            'function Get-RSProviderVendorOmegaState {}',
            [Text.UTF8Encoding]::new($false))

        (Test-RSNeutralityThrows -Action {
                Invoke-RSTerminalEngineCollectionNeutralityScan `
                    -RepoRoot $script:Fixture `
                    -AdditionalLiteral $script:Literal
            }) | Should Be $true
    }

    It 'rejects an embedded JSON-style key and artifact leaf' {
        [IO.File]::WriteAllText(
            $script:Target,
            "vendoromegaArchiveSha256 = 'adapter-vendoromega.dll'",
            [Text.UTF8Encoding]::new($false))

        (Test-RSNeutralityThrows -Action {
                Invoke-RSTerminalEngineCollectionNeutralityScan `
                    -RepoRoot $script:Fixture `
                    -AdditionalLiteral $script:Literal
            }) | Should Be $true
    }

    It 'rejects a constructed string literal' {
        [IO.File]::WriteAllText(
            $script:Target,
            "[string]::Concat('vendor','omega')",
            [Text.UTF8Encoding]::new($false))

        (Test-RSNeutralityThrows -Action {
                Invoke-RSTerminalEngineCollectionNeutralityScan `
                    -RepoRoot $script:Fixture `
                    -AdditionalLiteral $script:Literal
            }) | Should Be $true
    }

    It 'rejects adjacent C-family quoted literals and the distinctive core alias' {
        $cppRelative =
            'Tools/TerminalEngine/TerminalEngineGate0Harness.cpp'
        $cppPath = Join-Path $script:Fixture (
            $cppRelative.Replace('/', '\'))
        [IO.File]::WriteAllText(
            $cppPath,
            'auto value = "vendor" "-terminal";',
            [Text.UTF8Encoding]::new($false))
        $governance = [ordered]@{
            dependencyId = 'vendor-terminal-core'
        }

        (Test-RSNeutralityThrows -Action {
                Invoke-RSTerminalEngineCollectionNeutralityScan `
                    -RepoRoot $script:Fixture `
                    -GovernanceObject @($governance)
            }) | Should Be $true
    }

    It 'does not turn a distinctive core alias into the generic terminal word' {
        [IO.File]::WriteAllText(
            $script:Target,
            'function Invoke-TerminalEngineCollection {}',
            [Text.UTF8Encoding]::new($false))
        $governance = [ordered]@{
            dependencyId = 'vendor-terminal-core'
        }
        $forbidden = Get-RSTerminalEngineCollectionForbiddenLiteral `
            -GovernanceObject @($governance)

        ($forbidden -ccontains 'vendor-terminal') | Should Be $true
        ($forbidden -ccontains 'terminal') | Should Be $false
        $result = Invoke-RSTerminalEngineCollectionNeutralityScan `
            -RepoRoot $script:Fixture `
            -GovernanceObject @($governance)
        $result.status | Should Be 'verified'
    }

    It 'accepts only an exact used line-hash exception' {
        $line = 'function Get-RSProviderVendorOmegaState {}'
        $text = "first neutral line`nsecond neutral line`n$line"
        [IO.File]::WriteAllText(
            $script:Target,
            $text,
            [Text.UTF8Encoding]::new($false))
        $lineBytes = [Text.UTF8Encoding]::new($false).GetBytes($line)
        $exception = [ordered]@{
            path = $script:TargetRelative
            line = 3
            literal = $script:Literal
            lineSha256 = [Convert]::ToHexString(
                [Security.Cryptography.SHA256]::HashData($lineBytes)).
                ToLowerInvariant()
        }

        $result = Invoke-RSTerminalEngineCollectionNeutralityScan `
            -RepoRoot $script:Fixture `
            -AdditionalLiteral $script:Literal `
            -Exception @($exception)
        $result.status | Should Be 'verified'

        $exception.line = 1
        (Test-RSNeutralityThrows -Action {
                Invoke-RSTerminalEngineCollectionNeutralityScan `
                    -RepoRoot $script:Fixture `
                    -AdditionalLiteral $script:Literal `
                    -Exception @($exception)
            }) | Should Be $true
    }

    It 'validates the frozen policy surface and exact baseline proof' {
        $baseline = Invoke-RSTerminalEngineCollectionNeutralityScan `
            -RepoRoot $script:Fixture `
            -AdditionalLiteral $script:Literal
        $policy = [ordered]@{
            formatId =
                'red-salamander-terminal-engine-collection-neutrality'
            version = 1
            surfacePaths = [string[]]@(
                Get-RSTerminalEngineCollectionNeutralitySurface)
            baselineAdditionalLiterals = [string[]]@($script:Literal)
            exceptions = [object[]]@()
            baselineProof = [ordered]@{
                algorithmId = [string]$baseline.algorithmId
                surfaceSetSha256 = [string]$baseline.surfaceSetSha256
                forbiddenLiteralSetSha256 =
                    [string]$baseline.forbiddenLiteralSetSha256
                exceptionSetSha256 =
                    [string]$baseline.exceptionSetSha256
                findingCount = [int]$baseline.findingCount
                status = [string]$baseline.status
                proofSha256 = [string]$baseline.proofSha256
            }
        }

        $result = Invoke-RSTerminalEngineCollectionNeutralityScan `
            -RepoRoot $script:Fixture `
            -PolicyRecord $policy `
            -ValidateBaselineProof
        $result.proofSha256 | Should Be $baseline.proofSha256

        $policy.baselineProof.proofSha256 = '0' * 64
        (Test-RSNeutralityThrows -Action {
                Invoke-RSTerminalEngineCollectionNeutralityScan `
                    -RepoRoot $script:Fixture `
                    -PolicyRecord $policy `
                    -ValidateBaselineProof
            }) | Should Be $true
    }

    It 'rejects a stale exact exception' {
        $exception = [ordered]@{
            path = $script:TargetRelative
            line = 1
            literal = $script:Literal
            lineSha256 = '0' * 64
        }

        (Test-RSNeutralityThrows -Action {
                Invoke-RSTerminalEngineCollectionNeutralityScan `
                    -RepoRoot $script:Fixture `
                    -AdditionalLiteral $script:Literal `
                    -Exception @($exception)
            }) | Should Be $true
    }

    It 'rejects an incomplete exact host-surface closure' {
        Remove-Item -LiteralPath $script:Target -Force

        (Test-RSNeutralityThrows -Action {
                Invoke-RSTerminalEngineCollectionNeutralityScan `
                    -RepoRoot $script:Fixture `
                    -AdditionalLiteral $script:Literal
            }) | Should Be $true
    }
}
