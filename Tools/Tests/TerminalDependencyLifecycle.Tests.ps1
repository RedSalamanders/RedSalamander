Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$lifecycleModule = Join-Path $repoRoot 'Tools\TerminalEngine\TerminalEngineLifecycle.psm1'
$peModule = Join-Path $repoRoot 'Tools\TerminalEngine\TerminalEnginePeIdentity.psm1'
Import-Module $lifecycleModule -Force
Import-Module $peModule -Force
Import-Module (Join-Path $repoRoot 'Tools\TerminalEvidence.psm1') -Force

function Write-TestJson {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [object]$Value
    )

    [void](New-Item -ItemType Directory -Path (Split-Path -Parent $Path) -Force)
    $compatible = $Value |
        ConvertTo-Json -Depth 100 -Compress |
        ConvertFrom-Json -AsHashtable -Depth 100 -NoEnumerate -DateKind String
    [IO.File]::WriteAllBytes(
        $Path,
        ([RedSalamander.Tools.TerminalEvidenceJcs]::Serialize($compatible)))
}

function Get-TestRelativePath {
    param([Parameter(Mandatory)][string]$Path)

    return [IO.Path]::GetRelativePath($repoRoot, $Path).Replace('\', '/')
}

function Get-TestEmptySha256 {
    $algorithm = [Security.Cryptography.SHA256]::Create()
    try {
        return [Convert]::ToHexString(
            $algorithm.ComputeHash([byte[]]::new(0))).ToLowerInvariant()
    }
    finally {
        $algorithm.Dispose()
    }
}

function Add-TestInputToTerminalCache {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $sha256 = Get-RSSha256 -Path $Path
    $cachePath = Get-RSTerminalInputCachePath -RepoRoot $repoRoot -Sha256 $sha256
    [void](New-Item -ItemType Directory -Path (Split-Path -Parent $cachePath) -Force)
    if (-not (Test-Path -LiteralPath $cachePath -PathType Leaf)) {
        [IO.File]::Copy($Path, $cachePath, $false)
    }
    return $cachePath
}

function New-TestTerminalLock {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Pin,

        [Parameter(Mandatory)]
        [string]$SourcePath,

        [Parameter(Mandatory)]
        [string]$ToolX64Path,

        [Parameter(Mandatory)]
        [string]$ToolArm64Path,

        [Parameter(Mandatory)]
        [string]$FeedPath,

        [Parameter(Mandatory)]
        [string]$BasedOnLockSha256
    )

    $sourceSha = Get-RSSha256 -Path $SourcePath
    $toolSha = Get-RSSha256 -Path $ToolX64Path
    $toolSemantic = Get-RSPeSemanticSha256 -Path $ToolX64Path
    $toolSize = (Get-Item -LiteralPath $ToolX64Path).Length.ToString(
        [Globalization.CultureInfo]::InvariantCulture)
    $sourceSize = (Get-Item -LiteralPath $SourcePath).Length.ToString(
        [Globalization.CultureInfo]::InvariantCulture)
    $emptySha = Get-TestEmptySha256
    [void](Add-TestInputToTerminalCache -Path $SourcePath)
    [void](Add-TestInputToTerminalCache -Path $ToolX64Path)
    [void](Add-TestInputToTerminalCache -Path $ToolArm64Path)
    $smoke = [ordered]@{
        executableInputId = 'tool-x64'
        arguments = @('/d', '/c', 'exit 0')
        timeoutSeconds = 30
        expectedStdoutSha256 = $emptySha
        expectedStderrSha256 = $emptySha
    }
    $outputs = @(
        [ordered]@{
            id = 'engine-x64-debug'
            platform = 'x64'
            configuration = 'Debug'
            kind = 'pe'
            builtRelativePath = 'tool/cmd.exe'
            relativePath = 'ghostty-vt.dll'
            rawSha256Policy = 'exact'
            sha256 = $toolSha
            semanticSha256 = $toolSemantic
            machine = 'AMD64'
            requiredExports = @()
            smoke = $smoke
        },
        [ordered]@{
            id = 'engine-x64-release'
            platform = 'x64'
            configuration = 'Release'
            kind = 'pe'
            builtRelativePath = 'tool/cmd.exe'
            relativePath = 'ghostty-vt.dll'
            rawSha256Policy = 'exact'
            sha256 = $toolSha
            semanticSha256 = $toolSemantic
            machine = 'AMD64'
            requiredExports = @()
            smoke = $smoke
        },
        [ordered]@{
            id = 'engine-x64-asan'
            platform = 'x64'
            configuration = 'ASan Debug'
            kind = 'pe'
            sourceOutputId = 'engine-x64-debug'
            relativePath = 'ghostty-vt.dll'
            rawSha256Policy = 'exact'
            sha256 = $toolSha
            semanticSha256 = $toolSemantic
            machine = 'AMD64'
            requiredExports = @()
            smoke = $smoke
        },
        [ordered]@{
            id = 'engine-arm64-debug'
            platform = 'ARM64'
            configuration = 'Debug'
            kind = 'pe'
            builtRelativePath = 'tool/cmd.exe'
            relativePath = 'ghostty-vt.dll'
            rawSha256Policy = 'exact'
            sha256 = $toolSha
            semanticSha256 = $toolSemantic
            machine = 'ARM64'
            requiredExports = @()
        },
        [ordered]@{
            id = 'engine-arm64-release'
            platform = 'ARM64'
            configuration = 'Release'
            kind = 'pe'
            builtRelativePath = 'tool/cmd.exe'
            relativePath = 'ghostty-vt.dll'
            rawSha256Policy = 'exact'
            sha256 = $toolSha
            semanticSha256 = $toolSemantic
            machine = 'ARM64'
            requiredExports = @()
        })
    $lanes = @(
        [ordered]@{
            platform = 'x64'
            configuration = 'Debug'
            reuseFrom = $null
            materializations = @([ordered]@{
                    inputId = 'tool-x64'
                    mode = 'copy-file'
                    destination = 'tool/cmd.exe'
                })
            executableRelativePath = 'tool/cmd.exe'
            arguments = @('/d', '/c', 'exit 0')
            timeoutSeconds = 30
            outputIds = @('engine-x64-debug')
        },
        [ordered]@{
            platform = 'x64'
            configuration = 'Release'
            reuseFrom = $null
            materializations = @([ordered]@{
                    inputId = 'tool-x64'
                    mode = 'copy-file'
                    destination = 'tool/cmd.exe'
                })
            executableRelativePath = 'tool/cmd.exe'
            arguments = @('/d', '/c', 'exit 0')
            timeoutSeconds = 30
            outputIds = @('engine-x64-release')
        },
        [ordered]@{
            platform = 'x64'
            configuration = 'ASan Debug'
            reuseFrom = [ordered]@{ platform = 'x64'; configuration = 'Debug' }
            outputIds = @('engine-x64-asan')
        },
        [ordered]@{
            platform = 'ARM64'
            configuration = 'Debug'
            reuseFrom = $null
            materializations = @([ordered]@{
                    inputId = 'tool-arm64'
                    mode = 'copy-file'
                    destination = 'tool/cmd.exe'
                })
            executableRelativePath = 'tool/cmd.exe'
            arguments = @('/d', '/c', 'exit 1')
            outputIds = @('engine-arm64-debug')
        },
        [ordered]@{
            platform = 'ARM64'
            configuration = 'Release'
            reuseFrom = $null
            materializations = @([ordered]@{
                    inputId = 'tool-arm64'
                    mode = 'copy-file'
                    destination = 'tool/cmd.exe'
                })
            executableRelativePath = 'tool/cmd.exe'
            arguments = @('/d', '/c', 'exit 1')
            outputIds = @('engine-arm64-release')
        })
    $lock = [ordered]@{
        formatId = 'red-salamander-terminal-engine-lock'
        schemaVersion = 1
        engine = [ordered]@{
            candidateId = 'ghostty-patched-test'
            pin = $Pin
            pluginLinkage = 'private-dynamic'
            privateRuntimePath = 'Plugins/TerminalRuntime/ghostty-vt.dll'
        }
        inputs = @(
            [ordered]@{
                id = 'engine-source'
                category = 'engine-source'
                kind = 'file'
                source = [ordered]@{
                    type = 'https'
                    uri = "https://fixtures.invalid/$([IO.Path]::GetFileName($SourcePath))"
                }
                sha256 = $sourceSha
                sizeBytes = $sourceSize
            },
            [ordered]@{
                id = 'tool-x64'
                category = 'build-tool'
                kind = 'file'
                source = [ordered]@{
                    type = 'https'
                    uri = 'https://fixtures.invalid/tool-x64.exe'
                }
                sha256 = $toolSha
                sizeBytes = $toolSize
            },
            [ordered]@{
                id = 'tool-arm64'
                category = 'build-tool'
                kind = 'file'
                source = [ordered]@{
                    type = 'https'
                    uri = 'https://fixtures.invalid/tool-arm64.exe'
                }
                sha256 = Get-RSSha256 -Path $ToolArm64Path
                sizeBytes = (Get-Item -LiteralPath $ToolArm64Path).Length.ToString(
                    [Globalization.CultureInfo]::InvariantCulture)
            })
        build = [ordered]@{
            sourceInputId = 'engine-source'
            materializations = @([ordered]@{
                    inputId = 'engine-source'
                    mode = 'copy-file'
                    destination = 'source/source.bin'
                })
            workingDirectoryRelativePath = '.'
            environment = [ordered]@{}
        }
        sourceAdaptation = [ordered]@{
            upstreamTree = [ordered]@{
                algorithm = 'git-sha1'
                digest = ('1' * 40)
            }
            patchSha256 = ('2' * 64)
            resultTree = [ordered]@{
                algorithm = 'git-sha1'
                digest = ('3' * 40)
            }
            adaptedLineCount = '100'
            adaptedFileCount = '2'
            concernMappingSha256 = ('4' * 64)
        }
        derivedInputs = @([ordered]@{
                id = 'engine-source-file-set'
                kind = 'extracted-file-set'
                inputIds = @('engine-source')
                sha256 = ('5' * 64)
                fileCount = '1'
            })
        installedToolchains = @([ordered]@{
                id = 'test-msvc'
                kind = 'visual-studio-msvc'
                version = 'test'
                identitySha256 = ('6' * 64)
                platforms = @('x64', 'ARM64')
            })
        lanes = $lanes
        outputs = $outputs
        dependencies = @([ordered]@{
                id = 'ghostty'
                pin = $Pin
                inputIds = @('engine-source')
                license = [ordered]@{
                    spdx = 'MIT'
                    licenseInputIds = @('engine-source')
                    noticeInputIds = @()
                }
                updateDiscovery = @([ordered]@{
                        id = 'upstream-head'
                        kind = 'file-text'
                        uri = Get-TestRelativePath -Path $FeedPath
                        currentValue = $Pin
                    })
            })
        abi = [ordered]@{
            headerSetSha256 = $sourceSha
            enumIdentity = '0000000000000001'
            typeLayoutIdentity = '0000000000000002'
            callingConvention = 'c'
            capabilities = @('vt-parser', 'resource-limits')
        }
        runtimeClosure = @($outputs | ForEach-Object {
                [ordered]@{
                    platform = $_.platform
                    configuration = $_.configuration
                    module = 'ghostty-vt.dll'
                    system = $false
                    outputId = $_.id
                    loadSource = 'explicit-load'
                    loadedBy = '$product'
                }
            })
        qualification = [ordered]@{
            security = [ordered]@{ status = 'pass' }
            conformance = [ordered]@{ status = 'pass' }
            fuzz = [ordered]@{ status = 'pass' }
            performance = [ordered]@{ status = 'pass' }
            memory = [ordered]@{ status = 'pass' }
            packageSize = [ordered]@{ status = 'pass' }
        }
        upgrade = [ordered]@{
            basedOnLockSha256 = $BasedOnLockSha256
            targetSourceSha256 = $sourceSha
            persistedStateSchemaChanged = $false
            requiredChanges = @(
                'source',
                'submodules',
                'dependencySbom',
                'abi')
        }
    }
    Write-TestJson -Path $Path -Value $lock
    return $lock
}

function Get-TestUpgradeSnapshot {
    param([Parameter(Mandatory)][object]$Lock)

    $sourceInput = @($Lock.inputs | Where-Object { $_.id -eq 'engine-source' })[0]
    $tools = @($Lock.inputs | Where-Object { $_.category -eq 'build-tool' } |
            ForEach-Object { [ordered]@{ id = $_.id; sha256 = $_.sha256 } })
    return [ordered]@{
        source = Get-RSCanonicalObjectSha256 -InputObject ([ordered]@{
                pin = [string]$Lock.engine.pin
                sourceInputId = 'engine-source'
                sourceSha256 = [string]$sourceInput.sha256
                adaptation = $Lock.sourceAdaptation
                derivedInputs = @($Lock.derivedInputs)
            })
        submodules = Get-RSCanonicalObjectSha256 -InputObject @(
            $Lock.dependencies | ForEach-Object {
                [ordered]@{ id = $_.id; pin = $_.pin }
            })
        tools = Get-RSCanonicalObjectSha256 -InputObject ([ordered]@{
                cachedInputs = $tools
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
        security = Get-RSCanonicalObjectSha256 -InputObject $Lock.qualification.security
        conformance = Get-RSCanonicalObjectSha256 -InputObject $Lock.qualification.conformance
        fuzz = Get-RSCanonicalObjectSha256 -InputObject $Lock.qualification.fuzz
        performance = Get-RSCanonicalObjectSha256 -InputObject $Lock.qualification.performance
        memory = Get-RSCanonicalObjectSha256 -InputObject $Lock.qualification.memory
        packageSize = Get-RSCanonicalObjectSha256 -InputObject $Lock.qualification.packageSize
    }
}

function New-TestUpgradeLaneManifestPair {
    param(
        [Parameter(Mandatory)]
        [string]$Root,

        [Parameter(Mandatory)]
        [string]$Role,

        [Parameter(Mandatory)]
        [string]$Platform,

        [Parameter(Mandatory)]
        [string]$Configuration,

        [Parameter(Mandatory)]
        [string]$LockPath,

        [Parameter(Mandatory)]
        [string]$LockSha256,

        [Parameter(Mandatory)]
        [string]$Pin
    )

    $laneRoot = Join-Path (
        Join-Path (
            Join-Path $Root $Role) $Platform) $Configuration
    $buildPath = Join-Path $laneRoot 'terminal-engine-build-manifest.json'
    $verifyPath = Join-Path $laneRoot 'terminal-engine-verify-manifest.json'
    $offline = [ordered]@{
        requested = $true
        provider = 'test-isolation'
        providerStatus = 'pass'
        isolationVerified = $true
        descendantCoverage = 'test-complete'
        rootStdoutSha256 = ('0' * 64)
        rootStderrSha256 = ('0' * 64)
    }
    $lockRecord = [ordered]@{
        path = Get-TestRelativePath -Path $LockPath
        sha256 = $LockSha256
        candidateId = "test-$Role"
        pin = $Pin
    }
    $outputRecord = [ordered]@{
        id = "$Role-$($Platform.ToLowerInvariant())-$($Configuration.Replace(' ', '-').ToLowerInvariant())"
        sha256 = ('a' * 64)
    }
    Write-TestJson -Path $buildPath -Value ([ordered]@{
            formatId = 'red-salamander-terminal-engine-build'
            schemaVersion = 1
            createdUtc = '2026-07-30T00:00:00Z'
            lock = $lockRecord
            platform = $Platform
            configuration = $Configuration
            clean = $true
            offline = $offline
            inputs = @()
            command = [ordered]@{ reused = $false }
            outputs = @($outputRecord)
        })
    $buildSha256 = Get-RSSha256 -Path $buildPath
    $nativeArchitecture = if ($Platform -ceq 'x64') { 'X64' } else { 'Arm64' }
    Write-TestJson -Path $verifyPath -Value ([ordered]@{
            formatId = 'red-salamander-terminal-engine-verification'
            schemaVersion = 1
            createdUtc = '2026-07-30T00:00:00Z'
            lock = $lockRecord
            platform = $Platform
            configuration = $Configuration
            buildManifestSha256 = $buildSha256
            offline = $offline
            runSmoke = $true
            native = [ordered]@{
                process = $nativeArchitecture
                os = $nativeArchitecture
            }
            outputs = @($outputRecord)
            runtimeIdentity = @()
            smoke = @([ordered]@{
                    outputId = $outputRecord.id
                    exitCode = 0
                    stdoutSha256 = ('0' * 64)
                    stderrSha256 = ('0' * 64)
                    verified = $true
                })
            status = 'pass'
        })
    return [ordered]@{
        buildManifest = [ordered]@{
            path = Get-TestRelativePath -Path $buildPath
            sha256 = $buildSha256
        }
        verifyManifest = [ordered]@{
            path = Get-TestRelativePath -Path $verifyPath
            sha256 = Get-RSSha256 -Path $verifyPath
        }
    }
}

Describe 'Terminal engine dependency lifecycle' {
    BeforeAll {
        $script:testRoot = Join-Path $repoRoot (
            '.build\TerminalLifecycleTests\' + [Guid]::NewGuid().ToString('N'))
        $script:inputRoot = Join-Path $script:testRoot 'inputs'
        [void](New-Item -ItemType Directory -Path $script:inputRoot -Force)
        $systemCmd = Join-Path $env:SystemRoot 'System32\cmd.exe'
        $script:toolX64 = Join-Path $script:inputRoot 'tool-x64.exe'
        $script:toolArm64 = Join-Path $script:inputRoot 'tool-arm64.exe'
        [IO.File]::Copy($systemCmd, $script:toolX64)
        [IO.File]::Copy($systemCmd, $script:toolArm64)
        $script:currentSource = Join-Path $script:inputRoot 'source-current.bin'
        $script:candidateSource = Join-Path $script:inputRoot 'source-candidate.bin'
        [IO.File]::WriteAllText(
            $script:currentSource,
            'current-source',
            [Text.UTF8Encoding]::new($false))
        [IO.File]::WriteAllText(
            $script:candidateSource,
            'candidate-source',
            [Text.UTF8Encoding]::new($false))
        $script:feed = Join-Path $script:testRoot 'fake-upstream.txt'
        [IO.File]::WriteAllText(
            $script:feed,
            'pin-next',
            [Text.UTF8Encoding]::new($false))

        $script:currentLockPath = Join-Path $script:testRoot 'current.lock.json'
        $script:currentLockObject = New-TestTerminalLock `
            -Path $script:currentLockPath `
            -Pin 'pin-current' `
            -SourcePath $script:currentSource `
            -ToolX64Path $script:toolX64 `
            -ToolArm64Path $script:toolArm64 `
            -FeedPath $script:feed `
            -BasedOnLockSha256 ('0' * 64)
        $script:currentLockSha = Get-RSSha256 -Path $script:currentLockPath
        $script:currentLockObject = Read-RSTerminalEngineLock `
            -LockFile $script:currentLockPath
        $script:candidateLockPath = Join-Path $script:testRoot 'candidate.lock.json'
        $script:candidateLockObject = New-TestTerminalLock `
            -Path $script:candidateLockPath `
            -Pin 'pin-next' `
            -SourcePath $script:candidateSource `
            -ToolX64Path $script:toolX64 `
            -ToolArm64Path $script:toolArm64 `
            -FeedPath $script:feed `
            -BasedOnLockSha256 $script:currentLockSha
        $script:candidateLockSha = Get-RSSha256 -Path $script:candidateLockPath
        $script:candidateLockObject = Read-RSTerminalEngineLock `
            -LockFile $script:candidateLockPath
        $script:candidateSourceSha = Get-RSSha256 -Path $script:candidateSource
        $script:buildRoot = Join-Path $script:testRoot 'build'
    }

    AfterAll {
        $expectedRoot = [IO.Path]::GetFullPath(
            (Join-Path $repoRoot '.build\TerminalLifecycleTests'))
        $resolved = [IO.Path]::GetFullPath($script:testRoot)
        if (-not $resolved.StartsWith(
                $expectedRoot + [IO.Path]::DirectorySeparatorChar,
                [StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing to clean unexpected test root '$resolved'."
        }
        if (Test-Path -LiteralPath $resolved) {
            Remove-Item -LiteralPath $resolved -Recurse -Force
        }
    }

    It 'validates a synthetic private-dynamic lock and platform-specific tool inputs' {
        $lock = Read-RSTerminalEngineLock `
            -LockFile $script:currentLockPath `
            -ValidateInputs

        $lock.engine.pluginLinkage | Should Be 'private-dynamic'
        @($lock.verifiedInputs).Count | Should Be 3
        [string](Get-RSTerminalLane `
            -Lock $lock `
            -Platform x64 `
            -Configuration Release).materializations[0].inputId |
            Should Be 'tool-x64'
        [string](Get-RSTerminalLane `
            -Lock $lock `
            -Platform ARM64 `
            -Configuration Release).materializations[0].inputId |
            Should Be 'tool-arm64'
    }

    It 'keeps the lock schema closed and accepts the frozen mutation corpus' {
        $schemaPath = Join-Path $repoRoot 'Specs\Terminal\TerminalEngineLock.schema.json'
        $corpusPath = Join-Path $repoRoot 'Specs\Terminal\TerminalEngineLockCorpus.v1.json'
        $schema = Get-Content -LiteralPath $schemaPath -Raw |
            ConvertFrom-Json -Depth 100
        $corpus = Get-Content -LiteralPath $corpusPath -Raw |
            ConvertFrom-Json -Depth 100
        $schema.additionalProperties | Should Be $false
        $schema.'$schema' | Should Be 'https://json-schema.org/draft/2020-12/schema'
        $corpus.schema | Should Be 'Specs/Terminal/TerminalEngineLock.schema.json'
        @($corpus.cases).Count | Should Be 7
        (Get-Content -LiteralPath $script:currentLockPath -Raw |
                Test-Json -SchemaFile $schemaPath) | Should Be $true

        foreach ($case in @($corpus.cases)) {
            $candidate = Get-Content -LiteralPath $script:currentLockPath -Raw |
                ConvertFrom-Json -Depth 100
            switch ([string]$case.operation) {
                'none' {}
                'add' {
                    $candidate | Add-Member `
                        -NotePropertyName ([string]$case.path) `
                        -NotePropertyValue $case.value
                }
                'replace-source-with-path' {
                    $candidate.inputs[0] | Add-Member `
                        -NotePropertyName path `
                        -NotePropertyValue ([string]$case.value)
                    $candidate.inputs[0].PSObject.Properties.Remove('source')
                }
                'replace' {
                    if ([string]$case.path -eq 'inputs[0].source.uri') {
                        $candidate.inputs[0].source.uri = [string]$case.value
                    }
                    elseif ([string]$case.path -eq 'derivedInputs[0].inputIds[0]') {
                        $candidate.derivedInputs[0].inputIds[0] = [string]$case.value
                    }
                    else {
                        throw "Unknown lock-corpus replace path '$($case.path)'."
                    }
                }
                'remove-lane' {
                    $candidate.lanes = @($candidate.lanes | Where-Object {
                            -not (
                                [string]$_.platform -ceq 'x64' -and
                                [string]$_.configuration -ceq 'ASan Debug')
                        })
                }
                default {
                    throw "Unknown lock-corpus operation '$($case.operation)'."
                }
            }
            $casePath = Join-Path $script:testRoot "corpus-$($case.id).json"
            Write-TestJson -Path $casePath -Value $candidate
            $accepted = $true
            $message = ''
            try {
                Read-RSTerminalEngineLock -LockFile $casePath | Out-Null
            }
            catch {
                $accepted = $false
                $message = $_.Exception.Message
            }
            $accepted | Should Be ([bool]$case.valid)
            if (-not [bool]$case.valid) {
                $message | Should Match ([string]$case.errorPattern)
            }
        }

        $canonicalBytes = [IO.File]::ReadAllBytes($script:currentLockPath)
        $canonicalText = [Text.UTF8Encoding]::new($false, $true).GetString(
            $canonicalBytes)
        $malformedCases = @(
            [ordered]@{
                id = 'bom'
                bytes = [byte[]]@(0xef, 0xbb, 0xbf) + $canonicalBytes
                pattern = 'without BOM'
            },
            [ordered]@{
                id = 'noncanonical-whitespace'
                bytes = [Text.UTF8Encoding]::new($false).GetBytes(
                    $canonicalText + "`n")
                pattern = 'exact RFC 8785 JCS'
            },
            [ordered]@{
                id = 'duplicate-property'
                bytes = [Text.UTF8Encoding]::new($false).GetBytes(
                    $canonicalText.Insert(
                        1,
                        '"formatId":"red-salamander-terminal-engine-lock",'))
                pattern = 'exact RFC 8785 JCS'
            },
            [ordered]@{
                id = 'invalid-utf8'
                bytes = [byte[]]@(0xff)
                pattern = 'duplicate-free UTF-8 JSON'
            })
        foreach ($case in $malformedCases) {
            $casePath = Join-Path $script:testRoot "raw-$($case.id).json"
            [IO.File]::WriteAllBytes($casePath, [byte[]]$case.bytes)
            {
                Read-RSTerminalEngineLock -LockFile $casePath
            } | Should Throw ([string]$case.pattern)
        }
    }

    It 'distinguishes durable repository files from content-addressed HTTPS cache inputs' {
        $repositoryLockPath = Join-Path $script:testRoot 'repository-source.lock.json'
        $repositoryLock = Get-Content -LiteralPath $script:currentLockPath -Raw |
            ConvertFrom-Json -Depth 100
        $repositoryInputPath = Join-Path (
            Join-Path $repoRoot 'Tools\Tests') 'TerminalDependencyLifecycle.Tests.ps1'
        $repositoryLock.inputs[0].source = [pscustomobject][ordered]@{
            type = 'repository-file'
            path = Get-TestRelativePath -Path $repositoryInputPath
        }
        $repositoryLock.inputs[0].sha256 = Get-RSSha256 -Path $repositoryInputPath
        $repositoryLock.inputs[0].sizeBytes = (
            Get-Item -LiteralPath $repositoryInputPath).Length.ToString(
            [Globalization.CultureInfo]::InvariantCulture)
        Write-TestJson -Path $repositoryLockPath -Value $repositoryLock

        $validated = Read-RSTerminalEngineLock `
            -LockFile $repositoryLockPath `
            -ValidateInputs
        $validated.verifiedInputs[0].sourceType | Should Be 'repository-file'
        $cachePath = Get-RSTerminalInputCachePath `
            -RepoRoot $repoRoot `
            -Sha256 ([string]$script:currentLockObject.inputs[0].sha256)
        $cachePath | Should Be (
            Join-Path (
                Join-Path $repoRoot '.build\TerminalEngineInputCache\v1\sha256') (
                [string]$script:currentLockObject.inputs[0].sha256))

        $moduleSource = Get-Content -LiteralPath $lifecycleModule -Raw
        $restoreSource = Get-Content -LiteralPath (
            Join-Path $repoRoot 'Tools\Restore-TerminalEngineInputs.ps1') -Raw
        $moduleSource | Should Match 'AllowAutoRedirect = \$false'
        $moduleSource | Should Match 'UseDefaultCredentials = \$false'
        $moduleSource | Should Match 'UseProxy = \$false'
        $moduleSource | Should Match '\[int\]\$response\.StatusCode -ne 200'
        $moduleSource | Should Match (
            'CancellationTokenSource\]::new\(\s*' +
            '\[TimeSpan\]::FromMinutes\(10\)\)')
        $moduleSource | Should Match (
            'ResponseHeadersRead,\s*\$downloadCancellation\.Token')
        $moduleSource | Should Match (
            'ReadAsync\([\s\S]*?\$downloadCancellation\.Token\)')
        $moduleSource | Should Match (
            "HTTPS input '\`$id' timed out after 600 seconds")
        $restoreSource | Should Match 'Restore-RSTerminalEngineInputs'
    }

    It 'rejects incomplete dependency provenance and invalid recursive load declarations' {
        $invalidLockPath = Join-Path $script:testRoot 'invalid-lifecycle.lock.json'
        $cases = @(
            [ordered]@{
                Name = 'duplicate dependency id'
                Pattern = 'Duplicate dependency id'
                Mutate = {
                    param($Lock)
                    $duplicate = $Lock.dependencies[0] |
                        ConvertTo-Json -Depth 100 |
                        ConvertFrom-Json -Depth 100
                    $duplicate.id = 'GHOSTTY'
                    $Lock.dependencies = @(
                        $Lock.dependencies[0],
                        $duplicate)
                }
            },
            [ordered]@{
                Name = 'empty dependency pin'
                Pattern = 'pin may not be empty'
                Mutate = {
                    param($Lock)
                    $Lock.dependencies[0].pin = ''
                }
            },
            [ordered]@{
                Name = 'unknown dependency input'
                Pattern = 'unknown input'
                Mutate = {
                    param($Lock)
                    $Lock.dependencies[0].inputIds = @('missing-input')
                }
            },
            [ordered]@{
                Name = 'duplicate dependency input'
                Pattern = 'empty or duplicate inputId'
                Mutate = {
                    param($Lock)
                    $Lock.dependencies[0].inputIds = @(
                        'engine-source',
                        'engine-source')
                }
            },
            [ordered]@{
                Name = 'empty SPDX'
                Pattern = 'no SPDX identity'
                Mutate = {
                    param($Lock)
                    $Lock.dependencies[0].license.spdx = ' '
                }
            },
            [ordered]@{
                Name = 'missing licenseInputIds'
                Pattern = 'licenseInputIds must be explicitly declared'
                Mutate = {
                    param($Lock)
                    $Lock.dependencies[0].license.PSObject.Properties.Remove(
                        'licenseInputIds')
                }
            },
            [ordered]@{
                Name = 'empty licenseInputIds'
                Pattern = 'licenseInputIds may not be empty'
                Mutate = {
                    param($Lock)
                    $Lock.dependencies[0].license.licenseInputIds = @()
                }
            },
            [ordered]@{
                Name = 'license input outside dependency inputs'
                Pattern = 'must refer exactly'
                Mutate = {
                    param($Lock)
                    $Lock.dependencies[0].license.licenseInputIds = @('tool-x64')
                }
            },
            [ordered]@{
                Name = 'notice input outside dependency inputs'
                Pattern = 'must refer exactly'
                Mutate = {
                    param($Lock)
                    $Lock.dependencies[0].license.noticeInputIds = @('tool-arm64')
                }
            },
            [ordered]@{
                Name = 'unsupported load source'
                Pattern = 'unsupported loadSource'
                Mutate = {
                    param($Lock)
                    $Lock.runtimeClosure[0].loadSource = 'implicit-magic'
                }
            },
            [ordered]@{
                Name = 'missing recursive parent'
                Pattern = 'loaded by missing parent'
                Mutate = {
                    param($Lock)
                    $Lock.runtimeClosure[0].loadedBy = 'missing-parent.dll'
                }
            },
            [ordered]@{
                Name = 'recursive load cycle'
                Pattern = 'contains a load cycle'
                Mutate = {
                    param($Lock)
                    $Lock.runtimeClosure[0].loadedBy = 'ghostty-vt.dll'
                }
            })
        foreach ($case in $cases) {
            $invalid = Get-Content -LiteralPath $script:currentLockPath -Raw |
                ConvertFrom-Json -Depth 100
            & $case.Mutate $invalid
            Write-TestJson -Path $invalidLockPath -Value $invalid
            $rejected = $false
            $message = ''
            try {
                Read-RSTerminalEngineLock -LockFile $invalidLockPath |
                    Out-Null
            }
            catch {
                $rejected = $true
                $message = $_.Exception.Message
            }
            $rejected | Should Be $true
            $message | Should Match $case.Pattern
        }
    }

    It 'accepts source-static only with no private path or app-local runtime leaf' {
        $staticLockPath = Join-Path $script:testRoot 'source-static.lock.json'
        $staticLock = Get-Content -LiteralPath $script:currentLockPath -Raw |
            ConvertFrom-Json -Depth 100
        $staticLock.engine.pluginLinkage = 'source-static'
        $staticLock.engine.privateRuntimePath = $null
        $staticLock.runtimeClosure = @()
        Write-TestJson -Path $staticLockPath -Value $staticLock

        $validated = Read-RSTerminalEngineLock -LockFile $staticLockPath
        $validated.engine.pluginLinkage | Should Be 'source-static'
        $null -eq $validated.engine.privateRuntimePath | Should Be $true

        $staticLock.engine.privateRuntimePath =
            'Plugins/TerminalRuntime/forbidden.dll'
        Write-TestJson -Path $staticLockPath -Value $staticLock
        $pathRejected = $false
        try {
            Read-RSTerminalEngineLock -LockFile $staticLockPath | Out-Null
        }
        catch {
            $pathRejected = $true
        }
        $pathRejected | Should Be $true

        $staticLock.engine.privateRuntimePath = $null
        $staticLock.runtimeClosure = @([ordered]@{
                platform = 'x64'
                configuration = 'Release'
                module = 'forbidden.dll'
                system = $false
                outputId = 'engine-x64-release'
            })
        Write-TestJson -Path $staticLockPath -Value $staticLock
        $leafRejected = $false
        try {
            Read-RSTerminalEngineLock -LockFile $staticLockPath | Out-Null
        }
        catch {
            $leafRejected = $true
        }
        $leafRejected | Should Be $true
    }

    It 'builds and verifies the exact requested candidate lock under .build' {
        $buildManifest = & (Join-Path $repoRoot 'Tools\Build-TerminalEngine.ps1') `
            -Platform x64 `
            -Configuration Release `
            -LockFile $script:candidateLockPath `
            -OutputRoot $script:buildRoot `
            -Clean `
            -PassThruManifest
        $verifyManifest = & (Join-Path $repoRoot 'Tools\Verify-TerminalEngine.ps1') `
            -Platform x64 `
            -Configuration Release `
            -LockFile $script:candidateLockPath `
            -OutputRoot $script:buildRoot `
            -RunSmoke `
            -PassThruManifest

        (Test-Path -LiteralPath $buildManifest -PathType Leaf) | Should Be $true
        (Test-Path -LiteralPath $verifyManifest -PathType Leaf) | Should Be $true
        $build = Get-Content -LiteralPath $buildManifest -Raw | ConvertFrom-Json
        $verify = Get-Content -LiteralPath $verifyManifest -Raw | ConvertFrom-Json
        $build.lock.sha256 | Should Be $script:candidateLockSha
        $verify.lock.sha256 | Should Be $script:candidateLockSha
        $verify.status | Should Be 'pass'
        @($verify.smoke).Count | Should Be 1
    }

    It 'materializes the explicit x64 ASan-to-Debug reuse rule at the ASan path' {
        $manifestPath = & (Join-Path $repoRoot 'Tools\Build-TerminalEngine.ps1') `
            -Platform x64 `
            -Configuration 'ASan Debug' `
            -LockFile $script:currentLockPath `
            -OutputRoot $script:buildRoot `
            -Clean `
            -PassThruManifest
        $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
        $debugDll = Join-Path $script:buildRoot 'x64\Debug\ghostty-vt.dll'
        $asanDll = Join-Path $script:buildRoot 'x64\ASan Debug\ghostty-vt.dll'

        $manifest.command.reused | Should Be $true
        (Get-RSSha256 -Path $asanDll) | Should Be (Get-RSSha256 -Path $debugDll)
    }

    It 'fails closed for unsafe roots, input drift, and offline use without runner acknowledgement' {
        $unsafeRootRejected = $false
        try {
            Assert-RSTerminalSafeBuildRoot `
                -RepoRoot $repoRoot `
                -BuildRoot $repoRoot
        }
        catch {
            $unsafeRootRejected = $true
        }
        $unsafeRootRejected | Should Be $true

        $buildSource = Get-Content -LiteralPath (
            Join-Path $repoRoot 'Tools\Build-TerminalEngine.ps1') -Raw
        $buildSource | Should Match '\-Offline requires \-ExclusiveDisposableRunner'

        $currentSourceCache = Get-RSTerminalInputCachePath `
            -RepoRoot $repoRoot `
            -Sha256 (Get-RSSha256 -Path $script:currentSource)
        $original = [IO.File]::ReadAllBytes($currentSourceCache)
        try {
            [IO.File]::AppendAllText($currentSourceCache, 'drift')
            $driftRejected = $false
            try {
                Read-RSTerminalEngineLock `
                    -LockFile $script:currentLockPath `
                    -ValidateInputs
            }
            catch {
                $driftRejected = $true
            }
            $driftRejected | Should Be $true
        }
        finally {
            [IO.File]::WriteAllBytes($currentSourceCache, $original)
        }
    }

    It 'separates structural locked status from explicit cache verification' {
        $sourceSha = Get-RSSha256 -Path $script:currentSource
        $cachePath = Get-RSTerminalInputCachePath -RepoRoot $repoRoot -Sha256 $sourceSha
        $backupPath = Join-Path $script:testRoot 'source-cache-backup.bin'
        [IO.File]::Move($cachePath, $backupPath)
        try {
            & (Join-Path $repoRoot 'Tools\Get-TerminalDependencyStatus.ps1') `
                -Mode Locked `
                -EngineLock $script:currentLockPath
            $rejected = $false
            try {
                & (Join-Path $repoRoot 'Tools\Get-TerminalDependencyStatus.ps1') `
                    -Mode Locked `
                    -EngineLock $script:currentLockPath `
                    -RequireCache
            }
            catch {
                $rejected = $true
            }
            $rejected | Should Be $true

            $offlineRejected = $false
            try {
                & (Join-Path $repoRoot 'Tools\Restore-TerminalEngineInputs.ps1') `
                    -LockFile $script:currentLockPath `
                    -Offline
            }
            catch {
                $offlineRejected = $true
            }
            $offlineRejected | Should Be $true
        }
        finally {
            [IO.File]::Move($backupPath, $cachePath)
        }
        @(& (Join-Path $repoRoot 'Tools\Restore-TerminalEngineInputs.ps1') `
                -LockFile $script:currentLockPath `
                -Offline `
                -PassThru).Count | Should Be 3
    }

    It 'keeps locked status read-only and emits an advisory-only local online report' {
        $statusRoot = Join-Path $script:testRoot 'status'
        & (Join-Path $repoRoot 'Tools\Get-TerminalDependencyStatus.ps1') `
            -Mode Locked `
            -EngineLock $script:currentLockPath
        (Test-Path -LiteralPath $statusRoot) | Should Be $false

        $reportPath = & (Join-Path $repoRoot 'Tools\Get-TerminalDependencyStatus.ps1') `
            -Mode Online `
            -EngineLock $script:currentLockPath `
            -PassThruReport `
            -ReportRoot $statusRoot
        $report = Get-Content -LiteralPath $reportPath -Raw | ConvertFrom-Json
        $statusLock = Read-RSTerminalEngineLock `
            -LockFile $script:currentLockPath `
            -ValidateInputs

        $report.lockedStatus | Should Be 'pass'
        $report.schemaVersion | Should Be 3
        $report.cacheRequired | Should Be $false
        $report.cacheStatus | Should Be 'present-unverified'
        $report.sourceAdaptationSha256 | Should Match '^[0-9a-f]{64}$'
        $report.derivedInputSetSha256 | Should Match '^[0-9a-f]{64}$'
        $report.installedToolchainSetSha256 | Should Match '^[0-9a-f]{64}$'
        $report.licenseClosureSha256 | Should Be (
            Get-RSCanonicalObjectSha256 -InputObject @(
                Get-RSTerminalLicenseClosure -Lock $statusLock))
        $report.noticeClosureSha256 | Should Be (
            Get-RSCanonicalObjectSha256 -InputObject @(
                Get-RSTerminalNoticeClosure -Lock $statusLock))
        $report.runtimeLoadGraphSha256 | Should Match '^[0-9a-f]{64}$'
        $report.runtimeClosureEvidence | Should Be 'declarative-lock-graph'
        @($report.observations).Count | Should Be 1
        $report.observations[0].status | Should Be 'update-observed'
        [IO.File]::ReadAllText($script:feed).Trim() | Should Be 'pin-next'
    }

    It 'reopens all five exact lane manifests and proves previous-lock rollback' {
        $snapshot = [pscustomobject](Get-TestUpgradeSnapshot `
                -Lock $script:candidateLockObject)
        $evidenceRoot = Join-Path $script:testRoot 'upgrade-evidence'
        $tuples = @(
            @{ Platform = 'x64'; Configuration = 'Debug'; Architecture = 'X64' },
            @{ Platform = 'x64'; Configuration = 'Release'; Architecture = 'X64' },
            @{ Platform = 'x64'; Configuration = 'ASan Debug'; Architecture = 'X64' },
            @{ Platform = 'ARM64'; Configuration = 'Debug'; Architecture = 'Arm64' },
            @{ Platform = 'ARM64'; Configuration = 'Release'; Architecture = 'Arm64' })
        $evidencePaths = @()
        foreach ($tuple in $tuples) {
            $candidateManifests = New-TestUpgradeLaneManifestPair `
                -Root $evidenceRoot `
                -Role 'candidate' `
                -Platform $tuple.Platform `
                -Configuration $tuple.Configuration `
                -LockPath $script:candidateLockPath `
                -LockSha256 $script:candidateLockSha `
                -Pin 'pin-next'
            $rollbackManifests = New-TestUpgradeLaneManifestPair `
                -Root $evidenceRoot `
                -Role 'rollback' `
                -Platform $tuple.Platform `
                -Configuration $tuple.Configuration `
                -LockPath $script:currentLockPath `
                -LockSha256 $script:currentLockSha `
                -Pin 'pin-current'
            $path = Join-Path $evidenceRoot (
                "$($tuple.Platform)-$($tuple.Configuration.Replace(' ', '-')).json")
            Write-TestJson -Path $path -Value ([ordered]@{
                    formatId = 'red-salamander-terminal-engine-upgrade-collection'
                    schemaVersion = 2
                    status = 'pass'
                    currentLockSha256 = $script:currentLockSha
                    candidateLockPath = Get-TestRelativePath -Path $script:candidateLockPath
                    candidateLockSha256 = $script:candidateLockSha
                    targetPin = 'pin-next'
                    targetPinDigest = (& {
                            $algorithm = [Security.Cryptography.SHA256]::Create()
                            try {
                                [Convert]::ToHexString($algorithm.ComputeHash(
                                    [Text.UTF8Encoding]::new($false).GetBytes('pin-next'))).
                                    ToLowerInvariant()
                            }
                            finally { $algorithm.Dispose() }
                        })
                    targetSourceSha256 = $script:candidateSourceSha
                    platform = $tuple.Platform
                    configuration = $tuple.Configuration
                    native = [ordered]@{
                        process = $tuple.Architecture
                        os = $tuple.Architecture
                    }
                    policy = [ordered]@{
                        bootstrap = $true
                        clean = $true
                        offline = $true
                        networkIsolation = 'runner-enforced'
                    }
                    candidate = [ordered]@{
                        buildManifest = $candidateManifests.buildManifest
                        verifyManifest = $candidateManifests.verifyManifest
                        snapshot = $snapshot
                    }
                    rollback = [ordered]@{
                        previousLockSha256 = $script:currentLockSha
                        buildManifest = $rollbackManifests.buildManifest
                        verifyManifest = $rollbackManifests.verifyManifest
                        persistedStateSchemaChanged = $false
                        status = 'pass'
                    }
                })
            $evidencePaths += $path
        }
        $upgradeRoot = Join-Path $script:testRoot 'upgrade'
        $comparisonPath = & (Join-Path $repoRoot 'Tools\Prepare-TerminalEngineUpgrade.ps1') `
            -Mode Finalize `
            -TargetPin 'pin-next' `
            -TargetSourceSha256 $script:candidateSourceSha `
            -ExpectedCurrentLockSha256 $script:currentLockSha `
            -EngineLock $script:currentLockPath `
            -EvidenceManifest $evidencePaths `
            -UpgradeRoot $upgradeRoot `
            -RequireAllEvidence `
            -PassThruManifest
        $comparison = Get-Content -LiteralPath $comparisonPath -Raw |
            ConvertFrom-Json -Depth 100

        $comparison.status | Should Be 'pass'
        $comparison.rollback.status | Should Be 'pass'
        $comparison.rollback.previousLockSha256 | Should Be $script:currentLockSha
        $comparison.rollback.everyLaneReconstructed | Should Be $true
        @($comparison.inputs).Count | Should Be 5
        @($comparison.laneEvidence).Count | Should Be 5
        @($comparison.rollback.lanes).Count | Should Be 5
        @($comparison.changedComparisonFields) -join ',' |
            Should Be 'abi,dependencySbom,source,submodules'
        @($comparison.requiredRepositoryChangeCategories) -join ',' |
            Should Be 'abi,dependencySbom,source,submodules'

        $finalizeArguments = @{
            Mode = 'Finalize'
            TargetPin = 'pin-next'
            TargetSourceSha256 = $script:candidateSourceSha
            ExpectedCurrentLockSha256 = $script:currentLockSha
            EngineLock = $script:currentLockPath
            EvidenceManifest = $evidencePaths
            UpgradeRoot = $upgradeRoot
            RequireAllEvidence = $true
        }
        $assertFinalizeRejected = {
            param([Parameter(Mandatory)][string]$Pattern)

            $rejected = $false
            $message = ''
            try {
                & (Join-Path $repoRoot 'Tools\Prepare-TerminalEngineUpgrade.ps1') `
                    @finalizeArguments |
                    Out-Null
            }
            catch {
                $rejected = $true
                $message = $_.Exception.Message
            }
            $rejected | Should Be $true
            $message | Should Match $Pattern
        }
        $finalizeArguments.EvidenceManifest = @($evidencePaths[0..3])
        & $assertFinalizeRejected 'exactly five'
        $finalizeArguments.EvidenceManifest = $evidencePaths

        $originalCandidateLock = Get-Content `
            -LiteralPath $script:candidateLockPath `
            -Raw
        $mismatchedLock = $originalCandidateLock |
            ConvertFrom-Json -Depth 100
        $mismatchedLock.upgrade.requiredChanges = @('source')
        Write-TestJson `
            -Path $script:candidateLockPath `
            -Value $mismatchedLock
        $mismatchedLockSha = Get-RSSha256 -Path $script:candidateLockPath
        foreach ($evidencePath in $evidencePaths) {
            $record = Get-Content -LiteralPath $evidencePath -Raw |
                ConvertFrom-Json -Depth 100
            $candidateBuildPath = Join-Path `
                $repoRoot `
                ([string]$record.candidate.buildManifest.path)
            $candidateVerifyPath = Join-Path `
                $repoRoot `
                ([string]$record.candidate.verifyManifest.path)
            $candidateBuild = Get-Content `
                -LiteralPath $candidateBuildPath `
                -Raw |
                ConvertFrom-Json -Depth 100
            $candidateBuild.lock.sha256 = $mismatchedLockSha
            Write-TestJson -Path $candidateBuildPath -Value $candidateBuild
            $candidateBuildSha = Get-RSSha256 -Path $candidateBuildPath
            $candidateVerify = Get-Content `
                -LiteralPath $candidateVerifyPath `
                -Raw |
                ConvertFrom-Json -Depth 100
            $candidateVerify.lock.sha256 = $mismatchedLockSha
            $candidateVerify.buildManifestSha256 = $candidateBuildSha
            Write-TestJson -Path $candidateVerifyPath -Value $candidateVerify
            $record.candidateLockSha256 = $mismatchedLockSha
            $record.candidate.buildManifest.sha256 = $candidateBuildSha
            $record.candidate.verifyManifest.sha256 =
                Get-RSSha256 -Path $candidateVerifyPath
            Write-TestJson -Path $evidencePath -Value $record
        }
        & $assertFinalizeRejected 'requiredChanges must exactly equal'

        [IO.File]::WriteAllText(
            $script:candidateLockPath,
            $originalCandidateLock,
            [Text.UTF8Encoding]::new($false))
        foreach ($evidencePath in $evidencePaths) {
            $record = Get-Content -LiteralPath $evidencePath -Raw |
                ConvertFrom-Json -Depth 100
            $candidateManifests = New-TestUpgradeLaneManifestPair `
                -Root $evidenceRoot `
                -Role 'candidate' `
                -Platform ([string]$record.platform) `
                -Configuration ([string]$record.configuration) `
                -LockPath $script:candidateLockPath `
                -LockSha256 $script:candidateLockSha `
                -Pin 'pin-next'
            $record.candidateLockSha256 = $script:candidateLockSha
            $record.candidate.buildManifest =
                $candidateManifests.buildManifest
            $record.candidate.verifyManifest =
                $candidateManifests.verifyManifest
            Write-TestJson -Path $evidencePath -Value $record
        }

        $broken = Get-Content -LiteralPath $evidencePaths[0] -Raw |
            ConvertFrom-Json -Depth 100
        $broken.rollback.status = 'fail'
        Write-TestJson -Path $evidencePaths[0] -Value $broken
        & $assertFinalizeRejected 'rollback proof'

        $broken.rollback.status = 'pass'
        Write-TestJson -Path $evidencePaths[0] -Value $broken
        $rollbackVerifyPath = Join-Path `
            $repoRoot `
            ([string]$broken.rollback.verifyManifest.path)
        $rollbackVerify = Get-Content -LiteralPath $rollbackVerifyPath -Raw |
            ConvertFrom-Json -Depth 100
        $rollbackVerify.status = 'fail'
        Write-TestJson -Path $rollbackVerifyPath -Value $rollbackVerify
        $broken.rollback.verifyManifest.sha256 =
            Get-RSSha256 -Path $rollbackVerifyPath
        Write-TestJson -Path $evidencePaths[0] -Value $broken
        & $assertFinalizeRejected 'rollback .* verify manifest'

        $rollbackVerify.status = 'pass'
        Write-TestJson -Path $rollbackVerifyPath -Value $rollbackVerify
        $broken.rollback.verifyManifest.sha256 =
            Get-RSSha256 -Path $rollbackVerifyPath
        Write-TestJson -Path $evidencePaths[0] -Value $broken
        [IO.File]::AppendAllText(
            $rollbackVerifyPath,
            ' ',
            [Text.UTF8Encoding]::new($false))
        & $assertFinalizeRejected 'digest differs'
    }

    It 'requires an exact candidate lock for collection and documents the selected runtime workflow' {
        $source = Get-Content -LiteralPath (
            Join-Path $repoRoot 'Tools\Prepare-TerminalEngineUpgrade.ps1') -Raw
        $source | Should Match 'Collect requires an exact pre-reviewed \-CandidateLockFile'
        $source | Should Match 'Offline upgrade collection requires \-ExclusiveDisposableRunner'
        $source | Should Match '\-Bootstrap \-Clean \-Offline'
        $source | Should Match 'schemaVersion = 2'
        $source | Should Match 'Get-RSUpgradeManifestReference'
        $source | Should Match '\$EvidenceManifest\.Count -ne 5'
        $source | Should Match 'Test-RSUpgradeLaneManifestPair'
        $source | Should Match 'requiredChanges must exactly equal the derived changed fields'

        $terminalSpec = [IO.File]::ReadAllText(
            (Join-Path $repoRoot 'Specs\Terminal\Terminal_EmbeddedPlugin.md'))
        $decision = [IO.File]::ReadAllText(
            (Join-Path $repoRoot 'Specs\Terminal\TerminalEngineDecision.md'))
        $builder = [IO.File]::ReadAllText(
            (Join-Path $repoRoot 'Tools\Build-GhosttyTerminalRuntime.ps1'))
        $loader = [IO.File]::ReadAllText(
            (Join-Path $repoRoot 'Plugins\Terminal\TerminalRuntimeLoader.cpp'))

        $decision | Should Match 'Private `ghostty-vt\.dll` behind `Terminal\.dll`[\s\S]+\*\*Selected\*\*'
        $terminalSpec | Should Match 'Build-GhosttyTerminalRuntime\.ps1 -Platform x64 -Offline -PassThru'
        $terminalSpec | Should Match 'Build-GhosttyTerminalRuntime\.ps1 -Platform ARM64 -Offline -PassThru'
        $terminalSpec | Should Match '\-Offline \-Force'
        $terminalSpec | Should Match 'generated identity header'
        $terminalSpec | Should Match 'all\s+five notices as one\s+atomic\s+manifest unit'
        $builder | Should Match '\[switch\]\$Offline'
        $builder | Should Match 'schemaVersion\s*=\s*5'
        $builder | Should Match "pairingContract\s*=\s*'generated-raw-sha256-v1'"
        $builder | Should Match 'Get-TerminalRuntimeIdentityHeaderText'
        $builder | Should Match '\$generatedIdentityHeader'
        $builder | Should Match '\$expectedLicenseHashes'
        $loader | Should Match 'BCryptOpenAlgorithmProvider'
        $loader | Should Match 'GetFinalPathNameByHandleW'
    }
}
