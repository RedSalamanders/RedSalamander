$script:RepoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$script:ModulePath = Join-Path $script:RepoRoot (
    'Tools\TerminalEngine\TerminalEngineCollectionDescriptor.psm1')
$script:SchemaPath = Join-Path $script:RepoRoot (
    'Specs\Terminal\TerminalEngineCollectionDescriptor.schema.json')
Import-Module (Join-Path $script:RepoRoot 'Tools\TerminalEvidence.psm1') -Force -Global
Import-Module $script:ModulePath -Force

function New-RSDescriptorFixture {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root,

        [scriptblock]$Mutate = {}
    )

        [void](New-Item -ItemType Directory -Path $Root -Force)
        $licensePath = Join-Path $Root 'LICENSE.txt'
        [IO.File]::WriteAllText(
            $licensePath,
            'synthetic license',
            [Text.UTF8Encoding]::new($false))
        $license = Get-Item -LiteralPath $licensePath -Force
        $descriptor = [ordered]@{
            formatId = 'red-salamander-terminal-engine-collection-descriptor'
            version = 1
            candidateId = 'candidate-alpha'
            headerPackages = [object[]]@()
            legal = [ordered]@{
                closureId = 'closure-alpha'
                outputRelativePath =
                    'red-salamander/generated/third-party-notices.txt'
                dependencies = [object[]]@(
                    [ordered]@{
                        dependencyId = 'dependency-alpha'
                        spdx = 'MIT'
                        inputs = [object[]]@(
                            [ordered]@{
                                inputId = 'license-alpha'
                                kind = 'license'
                                destinationLeaf = 'license-alpha.txt'
                                sha256 = (Get-FileHash `
                                        -LiteralPath $licensePath `
                                        -Algorithm SHA256).Hash.ToLowerInvariant()
                                sizeBytes = $license.Length.ToString(
                                    [Globalization.CultureInfo]::InvariantCulture)
                                source = [ordered]@{
                                    sourceKind = 'candidate-source'
                                    origin = 'upstream'
                                    sourceRelativePath = 'LICENSE.txt'
                                }
                            })
                    })
            }
            build = [ordered]@{
                kind = 'provider-generic-cmake-msvc-v1'
                buildTargets = [string[]]@('adapter-alpha')
                pathBindings = [object[]]@()
                adapter = [ordered]@{
                    dllLeaf = 'adapter-alpha.dll'
                    projectLeaf = 'adapter-alpha.vcxproj'
                }
                staticLibraries = [object[]]@()
                privateRuntimeLibraries = [object[]]@()
            }
        }
        & $Mutate $descriptor
        $descriptorPath = Join-Path $Root (
            'red-salamander\collection-descriptor.v1.json')
        [void](New-Item -ItemType Directory -Path (
                Split-Path -Parent $descriptorPath) -Force)
        [IO.File]::WriteAllBytes(
            $descriptorPath,
            (ConvertTo-RSJcsUtf8Bytes -InputObject $descriptor))
        return [pscustomobject]@{
            Descriptor = $descriptor
            DescriptorPath = $descriptorPath
            ChangedPaths = [string[]]@(
                'red-salamander/collection-descriptor.v1.json')
        }
}

function Read-RSDescriptorFixture {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [object]$Fixture
    )

    return Read-RSTerminalEngineCollectionDescriptor `
        -LiteralPath $Fixture.DescriptorPath `
        -CandidateId 'candidate-alpha' `
        -SchemaFile $script:SchemaPath `
        -RepoRoot $script:RepoRoot
}

Describe 'Terminal-engine collection descriptor shared contract' {
    It 'accepts a bounded synthetic descriptor and exact path closure' {
        $fixture = New-RSDescriptorFixture -Root (
            Join-Path $TestDrive 'valid')
        $descriptor = Read-RSDescriptorFixture -Fixture $fixture

        {
            Assert-RSTerminalEngineCollectionDescriptorPathClosure `
                -Descriptor $descriptor `
                -SourceRoot (Split-Path -Parent (
                        Split-Path -Parent $fixture.DescriptorPath)) `
                -CandidateAuthoredPath $fixture.ChangedPaths `
                -CandidateId 'candidate-alpha'
        } | Should Not Throw
    }

    It 'rejects reserved device aliases even when the JSON schema accepts them' {
        $fixture = New-RSDescriptorFixture `
            -Root (Join-Path $TestDrive 'device') `
            -Mutate {
                param($descriptor)
                $descriptor.legal.dependencies[0].inputs[0].
                    destinationLeaf = 'NUL.txt'
            }

        $threw = $false
        try {
            $null = Read-RSDescriptorFixture -Fixture $fixture
        }
        catch {
            $threw = $true
        }
        $threw | Should Be $true
    }

    It 'rejects private-runtime leaves colliding with an adapter artifact' {
        $fixture = New-RSDescriptorFixture `
            -Root (Join-Path $TestDrive 'collision') `
            -Mutate {
                param($descriptor)
                $descriptor.build.privateRuntimeLibraries = [object[]]@(
                    [ordered]@{
                        fileName = 'ADAPTER-ALPHA.dll'
                        relativeBuildPath = 'bin/ADAPTER-ALPHA.dll'
                    })
            }

        $threw = $false
        try {
            $null = Read-RSDescriptorFixture -Fixture $fixture
        }
        catch {
            $threw = $true
        }
        $threw | Should Be $true
    }

    It 'rejects an upstream legal input listed in the exact patch paths' {
        $fixture = New-RSDescriptorFixture -Root (
            Join-Path $TestDrive 'upstream-authored')
        $descriptor = Read-RSDescriptorFixture -Fixture $fixture

        $threw = $false
        try {
            Assert-RSTerminalEngineCollectionDescriptorPathClosure `
                -Descriptor $descriptor `
                -SourceRoot (Split-Path -Parent (
                        Split-Path -Parent $fixture.DescriptorPath)) `
                -CandidateAuthoredPath (
                    $fixture.ChangedPaths + @('LICENSE.txt')) `
                -CandidateId 'candidate-alpha'
        }
        catch {
            $threw = $true
        }
        $threw | Should Be $true
    }

    It 'rejects a reparse parent on the generated-output path' {
        $fixture = New-RSDescriptorFixture -Root (
            Join-Path $TestDrive 'reparse')
        $root = Split-Path -Parent (
            Split-Path -Parent $fixture.DescriptorPath)
        $target = Join-Path $TestDrive 'outside-output'
        [void](New-Item -ItemType Directory -Path $target)
        $generated = Join-Path $root 'red-salamander\generated'
        [void](New-Item -ItemType Junction -Path $generated -Target $target)
        $descriptor = Read-RSDescriptorFixture -Fixture $fixture

        $threw = $false
        try {
            Assert-RSTerminalEngineCollectionDescriptorPathClosure `
                -Descriptor $descriptor `
                -SourceRoot $root `
                -CandidateAuthoredPath $fixture.ChangedPaths `
                -CandidateId 'candidate-alpha'
        }
        catch {
            $threw = $true
        }
        $threw | Should Be $true
    }
}
