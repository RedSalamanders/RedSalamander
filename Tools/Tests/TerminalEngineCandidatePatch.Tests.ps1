Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. (Join-Path $repoRoot 'Tools\TestRunPlan.ps1')
Import-Module (Join-Path $repoRoot 'Tools\TerminalEvidence.psm1') -Force -Global

function New-RSTerminalCandidatePatchFixture {
    param(
        [ValidateSet('collect', 'adaptation-rejected')]
        [string]$Disposition = 'collect'
    )

    $scratch = New-RSTestSandboxScratchDirectory `
        -RepoRoot $repoRoot `
        -Harness 'terminal-engine-candidate-patch-pester' `
        -Case ([Guid]::NewGuid().ToString('N'))
    $fixtureRoot = Join-Path $scratch 'repo'
    foreach ($leaf in @(
            'Tools\TerminalEngine\Test-TerminalEngineCandidatePatch.ps1',
            'Tools\TerminalEngine\TerminalEngineCollectionDescriptor.psm1',
            'Tools\TerminalEvidence.psm1',
            'Specs\Terminal\TerminalEngineConcernMapping.schema.json',
            'Specs\Terminal\TerminalEngineCollectionDescriptor.schema.json')) {
        $target = Join-Path $fixtureRoot $leaf
        [void](New-Item -ItemType Directory -Path (
                [IO.Path]::GetDirectoryName($target)) -Force)
        Copy-Item -LiteralPath (Join-Path $repoRoot $leaf) -Destination $target
    }
    $source = Join-Path $fixtureRoot 'source'
    [void](New-Item -ItemType Directory -Path $source -Force)
    [IO.File]::WriteAllText(
        (Join-Path $source 'engine.txt'),
        "baseline`n",
        [Text.UTF8Encoding]::new($false))
    & git init -q $source
    & git -C $source config core.longpaths true
    & git -C $source config user.name 'Terminal Candidate Patch Test'
    & git -C $source config user.email 'terminal-patch-test@example.invalid'
    & git -C $source config core.autocrlf false
    & git -C $source add -- engine.txt
    & git -C $source commit -q -m upstream
    if ($LASTEXITCODE -ne 0) {
        throw 'Unable to create the candidate-patch test upstream commit.'
    }
    $upstreamCommit = [string](& git -C $source rev-parse HEAD)
    $upstreamTree = [string](& git -C $source rev-parse 'HEAD^{tree}')
    [IO.File]::WriteAllText(
        (Join-Path $source 'engine.txt'),
        "baseline`nadapted`n",
        [Text.UTF8Encoding]::new($false))
    $enginePath = Join-Path $source 'engine.txt'
    $engineItem = Get-Item -LiteralPath $enginePath -Force
    $engineSha256 = (Get-FileHash `
            -LiteralPath $enginePath `
            -Algorithm SHA256).Hash.ToLowerInvariant()
    $descriptorSha256 = ''
    if ($Disposition -ceq 'collect') {
        $descriptor = [ordered]@{
            formatId = 'red-salamander-terminal-engine-collection-descriptor'
            version = 1
            candidateId = 'fixture'
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
                                sha256 = $engineSha256
                                sizeBytes = $engineItem.Length.ToString(
                                    [Globalization.CultureInfo]::InvariantCulture)
                                source = [ordered]@{
                                    sourceKind = 'candidate-source'
                                    origin = 'candidate-patch'
                                    sourceRelativePath = 'engine.txt'
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
        $descriptorPath = Join-Path $source (
            'red-salamander\collection-descriptor.v1.json')
        [void](New-Item -ItemType Directory -Path (
                Split-Path -Parent $descriptorPath) -Force)
        [IO.File]::WriteAllBytes(
            $descriptorPath,
            (ConvertTo-RSJcsUtf8Bytes -InputObject $descriptor))
        $descriptorSha256 = (Get-FileHash `
                -LiteralPath $descriptorPath `
                -Algorithm SHA256).Hash.ToLowerInvariant()
    }
    & git -C $source add -- .
    $resultTree = [string](& git -C $source write-tree)

    $candidateRoot = Join-Path $fixtureRoot (
        'External\TerminalEngine\CandidatePatches')
    [void](New-Item -ItemType Directory -Path $candidateRoot -Force)
    $patch = Join-Path $candidateRoot 'fixture.patch'
    & git -C $source -c core.autocrlf=false diff --cached --binary `
        --full-index --no-renames "--output=$patch" $upstreamCommit
    if ($LASTEXITCODE -ne 0) {
        throw 'Unable to create the candidate-patch test diff.'
    }
    $patchSha256 = (Get-FileHash -LiteralPath $patch -Algorithm SHA256).
        Hash.ToLowerInvariant()
    $pathConcerns = [Collections.Generic.List[object]]::new()
    $pathConcerns.Add([ordered]@{
            path = 'engine.txt'
            concerns = @('abi-snapshot-identity')
        })
    if ($Disposition -ceq 'collect') {
        $pathConcerns.Add([ordered]@{
                path =
                    'red-salamander/collection-descriptor.v1.json'
                concerns = @('abi-snapshot-identity')
            })
    }
    $mappingPayload = [ordered]@{
        candidateId = 'fixture'
        upstreamTree = $upstreamTree.Trim()
        resultTree = $resultTree.Trim()
        patchSha256 = $patchSha256
        pathConcerns = [object[]]$pathConcerns.ToArray()
    }
    $mapping = [ordered]@{
        formatId = 'red-salamander-terminal-engine-concern-mapping'
        version = 1
        candidateId = 'fixture'
        upstreamTree = $mappingPayload.upstreamTree
        resultTree = $mappingPayload.resultTree
        patchSha256 = $patchSha256
        pathConcerns = $mappingPayload.pathConcerns
        mappingSha256 = Get-RSJcsSha256 -InputObject $mappingPayload
    }
    $mappingPath = Join-Path $candidateRoot 'fixture.concerns.json'
    [IO.File]::WriteAllBytes(
        $mappingPath,
        (ConvertTo-RSJcsUtf8Bytes -InputObject $mapping))
    return [pscustomobject]@{
        Script = Join-Path $fixtureRoot (
            'Tools\TerminalEngine\Test-TerminalEngineCandidatePatch.ps1')
        Source = $source
        UpstreamCommit = $upstreamCommit.Trim()
        UpstreamTree = $upstreamTree.Trim()
        ResultTree = $resultTree.Trim()
        Patch = $patch
        PatchSha256 = $patchSha256
        Mapping = $mappingPath
        DescriptorSha256 = $descriptorSha256
        Disposition = $Disposition
        ChangedMaintainedLines = if ($Disposition -ceq 'collect') { 2 } else { 1 }
        ChangedMaintainedFiles = if ($Disposition -ceq 'collect') { 2 } else { 1 }
    }
}

function Invoke-RSTerminalCandidatePatchFixture {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Fixture,

        [string]$ExpectedResultTree = '',

        [string]$ExpectedDescriptorSha256 = ''
    )

    if ([string]::IsNullOrEmpty($ExpectedResultTree)) {
        $ExpectedResultTree = [string]$Fixture.ResultTree
    }
    if ([string]::IsNullOrEmpty($ExpectedDescriptorSha256) -and
        $Fixture.Disposition -ceq 'collect') {
        $ExpectedDescriptorSha256 = [string]$Fixture.DescriptorSha256
    }
    $arguments = @{
        CandidateId = 'fixture'
        Disposition = [string]$Fixture.Disposition
        SourceRepository = $Fixture.Source
        UpstreamCommit = $Fixture.UpstreamCommit
        ExpectedUpstreamTree = $Fixture.UpstreamTree
        PatchPath = $Fixture.Patch
        ExpectedPatchSha256 = $Fixture.PatchSha256
        ExpectedResultTree = $ExpectedResultTree
        ExpectedChangedMaintainedLines = $Fixture.ChangedMaintainedLines
        ExpectedChangedMaintainedFiles = $Fixture.ChangedMaintainedFiles
        ExpectedBinaryFiles = 0
        ConcernMappingPath = $Fixture.Mapping
    }
    if (-not [string]::IsNullOrEmpty($ExpectedDescriptorSha256)) {
        $arguments.ExpectedCollectionDescriptorSha256 =
            $ExpectedDescriptorSha256
    }
    return & $Fixture.Script @arguments
}

Describe 'TerminalEngine exact candidate patch proof' {
    It 'reapplies the canonical patch and proves tree, metrics, and mapping closure' {
        $fixture = New-RSTerminalCandidatePatchFixture
        $proof = Invoke-RSTerminalCandidatePatchFixture -Fixture $fixture

        $proof.status | Should Be 'verified'
        $proof.resultTree | Should Be $fixture.ResultTree
        $proof.changedMaintainedLines | Should Be 2
        $proof.changedMaintainedFiles | Should Be 2
        $proof.binaryFiles | Should Be 0
        $proof.collectionDescriptorRelativePath |
            Should Be 'red-salamander/collection-descriptor.v1.json'
        $proof.collectionDescriptorSha256 |
            Should Be $fixture.DescriptorSha256
        $proof.buildKind | Should Be 'provider-generic-cmake-msvc-v1'
        $proof.maintainedBuildWrapper | Should Be $false
        $proof.disposition | Should Be 'collect'
    }

    It 'proves an adaptation rejection without requiring a collection descriptor' {
        $fixture = New-RSTerminalCandidatePatchFixture `
            -Disposition adaptation-rejected
        $proof = Invoke-RSTerminalCandidatePatchFixture -Fixture $fixture

        $proof.status | Should Be 'verified'
        $proof.disposition | Should Be 'adaptation-rejected'
        $proof.changedMaintainedLines | Should Be 1
        $proof.changedMaintainedFiles | Should Be 1
        $proof.PSObject.Properties.Name -ccontains
            'collectionDescriptorRelativePath' | Should Be $false
        $proof.PSObject.Properties.Name -ccontains
            'maintainedBuildWrapper' | Should Be $false
    }

    It 'rejects a collection descriptor in an adaptation-rejected result tree' {
        $fixture = New-RSTerminalCandidatePatchFixture
        $fixture.Disposition = 'adaptation-rejected'
        $caught = $null
        try {
            Invoke-RSTerminalCandidatePatchFixture `
                -Fixture $fixture | Out-Null
        }
        catch {
            $caught = $_.Exception
        }
        $caught | Should Not BeNullOrEmpty
        $caught.Message | Should Match 'must not contain a collection descriptor'
    }

    It 'rejects a collection descriptor digest mismatch' {
        $fixture = New-RSTerminalCandidatePatchFixture
        $caught = $null
        try {
            Invoke-RSTerminalCandidatePatchFixture `
                -Fixture $fixture `
                -ExpectedDescriptorSha256 ('f' * 64) | Out-Null
        }
        catch {
            $caught = $_.Exception
        }
        $caught | Should Not BeNullOrEmpty
        $caught.Message | Should Match 'collection descriptor SHA-256 differs'
    }

    It 'rejects a frozen result-tree mismatch' {
        $fixture = New-RSTerminalCandidatePatchFixture
        $caught = $null
        try {
            Invoke-RSTerminalCandidatePatchFixture `
                -Fixture $fixture `
                -ExpectedResultTree ('f' * 40) | Out-Null
        } catch {
            $caught = $_.Exception
        }
        $caught | Should Not BeNullOrEmpty
        $caught.Message | Should Match 'result tree differs'
    }

    It 'rejects a concern mapping that is not the exact patch-path closure' {
        $fixture = New-RSTerminalCandidatePatchFixture
        $mapping = Get-Content -LiteralPath $fixture.Mapping -Raw |
            ConvertFrom-Json -AsHashtable -Depth 64
        $mapping.pathConcerns[0].path = 'other.txt'
        $mappingPayload = [ordered]@{
            candidateId = [string]$mapping.candidateId
            upstreamTree = [string]$mapping.upstreamTree
            resultTree = [string]$mapping.resultTree
            patchSha256 = [string]$mapping.patchSha256
            pathConcerns = [object[]]@($mapping.pathConcerns)
        }
        $mapping.mappingSha256 = Get-RSJcsSha256 -InputObject $mappingPayload
        [IO.File]::WriteAllBytes(
            $fixture.Mapping,
            (ConvertTo-RSJcsUtf8Bytes -InputObject $mapping))

        $caught = $null
        try {
            Invoke-RSTerminalCandidatePatchFixture -Fixture $fixture | Out-Null
        } catch {
            $caught = $_.Exception
        }
        $caught | Should Not BeNullOrEmpty
        $caught.Message | Should Match 'exact sorted patch-path closure'
    }
}
