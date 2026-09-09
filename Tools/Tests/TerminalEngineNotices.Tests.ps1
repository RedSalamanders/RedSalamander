Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$modulePath = Join-Path $repoRoot (
    'Tools\TerminalEngine\TerminalEngineNotices.psm1')
Import-Module $modulePath -Force

function Get-RSNoticeTestRelativePath {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    return [IO.Path]::GetRelativePath($repoRoot, $Path).Replace('\', '/')
}

function Write-RSNoticeTestUtf8 {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Text
    )

    $parent = [IO.Path]::GetDirectoryName($Path)
    [void](New-Item -ItemType Directory -Path $parent -Force)
    [IO.File]::WriteAllText(
        $Path,
        $Text,
        [Text.UTF8Encoding]::new($false))
}

function Write-RSNoticeTestJson {
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [object]$Value
    )

    $json = $Value | ConvertTo-Json -Depth 100
    Write-RSNoticeTestUtf8 -Path $Path -Text $json
}

function New-RSNoticeTestInput {
    param(
        [Parameter(Mandatory)]
        [string]$Root,

        [Parameter(Mandatory)]
        [string]$DependencyId,

        [Parameter(Mandatory)]
        [string]$InputId,

        [Parameter(Mandatory)]
        [ValidateSet('license', 'notice')]
        [string]$Kind,

        [Parameter(Mandatory)]
        [string]$Text
    )

    $path = Join-Path $Root "$DependencyId-$InputId.txt"
    Write-RSNoticeTestUtf8 -Path $path -Text $Text
    $item = Get-Item -LiteralPath $path
    return [ordered]@{
        id = $InputId
        kind = $Kind
        path = Get-RSNoticeTestRelativePath -Path $path
        sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).
            Hash.ToLowerInvariant()
        sizeBytes = $item.Length.ToString(
            [Globalization.CultureInfo]::InvariantCulture)
        sourceUrl = "https://example.invalid/$DependencyId/$InputId.txt"
        retrievedUtc = '2026-07-30T11:42:56Z'
        sourceSha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).
            Hash.ToLowerInvariant()
        sourceMember = ''
    }
}

function New-RSNoticeTestFixture {
    param(
        [Parameter(Mandatory)]
        [string]$Root
    )

    [void](New-Item -ItemType Directory -Path $Root -Force)
    $requiredIds = @(
        'boxed-cpp',
        'candidate-alpha',
        'gsl',
        'libunicode',
        'reflection-cpp',
        'unicode-ucd',
        'zlib')
    $pins = @{
        'boxed-cpp' = '1.4.3'
        'candidate-alpha' = '0397f9cfbfcc94d46103d3cff70ebf59326dd695'
        gsl = '3.1.0'
        libunicode = '0.9.2'
        'reflection-cpp' = '0.4.0'
        'unicode-ucd' = '17.0.0'
        zlib = '1.3.2'
    }
    $spdx = @{
        'boxed-cpp' = 'Apache-2.0'
        'candidate-alpha' = 'Apache-2.0'
        gsl = 'MIT'
        libunicode = 'Apache-2.0'
        'reflection-cpp' = 'Apache-2.0'
        'unicode-ucd' = 'Unicode-3.0'
        zlib = 'Zlib'
    }
    $dependencies = [Collections.Generic.List[object]]::new()
    foreach ($dependencyId in $requiredIds) {
        $inputs = [Collections.Generic.List[object]]::new()
        $inputs.Add((New-RSNoticeTestInput `
                    -Root $Root `
                    -DependencyId $dependencyId `
                    -InputId 'license' `
                    -Kind license `
                    -Text "License text for $dependencyId.`r`nExact pin $($pins[$dependencyId]).`r`n"))
        if ($dependencyId -ceq 'gsl') {
            $inputs.Add((New-RSNoticeTestInput `
                        -Root $Root `
                        -DependencyId $dependencyId `
                        -InputId 'third-party' `
                        -Kind notice `
                        -Text "Additional notice for $dependencyId.`n"))
        }
        $dependencies.Add([ordered]@{
                id = $dependencyId
                pin = $pins[$dependencyId]
                spdx = $spdx[$dependencyId]
                inputs = [object[]]$inputs.ToArray()
            })
    }
    $manifest = [ordered]@{
        formatId = 'red-salamander-terminal-engine-notices'
        schemaVersion = 1
        closureId = 'candidate-alpha-source-static-v1'
        dependencies = [object[]]$dependencies.ToArray()
    }
    $manifestPath = Join-Path $Root 'notices.json'
    Write-RSNoticeTestJson -Path $manifestPath -Value $manifest
    return [pscustomobject]@{
        Root = $Root
        RequiredIds = $requiredIds
        Manifest = $manifest
        ManifestPath = $manifestPath
    }
}

Describe 'Terminal-engine deterministic third-party notices' {
    BeforeAll {
        $script:testRoot = Join-Path $repoRoot (
            '.build\TerminalNoticesTests\' + [Guid]::NewGuid().ToString('N'))
        $script:fixture = New-RSNoticeTestFixture -Root $script:testRoot
    }

    AfterAll {
        $expectedRoot = [IO.Path]::GetFullPath(
            (Join-Path $repoRoot '.build\TerminalNoticesTests'))
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

    It 'materializes the complete seven-dependency closure byte-identically' {
        $outputOne = Join-Path $script:testRoot 'run-one\THIRD-PARTY-NOTICES.txt'
        $receiptOne = Join-Path $script:testRoot 'run-one\receipt.json'
        $outputTwo = Join-Path $script:testRoot 'run-two\THIRD-PARTY-NOTICES.txt'
        $receiptTwo = Join-Path $script:testRoot 'run-two\receipt.json'

        $first = Invoke-RSTerminalNoticeMaterialization `
            -ManifestFile $script:fixture.ManifestPath `
            -OutputFile $outputOne `
            -ReceiptFile $receiptOne `
            -RequiredDependencyId $script:fixture.RequiredIds `
            -RepoRoot $repoRoot
        $second = Invoke-RSTerminalNoticeMaterialization `
            -ManifestFile $script:fixture.ManifestPath `
            -OutputFile $outputTwo `
            -ReceiptFile $receiptTwo `
            -RequiredDependencyId $script:fixture.RequiredIds `
            -RepoRoot $repoRoot

        $first.OutputSha256 | Should Be $second.OutputSha256
        $first.ReceiptSha256 | Should Be $second.ReceiptSha256
        [Convert]::ToHexString([IO.File]::ReadAllBytes($outputOne)) |
            Should Be ([Convert]::ToHexString([IO.File]::ReadAllBytes($outputTwo)))
        [Convert]::ToHexString([IO.File]::ReadAllBytes($receiptOne)) |
            Should Be ([Convert]::ToHexString([IO.File]::ReadAllBytes($receiptTwo)))

        $receiptText = [IO.File]::ReadAllText($receiptOne)
        $receipt = $receiptText |
            ConvertFrom-Json -Depth 100 -NoEnumerate -DateKind String
        $receipt.formatId |
            Should Be 'red-salamander-terminal-engine-notice-receipt'
        $receipt.dependencyCount | Should Be '7'
        @($receipt.dependencies).Count | Should Be 7
        $receipt.notice.sha256 | Should Be $first.OutputSha256
        $receipt.notice.sizeBytes |
            Should Be ([string](Get-Item -LiteralPath $outputOne).Length)
        $receiptText | Should Not Match [regex]::Escape($script:testRoot)
        $receiptText | Should Not Match 'License text for'

        $notice = [IO.File]::ReadAllText($outputOne)
        $previousOffset = -1
        foreach ($dependencyId in $script:fixture.RequiredIds) {
            $offset = $notice.IndexOf(
                "Dependency: $dependencyId",
                [StringComparison]::Ordinal)
            ($offset -gt $previousOffset) | Should Be $true
            $previousOffset = $offset
        }
        $notice | Should Match 'Additional notice for gsl\.'
        $notice.Contains("`r") | Should Be $false
        [IO.File]::ReadAllBytes($outputOne)[0] | Should Not Be 0xEF
    }

    It 'exposes the content-free receipt through the command wrapper' {
        $output = Join-Path $script:testRoot 'wrapper\THIRD-PARTY-NOTICES.txt'
        $receipt = Join-Path $script:testRoot 'wrapper\receipt.json'
        $receiptResult = & (Join-Path $repoRoot 'Tools\New-TerminalEngineNotices.ps1') `
            -ManifestFile $script:fixture.ManifestPath `
            -OutputFile $output `
            -ReceiptFile $receipt `
            -RequiredDependencyId $script:fixture.RequiredIds `
            -PassThruReceipt

        [IO.Path]::GetFullPath([string]$receiptResult) |
            Should Be ([IO.Path]::GetFullPath($receipt))
        (Test-Path -LiteralPath $output -PathType Leaf) | Should Be $true
        (Test-Path -LiteralPath $receipt -PathType Leaf) | Should Be $true
    }

    It 'rejects wrong hashes, incomplete closure, duplicates, reordering, and path escape' {
        $invalidPath = Join-Path $script:testRoot 'invalid.json'
        $cases = @(
            [ordered]@{
                Name = 'wrong input hash'
                Pattern = 'SHA-256 differs'
                Mutate = {
                    param($Manifest)
                    $Manifest.dependencies[0].inputs[0].sha256 = ('0' * 64)
                }
            },
            [ordered]@{
                Name = 'missing dependency'
                Pattern = 'required closure has 7'
                Mutate = {
                    param($Manifest)
                    $Manifest.dependencies = @(
                        $Manifest.dependencies | Select-Object -First 6)
                }
            },
            [ordered]@{
                Name = 'extra dependency'
                Pattern = 'required closure has 7'
                Mutate = {
                    param($Manifest)
                    $extra = $Manifest.dependencies[-1] |
                        ConvertTo-Json -Depth 100 |
                        ConvertFrom-Json -AsHashtable -Depth 100
                    $extra.id = 'zz-extra'
                    $Manifest.dependencies = @($Manifest.dependencies) + @($extra)
                }
            },
            [ordered]@{
                Name = 'duplicate dependency'
                Pattern = 'does not match required ordinal closure entry'
                Mutate = {
                    param($Manifest)
                    $Manifest.dependencies[1].id = $Manifest.dependencies[0].id
                }
            },
            [ordered]@{
                Name = 'reordered dependency'
                Pattern = 'does not match required ordinal closure entry'
                Mutate = {
                    param($Manifest)
                    $first = $Manifest.dependencies[0]
                    $Manifest.dependencies[0] = $Manifest.dependencies[1]
                    $Manifest.dependencies[1] = $first
                }
            },
            [ordered]@{
                Name = 'path escape'
                Pattern = 'escapes the repository root'
                Mutate = {
                    param($Manifest)
                    $Manifest.dependencies[0].inputs[0].path = '../outside.txt'
                }
            })

        foreach ($case in $cases) {
            $invalid = $script:fixture.Manifest |
                ConvertTo-Json -Depth 100 |
                ConvertFrom-Json -AsHashtable -Depth 100
            & $case.Mutate $invalid
            Write-RSNoticeTestJson -Path $invalidPath -Value $invalid
            $rejected = $false
            $message = ''
            try {
                Invoke-RSTerminalNoticeMaterialization `
                    -ManifestFile $invalidPath `
                    -OutputFile (Join-Path $script:testRoot 'invalid-notices.txt') `
                    -ReceiptFile (Join-Path $script:testRoot 'invalid-receipt.json') `
                    -RequiredDependencyId $script:fixture.RequiredIds `
                    -RepoRoot $repoRoot |
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
}
