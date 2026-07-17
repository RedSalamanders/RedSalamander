Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$policyScript = Join-Path $repoRoot 'Tools\ReleaseArtifactPolicy.ps1'
$testRunPlanScript = Join-Path $repoRoot 'Tools\TestRunPlan.ps1'

. $testRunPlanScript
. $policyScript
Add-Type -AssemblyName System.IO.Compression.FileSystem

function New-RSReleasePolicyTestRoot {
    return (New-RSTestSandboxScratchDirectory `
            -RepoRoot $repoRoot `
            -Harness 'tools-pester' `
            -Case "release-policy-$([Guid]::NewGuid().ToString('N'))")
}

function New-RSFakePeBytes {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet(0x8664, 0xAA64)]
        [int]$Machine
    )

    $bytes = [byte[]]::new(512)
    $bytes[0] = 0x4D
    $bytes[1] = 0x5A
    [BitConverter]::GetBytes([int]128).CopyTo($bytes, 0x3C)
    $bytes[128] = 0x50
    $bytes[129] = 0x45
    $bytes[130] = 0
    $bytes[131] = 0
    [BitConverter]::GetBytes([uint16]$Machine).CopyTo($bytes, 132)
    return ,$bytes
}

function New-RSFakePortableArtifact {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root,

        [Parameter(Mandatory = $true)]
        [string]$ReleaseVersion,

        [Parameter(Mandatory = $true)]
        [ValidateSet('x64', 'ARM64')]
        [string]$Platform,

        [int]$Machine = 0
    )

    if ($Machine -eq 0) {
        $Machine = if ($Platform -eq 'ARM64') { 0xAA64 } else { 0x8664 }
    }
    $path = Join-Path $Root (Get-RSPortableReleaseFileName -ReleaseVersion $ReleaseVersion -Platform $Platform)
    $stream = [IO.File]::Open($path, [IO.FileMode]::CreateNew, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    $archive = [IO.Compression.ZipArchive]::new($stream, [IO.Compression.ZipArchiveMode]::Create, $false)
    try {
        $entry = $archive.CreateEntry('RedSalamander.exe')
        $entryStream = $entry.Open()
        try {
            $bytes = New-RSFakePeBytes -Machine $Machine
            $entryStream.Write($bytes, 0, $bytes.Length)
        }
        finally {
            $entryStream.Dispose()
        }
    }
    finally {
        $archive.Dispose()
        $stream.Dispose()
    }

    return $path
}

function New-RSFakeMsixArtifact {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root,

        [Parameter(Mandatory = $true)]
        [string]$ReleaseVersion,

        [Parameter(Mandatory = $true)]
        [ValidateSet('x64', 'ARM64')]
        [string]$Platform,

        [string]$IdentityVersion = '',

        [string]$ProcessorArchitecture = ''
    )

    if (-not $IdentityVersion) {
        $IdentityVersion = "$ReleaseVersion.0"
    }
    if (-not $ProcessorArchitecture) {
        $ProcessorArchitecture = if ($Platform -eq 'ARM64') { 'arm64' } else { 'x64' }
    }

    $path = Join-Path $Root (Get-RSMsixReleaseFileName -ReleaseVersion $ReleaseVersion -Platform $Platform)
    $stream = [IO.File]::Open($path, [IO.FileMode]::CreateNew, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    $archive = [IO.Compression.ZipArchive]::new($stream, [IO.Compression.ZipArchiveMode]::Create, $false)
    try {
        $entry = $archive.CreateEntry('AppxManifest.xml')
        $entryStream = $entry.Open()
        $writer = [IO.StreamWriter]::new($entryStream, [Text.UTF8Encoding]::new($false))
        try {
            $writer.Write("<?xml version=`"1.0`" encoding=`"utf-8`"?><Package xmlns=`"http://schemas.microsoft.com/appx/manifest/foundation/windows10`"><Identity Name=`"RedSalamander`" Publisher=`"CN=RedSalmanders`" Version=`"$IdentityVersion`" ProcessorArchitecture=`"$ProcessorArchitecture`" /></Package>")
        }
        finally {
            $writer.Dispose()
            $entryStream.Dispose()
        }
    }
    finally {
        $archive.Dispose()
        $stream.Dispose()
    }

    return $path
}

Describe 'Release artifact fail-closed policy' {
    It 'accepts the exact x64-only portable matrix when ARM64 and MSIX are disabled' {
        $root = New-RSReleasePolicyTestRoot
        try {
            New-RSFakePortableArtifact -Root $root -ReleaseVersion '7.0.42' -Platform x64 | Out-Null
            $artifacts = @(Assert-RSReleaseArtifacts -ArtifactsRoot $root -ReleaseVersion '7.0.42' -BuildArm64 $false -BuildMsix $false)
            $artifacts.Count | Should Be 1
            $artifacts[0].Platform | Should Be 'x64'
        }
        finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'accepts the exact x64 and ARM64 portable plus MSIX matrix' {
        $root = New-RSReleasePolicyTestRoot
        try {
            foreach ($platform in @('x64', 'ARM64')) {
                New-RSFakePortableArtifact -Root $root -ReleaseVersion '7.0.42' -Platform $platform | Out-Null
                New-RSFakeMsixArtifact -Root $root -ReleaseVersion '7.0.42' -Platform $platform | Out-Null
            }
            $artifacts = @(Assert-RSReleaseArtifacts -ArtifactsRoot $root -ReleaseVersion '7.0.42' -BuildArm64 $true -BuildMsix $true)
            $artifacts.Count | Should Be 4
        }
        finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'rejects a missing requested x64 portable artifact' {
        $root = New-RSReleasePolicyTestRoot
        try {
            New-RSFakePortableArtifact -Root $root -ReleaseVersion '7.0.42' -Platform ARM64 | Out-Null
            { Assert-RSReleaseArtifacts -ArtifactsRoot $root -ReleaseVersion '7.0.42' -BuildArm64 $false -BuildMsix $false } |
                Should Throw "Expected exactly one release artifact named 'RedSalamander-7.0.42-x64-Portable.zip'; found 0."
        }
        finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'rejects a missing requested ARM64 portable artifact' {
        $root = New-RSReleasePolicyTestRoot
        try {
            New-RSFakePortableArtifact -Root $root -ReleaseVersion '7.0.42' -Platform x64 | Out-Null
            { Assert-RSReleaseArtifacts -ArtifactsRoot $root -ReleaseVersion '7.0.42' -BuildArm64 $true -BuildMsix $false } |
                Should Throw 'Release artifact count mismatch. Expected [RedSalamander-7.0.42-ARM64-Portable.zip, RedSalamander-7.0.42-x64-Portable.zip]; found [RedSalamander-7.0.42-x64-Portable.zip].'
        }
        finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'allows omitted MSIX only when the explicit policy input disables it' {
        $root = New-RSReleasePolicyTestRoot
        try {
            New-RSFakePortableArtifact -Root $root -ReleaseVersion '7.0.42' -Platform x64 | Out-Null
            @(Assert-RSReleaseArtifacts -ArtifactsRoot $root -ReleaseVersion '7.0.42' -BuildArm64 $false -BuildMsix $false).Count |
                Should Be 1
            { Assert-RSReleaseArtifacts -ArtifactsRoot $root -ReleaseVersion '7.0.42' -BuildArm64 $false -BuildMsix $true } |
                Should Throw 'Release artifact count mismatch. Expected [RedSalamander-7.0.42-x64-Portable.zip, RedSalamander-7.0.42-x64.msix]; found [RedSalamander-7.0.42-x64-Portable.zip].'
        }
        finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'rejects a portable filename whose PE machine does not match its platform' {
        $root = New-RSReleasePolicyTestRoot
        try {
            New-RSFakePortableArtifact -Root $root -ReleaseVersion '7.0.42' -Platform x64 -Machine 0xAA64 | Out-Null
            { Assert-RSReleaseArtifacts -ArtifactsRoot $root -ReleaseVersion '7.0.42' -BuildArm64 $false -BuildMsix $false } |
                Should Throw "Portable artifact 'RedSalamander-7.0.42-x64-Portable.zip' has PE machine 0xAA64; expected 0x8664 for x64."
        }
        finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'rejects MSIX identity version and architecture mismatches' {
        $root = New-RSReleasePolicyTestRoot
        try {
            New-RSFakePortableArtifact -Root $root -ReleaseVersion '7.0.42' -Platform x64 | Out-Null
            New-RSFakeMsixArtifact -Root $root -ReleaseVersion '7.0.42' -Platform x64 -IdentityVersion '7.0.41.0' | Out-Null
            { Assert-RSReleaseArtifacts -ArtifactsRoot $root -ReleaseVersion '7.0.42' -BuildArm64 $false -BuildMsix $true } |
                Should Throw "MSIX 'RedSalamander-7.0.42-x64.msix' has Identity Version '7.0.41.0'; expected '7.0.42.0'."
        }
        finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'writes and revalidates exact SHA256 entries and rejects tampering' {
        $root = New-RSReleasePolicyTestRoot
        try {
            foreach ($platform in @('x64', 'ARM64')) {
                New-RSFakePortableArtifact -Root $root -ReleaseVersion '7.0.42' -Platform $platform | Out-Null
            }
            New-RSReleaseChecksumFile -ArtifactsRoot $root -ReleaseVersion '7.0.42' -BuildArm64 $true -BuildMsix $false | Out-Null
            Assert-RSReleaseChecksumFile -ArtifactsRoot $root -ReleaseVersion '7.0.42' -BuildArm64 $true -BuildMsix $false |
                Should Be $true

            $checksumPath = Join-Path $root 'checksums.sha256'
            $lines = @(Get-Content $checksumPath)
            $lines[0] = ('0' * 64) + $lines[0].Substring(64)
            [IO.File]::WriteAllLines($checksumPath, $lines, [Text.UTF8Encoding]::new($false))
            $expectedMessage = "Release checksum mismatch at line 1. Expected '$((Get-FileHash -LiteralPath (Join-Path $root 'RedSalamander-7.0.42-ARM64-Portable.zip') -Algorithm SHA256).Hash)  RedSalamander-7.0.42-ARM64-Portable.zip'; found '$($lines[0])'."
            { Assert-RSReleaseChecksumFile -ArtifactsRoot $root -ReleaseVersion '7.0.42' -BuildArm64 $true -BuildMsix $false } |
                Should Throw $expectedMessage
        }
        finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

Describe 'Release workflow source contracts' {
    BeforeAll {
        $releaseWorkflowPath = Join-Path $repoRoot '.github\workflows\release.yml'
        $releaseWorkflow = Get-Content -LiteralPath $releaseWorkflowPath -Raw
    }

    It 'uses normal dependency success semantics and no permissive artifact download' {
        $releaseWorkflow | Should Not Match 'continue-on-error:\s*true'
        $releaseWorkflow | Should Not Match '(?m)^\s*if:\s*>?-?\s*$[\s\S]{0,120}always\(\)'
        $releaseWorkflow | Should Not Match '!cancelled\(\)'
        $releaseWorkflow | Should Match 'MSIX packaging disabled by policy'
        $releaseWorkflow | Should Match 'Assert-RSReleaseArtifacts'
        $releaseWorkflow | Should Match 'Assert-RSReleaseChecksumFile'
    }

    It 'uploads deterministic package names and fails when an upload path is missing' {
        $releaseWorkflow | Should Match 'RedSalamander-\$\{\{ needs\.version-info\.outputs\.release_version \}\}-\$\{\{ matrix\.platform \}\}-Portable\.zip'
        $releaseWorkflow | Should Match 'RedSalamander-\$\{\{ needs\.version-info\.outputs\.release_version \}\}-\$\{\{ matrix\.platform \}\}\.msix'
        [regex]::Matches($releaseWorkflow, 'if-no-files-found:\s*error').Count | Should Be 2
        $releaseWorkflow | Should Match 'fail_on_unmatched_files:\s*true'
    }

    It 'keeps every active third-party action pinned to a full SHA with an exact version comment' {
        $workflowFiles = @(Get-ChildItem (Join-Path $repoRoot '.github\workflows') -File |
                Where-Object { $_.Extension -in '.yml', '.yaml' })
        $externalCount = 0
        foreach ($workflowFile in $workflowFiles) {
            $content = Get-Content -LiteralPath $workflowFile.FullName -Raw
            foreach ($match in [regex]::Matches($content, '(?m)^\s*uses:\s*(?<reference>[^\s#]+)(?<comment>\s+#\s+[^\r\n]+)?\s*$')) {
                $reference = $match.Groups['reference'].Value
                if ($reference.StartsWith('./')) {
                    continue
                }
                ++$externalCount
                if ($reference -notmatch '@[0-9a-f]{40}$') {
                    throw "$($workflowFile.Name) contains mutable action reference '$reference'."
                }
                if ($match.Groups['comment'].Value -notmatch '#\s+v\d+\.\d+\.\d+\s*$') {
                    throw "$($workflowFile.Name) action '$reference' lacks an exact version comment."
                }
            }
        }
        ($externalCount -gt 0) | Should Be $true
    }

    It 'removes the five no-op Squad workflows and broken promotion workflow' {
        $removed = @(
            'squad-ci.yml',
            'squad-docs.yml',
            'squad-preview.yml',
            'squad-insider-release.yml',
            'squad-release.yml',
            'squad-promote.yml'
        )
        foreach ($name in $removed) {
            (Test-Path -LiteralPath (Join-Path $repoRoot ".github\workflows\$name")) | Should Be $false
        }
    }

    It 'enables GitHub Actions Dependabot and requires workflow owner review' {
        $dependabot = Get-Content -LiteralPath (Join-Path $repoRoot '.github\dependabot.yml') -Raw
        $codeowners = Get-Content -LiteralPath (Join-Path $repoRoot '.github\CODEOWNERS') -Raw
        $dependabot | Should Match 'package-ecosystem:\s*github-actions'
        $dependabot | Should Match '(?ms)package-ecosystem:\s*github-actions.*?directory:\s*/'
        $codeowners | Should Match '(?m)^/\.github/workflows/\s+@[A-Za-z0-9-]+\s*$'
        $codeowners | Should Match '(?m)^/\.github/dependabot\.yml\s+@[A-Za-z0-9-]+\s*$'
    }
}
