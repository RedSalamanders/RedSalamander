Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$policyModule = Join-Path $repoRoot 'Tools\Modules\Packaging\ReleaseArtifactPolicy.psm1'
$testRunPlanModule = Join-Path $repoRoot 'Tools\Modules\Testing\TestRunPlan.psm1'
$buildEvidenceModule = Join-Path $repoRoot 'Tools\Modules\Build\BuildEvidence.psm1'

Import-Module $testRunPlanModule -Force -ErrorAction Stop
Import-Module $policyModule -Force -ErrorAction Stop
Import-Module $buildEvidenceModule -Force -ErrorAction Stop
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
    It 'keeps native Full tooling, rebuild integration and external evidence available' {
        $workflow = Get-Content -LiteralPath (Join-Path $repoRoot '.github/workflows/build-reusable.yml') -Raw
        $workflow | Should Match 'Import-Module Pester -RequiredVersion 3\.4\.0 -Force'
        $workflow | Should Match 'Install-Module Pester -RequiredVersion 3\.4\.0'
        $workflow | Should Match 'GetEnvironmentVariable\(\$name, ''Process''\)'
        $workflow | Should Match '>> \$env:GITHUB_ENV'
        $workflow | Should Match 'Native \$env:RS_TEST_SUITE suite failed with exit code'
        $workflow | Should Match 'MaxCpuCount\s*=\s*2'
        # The ARM64 heap exhaustion (C1002) was the 32-bit x86-hosted cross compiler, which
        # Directory.Build.props now avoids on ARM64 hosts; node count is not the lever.
        $props = Get-Content -LiteralPath (Join-Path $repoRoot 'Directory.Build.props') -Raw
        $expectedArm64HostRule = '<PreferredToolArchitecture Condition="''$(PreferredToolArchitecture)''=='''' and ''$(Platform)''==''ARM64'' and (''$(PROCESSOR_ARCHITECTURE)''==''ARM64'' or ''$(PROCESSOR_ARCHITEW6432)''==''ARM64'')">arm64</PreferredToolArchitecture>'
        $props | Should Match ([regex]::Escape($expectedArm64HostRule))
        # The toolset honors that rule only from an ARM64 MSBuild process, so every MSBuild
        # selection prefers the host-matching 64-bit executable over the x86 Bin\MSBuild.exe.
        $workflow | Should Match "if \(\`$hostArchitecture -eq 'Arm64'\) \{\s*\r?\n\s*@\('MSBuild\\Current\\Bin\\arm64\\MSBuild\.exe', 'MSBuild\\Current\\Bin\\amd64\\MSBuild\.exe'\)"
        $workflow | Should Match "@\('MSBuild\\Current\\Bin\\amd64\\MSBuild\.exe'\)"
        $workflow | Should Match "\[string\[\]\]\`$preferredRelativePaths = if"
        $workflow | Should Match "\(\`$preferredRelativePaths \+ \[string\[\]\]@\('MSBuild\\Current\\Bin\\MSBuild\.exe'\)\)"
        # build.ps1 routes every discovery strategy (vswhere, common installation roots,
        # VSINSTALLDIR) through one host-ordered candidate list.
        $buildScript = Get-Content -LiteralPath (Join-Path $repoRoot 'build.ps1') -Raw
        $buildScript | Should Match 'function Get-RSHostOrderedMSBuildRelativePaths \{'
        # ARM64 hosts also accept the emulated amd64 executable; x64 hosts cannot run arm64\, so it
        # is never a candidate there. Every runnable 64-bit executable, including the legacy 15.0
        # layout, outranks both x86 executables.
        $buildScript | Should Match "if \(\`$isArm64Host\) \{\s*\r?\n\s*@\('MSBuild\\Current\\Bin\\arm64\\MSBuild\.exe', 'MSBuild\\Current\\Bin\\amd64\\MSBuild\.exe'\)\s*\r?\n\s*\} else \{\s*\r?\n\s*@\('MSBuild\\Current\\Bin\\amd64\\MSBuild\.exe'\)"
        $buildScript | Should Match "return \`$candidates \+ \[string\[\]\]@\(\s*\r?\n\s*'MSBuild\\15\.0\\Bin\\amd64\\MSBuild\.exe',\s*\r?\n\s*'MSBuild\\Current\\Bin\\MSBuild\.exe',\s*\r?\n\s*'MSBuild\\15\.0\\Bin\\MSBuild\.exe'\s*\r?\n\s*\)"
        $buildScript | Should Not Match '\$other'
        # A PATH hit (developer prompts expose the x86 MSBuild\Current\Bin) is mapped to the best
        # host-ordered executable of the same installation before it is returned.
        ([regex]::Matches($buildScript, 'Resolve-RSHostOrderedMSBuildPath -Path \$msbuildInPath\.Source')).Count | Should Be 2
        $buildScript | Should Not Match 'Path = \$msbuildInPath\.Source'
        $tokens = $null; $parseErrors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $repoRoot 'build.ps1'), [ref]$tokens, [ref]$parseErrors)
        @($parseErrors).Count | Should Be 0
        $helpers = @($ast.FindAll({ param($node)
                    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                    $node.Name -in @('Get-RSHostOrderedMSBuildRelativePaths', 'Resolve-RSHostOrderedMSBuildPath')
                }, $true) | ForEach-Object { $_.Extent.Text })
        $helpers.Count | Should Be 2
        $installRoot = Join-Path $TestDrive 'Microsoft Visual Studio\18\Community'
        $x86 = Join-Path $installRoot 'MSBuild\Current\Bin\MSBuild.exe'
        $native = Join-Path $installRoot ('MSBuild\Current\Bin\{0}\MSBuild.exe' -f $(if ([Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString() -eq 'Arm64') { 'arm64' } else { 'amd64' }))
        foreach ($file in @($x86, $native)) {
            [void](New-Item -ItemType Directory -Path (Split-Path -Parent $file) -Force)
            Set-Content -LiteralPath $file -Value 'stub'
        }
        $resolved = & ([scriptblock]::Create(($helpers -join "`n") + "`nResolve-RSHostOrderedMSBuildPath -Path '$x86'"))
        $resolved | Should Be $native
        $unrelated = Join-Path $TestDrive 'elsewhere\msbuild.exe'
        (& ([scriptblock]::Create(($helpers -join "`n") + "`nResolve-RSHostOrderedMSBuildPath -Path '$unrelated'"))) | Should Be $unrelated
        ([regex]::Matches($buildScript, 'Get-RSHostOrderedMSBuildRelativePaths \| ForEach-Object')).Count | Should Be 3
        $buildScript | Should Not Match '\\MSBuild\\Current\\Bin\\MSBuild\.exe",'
        # The MSIX packaging job selects MSBuild with the same host order.
        $releaseWorkflow = Get-Content -LiteralPath (Join-Path $repoRoot '.github/workflows/release.yml') -Raw
        $releaseWorkflow | Should Match "if \(\`$hostArchitecture -eq 'Arm64'\) \{\s*\r?\n\s*@\('MSBuild\\Current\\Bin\\arm64\\MSBuild\.exe', 'MSBuild\\Current\\Bin\\amd64\\MSBuild\.exe'\)"
        $releaseWorkflow | Should Match "\(\`$preferredRelativePaths \+ \[string\[\]\]@\('MSBuild\\Current\\Bin\\MSBuild\.exe'\)\)"
        $releaseWorkflow | Should Not Match "-find 'MSBuild\\Current\\Bin\\MSBuild\.exe'"
        $restoreScript = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools/Restore-DxUi.ps1') -Raw
        $restoreScript | Should Match "OSArchitecture\.ToString\(\) -eq 'Arm64'\) \{\s*\r?\n\s*'MSBuild/Current/Bin/arm64/MSBuild\.exe'"
        $evidence = [regex]::Match($workflow, '(?s)- name: Upload native qualification evidence(.*?)(?=      - name:)').Groups[1].Value
        $evidence | Should Match 'steps\.test_root\.outputs\.path'
        $evidence | Should Not Match '\.build/'
        $diagnostics = [regex]::Match($workflow, '(?s)- name: Upload native build diagnostics(.*?)(?=      - name:)').Groups[1].Value
        $diagnostics | Should Match 'include-hidden-files: true'
        $diagnostics | Should Not Match 'steps\.test_root\.outputs\.path'
    }

    It 'maps every MSIX profile to matching product inputs without automatic sanitizer packaging' {
        [xml]$project = Get-Content -LiteralPath (Join-Path $repoRoot 'Installer/msix/RedSalamanderInstaller.wapproj') -Raw
        $solution = Get-Content -LiteralPath (Join-Path $repoRoot 'RedSalamander.sln') -Raw
        foreach ($platform in @('x64', 'ARM64')) {
            foreach ($configuration in @('Debug', 'Release', 'ASan Debug')) {
                $profile = "$configuration|$platform"
                @($project.SelectNodes('//*[local-name()="ProjectConfiguration"]') | Where-Object Include -EQ $profile).Count | Should Be 1
                $mapping = '{9A97489C-F552-4B8C-875A-199D3E5A1E3A}.' + $profile + '.ActiveCfg = ' + $profile
                $solution | Should Match ([regex]::Escape($mapping))
            }
        }
        $solution | Should Not Match '9A97489C-F552-4B8C-875A-199D3E5A1E3A\}\.ASan Debug\|[^\r\n]+\.(Build|Deploy)\.0'
    }

    BeforeAll {
        $releaseWorkflowPath = Join-Path $repoRoot '.github\workflows\release.yml'
        $releaseWorkflow = Get-Content -LiteralPath $releaseWorkflowPath -Raw
        $ghosttyStatusWorkflowPath = Join-Path $repoRoot '.github\workflows\ghostty-runtime-status.yml'
        $ghosttyStatusWorkflow = Get-Content -LiteralPath $ghosttyStatusWorkflowPath -Raw
        $ghosttyNativeWorkflowPath = Join-Path $repoRoot '.github\workflows\ghostty-runtime-native-qualification.yml'
        $ghosttyNativeWorkflow = Get-Content -LiteralPath $ghosttyNativeWorkflowPath -Raw
        $ghosttyProductWorkflowPath = Join-Path $repoRoot '.github\workflows\ghostty-runtime-product-qualification.yml'
        $ghosttyProductWorkflow = Get-Content -LiteralPath $ghosttyProductWorkflowPath -Raw
    }

    It 'uses normal dependency success semantics and no permissive artifact download' {
        $releaseWorkflow | Should Not Match 'continue-on-error:\s*true'
        $releaseWorkflow | Should Not Match '!cancelled\(\)'
        $releaseWorkflow | Should Match 'MSIX packaging disabled by policy'
        $releaseWorkflow | Should Match 'Assert-RSReleaseArtifacts'
        $releaseWorkflow | Should Match 'Assert-RSReleaseChecksumFile'
    }

    It 'uploads deterministic package names and fails when an upload path is missing' {
        $releaseWorkflow | Should Match 'RedSalamander-\$\{\{ needs\.version-info\.outputs\.release_version \}\}-\$\{\{ matrix\.platform \}\}-Portable\.zip'
        $releaseWorkflow | Should Match 'RedSalamander-\$\{\{ needs\.version-info\.outputs\.release_version \}\}-\$\{\{ matrix\.platform \}\}\.msix'
        [regex]::Matches($releaseWorkflow, 'if-no-files-found:\s*error').Count | Should Be 3
        $releaseWorkflow | Should Match 'fail_on_unmatched_files:\s*true'
    }

    It 'keeps Ghostty observation read-only, scheduled, and explicit about candidate identity' {
        $ghosttyStatusWorkflow | Should Match 'cron:\s*"17 03 \* \* \*"'
        $ghosttyStatusWorkflow | Should Match 'cron:\s*"41 04 1 \* \*"'
        $ghosttyStatusWorkflow | Should Match 'options:\s*\[Locked, Online, Candidate\]'
        $ghosttyStatusWorkflow | Should Match 'AdvisoryOnly'
        $ghosttyStatusWorkflow | Should Match 'CandidateCommit'
        $ghosttyStatusWorkflow | Should Match "Candidate mode requires one lowercase full commit"
        $ghosttyStatusWorkflow | Should Match 'candidate_commit is valid only when mode=Candidate'
        $ghosttyStatusWorkflow | Should Match '(?ms)^permissions:\s+contents:\s*read\s+actions:\s*read'
        $ghosttyStatusWorkflow | Should Not Match '(?m)^\s*(issues|pull-requests|deployments|packages|security-events):\s*write'
        $ghosttyStatusWorkflow | Should Not Match '(?i)git\s+(push|checkout\s+-b)|gh\s+(issue|pr)|GhosttyRuntimeLock.*Set-Content'
        $ghosttyStatusWorkflow | Should Match 'result -eq "update-observed"'
        $ghosttyStatusWorkflow | Should Match '(?s)result -in @\("security-review", "required-manual", "unreachable"\).*event_name.*schedule'
        $ghosttyStatusWorkflow | Should Match '(?s)name: Upload status report\s+if: always\(\).*?path: \.build/TerminalEngineStatus/workflow/status\.json'
        $ghosttyStatusWorkflow | Should Match '(?s)name: Enforce scheduled observation result.*?steps\.status\.outputs\.gate_failure'
        $ghosttyStatusWorkflow.IndexOf('name: Upload status report') | Should BeLessThan $ghosttyStatusWorkflow.IndexOf('name: Enforce scheduled observation result')
        $releaseWorkflow | Should Match '(?s)name: Upload Ghostty release preflight\s+if: always\(\).*?path: \.build/TerminalEngineStatus/release/status\.json'
        $releaseWorkflow | Should Match '(?s)name: Enforce Ghostty release preflight.*?source_binding_valid.*?lock_binding_valid.*?result_valid'
        $releaseWorkflow.IndexOf('name: Upload Ghostty release preflight') | Should BeLessThan $releaseWorkflow.IndexOf('name: Enforce Ghostty release preflight')
    }

    It 'qualifies reviewed Ghostty candidates on native ARM64 with read-only artifact evidence' {
        $ghosttyNativeWorkflow | Should Match 'workflow_dispatch:'
        $ghosttyNativeWorkflow | Should Not Match '(?m)^  push:'
        $ghosttyNativeWorkflow | Should Not Match 'vars\.RS_GHOSTTY'
        $ghosttyNativeWorkflow | Should Match 'candidate_commit:'
        $ghosttyNativeWorkflow | Should Match 'candidate_lock_sha256:'
        $ghosttyNativeWorkflow | Should Match 'candidate_patch_sha256:'
        $ghosttyNativeWorkflow | Should Match 'qualify_product:'
        $ghosttyNativeWorkflow | Should Match 'build_number:'
        $ghosttyNativeWorkflow | Should Match '(?ms)^permissions:\s+contents:\s*read'
        $ghosttyNativeWorkflow | Should Not Match '(?m)^\s*(contents|issues|pull-requests|deployments|packages|security-events):\s*write'
        $ghosttyNativeWorkflow | Should Match 'runs-on:\s*windows-11-vs2026-arm'
        $ghosttyNativeWorkflow | Should Match 'git config --global core\.longpaths true'
        $ghosttyNativeWorkflow | Should Match 'persist-credentials:\s*false'
        $ghosttyNativeWorkflow | Should Match 'OSArchitecture\.ToString\(\)'
        $ghosttyNativeWorkflow | Should Match 'ProcessArchitecture\.ToString\(\)'
        $ghosttyNativeWorkflow | Should Match 'platform\.machine\(\)\.upper\(\)'
        $ghosttyNativeWorkflow | Should Match 'RS_GHOSTTY_CANDIDATE_LOCK_B64'
        $ghosttyNativeWorkflow | Should Match 'RS_GHOSTTY_CANDIDATE_PATCH_B64'
        $ghosttyNativeWorkflow | Should Match '(?s)Build-GhosttyTerminalRuntime\.ps1.*?-Platform ARM64.*?-CandidateLockPath.*?-CandidateOutputRoot.*?-Force'
        $ghosttyNativeWorkflow | Should Match 'test-lib-vt test-lib-vt-build test-lib-vt-schema'
        $ghosttyNativeWorkflow | Should Match "RS_TEST_SEED: '0x97d35bc7'"
        $ghosttyNativeWorkflow | Should Match 'build --color off --seed \$env:RS_TEST_SEED -j1 test-lib-vt'
        $ghosttyNativeWorkflow | Should Match 'Dtarget=aarch64-windows-msvc'
        $ghosttyNativeWorkflow | Should Match 'Push-Location -LiteralPath \$sourceRoot'
        $ghosttyNativeWorkflow | Should Match 'lib\\compiler\\test_runner\.zig'
        $ghosttyNativeWorkflow | Should Match 'runnerSha256 = \(Get-FileHash'
        $ghosttyNativeWorkflow | Should Match '\.mode = \.simple'
        $ghosttyNativeWorkflow | Should Match 'red_salamander_vt_c_test_root\.init\(b, \.\{ \.existing = mod\.vt_c \}\)'
        $ghosttyNativeWorkflow | Should Match 'red_salamander_vt_c_test_root\.strip = false'
        $ghosttyNativeWorkflow | Should Match 'red_salamander_vt_c_test_root\.omit_frame_pointer = false'
        $ghosttyNativeWorkflow | Should Match 'red_salamander_vt_c_test_root\.unwind_tables = \.sync'
        $ghosttyNativeWorkflow | Should Match 'transportOnlyOverride = \$true'
        $ghosttyNativeWorkflow | Should Match 'cAbiDiagnosticModuleClone = \$true'
        $ghosttyNativeWorkflow | Should Match 'codeFiveLines\.Count -eq 1'
        $ghosttyNativeWorkflow | Should Match 'failedCommandLines\.Count -eq 1'
        $ghosttyNativeWorkflow | Should Match 'zeroFailureSummaries\.Count -eq 1'
        $ghosttyNativeWorkflow | Should Match 'Retry passed only as diagnostic evidence; qualification requires a clean first execution'
        $ghosttyNativeWorkflow | Should Match 'schemaSummaries\.Count -eq 1'
        $ghosttyNativeWorkflow | Should Match 'reportedTestFailures\.Count -eq 0'
        $ghosttyNativeWorkflow | Should Match 'Join-Path \$sourceRoot ''\.zig-cache'''
        $ghosttyNativeWorkflow | Should Match 'Start-Sleep -Seconds 2'
        $ghosttyNativeWorkflow | Should Match '& \$failedTestPath "--seed=\$env:RS_TEST_SEED"'
        $ghosttyNativeWorkflow | Should Match 'Windows Error Reporting\\LocalDumps\\test\.exe'
        $ghosttyNativeWorkflow | Should Match 'DumpType'
        $ghosttyNativeWorkflow | Should Match 'test_zcu\.obj'
        $ghosttyNativeWorkflow | Should Match 'retryExitCode = \$retryExitCode'
        $ghosttyNativeWorkflow | Should Match 'ghostty-native-arm64-crash-\$\{\{ github\.run_id \}\}'
        $ghosttyNativeWorkflow | Should Match 'if:\s*\$\{\{ failure\(\) \}\}'
        $ghosttyNativeWorkflow | Should Match 'codeFiveRetryExecutableSha256 ='
        $ghosttyNativeWorkflow | Should Not Match 'codeFiveRetryStatus = ''passed'''
        $ghosttyNativeWorkflow | Should Match 'windowsUnmaterializablePaths'
        $ghosttyNativeWorkflow | Should Match 'Compare-Object[\s\S]+?-CaseSensitive'
        $ghosttyNativeWorkflow | Should Match 'status --porcelain=v1 --untracked-files=normal'
        $ghosttyNativeWorkflow | Should Match "Native qualification source differs from the builder's exact Windows allowlist"
        $ghosttyNativeWorkflow | Should Match 'Native qualification left candidate-source changes beyond the exact Windows allowlist'
        $ghosttyNativeWorkflow | Should Match "Get-WinEvent -FilterHashtable"
        $ghosttyNativeWorkflow | Should Match "Id = 1000"
        $ghosttyNativeWorkflow | Should Match "test\\\.exe"
        $ghosttyNativeWorkflow | Should Match 'GhosttyRuntimeNativeQualification\.schema\.json'
        $ghosttyNativeWorkflow | Should Match 'Write-RSAtomicCanonicalJsonFile'
        $ghosttyNativeWorkflow | Should Match 'actions/upload-artifact@043fb46d1a93c77aae656e7c1c64a875d1fc6a0a'
        $ghosttyNativeWorkflow | Should Match 'if-no-files-found:\s*error'
        $ghosttyNativeWorkflow | Should Match '(?ms)^  qualify-product:.*?if:\s*\$\{\{ inputs\.qualify_product \}\}.*?needs:\s*qualify-arm64'
        $ghosttyNativeWorkflow | Should Match 'uses:\s*\./\.github/workflows/ghostty-runtime-product-qualification\.yml'
    }

    It 'qualifies the admitted Ghostty product and portable package on native ARM64' {
        $ghosttyProductWorkflow | Should Match 'workflow_dispatch:'
        $ghosttyProductWorkflow | Should Match 'workflow_call:'
        $ghosttyProductWorkflow | Should Not Match '(?m)^  push:'
        $ghosttyProductWorkflow | Should Match 'build_number:'
        $ghosttyProductWorkflow | Should Match '(?ms)^permissions:\s+contents:\s*read'
        $ghosttyProductWorkflow | Should Not Match '(?m)^\s*(contents|issues|pull-requests|deployments|packages|security-events):\s*write'
        $ghosttyProductWorkflow | Should Match 'runs-on:\s*windows-11-vs2026-arm'
        $ghosttyProductWorkflow | Should Match 'persist-credentials:\s*false'
        $ghosttyProductWorkflow | Should Match "Get-Content -LiteralPath \.\\vcpkg\.json -Raw"
        $ghosttyProductWorkflow | Should Match "'builtin-baseline'"
        $ghosttyProductWorkflow | Should Match 'Get-Content -LiteralPath \.\\vcpkg-tool\.json -Raw'
        $ghosttyProductWorkflow | Should Match 'foreach \(\$requiredCommit in @\(\$toolCommit, \$portsBaseline\)'
        $ghosttyProductWorkflow | Should Match '\$toolMetadata = git show "\$\{toolCommit\}:scripts/vcpkg-tools\.json" \| ConvertFrom-Json'
        $ghosttyProductWorkflow | Should Match '\$portsMetadata = git show "\$\{portsBaseline\}:scripts/vcpkg-tools\.json" \| ConvertFrom-Json'
        $ghosttyProductWorkflow | Should Match '\$toolSchemaVersion -lt \$portsSchemaVersion'
        $ghosttyProductWorkflow | Should Match 'Update vcpkg-tool\.json with the baseline'
        $ghosttyProductWorkflow | Should Match 'git checkout --detach \$toolCommit'
        $ghosttyProductWorkflow | Should Match 'actualToolCommit.*-cne \$toolCommit'
        $ghosttyProductWorkflow | Should Match 'New-Item -ItemType Directory -Force -Path \.\\\.build'
        $ghosttyProductWorkflow | Should Match 'git clone \$toolRepository \.\\\.build\\vcpkg-tool'
        $ghosttyProductWorkflow | Should Match 'Push-Location \.\\\.build\\vcpkg-tool'
        $ghosttyProductWorkflow | Should Match 'git worktree add --detach \.\.\\vcpkg-registry \$portsBaseline'
        $ghosttyProductWorkflow | Should Match 'actualPortsBaseline.*-cne \$portsBaseline'
        $ghosttyProductWorkflow | Should Not Match '\$env:VCPKG_ROOT\s*='
        $ghosttyProductWorkflow | Should Match '(?s)vcpkg-install\.ps1.*?-Platform ARM64.*?-VcpkgExe \.\\\.build\\vcpkg-tool\\vcpkg\.exe.*?-VcpkgRoot \.\\\.build\\vcpkg-registry'
        $ghosttyProductWorkflow.IndexOf('$toolSchemaVersion -lt $portsSchemaVersion') | Should BeLessThan $ghosttyProductWorkflow.IndexOf('git checkout --detach $toolCommit')
        $ghosttyProductWorkflow.IndexOf('.\vcpkg-install.ps1') | Should BeLessThan $ghosttyProductWorkflow.IndexOf('Build admitted Debug ARM64 product with tests')
        $ghosttyProductWorkflow | Should Match 'Configure pinned ARM64 vcpkg build integration'
        $ghosttyProductWorkflow | Should Match "Resolve-Path -LiteralPath \.\\\.build\\vcpkg_installed\\arm64"
        $ghosttyProductWorkflow | Should Match 'Join-Path \$toolRoot ''scripts\\buildsystems\\msbuild'''
        $ghosttyProductWorkflow | Should Match 'ForceImportBeforeCppTargets\s*=\s*\$vcpkgProps'
        $ghosttyProductWorkflow | Should Match 'ForceImportAfterCppTargets\s*=\s*\$vcpkgTargets'
        $ghosttyProductWorkflow | Should Match 'VcpkgManifestInstall\s*=\s*''false'''
        $ghosttyProductWorkflow | Should Match 'VcpkgXUseBuiltInApplocalDeps\s*=\s*''false'''
        $ghosttyProductWorkflow | Should Match 'VCPKG_INSTALLED_DIR\s*=\s*\$installedRoot'
        $ghosttyProductWorkflow | Should Match 'VcpkgInstalledDir\s*=\s*\$installedRoot\.TrimEnd\(''\\''\) \+ ''\\'''
        $ghosttyProductWorkflow | Should Match 'CL\s*=.*?/I"\{0\}".*?\$includeRoot'
        $ghosttyProductWorkflow.IndexOf('Configure pinned ARM64 vcpkg build integration') | Should BeLessThan $ghosttyProductWorkflow.IndexOf('Build admitted Debug ARM64 product with tests')
        $ghosttyProductWorkflow | Should Match 'RSBuildEnableTests:\s*''true'''
        $ghosttyProductWorkflow | Should Match '(?s)build\.ps1.*?-Configuration Debug.*?-Platform ARM64.*?-Rebuild'
        $ghosttyProductWorkflow | Should Match 'Build-GhosttyTerminalRuntime\.ps1 -Platform ARM64 -Offline -PassThru'
        $ghosttyProductWorkflow | Should Match 'PluginContractTests\.exe --terminal-selftests'
        $ghosttyProductWorkflow | Should Match "'terminal_'"
        $ghosttyProductWorkflow | Should Match "'cmd_pane_embedded_terminal_'"
        $ghosttyProductWorkflow | Should Match "'cmd_pane_open_command_shell_prefers_windows_terminal'"
        $ghosttyProductWorkflow | Should Match 'Select test-data root'
        $ghosttyProductWorkflow | Should Match 'Join-Path \(\[IO\.Path\]::GetPathRoot\(\$env:GITHUB_WORKSPACE\)\) ''RedSalamander\.Perf'''
        $ghosttyProductWorkflow | Should Match '(?s)Initialize-RSTestSandboxRoot.*?-TestRoot \$candidate.*?-AllowExternalInitialization'
        $ghosttyProductWorkflow | Should Match 'REDSALAMANDER_TEST_ROOT=\$root'
        $ghosttyProductWorkflow | Should Match '-AllowExternalTestRoot'
        $ghosttyProductWorkflow.IndexOf('Initialize-RSTestSandboxRoot') | Should BeLessThan $ghosttyProductWorkflow.IndexOf('PluginContractTests.exe --terminal-selftests')
        $ghosttyProductWorkflow | Should Not Match '\$caseIndex|ghostty-product[\\\\/]\{0:D2\}'
        $ghosttyProductWorkflow | Should Match 'RSBuildEnableTests:\s*''false'''
        $ghosttyProductWorkflow | Should Match '(?s)build\.ps1.*?-Configuration Release.*?-Platform ARM64.*?-Rebuild'
        $ghosttyProductWorkflow | Should Match '(?s)build-zip\.ps1.*?-Configuration Release.*?-Platform ARM64.*?-BuildNumber'
        $ghosttyProductWorkflow | Should Match 'actions/upload-artifact@043fb46d1a93c77aae656e7c1c64a875d1fc6a0a'
        $ghosttyProductWorkflow | Should Match '\.build/TerminalEngine/Product/ARM64/terminal-runtime-identity\.json'
        $ghosttyProductWorkflow | Should Match 'if-no-files-found:\s*error'
        $ghosttyProductWorkflow | Should Not Match '\.build/ARM64/(Debug|Release)/terminal-runtime-identity\.json'
    }

    It 'runs a source-and-lock-bound Ghostty preflight before release artifact builds' {
        $releaseWorkflow | Should Match '(?ms)^  ghostty-runtime-preflight:.*?Get-GhosttyTerminalRuntimeStatus\.ps1.*?-Mode Online.*?-ExpectedSourceCommit \$env:GITHUB_SHA'
        $releaseWorkflow | Should Match '(?ms)^  ghostty-runtime-preflight:.*?Get-FileHash.*?canonicalLock\.sha256.*?security-review.*?required-manual.*?unreachable'
        $releaseWorkflow | Should Match '(?ms)^  ghostty-product-qualification:.*?needs:\s*\[version-info, ghostty-runtime-preflight\].*?uses:\s*\./\.github/workflows/ghostty-runtime-product-qualification\.yml.*?build_number:\s*\$\{\{ needs\.version-info\.outputs\.build_number \}\}'
        $releaseWorkflow | Should Match '(?m)^  build:\s*\r?\n\s+needs:\s*\[version-info, ghostty-runtime-preflight, ghostty-product-qualification\]'
        $releaseWorkflow | Should Match 'ghostty-runtime-release-preflight-\$\{\{ github\.run_id \}\}'
        $releaseWorkflow | Should Match 'sourceBinding\.repositoryCommit -cne \$env:GITHUB_SHA'
        $releaseWorkflow | Should Match 'canonicalLock\.sha256 -cne \$lockSha256'
    }

    It 'publishes the validated asset matrix atomically and invokes Winget directly' {
        [regex]::Matches($releaseWorkflow, 'uses:\s*softprops/action-gh-release@').Count | Should Be 1
        $releaseWorkflow | Should Match 'files:\s*artifacts/\*'
        $releaseWorkflow | Should Match '(?ms)publish-winget:\s+if:\s*\$\{\{ github\.event_name != ''workflow_dispatch'' \|\| inputs\.build_arm64 \}\}\s+needs:\s*\[version-info, release\].*?uses:\s*\./\.github/workflows/winget-release\.yml'
        $releaseWorkflow | Should Match 'submit:\s*true'
        $releaseWorkflow | Should Match '(?ms)concurrency:\s+group:\s*release-RedSalamander-\$\{\{ github\.run_number \}\}\s+cancel-in-progress:\s*false'
    }

    It 'grants repository write only to publishing jobs and disables checkout persistence elsewhere' {
        $ciWorkflow = Get-Content -LiteralPath (Join-Path $repoRoot '.github\workflows\ci.yml') -Raw
        $buildWorkflow = Get-Content -LiteralPath (Join-Path $repoRoot '.github\workflows\build-reusable.yml') -Raw

        $ciWorkflow | Should Match '(?ms)^permissions:\s+contents:\s*read'
        $formatJob = [regex]::Match($ciWorkflow, '(?ms)^  format:\r?\n(?<body>.*?)(?=^  [a-zA-Z][\w-]*:|\z)').Groups['body'].Value
        $formatJob | Should Match 'Check changed native sources'
        $formatJob | Should Match 'persist-credentials:\s*false'
        $formatJob | Should Not Match 'contents:\s*write|git\s+push'
        $releaseWorkflow | Should Match '(?ms)^permissions:\s+contents:\s*read'
        $releaseWorkflow | Should Match '(?ms)^  release:.*?permissions:\s+contents:\s*write'
        $buildWorkflow | Should Match 'persist-credentials:\s*false'
        $releaseWorkflow | Should Not Match '(?ms)^permissions:\s+contents:\s*write'
    }

    It 'pins installer packages, caches the exact active Terminal input closure, and records canonical toolchain evidence' {
        $buildWorkflow = Get-Content -LiteralPath (Join-Path $repoRoot '.github\workflows\build-reusable.yml') -Raw
        foreach ($workflow in @($buildWorkflow, $releaseWorkflow)) {
            $workflow | Should Match 'visualstudio2026buildtools -PackageVersion 118\.8\.1'
            $workflow | Should Match 'visualstudio2026-workload-vctools -PackageVersion 1\.0\.0'
        }
        $releaseWorkflow | Should Match 'visualstudio2026-workload-universalbuildtools -PackageVersion 1\.0\.0'
        $buildWorkflow | Should Match 'Cache pinned Terminal runtime inputs and product'

        $cacheKeyMatch = [regex]::Match(
            $buildWorkflow,
            "(?m)^\s*key:\s*`\$\{\{ runner\.os \}\}-terminal-runtime-.*?hashFiles\((?<inputs>[^)]*)\)")
        $cacheKeyMatch.Success | Should Be $true
        $actualInputs = @([regex]::Matches($cacheKeyMatch.Groups['inputs'].Value, "'(?<path>[^']+)'") |
            ForEach-Object { $_.Groups['path'].Value })
        $expectedInputs = @((Get-RSGhosttyRuntimeIdentityInputPaths) |
            ForEach-Object { $_.Replace('\', '/') } |
            ForEach-Object {
                if ($_ -ceq 'External/TerminalEngine/Notices/uucode-bjoern-hoehrmann.txt') {
                    'External/TerminalEngine/Notices/*'
                }
                elseif ($_ -cne 'External/TerminalEngine/Notices/unicode-license-3.0-2025.txt') {
                    $_
                }
            })
        @($actualInputs | Sort-Object -Unique).Count | Should Be $actualInputs.Count
        (($actualInputs | Sort-Object) -join "`0") | Should Be (($expectedInputs | Sort-Object) -join "`0")
        $actualInputs | Should Not Contain 'Tools/TerminalEngine/Patches/Ghostty-WindowsARM64-SharedOnly.patch'
        $buildWorkflow | Should Match 'id:\s*toolchain_identity'
        $buildWorkflow | Should Match 'Get-RSBuildToolchainIdentity'
        $buildWorkflow | Should Match 'steps\.toolchain_identity\.outputs\.digest'
        $buildWorkflow | Should Match 'ToolchainIdentity\.schema\.json'
    }

    It 'keeps one Ghostty Windows overlay with the reviewed private-runtime linker policy' {
        $activePatchPath = Join-Path $repoRoot 'Tools\TerminalEngine\Patches\Ghostty-WindowsPrivateRuntime.patch'
        $supersededPatchPath = Join-Path $repoRoot 'Tools\TerminalEngine\Patches\Ghostty-WindowsARM64-SharedOnly.patch'
        $activePatch = Get-Content -LiteralPath $activePatchPath -Raw

        Test-Path -LiteralPath $supersededPatchPath | Should Be $false
        $activePatch | Should Not Match 'if \(config\.target\.result\.os\.tag != \.windows\)'
        $activePatch | Should Not Match 'GhosttyLibVt\.initStatic'
        $activePatch | Should Match 'lib\.use_llvm = true'
        $activePatch | Should Match 'lib\.use_lld = true'
        $activePatch | Should Match 'lib\.root_module\.strip = true'
    }

    It 'scopes Ghostty shared state by repository and takes the session drive lease only for rebuilds' {
        $builderPath = Join-Path $repoRoot 'Tools\TerminalEngine\GhosttyRuntimeLifecycle.psm1'
        $builderSource = Get-Content -LiteralPath $builderPath -Raw
        $tokens = $null
        $parseErrors = $null
        $builderAst = [Management.Automation.Language.Parser]::ParseFile(
            $builderPath,
            [ref]$tokens,
            [ref]$parseErrors)
        @($parseErrors).Count | Should Be 0
        $mutexFunction = $builderAst.Find({
                param($node)
                $node -is [Management.Automation.Language.FunctionDefinitionAst] -and
                    $node.Name -eq 'Get-RSGhosttyRuntimeRepositoryMutexName'
            }, $true)
        $mutexFunction | Should Not BeNullOrEmpty

        . ([scriptblock]::Create($mutexFunction.Extent.Text))
        try {
            $firstWorktree = Join-Path $TestDrive 'first-worktree'
            $secondWorktree = Join-Path $TestDrive 'second-worktree'
            $firstName = Get-RSGhosttyRuntimeRepositoryMutexName -RepoRoot $firstWorktree
            $sameRootName = Get-RSGhosttyRuntimeRepositoryMutexName -RepoRoot ($firstWorktree + '\')
            $secondName = Get-RSGhosttyRuntimeRepositoryMutexName -RepoRoot $secondWorktree
            $firstName | Should Be $sameRootName
            $firstName | Should Not Be $secondName
        }
        finally {
            Remove-Item Function:\Get-RSGhosttyRuntimeRepositoryMutexName -ErrorAction SilentlyContinue
        }

        $rebuildBranchIndex = $builderSource.IndexOf('if (-not $canReuse) {', [StringComparison]::Ordinal)
        $driveLeaseIndex = $builderSource.IndexOf("`$canonicalDriveMutexName = 'Local\RedSalamander.GhosttyRuntime.CanonicalDrive.R'", [StringComparison]::Ordinal)
        $substCreateIndex = $builderSource.IndexOf('-ArgumentList @($canonicalBuildDrive, $repoRoot)', [StringComparison]::Ordinal)
        $substRemoveIndex = $builderSource.IndexOf("-ArgumentList @(`$canonicalBuildDrive, '/D')", [StringComparison]::Ordinal)
        ($rebuildBranchIndex -ge 0) | Should Be $true
        ($driveLeaseIndex -gt $rebuildBranchIndex) | Should Be $true
        ($substCreateIndex -gt $driveLeaseIndex) | Should Be $true
        ($substRemoveIndex -gt $substCreateIndex) | Should Be $true
        [regex]::Matches($builderSource, [regex]::Escape('Local\RedSalamander.GhosttyRuntime.CanonicalDrive.R')).Count | Should Be 1
        $builderSource | Should Not Match 'GhosttyRuntime\.ReproducibleBuild'
        $builderSource | Should Match 'Start-RSContainedProcess\s+-ProcessStartInfo'
        $builderSource | Should Not Match '\[Diagnostics\.Process\]::new\(\)'
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
