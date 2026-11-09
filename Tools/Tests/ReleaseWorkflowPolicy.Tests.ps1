Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$policyModule = Join-Path $repoRoot 'Tools\Modules\Packaging\ReleaseArtifactPolicy.psm1'
$testRunPlanModule = Join-Path $repoRoot 'Tools\Modules\Testing\TestRunPlan.psm1'
$buildEvidenceModule = Join-Path $repoRoot 'Tools\Modules\Build\BuildEvidence.psm1'

Import-Module $testRunPlanModule -Force -ErrorAction Stop
Import-Module (Join-Path $repoRoot 'Tools\Modules\Testing\TestSandbox.psm1') -Force -ErrorAction Stop
Import-Module $policyModule -Force -ErrorAction Stop
Import-Module $buildEvidenceModule -Force -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot 'TestSupport.psm1') -Force -ErrorAction Stop
Add-Type -AssemblyName System.IO.Compression.FileSystem

function New-RSReleasePolicyTestRoot {
    return (New-RSTestSandboxScratchDirectory `
            -RepoRoot $repoRoot `
            -Harness 'tools-pester' `
            -Case "release-policy-$([Guid]::NewGuid().ToString('N'))")
}

Describe 'Project CI engine path resolution' {
    It 'resolves relative engine paths from the PowerShell filesystem location and preserves absolute paths' {
        $scratch = New-RSReleasePolicyTestRoot
        $psLocation = Join-Path $scratch 'powershell'
        $nativeLocation = Join-Path $scratch 'native'
        foreach ($entry in @(@($psLocation, 'EXPECTED_ENGINE'), @($nativeLocation, 'WRONG_ENGINE'))) {
            $moduleDir = Join-Path $entry[0] 'engine/Tools/Modules'
            [void](New-Item -ItemType Directory -Path $moduleDir -Force)
            [IO.File]::WriteAllText((Join-Path $moduleDir 'Ci.Policy.psm1'), "throw '$($entry[1])'", [Text.UTF8Encoding]::new($false))
        }
        $priorLocation = Get-Location
        $priorNativeLocation = [Environment]::CurrentDirectory
        try {
            [Environment]::CurrentDirectory = $nativeLocation
            Set-Location -LiteralPath $psLocation
            $adapter = Join-Path $repoRoot 'Tools/Invoke-ProjectCi.ps1'
            { & $adapter -Phase Static -EngineRoot './engine' -DevelopmentEngine } | Should Throw 'EXPECTED_ENGINE'
            { & $adapter -Phase Static -EngineRoot (Join-Path $psLocation 'engine') -DevelopmentEngine } | Should Throw 'EXPECTED_ENGINE'
            { & $adapter -Phase Static -EngineRoot 'Env:PATH' -DevelopmentEngine } | Should Throw 'EngineRoot requires the FileSystem provider'
        } finally {
            Set-Location -LiteralPath $priorLocation.Path
            [Environment]::CurrentDirectory = $priorNativeLocation
        }
    }
}

function New-RSFakePeBytes {
    # Minimal PE32+ image: an export section naming $Export and a resource section that holds
    # the ENABLE_TESTS version marker when -TestSurfaceMarker is set.
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet(0x8664, 0xAA64)]
        [int]$Machine,

        [string[]]$Export = @(),

        [switch]$TestSurfaceMarker
    )

    $peOffset = 0x80
    $optionalHeader = $peOffset + 24
    $optionalHeaderSize = 0xF0
    $sectionTable = $optionalHeader + $optionalHeaderSize
    $exportRva = 0x1000
    $resourceRva = 0x2000

    $exportSection = [IO.MemoryStream]::new()
    $exportWriter = [IO.BinaryWriter]::new($exportSection)
    $nameTableRva = $exportRva + 40
    $nameRva = $nameTableRva + (4 * $Export.Count)
    $exportWriter.Write([byte[]]::new(24))
    $exportWriter.Write([uint32]$Export.Count)
    $exportWriter.Write([uint32]0)
    $exportWriter.Write([uint32]$nameTableRva)
    $exportWriter.Write([uint32]0)
    foreach ($name in $Export) {
        $exportWriter.Write([uint32]$nameRva)
        $nameRva += [Text.Encoding]::ASCII.GetByteCount($name) + 1
    }
    foreach ($name in $Export) {
        $exportWriter.Write([Text.Encoding]::ASCII.GetBytes($name))
        $exportWriter.Write([byte]0)
    }
    $exportWriter.Flush()
    $exportBytes = $exportSection.ToArray()

    $resourceText = if ($TestSurfaceMarker) { 'SpecialBuild RedSalamander.TestSurface=ENABLE_TESTS' } else { 'ProductName RedSalamander' }
    $resourceBytes = [byte[]]([byte[]]::new(16) + [Text.Encoding]::Unicode.GetBytes($resourceText) + [byte[]]::new(2))

    $alignUp = { param([int]$value) [int]([Math]::Ceiling($value / 512.0) * 512) }
    $exportRaw = 0x200
    $exportRawSize = & $alignUp ([Math]::Max(1, $exportBytes.Length))
    $resourceRaw = $exportRaw + $exportRawSize
    $resourceRawSize = & $alignUp $resourceBytes.Length

    $bytes = [byte[]]::new($resourceRaw + $resourceRawSize)
    $bytes[0] = 0x4D
    $bytes[1] = 0x5A
    [BitConverter]::GetBytes([int]$peOffset).CopyTo($bytes, 0x3C)
    $bytes[$peOffset] = 0x50
    $bytes[$peOffset + 1] = 0x45
    [BitConverter]::GetBytes([uint16]$Machine).CopyTo($bytes, $peOffset + 4)
    [BitConverter]::GetBytes([uint16]2).CopyTo($bytes, $peOffset + 6)
    [BitConverter]::GetBytes([uint16]$optionalHeaderSize).CopyTo($bytes, $peOffset + 20)
    [BitConverter]::GetBytes([uint16]0x20B).CopyTo($bytes, $optionalHeader)
    [BitConverter]::GetBytes([uint32]16).CopyTo($bytes, $optionalHeader + 108)
    if ($Export.Count -gt 0) {
        [BitConverter]::GetBytes([uint32]$exportRva).CopyTo($bytes, $optionalHeader + 112)
        [BitConverter]::GetBytes([uint32]$exportBytes.Length).CopyTo($bytes, $optionalHeader + 116)
    }
    [BitConverter]::GetBytes([uint32]$resourceRva).CopyTo($bytes, $optionalHeader + 128)
    [BitConverter]::GetBytes([uint32]$resourceBytes.Length).CopyTo($bytes, $optionalHeader + 132)

    $sections = @(
        @{ Name = '.edata'; Rva = $exportRva; Raw = $exportRaw; RawSize = $exportRawSize; Data = $exportBytes },
        @{ Name = '.rsrc'; Rva = $resourceRva; Raw = $resourceRaw; RawSize = $resourceRawSize; Data = $resourceBytes }
    )
    for ($index = 0; $index -lt $sections.Count; ++$index) {
        $header = $sectionTable + (40 * $index)
        $section = $sections[$index]
        [Text.Encoding]::ASCII.GetBytes($section.Name).CopyTo($bytes, $header)
        [BitConverter]::GetBytes([uint32]$section.RawSize).CopyTo($bytes, $header + 8)
        [BitConverter]::GetBytes([uint32]$section.Rva).CopyTo($bytes, $header + 12)
        [BitConverter]::GetBytes([uint32]$section.RawSize).CopyTo($bytes, $header + 16)
        [BitConverter]::GetBytes([uint32]$section.Raw).CopyTo($bytes, $header + 20)
        $section.Data.CopyTo($bytes, $section.Raw)
    }
    return ,$bytes
}

function Add-RSFakeArchiveEntry {
    param(
        [Parameter(Mandatory = $true)]
        [IO.Compression.ZipArchive]$Archive,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [byte[]]$Bytes
    )

    $entryStream = $Archive.CreateEntry($Name).Open()
    try {
        $entryStream.Write($Bytes, 0, $Bytes.Length)
    }
    finally {
        $entryStream.Dispose()
    }
}

function Get-RSThrownMessage {
    param([Parameter(Mandatory = $true)][scriptblock]$Action)

    try {
        & $Action | Out-Null
    }
    catch {
        return $_.Exception.Message
    }
    return $null
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

        [int]$Machine = 0,

        [switch]$TestEnabledApplication,

        [string[]]$PluginExport = @('RedSalamanderCreate', 'RedSalamanderEnumeratePlugins')
    )

    if ($Machine -eq 0) {
        $Machine = if ($Platform -eq 'ARM64') { 0xAA64 } else { 0x8664 }
    }
    $path = Join-Path $Root (Get-RSPortableReleaseFileName -ReleaseVersion $ReleaseVersion -Platform $Platform)
    $stream = [IO.File]::Open($path, [IO.FileMode]::CreateNew, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    $archive = [IO.Compression.ZipArchive]::new($stream, [IO.Compression.ZipArchiveMode]::Create, $false)
    try {
        Add-RSFakeArchiveEntry -Archive $archive -Name 'RedSalamander.exe' `
            -Bytes (New-RSFakePeBytes -Machine $Machine -TestSurfaceMarker:$TestEnabledApplication)
        Add-RSFakeArchiveEntry -Archive $archive -Name 'Plugins/FileSystem.dll' `
            -Bytes (New-RSFakePeBytes -Machine $Machine -Export $PluginExport)
    }
    finally {
        $archive.Dispose()
        $stream.Dispose()
    }

    return $path
}

Describe 'Release artifact fail-closed policy' {
    It 'accepts the exact x64-only portable matrix when ARM64 is disabled' {
        $root = New-RSReleasePolicyTestRoot
        try {
            New-RSFakePortableArtifact -Root $root -ReleaseVersion '7.0.42' -Platform x64 | Out-Null
            $artifacts = @(Assert-RSReleaseArtifacts -ArtifactsRoot $root -ReleaseVersion '7.0.42' -BuildArm64 $false)
            $artifacts.Count | Should Be 1
            $artifacts[0].Platform | Should Be 'x64'
        }
        finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'accepts the exact x64 and ARM64 portable ZIP matrix' {
        $root = New-RSReleasePolicyTestRoot
        try {
            foreach ($platform in @('x64', 'ARM64')) {
                New-RSFakePortableArtifact -Root $root -ReleaseVersion '7.0.42' -Platform $platform | Out-Null
            }
            $artifacts = @(Assert-RSReleaseArtifacts -ArtifactsRoot $root -ReleaseVersion '7.0.42' -BuildArm64 $true)
            $artifacts.Count | Should Be 2
            @($artifacts.Platform | Sort-Object) -join ',' | Should Be 'ARM64,x64'
        }
        finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'rejects a missing requested x64 portable artifact' {
        $root = New-RSReleasePolicyTestRoot
        try {
            New-RSFakePortableArtifact -Root $root -ReleaseVersion '7.0.42' -Platform ARM64 | Out-Null
            { Assert-RSReleaseArtifacts -ArtifactsRoot $root -ReleaseVersion '7.0.42' -BuildArm64 $false } |
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
            { Assert-RSReleaseArtifacts -ArtifactsRoot $root -ReleaseVersion '7.0.42' -BuildArm64 $true } |
                Should Throw 'Release artifact count mismatch. Expected [RedSalamander-7.0.42-ARM64-Portable.zip, RedSalamander-7.0.42-x64-Portable.zip]; found [RedSalamander-7.0.42-x64-Portable.zip].'
        }
        finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'rejects a portable filename whose PE machine does not match its platform' {
        $root = New-RSReleasePolicyTestRoot
        try {
            New-RSFakePortableArtifact -Root $root -ReleaseVersion '7.0.42' -Platform x64 -Machine 0xAA64 | Out-Null
            { Assert-RSReleaseArtifacts -ArtifactsRoot $root -ReleaseVersion '7.0.42' -BuildArm64 $false } |
                Should Throw "Portable artifact 'RedSalamander-7.0.42-x64-Portable.zip' has PE machine 0xAA64; expected 0x8664 for x64."
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
            New-RSReleaseChecksumFile -ArtifactsRoot $root -ReleaseVersion '7.0.42' -BuildArm64 $true | Out-Null
            Assert-RSReleaseChecksumFile -ArtifactsRoot $root -ReleaseVersion '7.0.42' -BuildArm64 $true |
                Should Be $true

            $checksumPath = Join-Path $root 'checksums.sha256'
            $lines = @(Get-Content $checksumPath)
            $lines[0] = ('0' * 64) + $lines[0].Substring(64)
            [IO.File]::WriteAllLines($checksumPath, $lines, [Text.UTF8Encoding]::new($false))
            $expectedMessage = "Release checksum mismatch at line 1. Expected '$((Get-FileHash -LiteralPath (Join-Path $root 'RedSalamander-7.0.42-ARM64-Portable.zip') -Algorithm SHA256).Hash)  RedSalamander-7.0.42-ARM64-Portable.zip'; found '$($lines[0])'."
            { Assert-RSReleaseChecksumFile -ArtifactsRoot $root -ReleaseVersion '7.0.42' -BuildArm64 $true } |
                Should Throw $expectedMessage
        }
        finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

Describe 'Distributable test surface (ENABLE_TESTS) policy' {
    It 'reports reserved test exports and the test-build marker but not production exports' {
        $production = Get-RSPeTestSurface -Name 'FileSystem.dll' -Bytes (New-RSFakePeBytes -Machine 0x8664 -Export @(
                'RedSalamanderCreate', 'RedSalamanderEnumeratePlugins', 'RedSalamanderGetConfigurationSchema',
                'RedSalamanderPluginShutdown', 'RedSalamanderPluginCanUnloadNow',
                'RedSalamanderPluginRetainModuleUntilProcessExit', 'RedSalamanderBrowseConnectionTargets'))
        $production.TestBuildMarker | Should Be $false
        $production.TestExports.Count | Should Be 0

        $testExports = @(
            'RedSalamanderFileSystemTestGetRecycleBinBatchSnapshot',
            'RedSalamanderCurlStartFakeFtpForSelfTest',
            'RedSalamanderS3MultipartSelfTests',
            'RedSalamanderDebugCurlRuntimeProbe',
            'RedSalamanderTerminalDebugCommandSurfaceJoinTid',
            '?DebugUseFixedSettingsBackupTimestampForTest@Settings@Common@@YAX_N@Z')
        $tested = Get-RSPeTestSurface -Name 'FileSystem.dll' -Bytes (New-RSFakePeBytes -Machine 0xAA64 -Export (@('RedSalamanderCreate') + $testExports))
        $tested.TestBuildMarker | Should Be $false
        ($tested.TestExports -join '|') | Should BeExactly ($testExports -join '|')

        $marked = Get-RSPeTestSurface -Name 'RedSalamander.exe' -Bytes (New-RSFakePeBytes -Machine 0x8664 -TestSurfaceMarker)
        $marked.TestBuildMarker | Should Be $true
        $marked.TestExports.Count | Should Be 0
    }

    It 'rejects a release portable ZIP whose plugin exports a test entry point' {
        $root = New-RSReleasePolicyTestRoot
        try {
            New-RSFakePortableArtifact -Root $root -ReleaseVersion '7.0.42' -Platform x64 `
                -PluginExport @('RedSalamanderCreate', 'RedSalamanderFileSystemTestAdmissionProbe') | Out-Null
            $message = Get-RSThrownMessage { Assert-RSReleaseArtifacts -ArtifactsRoot $root -ReleaseVersion '7.0.42' -BuildArm64 $false }
            $message | Should Match 'Distributable package carries a test surface'
            $message | Should Match ([regex]::Escape('Plugins/FileSystem.dll: test exports [RedSalamanderFileSystemTestAdmissionProbe]'))
        }
        finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'rejects a release portable ZIP whose application was built with ENABLE_TESTS' {
        $root = New-RSReleasePolicyTestRoot
        try {
            New-RSFakePortableArtifact -Root $root -ReleaseVersion '7.0.42' -Platform ARM64 -TestEnabledApplication | Out-Null
            $message = Get-RSThrownMessage { Assert-RSReleaseArtifacts -ArtifactsRoot $root -ReleaseVersion '7.0.42' -BuildArm64 $true }
            # The x64 archive is missing, so build the full matrix to reach the content check.
            $message | Should Match 'Release artifact count mismatch'
            New-RSFakePortableArtifact -Root $root -ReleaseVersion '7.0.42' -Platform x64 | Out-Null
            $message = Get-RSThrownMessage { Assert-RSReleaseArtifacts -ArtifactsRoot $root -ReleaseVersion '7.0.42' -BuildArm64 $true }
            $message | Should Match ([regex]::Escape('RedSalamander.exe: built with ENABLE_TESTS'))
        }
        finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'checks staged package directories and fails closed on malformed or missing binaries' {
        $root = New-RSReleasePolicyTestRoot
        try {
            $stage = Join-Path $root 'stage'
            $plugins = Join-Path $stage 'Plugins'
            [void][IO.Directory]::CreateDirectory($plugins)
            [IO.File]::WriteAllBytes((Join-Path $stage 'RedSalamander.exe'), (New-RSFakePeBytes -Machine 0x8664))
            [IO.File]::WriteAllBytes((Join-Path $plugins 'Terminal.dll'), (New-RSFakePeBytes -Machine 0x8664 -Export @('RedSalamanderCreate')))
            Get-RSThrownMessage { Assert-RSShippingPackageHasNoTestSurface -LiteralPath $stage } | Should BeNullOrEmpty

            [IO.File]::WriteAllBytes((Join-Path $plugins 'Terminal.dll'), (New-RSFakePeBytes -Machine 0x8664 -Export @('RedSalamanderTerminalDebugSelfTests')))
            Get-RSThrownMessage { Assert-RSShippingPackageHasNoTestSurface -LiteralPath $stage } |
                Should Match ([regex]::Escape('Plugins\Terminal.dll: test exports [RedSalamanderTerminalDebugSelfTests]'))

            [IO.File]::WriteAllBytes((Join-Path $plugins 'Terminal.dll'), [byte[]](1, 2, 3, 4))
            Get-RSThrownMessage { Assert-RSShippingPackageHasNoTestSurface -LiteralPath $stage } |
                Should Match ([regex]::Escape("PE image 'Plugins\Terminal.dll' does not start with an MZ header."))

            $empty = Join-Path $root 'empty'
            [void][IO.Directory]::CreateDirectory($empty)
            Get-RSThrownMessage { Assert-RSShippingPackageHasNoTestSurface -LiteralPath $empty } |
                Should Match 'contains no \.exe or \.dll to inspect'
        }
        finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'keeps every first-party native C export production-reviewed or in a reserved test-only form' {
        # Test-only exports are found in binaries by name, so a test entry point that escapes the
        # reserved forms would ship undetected. A new production export must be reviewed here.
        $production = @(
            'RedSalamanderCreate', 'RedSalamanderEnumeratePlugins', 'RedSalamanderGetConfigurationSchema',
            'RedSalamanderPluginShutdown', 'RedSalamanderPluginCanUnloadNow',
            'RedSalamanderPluginRetainModuleUntilProcessExit', 'RedSalamanderBrowseConnectionTargets')
        $exportPattern = '(?:__declspec\(dllexport\)|PLUGFACTORY_API)\s+(?:[\w:]+[\s\*&]+)+?(?:__stdcall\s+)?(\w+)\s*\('
        $found = @{}
        foreach ($sourceRoot in @('Common', 'Plugins', 'RedSalamander', 'RedSalamanderMonitor', 'RedSalamanderSearchService', 'RedLauncher', 'RedConfigure')) {
            $path = Join-Path $repoRoot $sourceRoot
            if (-not (Test-Path -LiteralPath $path -PathType Container)) { continue }
            foreach ($file in @(Get-ChildItem -LiteralPath $path -File -Recurse -Include *.cpp, *.h)) {
                foreach ($match in [regex]::Matches([IO.File]::ReadAllText($file.FullName), $exportPattern)) {
                    $found[$match.Groups[1].Value] = $file.FullName.Substring($repoRoot.Length + 1)
                }
            }
        }

        $unreviewed = @($found.Keys | Where-Object {
                $production -cnotcontains $_ -and $_ -cnotmatch 'SelfTest|ForTest|^RedSalamander\w*(Test|Debug)'
            } | Sort-Object | ForEach-Object { "$_ ($($found[$_]))" })
        ($unreviewed -join '; ') | Should BeNullOrEmpty
        @($production | Where-Object { -not $found.ContainsKey($_) }).Count | Should Be 0
        ($found.Count - $production.Count) | Should BeGreaterThan 40
    }

    It 'keeps the portable ZIP packaging path test-disabled and checked' {
        $versionInfo = Get-Content -LiteralPath (Join-Path $repoRoot 'Common\VersionInfo.rc') -Raw
        $versionInfo | Should Match '#if defined\(ENABLE_TESTS\)\s+VALUE "SpecialBuild", L"RedSalamander\.TestSurface=ENABLE_TESTS"\s+#endif'
        $versionInfo | Should Match 'FILEFLAGS \(RS_VERSION_FILEFLAGS \| RS_VERSION_TEST_FILEFLAGS\)'

        $build = Get-Content -LiteralPath (Join-Path $repoRoot 'build.ps1') -Raw
        $build | Should Not Match '\$Msix|\$Msi|RSAppxManifestPath|BuildMsix'
        $build | Should Match '\$Zip'
        $build | Should Match '\$buildParams \+= "/p:RSBuildEnableTests=\$\(\$effectiveTestsEnabled\.ToString\(\)\.ToLowerInvariant\(\)\)"'
        $build | Should Match '(?s)if \(\(\$Zip -or \$GenerateWingetManifest -or \$OfficialRelease\) -and\s+\$receiptTestsEnabled\) \{\s+throw ''A distributable build contains a test-enabled output; no package receipt will be published\.'''
        $zip = Get-Content -LiteralPath (Join-Path $repoRoot 'Installer\zip\build-zip.ps1') -Raw
        $zip | Should Match "Compress-Archive[\s\S]+if \(\`$Configuration -eq 'Release'\) \{\s+Assert-RSShippingPackageHasNoTestSurface -LiteralPath \`$StagedZipPath[\s\S]+Test-RSPortablePackage"

        $releaseWorkflow = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools/Invoke-ProjectCi.ps1') -Raw
        $releaseWorkflow | Should Not Match 'run_full_tests|BuildMsix|\.msix|\.msi'
        $releaseWorkflow | Should Match 'Write-CiProductionReceipt'
    }
}

Describe 'Release workflow source contracts' {
    BeforeAll {
        $releaseWorkflowPath = Join-Path $repoRoot '.ci-engine/.github/workflows/publish.yml'
        $releaseWorkflow = Get-Content -LiteralPath $releaseWorkflowPath -Raw
        $ghosttyStatusWorkflowPath = Join-Path $repoRoot '.github/workflows/ghostty-runtime-status.yml'
        $ghosttyStatusWorkflow = Get-Content -LiteralPath $ghosttyStatusWorkflowPath -Raw
        $ghosttyNativeWorkflowPath = Join-Path $repoRoot '.github/workflows/ghostty-runtime-native-qualification.yml'
        $ghosttyNativeWorkflow = Get-Content -LiteralPath $ghosttyNativeWorkflowPath -Raw
        $ghosttyProductWorkflowPath = Join-Path $repoRoot '.github/workflows/ghostty-runtime-product-qualification.yml'
        $ghosttyProductWorkflow = Get-Content -LiteralPath $ghosttyProductWorkflowPath -Raw
    }

    It 'restores exact tracked bytes through a deep checkout with Git long paths enabled' {
        $fixture = New-RSReleasePolicyTestRoot
        $relative = ((@('a', 'b', 'c', 'd') | ForEach-Object { 'deep-path-' + ($_ * 70) }) -join '/') + '/proof.txt'
        $relative.Length | Should BeGreaterThan 260
        $file = Join-Path $fixture $relative
        $file.Length | Should BeGreaterThan 260
        [void][IO.Directory]::CreateDirectory((Split-Path -Parent $file))
        [IO.File]::WriteAllText($file, 'tracked deep-path fixture', [Text.UTF8Encoding]::new($false))
        $expectedHash = (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash
        & git -c core.longpaths=true -C $fixture init --quiet
        $LASTEXITCODE | Should Be 0
        & git -c core.longpaths=true -C $fixture add -- $relative
        $LASTEXITCODE | Should Be 0
        [IO.File]::Delete($file)
        $directory = Split-Path -Parent $file
        foreach ($level in 1..4) {
            [IO.Directory]::Delete($directory, $false)
            $directory = Split-Path -Parent $directory
        }
        [IO.File]::Exists($file) | Should Be $false
        & git -c core.longpaths=true -C $fixture checkout -- $relative
        $LASTEXITCODE | Should Be 0
        (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash | Should Be $expectedHash
    }

    It 'configures long paths before the scheduled Windows checkouts and bounds flake evidence uploads' {
        foreach ($relative in @('.github/workflows/nightly-flake.yml', '.github/workflows/ghostty-runtime-status.yml')) {
            $workflow = Get-Content -LiteralPath (Join-Path $repoRoot $relative) -Raw
            $setup = $workflow.IndexOf('git config --global core.longpaths true')
            $checkout = $workflow.IndexOf('uses: actions/checkout@')
            $setup | Should BeGreaterThan -1
            $checkout | Should BeGreaterThan $setup
            $workflow | Should Not Match 'continue-on-error:\s*true'
        }
        $flake = Get-Content -LiteralPath (Join-Path $repoRoot '.github/workflows/nightly-flake.yml') -Raw
        $flake | Should Match 'native_runner:\s*windows-2025-vs2026'
        $flake | Should Match 'runs-on:\s*windows-2025-vs2026'
        $flake.IndexOf('uses: ./.github/actions/setup-msvc') | Should BeGreaterThan -1
        $flake.IndexOf('uses: ./.github/actions/setup-msvc') | Should BeLessThan $flake.IndexOf('name: Verify and rebind portable build attestation')
          $flake.IndexOf('./Tools/Restore-DxUi.ps1 -Platform x64 -Configuration Debug') | Should BeGreaterThan -1
          $flake.IndexOf('./Tools/Restore-DxUi.ps1 -Platform x64 -Configuration Debug') | Should BeLessThan $flake.IndexOf('Import-RSPortableBuildAttestation')
        $flake | Should Match 'Initialize-RSTestSandboxRoot -TestRoot \$root -RepoRoot \$env:GITHUB_WORKSPACE -AllowExternalInitialization'
        $flake.IndexOf('Initialize-RSTestSandboxRoot') | Should BeLessThan $flake.IndexOf('name: Run nightly shuffle repeat lane')
        $flake | Should Match 'if:\s*\$\{\{ always\(\) && steps\.test-root\.outputs\.path != '''' \}\}'
        $flake | Should Match 'if-no-files-found:\s*error'
        $uploadStart = $flake.IndexOf('name: Upload nightly shuffle artifacts', [StringComparison]::Ordinal)
        $uploadStart | Should BeGreaterThan -1
        $upload = $flake.Substring($uploadStart)
        $upload | Should Match 'include-hidden-files:\s*true'
        $upload | Should Match '(?m)^\s+path:\s*\|\s*$'
        foreach ($suffix in @('/evidence/', '/runs/*/artifacts/selftest/', '/runs/*/artifacts/native-process-exits/', '/runs/*/artifacts/selftest/**/work/**',
                '/evidence/completed-runs/*/scratch/**', '/evidence/completed-runs/*/artifacts/selftest/**/work/**',
                '/**/*.vhd', '/**/*.vhdx', '/**/*.avhdx')) {
            $prefix = if ($suffix -in @('/evidence/', '/runs/*/artifacts/selftest/', '/runs/*/artifacts/native-process-exits/')) { '' } else { '!' }
            $upload | Should Match ([regex]::Escape($prefix + '${{ steps.test-root.outputs.path }}' + $suffix))
        }
        $upload | Should Not Match '(?m)^\s+path:\s*\$\{\{ steps\.test-root\.outputs\.path \}\}[\\/]?\s*$'
        $flake | Should Match 'Import-RSPortableBuildAttestation'
        $flake | Should Match '-ExpectedAttestationId'
        $flake | Should Match '-SelfTestRepeat 5 -SelfTestShuffleSeed \$seed -ClassifyFailures'
        $flake | Should Match '-AllowExternalTestRoot -AllowAlternateVolumeTestRoot'
        $flake | Should Match 'Invoke-RSWithHostedTestDesktop -Action'
          $flake | Should Match '(?s)Invoke-RSWithHostedTestDesktop -Action \{.*?\.\\Tools\\Run-AllTests\.ps1[^\r\n]*\r?\n\s*if \(\$LASTEXITCODE -ne 0\) \{ throw "Nightly shuffle suite failed'
    }

      It 'checks native exit status inside the protected action before reporting success' {
          $build = Get-Content -LiteralPath (Join-Path $repoRoot '.github/workflows/build-reusable.yml') -Raw
          $build | Should Match '(?s)\$runTests = \{\s*\./Tools/Run-AllTests\.ps1[^\r\n]*\r?\n\s*if \(\$LASTEXITCODE -ne 0\) \{ throw "Native'
          $build | Should Match '(?s)Invoke-RSWithHostedTestDesktop -Action \{.*?& \$runTests.*?exit 0'
          $uploadStart = $build.IndexOf('name: Upload native qualification evidence', [StringComparison]::Ordinal)
          $uploadEnd = $build.IndexOf('name: Upload native build diagnostics', $uploadStart, [StringComparison]::Ordinal)
          ($uploadStart -ge 0 -and $uploadEnd -gt $uploadStart) | Should Be $true
          $build.Substring($uploadStart, $uploadEnd - $uploadStart) | Should Match ([regex]::Escape('${{ steps.test_root.outputs.path }}/runs/*/artifacts/native-process-exits/'))
      }

    It 'selects and restores the hosted request inside all three actual direct workflow actions' {
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/HostedTestDesktop.psm1') -Force
        $saved = @{}
        foreach ($name in @('GITHUB_ACTIONS','RUNNER_ENVIRONMENT','GITHUB_REPOSITORY','REDSALAMANDER_HOSTED_TEST_DESKTOP')) {
            $saved[$name] = [Environment]::GetEnvironmentVariable($name, 'Process')
        }
        $priorExit = Get-Variable LASTEXITCODE -Scope Global -ValueOnly -ErrorAction SilentlyContinue
        function Invoke-RSWithHostedTestDesktop { param([scriptblock]$Action) & $Action }
        try {
            foreach ($relative in @('.github/workflows/nightly-flake.yml', '.github/workflows/build-reusable.yml', '.github/workflows/ghostty-runtime-product-qualification.yml')) {
                $text = Get-Content -LiteralPath (Join-Path $repoRoot $relative) -Raw
                $start = $text.IndexOf('Invoke-RSWithHostedTestDesktop -Action {', [StringComparison]::Ordinal)
                $endMarker = if ($relative -like '*nightly*') { '          exit 0' } elseif ($relative -like '*ghostty*') { '      - name: Build production Release ARM64 product' } else { '          } else {' }
                $end = $text.IndexOf($endMarker, $start, [StringComparison]::Ordinal)
                ($start -ge 0 -and $end -gt $start) | Should Be $true
                $actionText = $text.Substring($start, $end - $start)
                if ($relative -like '*nightly*') {
                    $nativeLine = '(?m)^\s*\.\\Tools\\Run-AllTests\.ps1[^\r\n]*'
                    ([regex]::Matches($actionText, $nativeLine).Count) | Should Be 1
                    $actionText = [regex]::Replace($actionText, $nativeLine, '              & $fakeNative')
                } elseif ($relative -like '*ghostty*') {
                    $nativeLine = '(?ms)^\s*\.\\Tools\\Run-AllTests\.ps1\s*`\s*\r?\n.*?^\s*-BuildNumber \(\[int\]\$env:RS_BUILD_NUMBER\)'
                    ([regex]::Matches($actionText, $nativeLine).Count) | Should Be 1
                    $actionText = [regex]::Replace($actionText, $nativeLine, '                & $fakeNative')
                }
                $action = [scriptblock]::Create($actionText)
                foreach ($hosted in @($true, $false)) {
                    foreach ($fail in @($false, $true)) {
                      foreach ($prior in @('default', 'prior-value', '')) {
                        $env:GITHUB_ACTIONS = 'true'
                        $env:RUNNER_ENVIRONMENT = if ($hosted) { 'github-hosted' } else { 'self-hosted' }
                        $env:GITHUB_REPOSITORY = 'RedSalamanders/RedSalamander'
                        [Environment]::SetEnvironmentVariable('REDSALAMANDER_HOSTED_TEST_DESKTOP', $(if ($prior) { $prior } else { $null }), 'Process')
                        $observed = [Collections.Generic.List[string]]::new()
                        $fakeNative = {
                            $observed.Add([string][Environment]::GetEnvironmentVariable('REDSALAMANDER_HOSTED_TEST_DESKTOP', 'Process'))
                            if ($fail) { $global:LASTEXITCODE = 9; throw 'seeded native failure' }
                            $global:LASTEXITCODE = 0
                        }
                        $runTests = $fakeNative
                        $result = Invoke-RSTestBlock { & $action }
                        $result.Succeeded | Should Be (-not $fail)
                        if ($fail) { $result.Error.Exception.Message | Should Match 'seeded native failure' }
                        $expectedCount = if ($relative -like '*ghostty*' -and -not $fail) { 3 } else { 1 }
                        $observed.Count | Should Be $expectedCount
                        foreach ($actual in $observed) { $actual | Should Be $(if ($hosted) { 'default' } else { '' }) }
                        ([string][Environment]::GetEnvironmentVariable('REDSALAMANDER_HOSTED_TEST_DESKTOP', 'Process')) | Should Be $prior
                      }
                    }
                }
            }
        } finally {
            foreach ($name in $saved.Keys) { [Environment]::SetEnvironmentVariable($name, $saved[$name], 'Process') }
            $global:LASTEXITCODE = $priorExit
        }
    }

    It 'keeps current artifact downloaders pinned and the exact sealed Gate0 workflow intact' {
        $downloaders = @(Get-ChildItem -LiteralPath (Join-Path $repoRoot '.github/workflows') -Filter '*.yml' |
            Select-String -Pattern 'uses:\s*actions/download-artifact@(\S+)')
        $downloaders.Count | Should BeGreaterThan 0
        foreach ($downloader in $downloaders) {
            $expectedPin = '9000827ccba6bdab643e8b6fd33ac0654aef8333'
            if ($downloader.Filename -ceq 'terminal-engine-gate0.yml') {
                # Historical pins are accepted only with the exact sealed workflow bytes.
                (Get-FileHash -Algorithm SHA256 -LiteralPath $downloader.Path).Hash.ToLowerInvariant() |
                    Should Be 'e6f15b6fef9c0d367e12dc7f43cb769106dc982934658f501120728599f32d1c'
                $expectedPin = '3e5f45b2cfb9172054b4087a40e8e0b5a5461e7c'
            }
            $downloader.Matches[0].Groups[1].Value | Should Be $expectedPin
        }
    }

    It 'keeps current artifact uploaders pinned and the exact sealed Gate0 workflow intact' {
        $uploaders = @(Get-ChildItem -LiteralPath (Join-Path $repoRoot '.github/workflows') -Filter '*.yml' |
            Select-String -Pattern 'uses:\s*actions/upload-artifact@(\S+)')
        $uploaders.Count | Should BeGreaterThan 0
        foreach ($uploader in $uploaders) {
            $expectedPin = 'cf430e030ddbb5b0abf93d22962f4752f3646cd9'
            if ($uploader.Filename -ceq 'terminal-engine-gate0.yml') {
                (Get-FileHash -Algorithm SHA256 -LiteralPath $uploader.Path).Hash.ToLowerInvariant() |
                    Should Be 'e6f15b6fef9c0d367e12dc7f43cb769106dc982934658f501120728599f32d1c'
                $expectedPin = '043fb46d1a93c77aae656e7c1c64a875d1fc6a0a'
            }
            $uploader.Matches[0].Groups[1].Value | Should Be $expectedPin
        }
    }

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
        # Public CI and local publication use the portable ZIP.
        $restoreScript = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools/Restore-DxUi.ps1') -Raw
        $restoreScript | Should Match "OSArchitecture\.ToString\(\) -eq 'Arm64'\) \{\s*\r?\n\s*'MSBuild/Current/Bin/arm64/MSBuild\.exe'"
        $evidence = [regex]::Match($workflow, '(?s)- name: Upload native qualification evidence(.*?)(?=      - name:)').Groups[1].Value
        $evidence | Should Match 'steps\.test_root\.outputs\.path'
        $evidence | Should Not Match '\.build/'
        $diagnostics = [regex]::Match($workflow, '(?s)- name: Upload native build diagnostics(.*?)(?=      - name:)').Groups[1].Value
        $diagnostics | Should Match 'include-hidden-files: true'
        $diagnostics | Should Not Match 'steps\.test_root\.outputs\.path'
    }

    It 'uses normal dependency success semantics and no permissive artifact download' {
        $releaseWorkflow | Should Not Match 'continue-on-error:\s*true'
        $native = Get-Content (Join-Path $repoRoot '.ci-engine/.github/workflows/native.yml') -Raw
        $native | Should Match 'needs: \[admission, test-gate\]'
        $native | Should Match "if-no-files-found: error"
        $native | Should Not Match 'continue-on-error:\s*true'
        $artifacts = Get-Content (Join-Path $repoRoot '.ci-engine/Tools/Modules/Ci.Artifacts.psm1') -Raw
        $artifacts | Should Match 'Test-CiRegressionReceipt'
        $artifacts | Should Match 'Get-FileHash'
    }

    It 'uploads deterministic package names and fails when an upload path is missing' {
        $native = Get-Content (Join-Path $repoRoot '.ci-engine/.github/workflows/native.yml') -Raw
        $adapter = Get-Content (Join-Path $repoRoot 'Tools/Invoke-ProjectCi.ps1') -Raw
        $package = Get-Content (Join-Path $repoRoot 'Installer/zip/build-zip.ps1') -Raw
        $native | Should Match 'name: production-'
        $native | Should Match 'archive_directory'
        $native | Should Match 'if-no-files-found: error'
        $adapter | Should Match 'Installer/zip/build-zip.ps1'
        $adapter | Should Match 'Write-CiProductionReceipt'
        $package | Should Match 'RedSalamander-\$Version-\$Platform-Portable\.zip|\$zipName'
        $publication = Get-Content (Join-Path $repoRoot '.ci-engine/Tools/Modules/Ci.Publication.psm1') -Raw
        $publication | Should Match 'Existing release contains unqualified assets'
        $publication | Should Match 'Existing public release is incomplete'
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
        $releaseWorkflow | Should Match '(?s)name: Retain Ghostty publication preflight.*?if: always\(\).*?path: \.build/TerminalEngineStatus/publication/status\.json'
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
        $ghosttyNativeWorkflow | Should Match 'build --color off @seedArguments -j1 test-lib-vt'
        $ghosttyNativeWorkflow | Should Match 'use_repository_lock:'
        $ghosttyNativeWorkflow | Should Match 'RS_USE_REPOSITORY_LOCK -ceq ''true'''
        $ghosttyNativeWorkflow | Should Match 'External/TerminalEngine/GhosttyRuntimeLock.v1.json'
        $ghosttyNativeWorkflow | Should Match 'ZIG_GLOBAL_CACHE_DIR'
        $ghosttyNativeWorkflow | Should Match 'Expand-RSGhosttyRuntimeLockedPackages'
        $ghosttyNativeWorkflow | Should Match '@\(''--system'', \$packages\)'
        $ghosttyNativeWorkflow | Should Match 'Initialize-RSTestSandboxRoot'
        $ghosttyNativeWorkflow | Should Match 'New-RSTestSandboxScratchDirectory'
        $ghosttyNativeWorkflow | Should Match 'setCwd'
        $ghosttyNativeWorkflow | Should Match 'addArg\(`"--seed=\$env:RS_TEST_SEED'
        $ghosttyNativeWorkflow | Should Match 'src/terminal/snapshot/testdata'
        $ghosttyNativeWorkflow | Should Match 'RS_NATIVE_TEST_SCRATCH.*native-arm64-qualification.json'
        $ghosttyNativeWorkflow | Should Not Match 'runner.temp.*ghostty-native-crash'
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
        $ghosttyNativeWorkflow | Should Match '\$nativeLocalCache = \[IO\.Path\]::GetFullPath\(\(Join-Path \$candidateRoot ''native-local-cache''\)\)'
        $ghosttyNativeWorkflow | Should Match '\$env:ZIG_LOCAL_CACHE_DIR = \$nativeLocalCache'
        $ghosttyNativeWorkflow | Should Match '\$allowedCacheRoot = \$nativeLocalCache \+ \[IO\.Path\]::DirectorySeparatorChar'
        $ghosttyNativeWorkflow | Should Match '\.Message\.Contains\(\$nativeLocalCache, \[StringComparison\]::OrdinalIgnoreCase\)'
        $ghosttyNativeWorkflow | Should Not Match 'Join-Path \$sourceRoot ''\.zig-cache'''
        $ghosttyNativeWorkflow | Should Match 'Start-Sleep -Seconds 2'
        $ghosttyNativeWorkflow | Should Match '& \$failedTestPath "--seed=\$env:RS_TEST_SEED"'
        $ghosttyNativeWorkflow | Should Match 'Windows Error Reporting\\LocalDumps\\\$_'
        $ghosttyNativeWorkflow | Should Match "@\('test.exe', 'ghostty-vt-test.exe', 'ghostty-vt-c-test.exe'\)"
        $ghosttyNativeWorkflow | Should Match 'Existing WER configuration must be preserved'
        $ghosttyNativeWorkflow | Should Match '\[Microsoft\.Win32\.Registry\]::CurrentUser\.CreateSubKey\('
        $ghosttyNativeWorkflow | Should Match '\$dumpParent\.Dispose\(\)'
        $ghosttyNativeWorkflow | Should Match 'New-Item -Path \$localDumpKey\)'
        $ghosttyNativeWorkflow | Should Not Match 'New-Item -Path \$localDumpKey -Force'
        $ghosttyNativeWorkflow | Should Match 'DumpType'
        $ghosttyNativeWorkflow | Should Match 'test_zcu\.obj'
        $ghosttyNativeWorkflow | Should Match 'retryExitCode = \$retryExitCode'
        $ghosttyNativeWorkflow | Should Match 'ghostty-native-arm64-crash-\$\{\{ github\.run_id \}\}'
        $ghosttyNativeWorkflow | Should Match 'if:\s*\$\{\{ failure\(\) && !inputs\.diagnostics_only && steps\.native_tests\.outcome == ''failure'' && env\.RS_NATIVE_TEST_SCRATCH != '''' \}\}'
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
        $nativeEvidence = [regex]::Match($ghosttyNativeWorkflow, '(?s)- name: Upload native qualification evidence(.*?)(?=\r?\n  qualify-product:)').Groups[1].Value
        $nativeEvidence | Should Not Match '\.build/'
        ([regex]::Matches($nativeEvidence, '\$\{\{ env\.RS_NATIVE_TEST_SCRATCH \}\}/')).Count | Should Be 3
        $ghosttyNativeWorkflow | Should Match 'actions/upload-artifact@cf430e030ddbb5b0abf93d22962f4752f3646cd9'
        $ghosttyNativeWorkflow | Should Match 'if-no-files-found:\s*error'
        $ghosttyNativeWorkflow | Should Match '(?ms)^  qualify-product:.*?if:\s*\$\{\{ inputs\.qualify_product && !inputs\.diagnostics_only \}\}.*?needs:\s*qualify-arm64'
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
        $ghosttyProductWorkflow | Should Match 'Import-Module.*VcpkgInstallSafety\.psm1'
        $ghosttyProductWorkflow | Should Match 'Initialize-RSVcpkgToolContext -RepoRoot \$env:GITHUB_WORKSPACE'
        $ghosttyProductWorkflow | Should Match '\(''VCPKG_REGISTRY_ROOT='' \+ \$context\.RegistryRoot\)'
        $ghosttyProductWorkflow | Should Match '(?s)vcpkg-install\.ps1.*?-Platform ARM64.*?-VcpkgExe \.\\\.build\\vcpkg-tool\\vcpkg\.exe.*?-VcpkgRoot \.\\\.build\\vcpkg-registry'
        $ghosttyProductWorkflow | Should Not Match '\$env:VCPKG_ROOT\s*='
        $ghosttyProductWorkflow.IndexOf('Bootstrap the pinned vcpkg tool') | Should BeLessThan $ghosttyProductWorkflow.IndexOf('Install pinned ARM64 dependencies')
        $ghosttyProductWorkflow.IndexOf('Initialize-RSVcpkgToolContext') | Should BeLessThan $ghosttyProductWorkflow.IndexOf('.\vcpkg-install.ps1')
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
        $ghosttyProductWorkflow | Should Match 'actions/upload-artifact@cf430e030ddbb5b0abf93d22962f4752f3646cd9'
        $ghosttyProductWorkflow | Should Match '\.build/TerminalEngine/Product/ARM64/terminal-runtime-identity\.json'
        $ghosttyProductWorkflow | Should Match 'if-no-files-found:\s*error'
        $ghosttyProductWorkflow | Should Not Match '\.build/ARM64/(Debug|Release)/terminal-runtime-identity\.json'
    }

    It 'archives native product receipt pointers and immutable records from the canonical profile stores' {
        $workflow = Get-Content -LiteralPath (Join-Path $repoRoot '.github/workflows/ghostty-runtime-product-qualification.yml') -Raw
        $upload = [regex]::Match($workflow, '(?s)- name: Upload native product qualification evidence(.*?)(?=\r?\n      # The test root)').Groups[1].Value
        $paths = @([regex]::Matches($upload, '(?m)^ {12}(\.build/[^\r\n]+)') | ForEach-Object { $_.Groups[1].Value })
        $module = Get-Module | Where-Object Path -EQ $buildEvidenceModule | Select-Object -First 1
        ($null -ne $module) | Should Be $true
        foreach ($configuration in @('Debug','Release')) {
            $profile = Join-Path $TestDrive ('.build/ARM64/' + $configuration)
            $store = & $module { param($directory) Get-RSBuildReceiptStore -BuildOutputDir $directory } $profile
            $pointer = [IO.Path]::GetRelativePath($TestDrive, $store.PointerPath).Replace('\','/')
            $records = [IO.Path]::GetRelativePath($TestDrive, $store.ReceiptDirectory).Replace('\','/') + '/'
            ($paths -ccontains $pointer) | Should Be $true
            ($paths -ccontains $records) | Should Be $true
            ($paths -ccontains ('.build/ARM64/' + $configuration + '/toolchain-identity.json')) | Should Be $true
        }
        $upload | Should Not Match '\.build/BuildReceipts/'
    }

    It 'keeps the source-and-lock-bound online Ghostty preflight in read-only publication preparation' {
        $publication = Get-Content (Join-Path $repoRoot '.ci-engine/Tools/Modules/Ci.Publication.psm1') -Raw
        $prepare = Get-Content (Join-Path $repoRoot '.ci-engine/Tools/Invoke-Publication.ps1') -Raw
        $publication | Should Match 'Get-GhosttyTerminalRuntimeStatus\.ps1'
        $publication | Should Match "'-Mode','Online'"
        $publication | Should Match '''-ExpectedSourceCommit'',\$Plan\.batch\.commit'
        $publication | Should Match 'sourceBinding\.repositoryCommit -cne \$Plan\.batch\.commit'
        $publication | Should Match 'canonicalLock\.sha256 -cne \$lockHash'
        $publication | Should Match 'security-review.*required-manual.*unreachable'
        $prepare | Should Match 'Test-CiProjectPublicationPreflight'
        $prepare | Should Not Match 'build\.ps1|Build-GhosttyTerminalRuntime'
        $native = Get-Content (Join-Path $repoRoot '.ci-engine/.github/workflows/native.yml') -Raw
        $native | Should Not Match 'ghostty-product-qualification|ghostty-runtime-native-qualification'
    }

    It 'publishes only an explicitly selected qualified ZIP matrix with a separate winget action' {
        $caller = Get-Content (Join-Path $repoRoot '.github/workflows/publish.yml') -Raw
        $caller | Should Match 'workflow_dispatch:'
        $caller | Should Not Match '(?m)^  (push|release|schedule):'
        $caller | Should Match 'default: Prepare'
        $caller | Should Match 'cancel-in-progress: false'
        $releaseWorkflow | Should Match "if: inputs.operation == 'Release'"
        $releaseWorkflow | Should Match "if: inputs.operation == 'Winget'"
        $releaseWorkflow | Should Match 'needs: \[prepare, install\]'
        $publication = Get-Content (Join-Path $repoRoot '.ci-engine/Tools/Modules/Ci.Publication.psm1') -Raw
        $publication | Should Match 'draft=\$true'
        $publication | Should Match "sha256:"
        $publication | Should Match 'Assert-CiPublicationCurrent'
        $publication | Should Match 'Existing version tag targets different source'
    }

    It 'separates read-only project execution from trusted state, format and publication writers' {
        $pr = Get-Content (Join-Path $repoRoot '.ci-engine/.github/workflows/pr.yml') -Raw
        $pr | Should Match '(?ms)^permissions:\s+contents:\s*read'
        $pr | Should Match 'persist-credentials:\s*false'
        $pr | Should Not Match 'contents:\s*write|git\s+push|checks:\s*write'
        $releaseWorkflow | Should Match '(?ms)^permissions:\s+contents:\s*read'
        $releaseWorkflow | Should Match '(?ms)^  release:.*?permissions:\s+contents:\s*write'
        $releaseWorkflow | Should Not Match '(?ms)^permissions:\s+contents:\s*write'
        $releaseWorkflow | Should Match 'Publish-Release.ps1'
        $releaseJob = [regex]::Match($releaseWorkflow,'(?ms)^  release:.*?(?=^  [A-Za-z0-9_-]+:|\z)').Value
        $releaseJob | Should Not Match 'manifest.adapter|Invoke-ProjectCi|build.ps1'
        $format = Get-Content (Join-Path $repoRoot '.github/workflows/format.yml') -Raw
        $format | Should Match "github.ref == 'refs/heads/main'"
        $format | Should Match 'Invoke-Formatting.ps1'
        $admission = Get-Content (Join-Path $repoRoot '.ci-engine/Tools/Modules/Ci.GitHub.psm1') -Raw
        $admission | Should Match 'force = \$false'
    }

    It 'preserves Terminal cache identity and restricts ABI-package fallback to its bounded dependency bag' {
        $buildWorkflow = Get-Content (Join-Path $repoRoot '.github/workflows/build-reusable.yml') -Raw
        $buildWorkflow | Should Match 'visualstudio2026buildtools -PackageVersion 118\.8\.1'
        $buildWorkflow | Should Match 'visualstudio2026-workload-vctools -PackageVersion 1\.0\.0'
        $buildWorkflow | Should Match 'Cache pinned Terminal runtime inputs and product'
        $cacheKeyMatch = [regex]::Match($buildWorkflow, "(?m)^\s*key:.*-terminal-runtime-.*?hashFiles\((?<inputs>[^)]*)\)")
        $cacheKeyMatch.Success | Should Be $true
        $actualInputs = @([regex]::Matches($cacheKeyMatch.Groups['inputs'].Value, "'(?<path>[^']+)'") | ForEach-Object { $_.Groups['path'].Value })
        $expectedInputs = @((Get-RSGhosttyRuntimeIdentityInputPaths) | ForEach-Object { $_.Replace('\','/') } | ForEach-Object {
            if ($_ -ceq 'External/TerminalEngine/Notices/uucode-bjoern-hoehrmann.txt') { 'External/TerminalEngine/Notices/*' }
            elseif ($_ -cne 'External/TerminalEngine/Notices/unicode-license-3.0-2025.txt') { $_ }
        })
        @($actualInputs | Sort-Object -Unique).Count | Should Be $actualInputs.Count
        (($actualInputs | Sort-Object) -join [char]0) | Should Be (($expectedInputs | Sort-Object) -join [char]0)
        $buildWorkflow | Should Match 'Get-RSBuildToolchainIdentity'
        $buildWorkflow | Should Match 'ToolchainIdentity\.schema\.json'
        $native = Get-Content (Join-Path $repoRoot '.ci-engine/.github/workflows/native.yml') -Raw
        $native | Should Match 'actions/cache/restore@[a-f0-9]{40}'
        $native | Should Match 'actions/cache/save@[a-f0-9]{40}'
        $native | Should Match "'Tools/\*\*'"
        foreach ($workflow in @(
            @{text=$native; expected=2},
            @{text=(Get-Content (Join-Path $repoRoot '.ci-engine/.github/workflows/qualification.yml') -Raw); expected=1}
        )) {
            $fallbackSteps = @([regex]::Matches($workflow.text, '(?ms)^      - .*?(?=^      - |\z)') |
                Where-Object { $_.Value -match '(?m)^          restore-keys:' })
            $fallbackSteps.Count | Should Be $workflow.expected
            foreach ($step in $fallbackSteps) {
                $step.Value | Should Match '(?m)^(?:      - |        )uses: actions/cache/restore@[a-f0-9]{40}'
                $step.Value | Should Match '(?m)^          path: \.build/ci/vcpkg-cache\r?$'
                $prefix = [regex]::Match($step.Value, '(?m)^          restore-keys: \|\r?\n(?<keys>(?: {12}[^\r\n]*\r?\n)+)')
                $prefix.Success | Should Be $true
                $prefix.Groups['keys'].Value.Trim() | Should Be 'ci-deps-${{ runner.os }}-${{ runner.arch }}-${{ matrix.platform }}-'
                $step.Value | Should Not Match 'path: \.build/(TerminalEngine|dependencies|x64|ARM64)'
            }
            $workflow.text | Should Match 'VCPKG_BINARY_SOURCES: clear;files,'
        }
    }

    It 'executes the CI matrix gate through the verified short source from a foreign directory' {
        $workflowLines = @(Get-Content (Join-Path $repoRoot '.github/workflows/build-reusable.yml'))
        $start = @(for ($index = 0; $index -lt $workflowLines.Count; $index++) {
            if ($workflowLines[$index] -match "^\s+Import-Module .*'Tools/Modules/Build/DxUiDependency[.]psm1'") { $index }
        })
        $start.Count | Should Be 1
        $block = ($workflowLines[$start[0]..($start[0] + 3)] -replace '^          ', '') -join "`n"
        $block | Should Match 'Get-RSDxUiSourceIdentity'
        $block | Should Match 'validate-build-matrix[.]ps1'
        $fixture = New-RSReleasePolicyTestRoot
        $script = Join-Path $fixture 'Invoke-CiMatrixGate.ps1'
        [IO.File]::WriteAllText($script, ("`$ErrorActionPreference = 'Stop'`n" + $block + "`nif ((Split-Path -Leaf `$dxUiSourceIdentity.Source).Length -ne 16) { throw 'CI did not use the short source locator.' }`n"))
        $startInfo = [Diagnostics.ProcessStartInfo]::new((Join-Path $PSHOME $(if ($IsWindows) { 'pwsh.exe' } else { 'pwsh' })))
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true
        $startInfo.WorkingDirectory = $fixture
        $startInfo.Environment['GITHUB_WORKSPACE'] = $repoRoot
        foreach ($argument in @('-NoProfile', '-NonInteractive', '-File', $script)) { $startInfo.ArgumentList.Add($argument) }
        $process = [Diagnostics.Process]::new()
        $process.StartInfo = $startInfo
        try {
            [void]$process.Start()
            $stdout = $process.StandardOutput.ReadToEndAsync()
            $stderr = $process.StandardError.ReadToEndAsync()
            if (-not $process.WaitForExit(30000)) { $process.Kill($true); throw 'Owned read-only CI matrix gate timed out.' }
            $output = $stdout.GetAwaiter().GetResult() + $stderr.GetAwaiter().GetResult()
            [IO.File]::WriteAllText((Join-Path $fixture 'matrix-gate.log'), $output)
            $process.ExitCode | Should Be 0
            $output | Should Match '0 six-configuration mapping errors'
        } finally { $process.Dispose() }
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

    It 'passes each native test seed option as complete argv elements in both Zig branches' {
        $workflow = Get-Content -LiteralPath (Join-Path $repoRoot '.github/workflows/ghostty-runtime-native-qualification.yml') -Raw
        $assignment = [regex]::Match($workflow, '(?m)^[ \t]*(?:\[string\[\]\][ \t]*)?\$seedArguments[ \t]*=[ \t]*if[^\r\n]+')
        $assignment.Success | Should Be $true
        $echo = Join-Path $TestDrive 'echo-native-seed.ps1'
        '[Console]::WriteLine((ConvertTo-Json -InputObject @($args) -Compress))' | Set-Content -LiteralPath $echo -Encoding utf8NoBOM
        $invoke = [scriptblock]::Create('param($child,$echo,[bool]$modernZig)' + "`n" + $assignment.Value + "`n" +
            '& $child -NoLogo -NoProfile -NonInteractive -File $echo build --color off @seedArguments -j1')
        $previousSeed = $env:RS_TEST_SEED
        try {
            $env:RS_TEST_SEED = '0x97d35bc7'
            foreach ($modern in @($true, $false)) {
                $json = & $invoke (Join-Path $PSHOME 'pwsh.exe') $echo $modern
                $LASTEXITCODE | Should Be 0
                $actual = @($json | ConvertFrom-Json)
                $expected = if ($modern) {
                    @('build','--color','off','--seed=0x97d35bc7','-j1')
                } else {
                    @('build','--color','off','--seed','0x97d35bc7','-j1')
                }
                ($actual -join '|') | Should Be ($expected -join '|')
            }
        }
        finally { $env:RS_TEST_SEED = $previousSeed }
    }

    It 'admits absolute and relative diagnostic executables only inside the exact native cache' {
        $workflow = Get-Content -LiteralPath (Join-Path $repoRoot '.github/workflows/ghostty-runtime-native-qualification.yml') -Raw
        $guard = [regex]::Match($workflow, '(?ms)^ {16}\$failedTestPath =.*?(?=^ {16}# Zig truncates)')
        $guard.Success | Should Be $true
        $invoke = [scriptblock]::Create('param($sourceRoot,$nativeLocalCache,$failedCommandMatch)' + "`n" + $guard.Value + "`n$" + 'failedTestPath')
        $sourceRoot = Join-Path $TestDrive 'source'
        $nativeLocalCache = Join-Path $TestDrive 'native-local-cache'
        $image = Join-Path $nativeLocalCache 'o/object/ghostty-vt-c-test.exe'
        [void][IO.Directory]::CreateDirectory((Split-Path -Parent $image))
        [IO.File]::WriteAllText($image, 'diagnostic path fixture')
        foreach ($path in @($image, [IO.Path]::GetRelativePath($sourceRoot, $image))) {
            $match = [regex]::Match($path, '^(.+)$')
            (& $invoke $sourceRoot $nativeLocalCache $match) | Should Be $image
        }
        $sibling = Join-Path $TestDrive 'native-local-cache-other/test.exe'
        [void][IO.Directory]::CreateDirectory((Split-Path -Parent $sibling))
        [IO.File]::WriteAllText($sibling, 'outside fixture')
        foreach ($path in @($sibling, (Join-Path $nativeLocalCache 'missing.exe'), '../native-local-cache-other/test.exe')) {
            $match = [regex]::Match($path, '^(.+)$')
            { & $invoke $sourceRoot $nativeLocalCache $match } | Should Throw 'escaped the candidate cache'
        }
    }

    It 'stages the native identity byte-for-byte beside the report and rejects stale or mismatched evidence' {
        $workflow = Get-Content -LiteralPath (Join-Path $repoRoot '.github/workflows/ghostty-runtime-native-qualification.yml') -Raw
        $staging = [regex]::Match($workflow, '(?ms)^ {10}\$identityEvidencePath =.*?(?=^ {10}Write-RSAtomicCanonicalJsonFile)')
        $staging.Success | Should Be $true
        $invoke = [scriptblock]::Create('param($identityPath,$report)' + "`n" + $staging.Value)
        $identity = Join-Path $TestDrive 'workspace/terminal-runtime-identity.json'
        [void][IO.Directory]::CreateDirectory((Split-Path -Parent $identity))
        [IO.File]::WriteAllText($identity, '{"runtime":"fixture"}', [Text.UTF8Encoding]::new($false))
        $report = [pscustomobject]@{ runtime = [pscustomobject]@{ identitySha256 = (Get-FileHash -LiteralPath $identity -Algorithm SHA256).Hash.ToLowerInvariant() } }
        $previousScratch = $env:RS_NATIVE_TEST_SCRATCH
        try {
            $env:RS_NATIVE_TEST_SCRATCH = Join-Path $TestDrive 'native-evidence'
            [void][IO.Directory]::CreateDirectory($env:RS_NATIVE_TEST_SCRATCH)
            & $invoke $identity $report
            $copy = Join-Path $env:RS_NATIVE_TEST_SCRATCH 'terminal-runtime-identity.json'
            [Convert]::ToBase64String([IO.File]::ReadAllBytes($copy)) | Should Be ([Convert]::ToBase64String([IO.File]::ReadAllBytes($identity)))
            { & $invoke $identity $report } | Should Throw 'identity evidence already exists'
            [Convert]::ToBase64String([IO.File]::ReadAllBytes($copy)) | Should Be ([Convert]::ToBase64String([IO.File]::ReadAllBytes($identity)))
            $env:RS_NATIVE_TEST_SCRATCH = Join-Path $TestDrive 'mismatched-evidence'
            [void][IO.Directory]::CreateDirectory($env:RS_NATIVE_TEST_SCRATCH)
            $report.runtime.identitySha256 = '0' * 64
            { & $invoke $identity $report } | Should Throw 'does not match the tested runtime'
        }
        finally { $env:RS_NATIVE_TEST_SCRATCH = $previousScratch }
    }

    It 'keeps build-only diagnostics outside native qualification and product continuation' {
        $workflow = Get-Content (Join-Path $repoRoot '.github/workflows/ghostty-runtime-native-qualification.yml') -Raw
        $blocks = @([regex]::Split($workflow, '(?m)(?=^      - name: )'))
        foreach ($name in @('Run upstream native ARM64 lib-vt targets', 'Publish canonical native qualification evidence', 'Upload native qualification evidence')) {
            $block = @($blocks | Where-Object { $_.StartsWith("      - name: $name`n") -or $_.StartsWith("      - name: $name`r`n") })
            $block.Count | Should Be 1
            $block[0] | Should Match '(?m)^        if: \$\{\{ !inputs\.diagnostics_only \}\}\s*$'
        }
        $workflow | Should Match 'name: \$\{\{ inputs\.diagnostics_only && ''ARM64 build diagnostics \(not qualification\)'''
        $workflow | Should Match '(?m)^    if: \$\{\{ inputs\.qualify_product && !inputs\.diagnostics_only \}\}'
        $diagnostic = @($blocks | Where-Object { $_.StartsWith('      - name: Collect bounded native build diagnostics only') })[0]
        $diagnostic | Should Match 'Invoke-RSStreamingProcess'
        $diagnostic | Should Match '-TimeoutMilliseconds 1200000'
        $diagnostic | Should Match 'qualification = \$false'
        $upload = @($blocks | Where-Object { $_.StartsWith('      - name: Upload build-only diagnostic evidence') })[0]
        $upload | Should Match 'always\(\) && inputs\.diagnostics_only && env\.RS_NATIVE_TEST_SCRATCH != '''''
        $upload | Should Match 'ghostty-native-build-diagnostics-'
        $workflow.IndexOf('- name: Select owned native test-data root') |
            Should BeLessThan $workflow.IndexOf('- name: Build exact candidate natively')
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
                if ($reference.StartsWith('RedSalamanders/.github/')) {
                    $pin = (Get-Content (Join-Path $repoRoot '.ci/project.json') -Raw | ConvertFrom-Json).engine.ref
                    $reference | Should Match ([regex]::Escape('@'+$pin)+'$')
                    continue
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
