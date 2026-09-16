$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $PSScriptRoot 'TestSupport.psm1') -Force
Import-Module (Join-Path $repoRoot 'Tools\Modules\Build\BuildEvidence.psm1') -Force

function Test-RSActionThrows {
    param([Parameter(Mandatory = $true)][scriptblock]$Action)
    $threw = $false
    try { & $Action | Out-Null } catch { $threw = $true }
    return $threw
}

function New-RSBuildArtifactManifestFixture {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$Output,
        [Parameter(Mandatory = $true)][hashtable]$ProjectTarget,
        [hashtable]$TestsEnabled = @{},
        [hashtable]$VcpkgTlog = @{}
    )

    $runtimeManifest = Join-Path $Root 'RuntimeDependencies.props'
    if (-not (Test-Path -LiteralPath $runtimeManifest -PathType Leaf)) {
        Copy-Item -LiteralPath (Join-Path $repoRoot 'RuntimeDependencies.props') -Destination $runtimeManifest
    }
    $manifestRoot = Join-Path $Root ('.build\manifest-fixture-' + [Guid]::NewGuid().ToString('N'))
    [void](New-Item -ItemType Directory -Path $manifestRoot -Force)
    foreach ($project in @($ProjectTarget.Keys | Sort-Object)) {
        $enabled = if ($TestsEnabled.ContainsKey($project)) { [bool]$TestsEnabled[$project] } else { $true }
        $tlogPath = Join-Path $manifestRoot ($project + '.write.1u.tlog')
        if ($VcpkgTlog.ContainsKey($project)) {
            [IO.File]::WriteAllLines($tlogPath, @($VcpkgTlog[$project]), [Text.Encoding]::Unicode)
        } else {
            $tlogPath = Join-Path $manifestRoot ($project + '.missing.write.1u.tlog')
        }
        $lines = @(
            "project=$project"
            "target=$(Join-Path $Output $ProjectTarget[$project])"
            "tests_enabled=$($enabled.ToString().ToLowerInvariant())"
            "vcpkg_tlog=$tlogPath"
        )
        [IO.File]::WriteAllLines((Join-Path $manifestRoot ($project + '.manifest')), $lines, [Text.UTF8Encoding]::new($false))
    }
    return $manifestRoot
}

function New-RSToolchainIdentityFixture {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$Output
    )

    $inputRoot = Join-Path $Root ('.build\toolchain-input-' + [Guid]::NewGuid().ToString('N'))
    [void](New-Item -ItemType Directory -Path $inputRoot -Force)
    $paths = [ordered]@{
        MSBuild = Join-Path $inputRoot 'MSBuild.exe'
        Compiler = Join-Path $inputRoot 'cl.exe'
        Linker = Join-Path $inputRoot 'link.exe'
        ResourceCompiler = Join-Path $inputRoot 'rc.exe'
        WindowsHeader = Join-Path $inputRoot 'Windows.h'
        UmImportLibrary = Join-Path $inputRoot 'kernel32.lib'
        UcrtImportLibrary = Join-Path $inputRoot 'ucrt.lib'
    }
    $index = 1
    foreach ($path in $paths.Values) {
        [IO.File]::WriteAllBytes($path, [byte[]]($index, ($index + 16), ($index + 32)))
        $index++
    }
    $record = New-RSBuildToolchainIdentityRecord -Platform x64 `
        -VcToolsVersion '14.51.36231' -WindowsSdkVersion '10.0.26100.0' `
        -MSBuildPath $paths.MSBuild -CompilerPath $paths.Compiler -LinkerPath $paths.Linker `
        -ResourceCompilerPath $paths.ResourceCompiler -WindowsHeaderPath $paths.WindowsHeader `
        -UmImportLibraryPath $paths.UmImportLibrary -UcrtImportLibraryPath $paths.UcrtImportLibrary
    $schema = Join-Path $repoRoot 'Specs\Build\ToolchainIdentity.schema.json'
    $path = Publish-RSBuildToolchainIdentity -BuildOutputDir $Output -IdentityRecord $record -SchemaPath $schema
    return [pscustomobject]@{
        Record = $record
        Digest = Get-RSCanonicalJsonDigest -InputObject $record
        Path = $path
        Inputs = $paths
        Schema = $schema
    }
}

Describe 'Ghostty build evidence identity closure' {
    It 'binds every active product lock builder schema module overlay notice and staging input' {
        $expected = [string[]]@(
            'External\TerminalEngine\GhosttyRuntimeLock.v1.json',
            'Specs\Terminal\GhosttyRuntimeLock.schema.json',
            'Tools\Build-GhosttyTerminalRuntime.ps1',
            'Tools\Modules\Build\SanitizedEnvironment.psm1',
            'Tools\Modules\Testing\ValidationFingerprint.psm1',
            'Tools\TerminalEngine\GhosttyRuntimeLifecycle.psm1',
            'Tools\TerminalEngine\Patches\Ghostty-WindowsPrivateRuntime.patch',
            'Tools\TerminalEngine\TerminalEngineArchive.psm1',
            'Tools\TerminalEngine\TerminalEnginePeIdentity.psm1',
            'Tools\TerminalEngine\TerminalEngineTextIdentity.psm1',
            'Tools\TerminalEngine\TerminalRuntimeExportIdentity.psm1',
            'Tools\TerminalEngine\TerminalRuntimeImportPolicy.psm1',
            'External\TerminalEngine\Notices\uucode-bjoern-hoehrmann.txt',
            'External\TerminalEngine\Notices\unicode-license-3.0-2025.txt',
            'RuntimeDependencies.props')
        $actual = [string[]]@(Get-RSGhosttyRuntimeIdentityInputPaths)
        ($actual -join "`0") | Should Be ($expected -join "`0")
        $actual | Should Not Contain 'TerminalEngine.lock.json'
        foreach ($relativePath in $actual) {
            Test-Path -LiteralPath (Join-Path $repoRoot $relativePath) -PathType Leaf | Should Be $true
        }
    }

    It 'changes the Ghostty identity when any direct input byte changes' {
        $fixtureRoot = Join-Path $TestDrive 'ghostty-identity-closure'
        $paths = [string[]]@(Get-RSGhosttyRuntimeIdentityInputPaths)
        foreach ($relativePath in $paths) {
            $source = Join-Path $repoRoot $relativePath
            $destination = Join-Path $fixtureRoot $relativePath
            [void](New-Item -ItemType Directory -Path (Split-Path -Parent $destination) -Force)
            Copy-Item -LiteralPath $source -Destination $destination
        }
        $baseline = Get-RSGhosttyRuntimeIdentityDigest -RepoRoot $fixtureRoot
        foreach ($relativePath in $paths) {
            $path = Join-Path $fixtureRoot $relativePath
            $original = [IO.File]::ReadAllBytes($path)
            try {
                [IO.File]::WriteAllBytes($path, [byte[]]($original + [byte]0xA5))
                (Get-RSGhosttyRuntimeIdentityDigest -RepoRoot $fixtureRoot) | Should Not Be $baseline
            }
            finally {
                [IO.File]::WriteAllBytes($path, $original)
            }
        }
    }
}

Describe 'Operation Startrail build output formatting' {
    It 'does not unload the runner fingerprint command surface when build evidence is re-imported' {
        $fingerprintModule = Join-Path $repoRoot 'Tools\Modules\Testing\ValidationFingerprint.psm1'
        $buildEvidenceModule = Join-Path $repoRoot 'Tools\Modules\Build\BuildEvidence.psm1'
        Import-Module $fingerprintModule -Force
        Import-Module $buildEvidenceModule -Force
        (Get-Command Write-RSAtomicCanonicalJsonFile -ErrorAction SilentlyContinue) | Should Not BeNullOrEmpty
    }

    It 'formats elapsed time without hour rollover' {
        Format-RSBuildDuration ([timespan]::FromSeconds(3599)) | Should Be '59:59'
        Format-RSBuildDuration ([timespan]::FromSeconds(3600)) | Should Be '1:00:00'
        Format-RSBuildDuration ([timespan]::FromSeconds(7509)) | Should Be '2:05:09'
    }

    It 'formats fixed two-decimal invariant MiB under another culture' {
        $prior = [Globalization.CultureInfo]::CurrentCulture
        try {
            [Globalization.CultureInfo]::CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('fr-FR')
            Format-RSBuildSizeMiB 1572864 | Should Be '1.50'
        } finally {
            [Globalization.CultureInfo]::CurrentCulture = $prior
        }
    }

    It 'escapes PowerShell interpolation and distinguishes runnable from data output' {
        $artifact = [pscustomobject]@{
            absolute_path_diagnostic = 'C:\path & more\dollar$ back`tick\app.exe'
            size_bytes = 1048576
            role = 'launchable'
        }
        $line = ConvertTo-RSBuildOutputLine -Artifact $artifact
        $line | Should Be '& "C:\path & more\dollar`$ back``tick\app.exe" # 1.00 MiB'
        $tokens = $null
        $errors = $null
        [void][Management.Automation.Language.Parser]::ParseInput($line, [ref]$tokens, [ref]$errors)
        @($errors).Count | Should Be 0
        $artifact.role = 'package'
        (ConvertTo-RSBuildOutputLine -Artifact $artifact) | Should Match '^Output: '
        $artifact.absolute_path_diagnostic = '\\server\share\path with spaces\app.exe'
        $artifact.role = 'launchable'
        (ConvertTo-RSBuildOutputLine -Artifact $artifact) | Should Be '& "\\server\share\path with spaces\app.exe" # 1.00 MiB'
    }

    It 'runs a copied harmless executable from a foreign current directory' {
        $fixtureDir = Join-Path $TestDrive 'space & dollar$ apostrophe'' tick`'
        [void](New-Item -ItemType Directory -Path $fixtureDir)
        $exe = Join-Path $fixtureDir 'whoami.exe'
        Copy-Item -LiteralPath (Join-Path $env:SystemRoot 'System32\whoami.exe') -Destination $exe
        $artifact = [pscustomobject]@{
            absolute_path_diagnostic = $exe
            size_bytes = (Get-Item -LiteralPath $exe).Length
            role = 'launchable'
        }
        $line = ConvertTo-RSBuildOutputLine -Artifact $artifact
        $foreign = Join-Path $TestDrive 'foreign'
        [void](New-Item -ItemType Directory -Path $foreign)
        Push-Location $foreign
        try {
            $output = Invoke-Expression $line
            $LASTEXITCODE | Should Be 0
            [string]::IsNullOrWhiteSpace(($output -join '')) | Should Be $false
        } finally {
            Pop-Location
        }
    }
}

Describe 'Operation Startrail build receipt records' {
    It 'normalizes current vcpkg source tlogs and legacy destination tlogs to deployed outputs' {
        $root = Join-Path $TestDrive 'vcpkg-tlog-repo'
        $output = Join-Path $root '.build\x64\Release'
        $pluginOutput = Join-Path $output 'Plugins'
        $vcpkgBin = Join-Path $root '.build\vcpkg_installed\x64\x64-windows\bin'
        [void](New-Item -ItemType Directory -Path $pluginOutput, $vcpkgBin -Force)
        $target = Join-Path $pluginOutput 'Plugin.dll'
        $source = Join-Path $vcpkgBin 'blake3.dll'
        $deployed = Join-Path $pluginOutput 'blake3.dll'
        [IO.File]::WriteAllBytes($target, [byte[]](1, 2, 3))
        [IO.File]::WriteAllBytes($source, [byte[]](4, 5, 6))
        [IO.File]::WriteAllBytes($deployed, [byte[]](4, 5, 6))

        $sourceManifest = New-RSBuildArtifactManifestFixture -Root $root -Output $output `
            -ProjectTarget @{ Plugin = 'Plugins\Plugin.dll' } -TestsEnabled @{ Plugin = $false } `
            -VcpkgTlog @{ Plugin = @($source) }
        $sourceArtifacts = @(Get-RSBuildArtifactRecords -RepoRoot $root -BuildOutputDir $output `
            -ArtifactManifestDir $sourceManifest -Configuration Release -Platform x64 -ProjectName Plugin)
        $plugin = @($sourceArtifacts | Where-Object relative_path -eq '.build/x64/Release/Plugins/Plugin.dll')
        $plugin.Count | Should Be 1
        ($plugin[0].runtime_closure -contains '.build/x64/Release/Plugins/blake3.dll') | Should Be $true

        $destinationManifest = New-RSBuildArtifactManifestFixture -Root $root -Output $output `
            -ProjectTarget @{ Plugin = 'Plugins\Plugin.dll' } -TestsEnabled @{ Plugin = $false } `
            -VcpkgTlog @{ Plugin = @($deployed) }
        $destinationArtifacts = @(Get-RSBuildArtifactRecords -RepoRoot $root -BuildOutputDir $output `
            -ArtifactManifestDir $destinationManifest -Configuration Release -Platform x64 -ProjectName Plugin)
        $plugin = @($destinationArtifacts | Where-Object relative_path -eq '.build/x64/Release/Plugins/Plugin.dll')
        ($plugin[0].runtime_closure -contains '.build/x64/Release/Plugins/blake3.dll') | Should Be $true
    }

    It 'rejects foreign vcpkg tlog sources and missing deployed outputs' {
        $root = Join-Path $TestDrive 'vcpkg-tlog-rejection-repo'
        $output = Join-Path $root '.build\x64\Release'
        $vcpkgBin = Join-Path $root '.build\vcpkg_installed\x64\x64-windows\bin'
        $foreignBin = Join-Path $root 'foreign'
        [void](New-Item -ItemType Directory -Path $output, $vcpkgBin, $foreignBin -Force)
        [IO.File]::WriteAllBytes((Join-Path $output 'App.exe'), [byte[]](1))
        $source = Join-Path $vcpkgBin 'missing.dll'
        $foreign = Join-Path $foreignBin 'foreign.dll'
        [IO.File]::WriteAllBytes($source, [byte[]](2))
        [IO.File]::WriteAllBytes($foreign, [byte[]](3))

        $missingManifest = New-RSBuildArtifactManifestFixture -Root $root -Output $output `
            -ProjectTarget @{ App = 'App.exe' } -VcpkgTlog @{ App = @($source) }
        (Test-RSActionThrows { Get-RSBuildArtifactRecords -RepoRoot $root -BuildOutputDir $output `
                -ArtifactManifestDir $missingManifest -Configuration Release -Platform x64 }) | Should Be $true

        [IO.File]::WriteAllBytes((Join-Path $output 'foreign.dll'), [byte[]](3))
        $foreignManifest = New-RSBuildArtifactManifestFixture -Root $root -Output $output `
            -ProjectTarget @{ App = 'App.exe' } -VcpkgTlog @{ App = @($foreign) }
        (Test-RSActionThrows { Get-RSBuildArtifactRecords -RepoRoot $root -BuildOutputDir $output `
                -ArtifactManifestDir $foreignManifest -Configuration Release -Platform x64 }) | Should Be $true
    }

    It 'selects full and targeted artifacts deterministically without fabricating outputs' {
        $root = Join-Path $TestDrive 'repo'
        $output = Join-Path $root '.build\x64\Debug'
        [void](New-Item -ItemType Directory -Path (Join-Path $output 'Plugins') -Force)
        [IO.File]::WriteAllBytes((Join-Path $output 'App.exe'), [byte[]](1, 2, 3))
        [IO.File]::WriteAllBytes((Join-Path $output 'Other.exe'), [byte[]](4, 5))
        [IO.File]::WriteAllBytes((Join-Path $output 'Plugins\Plugin.dll'), [byte[]](6))

        $manifest = New-RSBuildArtifactManifestFixture -Root $root -Output $output -ProjectTarget @{
            App = 'App.exe'; Other = 'Other.exe'; Plugin = 'Plugins\Plugin.dll'
        }
        $full = @(Get-RSBuildArtifactRecords -RepoRoot $root -BuildOutputDir $output `
            -ArtifactManifestDir $manifest -Configuration Debug -Platform x64)
        (@($full.relative_path) -join '|') | Should Be '.build/x64/Debug/App.exe|.build/x64/Debug/Other.exe|.build/x64/Debug/Plugins/Plugin.dll'
        $target = @(Get-RSBuildArtifactRecords -RepoRoot $root -BuildOutputDir $output `
            -ArtifactManifestDir $manifest -Configuration Debug -Platform x64 -ProjectName Plugin)
        $target.Count | Should Be 3
        ($target.relative_path -contains '.build/x64/Debug/Plugins/Plugin.dll') | Should Be $true
        (Test-RSActionThrows { Get-RSBuildArtifactRecords -RepoRoot $root -BuildOutputDir $output `
                -ArtifactManifestDir $manifest -Configuration Debug -Platform x64 -ProjectName Missing }) | Should Be $true
    }

    It 'rejects undeclared binaries and measures the effective project test surface' {
        $root = Join-Path $TestDrive 'governed-output-repo'
        $output = Join-Path $root '.build\x64\Debug'
        [void](New-Item -ItemType Directory -Path (Join-Path $output 'Plugins') -Force)
        [IO.File]::WriteAllBytes((Join-Path $output 'App.exe'), [byte[]](1))
        [IO.File]::WriteAllBytes((Join-Path $output 'Plugins\Plugin.dll'), [byte[]](2))
        [IO.File]::WriteAllBytes((Join-Path $output 'RemovedProject.exe'), [byte[]](3))
        $manifest = New-RSBuildArtifactManifestFixture -Root $root -Output $output `
            -ProjectTarget @{ App = 'App.exe'; Plugin = 'Plugins\Plugin.dll' } `
            -TestsEnabled @{ App = $true; Plugin = $false }

        (Test-RSBuildArtifactManifestTestsEnabled -ArtifactManifestDir $manifest) | Should Be $false
        (Test-RSActionThrows { Get-RSBuildArtifactRecords -RepoRoot $root -BuildOutputDir $output `
                -ArtifactManifestDir $manifest -Configuration Debug -Platform x64 }) | Should Be $true
        Remove-Item -LiteralPath (Join-Path $output 'RemovedProject.exe')
        @(Get-RSBuildArtifactRecords -RepoRoot $root -BuildOutputDir $output `
                -ArtifactManifestDir $manifest -Configuration Debug -Platform x64).Count | Should Be 2

        $declaredRuntime = Join-Path $output 'declared-runtime.dll'
        [IO.File]::WriteAllBytes($declaredRuntime, [byte[]](4))
        Add-Content -LiteralPath (Join-Path $manifest 'App.manifest') -Value "runtime=$declaredRuntime" -Encoding utf8NoBOM
        $records = @(Get-RSBuildArtifactRecords -RepoRoot $root -BuildOutputDir $output `
                -ArtifactManifestDir $manifest -Configuration Debug -Platform x64)
        $records.Count | Should Be 3
        ($records.relative_path -contains '.build/x64/Debug/declared-runtime.dll') | Should Be $true
    }

    It 'measures ordinary and explicitly test-enabled Debug and Release surfaces' {
        $root = Join-Path $TestDrive 'test-surface-matrix'
        $output = Join-Path $root '.build\x64\Debug'
        [void](New-Item -ItemType Directory -Path $output -Force)
        [IO.File]::WriteAllBytes((Join-Path $output 'App.exe'), [byte[]](1))

        $ordinaryDebug = New-RSBuildArtifactManifestFixture -Root $root -Output $output `
            -ProjectTarget @{ App = 'App.exe' } -TestsEnabled @{ App = $true }
        (Test-RSBuildArtifactManifestTestsEnabled -ArtifactManifestDir $ordinaryDebug) | Should Be $true

        $productionRelease = New-RSBuildArtifactManifestFixture -Root $root -Output $output `
            -ProjectTarget @{ App = 'App.exe' } -TestsEnabled @{ App = $false }
        (Test-RSBuildArtifactManifestTestsEnabled -ArtifactManifestDir $productionRelease) | Should Be $false

        $testEnabledRelease = New-RSBuildArtifactManifestFixture -Root $root -Output $output `
            -ProjectTarget @{ App = 'App.exe' } -TestsEnabled @{ App = $true }
        (Test-RSBuildArtifactManifestTestsEnabled -ArtifactManifestDir $testEnabledRelease) | Should Be $true
    }

    It 'changes canonical toolchain identity for every compiler linker RC and SDK byte seam' {
        $root = Join-Path $TestDrive 'toolchain-identity-repo'
        $output = Join-Path $root '.build\x64\Debug'
        [void](New-Item -ItemType Directory -Path $output -Force)
        $fixture = New-RSToolchainIdentityFixture -Root $root -Output $output
        Test-Path -LiteralPath $fixture.Path -PathType Leaf | Should Be $true

        foreach ($inputName in @($fixture.Inputs.Keys)) {
            $path = $fixture.Inputs[$inputName]
            $original = [IO.File]::ReadAllBytes($path)
            try {
                [IO.File]::WriteAllBytes($path, [byte[]]($original + [byte]0xA5))
                $mutated = New-RSBuildToolchainIdentityRecord -Platform x64 `
                    -VcToolsVersion '14.51.36231' -WindowsSdkVersion '10.0.26100.0' `
                    -MSBuildPath $fixture.Inputs.MSBuild -CompilerPath $fixture.Inputs.Compiler `
                    -LinkerPath $fixture.Inputs.Linker -ResourceCompilerPath $fixture.Inputs.ResourceCompiler `
                    -WindowsHeaderPath $fixture.Inputs.WindowsHeader `
                    -UmImportLibraryPath $fixture.Inputs.UmImportLibrary `
                    -UcrtImportLibraryPath $fixture.Inputs.UcrtImportLibrary
                (Get-RSCanonicalJsonDigest -InputObject $mutated) | Should Not Be $fixture.Digest
            } finally {
                [IO.File]::WriteAllBytes($path, $original)
            }
        }
    }

    It 'changes the source snapshot for dirty staged untracked delete and rename mutations' {
        $fixture = New-RSValidationGitFixture -Root (Join-Path $TestDrive 'source-fixture')
        $mutated = Get-RSBuildSourceSnapshot -RepoRoot $fixture.Root
        & git -C $fixture.Root reset --hard --quiet HEAD
        & git -C $fixture.Root clean -fdq
        $clean = Get-RSBuildSourceSnapshot -RepoRoot $fixture.Root
        $mutated.SnapshotId | Should Not Be $clean.SnapshotId
        $mutated.GitHead | Should Be $clean.GitHead
    }

    It 'publishes immutable content-addressed receipts and rejects tampered artifacts' {
        $root = Join-Path $TestDrive 'receipt-repo'
        $output = Join-Path $root '.build\x64\Debug'
        [void](New-Item -ItemType Directory -Path $output -Force)
        $artifactPath = Join-Path $output 'App.exe'
        $libraryPath = Join-Path $output 'Library.dll'
        [IO.File]::WriteAllBytes($artifactPath, [byte[]](1, 2, 3, 4))
        [IO.File]::WriteAllBytes($libraryPath, [byte[]](5, 6, 7, 8))
        $manifest = New-RSBuildArtifactManifestFixture -Root $root -Output $output `
            -ProjectTarget @{ App = 'App.exe'; Library = 'Library.dll' }
        $artifacts = @(Get-RSBuildArtifactRecords -RepoRoot $root -BuildOutputDir $output `
            -ArtifactManifestDir $manifest -Configuration Debug -Platform x64)
        $digest = 'a' * 64
        $source = [pscustomobject]@{ SnapshotId = 'b' * 64; GitHead = 'c' * 40 }
        $identity = [pscustomobject]@{
            RepoIdentity = 'd' * 64
            ToolchainIdentityDigest = 'e' * 64
            VcpkgIdentityDigest = 'f' * 64
            GhosttyRuntimeIdentity = $digest
            ProjectGraphDigest = '1' * 64
            RuntimeDependencyDigest = '2' * 64
        }
        $started = [datetime]'2026-08-13T12:00:00Z'
        $receipt = New-RSBuildReceipt -SourceSnapshot $source -Identity $identity -Platform x64 -Configuration Debug `
            -Target solution -ProjectSelection @('RedSalamander.sln') -BuildArguments @('/t:Rebuild') `
            -TestsEnabled $true -OfficialRelease $false -StartedUtc $started -EndedUtc $started.AddSeconds(2) `
            -Duration ([timespan]::FromSeconds(2)) -Outputs $artifacts
        $schema = Join-Path $repoRoot 'Specs\Build\BuildReceipt.schema.json'
        $stored = Publish-RSBuildReceipt -BuildOutputDir $output -Receipt $receipt -SchemaPath $schema
        $stored.receipt_id | Should Be $receipt.receipt_id
        Test-Path -LiteralPath (Join-Path $output "receipts\$($receipt.receipt_id).json") | Should Be $true

        $pointerPath = Join-Path $output 'build-receipt.json'
        $pointerBytes = [IO.File]::ReadAllBytes($pointerPath)
        $wrongPointer = ConvertTo-RSCanonicalJson -InputObject $stored | ConvertFrom-Json -DateKind String
        $wrongPointer.outputs[0].absolute_path_diagnostic = 'C:\different\diagnostic.exe'
        Write-RSAtomicCanonicalJsonFile -LiteralPath $pointerPath -InputObject $wrongPointer `
            -SchemaPath $schema -AllowedRoot $output
        (Test-RSActionThrows { Read-RSBuildReceipt -BuildOutputDir $output -SchemaPath $schema }) | Should Be $true
        [IO.File]::WriteAllBytes($pointerPath, $pointerBytes)

        $truncated = New-RSBuildReceipt -SourceSnapshot $source -Identity $identity -Platform x64 -Configuration Debug `
            -Target solution -ProjectSelection @('RedSalamander.sln') -BuildArguments @('/t:Rebuild') `
            -TestsEnabled $true -OfficialRelease $false -StartedUtc $started -EndedUtc $started.AddSeconds(2) `
            -Duration ([timespan]::FromSeconds(2)) -Outputs @($artifacts[0])
        Write-RSAtomicCanonicalJsonFile -LiteralPath $pointerPath -InputObject $truncated `
            -SchemaPath $schema -AllowedRoot $output
        (Test-RSActionThrows { Read-RSBuildReceipt -BuildOutputDir $output -SchemaPath $schema }) | Should Be $true
        [IO.File]::WriteAllBytes($pointerPath, $pointerBytes)

        $duplicateIdOutputs = @($artifacts | ForEach-Object { $_ | ConvertTo-Json -Depth 10 | ConvertFrom-Json })
        $duplicateIdOutputs[1].id = $duplicateIdOutputs[0].id
        $duplicateId = New-RSBuildReceipt -SourceSnapshot $source -Identity $identity -Platform x64 -Configuration Debug `
            -Target solution -ProjectSelection @('RedSalamander.sln') -BuildArguments @('/t:Rebuild') `
            -TestsEnabled $true -OfficialRelease $false -StartedUtc $started -EndedUtc $started.AddSeconds(2) `
            -Duration ([timespan]::FromSeconds(2)) -Outputs $duplicateIdOutputs
        (Test-RSBuildReceiptForArtifactUse -Receipt $duplicateId -SourceSnapshot $source -Platform x64 `
            -Configuration Debug -RepoRoot $root) | Should Be $false

        $brokenClosureOutputs = @($artifacts | ForEach-Object { $_ | ConvertTo-Json -Depth 10 | ConvertFrom-Json })
        $brokenClosureOutputs[0].runtime_closure = @('.build/x64/Debug/Missing.dll')
        $brokenClosure = New-RSBuildReceipt -SourceSnapshot $source -Identity $identity -Platform x64 -Configuration Debug `
            -Target solution -ProjectSelection @('RedSalamander.sln') -BuildArguments @('/t:Rebuild') `
            -TestsEnabled $true -OfficialRelease $false -StartedUtc $started -EndedUtc $started.AddSeconds(2) `
            -Duration ([timespan]::FromSeconds(2)) -Outputs $brokenClosureOutputs
        (Test-RSBuildReceiptForArtifactUse -Receipt $brokenClosure -SourceSnapshot $source -Platform x64 `
            -Configuration Debug -RepoRoot $root) | Should Be $false
        $targetReceipt = New-RSBuildReceipt -SourceSnapshot $source -Identity $identity -Platform x64 -Configuration Debug `
            -Target common -ProjectSelection @('Common\Common.vcxproj') -BuildArguments @('/t:Build') `
            -TestsEnabled $true -OfficialRelease $false -StartedUtc $started -EndedUtc $started.AddSeconds(1) `
            -Duration ([timespan]::FromSeconds(1)) -Outputs $artifacts
        Publish-RSBuildReceipt -BuildOutputDir $output -Receipt $targetReceipt -SchemaPath $schema | Out-Null
        Test-Path -LiteralPath (Join-Path $output "receipts\$($receipt.receipt_id).json") | Should Be $true
        Test-Path -LiteralPath (Join-Path $output "receipts\$($targetReceipt.receipt_id).json") | Should Be $true
        $pointerBeforeFailure = [IO.File]::ReadAllBytes((Join-Path $output 'build-receipt.json'))
        $invalidReceipt = $receipt | ConvertTo-Json -Depth 20 | ConvertFrom-Json
        $invalidReceipt.receipt_id = '9' * 64
        $invalidReceipt.outputs[0].relative_path = '../escape.exe'
        (Test-RSActionThrows { Publish-RSBuildReceipt -BuildOutputDir $output -Receipt $invalidReceipt -SchemaPath $schema }) | Should Be $true
        [Convert]::ToHexString([IO.File]::ReadAllBytes((Join-Path $output 'build-receipt.json'))) | Should Be ([Convert]::ToHexString($pointerBeforeFailure))
        Suspend-RSBuildReceipt -BuildOutputDir $output
        Test-Path -LiteralPath (Join-Path $output 'build-receipt.json') | Should Be $false
        $stored = Publish-RSBuildReceipt -BuildOutputDir $output -Receipt $receipt -SchemaPath $schema
        (Test-RSBuildReceiptCompatible -Receipt $stored -SourceSnapshot $source -Identity $identity `
            -Platform x64 -Configuration Debug -Target solution -RepoRoot $root `
            -ProjectSelection @('RedSalamander.sln') -BuildArguments @('/t:Rebuild')) | Should Be $true
        (Test-RSBuildReceiptForArtifactUse -Receipt $stored -SourceSnapshot $source -Platform x64 `
            -Configuration Debug -RepoRoot $root -RequiredRelativePath @('.build/x64/Debug/App.exe')) | Should Be $true
        [IO.File]::WriteAllBytes($artifactPath, [byte[]](9, 9, 9))
        (Test-RSBuildReceiptCompatible -Receipt $stored -SourceSnapshot $source -Identity $identity `
            -Platform x64 -Configuration Debug -Target solution -RepoRoot $root `
            -ProjectSelection @('RedSalamander.sln') -BuildArguments @('/t:Rebuild')) | Should Be $false
        (Test-RSBuildReceiptForArtifactUse -Receipt $stored -SourceSnapshot $source -Platform x64 `
            -Configuration Debug -RepoRoot $root -RequiredRelativePath @('.build/x64/Debug/App.exe')) | Should Be $false
    }

    It 'accepts only same-workflow portable evidence and rebinds verified artifacts locally' {
        $root = Join-Path $TestDrive 'portable-repo'
        [void](New-Item -ItemType Directory -Path $root)
        [IO.File]::WriteAllText((Join-Path $root '.gitignore'), ".build/`n", [Text.UTF8Encoding]::new($false))
        [IO.File]::WriteAllText((Join-Path $root 'source.txt'), "source`n", [Text.UTF8Encoding]::new($false))
        Copy-Item -LiteralPath (Join-Path $repoRoot 'RuntimeDependencies.props') -Destination (Join-Path $root 'RuntimeDependencies.props')
        & git -C $root init --quiet
        & git -C $root config user.email 'startrail-fixture@invalid.example'
        & git -C $root config user.name 'Startrail Fixture'
        & git -C $root add -- .
        & git -C $root commit --quiet -m 'portable fixture'
        $source = Get-RSBuildSourceSnapshot -RepoRoot $root
        $output = Join-Path $root '.build\x64\Debug'
        [void](New-Item -ItemType Directory -Path $output -Force)
        $artifactPath = Join-Path $output 'App.exe'
        [IO.File]::WriteAllBytes($artifactPath, [byte[]](1, 3, 5, 7))
        $manifest = New-RSBuildArtifactManifestFixture -Root $root -Output $output -ProjectTarget @{ App = 'App.exe' }
        $artifacts = @(Get-RSBuildArtifactRecords -RepoRoot $root -BuildOutputDir $output `
            -ArtifactManifestDir $manifest -Configuration Debug -Platform x64)
        $toolchain = New-RSToolchainIdentityFixture -Root $root -Output $output
        $artifacts += New-RSBuildArtifactRecord -RepoRoot $root -LiteralPath $toolchain.Path
        $artifacts = @($artifacts | Sort-Object relative_path)
        $identity = [pscustomobject]@{
            RepoIdentity = '1' * 64; ToolchainIdentityDigest = $toolchain.Digest
            VcpkgIdentityDigest = '3' * 64; GhosttyRuntimeIdentity = '4' * 64
            ProjectGraphDigest = '5' * 64; RuntimeDependencyDigest = '6' * 64
        }
        $created = [datetime]'2026-08-13T12:00:00Z'
        $receipt = New-RSBuildReceipt -SourceSnapshot $source -Identity $identity -Platform x64 -Configuration Debug `
            -Target solution -ProjectSelection @('RedSalamander.sln') -BuildArguments @('ci=true') `
            -TestsEnabled $true -OfficialRelease $false -StartedUtc $created -EndedUtc $created `
            -Duration ([timespan]::Zero) -Outputs $artifacts
        $originalArtifactBytes = [IO.File]::ReadAllBytes($artifactPath)
        [IO.File]::WriteAllBytes($artifactPath, [byte[]](9, 9, 9, 9))
        (Test-RSActionThrows {
                New-RSPortableBuildAttestation -Receipt $receipt -RepoRoot $root -BuildOutputDir $output `
                    -ToolchainSchemaPath $toolchain.Schema -ProducerRepository 'owner/repo' `
                    -ProducerWorkflow 'CI' -ProducerRunId '1234' -ProducerRunAttempt 2 -ProducerJob 'build' `
                    -CreatedUtc $created
            }) | Should Be $true
        [IO.File]::WriteAllBytes($artifactPath, $originalArtifactBytes)
        Remove-Item -LiteralPath $artifactPath
        (Test-RSActionThrows {
                New-RSPortableBuildAttestation -Receipt $receipt -RepoRoot $root -BuildOutputDir $output `
                    -ToolchainSchemaPath $toolchain.Schema -ProducerRepository 'owner/repo' `
                    -ProducerWorkflow 'CI' -ProducerRunId '1234' -ProducerRunAttempt 2 -ProducerJob 'build' `
                    -CreatedUtc $created
            }) | Should Be $true
        [IO.File]::WriteAllBytes($artifactPath, $originalArtifactBytes)
        $portable = New-RSPortableBuildAttestation -Receipt $receipt -RepoRoot $root -BuildOutputDir $output `
            -ToolchainSchemaPath $toolchain.Schema `
            -ProducerRepository 'owner/repo' -ProducerWorkflow 'CI' -ProducerRunId '1234' `
            -ProducerRunAttempt 2 -ProducerJob 'build' -CreatedUtc $created
        $portableSchema = Join-Path $repoRoot 'Specs\Build\PortableBuildAttestation.schema.json'
        $receiptSchema = Join-Path $repoRoot 'Specs\Build\BuildReceipt.schema.json'
        $unreviewedProducerFile = Join-Path $output 'unreviewed.txt'
        [IO.File]::WriteAllText($unreviewedProducerFile, 'not in receipt', [Text.UTF8Encoding]::new($false))
        $payload = Publish-RSPortableBuildPayload -RepoRoot $root -BuildOutputDir $output `
            -Receipt $receipt -Attestation $portable -PortableSchemaPath $portableSchema
        Test-Path -LiteralPath (Join-Path $payload 'App.exe') -PathType Leaf | Should Be $true
        Test-Path -LiteralPath (Join-Path $payload 'toolchain-identity.json') -PathType Leaf | Should Be $true
        Test-Path -LiteralPath (Join-Path $payload 'portable-build-attestation.json') -PathType Leaf | Should Be $true
        Test-Path -LiteralPath (Join-Path $payload 'unreviewed.txt') | Should Be $false
        @(Get-ChildItem -LiteralPath $payload -File -Recurse).Count | Should Be 3
        Remove-Item -LiteralPath $unreviewedProducerFile
        $invalidPortable = $portable | ConvertTo-Json -Depth 20 | ConvertFrom-Json
        $invalidPortable.artifacts[0].relative_path = 'C:/absolute-authority.exe'
        (Test-RSActionThrows {
                Publish-RSPortableBuildAttestation -BuildOutputDir $output -Attestation $invalidPortable -SchemaPath $portableSchema
            }) | Should Be $true
        Publish-RSPortableBuildAttestation -BuildOutputDir $output -Attestation $portable -SchemaPath $portableSchema | Out-Null

        $import = {
            param($job, $workflow, $sourceCommit, $attestationId, $toolchainDigest)
            Import-RSPortableBuildAttestation -RepoRoot $root -BuildOutputDir $output `
                -PortableSchemaPath $portableSchema -BuildReceiptSchemaPath $receiptSchema `
                -ToolchainSchemaPath $toolchain.Schema `
                -ExpectedRepository 'owner/repo' -ExpectedWorkflow $workflow -ExpectedRunId '1234' `
                -ExpectedRunAttempt 2 -ExpectedProducerJob $job -ExpectedSourceCommit $sourceCommit `
                -ExpectedAttestationId $attestationId -ExpectedToolchainIdentityDigest $toolchainDigest `
                -ExpectedPlatform x64 -ExpectedConfiguration Debug
        }
        $goodArgs = @('build', 'CI', $source.GitHead, $portable.attestation_id, $portable.toolchain_identity_digest)
        (Test-RSActionThrows { & $import 'other-build' 'CI' $source.GitHead $portable.attestation_id $portable.toolchain_identity_digest }) | Should Be $true
        (Test-RSActionThrows { & $import 'build' 'Other workflow' $source.GitHead $portable.attestation_id $portable.toolchain_identity_digest }) | Should Be $true
        (Test-RSActionThrows { & $import 'build' 'CI' ('a' * 40) $portable.attestation_id $portable.toolchain_identity_digest }) | Should Be $true
        (Test-RSActionThrows { & $import 'build' 'CI' $source.GitHead ('f' * 64) $portable.toolchain_identity_digest }) | Should Be $true
        (Test-RSActionThrows { & $import 'build' 'CI' $source.GitHead $portable.attestation_id ('e' * 64) }) | Should Be $true
        $rebound = & $import @goodArgs
        $rebound.source_snapshot_id | Should Be $source.SnapshotId
        (@($rebound.outputs.absolute_path_diagnostic) -contains $artifactPath) | Should Be $true
        $unmanifested = Join-Path $output 'Themes\unmanifested.theme.json5'
        [void](New-Item -ItemType Directory -Path (Split-Path -Parent $unmanifested) -Force)
        [IO.File]::WriteAllText($unmanifested, '{}', [Text.UTF8Encoding]::new($false))
        (Test-RSActionThrows { & $import @goodArgs }) | Should Be $true
        Remove-Item -LiteralPath $unmanifested
        [IO.File]::WriteAllBytes($artifactPath, [byte[]](9, 9, 9, 9))
        (Test-RSActionThrows { & $import @goodArgs }) | Should Be $true
    }

    It 'stops the workflow after a failed build before another native command can replace its exit status' {
        $workflow = Get-Content -LiteralPath (Join-Path $repoRoot '.github\workflows\build-reusable.yml') -Raw
        $start = $workflow.IndexOf('& $buildScript @buildArgs', [StringComparison]::Ordinal)
        $end = $workflow.IndexOf('$pin = Get-Content Dependencies/DxUi.lock.json', $start, [StringComparison]::Ordinal)
        $invocation = [scriptblock]::Create($workflow.Substring($start, $end - $start))
        $buildScript = Join-Path $TestDrive 'failed-build.ps1'
        [IO.File]::WriteAllText($buildScript, 'exit 23')
        $buildArgs = @{}
        $failure = ''
        try { & $invocation } catch { $failure = $_.Exception.Message }
        $failure | Should Be 'Solution build failed with exit code 23.'
        [IO.File]::WriteAllText($buildScript, 'exit 0')
        (Test-RSActionThrows { & $invocation }) | Should Be $false
    }

    It 'keeps same-job CI receipt verification and cross-job nightly attestation ordered' {
        $reusable = Get-Content -LiteralPath (Join-Path $repoRoot '.github\workflows\build-reusable.yml') -Raw
        $reusable | Should Match 'portable_attestation_id:\s*\r?\n\s+description:'
        $reusable | Should Match 'id: portable_attestation'
        $reusable | Should Match 'New-RSPortableBuildAttestation'
        $reusable | Should Match 'Publish-RSPortableBuildPayload'
        $reusable | Should Match 'path:\s*\$\{\{ steps\.portable_attestation\.outputs\.payload_path \}\}/'
        $reusable | Should Not Match '(?s)Create portable same-workflow build attestation.+?Get-RSBuildArtifactRecords.+?Upload build output'
        $reusable | Should Match 'attestation_id=.*GITHUB_OUTPUT'
        $reusable.IndexOf('Create portable same-workflow build attestation') | Should BeLessThan $reusable.IndexOf('Upload build output')

        $reusable | Should Match 'Run-AllTests.ps1 -Suite Full -ValidationMode Fresh -SkipBuild'
        $reusable.IndexOf('Build solution') | Should BeLessThan $reusable.IndexOf('Run native Fresh Full suite')

        foreach ($relative in @('.github\workflows\nightly-flake.yml')) {
            $workflow = Get-Content -LiteralPath (Join-Path $repoRoot $relative) -Raw
            $workflow | Should Match 'Import-RSPortableBuildAttestation'
            $workflow | Should Match 'ToolchainIdentity\.schema\.json'
            $workflow | Should Match 'needs\.selftest-build\.outputs\.portable_attestation_id'
            $workflow | Should Match 'needs\.selftest-build\.outputs\.portable_toolchain_identity_digest'
            $workflow.IndexOf('Download build output') | Should BeLessThan $workflow.IndexOf('Verify and rebind portable build attestation')
            $workflow.IndexOf('Verify and rebind portable build attestation') | Should BeLessThan $workflow.IndexOf('Run-AllTests.ps1')
        }
    }

    It 'renders exactly one deterministic final success block from artifact records' {
        $lines = @(ConvertTo-RSBuildSummaryLines -Configuration 'ASan Debug' -Platform x64 `
            -Duration ([timespan]::FromSeconds(169)) -Artifacts @(
                [pscustomobject]@{ relative_path = 'b.exe'; absolute_path_diagnostic = 'C:\b.exe'; size_bytes = 2MB; role = 'launchable' },
                [pscustomobject]@{ relative_path = 'a.dll'; absolute_path_diagnostic = 'C:\a.dll'; size_bytes = 1MB; role = 'library' }))
        @($lines | Where-Object { $_ -eq 'Build completed successfully!' }).Count | Should Be 1
        $lines[2] | Should Be 'Configuration: ASan Debug | Platform: x64'
        $lines[3] | Should Be 'Build time: 02:49'
        $lines[-2] | Should Match '^Output: "C:\\a\.dll"'
        $lines[-1] | Should Match '^& "C:\\b\.exe"'
    }
}
