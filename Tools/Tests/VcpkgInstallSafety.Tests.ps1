Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$testRunPlanModule = Join-Path $repoRoot 'Tools\Modules\Testing\TestRunPlan.psm1'
$helperModule = Join-Path $repoRoot 'Tools\Modules\Build\VcpkgInstallSafety.psm1'
Import-Module $testRunPlanModule -Force -ErrorAction Stop
Import-Module $helperModule -Force -ErrorAction Stop
Import-Module (Join-Path $repoRoot 'Tools\Modules\Build\DxUiPaths.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot 'TestSupport.psm1') -Force

function New-RSTemporaryVcpkgInstallSafetyRoot {
    return (New-RSTestSandboxScratchDirectory `
            -RepoRoot $repoRoot `
            -Harness 'tools-pester' `
            -Case "vcpkg-install-safety-$([System.Guid]::NewGuid().ToString('N'))")
}

function Assert-ThrowsTerminatingError {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock
    )

    $threw = $false
    try {
        $null = & $ScriptBlock
    } catch {
        $threw = $true
    }

    $threw | Should Be $true
}

Describe 'Vcpkg install safety helper' {
    It 'contains vcpkg execution and acquires selected platform dependency locks in stable order' {
        $scriptPath = Join-Path $repoRoot 'vcpkg-install.ps1'
        $tokens = $null
        $parseErrors = $null
        $scriptAst = [System.Management.Automation.Language.Parser]::ParseFile(
            $scriptPath,
            [ref]$tokens,
            [ref]$parseErrors)
        @($parseErrors).Count | Should Be 0
        $source = Get-Content -LiteralPath $scriptPath -Raw

        $source | Should Match "Import-Module .*ArtifactOperationLock\.psm1"
        $source | Should Match "Import-Module .*ProcessStreaming\.psm1"
        $source | Should Match "coordination = 'shared-dependency'"
        $source | Should Match 'Sort-Object \{ if \(\$_ -eq ''x64''\) \{ 0 \} else \{ 1 \} \} -Unique'
        $source | Should Match 'Invoke-RSStreamingProcess\s+`\r?\n\s+-FilePath \$vcpkgExePath'
        $source | Should Match '(?s)Invoke-RSStreamingProcess.+?-DelegateArtifactOperation'
        $source | Should Not Match '& \$vcpkgExePath\s+install'
        $source | Should Match '\[string\]\$VcpkgRoot\s*=\s*\$null'
        $source | Should Match 'foreach \(\$requiredDirectory in @\(''ports'', ''versions''\)\)'
        $source | Should Match '\$vcpkgArguments\.Add\("--vcpkg-root=\$vcpkgRootPath"\)'
        $source | Should Match 'Assert-RSVcpkgManifestVersionConstraintsAvailable[\s\S]*?-RegistryRoot \$vcpkgRegistryRoot'
        $source.IndexOf('Assert-RSVcpkgManifestVersionConstraintsAvailable') | Should BeLessThan $source.IndexOf('$sharedDependencyLocks =')
        $source | Should Match '-Arguments \$vcpkgArguments\.ToArray\(\)'
        $source | Should Match '\$installBaseRoot\s*=.*\.build\\vcpkg_installed'
        $source | Should Match '\$platformScopeName\s*=\s*\(Get-RSVcpkgCoordinationPlatform -TripletName \$triplet\)\.ToLowerInvariant\(\)'
        $source | Should Match 'Resolve-RSVcpkgSafeChildPath -Root \$installBaseRoot -Child \$platformScopeName'
        $source | Should Match 'Resolve-RSVcpkgSafeChildPath -Root \$platformInstallRoot -Child \$triplet'
        $source | Should Match 'Assert-RSNoResidualArtifactToolProcesses[\s\S]*?-Scope \$sharedDependencyScope'
        $source | Should Match 'finally\s*\{[\s\S]*?Exit-RSArtifactOperationLock'
    }

    It 'accepts vcpkg triplet leaf names' {
        Assert-RSVcpkgTripletLeafName -Triplet 'x64-windows' | Should Be 'x64-windows'
        Assert-RSVcpkgTripletLeafName -Triplet 'arm64-windows' | Should Be 'arm64-windows'
        Assert-RSVcpkgTripletLeafName -Triplet 'x64-windows-static-md' | Should Be 'x64-windows-static-md'
    }

    It 'rejects triplet values that are paths' {
        Assert-ThrowsTerminatingError { Assert-RSVcpkgTripletLeafName -Triplet '..' }
        Assert-ThrowsTerminatingError { Assert-RSVcpkgTripletLeafName -Triplet '..\..' }
        Assert-ThrowsTerminatingError { Assert-RSVcpkgTripletLeafName -Triplet 'x64\windows' }
        Assert-ThrowsTerminatingError { Assert-RSVcpkgTripletLeafName -Triplet 'x64/windows' }
    }

    It 'resolves child paths under the intended root' {
        $root = New-RSTemporaryVcpkgInstallSafetyRoot
        try {
            $path = Resolve-RSVcpkgSafeChildPath -Root $root -Child 'x64-windows' -Description 'test triplet'

            $path | Should Be (Join-Path $root 'x64-windows')
        } finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'rejects child paths that escape the intended root' {
        $root = New-RSTemporaryVcpkgInstallSafetyRoot
        try {
            Assert-ThrowsTerminatingError { Resolve-RSVcpkgSafeChildPath -Root $root -Child '..\outside' -Description 'test triplet' }
            Assert-ThrowsTerminatingError { Resolve-RSVcpkgSafeChildPath -Root $root -Child '..\..' -Description 'test triplet' }
        } finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'accepts manifest version floors present in the selected registry' {
        $root = New-RSTemporaryVcpkgInstallSafetyRoot
        try {
            $manifestPath = Join-Path $root 'vcpkg.json'
            $registryRoot = Join-Path $root 'registry'
            $versionDirectory = Join-Path $registryRoot 'versions\a-'
            New-Item -ItemType Directory -Path $versionDirectory -Force | Out-Null
            [System.IO.File]::WriteAllText($manifestPath, '{"dependencies":[{"name":"alpha","version>=":"1.2.3#2"}]}')
            [System.IO.File]::WriteAllText(
                (Join-Path $versionDirectory 'alpha.json'),
                '{"versions":[{"version-semver":"1.2.3","port-version":2,"git-tree":"0123456789abcdef"}]}')

            $available = Invoke-RSTestBlock {
                Assert-RSVcpkgManifestVersionConstraintsAvailable -ManifestPath $manifestPath -RegistryRoot $registryRoot
            }
            $available.Succeeded | Should Be $true
        } finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'rejects a manifest floor absent from the selected registry before install' {
        $root = New-RSTemporaryVcpkgInstallSafetyRoot
        try {
            $manifestPath = Join-Path $root 'vcpkg.json'
            $registryRoot = Join-Path $root 'registry'
            $versionDirectory = Join-Path $registryRoot 'versions\a-'
            New-Item -ItemType Directory -Path $versionDirectory -Force | Out-Null
            [System.IO.File]::WriteAllText($manifestPath, '{"dependencies":[{"name":"alpha","version>=":"1.2.4"}]}')
            [System.IO.File]::WriteAllText(
                (Join-Path $versionDirectory 'alpha.json'),
                '{"versions":[{"version-semver":"1.2.3","port-version":2,"git-tree":"0123456789abcdef"}]}')

            $errorMessage = $null
            try {
                Assert-RSVcpkgManifestVersionConstraintsAvailable -ManifestPath $manifestPath -RegistryRoot $registryRoot
            } catch {
                $errorMessage = $_.Exception.Message
            }

            $errorMessage | Should Match 'Update vcpkg\.json builtin-baseline, dependency floors, and exact overrides together'
            $errorMessage | Should Match 'alpha >= 1\.2\.4'
        } finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'accepts an exact manifest override present in the selected registry' {
        $root = New-RSTemporaryVcpkgInstallSafetyRoot
        try {
            $manifestPath = Join-Path $root 'vcpkg.json'
            $registryRoot = Join-Path $root 'registry'
            $versionDirectory = Join-Path $registryRoot 'versions\a-'
            New-Item -ItemType Directory -Path $versionDirectory -Force | Out-Null
            [System.IO.File]::WriteAllText(
                $manifestPath,
                '{"dependencies":[{"name":"alpha","version>=":"1.2.3"}],"overrides":[{"name":"alpha","version":"1.2.3"}]}')
            [System.IO.File]::WriteAllText(
                (Join-Path $versionDirectory 'alpha.json'),
                '{"versions":[{"version":"1.2.5","port-version":0,"git-tree":"fedcba9876543210"},{"version":"1.2.3","port-version":0,"git-tree":"0123456789abcdef"}]}')

            $available = Invoke-RSTestBlock {
                Assert-RSVcpkgManifestVersionConstraintsAvailable -ManifestPath $manifestPath -RegistryRoot $registryRoot
            }
            $available.Succeeded | Should Be $true
        } finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'rejects an exact manifest override absent from the selected registry' {
        $root = New-RSTemporaryVcpkgInstallSafetyRoot
        try {
            $manifestPath = Join-Path $root 'vcpkg.json'
            $registryRoot = Join-Path $root 'registry'
            $versionDirectory = Join-Path $registryRoot 'versions\a-'
            New-Item -ItemType Directory -Path $versionDirectory -Force | Out-Null
            [System.IO.File]::WriteAllText(
                $manifestPath,
                '{"dependencies":[{"name":"alpha","version>=":"1.2.3"}],"overrides":[{"name":"alpha","version":"1.2.3"}]}')
            [System.IO.File]::WriteAllText(
                (Join-Path $versionDirectory 'alpha.json'),
                '{"versions":[{"version":"1.2.5","port-version":0,"git-tree":"fedcba9876543210"}]}')

            $errorMessage = $null
            try {
                Assert-RSVcpkgManifestVersionConstraintsAvailable -ManifestPath $manifestPath -RegistryRoot $registryRoot
            } catch {
                $errorMessage = $_.Exception.Message
            }

            $errorMessage | Should Match 'alpha override 1\.2\.3'
        } finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'keeps the audited AWS SDK serializer dependency on one exact override' {
        $manifest = Get-Content -LiteralPath (Join-Path $repoRoot 'vcpkg.json') -Raw | ConvertFrom-Json
        $dependency = @($manifest.dependencies | Where-Object { $_.name -eq 'aws-sdk-cpp' })
        $override = @($manifest.overrides | Where-Object { $_.name -eq 'aws-sdk-cpp' })

        $dependency.Count | Should Be 1
        $dependency[0].'version>=' | Should Be '1.11.880'
        $override.Count | Should Be 1
        $override[0].version | Should Be '1.11.880'
    }

    It 'merges a single source file while StrictMode is active' {
        $root = New-RSTemporaryVcpkgInstallSafetyRoot
        try {
            $src = Join-Path $root 'src\test-triplet'
            $dst = Join-Path $root 'dst\test-triplet'
            New-Item -ItemType Directory -Path (Join-Path $src 'include') -Force | Out-Null
            New-Item -ItemType Directory -Path $dst -Force | Out-Null
            [System.IO.File]::WriteAllText((Join-Path $src 'include\header.h'), '// v1')

            $result = Merge-RSVcpkgTripletSafe -SourcePath $src -DestinationPath $dst -TripletName 'test-triplet'

            $result.FilesTotal | Should Be 1
            $result.FilesCopied | Should Be 1
            Test-Path -LiteralPath (Join-Path $dst 'include\header.h') | Should Be $true
        } finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function New-RSFixtureVisualStudioInstallation {
    param([string]$Root, [string]$VersionFileText = '', [switch]$NoVersionFile)
    $installation = Join-Path $Root 'Microsoft Visual Studio\18\Insiders'
    $buildDirectory = Join-Path $installation 'VC\Auxiliary\Build'
    New-Item -ItemType Directory -Path $buildDirectory -Force | Out-Null
    if (-not $NoVersionFile) {
        [System.IO.File]::WriteAllText((Join-Path $buildDirectory 'Microsoft.VCToolsVersion.default.txt'), $VersionFileText)
    }
    return $installation
}

Describe 'Vcpkg default MSVC toolset pin' {
    It 'reads the default toolset from the installation, not the newest one installed beside it' {
        $root = New-RSTemporaryVcpkgInstallSafetyRoot
        try {
            $installation = New-RSFixtureVisualStudioInstallation -Root $root -VersionFileText "14.51.36231`r`n"
            # A newer toolset beside the default, as VS 18 Insiders 14.52 sits beside 14.51: it is never consulted.
            New-Item -ItemType Directory -Path (Join-Path $installation 'VC\Tools\MSVC\14.52.36725') -Force | Out-Null

            $toolset = Get-RSVcpkgDefaultToolset -Installation $installation
            $toolset.Version | Should Be '14.51.36231'
            $toolset.MajorMinor | Should Be '14.51'
            $toolset.Installation | Should Be $installation
            $toolset.VersionFile | Should Be (Join-Path $installation 'VC\Auxiliary\Build\Microsoft.VCToolsVersion.default.txt')
        } finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'fails with the file and the repair when the default toolset version is missing or malformed' {
        $root = New-RSTemporaryVcpkgInstallSafetyRoot
        try {
            $missing = New-RSFixtureVisualStudioInstallation -Root (Join-Path $root 'missing') -NoVersionFile
            { Get-RSVcpkgDefaultToolset -Installation $missing } | Should Throw 'does not exist'
            { Get-RSVcpkgDefaultToolset -Installation $missing } | Should Throw 'Install the Desktop development with C++ workload'

            foreach ($case in @(
                    @{ Name = 'empty'; Text = '' },
                    @{ Name = 'major.minor only'; Text = '14.51' },
                    @{ Name = 'a platform toolset name'; Text = 'v145' },
                    @{ Name = 'two lines'; Text = "14.51.36231`r`n14.52.36725" },
                    @{ Name = 'a suffix'; Text = '14.51.36231-preview' })) {
                $installation = New-RSFixtureVisualStudioInstallation -Root (Join-Path $root $case.Name.Replace(' ', '-').Replace('.', '-')) -VersionFileText $case.Text
                $message = $null
                try { Get-RSVcpkgDefaultToolset -Installation $installation } catch { $message = $_.Exception.Message }
                ('{0}: {1}' -f $case.Name, $message) | Should Match 'not a toolset version such as 14\.51\.36231'
                $message | Should Match 'Microsoft\.VCToolsVersion\.default\.txt'
                $message | Should Match 'Repair the installation in the Visual Studio Installer'
            }
        } finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'finds the installation the way the build does' {
        $savedActions = $env:GITHUB_ACTIONS
        $savedMsbuild = $env:MSBUILD_EXE_PATH
        try {
            # GitHub Actions: the workflow-selected MSBuild names the installation, whatever its host directory.
            $env:GITHUB_ACTIONS = 'true'
            foreach ($relative in @('MSBuild\Current\Bin\amd64\MSBuild.exe', 'MSBuild\Current\Bin\arm64\MSBuild.exe', 'MSBuild\Current\Bin\MSBuild.exe')) {
                $env:MSBUILD_EXE_PATH = 'C:\Program Files\Microsoft Visual Studio\18\Enterprise\' + $relative
                Get-RSVcpkgVisualStudioInstallation | Should Be 'C:\Program Files\Microsoft Visual Studio\18\Enterprise'
            }
            $env:MSBUILD_EXE_PATH = 'C:\Tools\MSBuild.exe'
            { Get-RSVcpkgVisualStudioInstallation } | Should Throw 'not inside a Visual Studio installation'

            # Anywhere else: vswhere's newest installation with MSBuild, which this machine must have to build at all.
            $env:GITHUB_ACTIONS = $null
            $env:MSBUILD_EXE_PATH = $null
            $installation = Get-RSVcpkgVisualStudioInstallation
            Test-Path -LiteralPath (Join-Path $installation 'MSBuild\Current\Bin') | Should Be $true
            Test-Path -LiteralPath (Join-Path $installation 'VC\Auxiliary\Build\Microsoft.VCToolsVersion.default.txt') | Should Be $true
        } finally {
            $env:GITHUB_ACTIONS = $savedActions
            $env:MSBUILD_EXE_PATH = $savedMsbuild
        }
    }

    It 'writes the stock triplet unchanged in its own line endings, then the two pins' {
        $installation = 'C:\Program Files\Microsoft Visual Studio\18\Insiders'
        $crlf = "set(VCPKG_TARGET_ARCHITECTURE arm64)`r`nset(VCPKG_CRT_LINKAGE dynamic)`r`n"
        $lf = "set(VCPKG_TARGET_ARCHITECTURE x64)`nset(VCPKG_CRT_LINKAGE dynamic)`n"

        $arm = Get-RSVcpkgOverlayTripletText -StockText $crlf -Installation $installation -ToolsetVersion '14.51'
        $arm.StartsWith($crlf) | Should Be $true
        $arm | Should Match 'set\(VCPKG_VISUAL_STUDIO_PATH "C:\\\\Program Files\\\\Microsoft Visual Studio\\\\18\\\\Insiders"\)\r\n'
        $arm | Should Match 'set\(VCPKG_PLATFORM_TOOLSET_VERSION "14\.51"\)\r\n$'
        ([regex]::Matches($arm, "(?<!`r)`n")).Count | Should Be 0

        $x64 = Get-RSVcpkgOverlayTripletText -StockText $lf -Installation $installation -ToolsetVersion '14.51'
        $x64.StartsWith($lf) | Should Be $true
        $x64.Contains("`r") | Should Be $false
        $x64 | Should Match 'set\(VCPKG_PLATFORM_TOOLSET_VERSION "14\.51"\)\n$'

        # A stock file without a final newline still gets its pins on their own lines; an empty one is only the pins.
        (Get-RSVcpkgOverlayTripletText -StockText 'set(A 1)' -Installation $installation -ToolsetVersion '14.51').StartsWith("set(A 1)`n`n# RedSalamander") | Should Be $true
        (Get-RSVcpkgOverlayTripletText -StockText '' -Installation $installation -ToolsetVersion '14.51').StartsWith('# RedSalamander') | Should Be $true
        # CMake quoting: a backslash, a quote and a dollar sign are escaped, nothing else.
        (Get-RSVcpkgOverlayTripletText -StockText '' -Installation 'C:\a "b" $c' -ToolsetVersion '14.51') | Should Match ([regex]::Escape('set(VCPKG_VISUAL_STUDIO_PATH "C:\\a \"b\" \$c")'))
        # The pin is major.minor: a full version is not accepted.
        { Get-RSVcpkgOverlayTripletText -StockText '' -Installation $installation -ToolsetVersion '14.51.36231' } | Should Throw 'ToolsetVersion'
    }

    It 'rewrites the overlay only when its bytes change' {
        $root = New-RSTemporaryVcpkgInstallSafetyRoot
        try {
            $stock = Join-Path $root 'triplets\arm64-windows.cmake'
            New-Item -ItemType Directory -Path (Split-Path -Parent $stock) -Force | Out-Null
            [System.IO.File]::WriteAllText($stock, "set(VCPKG_TARGET_ARCHITECTURE arm64)`r`n")
            $toolset = [pscustomobject]@{ Installation = 'C:\VS'; MajorMinor = '14.51' }
            $output = Join-Path $root 'overlay\arm64'

            $first = Update-RSVcpkgOverlayTriplet -StockTripletPath $stock -Toolset $toolset -OutputDirectory $output
            $first.Changed | Should Be $true
            $first.Triplet | Should Be 'arm64-windows'
            $first.Directory | Should Be $output
            $first.Path | Should Be (Join-Path $output 'arm64-windows.cmake')
            # Backdate the file: an unchanged restore must leave its timestamp alone, so a rewrite would show.
            $written = [datetime]::new(2020, 1, 2, 3, 4, 5, [System.DateTimeKind]::Utc)
            [System.IO.File]::SetLastWriteTimeUtc($first.Path, $written)

            $second = Update-RSVcpkgOverlayTriplet -StockTripletPath $stock -Toolset $toolset -OutputDirectory $output
            $second.Changed | Should Be $false
            (Get-Item -LiteralPath $second.Path).LastWriteTimeUtc | Should Be $written

            # A patch update of the default toolset keeps the major.minor pin, so it changes nothing; a minor update does.
            (Update-RSVcpkgOverlayTriplet -StockTripletPath $stock -Toolset ([pscustomobject]@{ Installation = 'C:\VS'; MajorMinor = '14.51' }) -OutputDirectory $output).Changed | Should Be $false
            $third = Update-RSVcpkgOverlayTriplet -StockTripletPath $stock -Toolset ([pscustomobject]@{ Installation = 'C:\VS'; MajorMinor = '14.52' }) -OutputDirectory $output
            $third.Changed | Should Be $true
            [System.IO.File]::ReadAllText($third.Path) | Should Match 'set\(VCPKG_PLATFORM_TOOLSET_VERSION "14\.52"\)'
            # The stock triplet itself is never modified.
            [System.IO.File]::ReadAllText($stock) | Should Be "set(VCPKG_TARGET_ARCHITECTURE arm64)`r`n"
            { Update-RSVcpkgOverlayTriplet -StockTripletPath (Join-Path $root 'absent.cmake') -Toolset $toolset -OutputDirectory $output } | Should Throw 'does not exist'
        } finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'finds the stock triplet in triplets, then triplets\community, and keeps the governed leaf-name rule' {
        $root = New-RSTemporaryVcpkgInstallSafetyRoot
        try {
            New-Item -ItemType Directory -Path (Join-Path $root 'triplets\community') -Force | Out-Null
            [System.IO.File]::WriteAllText((Join-Path $root 'triplets\x64-windows.cmake'), '# stock')
            [System.IO.File]::WriteAllText((Join-Path $root 'triplets\community\x64-windows-static.cmake'), '# community')
            [System.IO.File]::WriteAllText((Join-Path $root 'triplets\community\x64-windows.cmake'), '# shadowed by the stock one')

            Resolve-RSVcpkgStockTripletPath -RegistryRoot $root -Triplet 'x64-windows' | Should Be (Join-Path $root 'triplets\x64-windows.cmake')
            Resolve-RSVcpkgStockTripletPath -RegistryRoot $root -Triplet 'x64-windows-static' | Should Be (Join-Path $root 'triplets\community\x64-windows-static.cmake')
            { Resolve-RSVcpkgStockTripletPath -RegistryRoot $root -Triplet 'arm64-windows' } | Should Throw 'is not one of the triplets of the selected vcpkg root'
            { Resolve-RSVcpkgStockTripletPath -RegistryRoot $root -Triplet '..\x64-windows' } | Should Throw 'pass a triplet leaf name'
        } finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'pins every triplet the installer builds, before any lock, staging directory or download' {
        $scriptPath = Join-Path $repoRoot 'vcpkg-install.ps1'
        $source = Get-Content -LiteralPath $scriptPath -Raw

        $pin = $source.IndexOf('$msvcToolset = Get-RSVcpkgDefaultToolset')
        $pin | Should BeGreaterThan $source.IndexOf('Assert-RSVcpkgManifestVersionConstraintsAvailable')
        $pin | Should BeLessThan $source.IndexOf('$sharedDependencyLocks =')
        $pin | Should BeLessThan $source.IndexOf('New-Item -ItemType Directory -Path $stagingRoot')
        $source | Should Match '\$visualStudioInstallation = Get-RSVcpkgVisualStudioInstallation'

        # The overlay is written and passed for each triplet in the install loop, from the selected vcpkg root's own triplet.
        $loopStart = $source.IndexOf('foreach ($triplet in $tripletsToInstall) {')
        $overlay = $source.IndexOf('Update-RSVcpkgOverlayTriplet')
        $arguments = $source.IndexOf('$vcpkgArguments = [System.Collections.Generic.List[string]]::new()')
        $loopStart | Should BeGreaterThan 0
        $overlay | Should BeGreaterThan $loopStart
        $overlay | Should BeLessThan $arguments
        $source | Should Match 'Resolve-RSVcpkgStockTripletPath -RegistryRoot \$vcpkgRegistryRoot -Triplet \$triplet'
        $source | Should Match '"--overlay-triplets=\$\(\$tripletOverlay\.Directory\)"'
        $source | Should Match 'Resolve-RSVcpkgSafeChildPath -Root \$tripletOverlayRoot -Child \$platformScopeName'
        $source | Should Match 'Resolve-RSVcpkgSafeChildPath -Root \$repoRoot -Child "\.build\\vcpkg-triplets"'
        # The governed custom-triplet rules stay in force.
        $source | Should Match 'Assert-RSVcpkgTripletLeafName'
        $source | Should Match 'has no reviewed RedSalamander platform lock mapping'
    }

    $pin = (Get-Content -LiteralPath (Join-Path $repoRoot 'Dependencies\DxUi.lock.json') -Raw | ConvertFrom-Json).commit
    $dxuiSource = Get-DxUiSourcePath -DependencyRoot (Join-Path $repoRoot '.build\dependencies\DxUi') -Commit $pin
    $dxuiTools = Join-Path $dxuiSource 'Tools'
    $dxuiModulesRestored = (Test-Path -LiteralPath (Join-Path $dxuiTools 'VisualStudio.psm1')) -and (Test-Path -LiteralPath (Join-Path $dxuiTools 'VcpkgTriplet.psm1'))

    It 'matches the pinned DxUi modules it mirrors' -Skip:(-not $dxuiModulesRestored) {
        # The short directory is only a locator; verify the complete identity
        # before importing executable code from the restored dependency.
        $restoredHead = Invoke-RSTestFixtureGit -Arguments @('-C'; $dxuiSource; 'rev-parse'; 'HEAD') 2>$null
        $LASTEXITCODE | Should Be 0
        ([string]$restoredHead).Trim() | Should Be $pin
        # This installer cannot import the pinned checkout (it runs before the restore), so it keeps DxUi's contract itself. Whenever
        # the checkout is restored, both implementations must agree on the installation, the toolset and the overlay file they write.
        # Only the documented consumer interface is called: Get-DxUiVisualStudioInstallation, Get-DxUiDefaultToolset and
        # Update-DxUiVcpkgOverlayTriplet.
        Import-Module (Join-Path $dxuiTools 'VisualStudio.psm1') -Force
        Import-Module (Join-Path $dxuiTools 'VcpkgTriplet.psm1') -Force
        $root = New-RSTemporaryVcpkgInstallSafetyRoot
        $savedActions = $env:GITHUB_ACTIONS
        try {
            $env:GITHUB_ACTIONS = $null
            $installation = Get-RSVcpkgVisualStudioInstallation
            $installation | Should Be (Get-DxUiVisualStudioInstallation)
            $ours = Get-RSVcpkgDefaultToolset -Installation $installation
            $theirs = Get-DxUiDefaultToolset -Installation $installation
            $ours.Version | Should Be $theirs.Version
            $ours.MajorMinor | Should Be $theirs.MajorMinor
            $ours.VersionFile | Should Be $theirs.VersionFile

            # Each side writes its own header comment; the stock text and the two pins must match byte for byte.
            $withoutComment = { param($Path) [System.IO.File]::ReadAllText($Path) -replace '(?m)^# (DxUi|RedSalamander): [^\r\n]*(\r?\n)', "# pinned`$2" }
            $index = 0
            foreach ($stockText in @("set(A 1)`r`nset(B `"\x`")`r`n", "set(A 1)`nset(B 2)", '')) {
                $index++
                $stock = Join-Path $root "stock$index\triplets\sample.cmake"
                New-Item -ItemType Directory -Path (Split-Path -Parent $stock) -Force | Out-Null
                [System.IO.File]::WriteAllText($stock, $stockText)
                $toolset = [pscustomobject]@{ Installation = 'C:\VS "x" $y'; MajorMinor = '14.51' }
                $a = Update-RSVcpkgOverlayTriplet -StockTripletPath $stock -Toolset $toolset -OutputDirectory (Join-Path $root "ours$index")
                $b = Update-DxUiVcpkgOverlayTriplet -StockTripletPath $stock -Toolset $toolset -OutputDirectory (Join-Path $root "theirs$index")
                (& $withoutComment $a.Path) | Should Be (& $withoutComment $b.Path)
            }
        } finally {
            $env:GITHUB_ACTIONS = $savedActions
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

Describe 'Vcpkg dependency pins' {
    # Dependabot moved builtin-baseline to a vcpkg release tag older than the tool pin (#98, #111) and raised floors past the
    # pinned version database and past the audited AWS override (#112, #113) while CI could not start, so master stopped
    # restoring. These checks are offline: they read the repository's manifests and, when present, the pinned tool checkout.
    $manifest = Get-Content -LiteralPath (Join-Path $repoRoot 'vcpkg.json') -Raw | ConvertFrom-Json
    $toolPin = Get-Content -LiteralPath (Join-Path $repoRoot 'vcpkg-tool.json') -Raw | ConvertFrom-Json

    # vcpkg version text: a dotted or dashed base, then an optional #port-version. Numeric parts compare as numbers.
    function Compare-RSVcpkgVersionText {
        param([string]$Left, [string]$Right)

        $split = {
            param([string]$Text)
            $hash = $Text.IndexOf('#')
            $port = if ($hash -ge 0) { [int]$Text.Substring($hash + 1) } else { 0 }
            $base = if ($hash -ge 0) { $Text.Substring(0, $hash) } else { $Text }
            return @{ Parts = @($base -split '[.-]'); Port = $port }
        }
        $a = & $split $Left
        $b = & $split $Right
        for ($i = 0; $i -lt [Math]::Max($a.Parts.Count, $b.Parts.Count); $i++) {
            $x = if ($i -lt $a.Parts.Count) { $a.Parts[$i] } else { '0' }
            $y = if ($i -lt $b.Parts.Count) { $b.Parts[$i] } else { '0' }
            $xn = [long]0
            $yn = [long]0
            $order = if ([long]::TryParse($x, [ref]$xn) -and [long]::TryParse($y, [ref]$yn)) { $xn.CompareTo($yn) } else { [string]::CompareOrdinal($x, $y) }
            if ($order -ne 0) { return [Math]::Sign($order) }
        }
        return $a.Port.CompareTo($b.Port)
    }

    It 'compares vcpkg version text the way the pin checks need' {
        Compare-RSVcpkgVersionText '1.11.900' '1.11.880' | Should Be 1
        Compare-RSVcpkgVersionText '1.11.880' '1.11.880' | Should Be 0
        Compare-RSVcpkgVersionText '3.53.4#2' '3.53.4#1' | Should Be 1
        Compare-RSVcpkgVersionText '3.53.4' '3.53.4#0' | Should Be 0
        Compare-RSVcpkgVersionText '0.43.4' '0.43.10' | Should Be -1
    }

    It 'pins builtin-baseline to the vcpkg tool commit' {
        # A reviewed lane that needs the two pins to diverge records the exact baseline -> tool pair here and in
        # Build_Toolchain.md. Dependabot proposes the newest release tag, which can be OLDER than a master-commit pin.
        $reviewedPinDivergence = @{}

        $baseline = [string]$manifest.'builtin-baseline'
        $baseline | Should Match '^[0-9a-f]{40}$'
        [string]$toolPin.commit | Should Match '^[0-9a-f]{40}$'
        if ($reviewedPinDivergence.ContainsKey($baseline)) {
            [string]$toolPin.commit | Should Be $reviewedPinDivergence[$baseline]
        } else {
            $baseline | Should Be ([string]$toolPin.commit)
        }
    }

    It 'keeps every direct floor at or below its exact override and the audited AWS SDK and CRT pair together' {
        # The SDK tag's crt/aws-crt-cpp submodule names the CRT the pair was audited with (Build_Toolchain.md, FileSystem_S3.md).
        $auditedAwsPairs = @{ '1.11.880' = '0.43.4' }

        $overrides = @{}
        foreach ($override in @($manifest.overrides)) {
            $overrides[[string]$override.name] = [string]$override.version
        }
        $violations = [System.Collections.Generic.List[string]]::new()
        foreach ($dependency in @($manifest.dependencies)) {
            if ($dependency -is [string] -or $null -eq $dependency.PSObject.Properties['version>=']) { continue }
            $name = [string]$dependency.name
            if (-not $overrides.ContainsKey($name)) { continue }
            $floor = [string]$dependency.'version>='
            if ((Compare-RSVcpkgVersionText $floor $overrides[$name]) -gt 0) {
                # vcpkg silently builds the override, so the manifest would advertise a version that is never built.
                [void]$violations.Add("$name >= $floor exceeds its exact override $($overrides[$name])")
            }
        }
        ($violations -join '; ') | Should Be ''

        $overrides.ContainsKey('aws-sdk-cpp') | Should Be $true
        $auditedAwsPairs.ContainsKey($overrides['aws-sdk-cpp']) | Should Be $true
        $overrides['aws-crt-cpp'] | Should Be $auditedAwsPairs[$overrides['aws-sdk-cpp']]
    }

    It 'keeps Dependabot away from the vcpkg baseline and the audited AWS pair' {
        $text = Get-Content -LiteralPath (Join-Path $repoRoot '.github\dependabot.yml') -Raw
        $vcpkgBlocks = @([regex]::Split($text, '(?m)^\s*-\s*package-ecosystem:\s*') |
                Where-Object { $_ -match '^["'']?vcpkg["'']?\s*(\r?\n|$)' })
        $vcpkgBlocks.Count | Should Be 1
        $vcpkgBlocks[0] | Should Match '(?m)^\s*ignore:\s*$'
        foreach ($name in @('github.com/microsoft/vcpkg', 'aws-sdk-cpp', 'aws-crt-cpp')) {
            $vcpkgBlocks[0] | Should Match ('(?m)^\s*-\s*dependency-name:\s*["'']?' + [regex]::Escape($name) + '["'']?\s*$')
        }
    }
    It 'keeps Dependabot from proposing any vcpkg port or baseline update' {
        # A port floor is restorable only if the pinned baseline contains it, and Dependabot cannot move the baseline
        # together with the floor, so each of its vcpkg PRs is unmergeable alone (#111-#113, #117). The vcpkg block must
        # carry an unrestricted wildcard entry under ignore: (limit 0 would not cover security updates, and allow: would
        # permit the dependency instead of ignoring it).
        $text = Get-Content -LiteralPath (Join-Path $repoRoot '.github\dependabot.yml') -Raw
        $vcpkgBlocks = @([regex]::Split($text, '(?m)^\s*-\s*package-ecosystem:\s*') |
                Where-Object { $_ -match '^["'']?vcpkg["'']?\s*(\r?\n|$)' })
        $vcpkgBlocks.Count | Should Be 1
        # Cut out the ignore: list only: from the key up to the next non-comment line at the same or a lower indent.
        $ignoreLines = New-Object System.Collections.Generic.List[string]
        $ignoreIndent = -1
        foreach ($line in ($vcpkgBlocks[0] -split '\r?\n')) {
            if ($ignoreIndent -lt 0) {
                if ($line -match '^(\s*)ignore:\s*(#.*)?$') { $ignoreIndent = $Matches[1].Length }
                continue
            }
            if ($line -match '^\s*(#.*)?$') { continue }
            if ((($line -replace '^(\s*).*$', '$1').Length -le $ignoreIndent) -and ($line -notmatch '^\s*-\s')) { break }
            $ignoreLines.Add($line)
        }
        $ignoreIndent | Should Not BeLessThan 0
        # The wildcard entry counts only when no update-types/versions line narrows it before the next list entry.
        $silenced = $false
        foreach ($entry in [regex]::Split(($ignoreLines -join "`n"), '(?m)^\s*-\s*(?=dependency-name:)')) {
            if ($entry -match '^dependency-name:\s*["'']?\*["'']?\s*(\r?\n|$)' -and $entry -notmatch '(?m)^\s*(update-types|versions):') {
                $silenced = $true
            }
        }
        $silenced | Should Be $true
    }

    # These cases prove manifest data against immutable builtin-baseline Git objects.
    # They do not qualify the current vcpkg executable checkout or bootstrap state.
    $toolObjectStore = Join-Path $repoRoot '.build\vcpkg-tool'
    if (-not (Test-Path -LiteralPath (Join-Path $toolObjectStore '.git'))) {
        # Consumer restores keep their object store in the identity-verified short
        # output root. Its active checkout need not equal our baseline: only exact
        # immutable objects are read below, without fetch, checkout or executable use.
        $dependency = Join-Path $repoRoot '.build/dependencies/DxUi'
        $dependencyPin = Get-Content -LiteralPath (Join-Path $repoRoot 'Dependencies/DxUi.lock.json') -Raw | ConvertFrom-Json
        $source = Get-DxUiSourcePath -DependencyRoot $dependency -Commit $dependencyPin.commit
        foreach ($platform in @('x64', 'ARM64')) {
            $identityPath = Join-Path $dependency "DxUi.identity.$platform.json"
            if (-not (Test-Path -LiteralPath $identityPath -PathType Leaf)) { continue }
            $identity = Get-Content -LiteralPath $identityPath -Raw | ConvertFrom-Json
            if ($identity.Identity.commit -cne $dependencyPin.commit -or $identity.Identity.platform -cne $platform) {
                throw 'Pinned-baseline object-store locator has a stale DxUi dependency identity.'
            }
            $sourceHead = Invoke-RSTestFixtureGit -Arguments @('--no-optional-locks'; '--no-lazy-fetch'; '-C'; $source; 'rev-parse'; 'HEAD') 2>&1
            if ($LASTEXITCODE -ne 0 -or ([string]$sourceHead).Trim() -cne $dependencyPin.commit) {
                throw 'Pinned-baseline object-store locator does not match the complete DxUi source pin.'
            }
            $output = Read-DxUiResolvedOutputRoot -DependencyRoot $dependency -Platform $platform -BuildIdentity $identity
            $candidate = Join-Path $output 'vcpkg-tool'
            if (-not (Test-Path -LiteralPath (Join-Path $candidate '.git'))) { continue }
            $origin = Invoke-RSTestFixtureGit -Arguments @('--no-optional-locks'; '--no-lazy-fetch'; '-C'; $candidate; 'remote'; 'get-url'; 'origin') 2>&1
            $expectedOrigin = (Get-Content -LiteralPath (Join-Path $repoRoot 'vcpkg-tool.json') -Raw | ConvertFrom-Json).repository
            if ($LASTEXITCODE -ne 0 -or (([string]$origin).Trim().TrimEnd('/') -replace '\.git$', '') -cne $expectedOrigin) {
                throw 'Pinned-baseline object store does not have the pinned microsoft/vcpkg origin.'
            }
            $toolObjectStore = $candidate
            break
        }
    }
    $baselineCommit = [string]$manifest.'builtin-baseline'
    if ($baselineCommit -cnotmatch '^[0-9a-f]{40}$') { throw 'vcpkg builtin-baseline must be a full commit SHA before reading Git objects.' }
    Write-Host "Pinned-baseline data object store: $toolObjectStore; requested commit: $baselineCommit"
    $baselinePaths = [System.Collections.Generic.List[string]]::new()
    $packageNames = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $baselinePaths.Add('versions/baseline.json')
    foreach ($dependency in @($manifest.dependencies)) {
        if ($dependency -is [string] -or $null -eq $dependency.PSObject.Properties['version>=']) { continue }
        [void]$packageNames.Add([string]$dependency.name)
    }
    foreach ($override in @($manifest.overrides)) { [void]$packageNames.Add([string]$override.name) }
    foreach ($name in $packageNames) {
        if ($name -cnotmatch '^[a-z0-9][a-z0-9.+-]*$') { throw "Unsafe vcpkg package name in manifest: '$name'" }
        $prefix = $name.Substring(0, 1).ToLowerInvariant()
        $baselinePaths.Add("versions/$prefix-/$name.json")
    }
    foreach ($relative in @(
        'ports/aws-c-common/vcpkg.json',
        'ports/aws-c-common/portfile.cmake',
        'ports/aws-c-common/disable-internal-crt-option.patch'
    )) { $baselinePaths.Add($relative) }
    $baselinePaths = @($baselinePaths | Sort-Object -Unique)

    # --no-lazy-fetch keeps partial/promisor clones offline. A missing commit/tree/blob is
    # an explicit local-object prerequisite skip; an absent expected path or materialization error fails.
    $baselineObjectUnavailableReason = $null
    $baselineBlobIds = @{}
    if (-not (Test-Path -LiteralPath (Join-Path $toolObjectStore '.git'))) {
        $baselineObjectUnavailableReason = "No managed vcpkg Git object store is available at $toolObjectStore."
    } else {
        $commitCheck = @(Invoke-RSTestFixtureGit -Arguments @('--no-lazy-fetch'; '-C'; $toolObjectStore; 'cat-file'; '-e'; "$baselineCommit^{commit}") 2>&1)
        if ($LASTEXITCODE -ne 0) {
            $message = $commitCheck -join "`n"
            if ($message -match '(?i)not a valid object name|could not get object info|missing object|bad object|promisor remote|could not fetch') {
                $baselineObjectUnavailableReason = "Pinned builtin-baseline commit $baselineCommit is absent locally: $message"
            } else {
                throw "Could not inspect pinned builtin-baseline commit ${baselineCommit}: $message"
            }
        } else {
            foreach ($relative in $baselinePaths) {
                $listing = @(Invoke-RSTestFixtureGit -Arguments @('--no-lazy-fetch'; '-C'; $toolObjectStore; 'ls-tree'; '-r'; $baselineCommit; '--'; $relative) 2>&1)
                if ($LASTEXITCODE -ne 0) {
                    $message = $listing -join "`n"
                    if ($message -match '(?i)missing|not a valid|bad object|unable to read|not a tree object|promisor remote') {
                        $baselineObjectUnavailableReason = "Required pinned-baseline tree object for '$relative' is absent locally: $message"
                        break
                    }
                    throw "Could not inspect pinned baseline path '$relative': $message"
                }
                $entry = @($listing | Where-Object { $_ -match "`t$([regex]::Escape($relative))$" })
                if ($entry.Count -ne 1) { throw "Pinned baseline $baselineCommit does not contain exactly one expected file '$relative'." }
                if ($entry[0] -notmatch '^100644 blob ([0-9a-f]{40})\t') { throw "Pinned baseline path '$relative' is not a regular blob: $($entry[0])" }
                $blobId = $Matches[1]
                $blobCheck = @(Invoke-RSTestFixtureGit -Arguments @('--no-lazy-fetch'; '-C'; $toolObjectStore; 'cat-file'; '-e'; "$blobId^{blob}") 2>&1)
                if ($LASTEXITCODE -ne 0) {
                    $message = $blobCheck -join "`n"
                    if ($message -match '(?i)missing|not a valid|bad object|unable to read|promisor remote') {
                        $baselineObjectUnavailableReason = "Required pinned-baseline blob $blobId for '$relative' is absent locally: $message"
                        break
                    }
                    throw "Could not read pinned baseline blob $blobId for '$relative': $message"
                }
                $baselineBlobIds[$relative] = $blobId
            }
        }
    }

    if ($baselineObjectUnavailableReason) { Write-Warning "Skipping immutable vcpkg baseline data cases: $baselineObjectUnavailableReason" }

    # Allocate only inside an executing It, so an excluded Describe never creates scratch.
    # Both data assertions share this fixture; it has no live tool/bootstrap authority.
    $baselineFixtureState = @{ Value = $null }
    $getBaselineFixture = {
        if ($null -ne $baselineFixtureState.Value) { return $baselineFixtureState.Value }
        if ($baselineObjectUnavailableReason) { throw $baselineObjectUnavailableReason }
        $fixtureRoot = New-RSTemporaryVcpkgInstallSafetyRoot
        $contentRoot = Join-Path $fixtureRoot 'pinned-baseline'
        [void](New-Item -ItemType Directory -Path $contentRoot)
        # Reuse the owner's raw Git reader in module scope. Per-OID reads avoid archive
        # attribute traversal into unrelated missing objects in a partial clone.
        $blobReader = Import-Module (Join-Path $repoRoot 'Tools\Modules\Testing\ValidationFingerprint.psm1') -PassThru -ErrorAction Stop
        foreach ($relative in $baselinePaths) {
            $filePath = Resolve-RSVcpkgSafeChildPath -Root $contentRoot -Child $relative
            [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($filePath))
            [byte[]]$blobBytes = & $blobReader {
                param($ObjectStore, $BlobId)
                Invoke-RSValidationGitBinary -RepoRoot $ObjectStore -Argument @('--no-lazy-fetch', 'cat-file', 'blob', $BlobId)
            } $toolObjectStore $baselineBlobIds[$relative]
            [IO.File]::WriteAllBytes($filePath, $blobBytes)
            $actualBlob = Invoke-RSTestFixtureGit -Arguments @('--no-lazy-fetch'; '-C'; $toolObjectStore; 'hash-object'; '--no-filters'; '--'; $filePath)
            if ($LASTEXITCODE -ne 0 -or ([string]$actualBlob).Trim() -cne $baselineBlobIds[$relative]) {
                throw "Pinned baseline file '$relative' does not match its immutable Git blob $($baselineBlobIds[$relative])."
            }
        }
        $actualFiles = @(Get-ChildItem -LiteralPath $contentRoot -File -Recurse | ForEach-Object {
            [IO.Path]::GetRelativePath($contentRoot, $_.FullName).Replace('\', '/')
        } | Sort-Object)
        if (($actualFiles -join "`n") -cne ($baselinePaths -join "`n")) {
            throw "Pinned baseline fixture file set mismatch. Expected [$($baselinePaths -join ', ')], found [$($actualFiles -join ', ')]."
        }
        $baselineFixtureState.Value = [pscustomobject]@{ Root=$contentRoot; Baseline=$baselineCommit; BlobIds=$baselineBlobIds }
        return $baselineFixtureState.Value
    }

    It 'finds every floor and override in the immutable builtin-baseline version database (data only)' -Skip:([bool]$baselineObjectUnavailableReason) {
        $fixture = & $getBaselineFixture
        $errorMessage = $null
        try {
            Assert-RSVcpkgManifestVersionConstraintsAvailable -ManifestPath (Join-Path $repoRoot 'vcpkg.json') -RegistryRoot $fixture.Root
        } catch {
            $errorMessage = $_.Exception.Message
        }
        $errorMessage | Should BeNullOrEmpty
    }

    It 'keeps the ARM64 compatibility overlay on the builtin-baseline source and notices (data only)' -Skip:([bool]$baselineObjectUnavailableReason) {
        $fixture = & $getBaselineFixture
        $baseline = Get-Content -LiteralPath (Join-Path $fixture.Root 'versions/baseline.json') -Raw | ConvertFrom-Json
        $overlayRoot = Join-Path $repoRoot 'Build/vcpkg-ports/aws-c-common'
        $registryRoot = Join-Path $fixture.Root 'ports/aws-c-common'
        $overlay = Get-Content -LiteralPath (Join-Path $overlayRoot 'vcpkg.json') -Raw | ConvertFrom-Json
        $registry = Get-Content -LiteralPath (Join-Path $registryRoot 'vcpkg.json') -Raw | ConvertFrom-Json
        $overlay.version | Should Be $baseline.default.'aws-c-common'.baseline
        $overlay.version | Should Be $registry.version
        ($overlay | ConvertTo-Json -Depth 8 -Compress) | Should Be ($registry | ConvertTo-Json -Depth 8 -Compress)
        $overlayPort = [IO.File]::ReadAllText((Join-Path $overlayRoot 'portfile.cmake'))
        $registryPort = [IO.File]::ReadAllText((Join-Path $registryRoot 'portfile.cmake'))
        $archivePattern = '(?m)^\s*SHA512\s+([0-9a-f]{128})\s*$'
        $expected = [regex]::Match($registryPort, $archivePattern)
        $actual = [regex]::Match($overlayPort, $archivePattern)
        $expected.Success | Should Be $true
        $actual.Success | Should Be $true
        $actual.Groups[1].Value | Should Be $expected.Groups[1].Value
        $overlayPatch = [IO.File]::ReadAllText((Join-Path $overlayRoot 'disable-internal-crt-option.patch')).Replace("`r`n", "`n")
        $registryPatch = [IO.File]::ReadAllText((Join-Path $registryRoot 'disable-internal-crt-option.patch')).Replace("`r`n", "`n")
        $overlayPatch | Should Be $registryPatch
        $overlayPort | Should Match 'msvc1452-arm64\.cmake'
        $overlayPort | Should Match '"\$\{SOURCE_PATH\}/NOTICE"'
    }
}

Describe 'Hosted vcpkg install parity' {
    It 'delegates the hosted install to the pinned repository installer' {
        $workflow = Get-Content -LiteralPath (Join-Path $repoRoot '.github/workflows/build-reusable.yml') -Raw
        $workflow | Should Match '(?s)- name: Install dependencies.*?\$platform = ''\$\{\{ inputs\.platform \}\}''.*?vcpkg-install\.ps1 -Platform \$platform .*?-VcpkgRoot \$env:VCPKG_REGISTRY_ROOT -ReuseCurrent'
        $workflow | Should Not Match '--x-install-root|Normalize vcpkg install root'
        $source = Get-Content -LiteralPath (Join-Path $repoRoot 'vcpkg-install.ps1') -Raw
        $source | Should Match '--overlay-triplets'
        $source | Should Match '(?s)if \(\$triplet -like ''arm64-\*''\) \{.*?--overlay-ports'
        $source | Should Match 'throw "vcpkg install failed for triplet'
    }

    It 'invalidates the hosted dependency cache when the overlays or their helper change' {
        $workflow = Get-Content -LiteralPath (Join-Path $repoRoot '.github/workflows/build-reusable.yml')
        $keys = @($workflow | Where-Object { $_ -match '^\s*key:.*-vcpkg-' })
        $keys.Count | Should Be 1
        foreach ($input in @('vcpkg.json','vcpkg-tool.json','vcpkg-install.ps1','Tools/Modules/Build/VcpkgInstallSafety.psm1','Tools/Modules/Build/MSBuildInvocation.psm1','Build/vcpkg-ports/**')) {
            $keys[0].Contains("'$input'") | Should Be $true
        }
        $terminalKeys = @($workflow | Where-Object { $_ -match '^\s*key:.*-terminal-runtime-' })
        $terminalKeys.Count | Should Be 1
        $terminalKeys[0].Contains("'Tools/Modules/Build/MSBuildInvocation.psm1'") | Should Be $true
    }
}

Describe 'Pinned public build dependency binding' {
    BeforeAll {
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Build/SanitizedEnvironment.psm1') -ErrorAction Stop
        $source = [IO.File]::ReadAllText((Join-Path $repoRoot 'build.ps1'))
        $ast = [Management.Automation.Language.Parser]::ParseInput($source,[ref]$null,[ref]$null)
        foreach ($name in @('Initialize-RSBuildVcpkgEnvironment','Invoke-MSBuild')) {
            $definition = $ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq $name},$true)
            . ([scriptblock]::Create($definition.Extent.Text))
        }
    }

    It 'binds the prepared triplet without forced imports or ambient compiler and linker flags' {
        $root = New-RSTemporaryVcpkgInstallSafetyRoot
        $tool = Join-Path $root '.build/vcpkg-tool'
        $installed = Join-Path $root '.build/vcpkg_installed/x64/x64-windows'
        foreach ($directory in @('include','lib','bin')) {
            [void][IO.Directory]::CreateDirectory((Join-Path $installed $directory))
        }
        & {
            $observed=@{BootstrapCalls=0}
            function Initialize-RSVcpkgToolContext { param($RepoRoot) ++$observed.BootstrapCalls; [pscustomobject]@{ToolRoot=$tool;ToolCommit=('a'*40);Baseline=('b'*40)} }
            $parentCL=$env:CL; $parentLINK=$env:LINK; $parentRoot=$env:VcpkgRoot; $parentImports=$env:ForceImportBeforeCppTargets
            try {
                $env:CL='/DUNREVIEWED'; $env:LINK='/FORCE:UNRESOLVED'
                $env:ForceImportBeforeCppTargets = 'unreviewed.props'
                $binding=Initialize-RSBuildVcpkgEnvironment -RepoRoot $root -Platform x64 -Configuration Release
                $binding.ToolCommit | Should Be ('a'*40)
                $binding.Baseline | Should Be ('b'*40)
                $psi=New-RSProcessStartInfo -FilePath 'fixture.exe' -WorkingDirectory $root -AdditionalEnvironment $binding.Environment
                ($psi.Environment['Path'].Split([IO.Path]::PathSeparator)[0]) | Should Be (Join-Path $installed 'bin')
                foreach ($name in @('CL','LINK','ForceImportBeforeCppTargets','ForceImportAfterCppTargets')) {
                    $psi.Environment.ContainsKey($name) | Should Be $false
                }
                $env:CL | Should Be '/DUNREVIEWED'
                $env:LINK | Should Be '/FORCE:UNRESOLVED'
                $env:ForceImportBeforeCppTargets | Should Be 'unreviewed.props'
                Remove-Item -LiteralPath (Join-Path $installed 'lib')
                $r=Invoke-RSTestBlock { Initialize-RSBuildVcpkgEnvironment -RepoRoot $root -Platform x64 -Configuration Release }
                $r.Succeeded | Should Be $false
                $r.Error.Exception.Message | Should Match 'vcpkg-install.ps1 -Platform x64'
                $observed.BootstrapCalls | Should Be 1
            } finally {
                $env:CL=$parentCL; $env:LINK=$parentLINK; $env:VcpkgRoot=$parentRoot; $env:ForceImportBeforeCppTargets=$parentImports
            }
        }
        # The SDK reader must see raw overrides before canonical composition; its
        # existing broker negative controls continue to own override admission.
        $source=[IO.File]::ReadAllText((Join-Path $repoRoot 'build.ps1'))
        $source.IndexOf('Open-CiApplicationSdkInput') | Should BeLessThan $source.IndexOf('$vcpkgBinding = Initialize-RSBuildVcpkgEnvironment')
        $source.IndexOf('$vcpkgBinding = Initialize-RSBuildVcpkgEnvironment') | Should BeLessThan $source.IndexOf('$buildIdentity = Get-RSBuildIdentityBundle')
        $source | Should Match 'Get-RSBuildIdentityBundle[\s\S]*?-AdditionalEnvironment \$script:buildVcpkgEnvironment'
    }

    It 'passes the actual canonical child environment through direct and streaming MSBuild phases under the platform lock' {
        & {
            $SolutionDir=$TestDrive; $Configuration='Release'; $Platform='x64'; $Rebuild=$false; $repairProjectName='fixture'
            $artifactOperationScope=@{kind='build';configuration='Release';platform='x64'}
            $script:buildVcpkgEnvironment=@{Path=$TestDrive;CL=$null;LINK=$null}
            $events=[Collections.Generic.List[string]]::new()
            $children=[Collections.Generic.List[object]]::new()
            function Enter-RSArtifactOperationLock {param($RepoRoot,$Operation,$Scope) $events.Add('lock-'+$Scope.coordination); @{WasAbandoned=$false}}
            function Exit-RSArtifactOperationLock {param($Lock) $events.Add('unlock')}
            function Read-RSArtifactOperationContamination {param($RepoRoot,$Scope) $null}
            function Assert-RSNoResidualArtifactToolProcesses {param($RepoRoot,$Scope)}
            function New-ProcessLogPath {param($Prefix) Join-Path $TestDrive 'not-executed.log'}
            function Get-RSMSBuildInvocationPlan {param($UseInteractiveTerminal,$LogPath) @{UseDirectConsole=$UseInteractiveTerminal;AdditionalArguments=@()}}
            function Write-RSMSBuildDiagnosticSummary {param($LogPath)}
            function Assert-NativeBuildDiagnostics {param($LogPath)}
            function Invoke-RSProcess {
                param($FilePath,$Arguments,$WorkingDirectory,$AdditionalEnvironment)
                $events.Add('direct');$children.Add((New-RSProcessStartInfo -FilePath $FilePath -Arguments $Arguments -WorkingDirectory $WorkingDirectory -AdditionalEnvironment $AdditionalEnvironment));0
            }
            function Invoke-RSStreamingProcess {
                param($FilePath,$Arguments,$WorkingDirectory,$LogPath,$OutputLineCallback,$AdditionalEnvironment)
                $events.Add('streaming');$children.Add((New-RSProcessStartInfo -FilePath $FilePath -Arguments $Arguments -WorkingDirectory $WorkingDirectory -AdditionalEnvironment $AdditionalEnvironment));0
            }
            foreach($direct in @($false,$true)){
                $script:UseInteractiveTerminal=$direct
                (Invoke-MSBuild -MSBuildPath 'not-executed.exe' -Arguments @('/t:Build')) | Should Be 0
            }
            ($events -join ',') | Should Be 'lock-shared-dependency,streaming,unlock,lock-shared-dependency,direct,unlock'
            $children.Count | Should Be 2
            foreach($child in $children){
                $child.Environment['Path'] | Should Be $TestDrive
                $child.Environment.ContainsKey('CL') | Should Be $false
                $child.Environment.ContainsKey('LINK') | Should Be $false
                @($child.ArgumentList) -contains '/nr:false' | Should Be $true
            }
        }
    }
}

Describe 'Explicit installer overlay MSBuild properties' -Tag RequiresBuildToolchain {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
    $msbuild = $null
    if (Test-Path -LiteralPath $vswhere -PathType Leaf) {
        $installation = @(Invoke-RSTestFixtureNative -FilePath $vswhere -Arguments @('-latest'; '-prerelease'; '-products'; '*'; '-requires'; 'Microsoft.Component.MSBuild'; '-property'; 'installationPath')) |
            Where-Object { $_ } | Select-Object -First 1
        if ($installation) {
            $relative = if ([Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString() -eq 'Arm64') {
                'MSBuild\Current\Bin\arm64\MSBuild.exe'
            } else { 'MSBuild\Current\Bin\amd64\MSBuild.exe' }
            $candidate = Join-Path ([string]$installation).Trim() $relative
            if (Test-Path -LiteralPath $candidate -PathType Leaf) { $msbuild = $candidate }
        }
    }

    It 'evaluates explicit MSBuild library paths with integration disabled' -Skip:($null -eq $msbuild) {
        $project = Join-Path $repoRoot 'Plugins\FileSystemCurl\FileSystemCurl.vcxproj'
        foreach ($platform in @('x64','ARM64')) {
            $scope = $platform.ToLowerInvariant()
            foreach ($configuration in @('Release','Debug','ASan Debug')) {
                $output = Invoke-RSTestFixtureNative -FilePath $msbuild -Arguments @(
                    $project; '/nologo'; '/nr:false'; "/p:Configuration=$configuration"; "/p:Platform=$platform";
                    '-getProperty:VcpkgEnabled,VcpkgEnableManifest,VcpkgManifestInstall,RSVcpkgLinkLibraries')
                $LASTEXITCODE | Should Be 0
                $metadata = ($output -join [Environment]::NewLine | ConvertFrom-Json).Properties
                foreach ($name in @('VcpkgEnabled','VcpkgEnableManifest','VcpkgManifestInstall')) {
                    $metadata.$name | Should Be 'false'
                }
                $lib = if ($configuration -eq 'Release') { 'lib\libcurl.lib' } else { 'debug\lib\libcurl-d.lib' }
                $expected = Join-Path $repoRoot ".build\vcpkg_installed\$scope\$scope-windows\$lib"
                ($metadata.RSVcpkgLinkLibraries.Split(';') -contains $expected) | Should Be $true
            }
        }
    }
}

Describe 'MSVC 14.52 ARM64 AWS hash compiler workaround' {
    It 'limits expansion only for the affected compiler target and translation unit' {
        $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
        $installation = (Invoke-RSTestFixtureNative -FilePath $vswhere -Arguments @('-latest'; '-prerelease'; '-products'; '*'; '-requires'; 'Microsoft.Component.MSBuild'; '-property'; 'installationPath') | Select-Object -First 1)
        $cmake = Join-Path $installation 'Common7/IDE/CommonExtensions/Microsoft/CMake/CMake/bin/cmake.exe'
        (Test-Path -LiteralPath $cmake) | Should Be $true
        $fixture = New-RSTemporaryVcpkgInstallSafetyRoot
        $ninja = Join-Path $installation 'Common7/IDE/CommonExtensions/Microsoft/CMake/Ninja/ninja.exe'
        (Test-Path -LiteralPath $ninja) | Should Be $true
        $helper = (Join-Path $repoRoot 'Build/vcpkg-ports/aws-c-common/msvc1452-arm64.cmake').Replace('\','/')
        $cases = @(
            @{name='affected';msvc='TRUE';version='19.52.37025.0';arch='ARM64';expected='/Ob1'},
            @{name='x64';msvc='TRUE';version='19.52.37025.0';arch='x64';expected='NOTFOUND'},
            @{name='older';msvc='TRUE';version='19.51.36231.0';arch='ARM64';expected='NOTFOUND'},
            @{name='next-patch';msvc='TRUE';version='19.52.37026.0';arch='ARM64';expected='NOTFOUND'},
            @{name='newer';msvc='TRUE';version='19.53.37025.0';arch='ARM64';expected='NOTFOUND'},
            @{name='other-compiler';msvc='FALSE';version='19.52.37025.0';arch='ARM64';expected='NOTFOUND'}
        )
        foreach ($case in $cases) {
            $source = Join-Path $fixture $case.name
            [void][IO.Directory]::CreateDirectory((Join-Path $source 'source'))
            [IO.File]::WriteAllText((Join-Path $source 'source/hash_table.c'),'int fixture(void) { return 0; }')
            [IO.File]::WriteAllText((Join-Path $source 'source/other.c'),'int other(void) { return 0; }')
            $project = @"
cmake_minimum_required(VERSION 3.20)
project(CompilerScope NONE)
set(MSVC $($case.msvc))
set(CMAKE_C_COMPILER_VERSION $($case.version))
set(CMAKE_C_COMPILER_ARCHITECTURE_ID $($case.arch))
include("$helper")
get_source_file_property(affected "`${CMAKE_CURRENT_SOURCE_DIR}/source/hash_table.c" COMPILE_OPTIONS)
get_source_file_property(other "`${CMAKE_CURRENT_SOURCE_DIR}/source/other.c" COMPILE_OPTIONS)
file(WRITE "`${CMAKE_BINARY_DIR}/result.txt" "`${affected}\n`${other}\n")
"@
            [IO.File]::WriteAllText((Join-Path $source 'CMakeLists.txt'),$project)
            $build = Join-Path $source 'build'
            Invoke-RSTestFixtureNative -FilePath $cmake -Arguments @('-G'; 'Ninja'; ('-DCMAKE_MAKE_PROGRAM=' + $ninja.Replace('\','/')); '-S'; $source; '-B'; $build) *> (Join-Path $source 'configure.log')
            $LASTEXITCODE | Should Be 0
            $actual = @(Get-Content -LiteralPath (Join-Path $build 'result.txt'))
            $actual[0] | Should Be $case.expected
            $actual[1] | Should Be 'NOTFOUND'
        }
    }
}

Describe 'Captured vcpkg provisioning Git authority' {
    It 'keeps provisioning writes on delegated streaming with original merged diagnostics and exact failure' {
        $root=New-RSTemporaryVcpkgInstallSafetyRoot
        $bootstrapFixtureState=[pscustomobject]@{Exit=0;Calls=[Collections.Generic.List[object]]::new()}
        Mock Invoke-RSCapturedProcess -ModuleName VcpkgInstallSafety {throw 'Provisioning writes must not use read-only capture.'}
        Mock Invoke-RSStreamingProcess -ModuleName VcpkgInstallSafety ({
            param($FilePath,$Arguments,$WorkingDirectory,$OutputLineCallback,[switch]$DelegateArtifactOperation,[Text.Encoding]$StandardOutputEncoding)
            if ($FilePath -cne 'git' -or $WorkingDirectory -cne $root -or -not $DelegateArtifactOperation -or $StandardOutputEncoding.CodePage -ne 65001) { throw 'Bootstrap capture lost literal Git, CWD or delegated write authority.' }
            $bootstrapFixtureState.Calls.Add(@{argv=@($Arguments);delegated=[bool]$DelegateArtifactOperation})
            & $OutputLineCallback 'stdout é雪' $false
            & $OutputLineCallback 'stderr diagnostic' $true
            return $bootstrapFixtureState.Exit
        }.GetNewClosure())
        $module=Get-Module VcpkgInstallSafety
        $actual=& $module {param($Root)Invoke-RSVcpkgBootstrapGit $Root @('clone','--','literal * repository','owned checkout')} $root
        $actual|Should Be ('stdout é雪'+[Environment]::NewLine+'stderr diagnostic')
        ($bootstrapFixtureState.Calls[0].argv -join '|')|Should Be ('-C|'+$root+'|clone|--|literal * repository|owned checkout')
        $bootstrapFixtureState.Exit=17
        {& $module {param($Root)Invoke-RSVcpkgBootstrapGit $Root @('fetch','origin',('a'*40))} $root}|Should Throw 'Pinned vcpkg Git operation failed'
        ($bootstrapFixtureState.Calls[1].argv -join '|')|Should Be ('-C|'+$root+'|fetch|origin|'+('a'*40))
        $bootstrapFixtureState.Calls.Count|Should Be 2
    }
    It 'uses the actual cat-file exit to fetch only missing pinned objects without swallowing the recheck' {
        $t=$null;$e=$null;$ast=[Management.Automation.Language.Parser]::ParseFile($helperModule,[ref]$t,[ref]$e)
        $function=$ast.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -ceq 'Initialize-RSVcpkgToolContext'},$true)
        $loop=$function.Find({param($n)$n -is [Management.Automation.Language.ForEachStatementAst] -and $n.Variable.Extent.Text -ceq '$required'},$true)
        if(!$loop -or !$loop.Extent.Text.Contains('$commitProbe.ExitCode -ne 0')){throw 'Actual pinned-object probe boundary missing.'}
        $source=$loop.Extent.Text
        foreach($probeExitValue in @(0,17)){
            $probeWriteCalls=[Collections.Generic.List[object]]::new();$probeReadCalls=[Collections.Generic.List[object]]::new()
            & {
                $toolRoot=$TestDrive;$commit='a'*40;$baseline='b'*40
                function Invoke-RSCapturedProcess {param($FilePath,$WorkingDirectory,$Arguments,[Text.Encoding]$StandardOutputEncoding)
                    if ($FilePath -cne 'git' -or $WorkingDirectory -cne $toolRoot -or $StandardOutputEncoding.CodePage -ne 65001) { throw 'Object probe lost literal Git or CWD.' }
                    $probeReadCalls.Add(@($Arguments));return [pscustomobject]@{ExitCode=$probeExitValue;StdOut=@();StdErr=@('probe diagnostic must remain suppressed')}
                }
                function Invoke-RSVcpkgBootstrapGit {param($Root,$Arguments)$Root|Should Be $toolRoot;$probeWriteCalls.Add(@($Arguments));return ''}
                . ([scriptblock]::Create($source))
            }
            $probeReadCalls.Count|Should Be 2
            ($probeReadCalls[0] -join '|')|Should Be ('-C|'+$TestDrive+'|cat-file|-e|'+('a'*40)+'^{commit}')
            ($probeReadCalls[1] -join '|')|Should Be ('-C|'+$TestDrive+'|cat-file|-e|'+('b'*40)+'^{commit}')
            $probeWriteCalls.Count|Should Be $(if($probeExitValue){4}else{0})
            if($probeExitValue){($probeWriteCalls[0] -join '|')|Should Be ('fetch|origin|'+('a'*40));($probeWriteCalls[1] -join '|')|Should Be ('cat-file|-e|'+('a'*40)+'^{commit}')}
        }
    }
}
