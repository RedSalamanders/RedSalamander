Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$testRunPlanModule = Join-Path $repoRoot 'Tools\Modules\Testing\TestRunPlan.psm1'
$helperModule = Join-Path $repoRoot 'Tools\Modules\Build\VcpkgInstallSafety.psm1'
Import-Module $testRunPlanModule -Force -ErrorAction Stop
Import-Module $helperModule -Force -ErrorAction Stop

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
        $source | Should Match "Import-Module .*SanitizedEnvironment\.psm1"
        $source | Should Match "coordination = 'shared-dependency'"
        $source | Should Match 'Sort-Object \{ if \(\$_ -eq ''x64''\) \{ 0 \} else \{ 1 \} \} -Unique'
        $source | Should Match 'Invoke-RSProcess\s+`\r?\n\s+-FilePath \$vcpkgExePath'
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

            { Assert-RSVcpkgManifestVersionConstraintsAvailable -ManifestPath $manifestPath -RegistryRoot $registryRoot } |
                Should Not Throw
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

            { Assert-RSVcpkgManifestVersionConstraintsAvailable -ManifestPath $manifestPath -RegistryRoot $registryRoot } |
                Should Not Throw
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
    $dxuiTools = Join-Path $repoRoot ".build\dependencies\DxUi\source\$pin\Tools"
    $dxuiModulesRestored = (Test-Path -LiteralPath (Join-Path $dxuiTools 'VisualStudio.psm1')) -and (Test-Path -LiteralPath (Join-Path $dxuiTools 'VcpkgTriplet.psm1'))

    It 'matches the pinned DxUi modules it mirrors' -Skip:(-not $dxuiModulesRestored) {
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

    $toolCheckout = Join-Path $repoRoot '.build\vcpkg-tool'
    $toolCheckoutAtPin = $false
    if (Test-Path -LiteralPath (Join-Path $toolCheckout 'versions\baseline.json') -PathType Leaf) {
        $head = & git -C $toolCheckout rev-parse HEAD 2>$null
        $toolCheckoutAtPin = ($LASTEXITCODE -eq 0) -and ([string]$head).Trim() -eq [string]$toolPin.commit
    }

    It 'finds every floor and override in the pinned tool checkout''s version database' -Skip:(-not $toolCheckoutAtPin) {
        # Capture the error text: under PowerShell 7, Pester 3.4's 'Should Not Throw' passes even when the block throws.
        $errorMessage = $null
        try {
            Assert-RSVcpkgManifestVersionConstraintsAvailable -ManifestPath (Join-Path $repoRoot 'vcpkg.json') -RegistryRoot $toolCheckout
        } catch {
            $errorMessage = $_.Exception.Message
        }
        $errorMessage | Should BeNullOrEmpty
    }
}

Describe 'MSBuild manifest install toolset pin' {
    It 'refreshes the installer''s overlay triplet before MSBuild can install, inside the platform dependency lock' {
        $source = Get-Content -LiteralPath (Join-Path $repoRoot 'build.ps1') -Raw

        $invoke = $source.IndexOf('function Invoke-MSBuild {')
        $invoke | Should BeGreaterThan 0
        $lock = $source.IndexOf('$sharedDependencyLock = Enter-RSArtifactOperationLock', $invoke)
        $lock | Should BeGreaterThan $invoke
        $overlay = $source.IndexOf('Update-RSBuildVcpkgOverlayTriplet', $lock)
        $overlay | Should BeGreaterThan $lock
        $overlay | Should BeLessThan $source.IndexOf('$logPath = New-ProcessLogPath', $invoke)
        $source | Should Match '\$visualStudioInstallation = Get-RSVcpkgVisualStudioInstallation'
        $source | Should Match 'Get-RSVcpkgDefaultToolset -Installation \$visualStudioInstallation'
        $source | Should Match 'Update-RSVcpkgOverlayTriplet'
        $source | Should Match 'Resolve-RSVcpkgSafeChildPath -Root \$SolutionDir -Child "\.build\\vcpkg-triplets\\'
    }

    It 'initializes the once-per-build overlay flag so callers under StrictMode can build' {
        # Run-AllTests and Test-Changes invoke build.ps1 under Set-StrictMode -Version Latest; reading the flag before any
        # assignment then throws and every governed build fails. Run the real function text in its own script scope.
        $source = Get-Content -LiteralPath (Join-Path $repoRoot 'build.ps1') -Raw
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($source, [ref]$null, [ref]$null)
        $function = $ast.Find({
                param($node)
                $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Update-RSBuildVcpkgOverlayTriplet'
            }, $false)
        $function | Should Not BeNullOrEmpty
        # Script-level (possibly inside build.ps1's top-level try block), never inside a function body.
        $initialization = @($ast.FindAll({
                    param($node)
                    if ($node -isnot [System.Management.Automation.Language.AssignmentStatementAst] -or
                        $node.Extent.Text -notmatch '^\$script:vcpkgOverlayTripletRefreshed\s*=\s*\$false$') {
                        return $false
                    }
                    for ($parent = $node.Parent; $null -ne $parent; $parent = $parent.Parent) {
                        if ($parent -is [System.Management.Automation.Language.FunctionDefinitionAst]) { return $false }
                    }
                    return $true
                }, $true))
        $initialization.Count | Should Be 1
        $initialization[0].Extent.StartOffset | Should BeLessThan $function.Extent.StartOffset

        $root = New-RSTemporaryVcpkgInstallSafetyRoot
        try {
            $probe = Join-Path $root 'strict-overlay-flag-probe.ps1'
            Set-Content -LiteralPath $probe -Encoding utf8 -Value @(
                'Set-StrictMode -Version Latest'
                '$ErrorActionPreference = ''Stop'''
                $initialization[0].Extent.Text
                $function.Extent.Text
                '$env:VcpkgManifestInstall = ''false'''
                'Update-RSBuildVcpkgOverlayTriplet'
                'Update-RSBuildVcpkgOverlayTriplet'
                'if (-not $script:vcpkgOverlayTripletRefreshed) { throw ''overlay flag was not recorded'' }'
            )
            $output = @(& (Join-Path $PSHOME 'pwsh.exe') -NoProfile -NonInteractive -File $probe 2>&1)
            $exitCode = $LASTEXITCODE
            ($output -join [Environment]::NewLine) | Should Not Match 'cannot be retrieved because it has not been set'
            $exitCode | Should Be 0
        } finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
    $msbuild = $null
    if (Test-Path -LiteralPath $vswhere -PathType Leaf) {
        $installation = @(& $vswhere -latest -prerelease -products '*' -requires Microsoft.Component.MSBuild -property installationPath) |
            Where-Object { $_ } | Select-Object -First 1
        if ($installation) {
            $relative = if ([Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString() -eq 'Arm64') {
                'MSBuild\Current\Bin\arm64\MSBuild.exe'
            } else { 'MSBuild\Current\Bin\amd64\MSBuild.exe' }
            $candidate = Join-Path ([string]$installation).Trim() $relative
            if (Test-Path -LiteralPath $candidate -PathType Leaf) { $msbuild = $candidate }
        }
    }

    It 'passes the overlay to the MSBuild manifest install only when the overlay triplet exists' -Skip:($null -eq $msbuild) {
        $root = New-RSTemporaryVcpkgInstallSafetyRoot
        try {
            $buildRoot = (Join-Path $root 'build') + '\'
            $project = Join-Path $repoRoot 'Common\Common\Common.vcxproj'
            $evaluate = {
                param([string]$Platform)
                $output = & $msbuild $project /nologo '/p:Configuration=Debug' "/p:Platform=$Platform" "/p:RSBuildRoot=$buildRoot" '-getProperty:VcpkgAdditionalInstallOptions'
                if ($LASTEXITCODE -ne 0) { throw "Cannot evaluate $project for $Platform." }
                return (($output -join "`n").Trim())
            }

            # No overlay yet (neither vcpkg-install.ps1 nor build.ps1 ran): vcpkg must not get a directory that does not exist.
            (& $evaluate 'x64') | Should Not Match 'overlay-triplets'

            foreach ($platform in @(@{ Name = 'x64'; Scope = 'x64'; Triplet = 'x64-windows' }, @{ Name = 'ARM64'; Scope = 'arm64'; Triplet = 'arm64-windows' })) {
                $overlayDirectory = Join-Path $buildRoot "vcpkg-triplets\$($platform.Scope)"
                New-Item -ItemType Directory -Path $overlayDirectory -Force | Out-Null
                [System.IO.File]::WriteAllText((Join-Path $overlayDirectory "$($platform.Triplet).cmake"), "set(VCPKG_TARGET_ARCHITECTURE x)`n")
                # Quoted, and without a trailing backslash that would escape the closing quote on the vcpkg command line.
                (& $evaluate $platform.Name) | Should Match ('"--overlay-triplets=' + [regex]::Escape($overlayDirectory) + '"')
            }
        } finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

Describe 'MSVC 14.52 ARM64 AWS hash compiler workaround' {
    It 'limits expansion only for the affected compiler target and translation unit' {
        $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
        $installation = (& $vswhere -latest -prerelease -products '*' -requires Microsoft.Component.MSBuild -property installationPath | Select-Object -First 1)
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
            & $cmake -G Ninja ('-DCMAKE_MAKE_PROGRAM=' + $ninja.Replace('\','/')) -S $source -B $build *> (Join-Path $source 'configure.log')
            $LASTEXITCODE | Should Be 0
            $actual = @(Get-Content -LiteralPath (Join-Path $build 'result.txt'))
            $actual[0] | Should Be $case.expected
            $actual[1] | Should Be 'NOTFOUND'
        }
    }
}
