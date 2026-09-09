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
        $dependency[0].'version>=' | Should Be '1.11.840'
        $override.Count | Should Be 1
        $override[0].version | Should Be '1.11.840'
    }

    It 'pairs blake3 and ftxui floors with the registry commit that first published both' {
        $manifest = Get-Content -LiteralPath (Join-Path $repoRoot 'vcpkg.json') -Raw | ConvertFrom-Json
        $blake3 = @($manifest.dependencies | Where-Object { $_.name -eq 'blake3' })
        $ftxui = @($manifest.dependencies | Where-Object { $_.name -eq 'ftxui' })

        $manifest.'builtin-baseline' | Should Be '11159b9b7eed119f182612c9acdd8c864a21e9c6'
        $blake3.Count | Should Be 1
        $blake3[0].'version>=' | Should Be '1.8.7'
        $ftxui.Count | Should Be 1
        $ftxui[0].'version>=' | Should Be '7.0.3'
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
