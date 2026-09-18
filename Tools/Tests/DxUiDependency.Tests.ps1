$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'Tools/Modules/Build/DxUiDependency.psm1') -Force

Describe 'Canonical DxUi source ownership' {
    It 'keeps the shared implementation and control suites out of the consumer tracked tree' {
        foreach ($relative in @('Common/DxUi', 'Tests/DxUiTests')) {
            $tracked = @(& git -C $repoRoot ls-files -- $relative)
            if ($LASTEXITCODE -ne 0) { throw "Cannot enumerate tracked legacy path: $relative" }
            $tracked.Count | Should Be 0
        }
        $solution = Get-Content -LiteralPath (Join-Path $repoRoot 'RedSalamander.sln') -Raw
        $solution | Should Not Match 'Common[\\/]DxUi|Tests[\\/]DxUiTests'
        $solution | Should Match 'ProductUiTests'
    }

    It 'keeps owned project inputs separate from the restored library implementation' {
        $projects = @(& git -C $repoRoot ls-files -- '*.vcxproj')
        if ($LASTEXITCODE -ne 0) { throw 'Cannot enumerate owned projects.' }
        $projects.Count | Should BeGreaterThan 0
        foreach ($project in $projects) {
            [xml]$xml = Get-Content -LiteralPath (Join-Path $repoRoot $project) -Raw
            foreach ($inputNode in $xml.SelectNodes('//*[local-name()="ClCompile" or local-name()="ProjectReference"]/@Include')) {
                $inputNode.Value | Should Not Match 'Common[\\/]DxUi|Tests[\\/]DxUiTests|DxUi[\\/]src[\\/].*\.cpp'
            }
        }
    }
}

Describe 'Exact DxUi consumer dependency' {
    It 'renders a pending DxUi validation in red and validated upgrade commands in yellow' {
        $module = Get-Module DxUiDependency
        $unvalidated = & $module {
            Get-RSDxUiUpdateNoticePresentation -Notice ('DxUi main ' + ('b' * 12) + ' has no successful completed validation; keep pinned ' + ('a' * 12) + '.')
        }
        $available = & $module {
            Get-RSDxUiUpdateNoticePresentation -Notice ('DxUi update available: pinned ' + ('a' * 12) + ', available ' + ('b' * 12) +
                '. https://github.com/RedSalamanders/DxUi/compare/test . Update Z:\fixture\Dependencies\DxUi.lock.json on a branch and run the product regressions.')
        }
        $other = & $module {
            Get-RSDxUiUpdateNoticePresentation -Notice 'DxUi update check unavailable; the exact pinned dependency remains selected.'
        }

        $unvalidated.ForegroundColor.ToString() | Should Be 'Red'
        $available.ForegroundColor.ToString() | Should Be 'Yellow'
        $available.Notice | Should Match ([regex]::Escape("`nTo upgrade and run local validation: .\Tools\Update-DxUi.ps1"))
        $available.Notice | Should Match ([regex]::Escape('To upgrade only after equivalent product validation passed elsewhere: .\Tools\Update-DxUi.ps1 -UpdateOnly'))
        $other.ForegroundColor.ToString() | Should Be 'DarkYellow'
    }

    It 'rejects malformed locks before source lookup or restore' {
        $root = Join-Path $TestDrive 'invalid-pin'
        [void](New-Item -ItemType Directory -Path (Join-Path $root 'Dependencies') -Force)
        foreach ($pin in @(
            @{repository='https://example.invalid/fork';commit=('a'*40);apiRevision=2;targets=@('DxUi')},
            @{repository='https://github.com/RedSalamanders/DxUi';commit='../escape';apiRevision=2;targets=@('DxUi')},
            @{repository='https://github.com/RedSalamanders/DxUi';commit=('a'*40);apiRevision=1;targets=@('DxUi')},
            @{repository='https://github.com/RedSalamanders/DxUi';commit=('a'*40);apiRevision=2;targets=@('Foundation')}
        )) {
            $pin | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $root 'Dependencies/DxUi.lock.json')
            $rejected=$false; try { Get-RSDxUiPin -RepoRoot $root | Out-Null } catch { $rejected=$true }; $rejected | Should Be $true
        }
    }

    It 'requires source for artifact use and rejects tracked and untracked source changes' {
        $root = Join-Path $TestDrive 'source-identity'
        $origin = Join-Path $root 'fixture-origin'
        [void](New-Item -ItemType Directory -Path (Join-Path $origin 'include') -Force)
        [void](New-Item -ItemType Directory -Path (Join-Path $root 'Dependencies') -Force)
        'synthetic public header' | Set-Content -LiteralPath (Join-Path $origin 'include/Example.h')
        & git init --quiet $origin
        & git -C $origin add .
        & git -C $origin -c user.name=Fixture -c user.email=fixture@example.invalid commit --quiet -m fixture
        if ($LASTEXITCODE -ne 0) { throw 'Cannot initialize isolated source fixture.' }
        $head = (& git -C $origin rev-parse HEAD).Trim()
        @{repository='https://github.com/RedSalamanders/DxUi';commit=$head;apiRevision=2;targets=@('DxUi')} |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path $root 'Dependencies/DxUi.lock.json')
        $rejected=$false; try { Get-RSDxUiSourceIdentity -RepoRoot $root | Out-Null } catch { $rejected=$true }; $rejected | Should Be $true
        Get-RSDxUiSourceIdentity -RepoRoot $root -AllowMissing | Should BeNullOrEmpty
        $source = Join-Path $root ".build/dependencies/DxUi/source/$head"
        [void](New-Item -ItemType Directory -Path (Split-Path $source) -Force)
        & git -c core.longpaths=true clone --config core.longpaths=true --quiet --local $origin $source
        if ($LASTEXITCODE -ne 0) { throw 'Cannot clone isolated source fixture.' }
        (Get-RSDxUiSourceIdentity -RepoRoot $root).Pin.commit | Should Be $head
        $header = Join-Path $source 'include/Example.h'
        $original = [IO.File]::ReadAllBytes($header)
        Add-Content -LiteralPath $header -Value 'edited source'
        $rejected=$false; try { Get-RSDxUiSourceIdentity -RepoRoot $root | Out-Null } catch { $rejected=$true }; $rejected | Should Be $true
        [IO.File]::WriteAllBytes($header,$original)
        'untracked input' | Set-Content -LiteralPath (Join-Path $source 'extra.cpp')
        $rejected=$false; try { Get-RSDxUiSourceIdentity -RepoRoot $root | Out-Null } catch { $rejected=$true }; $rejected | Should Be $true
        Remove-Item -LiteralPath (Join-Path $source 'extra.cpp')
        (Get-RSDxUiSourceIdentity -RepoRoot $root).PublicHeadersGitTree | Should Match '^[0-9a-f]{40}$'
    }
    It 'binds module provenance to the exact archive and rejects stale or ambiguous inputs' {
        $root = Join-Path $TestDrive 'provenance'
        $origin = Join-Path $root 'fixture-origin'
        [void](New-Item -ItemType Directory -Path (Join-Path $origin 'include') -Force)
        [void](New-Item -ItemType Directory -Path (Join-Path $root 'Dependencies') -Force)
        'synthetic public header' | Set-Content -LiteralPath (Join-Path $origin 'include/Example.h')
        & git init --quiet $origin
        & git -C $origin add .
        & git -C $origin -c user.name=Fixture -c user.email=fixture@example.invalid commit --quiet -m fixture
        if ($LASTEXITCODE -ne 0) { throw 'Cannot initialize isolated provenance fixture.' }
        $head = (& git -C $origin rev-parse HEAD).Trim()
        @{repository='https://github.com/RedSalamanders/DxUi';commit=$head;apiRevision=2;targets=@('DxUi')} |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path $root 'Dependencies/DxUi.lock.json')
        $dependency = Join-Path $root '.build/dependencies/DxUi'
        $source = Join-Path $dependency "source/$head"
        [void](New-Item -ItemType Directory -Path (Split-Path $source) -Force)
        & git -c core.longpaths=true clone --config core.longpaths=true --quiet --local $origin $source
        if ($LASTEXITCODE -ne 0) { throw 'Cannot clone isolated provenance fixture.' }
        $identity = @{Fingerprint=('b'*64);Identity=@{commit=$head;platform='x64';disableStlAnnotations=$true}}
        $identityPath = Join-Path $dependency 'DxUi.identity.x64.json'
        $identity | ConvertTo-Json | Set-Content -LiteralPath $identityPath
        $archive = Join-Path $dependency "$($identity.Fingerprint)/x64/Debug/DxUi.lib"
        [void](New-Item -ItemType Directory -Path (Split-Path $archive) -Force)
        'synthetic archive bytes' | Set-Content -LiteralPath $archive
        $output = Join-Path $root '.build/x64/Debug'
        [void](New-Item -ItemType Directory -Path $output -Force)
        $module = Join-Path $output 'Pilot.exe'
        'synthetic PE bytes' | Set-Content -LiteralPath $module
        $tlog = Join-Path $output 'link.command.1.tlog'
        $link = 'link.exe "' + $archive + '"'
        $link | Set-Content -LiteralPath $tlog
        "target=$module`ndxui_link_tlog=$tlog" | Set-Content -LiteralPath (Join-Path $output 'Pilot.manifest')
        $arguments = @{RepoRoot=$root;ArtifactManifestDir=$output;Platform='x64';Configuration='Debug'}
        Write-RSDxUiModuleProvenance @arguments
        $record = Get-Content -LiteralPath "$module.DxUi.json" -Raw | ConvertFrom-Json
        $record.commit | Should Be $head
        $record.archiveSha256 | Should Be (Get-FileHash -LiteralPath $archive).Hash
        $record.moduleSha256 | Should Be (Get-FileHash -LiteralPath $module).Hash
        $record.publicHeadersGitTree | Should Be (Get-RSDxUiSourceIdentity -RepoRoot $root).PublicHeadersGitTree
        foreach ($invalidLink in @('link.exe "old/DxUi.lib"', ($link + ' "DxUi.lib" '), ($link + ' "Common/DxUi/legacy.obj"'))) {
            $invalidLink | Set-Content -LiteralPath $tlog
            $rejected=$false; try { Write-RSDxUiModuleProvenance @arguments } catch { $rejected=$true }; $rejected | Should Be $true
        }
        $link | Set-Content -LiteralPath $tlog
        foreach ($invalidIdentity in @(
            @{Fingerprint=$identity.Fingerprint;Identity=@{commit=('c'*40);platform='x64';disableStlAnnotations=$true}},
            @{Fingerprint=$identity.Fingerprint;Identity=@{commit=$head;platform='ARM64';disableStlAnnotations=$true}},
            @{Fingerprint=$identity.Fingerprint;Identity=@{commit=$head;platform='x64';disableStlAnnotations=$false}}
        )) {
            $invalidIdentity | ConvertTo-Json | Set-Content -LiteralPath $identityPath
            $rejected=$false; try { Write-RSDxUiModuleProvenance @arguments } catch { $rejected=$true }; $rejected | Should Be $true
        }
        $identity | ConvertTo-Json | Set-Content -LiteralPath $identityPath
        Write-RSDxUiModuleProvenance @arguments
    }
}


Describe 'DxUi consumer profile evaluation' {
    It 'uses the first-party host during restore and restores the caller environment after success or failure' {
        $root = Join-Path $TestDrive 'host-restore'
        # Full uses a deeper sandbox than focused Pester runs. Exercise that path length in both entrypoints.
        if ($root.Length -lt 160) { $root = Join-Path $root ('x' * (160 - $root.Length)) }
        $origin = Join-Path $root 'fixture-origin'
        foreach ($relative in @('RedSalamander','Dependencies','fixture-origin/include','fixture-origin/Tools')) {
            [void](New-Item -ItemType Directory -Path (Join-Path $root $relative) -Force)
        }
        # A real MSBuild evaluation with profile-dependent first-party policy; no compilation or SDK fixture is needed.
        @'
<Project><PropertyGroup>
  <PreferredToolArchitecture Condition="'$(PreferredToolArchitecture)'=='' and '$(Configuration)'=='Release'">x86</PreferredToolArchitecture>
  <PreferredToolArchitecture Condition="'$(PreferredToolArchitecture)'==''">x64</PreferredToolArchitecture>
</PropertyGroup></Project>
'@ | Set-Content -LiteralPath (Join-Path $root 'RedSalamander/RedSalamander.vcxproj')
        'synthetic public header' | Set-Content -LiteralPath (Join-Path $origin 'include/Example.h')
        # The pinned helper's environment boundary is the dependency under test; do not download or compile DxUi here.
        @'
function Get-DxUiConsumerBuildIdentity {
    param($DxUiRoot,$MSBuildPath,$Platform,[switch]$DisableStlAnnotations)
    if (Test-Path -LiteralPath (Join-Path (Split-Path (Split-Path $DxUiRoot)) 'fail-identity')) { throw 'Injected identity failure' }
    $selected = if ($env:PreferredToolArchitecture) { $env:PreferredToolArchitecture } else { 'x86' }
    return [pscustomobject]@{Fingerprint=('a'*64);Identity=@{
        toolset='v145';vcToolsVersion='fixture';windowsSdkVersion='fixture';preferredToolArchitecture=$selected
    }}
}
Export-ModuleMember -Function Get-DxUiConsumerBuildIdentity
'@ | Set-Content -LiteralPath (Join-Path $origin 'Tools/ConsumerBuild.psm1')
        & git init --quiet $origin
        & git -C $origin add .
        & git -C $origin -c user.name=Fixture -c user.email=fixture@example.invalid commit --quiet -m fixture
        if ($LASTEXITCODE -ne 0) { throw 'Cannot initialize compiler-host fixture.' }
        $head = (& git -C $origin rev-parse HEAD).Trim()
        @{repository='https://github.com/RedSalamanders/DxUi';commit=$head;apiRevision=2;targets=@('DxUi')} |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path $root 'Dependencies/DxUi.lock.json')
        $dependency = Join-Path $root '.build/dependencies/DxUi'
        $source = Join-Path $dependency "source/$head"
        [void](New-Item -ItemType Directory -Path (Split-Path $source) -Force)
        (Join-Path $source '.git/hooks/fsmonitor-watchman.sample').Length | Should BeGreaterThan 260
        & git -c core.longpaths=true clone --config core.longpaths=true --quiet --local $origin $source
        if ($LASTEXITCODE -ne 0) { throw 'Cannot clone compiler-host fixture.' }
        Mock Invoke-RSDxUiRestoreProcess {} -ModuleName DxUiDependency
        $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
        $installation = (& $vswhere -latest -prerelease -products '*' -requires Microsoft.Component.MSBuild -property installationPath).Trim()
        $msbuild = Join-Path $installation 'MSBuild/Current/Bin/MSBuild.exe'
        $previousHost = $env:PreferredToolArchitecture
        try {
            foreach ($requestedHost in @('', 'x86')) {
                foreach ($platform in @('x64','ARM64')) {
                    foreach ($configuration in @('Debug','Release','ASan Debug')) {
                        $env:PreferredToolArchitecture = $requestedHost
                        Restore-RSDxUiDependency -RepoRoot $root -MSBuildPath $msbuild -Platform $platform -Configuration $configuration
                        $resolved = Get-Content -LiteralPath (Join-Path $dependency "DxUi.identity.$platform.json") -Raw | ConvertFrom-Json
                        $expected = if ($requestedHost -or $configuration -eq 'Release') { 'x86' } else { 'x64' }
                        $resolved.Identity.preferredToolArchitecture | Should Be $expected
                        [string]$env:PreferredToolArchitecture | Should Be $requestedHost
                    }
                }
            }
            'injected failure outside the pinned checkout' | Set-Content -LiteralPath (Join-Path $dependency 'fail-identity')
            $env:PreferredToolArchitecture = $null
            { Restore-RSDxUiDependency -RepoRoot $root -MSBuildPath $msbuild -Platform ARM64 -Configuration Debug } | Should Throw 'Injected identity failure'
            [string]$env:PreferredToolArchitecture | Should Be ''
        } finally { $env:PreferredToolArchitecture = $previousHost }
    }

    It 'evaluates the six real consumer profiles with sanitizer settings and a 64-bit cross-compiler host' {
        $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
        $installation = (& $vswhere -latest -prerelease -products '*' -requires Microsoft.Component.MSBuild -property installationPath).Trim()
        $relative = if ([Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString() -eq 'Arm64') {
            'MSBuild/Current/Bin/MSBuild.exe'
        } else { 'MSBuild/Current/Bin/amd64/MSBuild.exe' }
        $msbuild = Join-Path $installation $relative
        $dxuiFixture = Join-Path $TestDrive 'reference-host'
        [void](New-Item -ItemType Directory -Path (Join-Path $dxuiFixture 'Build') -Force)
        '<Project><PropertyGroup><DxUiResolvedRoot>$(DxUiRoot)\</DxUiResolvedRoot></PropertyGroup></Project>' |
            Set-Content -LiteralPath (Join-Path $dxuiFixture 'Build/DxUi.Consumer.props')
        '<Project><ItemGroup><ProjectReference Include="$(DxUiResolvedRoot)src\DxUi.vcxproj"><AdditionalProperties>DxUiOutputRoot=fixture;Platform=$(Platform);Configuration=$(Configuration)</AdditionalProperties></ProjectReference></ItemGroup></Project>' |
            Set-Content -LiteralPath (Join-Path $dxuiFixture 'Build/DxUi.Consumer.targets')
        $previousHost = $env:PreferredToolArchitecture
        try {
            $env:PreferredToolArchitecture = $null
            foreach ($platform in @('x64','ARM64')) {
                foreach ($configuration in @('Debug','Release','ASan Debug')) {
                    $json = & $msbuild (Join-Path $repoRoot 'Tests/RedConfigureTests/RedConfigureTests.vcxproj') /nologo "/p:Configuration=$configuration" "/p:Platform=$platform" "/p:DxUiRoot=$dxuiFixture" '-getProperty:EnableASAN,PreferredToolArchitecture' '-getItem:ClCompile,ProjectReference'
                    if ($LASTEXITCODE -ne 0) { throw "Cannot evaluate consumer profile $platform/$configuration." }
                    $evaluated = ($json -join "`n") | ConvertFrom-Json
                    $hostArchitecture = [Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString()
                    if ($hostArchitecture -eq 'X64') {
                        $evaluated.Properties.PreferredToolArchitecture | Should Be 'x64'
                    } elseif ($hostArchitecture -eq 'Arm64' -and $platform -eq 'ARM64') {
                        # Never the 32-bit x86-hosted cross tools: their heap cannot optimize the
                        # largest Release self-test translation units (C1002 on the hosted runners).
                        $evaluated.Properties.PreferredToolArchitecture | Should Be 'arm64'
                    }
                    $restoreHost = & (Get-Module DxUiDependency) {
                        param($Root,$BuildTool,$Target,$Profile)
                        Get-RSDxUiConsumerCompilerHost -RepoRoot $Root -MSBuildPath $BuildTool -Platform $Target -Configuration $Profile
                    } $repoRoot $msbuild $platform $configuration
                    $restoreHost | Should Be $evaluated.Properties.PreferredToolArchitecture
                    $dxuiReference = @($evaluated.Items.ProjectReference | Where-Object { [IO.Path]::GetFileName($_.Identity) -eq 'DxUi.vcxproj' })
                    $dxuiReference.Count | Should Be 1
                    $dxuiReference[0].AdditionalProperties | Should Match (';PreferredToolArchitecture=' + $evaluated.Properties.PreferredToolArchitecture + '(;|$)')
                    @($evaluated.Items.ClCompile).Count | Should BeGreaterThan 0
                    foreach ($item in $evaluated.Items.ClCompile) {
                        if ($configuration -eq 'ASan Debug') {
                            $evaluated.Properties.EnableASAN | Should Be 'true'
                            $item.ForcedIncludeFiles | Should Match 'RequireAddressSanitizer\.h'
                            $item.RuntimeLibrary | Should Be 'MultiThreadedDebugDLL'
                            $item.Optimization | Should Be 'Disabled'
                            $item.BasicRuntimeChecks | Should Be 'Default'
                        } else {
                            $item.ForcedIncludeFiles | Should Not Match 'RequireAddressSanitizer\.h'
                        }
                    }
                    $explicitHost = & $msbuild (Join-Path $repoRoot 'Tests/RedConfigureTests/RedConfigureTests.vcxproj') /nologo "/p:Configuration=$configuration" "/p:Platform=$platform" "/p:DxUiRoot=$dxuiFixture" '/p:PreferredToolArchitecture=x86' '-getProperty:PreferredToolArchitecture' '-getItem:ProjectReference'
                    if ($LASTEXITCODE -ne 0) { throw "Cannot evaluate explicit compiler host for $platform/$configuration." }
                    $explicitEvaluation = ($explicitHost -join "`n") | ConvertFrom-Json
                    $explicitEvaluation.Properties.PreferredToolArchitecture | Should Be 'x86'
                    $explicitReference = @($explicitEvaluation.Items.ProjectReference | Where-Object { [IO.Path]::GetFileName($_.Identity) -eq 'DxUi.vcxproj' })
                    $explicitReference.Count | Should Be 1
                    $explicitReference[0].AdditionalProperties | Should Match ';PreferredToolArchitecture=x86(;|$)'
                    try {
                        $env:PreferredToolArchitecture = 'x86'
                        $explicitRestoreHost = & (Get-Module DxUiDependency) {
                            param($Root,$BuildTool,$Target,$Profile)
                            Get-RSDxUiConsumerCompilerHost -RepoRoot $Root -MSBuildPath $BuildTool -Platform $Target -Configuration $Profile
                        } $repoRoot $msbuild $platform $configuration
                        $explicitRestoreHost | Should Be 'x86'
                    } finally { $env:PreferredToolArchitecture = $null }
                }
            }
        } finally { $env:PreferredToolArchitecture = $previousHost }
    }
}

Describe 'Packaged DxUi module provenance' {
    It 'ships only included module sidecars and rejects mismatched pins or module bytes' {
        $root = Join-Path $TestDrive 'package-provenance'
        $output = Join-Path $root 'build'
        $package = Join-Path $root 'package'
        foreach ($path in @((Join-Path $root 'Dependencies'),(Join-Path $output 'Plugins'),(Join-Path $package 'Plugins'))) {
            [void](New-Item -ItemType Directory -Path $path -Force)
        }
        $pin = @{repository='https://github.com/RedSalamanders/DxUi';commit=('a'*40);apiRevision=2;targets=@('DxUi')}
        $pin | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $root 'Dependencies/DxUi.lock.json')
        foreach ($relative in @('RedSalamander.exe','Plugins/ViewerText.dll','ProductUiTests.exe')) {
            $module = Join-Path $output $relative
            'fixture binary' | Set-Content -LiteralPath $module
            $record = @{repository=$pin.repository;commit=$pin.commit;module=[IO.Path]::GetFileName($relative);moduleSha256=(Get-FileHash $module).Hash}
            $record | ConvertTo-Json | Set-Content -LiteralPath ($module+'.DxUi.json')
            if ($relative -ne 'ProductUiTests.exe') { Copy-Item -LiteralPath $module -Destination (Join-Path $package $relative) }
        }
        $arguments = @{RepoRoot=$root;BuildOutputDir=$output;PackageRoot=$package}
        Copy-RSDxUiPackageProvenance @arguments
        @(Get-ChildItem -LiteralPath $package -Filter '*.DxUi.json' -Recurse).Count | Should Be 2
        Test-Path -LiteralPath (Join-Path $package 'ProductUiTests.exe.DxUi.json') | Should Be $false
        $sidecar = Join-Path $output 'RedSalamander.exe.DxUi.json'
        $record = Get-Content -LiteralPath $sidecar -Raw | ConvertFrom-Json
        $record.commit = 'b'*40
        $record | ConvertTo-Json | Set-Content -LiteralPath $sidecar
        $rejected=$false; try { Copy-RSDxUiPackageProvenance @arguments } catch { $rejected=$true }; $rejected | Should Be $true
        $record.commit = $pin.commit
        $record | ConvertTo-Json | Set-Content -LiteralPath $sidecar
        Add-Content -LiteralPath (Join-Path $package 'Plugins/ViewerText.dll') -Value 'changed after build'
        $rejected=$false; try { Copy-RSDxUiPackageProvenance @arguments } catch { $rejected=$true }; $rejected | Should Be $true
    }
}
