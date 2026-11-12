$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'Tools/Modules/Build/DxUiDependency.psm1') -Force

Describe 'SDK consumer Windows header policy' {
    It 'retains application macro semantics and the shared ABI guards in both native Release profiles' {
        $sdk = Join-Path $TestDrive 'sdk-profile'
        [void][IO.Directory]::CreateDirectory($sdk)
        @'
<Project>
  <PropertyGroup><DxUiSdkRoot>$(MSBuildThisFileDirectory)</DxUiSdkRoot></PropertyGroup>
  <ItemDefinitionGroup><ClCompile><PreprocessorDefinitions>UNICODE;_UNICODE;WIN32_LEAN_AND_MEAN;NOMINMAX;NDEBUG;NOMINMAX_FEATURE_CONTROL;%(PreprocessorDefinitions)</PreprocessorDefinitions></ClCompile></ItemDefinitionGroup>
</Project>
'@ | Set-Content -LiteralPath (Join-Path $sdk 'DxUi.Sdk.props')
        $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
        $installation = (& $vswhere -latest -prerelease -products '*' -requires Microsoft.Component.MSBuild -property installationPath).Trim()
        $hostFolder = if ([Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString() -ceq 'Arm64') {'arm64'} else {'amd64'}
        $msbuild = Join-Path $installation "MSBuild/Current/Bin/$hostFolder/MSBuild.exe"
        $taskProbe = Join-Path $sdk 'definition-list.targets'
        @'
<Project>
  <Target Name="I61DefinitionListProbe">
    <WriteLinesToFile File="$(I61DefinitionLines)" Lines="%(ClCompile.PreprocessorDefinitions)" Overwrite="true" />
  </Target>
</Project>
'@ | Set-Content -LiteralPath $taskProbe
        foreach ($platform in @('x64','ARM64')) {
            $json = & $msbuild (Join-Path $repoRoot 'Tests/RedConfigureTests/RedConfigureTests.vcxproj') /nologo /p:Configuration=Release "/p:Platform=$platform" "/p:RSDxUiSdkRoot=$sdk" "/p:RSCiEngineRoot=$(Join-Path $repoRoot '.ci-engine')" '-getItem:ClCompile,ProjectReference'
            if ($LASTEXITCODE) {throw "Cannot evaluate SDK Windows macro policy for $platform."}
            $evaluated = ($json -join "`n") | ConvertFrom-Json
            @($evaluated.Items.ClCompile).Count | Should BeGreaterThan 0
            foreach ($item in $evaluated.Items.ClCompile) {
                $item.PreprocessorDefinitions | Should Not Match '(^|;)(WIN32_LEAN_AND_MEAN|NOMINMAX)(=1)?(;|$)'
                $item.PreprocessorDefinitions | Should Match '(^|;)NOMINMAX_FEATURE_CONTROL(;|$)'
                $item.PreprocessorDefinitions | Should Match '(^|;)NDEBUG(;|$)'
                $item.PreprocessorDefinitions | Should Match '(^|;)UNICODE(;|$)'
            }
            @($evaluated.Items.ProjectReference | Where-Object {[IO.Path]::GetFileName($_.Identity) -ceq 'DxUi.vcxproj'}).Count | Should Be 0
            # A getItem string hides escaped delimiters. Exercise an actual
            # MSBuild task's list binding, as the compiler task consumes it.
            $lines = Join-Path $sdk ("definitions-$platform.txt")
            $probeOutput = & $msbuild (Join-Path $repoRoot 'Tests/RedConfigureTests/RedConfigureTests.vcxproj') /nologo /p:Configuration=Release "/p:Platform=$platform" "/p:RSDxUiSdkRoot=$sdk" "/p:RSCiEngineRoot=$(Join-Path $repoRoot '.ci-engine')" "/p:ForceImportAfterCppTargets=$taskProbe" "/p:I61DefinitionLines=$lines" /t:I61DefinitionListProbe 2>&1
            if ($LASTEXITCODE) {throw "Cannot probe actual MSBuild definition list binding for $platform`: $($probeOutput -join [Environment]::NewLine)"}
            $definitions = @(Get-Content -LiteralPath $lines)
            $definitions.Count | Should BeGreaterThan 4
            foreach ($definition in $definitions) {$definition | Should Not Match ';'}
            ($definitions -ccontains 'NOMINMAX_FEATURE_CONTROL') | Should Be $true
            ($definitions -ccontains 'NDEBUG') | Should Be $true
            ($definitions -ccontains 'UNICODE') | Should Be $true
        }
        # Metadata evaluation above grants no build/ABI verdict. Actual SDK builds
        # must still execute the shared live-reader and per-TU admission targets.
        [IO.File]::ReadAllText((Join-Path $repoRoot 'Build/RedSalamander.DxUi.props')) | Should Match 'Ci\.SdkConsumer\.targets'
    }
}

Describe 'Folded DxUi change notes' {
    It 'directs an API mismatch to the generated changelog and pending fragments' {
        $message = ''
        try {
            & (Get-Module DxUiDependency) {
                Assert-RSDxUiApiRevision -Pin ([pscustomobject]@{apiRevision=4}) -Recorded 3 -Context 'fixture'
            }
        } catch { $message = $_.Exception.Message }
        $message | Should Match 'Read CHANGELOG.md and any pending Changes/'
        $message | Should Match 'apiRevision in the lock'
    }
}

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
            @{repository='https://example.invalid/fork';commit=('a'*40);apiRevision=3;targets=@('DxUi')},
            @{repository='https://github.com/RedSalamanders/DxUi';commit='../escape';apiRevision=3;targets=@('DxUi')},
            @{repository='https://github.com/RedSalamanders/DxUi';commit=('a'*40);apiRevision=0;targets=@('DxUi')},
            @{repository='https://github.com/RedSalamanders/DxUi';commit=('a'*40);apiRevision='3';targets=@('DxUi')},
            @{repository='https://github.com/RedSalamanders/DxUi';commit=('a'*40);targets=@('DxUi')},
            @{repository='https://github.com/RedSalamanders/DxUi';commit=('a'*40);apiRevision=3;targets=@('Foundation')}
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
        @{repository='https://github.com/RedSalamanders/DxUi';commit=$head;apiRevision=3;targets=@('DxUi')} |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path $root 'Dependencies/DxUi.lock.json')
        $rejected=$false; try { Get-RSDxUiSourceIdentity -RepoRoot $root | Out-Null } catch { $rejected=$true }; $rejected | Should Be $true
        Get-RSDxUiSourceIdentity -RepoRoot $root -AllowMissing | Should BeNullOrEmpty
        $source = Join-Path $root ".build/dependencies/DxUi/source/$($head.Substring(0,16))"
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
        @{repository='https://github.com/RedSalamanders/DxUi';commit=$head;apiRevision=3;targets=@('DxUi')} |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path $root 'Dependencies/DxUi.lock.json')
        $dependency = Join-Path $root '.build/dependencies/DxUi'
        $source = Join-Path $dependency "source/$($head.Substring(0,16))"
        [void](New-Item -ItemType Directory -Path (Split-Path $source) -Force)
        & git -c core.longpaths=true clone --config core.longpaths=true --quiet --local $origin $source
        if ($LASTEXITCODE -ne 0) { throw 'Cannot clone isolated provenance fixture.' }
        $identity = @{Fingerprint=('b'*64);Identity=@{commit=$head;apiRevision=3;platform='x64';disableStlAnnotations=$true}}
        $identityPath = Join-Path $dependency 'DxUi.identity.x64.json'
        $identity | ConvertTo-Json | Set-Content -LiteralPath $identityPath
        # The output directory is the fingerprint's 16-character prefix; the identity file keeps the full digest.
        $archive = Join-Path $dependency "$('b'*16)/x64/Debug/DxUi.lib"
        $fullFingerprintArchive = Join-Path $dependency "$($identity.Fingerprint)/x64/Debug/DxUi.lib"
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
        foreach ($invalidLink in @('link.exe "old/DxUi.lib"', ('link.exe "' + $fullFingerprintArchive + '"'),
            ($link + ' "DxUi.lib" '), ($link + ' "Common/DxUi/legacy.obj"'))) {
            $invalidLink | Set-Content -LiteralPath $tlog
            $rejected=$false; try { Write-RSDxUiModuleProvenance @arguments } catch { $rejected=$true }; $rejected | Should Be $true
        }
        $link | Set-Content -LiteralPath $tlog
        foreach ($invalidIdentity in @(
            @{Fingerprint=$identity.Fingerprint;Identity=@{commit=('c'*40);apiRevision=3;platform='x64';disableStlAnnotations=$true}},
            @{Fingerprint=$identity.Fingerprint;Identity=@{commit=$head;apiRevision=3;platform='ARM64';disableStlAnnotations=$true}},
            @{Fingerprint=$identity.Fingerprint;Identity=@{commit=$head;apiRevision=3;platform='x64';disableStlAnnotations=$false}},
            # A resolved identity from another API revision than the lock records, or one that records none.
            @{Fingerprint=$identity.Fingerprint;Identity=@{commit=$head;apiRevision=2;platform='x64';disableStlAnnotations=$true}},
            @{Fingerprint=$identity.Fingerprint;Identity=@{commit=$head;platform='x64';disableStlAnnotations=$true}},
            @{Fingerprint=('b'*16);Identity=@{commit=$head;apiRevision=3;platform='x64';disableStlAnnotations=$true}}
        )) {
            $invalidIdentity | ConvertTo-Json | Set-Content -LiteralPath $identityPath
            $rejected=$false; try { Write-RSDxUiModuleProvenance @arguments } catch { $rejected=$true }; $rejected | Should Be $true
        }
        $identity | ConvertTo-Json | Set-Content -LiteralPath $identityPath
        Write-RSDxUiModuleProvenance @arguments
    }
}


Describe 'Lighter DxUi source restore' {
    function New-RSDxUiRestoreOrigin {
        # A library fixture: what restore and build read, a kept directory deep enough to need Git long paths once it is
        # restored beneath a deep consumer root, and the three directories the restore leaves out. A full run's test drive
        # is itself deep, so the fixture's own repository uses long paths too; the restored checkout never inherits that.
        param([string]$Path, [string]$DeepDirectory, [string[]]$Omit = @())
        $files = [ordered]@{
            'capabilities.json' = '{ "apiRevision": 3 }'
            'vcpkg.json' = '{}'
            'vcpkg-install.ps1' = '# fixture'
            'Tools/ConsumerBuild.psm1' = '# fixture'
            'Build/DxUi.Consumer.props' = '<Project />'
            'Build/DxUi.Consumer.targets' = '<Project />'
            'src/DxUi.vcxproj' = '<Project />'
            'include/DxUi/Example.h' = '// fixture'
            'docs/README.md' = 'kept: only docs/gallery is left out'
            "$DeepDirectory/Deep.psm1" = '# a kept file whose restored path passes MAX_PATH'
            'Measurements/2026-09-21/raw/receipt.txt' = 'retained measurement'
            'docs/gallery/sheet.png' = 'generated sheet'
            'Specs/Plan.md' = 'specification'
        }
        foreach ($relative in $files.Keys) {
            if ($Omit -contains $relative.Split('/')[0]) { continue }
            $target = Join-Path $Path $relative
            [void](New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force)
            [IO.File]::WriteAllText($target, $files[$relative])
        }
        & git -c core.longpaths=true init --quiet $Path
        if ($LASTEXITCODE -ne 0) { throw 'Cannot initialize the restore fixture.' }
        & git -C $Path config core.longpaths true
        if ($LASTEXITCODE -ne 0) { throw 'Cannot initialize the restore fixture.' }
        & git -C $Path add .
        if ($LASTEXITCODE -ne 0) { throw 'Cannot initialize the restore fixture.' }
        & git -C $Path -c user.name=Fixture -c user.email=fixture@example.invalid commit --quiet -m fixture
        if ($LASTEXITCODE -ne 0) { throw 'Cannot initialize the restore fixture.' }
        return (& git -C $Path rev-parse HEAD).Trim()
    }

    It 'restores beneath a deep root with Git long paths, without measurements, specifications or gallery sheets' {
        $deepDirectory = 'Tools/' + ('d' * 45) + '/' + ('e' * 45)
        $origin = Join-Path $TestDrive 'lighter-origin'
        $head = New-RSDxUiRestoreOrigin -Path $origin -DeepDirectory $deepDirectory
        $root = Join-Path $TestDrive 'deep-restore'
        # Full runs deeper sandboxes than a focused run; this root alone puts the kept deep file past MAX_PATH.
        if ($root.Length -lt 160) { $root = Join-Path $root ('x' * (160 - $root.Length)) }
        [void](New-Item -ItemType Directory -Path (Join-Path $root 'Dependencies') -Force)
        @{repository='https://github.com/RedSalamanders/DxUi';commit=$head;apiRevision=3;targets=@('DxUi')} |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path $root 'Dependencies/DxUi.lock.json')
        $source = Join-Path $root ".build/dependencies/DxUi/source/$($head.Substring(0,16))"
        (Join-Path $source "$deepDirectory/Deep.psm1").Length | Should BeGreaterThan 260

        & (Get-Module DxUiDependency) { param($Root, $Origin, $Commit, $Checkout)
            Install-RSDxUiPinnedSource -RepoRoot $Root -Repository $Origin -Commit $Commit -Source $Checkout } $root $origin $head $source

        Test-Path -LiteralPath (Join-Path $source "$deepDirectory/Deep.psm1") | Should Be $true
        foreach ($omitted in @('Measurements', 'docs/gallery', 'Specs')) {
            Test-Path -LiteralPath (Join-Path $source $omitted) | Should Be $false
        }
        foreach ($kept in @('capabilities.json', 'vcpkg-install.ps1', 'Tools/ConsumerBuild.psm1', 'src/DxUi.vcxproj', 'include/DxUi/Example.h', 'docs/README.md')) {
            Test-Path -LiteralPath (Join-Path $source $kept) | Should Be $true
        }
        # Long paths live in the checkout's own configuration; no user or global configuration is involved.
        (& git -C $source config --local --get core.longpaths) | Should Be 'true'
        # The omitted paths are not deletions: the restore's own clean-source check accepts the sparse checkout.
        (Get-RSDxUiSourceIdentity -RepoRoot $root).Pin.commit | Should Be $head
        (& git -C $source rev-parse HEAD).Trim() | Should Be $head
    }

    It 'removes the partial checkout it created when the restore fails, and never one it did not' {
        $origin = Join-Path $TestDrive 'failing-origin'
        # Without include/ the sparse checkout lacks a path the build reads.
        $head = New-RSDxUiRestoreOrigin -Path $origin -DeepDirectory 'Tools/deep' -Omit @('include')
        $root = Join-Path $TestDrive 'failing-restore'
        $source = Join-Path $root ".build/dependencies/DxUi/source/$($head.Substring(0,16))"
        $install = { param($Root, $Repository, $Commit, $Checkout)
            Install-RSDxUiPinnedSource -RepoRoot $Root -Repository $Repository -Commit $Commit -Source $Checkout }
        { & (Get-Module DxUiDependency) $install $root $origin $head $source } | Should Throw "lacks 'include'"
        Test-Path -LiteralPath $source | Should Be $false
        # A commit the origin does not have fails in git itself, and leaves nothing behind either.
        { & (Get-Module DxUiDependency) $install $root $origin ('9' * 40) $source } | Should Throw 'DxUi restore failed'
        Test-Path -LiteralPath $source | Should Be $false
        # Its sibling checkouts are not its to remove.
        $sibling = Join-Path $root ('.build/dependencies/DxUi/source/' + ('8' * 40))
        [void](New-Item -ItemType Directory -Path $sibling -Force)
        'untouched' | Set-Content -LiteralPath (Join-Path $sibling 'marker.txt')
        { & (Get-Module DxUiDependency) $install $root $origin ('9' * 40) $source } | Should Throw 'DxUi restore failed'
        Test-Path -LiteralPath (Join-Path $sibling 'marker.txt') | Should Be $true
    }
}

Describe 'DxUi consumer profile evaluation' {
    It 'uses the first-party host during restore and restores the caller environment after success or failure' {
        $root = Join-Path $TestDrive ('host-restore-' + [guid]::NewGuid().ToString('N'))
        # Exercise a genuinely long Git metadata path in both focused and deeper Full sandboxes.
        # Size the padding from the shortened source locator rather than the old 40-character folder.
        $metadataSuffix = '.build/dependencies/DxUi/source/' + ('0' * 16) + '/.git/hooks/fsmonitor-watchman.sample'
        $metadataPath = Join-Path $root $metadataSuffix
        if ($metadataPath.Length -lt 280) { $root = Join-Path $root ('x' * (280 - $metadataPath.Length)) }
        $origin = Join-Path $root 'fixture-origin'
        foreach ($relative in @('RedSalamander','Dependencies','Build','fixture-origin/include','fixture-origin/Tools')) {
            [void](New-Item -ItemType Directory -Path (Join-Path $root $relative) -Force)
        }
        foreach ($relative in @('Build/NativeBuildPolicy.targets','Build/PinnedDxUiPolicy.targets')) {
            Copy-Item -LiteralPath (Join-Path $repoRoot $relative) -Destination (Join-Path $root $relative)
        }
        $expectedPolicy = & (Get-Module DxUiDependency) { param($Root) Get-DxUiNativePolicyFields -RepoRoot $Root } $root
        '<Project><PropertyGroup><RequiredMsvcToolsVersion>14.52</RequiredMsvcToolsVersion></PropertyGroup></Project>' |
            Set-Content -LiteralPath (Join-Path $root 'Directory.Build.props')
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
    # The revision the pinned source reports; a file beside the checkouts makes the fixture report another one.
    $override = Join-Path (Split-Path (Split-Path $DxUiRoot)) 'api-revision'
    $revision = if (Test-Path -LiteralPath $override) { [int](Get-Content -LiteralPath $override -Raw) } else { 3 }
    $toolsOverride = Join-Path (Split-Path (Split-Path $DxUiRoot)) 'tool-version'
    $tools = if (Test-Path -LiteralPath $toolsOverride) { (Get-Content -LiteralPath $toolsOverride -Raw).Trim() } else { '14.52.37025' }
    return [pscustomobject]@{Fingerprint=('a'*64);Identity=@{
        apiRevision=$revision;toolset='v145';vcToolsVersion=$tools;windowsSdkVersion='fixture';preferredToolArchitecture=$selected
    }}
}
Export-ModuleMember -Function Get-DxUiConsumerBuildIdentity
'@ | Set-Content -LiteralPath (Join-Path $origin 'Tools/ConsumerBuild.psm1')
        # The fixture intentionally exceeds MAX_PATH, including Git's own object database.
        # Configure each bootstrap command before it can read a repository-local setting.
        & git -c core.longpaths=true init --quiet $origin
        if ($LASTEXITCODE -ne 0) { throw 'Cannot initialize long-path compiler-host fixture.' }
        & git -c core.longpaths=true -C $origin config core.longpaths true
        if ($LASTEXITCODE -ne 0) { throw 'Cannot configure long-path compiler-host fixture.' }
        & git -c core.longpaths=true -C $origin add .
        if ($LASTEXITCODE -ne 0) { throw 'Cannot stage long-path compiler-host fixture.' }
        & git -c core.longpaths=true -C $origin -c user.name=Fixture -c user.email=fixture@example.invalid commit --quiet -m fixture
        if ($LASTEXITCODE -ne 0) { throw 'Cannot initialize compiler-host fixture.' }
        $head = (& git -C $origin rev-parse HEAD).Trim()
        @{repository='https://github.com/RedSalamanders/DxUi';commit=$head;apiRevision=3;targets=@('DxUi')} |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path $root 'Dependencies/DxUi.lock.json')
        $dependency = Join-Path $root '.build/dependencies/DxUi'
        $source = Join-Path $dependency "source/$($head.Substring(0,16))"
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
                        $resolved.Identity.releaseOptimizationPolicy | Should Be 'Gy-REF-GL-LTCG-full-v1'
                        $resolved.Identity.nativeBuildPolicy | Should Be $expectedPolicy.nativeBuildPolicy
                        $resolved.Identity.nativeBuildPolicySha256 | Should Be $expectedPolicy.nativeBuildPolicySha256
                        $resolved.Fingerprint | Should Not Be ('a' * 64)
                        [string]$env:PreferredToolArchitecture | Should Be $requestedHost
                    }
                }
            }
            # A pin that crossed an API revision fails before the library's own dependency restore, with the way forward.
            '4' | Set-Content -LiteralPath (Join-Path $dependency 'api-revision')
            { Restore-RSDxUiDependency -RepoRoot $root -MSBuildPath $msbuild -Platform x64 -Configuration Debug } |
                Should Throw 'reports API revision 4, but Dependencies/DxUi.lock.json records API revision 3'
            Remove-Item -LiteralPath (Join-Path $dependency 'api-revision')
            '14.51.36231' | Set-Content -LiteralPath (Join-Path $dependency 'tool-version')
            { Restore-RSDxUiDependency -RepoRoot $root -MSBuildPath $msbuild -Platform x64 -Configuration Release } |
                Should Throw 'requires MSVC 14.52.x'
            [string]$env:PreferredToolArchitecture | Should Be 'x86'
            Remove-Item -LiteralPath (Join-Path $dependency 'tool-version')
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
            'MSBuild/Current/Bin/arm64/MSBuild.exe'
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

Describe 'Qualified SDK CI preparation routing' {
    It 'preserves actual normalization and build arguments across sixty role, profile, architecture and availability combinations' {
        $adapter=Join-Path $repoRoot 'Tools/Invoke-ProjectCi.ps1'
        $originalHash=(Get-FileHash -LiteralPath $adapter).Hash
        $tokens=$null;$errors=$null
        $ast=[Management.Automation.Language.Parser]::ParseFile($adapter,[ref]$tokens,[ref]$errors)
        @($errors).Count | Should Be 0
        $normalize=@($ast.FindAll({param($n) $n -is [Management.Automation.Language.IfStatementAst] -and $n.Extent.Text.StartsWith("if (`$Phase -ne 'Qualification')")},$true))
        $prepare=@($ast.FindAll({param($n) $n -is [Management.Automation.Language.IfStatementAst] -and $n.Extent.Text.StartsWith("if(`$Configuration -ceq 'Release' -and `$Phase -cin @('ReleaseTestBuild','Production'))")},$true))
        $commands=@($ast.FindAll({param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -ceq 'Run' -and $n.CommandElements[1].Extent.Text -ceq "'build.ps1'"},$true))
        $normalize.Count | Should Be 1
        $prepare.Count | Should Be 1
        $commands.Count | Should Be 2
        $production=@($commands|Where-Object {$_.Extent.Text.Contains('-ProductionOnly')})
        $testBuild=@($commands|Where-Object {-not $_.Extent.Text.Contains('-ProductionOnly')})
        $production.Count | Should Be 1
        $testBuild.Count | Should Be 1
        $stage=[scriptblock]::Create(@'
param($Phase,$Configuration,$Platform,$Available,$Probe,$Pin)
$root='Z:/Synthetic workspace';$EngineRoot='Z:/Controlled engine';$context=@{fixture=$true};$number='123'
function Import-Module {param($Name) if(-not($Name.EndsWith('Ci.SdkPreparation.psm1') -or $Name.EndsWith('DxUiDependency.psm1'))){throw 'Unexpected SDK module'};$Probe.imports++}
function Get-RSDxUiPin {param($RepoRoot) return $Pin}
function Initialize-CiPinnedSdkInput {param($Context,$SourceCommit,$ApiRevision) $Probe.prepares++;$Probe.source=$SourceCommit;$Probe.api=$ApiRevision;if($Available){'Z:/Synthetic workspace/.build/input descriptor.json'}}
function Run {param($Script,$CommandArguments) $Probe.commands.Add(@{script=$Script;arguments=@($CommandArguments)})}
'@ + "`n" + $normalize[0].Extent.Text + "`n" + '$profile=@(''-Platform'',$Platform,''-Configuration'',$Configuration);$sdkBuildArguments=@()' + "`n" + $prepare[0].Extent.Text + "`n" +
            'if($Phase -ceq ''Production''){' + $production[0].Extent.Text + '}' + "`n" +
            'if($Phase -cin @(''ReleaseTestBuild'',''Qualification'')){' + $testBuild[0].Extent.Text + '}')
        $pin=Get-Content (Join-Path $repoRoot 'Dependencies/DxUi.lock.json') -Raw|ConvertFrom-Json
        $pin.commit | Should Be '4cdf1ce52c7785446698e346b8aec058ed6a9283'
        $pin.apiRevision | Should Be 4
        $count=0
        foreach($phase in @('Static','ReleaseTestBuild','ReleaseTest','Production','Qualification')){
            foreach($configuration in @('Debug','Release','ASan Debug')){foreach($platform in @('x64','ARM64')){foreach($available in @($true,$false)){
                $probe=@{imports=0;prepares=0;source='';api=0;commands=[Collections.Generic.List[object]]::new()}
                & $stage $phase $configuration $platform $available $probe $pin
                $eligible=$phase -cin @('ReleaseTestBuild','Production')
                $probe.prepares | Should Be ([int]$eligible)
                $probe.imports | Should Be $(if($eligible){2}else{0})
                if($eligible){$probe.source|Should Be $pin.commit;$probe.api|Should Be $pin.apiRevision}
                $probe.commands.Count | Should Be $(if($phase -cin @('ReleaseTestBuild','Production','Qualification')){1}else{0})
                foreach($command in $probe.commands){
                    $hasDescriptor='-DxUiSdkInputDescriptor' -cin $command.arguments
                    $hasDescriptor | Should Be ($eligible -and $available)
                    if($hasDescriptor){('Z:/Synthetic workspace/.build/input descriptor.json' -cin $command.arguments)|Should Be $true}
                    $effective=if($phase -ceq 'Qualification'){$configuration}else{'Release'}
                    ($command.arguments[0..3] -join '|')|Should Be ("-Platform|$platform|-Configuration|$effective")
                    ('-ProductionOnly' -cin $command.arguments)|Should Be ($phase -ceq 'Production')
                }
                $count++
            }}}
        }
        $count | Should Be 60
        (Get-FileHash -LiteralPath $adapter).Hash | Should Be $originalHash
    }

    It 'prepares SDK inputs before applying integration and preserves inherited compiler overrides' {
        $tokens=$null;$errors=$null
        $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repoRoot 'Tools/Invoke-ProjectCi.ps1'),[ref]$tokens,[ref]$errors)
        @($errors).Count | Should Be 0
        $guard=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.IfStatementAst] -and $n.Extent.Text.StartsWith('if($sdkBuildArguments.Count -gt 0')},$true))
        $guard.Count|Should Be 1
        $prepare=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -ceq 'Initialize-CiPinnedSdkInput'},$true))
        $apply=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.ForEachStatementAst] -and $n.Extent.Text.Contains('$environment.GetEnumerator()')},$true))
        $prepare.Count|Should Be 1
        $apply.Count|Should Be 1
        ($prepare[0].Extent.StartOffset -lt $guard[0].Extent.StartOffset)|Should Be $true
        ($guard[0].Extent.EndOffset -lt $apply[0].Extent.StartOffset)|Should Be $true
        $body=[scriptblock]::Create($guard[0].Extent.Text)
        $savedCL=$env:CL
        try{
            $env:CL='inherited override remains visible'
            foreach($phase in @('ReleaseTestBuild','ReleaseTest','Production','Qualification')){
                foreach($sdk in @($false,$true)){foreach($record in @($false,$true)){
                    $result=& {
                        param($Body,$Phase,$Sdk,$Record)
                        $root='Z:/Recording fixture';$Platform='x64';$Configuration='Release'
                        $sdkBuildArguments=@(if($Sdk){'-DxUiSdkInputDescriptor';'fixture.json'})
                        $environment=[ordered]@{CL='synthetic vcpkg include';LINK='original libraries';ForceImportBeforeCppTargets='canonical.props'}
                        function Test-Path {param($LiteralPath,$PathType) return $Record}
                        . $Body
                        return $environment
                    } $body $phase $sdk $record
                    $result.Contains('CL')|Should Be (-not($sdk -or ($phase -ceq 'ReleaseTest' -and $record)))
                    $result.LINK|Should Be 'original libraries'
                    $result.ForceImportBeforeCppTargets|Should Be 'canonical.props'
                    $env:CL|Should Be 'inherited override remains visible'
                }}
            }
        }finally{$env:CL=$savedCL}
    }
}

Describe 'SDK input data records and fail-closed artifact routing' {
    It 'writes a strict receipt input record and refuses unknown, stale or executable-selecting data before SDK execution' {
        $root=Join-Path $TestDrive 'sdk-record'
        $output=Join-Path $root '.build/x64/Release'
        $sdk=Join-Path $root '.build/sdk'
        foreach($directory in @($output,(Join-Path $sdk 'lib/x64/Release'),(Join-Path $root 'Specs/Build'),(Join-Path $root 'Dependencies'))){[void][IO.Directory]::CreateDirectory($directory)}
        Copy-Item -LiteralPath (Join-Path $repoRoot 'Specs/Build/Build_DxUiSdkInput.schema.json') -Destination (Join-Path $root 'Specs/Build/Build_DxUiSdkInput.schema.json')
        $pin=@{repository='https://github.com/RedSalamanders/DxUi';commit=('a'*40);apiRevision=4;targets=@('DxUi')}
        [IO.File]::WriteAllText((Join-Path $root 'Dependencies/DxUi.lock.json'),($pin|ConvertTo-Json))
        $descriptor=Join-Path $root '.build/input.json'
        [IO.File]::WriteAllText($descriptor,'{"role":"fixture-data-only"}')
        $tool=Join-Path $root '.build/unexpected-tool.exe'
        [IO.File]::WriteAllText($tool,'Non-executable fixture data; must never be launched')
        [IO.File]::WriteAllText((Join-Path $sdk 'lib/x64/Release/DxUi.lib'),'fixture library bytes, not native proof')
        $reader=[pscustomobject]@{Closed=$false;SdkRoot=$sdk;Descriptor=@{Path=$descriptor;Hash=(Get-FileHash $descriptor).Hash;Value=@{sourceCommit=$pin.commit;apiRevision=4}};
            Evidence=@{recipe='PinnedQualifiedReleaseSdkV2';inputKey=('b'*64);manifestSha256=('c'*64);archiveSha256=('d'*64);toolchainReference=('e'*64)}}
        $sdkRecordArguments=@{RepoRoot=$root;BuildOutputDir=$output;SdkInput=$reader;MSBuildPath=$tool;Platform='x64'}
        $recordPath=Write-RSDxUiSdkBuildRecord @sdkRecordArguments
        $record=Get-Content -LiteralPath $recordPath -Raw|ConvertFrom-Json
        $record.inputKey | Should Be ('b'*64)
        $record.descriptorSha256 | Should Be (Get-FileHash $descriptor).Hash.ToLowerInvariant()
        $record.librarySha256 | Should Be (Get-FileHash (Join-Path $sdk 'lib/x64/Release/DxUi.lib')).Hash.ToLowerInvariant()
        $reader.Closed=$true
        $reader.Closed | Should Be $true
        $sdkRecordArguments.SdkInput.Closed | Should Be $true
        {Write-RSDxUiSdkBuildRecord @sdkRecordArguments}|Should Throw 'active qualified SDK reader'
        $reader.Closed=$false
        $reader.Evidence.recipe='PinnedReleaseSdkV1'
        {Write-RSDxUiSdkBuildRecord @sdkRecordArguments}|Should Throw 'active qualified SDK reader'
        $reader.Evidence.recipe='PinnedQualifiedReleaseSdkV2'
        $toolchain=Join-Path $output 'toolchain-identity.json'
        [IO.File]::WriteAllText($toolchain,(@{msbuild=@{sha256=$record.msbuildSha256}}|ConvertTo-Json))
        $receipt=[pscustomobject]@{platform='x64';configuration='Release';build_arguments=@('dxui-recipe=PinnedQualifiedReleaseSdkV2',"dxui-sdk-input-key=$($record.inputKey)","dxui-sdk-manifest=$($record.manifestSha256)","dxui-sdk-descriptor-sha256=$($record.descriptorSha256)");
            outputs=@(@{relative_path='.build/x64/Release/DxUi.sdk-input.json';sha256=(Get-FileHash $recordPath).Hash},@{relative_path='.build/x64/Release/toolchain-identity.json';sha256=(Get-FileHash $toolchain).Hash})}
        {Open-RSDxUiSdkArtifactInput -RepoRoot $root -Receipt $receipt}|Should Throw 'cannot select another executable'
        $original=[IO.File]::ReadAllText($recordPath)
        foreach($change in @('schema','pin','api','profile','descriptor')){
            $bad=$original|ConvertFrom-Json
            switch($change){'schema'{$bad|Add-Member -NotePropertyName passed -NotePropertyValue $true};'pin'{$bad.sourceCommit='f'*40};'api'{$bad.apiRevision=5};'profile'{$bad.platform='ARM64'};'descriptor'{$bad.descriptorSha256='0'*64}}
            [IO.File]::WriteAllText($recordPath,($bad|ConvertTo-Json))
            $receipt.outputs[0].sha256=(Get-FileHash $recordPath).Hash
            $message=switch($change){'schema'{'schema'};'pin'{'product pin or native profile'};'api'{'product pin or native profile'};'profile'{'product pin or native profile'};'descriptor'{'arguments'}}
            {Open-RSDxUiSdkArtifactInput -RepoRoot $root -Receipt $receipt}|Should Throw $message
        }
        [IO.File]::WriteAllText($recordPath,$original)
        $receipt.outputs[0].sha256=(Get-FileHash $recordPath).Hash
        [IO.File]::AppendAllText($descriptor,'changed')
        {Open-RSDxUiSdkArtifactInput -RepoRoot $root -Receipt $receipt}|Should Throw 'descriptor changed'
        $receipt.build_arguments=@('dxui-recipe=Unknown')
        {Open-RSDxUiSdkArtifactInput -RepoRoot $root -Receipt $receipt}|Should Throw 'Unknown or unsupported'
        $receipt.build_arguments=@()
        {Open-RSDxUiSdkArtifactInput -RepoRoot $root -Receipt $receipt}|Should Throw 'cannot downgrade'
        $receipt.outputs=@()
        {Open-RSDxUiSdkArtifactInput -RepoRoot $root -Receipt $receipt}|Should Throw 'cannot downgrade'
        Remove-Item -LiteralPath $recordPath
        (Open-RSDxUiSdkArtifactInput -RepoRoot $root -Receipt $receipt) | Should Be $null
        Close-RSDxUiSdkArtifactInput $null
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
        $pin = @{repository='https://github.com/RedSalamanders/DxUi';commit=('a'*40);apiRevision=3;targets=@('DxUi')}
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
