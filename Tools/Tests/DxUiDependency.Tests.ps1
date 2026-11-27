Set-StrictMode -Version Latest
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$script:dxUiTestModule = Import-Module (Join-Path $repoRoot 'Tools/Modules/Build/DxUiDependency.psm1') -Force -PassThru
$script:dxUiMetadataAudit = $null

function Invoke-RSDxUiTestGit {
    # Fixture setup may write only its sandbox repository; use the same safe
    # capture boundary rather than executing a legacy bare-CUI baseline.
    $capture = & $script:dxUiTestModule {
        param($Directory, $GitArguments)
        Invoke-RSDxUiSourceGitCapture -WorkingDirectory $Directory -Arguments $GitArguments
    } $repoRoot ([string[]]$args)
    if ($null -ne $script:dxUiMetadataAudit -and $script:dxUiMetadataAudit.Active) {
        $record = Assert-RSTestSandboxContainedPath -TestRoot $script:dxUiMetadataAudit.TestRoot `
            -Path (Join-Path $script:dxUiMetadataAudit.Directory 'git-captures.jsonl') -RequireDescendant
        [ordered]@{arguments=@($args);cwd=$repoRoot;exit=$capture.ExitCode;stdout=@($capture.StdOut);stderr=@($capture.StdErr)} |
            ConvertTo-Json -Depth 5 -Compress | Add-Content -LiteralPath $record -Encoding utf8
    }
    $global:LASTEXITCODE = $capture.ExitCode
    return $capture.StdOut
}



function Invoke-RSDxUiTestMetadataCapture {
    param([string]$FilePath, [string[]]$Arguments, [string]$WorkingDirectory=$repoRoot)
    return & $script:dxUiTestModule {
        param($File,$Values,$Directory)
        Invoke-RSDxUiReadOnlyCapture -FilePath $File -Arguments $Values -WorkingDirectory $Directory
    } $FilePath $Arguments $WorkingDirectory
}

function Start-RSDxUiMetadataAudit {
    # Test-only interception; canonical starter still owns launch and lifetime.
    $directory = New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness d -Case ('metadata-'+[guid]::NewGuid().ToString('N').Substring(0,8))
    $testRoot = Get-RSTestSandboxRoot -RepoRoot $repoRoot
    $record = Assert-RSTestSandboxContainedPath -TestRoot $testRoot -Path (Join-Path $directory 'launches.jsonl') -RequireDescendant
    $state = [pscustomobject]@{PreviousRecord=$env:RS_DXUI_METADATA_RECORD;PreviousStarter=$env:RS_DXUI_METADATA_STARTER;
        PreviousAudit=$script:dxUiMetadataAudit;Directory=$directory;TestRoot=$testRoot;Active=$true;
        Record=$record;Starter=(Join-Path $repoRoot 'Tools/Modules/Build/SanitizedEnvironment.psm1')}
    $script:dxUiMetadataAudit = $state
    $env:RS_DXUI_METADATA_RECORD = Join-Path $directory 'launches.jsonl'
    $env:RS_DXUI_METADATA_STARTER = Join-Path $repoRoot 'Tools/Modules/Build/SanitizedEnvironment.psm1'
    $installed = $false
    try {
    Mock Start-RSContainedProcess -ModuleName ProcessStreaming ({
        param($ProcessStartInfo,[switch]$DelegateArtifactOperation)
        if(-not $state.Active){throw 'Metadata audit lease is inactive.'}
        if($ProcessStartInfo.UseShellExecute -or -not $ProcessStartInfo.CreateNoWindow -or
            -not $ProcessStartInfo.RedirectStandardOutput -or -not $ProcessStartInfo.RedirectStandardError -or $DelegateArtifactOperation){
            throw 'Refusing unsafe DxUi metadata fixture launch.'
        }
        $modules = @(Get-Module -All | Where-Object { $_.Path -ceq $state.Starter })
        if($modules.Count -ne 1){throw 'Expected one canonical metadata starter module.'}
        $starter = $modules[0].ExportedFunctions['Start-RSContainedProcess']
        $record = Assert-RSTestSandboxContainedPath -TestRoot $state.TestRoot -Path (Join-Path $state.Directory 'launches.jsonl') -RequireDescendant
        $owned = & $starter -ProcessStartInfo $ProcessStartInfo
        $transferred = $false
        try {
            [ordered]@{process=$owned.Id;exact_job=$owned.IsInContainmentJob;file=$ProcessStartInfo.FileName;cwd=$ProcessStartInfo.WorkingDirectory;
                arguments=@($ProcessStartInfo.ArgumentList);no_window=$ProcessStartInfo.CreateNoWindow;stdout=$ProcessStartInfo.RedirectStandardOutput;
                stderr=$ProcessStartInfo.RedirectStandardError;delegated=[bool]$DelegateArtifactOperation} |
                ConvertTo-Json -Compress | Add-Content -LiteralPath $record -Encoding utf8
            $transferred = $true
            return $owned
        } finally {if(-not $transferred){$owned.Dispose()}}
    }.GetNewClosure())
        $installed = $true
        return $state
    } finally {
        if(-not $installed){Stop-RSDxUiMetadataAudit -State $state}
    }
}

function Stop-RSDxUiMetadataAudit {
    param($State)
    try {
        $State.Active = $false
        $script:dxUiMetadataAudit = $State.PreviousAudit
    } finally {
        $env:RS_DXUI_METADATA_RECORD = $State.PreviousRecord
        $env:RS_DXUI_METADATA_STARTER = $State.PreviousStarter
    }
}

function Assert-RSDxUiMetadataAudit {
    param($State, [int]$ExpectedMSBuild)
    $rows=@(Get-Content (Join-Path $State.Directory 'launches.jsonl')|ForEach-Object{$_|ConvertFrom-Json})
    @($rows|Where-Object{[IO.Path]::GetFileName($_.file) -ieq 'MSBuild.exe'}).Count|Should Be $ExpectedMSBuild
    @($rows|Where-Object{[IO.Path]::GetFileName($_.file) -ieq 'vswhere.exe'}).Count|Should Be 1
    @($rows|Where-Object{-not $_.no_window -or -not $_.stdout -or -not $_.stderr -or -not $_.exact_job -or $_.delegated}).Count|Should Be 0
}

Describe 'Captured DxUi source Git identity' {
    function New-RSDxUiCaptureFixture {
        param([string]$Root)
        [void][IO.Directory]::CreateDirectory((Join-Path $Root 'Dependencies'))
        $commit = 'a' * 40
        @{repository='https://github.com/RedSalamanders/DxUi';commit=$commit;apiRevision=3;targets=@('DxUi')} |
            ConvertTo-Json | Set-Content (Join-Path $Root 'Dependencies/DxUi.lock.json') -Encoding utf8
        [void][IO.Directory]::CreateDirectory((Join-Path $Root ('.build/dependencies/DxUi/source/' + $commit.Substring(0,16))))
        return $Root
    }

    It 'captures separate line arrays and actual exit using the first Git application without host output or delegation' {
        $captureRoot = Join-Path $TestDrive 'capture-streams'
        [void][IO.Directory]::CreateDirectory($captureRoot)
        Mock Get-Command -ModuleName DxUiDependency {
            @([pscustomobject]@{Source='C:\first Git\git.exe'},[pscustomobject]@{Source='C:\second Git\git.exe'})
        } -ParameterFilter { $Name -ceq 'git' -and $CommandType -eq 'Application' -and $env:RS_DXUI_FAKE_GIT -ceq '1' }
        Mock Invoke-RSStreamingProcess -ModuleName ProcessStreaming {
            param($FilePath,$Arguments,$WorkingDirectory,$OutputLineCallback,[switch]$DelegateArtifactOperation)
            @{file=$FilePath;arguments=@($Arguments);cwd=$WorkingDirectory;delegated=[bool]$DelegateArtifactOperation} |
                ConvertTo-Json -Compress | Set-Content (Join-Path $WorkingDirectory 'recording.json') -Encoding utf8
            & $OutputLineCallback 'first line' $false
            & $OutputLineCallback '' $false
            & $OutputLineCallback 'snow 雪' $false
            & $OutputLineCallback 'separate error' $true
            return 17
        } -ParameterFilter { $WorkingDirectory -like '*capture-streams' }
        $priorFake = $env:RS_DXUI_FAKE_GIT
        try {
            $env:RS_DXUI_FAKE_GIT = '1'
            $result = @(& $script:dxUiTestModule {
                param($Root)
                Invoke-RSDxUiSourceGitCapture -WorkingDirectory $Root -Arguments @('literal space','snow 雪','$notAnExpression')
            } $captureRoot)
        } finally { $env:RS_DXUI_FAKE_GIT = $priorFake }
        $result.Count | Should Be 1
        $result[0].ExitCode | Should Be 17
        ($result[0].StdOut -join '|') | Should Be 'first line||snow 雪'
        ($result[0].StdErr -join '|') | Should Be 'separate error'
        $record = Get-Content (Join-Path $captureRoot 'recording.json') -Raw | ConvertFrom-Json
        $record.file | Should Be 'C:\first Git\git.exe'
        ($record.arguments -join '|') | Should Be 'literal space|snow 雪|$notAnExpression'
        $record.cwd | Should Be $captureRoot
        $record.delegated | Should Be $false
    }

    It 'keeps the three literal identity queries and clean empty output separate from stderr' {
        $root = New-RSDxUiCaptureFixture (Join-Path $TestDrive 'capture identity root')
        Mock Invoke-RSStreamingProcess -ModuleName ProcessStreaming {
            param($FilePath,$Arguments,$WorkingDirectory,$OutputLineCallback,[switch]$DelegateArtifactOperation)
            [IO.File]::AppendAllText((Join-Path $WorkingDirectory 'queries.jsonl'), ((@{arguments=@($Arguments);cwd=$WorkingDirectory;delegated=[bool]$DelegateArtifactOperation}|ConvertTo-Json -Compress)+"`n"))
            if ($Arguments[-1] -ceq 'HEAD') { & $OutputLineCallback ('a'*40) $false }
            elseif ($Arguments[-1] -ceq 'HEAD:include') { & $OutputLineCallback ('b'*40) $false }
            else { & $OutputLineCallback 'warning is not a dirty stdout row' $true }
            return 0
        } -ParameterFilter { $WorkingDirectory -like '*capture identity root' }
        $identity = Get-RSDxUiSourceIdentity -RepoRoot $root
        $identity.Pin.commit | Should Be ('a'*40)
        $identity.PublicHeadersGitTree | Should Be ('b'*40)
        $queries = @(Get-Content (Join-Path $root 'queries.jsonl') | ForEach-Object { $_ | ConvertFrom-Json })
        $queries.Count | Should Be 3
        ($queries[0].arguments -join '|') | Should Be ('-C|'+$identity.Source+'|rev-parse|HEAD')
        ($queries[1].arguments -join '|') | Should Be ('-C|'+$identity.Source+'|--no-optional-locks|status|--porcelain|--untracked-files=normal')
        ($queries[2].arguments -join '|') | Should Be ('-C|'+$identity.Source+'|rev-parse|HEAD:include')
        @($queries | Where-Object { $_.cwd -cne $root }).Count | Should Be 0
        @($queries | Where-Object delegated).Count | Should Be 0
        $work = New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness d -Case ('identity-'+[guid]::NewGuid().ToString('N').Substring(0,8))
        $queries | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $work 'identity-queries.json') -Encoding utf8
    }

    It 'fails closed on <Mode> without accepting captured stdout as success' -TestCases @(
        @{Mode='head-exit'}, @{Mode='pin-mismatch'}, @{Mode='dirty-status'}, @{Mode='status-exit'}, @{Mode='tree-exit'}
    ) {
        param($Mode)
        $root = New-RSDxUiCaptureFixture (Join-Path $TestDrive ('capture-'+$Mode))
        $Mode | Set-Content (Join-Path $root 'query-mode')
        Mock Invoke-RSStreamingProcess -ModuleName ProcessStreaming {
            param($FilePath,$Arguments,$WorkingDirectory,$OutputLineCallback)
            $mode = (Get-Content (Join-Path $WorkingDirectory 'query-mode') -Raw).Trim()
            if ($Arguments[-1] -ceq 'HEAD') {
                & $OutputLineCallback ($(if($mode -ceq 'pin-mismatch'){'c'*40}else{'a'*40})) $false
                if($mode -ceq 'head-exit'){return 128}
            } elseif ($Arguments[-1] -ceq 'HEAD:include') {
                if($mode -ceq 'tree-exit'){& $OutputLineCallback 'missing include tree' $true;return 128}
                & $OutputLineCallback ('b'*40) $false
            } else {
                if($mode -ceq 'dirty-status'){& $OutputLineCallback '?? untracked 雪.txt' $false}
                if($mode -ceq 'status-exit'){return 17}
            }
            return 0
        } -ParameterFilter { (Test-Path -LiteralPath (Join-Path $WorkingDirectory 'query-mode')) }
        $message = ''
        try { Get-RSDxUiSourceIdentity -RepoRoot $root | Out-Null } catch { $message = $_.Exception.Message }
        $expected = if($Mode -in @('head-exit','pin-mismatch')){'does not match the product pin'}elseif($Mode -eq 'tree-exit'){'Cannot identify the pinned DxUi public headers'}else{'restored DxUi source is dirty'}
        $message | Should Match ([regex]::Escape($expected))
    }

    It 'preflights actual Git PSI before launching and preserves Unicode, exit, environment and owned-job facts' {
        $expectedStreamPath = [IO.Path]::GetFullPath((Join-Path $repoRoot 'Tools/Modules/Build/ProcessStreaming.psm1'))
        # Resolve in the actual consumer session. Its current Get-Command can be
        # a legitimate Pester proxy; validate the canonical export owner instead.
        $streamCandidates = @(& $script:dxUiTestModule {
            Get-Module -Name ProcessStreaming -All | Where-Object { $_.ModuleType -eq 'Script' }
        })
        $streamCandidates.Count | Should Be 1
        $stream = $streamCandidates[0]
        $export = $stream.ExportedFunctions['Invoke-RSStreamingProcess']
        $validOwner = { param([object[]]$Candidates, [string]$ExpectedPath, $Command)
            if ($null -eq $Candidates -or $Candidates.Count -ne 1) { return $false }
            $module = $Candidates[0]
            if ($module -isnot [Management.Automation.PSModuleInfo] -or
                $module.ModuleType -ne 'Script' -or [string]::IsNullOrEmpty($module.Path)) { return $false }
            return [string]::Equals([IO.Path]::GetFullPath($module.Path), $ExpectedPath, [StringComparison]::OrdinalIgnoreCase) -and
                $Command -is [Management.Automation.FunctionInfo] -and [object]::ReferenceEquals($Command.Module, $module)
        }
        (& $validOwner $streamCandidates $expectedStreamPath $export) | Should Be $true
        (& $validOwner @() $expectedStreamPath $export) | Should Be $false
        (& $validOwner @($null) $expectedStreamPath $export) | Should Be $false
        (& $validOwner @([pscustomobject]@{Path=$expectedStreamPath;ModuleType='Script'}) $expectedStreamPath $export) | Should Be $false
        (& $validOwner $streamCandidates ($expectedStreamPath+'.other') $export) | Should Be $false
        (& $validOwner $streamCandidates $expectedStreamPath $script:dxUiTestModule.ExportedFunctions['Get-RSDxUiPin']) | Should Be $false
        (& $validOwner $streamCandidates $expectedStreamPath ([pscustomobject]@{Module=$stream})) | Should Be $false
        (& $validOwner @($stream,$stream) $expectedStreamPath $export) | Should Be $false
        $sandbox = New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness d -Case ('git-'+[guid]::NewGuid().ToString('N').Substring(0,8))
        [ordered]@{strict_entry='Latest';canonical_candidates=$streamCandidates.Count;export_owner_reference_match=[object]::ReferenceEquals($export.Module,$stream);
            missing_lookalike_wrongpath_wrongowner_proxy_ambiguity_rejected=$true} |
            ConvertTo-Json | Set-Content (Join-Path $sandbox 'module-ownership.json') -Encoding utf8
        $origin = Join-Path $sandbox 'origin 雪'
        [void][IO.Directory]::CreateDirectory((Join-Path $origin 'include'))
        $priorRecord = $env:RS_DXUI_CAPTURE_RECORD
        $priorModule = $env:RS_DXUI_CAPTURE_STARTER
        $priorSentinel = $env:RS_DXUI_CAPTURE_SENTINEL
        try {
            $env:RS_DXUI_CAPTURE_RECORD = Join-Path $sandbox 'launches.jsonl'
            $env:RS_DXUI_CAPTURE_STARTER = Join-Path $repoRoot 'Tools/Modules/Build/SanitizedEnvironment.psm1'
            $env:RS_DXUI_CAPTURE_SENTINEL = 'literal inherited 雪'
            Mock Start-RSContainedProcess -ModuleName ProcessStreaming {
                param($ProcessStartInfo,[switch]$DelegateArtifactOperation)
                if($ProcessStartInfo.UseShellExecute -or -not $ProcessStartInfo.CreateNoWindow -or
                    -not $ProcessStartInfo.RedirectStandardOutput -or -not $ProcessStartInfo.RedirectStandardError -or $DelegateArtifactOperation) {
                    throw 'Refusing unsafe DxUi Git fixture launch.'
                }
                if ($null -eq $ProcessStartInfo.StandardOutputEncoding -or $ProcessStartInfo.StandardOutputEncoding.CodePage -ne 65001 -or $null -ne $ProcessStartInfo.StandardErrorEncoding) { throw 'Git stdout must be UTF8 without changing stderr decoding.' }
                # Resolve the already-loaded canonical module without importing it
                # over this module-scoped interception after the first child.
                $starterModule = @(Get-Module -All | Where-Object { $_.Path -ceq $env:RS_DXUI_CAPTURE_STARTER })
                if ($starterModule.Count -ne 1) { throw 'Expected one canonical contained-process module.' }
                $starter = $starterModule[0].ExportedFunctions['Start-RSContainedProcess']
                $owned = & $starter -ProcessStartInfo $ProcessStartInfo
                $transferred = $false
                try {
                    [ordered]@{process=$owned.Id;exact_job=$owned.IsInContainmentJob;file=$ProcessStartInfo.FileName;
                        cwd=$ProcessStartInfo.WorkingDirectory;arguments=@($ProcessStartInfo.ArgumentList);no_window=$ProcessStartInfo.CreateNoWindow;
                        stdout=$ProcessStartInfo.RedirectStandardOutput;stderr=$ProcessStartInfo.RedirectStandardError;delegated=[bool]$DelegateArtifactOperation;
                        inherited_sentinel=$ProcessStartInfo.Environment['RS_DXUI_CAPTURE_SENTINEL']} |
                        ConvertTo-Json -Compress | Add-Content -LiteralPath $env:RS_DXUI_CAPTURE_RECORD -Encoding utf8
                    $transferred = $true
                    return $owned
                } finally { if(-not $transferred){$owned.Dispose()} }
            }
            # Assert the hook is in the actual module before any real fixture child.
            (& $stream { (Get-Command Start-RSContainedProcess).Definition }) | Should Match 'Invoke-Mock'
            $unsafe = [Diagnostics.ProcessStartInfo]::new()
            $unsafe.FileName = (Get-Command git -CommandType Application | Select-Object -First 1).Source
            $unsafe.UseShellExecute = $false
            $refused = $false
            try { & $stream {param($Psi) Start-RSContainedProcess -ProcessStartInfo $Psi} $unsafe | Out-Null }
            catch { $refused = $_.Exception.Message -like '*Refusing unsafe DxUi Git fixture*' }
            $refused | Should Be $true
            (Test-Path $env:RS_DXUI_CAPTURE_RECORD) | Should Be $false
            '// synthetic header' | Set-Content (Join-Path $origin 'include/Example.h') -Encoding utf8
            '// synthetic Unicode-name fixture' | Set-Content (Join-Path $origin 'include/資料 雪.sln') -Encoding utf8
            Invoke-RSDxUiTestGit -c core.longpaths=true init --quiet $origin
            Invoke-RSDxUiTestGit -C $origin add .
            Invoke-RSDxUiTestGit -C $origin -c user.name=Fixture -c user.email=fixture@example.invalid commit --quiet -m 'Unicode 雪 fixture'
            $LASTEXITCODE | Should Be 0
            $unicode = & $script:dxUiTestModule {param($Root) Invoke-RSDxUiSourceGitCapture -WorkingDirectory $Root -Arguments @('-C',$Root,'log','-1','--format=%s')} $origin
            $unicode.ExitCode | Should Be 0
            ($unicode.StdOut -join "`n") | Should Be 'Unicode 雪 fixture'
            $savedConsoleEncoding=[Console]::OutputEncoding
            $unicodeObservations=[Collections.Generic.List[object]]::new()
            try {
                foreach($encoding in @([Text.Encoding]::GetEncoding(437),[Text.UTF8Encoding]::new($false))) {
                    [Console]::OutputEncoding=$encoding
                    $paths = & $script:dxUiTestModule {param($Root) Invoke-RSDxUiSourceGitCapture -WorkingDirectory $Root -Arguments @('-c','core.quotepath=false','-C',$Root,'ls-files')} $origin
                    $paths.ExitCode | Should Be 0
                    (@($paths.StdOut) -ccontains 'include/資料 雪.sln') | Should Be $true
                    [Console]::OutputEncoding.CodePage | Should Be $encoding.CodePage
                    $unicodeObservations.Add(@{parentCodePage=$encoding.CodePage;stdout=@($paths.StdOut);exit=$paths.ExitCode})
                }
            } finally { [Console]::OutputEncoding=$savedConsoleEncoding }
            $unicodeObservations | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $sandbox 'git-stdout-unicode-codepages.json') -Encoding utf8
            $failure = & $script:dxUiTestModule {param($Root) Invoke-RSDxUiSourceGitCapture -WorkingDirectory $Root -Arguments @('-C',$Root,'rev-parse','HEAD:missing')} $origin
            ($failure.ExitCode -ne 0) | Should Be $true
            @($failure.StdErr).Count | Should BeGreaterThan 0
            $env:RS_DXUI_CAPTURE_SENTINEL | Should Be 'literal inherited 雪'
            $rows = @(Get-Content $env:RS_DXUI_CAPTURE_RECORD | ForEach-Object { $_ | ConvertFrom-Json })
            $rows.Count | Should Be 7
            @($rows | Where-Object { -not $_.no_window -or -not $_.stdout -or -not $_.stderr -or -not $_.exact_job -or $_.delegated }).Count | Should Be 0
            @($rows | Where-Object { $_.inherited_sentinel -cne 'literal inherited 雪' }).Count | Should Be 0
            $rows | ConvertTo-Json -Depth 7 | Set-Content (Join-Path $sandbox 'safe-git-observations.json') -Encoding utf8
        } finally {
            $env:RS_DXUI_CAPTURE_RECORD=$priorRecord
            $env:RS_DXUI_CAPTURE_STARTER=$priorModule
            $env:RS_DXUI_CAPTURE_SENTINEL=$priorSentinel
        }
    }
}

Describe 'SDK consumer Windows header policy' {
    BeforeEach { $script:dxUiMetadataAuditLease = $null; $script:dxUiMetadataAuditLease = Start-RSDxUiMetadataAudit }
    AfterEach { if($null -ne $script:dxUiMetadataAuditLease){Stop-RSDxUiMetadataAudit -State $script:dxUiMetadataAuditLease} }
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
        $installationCapture = Invoke-RSDxUiTestMetadataCapture -FilePath $vswhere -Arguments @('-latest','-prerelease','-products','*','-requires','Microsoft.Component.MSBuild','-property','installationPath')
        if($installationCapture.ExitCode -ne 0){throw ('Cannot locate the metadata fixture toolchain: '+($installationCapture.StdErr -join [Environment]::NewLine))}
        $installation = ($installationCapture.StdOut -join [Environment]::NewLine).Trim()
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
            $json = Invoke-RSDxUiTestMetadataCapture -FilePath $msbuild -Arguments @((Join-Path $repoRoot 'Tests/RedConfigureTests/RedConfigureTests.vcxproj'), '/nologo', '/p:Configuration=Release', "/p:Platform=$platform", "/p:RSDxUiSdkRoot=$sdk", "/p:RSCiEngineRoot=$(Join-Path $repoRoot '.ci-engine')", '-getItem:ClCompile,ProjectReference')
            if ($json.ExitCode -ne 0) {throw "Cannot evaluate SDK Windows macro policy for $platform."}
            $evaluated = ($json.StdOut -join "`n") | ConvertFrom-Json
            @($evaluated.Items.ClCompile).Count | Should BeGreaterThan 0
            foreach ($item in $evaluated.Items.ClCompile) {
                $item.PreprocessorDefinitions | Should Not Match '(^|;)(WIN32_LEAN_AND_MEAN|NOMINMAX)(=1)?(;|$)'
                $item.PreprocessorDefinitions | Should Match '(^|;)WIN32_LEAN_AND_MEAN=(;|$)'
                $item.PreprocessorDefinitions | Should Match '(^|;)NOMINMAX=(;|$)'
                $item.PreprocessorDefinitions | Should Match '(^|;)NOMINMAX_FEATURE_CONTROL(;|$)'
                $item.PreprocessorDefinitions | Should Match '(^|;)NDEBUG(;|$)'
                $item.PreprocessorDefinitions | Should Match '(^|;)UNICODE(;|$)'
            }
            @($evaluated.Items.ProjectReference | Where-Object {[IO.Path]::GetFileName($_.Identity) -ceq 'DxUi.vcxproj'}).Count | Should Be 0
            # A getItem string hides escaped delimiters. Exercise an actual
            # MSBuild task's list binding, as the compiler task consumes it.
            $lines = Join-Path $sdk ("definitions-$platform.txt")
            $probeOutput = Invoke-RSDxUiTestMetadataCapture -FilePath $msbuild -Arguments @((Join-Path $repoRoot 'Tests/RedConfigureTests/RedConfigureTests.vcxproj'), '/nologo', '/p:Configuration=Release', "/p:Platform=$platform", "/p:RSDxUiSdkRoot=$sdk", "/p:RSCiEngineRoot=$(Join-Path $repoRoot '.ci-engine')", "/p:ForceImportAfterCppTargets=$taskProbe", "/p:I61DefinitionLines=$lines", '/t:I61DefinitionListProbe')
            if ($probeOutput.ExitCode -ne 0) {throw "Cannot probe actual MSBuild definition list binding for $platform`: $((@($probeOutput.StdOut)+@($probeOutput.StdErr)) -join [Environment]::NewLine)"}
            $definitions = @(Get-Content -LiteralPath $lines)
            $definitions.Count | Should BeGreaterThan 4
            foreach ($definition in $definitions) {$definition | Should Not Match ';'}
            @($definitions | Where-Object { $_ -ceq 'WIN32_LEAN_AND_MEAN=' }).Count | Should Be 1
            @($definitions | Where-Object { $_ -ceq 'NOMINMAX=' }).Count | Should Be 1
            ($definitions -ccontains 'NOMINMAX_FEATURE_CONTROL') | Should Be $true
            ($definitions -ccontains 'NDEBUG') | Should Be $true
            ($definitions -ccontains 'UNICODE') | Should Be $true
        }
        # Metadata evaluation above grants no build/ABI verdict. Actual SDK builds
        # must still execute the shared live-reader and per-TU admission targets.
        [IO.File]::ReadAllText((Join-Path $repoRoot 'Build/RedSalamander.DxUi.props')) | Should Match 'Ci\.SdkConsumer\.targets'
        Assert-RSDxUiMetadataAudit -State $script:dxUiMetadataAudit -ExpectedMSBuild 4
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
            $tracked = @(Invoke-RSDxUiTestGit -C $repoRoot ls-files -- $relative)
            if ($LASTEXITCODE -ne 0) { throw "Cannot enumerate tracked legacy path: $relative" }
            $tracked.Count | Should Be 0
        }
        $solution = Get-Content -LiteralPath (Join-Path $repoRoot 'RedSalamander.sln') -Raw
        $solution | Should Not Match 'Common[\\/]DxUi|Tests[\\/]DxUiTests'
        $solution | Should Match 'ProductUiTests'
    }

    It 'keeps owned project inputs separate from the restored library implementation' {
        $projects = @(Invoke-RSDxUiTestGit -C $repoRoot ls-files -- '*.vcxproj')
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
        Invoke-RSDxUiTestGit init --quiet $origin
        Invoke-RSDxUiTestGit -C $origin add .
        Invoke-RSDxUiTestGit -C $origin -c user.name=Fixture -c user.email=fixture@example.invalid commit --quiet -m fixture
        if ($LASTEXITCODE -ne 0) { throw 'Cannot initialize isolated source fixture.' }
        $head = (Invoke-RSDxUiTestGit -C $origin rev-parse HEAD).Trim()
        @{repository='https://github.com/RedSalamanders/DxUi';commit=$head;apiRevision=3;targets=@('DxUi')} |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path $root 'Dependencies/DxUi.lock.json')
        $rejected=$false; try { Get-RSDxUiSourceIdentity -RepoRoot $root | Out-Null } catch { $rejected=$true }; $rejected | Should Be $true
        Get-RSDxUiSourceIdentity -RepoRoot $root -AllowMissing | Should BeNullOrEmpty
        $source = Join-Path $root ".build/dependencies/DxUi/source/$($head.Substring(0,16))"
        [void](New-Item -ItemType Directory -Path (Split-Path $source) -Force)
        Invoke-RSDxUiTestGit -c core.longpaths=true clone --config core.longpaths=true --quiet --local $origin $source
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
        Invoke-RSDxUiTestGit init --quiet $origin
        Invoke-RSDxUiTestGit -C $origin add .
        Invoke-RSDxUiTestGit -C $origin -c user.name=Fixture -c user.email=fixture@example.invalid commit --quiet -m fixture
        if ($LASTEXITCODE -ne 0) { throw 'Cannot initialize isolated provenance fixture.' }
        $head = (Invoke-RSDxUiTestGit -C $origin rev-parse HEAD).Trim()
        @{repository='https://github.com/RedSalamanders/DxUi';commit=$head;apiRevision=3;targets=@('DxUi')} |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path $root 'Dependencies/DxUi.lock.json')
        $dependency = Join-Path $root '.build/dependencies/DxUi'
        $source = Join-Path $dependency "source/$($head.Substring(0,16))"
        [void](New-Item -ItemType Directory -Path (Split-Path $source) -Force)
        Invoke-RSDxUiTestGit -c core.longpaths=true clone --config core.longpaths=true --quiet --local $origin $source
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
            'vcpkg-tool.json' = '{ "repository": "https://github.com/microsoft/vcpkg", "commit": "0000000000000000000000000000000000000000" }'
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
        Invoke-RSDxUiTestGit -c core.longpaths=true init --quiet $Path
        if ($LASTEXITCODE -ne 0) { throw 'Cannot initialize the restore fixture.' }
        Invoke-RSDxUiTestGit -C $Path config core.longpaths true
        if ($LASTEXITCODE -ne 0) { throw 'Cannot initialize the restore fixture.' }
        Invoke-RSDxUiTestGit -C $Path add .
        if ($LASTEXITCODE -ne 0) { throw 'Cannot initialize the restore fixture.' }
        Invoke-RSDxUiTestGit -C $Path -c user.name=Fixture -c user.email=fixture@example.invalid commit --quiet -m fixture
        if ($LASTEXITCODE -ne 0) { throw 'Cannot initialize the restore fixture.' }
        return (Invoke-RSDxUiTestGit -C $Path rev-parse HEAD).Trim()
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
        foreach ($kept in @('capabilities.json', 'vcpkg-install.ps1', 'vcpkg-tool.json', 'Tools/ConsumerBuild.psm1', 'src/DxUi.vcxproj', 'include/DxUi/Example.h', 'docs/README.md')) {
            Test-Path -LiteralPath (Join-Path $source $kept) | Should Be $true
        }
        # Long paths live in the checkout's own configuration; no user or global configuration is involved.
        (Invoke-RSDxUiTestGit -C $source config --local --get core.longpaths) | Should Be 'true'
        # The omitted paths are not deletions: the restore's own clean-source check accepts the sparse checkout.
        (Get-RSDxUiSourceIdentity -RepoRoot $root).Pin.commit | Should Be $head
        (Invoke-RSDxUiTestGit -C $source rev-parse HEAD).Trim() | Should Be $head
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
    BeforeEach { $script:dxUiMetadataAuditLease = $null; $script:dxUiMetadataAuditLease = Start-RSDxUiMetadataAudit }
    AfterEach { if($null -ne $script:dxUiMetadataAuditLease){Stop-RSDxUiMetadataAudit -State $script:dxUiMetadataAuditLease} }
    It 'uses the first-party host during restore and restores the caller environment after success or failure' {
        $root = New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness d -Case ('host-' + [guid]::NewGuid().ToString('N').Substring(0,8))
        # Exercise a genuinely long Git metadata path in both focused and deeper Full sandboxes.
        # Size the padding from the shortened source locator rather than the old 40-character folder.
        $metadataSuffix = '.build/dependencies/DxUi/source/' + ('0' * 16) + '/.git/hooks/fsmonitor-watchman.sample'
        $metadataPath = Join-Path $root $metadataSuffix
        # The contract is long Git metadata, not a >260-character Git -C directory.
        if ($metadataPath.Length -lt 270) { $root = Join-Path $root ('x' * (270 - $metadataPath.Length)) }
        $predictedSource = Join-Path $root ('.build/dependencies/DxUi/source/' + ('0' * 16))
        $metadataPath = Join-Path $root $metadataSuffix
        if ($predictedSource.Length -gt 240) { throw 'Compiler-host fixture exceeds the safe Git -C directory budget (240).' }
        $metadataPath.Length | Should BeGreaterThan 260
        $budgetRecord = Assert-RSTestSandboxContainedPath -TestRoot $script:dxUiMetadataAudit.TestRoot `
            -Path (Join-Path $script:dxUiMetadataAudit.Directory 'fixture-path-budget.json') -RequireDescendant
        @{source_length=$predictedSource.Length;source_limit=240;metadata_length=$metadataPath.Length;metadata_minimum_exclusive=260;
            placement='owned shallow source with long metadata; not Git -C>=265 support'} |
            ConvertTo-Json | Set-Content $budgetRecord -Encoding utf8
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
        Invoke-RSDxUiTestGit -c core.longpaths=true init --quiet $origin
        if ($LASTEXITCODE -ne 0) { throw 'Cannot initialize long-path compiler-host fixture.' }
        Invoke-RSDxUiTestGit -c core.longpaths=true -C $origin config core.longpaths true
        if ($LASTEXITCODE -ne 0) { throw 'Cannot configure long-path compiler-host fixture.' }
        Invoke-RSDxUiTestGit -c core.longpaths=true -C $origin add .
        if ($LASTEXITCODE -ne 0) { throw 'Cannot stage long-path compiler-host fixture.' }
        Invoke-RSDxUiTestGit -c core.longpaths=true -C $origin -c user.name=Fixture -c user.email=fixture@example.invalid commit --quiet -m fixture
        if ($LASTEXITCODE -ne 0) { throw 'Cannot initialize compiler-host fixture.' }
        $head = (Invoke-RSDxUiTestGit -C $origin rev-parse HEAD).Trim()
        @{repository='https://github.com/RedSalamanders/DxUi';commit=$head;apiRevision=3;targets=@('DxUi')} |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path $root 'Dependencies/DxUi.lock.json')
        $dependency = Join-Path $root '.build/dependencies/DxUi'
        $source = Join-Path $dependency "source/$($head.Substring(0,16))"
        [void](New-Item -ItemType Directory -Path (Split-Path $source) -Force)
        (Join-Path $source '.git/hooks/fsmonitor-watchman.sample').Length | Should BeGreaterThan 260
        # Match the canonical restore's two phases: initialize the long-path
        # repository before explicitly selecting its worktree and pinned commit.
        # Copy fixture objects rather than retaining links to the origin.
        Invoke-RSDxUiTestGit -c core.longpaths=true clone --config core.longpaths=true --quiet --local --no-hardlinks --no-checkout $origin $source
        if ($LASTEXITCODE -ne 0) { throw 'Cannot clone compiler-host fixture.' }
        Invoke-RSDxUiTestGit -c core.longpaths=true -C $source checkout --quiet --detach $head
        if ($LASTEXITCODE -ne 0) { throw 'Cannot check out compiler-host fixture.' }
        (Invoke-RSDxUiTestGit -c core.longpaths=true -C $source rev-parse HEAD).Trim() | Should Be $head
        Mock Invoke-RSDxUiRestoreProcess {} -ModuleName DxUiDependency
        $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
        $installationCapture = Invoke-RSDxUiTestMetadataCapture -FilePath $vswhere -Arguments @('-latest','-prerelease','-products','*','-requires','Microsoft.Component.MSBuild','-property','installationPath')
        if($installationCapture.ExitCode -ne 0){throw ('Cannot locate the metadata fixture toolchain: '+($installationCapture.StdErr -join [Environment]::NewLine))}
        $installation = ($installationCapture.StdOut -join [Environment]::NewLine).Trim()
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
        Assert-RSDxUiMetadataAudit -State $script:dxUiMetadataAudit -ExpectedMSBuild 15
    }

    It 'evaluates the six real consumer profiles with sanitizer settings and a 64-bit cross-compiler host' {
        $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
        $installationCapture = Invoke-RSDxUiTestMetadataCapture -FilePath $vswhere -Arguments @('-latest','-prerelease','-products','*','-requires','Microsoft.Component.MSBuild','-property','installationPath')
        if($installationCapture.ExitCode -ne 0){throw ('Cannot locate the metadata fixture toolchain: '+($installationCapture.StdErr -join [Environment]::NewLine))}
        $installation = ($installationCapture.StdOut -join [Environment]::NewLine).Trim()
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
                    $json = Invoke-RSDxUiTestMetadataCapture -FilePath $msbuild -Arguments @((Join-Path $repoRoot 'Tests/RedConfigureTests/RedConfigureTests.vcxproj'), '/nologo', "/p:Configuration=$configuration", "/p:Platform=$platform", "/p:DxUiRoot=$dxuiFixture", '-getProperty:EnableASAN,PreferredToolArchitecture', '-getItem:ClCompile,ProjectReference')
                    if ($json.ExitCode -ne 0) { throw "Cannot evaluate consumer profile $platform/$configuration." }
                    $evaluated = ($json.StdOut -join "`n") | ConvertFrom-Json
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
                    $explicitHost = Invoke-RSDxUiTestMetadataCapture -FilePath $msbuild -Arguments @((Join-Path $repoRoot 'Tests/RedConfigureTests/RedConfigureTests.vcxproj'), '/nologo', "/p:Configuration=$configuration", "/p:Platform=$platform", "/p:DxUiRoot=$dxuiFixture", '/p:PreferredToolArchitecture=x86', '-getProperty:PreferredToolArchitecture', '-getItem:ProjectReference')
                    if ($explicitHost.ExitCode -ne 0) { throw "Cannot evaluate explicit compiler host for $platform/$configuration." }
                    $explicitEvaluation = ($explicitHost.StdOut -join "`n") | ConvertFrom-Json
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
        Assert-RSDxUiMetadataAudit -State $script:dxUiMetadataAudit -ExpectedMSBuild 24
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
param($Phase,$Configuration,$Platform,$Available,$Probe,$Pin,$FixtureRoot)
$root=Join-Path $FixtureRoot 'Synthetic workspace';$EngineRoot=Join-Path $FixtureRoot 'Controlled engine';$context=@{fixture=$true};$number='123'
function Import-Module {param($Name) if(-not($Name.EndsWith('Ci.SdkPreparation.psm1') -or $Name.EndsWith('DxUiDependency.psm1'))){throw 'Unexpected SDK module'};$Probe.imports++}
function Get-RSDxUiPin {param($RepoRoot) return $Pin}
function Initialize-CiPinnedSdkInput {param($Context,$SourceCommit,$ApiRevision) $Probe.prepares++;$Probe.source=$SourceCommit;$Probe.api=$ApiRevision;if($Available){Join-Path $root '.build/input descriptor.json'}}
function Run {param($Script,$CommandArguments) $Probe.commands.Add(@{script=$Script;arguments=@($CommandArguments)})}
'@ + "`n" + $normalize[0].Extent.Text + "`n" + '$profile=@(''-Platform'',$Platform,''-Configuration'',$Configuration);$sdkBuildArguments=@()' + "`n" + $prepare[0].Extent.Text + "`n" +
            'if($Phase -ceq ''Production''){' + $production[0].Extent.Text + '}' + "`n" +
            'if($Phase -cin @(''ReleaseTestBuild'',''Qualification'')){' + $testBuild[0].Extent.Text + '}')
        $pin=Get-Content (Join-Path $repoRoot 'Dependencies/DxUi.lock.json') -Raw|ConvertFrom-Json
        $pin.commit | Should Be '7fb2ab71852087d93b3a3e5b3ba504c37f56267c'
        $pin.apiRevision | Should Be 4
        $count=0
        foreach($phase in @('Static','ReleaseTestBuild','ReleaseTest','Production','Qualification')){
            foreach($configuration in @('Debug','Release','ASan Debug')){foreach($platform in @('x64','ARM64')){foreach($available in @($true,$false)){
                $probe=@{imports=0;prepares=0;source='';api=0;commands=[Collections.Generic.List[object]]::new()}
                & $stage $phase $configuration $platform $available $probe $pin $TestDrive
                $eligible=$phase -cin @('ReleaseTestBuild','Production')
                $probe.prepares | Should Be ([int]$eligible)
                $probe.imports | Should Be $(if($eligible){2}else{0})
                if($eligible){$probe.source|Should Be $pin.commit;$probe.api|Should Be $pin.apiRevision}
                $probe.commands.Count | Should Be $(if($phase -cin @('ReleaseTestBuild','Production','Qualification')){1}else{0})
                foreach($command in $probe.commands){
                    $hasDescriptor='-DxUiSdkInputDescriptor' -cin $command.arguments
                    $hasDescriptor | Should Be ($eligible -and $available)
                    if($hasDescriptor){((Join-Path (Join-Path $TestDrive 'Synthetic workspace') '.build/input descriptor.json') -cin $command.arguments)|Should Be $true}
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

    It 'uses runtime-only vcpkg binding and preserves inherited compiler overrides' {
        $tokens=$null;$errors=$null
        $adapter=Join-Path $repoRoot 'Tools/Invoke-ProjectCi.ps1'
        $ast=[Management.Automation.Language.Parser]::ParseFile($adapter,[ref]$tokens,[ref]$errors)
        @($errors).Count | Should Be 0
        $runtime=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -ceq 'Get-RSVcpkgBuildEnvironment' -and @($n.CommandElements | Where-Object {$_ -is [Management.Automation.Language.CommandParameterAst] -and $_.ParameterName -ceq 'RuntimeOnly'}).Count -eq 1},$true))
        $prepare=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -ceq 'Initialize-CiPinnedSdkInput'},$true))
        $apply=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.ForEachStatementAst] -and $n.Extent.Text.Contains('$environment.GetEnumerator()')},$true))
        $runtime.Count|Should Be 1
        $prepare.Count|Should Be 1
        $apply.Count|Should Be 1
        ($runtime[0].Extent.StartOffset -lt $prepare[0].Extent.StartOffset)|Should Be $true
        ($prepare[0].Extent.EndOffset -lt $apply[0].Extent.StartOffset)|Should Be $true
        $body=[scriptblock]::Create($apply[0].Extent.Text)
        $runtimeModule=Import-Module (Join-Path $repoRoot 'Tools/Modules/Build/VcpkgInstallSafety.psm1') -PassThru -ErrorAction Stop
        $fixtureRoot=Join-Path $TestDrive 'runtime-only-vcpkg'
        $toolRoot=Join-Path $fixtureRoot '.build/vcpkg-tool'
        $installed=Join-Path $fixtureRoot '.build/vcpkg_installed/x64/x64-windows'
        $include=Join-Path $installed 'include';$lib=Join-Path $installed 'lib';$bin=Join-Path $installed 'bin'
        foreach($directory in @($include,$lib,$bin)){[void][IO.Directory]::CreateDirectory($directory)}
        $savedPath=$env:Path;$savedCL=$env:CL;$savedLINK=$env:LINK;$savedForceBefore=$env:ForceImportBeforeCppTargets;$savedForceAfter=$env:ForceImportAfterCppTargets
        try{
            $env:Path='inherited runtime path'
            $env:CL='inherited override remains visible'
            $env:LINK='inherited linker override remains visible'
            $env:ForceImportBeforeCppTargets='inherited before import remains visible'
            $env:ForceImportAfterCppTargets='inherited after import remains visible'
            $environment=& $runtimeModule { param($Root,$ToolRoot) Get-RSVcpkgBuildEnvironment -RepoRoot $Root -ToolRoot $ToolRoot -Platform x64 -Configuration Release -RuntimeOnly } $fixtureRoot $toolRoot
            @($environment.Keys).Count|Should Be 1
            @($environment.Keys)[0]|Should Be 'Path'
            $expectedPath=$bin+[IO.Path]::PathSeparator+'inherited runtime path'
            $environment.Path|Should Be $expectedPath
            $root=$fixtureRoot;$Platform='x64';$Configuration='Release'
            . $body
            $env:Path|Should Be $expectedPath
            $env:CL|Should Be 'inherited override remains visible'
            $env:LINK|Should Be 'inherited linker override remains visible'
            $env:ForceImportBeforeCppTargets|Should Be 'inherited before import remains visible'
            $env:ForceImportAfterCppTargets|Should Be 'inherited after import remains visible'
        }finally{$env:Path=$savedPath;$env:CL=$savedCL;$env:LINK=$savedLINK;$env:ForceImportBeforeCppTargets=$savedForceBefore;$env:ForceImportAfterCppTargets=$savedForceAfter}
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


Describe 'Captured DxUi metadata probes' {
    It 'does not grant Git recording authority to an inherited ambient path' {
        $root = Join-Path $TestDrive 'metadata-ambient-record'
        [void][IO.Directory]::CreateDirectory($root)
        $sentinel = Join-Path $root 'owned-sentinel'
        'unchanged' | Set-Content $sentinel -Encoding utf8
        Mock Invoke-RSStreamingProcess -ModuleName ProcessStreaming { return 0 } -ParameterFilter { $env:RS_DXUI_METADATA_AMBIENT_TEST -ceq '1' }
        $previousRecord = $env:RS_DXUI_METADATA_RECORD
        $previousAudit = $script:dxUiMetadataAudit
        $previousFixtureFlag = $env:RS_DXUI_METADATA_AMBIENT_TEST
        try {
            $env:RS_DXUI_METADATA_AMBIENT_TEST = '1'
            $env:RS_DXUI_METADATA_RECORD = $sentinel
            foreach($audit in @($null,[pscustomobject]@{Active=$false})) {
                $script:dxUiMetadataAudit = $audit
                Invoke-RSDxUiTestGit --version | Out-Null
                (Get-Content $sentinel -Raw).Trim() | Should Be 'unchanged'
                (Test-Path ($sentinel + '.git-captures.jsonl')) | Should Be $false
            }
        } finally {
            $script:dxUiMetadataAudit = $previousAudit
            $env:RS_DXUI_METADATA_RECORD = $previousRecord
            $env:RS_DXUI_METADATA_AMBIENT_TEST = $previousFixtureFlag
        }
    }

    It 'refuses unsafe metadata PSI before the actual starter receives a child' {
        $state = Start-RSDxUiMetadataAudit
        try {
            $stream = & $script:dxUiTestModule { Get-Module -Name ProcessStreaming }
            (& $stream { (Get-Command Start-RSContainedProcess).Definition }) | Should Match 'Invoke-Mock'
            $unsafe = [Diagnostics.ProcessStartInfo]::new()
            # A nonexistent file also prevents a child if the hook were absent.
            $unsafe.FileName = Join-Path $state.Directory 'must-not-start.exe'
            $unsafe.UseShellExecute = $false
            $message = ''
            try { & $stream {param($Psi) Start-RSContainedProcess -ProcessStartInfo $Psi} $unsafe | Out-Null }
            catch { $message = $_.Exception.Message }
            $message | Should Match 'Refusing unsafe DxUi metadata fixture launch'
            (Test-Path $env:RS_DXUI_METADATA_RECORD) | Should Be $false
        } finally { Stop-RSDxUiMetadataAudit -State $state }
    }

    It 'keeps compiler host <HostValue> arguments, CWD, stderr and parent environment independent' -TestCases @(
        @{HostValue='x86'},@{HostValue='x64'},@{HostValue='arm64'}
    ) {
        param($HostValue)
        $root=Join-Path $TestDrive ('metadata-host-'+$HostValue)
        [void][IO.Directory]::CreateDirectory($root)
        $HostValue|Set-Content (Join-Path $root 'metadata-mode')
        Mock Invoke-RSStreamingProcess -ModuleName ProcessStreaming {
            param($FilePath,$Arguments,$WorkingDirectory,$OutputLineCallback,[switch]$DelegateArtifactOperation)
            @{file=$FilePath;arguments=@($Arguments);cwd=$WorkingDirectory;delegated=[bool]$DelegateArtifactOperation;parent_host=$env:PreferredToolArchitecture} |
                ConvertTo-Json -Depth 5|Set-Content (Join-Path $WorkingDirectory 'metadata-record.json') -Encoding utf8
            & $OutputLineCallback (' '+(Get-Content (Join-Path $WorkingDirectory 'metadata-mode') -Raw).Trim()+' ') $false
            & $OutputLineCallback 'stderr is not a host value' $true
            return 0
        } -ParameterFilter { Test-Path -LiteralPath (Join-Path $WorkingDirectory 'metadata-mode') }
        $previousHost=$env:PreferredToolArchitecture;$previousExit=$global:LASTEXITCODE
        try {
            $env:PreferredToolArchitecture='literal & inherited';$global:LASTEXITCODE=73
            $actual=&$script:dxUiTestModule {param($r) Get-RSDxUiConsumerCompilerHost -RepoRoot $r -MSBuildPath 'C:\fixture\MSBuild.exe' -Platform ARM64 -Configuration 'ASan Debug'} $root
            $actual|Should Be $HostValue
            $env:PreferredToolArchitecture|Should Be 'literal & inherited'
            $record=Get-Content (Join-Path $root 'metadata-record.json') -Raw|ConvertFrom-Json
            $record.file|Should Be 'C:\fixture\MSBuild.exe'
            $record.cwd|Should Be $root
            ($record.arguments -join '|')|Should Be ((Join-Path $root 'RedSalamander/RedSalamander.vcxproj')+'|/nologo|/p:Configuration=ASan Debug|/p:Platform=ARM64|-getProperty:PreferredToolArchitecture')
            $record.delegated|Should Be $false
            $record.parent_host|Should Be 'literal & inherited'
        } finally {$env:PreferredToolArchitecture=$previousHost;$global:LASTEXITCODE=$previousExit}
    }

    It 'refuses metadata failure <Mode> with its actual exit and supported-host checks' -TestCases @(
        @{Mode='nonzero'},@{Mode='unsupported'}
    ) {
        param($Mode)
        $root=Join-Path $TestDrive ('metadata-refusal-'+$Mode)
        [void][IO.Directory]::CreateDirectory($root)
        $Mode|Set-Content (Join-Path $root 'metadata-failure')
        Mock Invoke-RSStreamingProcess -ModuleName ProcessStreaming {
            param($WorkingDirectory,$OutputLineCallback)
            $mode=(Get-Content (Join-Path $WorkingDirectory 'metadata-failure') -Raw).Trim()
            & $OutputLineCallback ($(if($mode -eq 'nonzero'){'x64'}else{'foreign-host'})) $false
            & $OutputLineCallback 'metadata failure detail' $true
            if($mode -eq 'nonzero'){return 17}else{return 0}
        } -ParameterFilter { Test-Path -LiteralPath (Join-Path $WorkingDirectory 'metadata-failure') }
        $message=''
        try { &$script:dxUiTestModule {param($r) Get-RSDxUiConsumerCompilerHost -RepoRoot $r -MSBuildPath 'C:\fixture\MSBuild.exe' -Platform x64 -Configuration Debug} $root|Out-Null }
        catch {$message=$_.Exception.Message}
        $message|Should Match ($(if($Mode -eq 'nonzero'){'(?s)Cannot evaluate the consumer compiler host.*metadata failure detail'}else{'Unsupported consumer compiler host: foreign-host'}))
    }

    It 'preserves JSON stdout separately from stderr and retains the exact nonzero exit' {
        $root=Join-Path $TestDrive 'metadata-json'
        [void][IO.Directory]::CreateDirectory($root)
        Mock Invoke-RSStreamingProcess -ModuleName ProcessStreaming {
            param($OutputLineCallback)
            & $OutputLineCallback '{"Properties":{"Value":"Unicode 雪"}}' $false
            & $OutputLineCallback 'not JSON: diagnostic stderr' $true
            return 29
        } -ParameterFilter { $WorkingDirectory -like '*metadata-json' }
        $capture=Invoke-RSDxUiTestMetadataCapture -FilePath 'C:\fixture\vswhere.exe' -Arguments @('literal space','*') -WorkingDirectory $root
        $capture.ExitCode|Should Be 29
        (($capture.StdOut -join [Environment]::NewLine)|ConvertFrom-Json).Properties.Value|Should Be 'Unicode 雪'
        ($capture.StdErr -join '|')|Should Be 'not JSON: diagnostic stderr'
    }
}
