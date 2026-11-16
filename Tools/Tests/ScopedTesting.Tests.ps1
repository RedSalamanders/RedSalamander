Import-Module (Join-Path $PSScriptRoot 'TestSupport.psm1') -Force -ErrorAction Stop
$repoRoot=Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestRunPlan.psm1') -Force
Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/ScopedTesting.psm1') -Force
Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestSandbox.psm1') -Force
Describe 'Scoped testing and PR coverage contracts' {
    It 'validates real fragments before affected execution and keeps explanation non-executing' {
        $fixture=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('fragment-admission-'+[guid]::NewGuid().ToString('N'))
        $moduleDirectory=Join-Path $fixture 'Tools/Modules/Testing'
        $documentationDirectory=Join-Path $fixture 'Tools/Modules/Documentation'
        [void](New-Item -ItemType Directory -Path $moduleDirectory,$documentationDirectory,(Join-Path $fixture 'Changes') -Force)
        foreach($relative in @('Test-Changes.ps1','Tools/Test-Changelog.ps1','Tools/Modules/Documentation/Changelog.psm1')) {
            Copy-Item -LiteralPath (Join-Path $repoRoot $relative) -Destination (Join-Path $fixture $relative)
        }
        'function Assert-ScopedTestNames { param($Root) return 1 }; Export-ModuleMember -Function Assert-ScopedTestNames' |
            Set-Content -LiteralPath (Join-Path $moduleDirectory 'ScopedTesting.psm1')
        @'
param([switch]$ExplainPlan)
'selected' | Set-Content -LiteralPath (Join-Path $PSScriptRoot '../launched.txt')
$global:LASTEXITCODE=2
'@ | Set-Content -LiteralPath (Join-Path $fixture 'Tools/Run-AllTests.ps1')
        'Fragment instructions' | Set-Content -LiteralPath (Join-Path $fixture 'Changes/README.md')
        "# Changelog`n`n## Unreleased`n" | Set-Content -LiteralPath (Join-Path $fixture 'CHANGELOG.md')
        $fragment=Join-Path $fixture 'Changes/2026-11-09-example.md'
        '- Valid fixture change.' | Set-Content -LiteralPath $fragment
        $child=Join-Path $PSHOME 'pwsh.exe'
        Invoke-RSTestFixtureNative -FilePath $child -Arguments @('-NoProfile'; '-File'; (Join-Path $fixture 'Test-Changes.ps1')) *> (Join-Path $fixture 'valid.log')
        $LASTEXITCODE | Should Be 2
        $launched=Join-Path $fixture 'launched.txt'
        (Test-Path -LiteralPath $launched) | Should Be $true
        Remove-Item -LiteralPath $launched
        'Malformed fixture with no bullet.' | Set-Content -LiteralPath $fragment
        Invoke-RSTestFixtureNative -FilePath $child -Arguments @('-NoProfile'; '-File'; (Join-Path $fixture 'Test-Changes.ps1')) *> (Join-Path $fixture 'invalid.log')
        ($LASTEXITCODE -ne 0) | Should Be $true
        (Test-Path -LiteralPath $launched) | Should Be $false
        Invoke-RSTestFixtureNative -FilePath $child -Arguments @('-NoProfile'; '-File'; (Join-Path $fixture 'Test-Changes.ps1'); '-Explain') *> (Join-Path $fixture 'explain.log')
        $LASTEXITCODE | Should Be 2
        (Test-Path -LiteralPath $launched) | Should Be $true
        $global:LASTEXITCODE=0
    }
    It 'forwards explicit alternate-volume opt-in through Full and focused FileOps without enabling it by default' {
        $fixture=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('alternate-forwarding-'+[guid]::NewGuid().ToString('N'))
        $succeeded=$false
        try {
            $moduleDirectory=Join-Path $fixture 'Tools/Modules/Testing'
            [void](New-Item -ItemType Directory -Path $moduleDirectory -Force)
            Copy-Item -LiteralPath (Join-Path $repoRoot 'Test-Changes.ps1') -Destination (Join-Path $fixture 'Test-Changes.ps1')
            'param($Base)' | Set-Content -LiteralPath (Join-Path $fixture 'Tools/Test-Changelog.ps1')
            @'
function Assert-ScopedTestNames { param($Root) return 1 }
Export-ModuleMember -Function Assert-ScopedTestNames
'@ | Set-Content -LiteralPath (Join-Path $moduleDirectory 'ScopedTesting.psm1')
            @'
param($Suite,$Configuration,$Platform,$ImpactBase,[switch]$SkipBuild,[switch]$NonInteractive,$ValidationMode,$TestRoot,
      [switch]$AllowAlternateVolumeTestRoot,$CaseFilter='',$CommandsFamily='')
[ordered]@{suite=$Suite;mode=$ValidationMode;testRoot=$TestRoot;caseFilter=$CaseFilter;alternate=[bool]$AllowAlternateVolumeTestRoot} |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $PSScriptRoot '../captured.json')
$global:LASTEXITCODE=2
'@ | Set-Content -LiteralPath (Join-Path $fixture 'Tools/Run-AllTests.ps1')
            foreach($focused in @($false,$true)) {
                foreach($optIn in @($false,$true)) {
                    $arguments=@('-NoProfile','-File',(Join-Path $fixture 'Test-Changes.ps1'),'-TestRoot',$env:REDSALAMANDER_TEST_ROOT)
                    if($focused){$arguments+=@('-Suite','FileOps','-CaseFilter','FileOps_IndependentVolumes')}
                    else{$arguments+=@('-Mode','Full')}
                    if($optIn){$arguments+='-AllowAlternateVolumeTestRoot'}
                    Invoke-RSTestFixtureNative -FilePath (Join-Path $PSHOME 'pwsh.exe') -Arguments @($arguments) *> (Join-Path $fixture "child-$focused-$optIn.log")
                    $LASTEXITCODE | Should Be 2
                    $captured=Get-Content -LiteralPath (Join-Path $fixture 'captured.json') -Raw | ConvertFrom-Json
                    $captured.alternate | Should Be $optIn
                    $captured.suite | Should Be $(if($focused){'FileOps'}else{'Full'})
                    $captured.mode | Should Be 'Fresh'
                    $captured.caseFilter | Should Be $(if($focused){'FileOps_IndependentVolumes'}else{''})
                    $captured.testRoot | Should Be $env:REDSALAMANDER_TEST_ROOT
                }
            }
            $succeeded=$true
        } finally {
            if($succeeded) {
                [void](Assert-RSTestSandboxContainedPath -TestRoot $env:REDSALAMANDER_TEST_ROOT -Path $fixture -RequireDescendant)
                $refusals=@(& (Get-Module TestSandbox) { param($Path) Remove-RSTestSandboxTreeWithoutFollowingLinks -Path $Path } $fixture)
                if($refusals.Count){throw 'Scoped-wrapper fixture cleanup was refused.'}
            }
            $global:LASTEXITCODE=0
        }
    }
    It 'accounts for all active native sources with contextual test names' {(Assert-ScopedTestNames $repoRoot) | Should BeGreaterThan 0}
    It 'defers only complete equal PR contracts and retains changed arguments, writers and selftests' {
        $same=New-RSTestRunPlanEntry -Id standalone.example -Name Example -Kind Executable -Path (Join-Path $repoRoot '.build/x64/Debug/example.exe')
        $changed=New-RSTestRunPlanEntry -Id standalone.changed -Name Changed -Kind Executable -Path (Join-Path $repoRoot '.build/x64/Debug/example.exe') -Arguments @('--extra-scenario')
        $prChanged=New-RSTestRunPlanEntry -Id standalone.changed -Name Changed -Kind Executable -Path (Join-Path $repoRoot '.build/x64/Debug/example.exe')
        $selftest=New-RSTestRunPlanEntry -Id selftest.commands -Name Commands -Kind SelfTest -Path (Join-Path $repoRoot '.build/x64/Debug/app.exe')
        $writer=New-RSTestRunPlanEntry -Id tooling.pester-build-toolchain -Name Writer -Kind Pester -Path (Join-Path $repoRoot 'Tools/Tests')
        $local=@(Get-RSPrDeferredTestPlan -FullPlan @($same,$changed,$selftest,$writer) -PrPlan @($same,$prChanged) -RepoRoot $repoRoot)
        @($local.Id) | Should Be @('standalone.changed','selftest.commands','tooling.pester-build-toolchain')
    }
    It 'preserves every full entry when no PR contract matches' {
        $full=@(Get-RSTestRunPlan -Suite Full -RepoRoot $repoRoot -Platform x64 -Configuration Debug -RedSalamanderExePath (Join-Path $repoRoot '.build/x64/Debug/RedSalamander.exe'))
        @(Get-RSPrDeferredTestPlan -FullPlan $full -PrPlan @() -RepoRoot $repoRoot).Count | Should Be $full.Count
    }
    It 'leaves long in-product scenarios local although the PR gate exists' {
        $full=@(Get-RSTestRunPlan -Suite Full -RepoRoot $repoRoot -Platform x64 -Configuration Debug -RedSalamanderExePath (Join-Path $repoRoot '.build/x64/Debug/RedSalamander.exe'))
        $pr=@(Get-RSTestRunPlan -Suite PR -RepoRoot $repoRoot -Platform x64 -Configuration Debug -RedSalamanderExePath (Join-Path $repoRoot '.build/x64/Debug/RedSalamander.exe'))
        $local=@(Get-RSPrDeferredTestPlan -FullPlan $full -PrPlan $pr -RepoRoot $repoRoot)
        @($local | Where-Object Kind -eq SelfTest).Count | Should Be 3
        $local.Count | Should BeLessThan $full.Count
        @($local | Where-Object Id -eq tooling.pester-build-toolchain).Count | Should Be 1
    }
    It 'rejects nonexistent scope before any runtime execution' {
        $manifest=Read-ScopedTestManifest $repoRoot
        $threw=$false
        try {Get-ScopedTestPlan -Manifest $manifest -ChangedPaths @() -Scopes @('Unknown') | Out-Null}
        catch [System.Management.Automation.RuntimeException] {$threw=$true}
        $threw | Should Be $true
    }
    It 'binds PR delegation to clean committed bytes despite concurrent changes' {
        function Assert-PrCoverage([bool]$Condition,[string]$Message){if(-not $Condition){throw $Message}}
        $fixture=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('pr-coverage-'+[guid]::NewGuid().ToString('N'))
        try {
            $delegate=Join-Path $fixture 'delegation'
            function Invoke-DelegateGit([string[]]$Arguments) {Invoke-RSTestFixtureGit -Arguments @('-C'; $delegate; $Arguments) *> $null;if($LASTEXITCODE){throw "Delegation fixture git failed: $Arguments"}}
            [void](New-Item -ItemType Directory -Path $delegate -Force)
            $workflow="name: fixture`non:`n  pull_request:`njobs:`n  tests:`n    uses: ./.github/workflows/build.yml`n"
            [void](New-Item -ItemType Directory -Path (Join-Path $delegate '.github/workflows') -Force)
            [IO.File]::WriteAllText((Join-Path $delegate '.github/workflows/ci.yml'),$workflow)
            [IO.File]::WriteAllText((Join-Path $delegate '.github/workflows/build.yml'),"on: workflow_call`njobs:`n  tests:`n    steps:`n      - run: actual-tests")
            [IO.File]::WriteAllText((Join-Path $delegate 'code.cpp'),'initial')
            Invoke-DelegateGit @('init','-q')
            Invoke-DelegateGit @('config','user.name','Delegation fixture')
            Invoke-DelegateGit @('config','user.email','fixture@example.invalid')
            Invoke-DelegateGit @('add','-A');Invoke-DelegateGit @('commit','-q','-m','baseline')
            Invoke-DelegateGit @('update-ref','refs/remotes/origin/main','HEAD')
            $coverage=[pscustomobject]@{repository='fixture/example';defaultBranch='main';prWorkflowDigest=(Get-ScopedPrWorkflowDigest $delegate);prCoverage=@([pscustomobject]@{platform='x64';configuration='Release';scopes=@('Example')})}
            $global:ScopedFixtureGhCalls=0;$global:ScopedFixtureGhMutation=''
            $global:ScopedFixtureGhRoot=$delegate
            Mock Invoke-ScopedWorkflowMetadata -ModuleName ScopedTesting {
                param([string]$Repository)
                $global:ScopedFixtureGhCalls++;$global:LASTEXITCODE=0
                if($global:ScopedFixtureGhMutation -eq 'untracked') {[IO.File]::WriteAllText((Join-Path $global:ScopedFixtureGhRoot 'during-api.cpp'),'changed')}
                if($global:ScopedFixtureGhMutation -eq 'commit') {
                    [IO.File]::WriteAllText((Join-Path $global:ScopedFixtureGhRoot 'code.cpp'),'concurrent committed change')
                    Invoke-RSTestFixtureGit -Arguments @('-C'; $global:ScopedFixtureGhRoot; 'add'; 'code.cpp') *> $null
                    Invoke-RSTestFixtureGit -Arguments @('-C'; $global:ScopedFixtureGhRoot; 'commit'; '-q'; '-m'; 'concurrent change') *> $null
                }
                return [pscustomobject]@{ExitCode=0;Output='{"state":"active"}'}
            }
            try {
                Assert-PrCoverage (@(Get-ScopedPrCoverage $delegate $coverage x64 Release) -contains 'Example') 'Clean committed candidate was not delegated'
                foreach($state in @('untracked','unstaged','staged')) {
                    $path=Join-Path $delegate $(if($state -eq 'untracked'){'pending.cpp'}else{'code.cpp'})
                    [IO.File]::WriteAllText($path,'pending')
                    if($state -eq 'staged'){Invoke-DelegateGit @('add','code.cpp')}
                    $calls=$global:ScopedFixtureGhCalls
                    Assert-PrCoverage (@(Get-ScopedPrCoverage $delegate $coverage x64 Release).Count -eq 0) "Dirty candidate delegated: $state"
                    Assert-PrCoverage ($global:ScopedFixtureGhCalls -eq $calls) 'Dirty candidate consulted CI before retaining obligations locally'
                    if($state -eq 'untracked'){Remove-Item -LiteralPath $path}else{Invoke-DelegateGit @('restore','--staged','--worktree','--','code.cpp')}
                }
                [IO.File]::WriteAllText((Join-Path $delegate '.github/workflows/build.yml'),"on: workflow_call`njobs: {}")
                Invoke-DelegateGit @('add','-A');Invoke-DelegateGit @('commit','-q','-m','remove tests from reusable gate')
                Assert-PrCoverage (@(Get-ScopedPrCoverage $delegate $coverage x64 Release).Count -eq 0) 'Changed committed reusable gate delegated local obligations'
                Invoke-DelegateGit @('revert','--no-edit','HEAD')
                Assert-PrCoverage (@(Get-ScopedPrCoverage $delegate $coverage x64 Release) -contains 'Example') 'Restored execution closure was not delegated'
                $global:ScopedFixtureGhMutation='untracked'
                Assert-PrCoverage (@(Get-ScopedPrCoverage $delegate $coverage x64 Release).Count -eq 0) 'Mutation during API lookup was delegated'
                Remove-Item -LiteralPath (Join-Path $delegate 'during-api.cpp')
                $global:ScopedFixtureGhMutation='commit'
                Assert-PrCoverage (@(Get-ScopedPrCoverage $delegate $coverage x64 Release).Count -eq 0) 'Concurrent committed candidate change was delegated'
            } finally {
                Remove-Variable -Name ScopedFixtureGhCalls,ScopedFixtureGhMutation,ScopedFixtureGhRoot -Scope Global
            }
        } finally {
            [void](Assert-RSTestSandboxContainedPath -TestRoot (Get-RSTestSandboxRoot -RepoRoot $repoRoot) -Path $fixture -RequireDescendant)
            # Owned fixtures stay with the run for evidence and governed run cleanup.
        }
    }

    It 'invalidates reviewed PR coverage for the complete execution dependency surface' {
        $fixture=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('pr-closure-'+[guid]::NewGuid().ToString('N'))
        try {
            [void](New-Item -ItemType Directory -Path (Join-Path $fixture '.github/workflows') -Force)
            [void](New-Item -ItemType Directory -Path (Join-Path $fixture '.github/actions/check') -Force)
            [void](New-Item -ItemType Directory -Path (Join-Path $fixture 'Tools/Modules/Testing') -Force)
            [void](New-Item -ItemType Directory -Path (Join-Path $fixture 'Tests') -Force)
            [IO.File]::WriteAllText((Join-Path $fixture '.github/workflows/ci.yml'), "on:`n  pull_request:`njobs:`n  native:`n    uses: ./.github/workflows/build.yml`n")
            [IO.File]::WriteAllText((Join-Path $fixture '.github/workflows/build.yml'), "on: workflow_call`njobs:`n  test:`n    steps:`n      - uses: ./.github/actions/check`n")
            [IO.File]::WriteAllText((Join-Path $fixture '.github/actions/check/action.yml'), "runs:`n  using: composite`n  steps:`n    - run: ./Tools/Run.ps1`n")
            [IO.File]::WriteAllText((Join-Path $fixture '.github/actions/check/Dockerfile'), 'FROM fixture')
            [IO.File]::WriteAllText((Join-Path $fixture 'Tools/Run.ps1'), 'Import-Module ./Tools/Modules/Testing/Policy.psm1')
            [IO.File]::WriteAllText((Join-Path $fixture 'Tools/Modules/Testing/Policy.psm1'), "function Test-Policy { 'qualified' }")
            [IO.File]::WriteAllText((Join-Path $fixture 'Fixture.sln'), 'test project selection')
            [IO.File]::WriteAllText((Join-Path $fixture 'Tests/test-scopes.json'), '{"version":1,"prWorkflowDigest":"old"}')
            Invoke-RSTestFixtureGit -Arguments @('-C'; $fixture; 'init'; '-q')
            Invoke-RSTestFixtureGit -Arguments @('-C'; $fixture; 'add'; '-A')
            $baseline=Get-ScopedPrWorkflowDigest $fixture
            foreach($relative in @('.github/workflows/build.yml','.github/actions/check/action.yml','.github/actions/check/Dockerfile','Tools/Run.ps1','Tools/Modules/Testing/Policy.psm1','Fixture.sln')) {
                $path=Join-Path $fixture $relative
                $original=[IO.File]::ReadAllText($path)
                [IO.File]::WriteAllText($path,$original+"`n# changed executable dependency")
                (Get-ScopedPrWorkflowDigest $fixture) | Should Not Be $baseline
                [IO.File]::WriteAllText($path,$original)
            }
            # The digest field cannot hash itself; all remaining scope-contract fields do participate.
            [IO.File]::WriteAllText((Join-Path $fixture 'Tests/test-scopes.json'), '{"version":1,"prWorkflowDigest":"new"}')
            (Get-ScopedPrWorkflowDigest $fixture) | Should Be $baseline
            [IO.File]::WriteAllText((Join-Path $fixture 'Tests/test-scopes.json'), '{"version":2,"prWorkflowDigest":"new"}')
            (Get-ScopedPrWorkflowDigest $fixture) | Should Not Be $baseline
            [IO.File]::WriteAllText((Join-Path $fixture 'Tests/test-scopes.json'), '{"version":1,"prWorkflowDigest":"old"}')
            # A product source edit does not change which tests the reviewed gate executes.
            [IO.File]::WriteAllText((Join-Path $fixture 'code.cpp'), 'updated implementation')
            (Get-ScopedPrWorkflowDigest $fixture) | Should Be $baseline
            Remove-Item -LiteralPath (Join-Path $fixture '.github/workflows/build.yml')
            $missingRejected=$false
            try { Get-ScopedPrWorkflowDigest $fixture | Out-Null } catch [System.Management.Automation.RuntimeException] { $missingRejected=$true }
            $missingRejected | Should Be $true
        } finally {
            [void](Assert-RSTestSandboxContainedPath -TestRoot (Get-RSTestSandboxRoot -RepoRoot $repoRoot) -Path $fixture -RequireDescendant)
            # Owned fixtures stay with the run for evidence and governed run cleanup.
        }
    }

    It 'skips scheduled unchanged code only after the latest attempt succeeds' {
        $success=[pscustomobject]@{head_sha='current';conclusion='success'}
        (Test-RSScheduledTestsRequired schedule current $success) | Should Be $false
        foreach($conclusion in @('failure','cancelled','timed_out','action_required')) {
            (Test-RSScheduledTestsRequired schedule current ([pscustomobject]@{head_sha='current';conclusion=$conclusion})) | Should Be $true
        }
        (Test-RSScheduledTestsRequired workflow_dispatch current $success) | Should Be $true
        (Test-RSScheduledTestsRequired schedule new $success) | Should Be $true
        (Test-RSScheduledTestsRequired schedule current $null) | Should Be $true
        (Test-RSScheduledTestsRequired schedule current ([pscustomobject]@{})) | Should Be $true
    }
}

Describe 'Canonical scoped evidence authority' {
    It 'keeps obsolete receipt and source identity exports absent' {
        $exports = @((Get-Module ScopedTesting).ExportedCommands.Keys)
        foreach ($name in @('Get-ScopedSourceIdentity','Get-ScopedArtifactIdentity','Get-ScopedRunIdentity',
                'Get-ScopedEnvironmentIdentity','Test-ScopedReceipt','Write-ScopedReceipt')) {
            ($exports -contains $name) | Should Be $false
        }
        $runner = [IO.File]::ReadAllText((Join-Path $repoRoot 'Tools/Run-AllTests.ps1'))
        $runner | Should Match 'hosted_stock_preparation=\[string\]\$env:RS_HOSTED_ARM64_STOCK_PREPARATION'
    }
}
Describe 'Engine-free scoped metadata transport' {
    It 'preserves raw NUL Unicode and newline payloads without a CI engine or caller mutation' {
        $fixture=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('scoped-native-'+[guid]::NewGuid().ToString('N'))
        $copy=Join-Path $fixture 'ScopedBootstrap.psm1'
        Copy-Item -LiteralPath (Join-Path $repoRoot 'Tools/Modules/Testing/ScopedTesting.psm1') -Destination $copy
        (Test-Path (Join-Path $fixture '.ci-engine'))|Should Be $false
        $module=Import-Module $copy -PassThru -Scope Local -Prefix ScopedBootstrap
        $cwd=(Get-Location).Path;$prior=[Environment]::GetEnvironmentVariable('RS_METADATA_PARENT_SENTINEL','Process')
        $savedConsoleEncoding=[Console]::OutputEncoding
        try {
            [Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
            [Environment]::SetEnvironmentVariable('RS_METADATA_PARENT_SENTINEL','fixture é雪','Process')
            $value='raw é雪'+[char]13+[char]10+'middle'+[char]10
            $actual=& $module {param($root,$value)Invoke-ScopedGit $root @('-c',('fixture.payload='+$value),'config','--null','--get','fixture.payload')} $fixture $value
            $actual|Should Be ($value+[char]0)
            (Get-Location).Path|Should Be $cwd
            [Environment]::GetEnvironmentVariable('RS_METADATA_PARENT_SENTINEL','Process')|Should Be 'fixture é雪'
            [Console]::OutputEncoding=[Text.Encoding]::GetEncoding(437)
            $legacy=& $module {param($root,$value)Invoke-ScopedGit $root @('-c',('fixture.payload='+$value),'config','--null','--get','fixture.payload')} $fixture $value
            $legacy|Should Be ($value+[char]0)
            [Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
            $null=& $module {param($root)Invoke-ScopedGit $root @('init','--quiet')} $fixture
            $path='café-Δ-雪.txt'
            [IO.File]::WriteAllText((Join-Path $fixture $path),'owned fixture',[Text.UTF8Encoding]::new($false))
            foreach($encoding in @([Text.Encoding]::GetEncoding(437),[Text.UTF8Encoding]::new($false))) {
                [Console]::OutputEncoding=$encoding
                $paths=@(& $module {param($root)Get-ScopedTrackedPaths $root} $fixture)
                ($paths -ccontains $path)|Should Be $true
                [Console]::OutputEncoding.CodePage|Should Be $encoding.CodePage
            }
            $failed=$false
            try {& $module {param($root)Invoke-ScopedGit $root @('rev-parse','--verify','missing-object')} $fixture}
            catch {$failed=$_.Exception.Message.StartsWith('Git scope discovery failed: ')}
            $failed|Should Be $true
        } finally {
            [Environment]::SetEnvironmentVariable('RS_METADATA_PARENT_SENTINEL',$prior,'Process')
            [Console]::OutputEncoding=$savedConsoleEncoding
            Remove-Module $module
        }
    }
    It 'keeps an unavailable workflow API local without making a network request' {
        $fixture=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('scoped-api-'+[guid]::NewGuid().ToString('N'))
        $copy=Join-Path $fixture 'ScopedApi.psm1'
        Copy-Item -LiteralPath (Join-Path $repoRoot 'Tools/Modules/Testing/ScopedTesting.psm1') -Destination $copy
        [void][IO.Directory]::CreateDirectory((Join-Path $fixture '.github/workflows'))
        [IO.File]::WriteAllText((Join-Path $fixture '.github/workflows/ci.yml'),('on:'+ [char]10+'  pull_request:'+ [char]10))
        $module=Import-Module $copy -PassThru -Scope Local -Prefix ScopedApi
        try {
            $result=& $module {
                param($root)
                $saved=@{}
                foreach($name in @('Invoke-ScopedGit','Get-ScopedPrWorkflowDigest','Invoke-ScopedWorkflowMetadata')){$saved[$name]=(Get-Command $name).ScriptBlock}
                $script:workflowFixtureCalls=0
                function script:Invoke-ScopedGit {param($Root,$Arguments)if($Arguments[0] -eq 'rev-parse'){return ('a'*40)};return ''}
                function script:Get-ScopedPrWorkflowDigest {param($Root)return 'reviewed'}
                function script:Invoke-ScopedWorkflowMetadata {param($Repository)$script:workflowFixtureCalls++;$global:LASTEXITCODE=9;return [pscustomobject]@{ExitCode=9;Output='{"state":"active"}'}}
                try {
                    $manifest=[pscustomobject]@{prCoverage=@([pscustomobject]@{platform='ARM64';configuration='Release';scopes=@('native')});prWorkflowDigest='reviewed';repository='fixture/example';defaultBranch='main'}
                    $coverage=@(Get-ScopedPrCoverage -Root $root -Manifest $manifest -Platform ARM64 -Configuration Release)
                    return [pscustomobject]@{count=$coverage.Count;calls=$script:workflowFixtureCalls;exit=$LASTEXITCODE}
                } finally {foreach($name in $saved.Keys){Set-Item -Path ('Function:script:'+$name) -Value $saved[$name]}}
            } $fixture
            $result.count|Should Be 0
            $result.calls|Should Be 1
            $result.exit|Should Be 9
        } finally {Remove-Module $module}
    }
}
