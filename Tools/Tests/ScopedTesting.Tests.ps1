$repoRoot=Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestRunPlan.psm1') -Force
Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/ScopedTesting.psm1') -Force
Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestSandbox.psm1') -Force
Describe 'Scoped testing and PR coverage contracts' {
    It 'forwards explicit alternate-volume opt-in through Full and focused FileOps without enabling it by default' {
        $fixture=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('alternate-forwarding-'+[guid]::NewGuid().ToString('N'))
        $succeeded=$false
        try {
            $moduleDirectory=Join-Path $fixture 'Tools/Modules/Testing'
            [void](New-Item -ItemType Directory -Path $moduleDirectory -Force)
            Copy-Item -LiteralPath (Join-Path $repoRoot 'Test-Changes.ps1') -Destination (Join-Path $fixture 'Test-Changes.ps1')
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
                    & (Join-Path $PSHOME 'pwsh.exe') @arguments *> (Join-Path $fixture "child-$focused-$optIn.log")
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
            function Invoke-DelegateGit([string[]]$Arguments) {& git -C $delegate @Arguments *> $null;if($LASTEXITCODE){throw "Delegation fixture git failed: $Arguments"}}
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
            $previousGh=Get-Item Function:\global:gh -ErrorAction SilentlyContinue
            $global:ScopedFixtureGhCalls=0;$global:ScopedFixtureGhMutation=''
            $global:ScopedFixtureGhRoot=$delegate
            function global:gh {
                param([Parameter(ValueFromRemainingArguments=$true)][string[]]$FixtureGhArguments)
                $global:ScopedFixtureGhCalls++;$global:LASTEXITCODE=0
                if($global:ScopedFixtureGhMutation -eq 'untracked') {[IO.File]::WriteAllText((Join-Path $global:ScopedFixtureGhRoot 'during-api.cpp'),'changed')}
                if($global:ScopedFixtureGhMutation -eq 'commit') {
                    [IO.File]::WriteAllText((Join-Path $global:ScopedFixtureGhRoot 'code.cpp'),'concurrent committed change')
                    & git -C $global:ScopedFixtureGhRoot add code.cpp *> $null
                    & git -C $global:ScopedFixtureGhRoot commit -q -m 'concurrent change' *> $null
                }
                '{"state":"active"}'
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
                if($previousGh){Set-Item Function:\global:gh -Value $previousGh.ScriptBlock}else{Remove-Item Function:\global:gh}
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
            & git -C $fixture init -q
            & git -C $fixture add -A
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
