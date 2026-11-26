Import-Module (Join-Path $PSScriptRoot 'TestSupport.psm1') -Force -ErrorAction Stop
$repoRoot=Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/ScopedTesting.psm1') -Force
Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestSandbox.psm1') -Force
Describe 'Scoped testing entry points' {
    It 'validates real fragments before affected execution and keeps explanation non-executing' {
        $fixture=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('fragment-admission-'+[guid]::NewGuid().ToString('N'))
        $moduleDirectory=Join-Path $fixture 'Tools/Modules/Testing'
        [void](New-Item -ItemType Directory -Path $moduleDirectory,(Join-Path $fixture '.ci'),(Join-Path $fixture 'Changes') -Force)
        foreach($relative in @('Test-Changes.ps1','test.ps1','Tools/Test-Changelog.ps1','.ci/project.json')) {
            Copy-Item -LiteralPath (Join-Path $repoRoot $relative) -Destination (Join-Path $fixture $relative)
        }
        # The Tools/Test-Changelog.ps1 engine shim runs the pinned engine from the fixture's own .ci-engine.
        Invoke-RSTestFixtureGit -Arguments @('clone';'--quiet';'--no-hardlinks';(Join-Path $repoRoot '.ci-engine');(Join-Path $fixture '.ci-engine')) *> (Join-Path $fixture 'engine-clone.log')
        $LASTEXITCODE | Should Be 0
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
    It 'preserves legacy profile controls and hands ordinary Full and PrePush to the engine' {
        $fixture=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('alternate-forwarding-'+[guid]::NewGuid().ToString('N'))
        $succeeded=$false
        try {
            $moduleDirectory=Join-Path $fixture 'Tools/Modules/Testing'
            [void](New-Item -ItemType Directory -Path $moduleDirectory -Force)
            foreach ($entry in 'Test-Changes.ps1','test.ps1') {
                Copy-Item -LiteralPath (Join-Path $repoRoot $entry) -Destination (Join-Path $fixture $entry)
            }
            'param($Base)' | Set-Content -LiteralPath (Join-Path $fixture 'Tools/Test-Changelog.ps1')
            @'
function Assert-ScopedTestNames { param($Root) return 1 }
Export-ModuleMember -Function Assert-ScopedTestNames
'@ | Set-Content -LiteralPath (Join-Path $moduleDirectory 'ScopedTesting.psm1')
            @'
param($Suite,$Configuration,$Platform,$ImpactBase,[switch]$SkipBuild,[switch]$NonInteractive,$ValidationMode,$TestRoot,
      [switch]$AllowAlternateVolumeTestRoot,$CaseFilter='',$CommandsFamily='',[int]$BuildNumber=0,[switch]$UseSharedAffectedPlanner)
[ordered]@{
    suite=$Suite; mode=$ValidationMode; testRoot=$TestRoot; caseFilter=$CaseFilter
    alternate=[bool]$AllowAlternateVolumeTestRoot; nonInteractive=[bool]$NonInteractive; buildNumber=$BuildNumber
    sharedAffected=[bool]$UseSharedAffectedPlanner
} |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $PSScriptRoot '../captured.json')
$global:LASTEXITCODE=2
'@ | Set-Content -LiteralPath (Join-Path $fixture 'Tools/Run-AllTests.ps1')
            $cases = @(
                @{ name='LegacyFull'; script='Test-Changes.ps1'; arguments=@('-Mode','Full'); suite='Full'; mode='Fresh'; filter=''; number=0 }
                @{ name='LegacyPrePush'; script='Test-Changes.ps1'; arguments=@('-Mode','PrePush'); suite='Full'; mode='Fresh'; filter=''; number=0 }
                @{ name='LegacyFocused'; script='Test-Changes.ps1'; arguments=@('-Suite','FileOps','-CaseFilter','FileOps_IndependentVolumes'); suite='FileOps'; mode='Fresh'; filter='FileOps_IndependentVolumes'; number=0 }
                @{ name='RootAffected'; script='test.ps1'; arguments=@(); suite='Full'; mode='Affected'; filter=''; number=0 }
                @{ name='RootFull'; script='test.ps1'; arguments=@('-Full','-BuildNumber','42'); suite='Full'; mode='Fresh'; filter=''; number=42 }
            )
            foreach ($case in $cases) {
                foreach ($optIn in @($false,$true)) {
                    $arguments = @('-NoProfile','-File',(Join-Path $fixture $case.script)) + $case.arguments
                    if ($case.name -cnotin @('LegacyPrePush','RootAffected') -or $optIn) { $arguments += @('-TestRoot',$env:REDSALAMANDER_TEST_ROOT) }
                    if ($optIn) { $arguments += '-AllowAlternateVolumeTestRoot' }
                    Invoke-RSTestFixtureNative -FilePath (Join-Path $PSHOME 'pwsh.exe') -Arguments $arguments *> (Join-Path $fixture "child-$($case.name)-$optIn.log")
                    $LASTEXITCODE | Should Be 2
                    $captured = Get-Content -LiteralPath (Join-Path $fixture 'captured.json') -Raw | ConvertFrom-Json
                    $captured.alternate | Should Be $optIn
                    $captured.suite | Should Be $case.suite
                    $captured.mode | Should Be $case.mode
                    $captured.sharedAffected | Should Be ($case.name -ceq 'RootAffected' -and -not $optIn)
                    $captured.caseFilter | Should Be $case.filter
                    $captured.buildNumber | Should Be $case.number
                    $captured.nonInteractive | Should Be $true
                    $captured.testRoot | Should Be $env:REDSALAMANDER_TEST_ROOT
                }
            }
            $engine = Join-Path $fixture '.ci-engine/Tools'
            [void](New-Item -ItemType Directory -Path (Join-Path $engine 'Modules') -Force)
            @'
function Resolve-CiPinnedEngineCommand {
    param($ProjectRoot,$Name)
    Join-Path $ProjectRoot ".ci-engine/Tools/$Name"
}
Export-ModuleMember -Function Resolve-CiPinnedEngineCommand
'@ | Set-Content -LiteralPath (Join-Path $engine 'Modules/Ci.EngineBinding.psm1')
            @'
param($Root,$Mode,[string]$Platform,[string]$Configuration,[string]$BaseRef,[switch]$SkipBuild,[switch]$Explain)
[ordered]@{
    mode=$Mode; platform=$Platform; configuration=$Configuration; base=$BaseRef
    skipBuild=[bool]$SkipBuild; explain=[bool]$Explain
} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $Root 'shared.json')
if ($BaseRef -ceq 'blocked-fixture') { exit 4 }
exit 0
'@ | Set-Content -LiteralPath (Join-Path $engine 'Test-Changes.ps1')
            $hostPlatform = if ([Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString() -ceq 'Arm64') { 'ARM64' } else { 'x64' }
            $sharedCases = @(
                @{ name='Full'; arguments=@('-Full','-Configuration','ASan Debug','-BaseRef','reviewed-base','-SkipBuild');
                    mode='Full'; platform=$hostPlatform; configuration='ASan Debug'; base='reviewed-base'; skipBuild=$true; explain=$false; exitCode=2 },
                @{ name='PrePush'; arguments=@('-Mode','PrePush','-Explain');
                    mode='PrePush'; platform=''; configuration=''; base=''; skipBuild=$false; explain=$true; exitCode=2 },
                @{ name='Blocked'; arguments=@('-Mode','Full','-Platform','x64','-BaseRef','blocked-fixture');
                    mode='Full'; platform='x64'; configuration='Debug'; base='blocked-fixture'; skipBuild=$false; explain=$false; exitCode=4 }
            )
            foreach ($case in $sharedCases) {
                $arguments = @('-NoProfile','-File',(Join-Path $fixture 'test.ps1')) + $case.arguments
                Invoke-RSTestFixtureNative -FilePath (Join-Path $PSHOME 'pwsh.exe') -Arguments $arguments *> (Join-Path $fixture "shared-$($case.name).log")
                $LASTEXITCODE | Should Be $case.exitCode
                $captured = Get-Content -LiteralPath (Join-Path $fixture 'shared.json') -Raw | ConvertFrom-Json
                foreach ($property in 'mode','platform','configuration','base','skipBuild','explain') {
                    $captured.$property | Should Be $case[$property]
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
}

Describe 'Shared scope data and report contract' {
    It 'matches canonical noninteractive entries in all six profiles and keeps declared desktop obligations' {
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestSuitePlan.psm1') -Scope Local
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/ValidationImpact.psm1') -Scope Local
        Import-Module (Join-Path $repoRoot '.ci-engine/Tools/Modules/Ci.ScopedTesting.psm1') -Scope Local -Prefix CI
        $shared = Read-CIScopedTestManifest -Root $repoRoot
        $shared.nativeProfileOwner | Should Be 'product'
        @($shared.prePushProfiles | Where-Object { $_.buildOnly -ceq 'cross' }).Count | Should Be 6
        $impact = Read-RSValidationImpactManifest -RepoRoot $repoRoot `
            -Path (Join-Path $repoRoot 'Tools/validation-impact.json') `
            -SchemaPath (Join-Path $repoRoot 'Specs/Testing/ValidationImpact.schema.json')
        $changed = @(
            @{ paths=@('RedSalamander/SelfTest/CompareDirectories/new.cpp') },
            @{ paths=@('RedSalamander/SelfTest/Commands/new.cpp') },
            @{ paths=@('Tests/SettingsSchemaTests/renamed.cpp','Tests/LocalizationTests/original.cpp') },
            @{ paths=@('Tests/Unregistered/new.cpp') },
            @{ paths=@('Plugins/ViewerSqlite/new.cpp') },
            @{ paths=@('Tools/Tests/deleted.ps1','Specs/Testing/foo.md') },
            @{ paths=@('Specs/Testing/new.json5') },
            @{ paths=@('Dependencies/RedMonitor.lock.json') },
            @{ paths=@('docs/usage.md','Changes/2026-11-26-fixture.md') },
            @{ paths=@('unclassified/new-input') }
        )
        $desktop = @((Get-CIScopedInteractiveObligations -Manifest $shared -Full).suite | Sort-Object)
        foreach ($platform in 'x64','ARM64') {
            foreach ($configuration in 'Debug','Release','ASan Debug') {
                $selectShared = New-RSSharedAffectedPlanner -RepoRoot $repoRoot -Configuration $configuration
                $executable = Join-Path $repoRoot ".build/$platform/$configuration/RedSalamander.exe"
                $plan = @(Get-RSTestRunPlan -Suite Full -RepoRoot $repoRoot -Platform $platform `
                        -Configuration $configuration -RedSalamanderExePath $executable)
                $portable = @($plan | Where-Object { -not $_.RequiresInteractiveDesktop })
                ($desktop -join ',') | Should Be (@($plan | Where-Object RequiresInteractiveDesktop | ForEach-Object Id | Sort-Object) -join ',')
                ((@($shared.scopes | Where-Object native | ForEach-Object name) -join ',')) | Should Be ($portable.Id -join ',')
                @($shared.scopes | Where-Object reuse).Count | Should Be 0
                foreach ($case in $changed) {
                    $tags = @(foreach ($path in $case.paths) {
                        Get-RSValidationImpactRulesForPath -Manifest $impact -Path $path | ForEach-Object impact_tags
                    })
                    $expected = @($portable | Where-Object { @($_.ImpactTags | Where-Object { $_ -cin $tags }).Count } | ForEach-Object Id)
                    $actual = @(& $selectShared -ChangedPaths $case.paths)
                    ($actual -join ',') | Should Be ($expected -join ',')
                }
                $selection = @(Select-RSScopedTestRunEntries -Entries $plan -EntryId @('standalone.settings-schema','tooling.pester-build-toolchain'))
                ($selection.Id -join ',') | Should Be 'tooling.pester-build-toolchain,standalone.settings-schema'
                @($selection[0].ArtifactWriteIds).Count | Should BeGreaterThan 0
                $invalidSelections = @(
                    @{ ids=@('selftest.commands'); message="Unknown or desktop-required scoped test entry 'selftest.commands'." },
                    @{ ids=@('missing.entry'); message="Unknown or desktop-required scoped test entry 'missing.entry'." },
                    @{ ids=@('selftest.compare-directories','selftest.compare-directories'); message="Duplicate scoped test entry 'selftest.compare-directories'." }
                )
                foreach ($invalid in $invalidSelections) {
                    $refusal = ''
                    try { $null = Select-RSScopedTestRunEntries -Entries $plan -EntryId $invalid.ids }
                    catch { $refusal = $_.Exception.Message }
                    $refusal | Should Be $invalid.message
                }
            }
        }
    }

    It 'keeps the affected build decision and existing native obligations when shared scopes apply' {
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestSuitePlan.psm1') -Scope Local
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/ValidationImpact.psm1') -Scope Local
        $manifest = Read-RSValidationImpactManifest -RepoRoot $repoRoot `
            -Path (Join-Path $repoRoot 'Tools/validation-impact.json') `
            -SchemaPath (Join-Path $repoRoot 'Specs/Testing/ValidationImpact.schema.json')
        $entries = @(Get-RSTestRunPlan -Suite Full -RepoRoot $repoRoot -Platform x64 -Configuration Debug `
                -RedSalamanderExePath (Join-Path $repoRoot '.build/x64/Debug/RedSalamander.exe') | Where-Object { -not $_.RequiresInteractiveDesktop })
        $plan = @{ plan_digest=('0'*64); entries=@($entries | ForEach-Object { @{ id=$_.Id; impact_tags=$_.ImpactTags } }) }
        $snapshot = @{ snapshot_id=('1'*64); files=@(@{ path='Tests/SettingsSchemaTests/new.cpp'; old_path='' }) }
        $options = @{ Manifest=$manifest; ValidationPlan=$plan; WorkspaceSnapshot=$snapshot; PlannerMode='enforced' }
        $legacy = Get-RSValidationImpactDecision @options
        $withShared = Get-RSValidationImpactDecision @options -SelectSharedAffectedEntries { param($ChangedPaths) 'standalone.localization' }
        $withShared.advisory_build_action | Should Be $legacy.advisory_build_action
        $affected = @($withShared.entries | Where-Object affected | ForEach-Object id)
        ($affected -ccontains 'standalone.localization') | Should Be $true
        foreach ($entry in @($legacy.entries | Where-Object affected)) { ($affected -ccontains $entry.id) | Should Be $true }
        (@($legacy.entries | Where-Object affected | ForEach-Object id) -ccontains 'standalone.settings-schema') | Should Be $true
        $snapshot.files[0].path = 'docs/UserGuide.md'
        $prose = Get-RSValidationImpactDecision @options -SelectSharedAffectedEntries { throw 'A compiler-free change must not run native scope selection.' }
        $prose.advisory_build_action | Should Be 'none'
        @($prose.entries | Where-Object affected).Count | Should Be 0
    }

    It 'projects only fresh promotions and retains scoped skips and failure' {
        $module = Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestRunSummary.psm1') -Scope Local -PassThru
        $schema = Join-Path $repoRoot 'Specs/Testing/ScopedSuiteReport.schema.json'
        & $module {
            param($SchemaPath)
            $runId = '20261126T120000Z-1-' + ('3' * 32)
            $entry = [pscustomobject]@{ Id='standalone.settings-schema'; Name='SettingsSchemaTests'; RequiresInteractiveDesktop=$false }
            $summary = [pscustomobject]@{
                schema='red-salamander.run-all-tests.v2'; validation_mode='Fresh'; requested_gate='Full'; coverage_scope='filtered'
                run_id=$runId; repository_verdict='NOT_EVALUATED'; change_verdict='PASSED'
                entries=@([pscustomobject]@{ id=$entry.Id; execution='executed'; promotion_state='promoted'; origin_run_id=$runId; evidence_digest=('a'*64) })
            }
            $legacy = [pscustomobject]@{
                run_id=$runId; platform='ARM64'; configuration='ASan Debug'
                suites=@([pscustomobject]@{ suite=$entry.Name; exit_code=0; failed=0; skipped=2 })
            }
            $arguments = @{ Summary=$summary; LegacySummary=$legacy; Entry=@($entry); SummarySha256=('b'*64); RunExitCode=2 }
            $report = Get-RSScopedSuiteReports @arguments
            $report.exitCode | Should Be 0
            $report.skips.Count | Should Be 1
            $report.platform | Should Be 'ARM64'
            $report.configuration | Should Be 'ASan Debug'
            $report.summarySha256 | Should Be ('b'*64)
            $report.evidenceDigest | Should Be ('a'*64)
            (Test-Json -Json (ConvertTo-Json -InputObject $report -Depth 6) -SchemaFile $SchemaPath) | Should Be $true
            $legacy.suites[0].skipped = 0
            (Get-RSScopedSuiteReports @arguments).skips.Count | Should Be 0
            $summary.entries[0].execution = 'reused'
            (Get-RSScopedSuiteReports @arguments).exitCode | Should Be 1
            $summary.entries[0].execution = 'executed'
            $summary.entries[0].promotion_state = 'incomplete'
            (Get-RSScopedSuiteReports @arguments).exitCode | Should Be 1
            $summary.entries[0].promotion_state = 'promoted'
            $legacy.suites[0].failed = 1
            (Get-RSScopedSuiteReports @arguments).exitCode | Should Be 1
            $legacy.suites[0].failed = 0
            $arguments.RunExitCode = 5
            (Get-RSScopedSuiteReports @arguments).exitCode | Should Be 1
            $legacy.run_id = 'another-run'
            $refusal = ''
            try { $null = Get-RSScopedSuiteReports @arguments }
            catch { $refusal = $_.Exception.Message }
            $refusal | Should Be 'Shared suite reports require a matching Fresh selected run.'
        } $schema
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
}
