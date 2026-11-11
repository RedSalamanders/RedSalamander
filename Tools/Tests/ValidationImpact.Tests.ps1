$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $PSScriptRoot 'TestSupport.psm1') -Force
Import-Module (Join-Path $repoRoot 'Tools\Modules\Testing\ValidationFingerprint.psm1') -Force
Import-Module (Join-Path $repoRoot 'Tools\Modules\Testing\ValidationImpact.psm1') -Force
Import-Module (Join-Path $repoRoot 'Tools\Modules\Testing\TestRunPlan.psm1') -Force

function New-RSImpactSnapshot([object[]]$Files) {
    [pscustomobject]@{ snapshot_id = 'a' * 64; files = @($Files) }
}

function New-RSImpactPlan {
    $entries = @(
        (New-RSTestRunPlanEntry -Id 'selftest.commands' -Name Commands -Kind SelfTest -Path (Join-Path $repoRoot 'app.exe'))
        (New-RSTestRunPlanEntry -Id 'selftest.file-operations' -Name FileOps -Kind SelfTest -Path (Join-Path $repoRoot 'app.exe'))
        (New-RSTestRunPlanEntry -Id 'tooling.pester' -Name Tooling -Kind Pester -Path (Join-Path $repoRoot 'Tools\Tests'))
    )
    return New-RSValidationPlan -RepoRoot $repoRoot -RequestedSuite Full -Entries $entries -ValidationMode Affected
}

Describe 'Affected native build selection before compilation' {
    BeforeAll {
        $script:buildEntries = @(Get-RSTestRunPlan -Suite Full -RepoRoot $repoRoot -Platform x64 -Configuration Debug `
            -RedSalamanderExePath (Join-Path $repoRoot '.build/x64/Debug/RedSalamander.exe'))
        $script:buildPlan = New-RSValidationPlan -RepoRoot $repoRoot -RequestedSuite Full -Entries $script:buildEntries -ValidationMode Affected
        $script:buildManifest = Read-RSValidationImpactManifest -Path (Join-Path $repoRoot 'Tools/validation-impact.json') `
            -SchemaPath (Join-Path $repoRoot 'Specs/Testing/ValidationImpact.schema.json') -RepoRoot $repoRoot
        $script:solution = Join-Path $repoRoot 'RedSalamander.sln'
    }

    It 'binds every actual native entry to a resolvable reviewed producer and portable contract' {
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Build/BuildProjectSelection.psm1') -ErrorAction Stop
        foreach ($entry in @($script:buildEntries | Where-Object Kind -in @('SelfTest','Executable','CppUnitTest'))) {
            @($entry.BuildProjectNames).Count | Should BeGreaterThan 0
            $union = Get-RSBuildProjectUnion -SolutionPath $script:solution -ProjectNames $entry.BuildProjectNames
            @($union.ProjectSelection).Count | Should BeGreaterThan 0
            $portable = @($script:buildPlan.entries | Where-Object id -eq $entry.Id)[0]
            ($portable.build_project_names -join '|') | Should Be ($entry.BuildProjectNames -join '|')
        }
    }

    It 'uses the actual impact graph to select the ProductUiTests union before any native build' {
        $snapshot = New-RSImpactSnapshot @([pscustomobject]@{path='Tests/ProductUiTests/selected.cpp';old_path=$null})
        $decision = Get-RSValidationImpactDecision -Manifest $script:buildManifest -WorkspaceSnapshot $snapshot -ValidationPlan $script:buildPlan
        $selected = Get-RSValidationAffectedBuildSelection -Decision $decision -ExecutionEntries $script:buildEntries -SolutionPath $script:solution
        $selected.UseSelectedBuild | Should Be $true
        ($selected.Selection.ProjectNames -join '|') | Should Be 'ProductUiTests'
        ($selected.EntryIds -join '|') | Should Be 'standalone.product-ui'
    }

    It 'retains the full build for shared or unknown source and artifact writers' {
        foreach ($path in @('Common/unknown.h','unknown-executable-input.xyz','Tools/Tests/changed.Tests.ps1')) {
            $snapshot = New-RSImpactSnapshot @([pscustomobject]@{path=$path;old_path=$null})
            $decision = Get-RSValidationImpactDecision -Manifest $script:buildManifest -WorkspaceSnapshot $snapshot -ValidationPlan $script:buildPlan
            $selected = Get-RSValidationAffectedBuildSelection -Decision $decision -ExecutionEntries $script:buildEntries -SolutionPath $script:solution
            $selected.UseSelectedBuild | Should Be $false
        }
        $entry = New-RSTestRunPlanEntry -Id fixture.writer -Name Writer -Kind Executable -Path (Join-Path $repoRoot 'writer.exe') -BuildProjectNames ProductUiTests
        $entry.ArtifactWriteIds=@('artifact.fixture')
        $decision=[pscustomobject]@{advisory_build_action='targeted';entries=@([pscustomobject]@{id=$entry.Id;affected=$true})}
        $selected=Get-RSValidationAffectedBuildSelection -Decision $decision -ExecutionEntries @($entry) -SolutionPath $script:solution
        $selected.UseSelectedBuild | Should Be $false
        $selected.Reason | Should Match 'artifact writer'
    }

    It 'widens an absent or invalid native producer binding and undeclared script artifact reads' {
        foreach ($names in @(@(),@('MissingProject'),@('ProductUiTests;Build'))) {
            $entry=New-RSTestRunPlanEntry -Id fixture.native -Name Native -Kind Executable -Path (Join-Path $repoRoot 'native.exe') -BuildProjectNames $names
            $decision=[pscustomobject]@{advisory_build_action='targeted';entries=@([pscustomobject]@{id=$entry.Id;affected=$true})}
            (Get-RSValidationAffectedBuildSelection -Decision $decision -ExecutionEntries @($entry) -SolutionPath $script:solution).UseSelectedBuild | Should Be $false
        }
        $entry=New-RSTestRunPlanEntry -Id fixture.script -Name Script -Kind PowerShellScript -Path (Join-Path $repoRoot 'script.ps1')
        $entry.ArtifactReadIds=@('artifact.fixture')
        $decision=[pscustomobject]@{advisory_build_action='targeted';entries=@([pscustomobject]@{id=$entry.Id;affected=$true})}
        (Get-RSValidationAffectedBuildSelection -Decision $decision -ExecutionEntries @($entry) -SolutionPath $script:solution).UseSelectedBuild | Should Be $false
    }

    It 'refuses a decision describing missing, extra or duplicate entries' {
        $entry=New-RSTestRunPlanEntry -Id fixture.native -Name Native -Kind Executable -Path (Join-Path $repoRoot 'native.exe') -BuildProjectNames ProductUiTests
        foreach ($decisionEntries in @(@(),@([pscustomobject]@{id='fixture.other';affected=$true}),@([pscustomobject]@{id=$entry.Id;affected=$true},[pscustomobject]@{id=$entry.Id;affected=$true}))) {
            $decision=[pscustomobject]@{advisory_build_action='targeted';entries=$decisionEntries}
            $attempt=Invoke-RSTestBlock {Get-RSValidationAffectedBuildSelection -Decision $decision -ExecutionEntries @($entry) -SolutionPath $script:solution}
            $attempt.Succeeded | Should Be $false
        }
    }
}

Describe 'Operation Startrail governed impact graph' {
    BeforeAll {
        $script:manifest = Read-RSValidationImpactManifest `
            -Path (Join-Path $repoRoot 'Tools\validation-impact.json') `
            -SchemaPath (Join-Path $repoRoot 'Specs\Testing\ValidationImpact.schema.json') `
            -RepoRoot $repoRoot
        $script:plan = New-RSImpactPlan
    }

    It 'preserves literal, wildcard, Unicode and invariant case decisions across cache hits and eviction' {
        $module = Get-Module -Name ValidationImpact -All | Where-Object Path -eq (Join-Path $repoRoot 'Tools\Modules\Testing\ValidationImpact.psm1') | Select-Object -First 1
        $matchPattern = & $module { (Get-Command Test-RSValidationImpactPattern).ScriptBlock }
        $cases = @(
            @('**/*.cpp','case.cpp',$true), @('**/*.cpp','Tests/case.cpp',$true),
            @('Tests/*.cpp','Tests/nested/case.cpp',$false), @('Tests/?.cpp','Tests/a.cpp',$true),
            @('Tests/?.cpp','Tests/ab.cpp',$false), @('**/a+[1](x).cpp','Tests/a+[1](x).cpp',$true),
            @('**/a+[1](x).cpp','Tests/aaaa1x.cpp',$false), @('**/日本語?.md','docs/日本語1.md',$true),
            @('I/**','i/file.cpp',$true), @('I/**','ı/file.cpp',$false),
            @('Tests/*.cpp','Tests\case.cpp',$false)
        )
        $culture = [Globalization.CultureInfo]::CurrentCulture
        try {
            foreach ($name in @('en-US','tr-TR')) {
                [Globalization.CultureInfo]::CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo($name)
                foreach ($round in 1..2) {
                    foreach ($case in $cases) {
                        (& $matchPattern -Pattern $case[0] -Path $case[1]) | Should Be $case[2]
                    }
                }
                foreach ($index in 1..1100) {
                    (& $matchPattern -Pattern "eviction/$index/*.cpp" -Path "eviction/$index/case.cpp") | Should Be $true
                }
                foreach ($case in $cases) {
                    (& $matchPattern -Pattern $case[0] -Path $case[1]) | Should Be $case[2]
                }
            }
        } finally { [Globalization.CultureInfo]::CurrentCulture = $culture }
    }

    It 'maps native formatter acquisition inputs to tooling coverage' {
        foreach ($path in @('Build/clang-format.json', 'Tools/Install-ClangFormat.ps1')) {
            $decision = Get-RSValidationImpactDecision -Manifest $manifest `
                -WorkspaceSnapshot (New-RSImpactSnapshot @([pscustomobject]@{path=$path; old_path=$null})) `
                -ValidationPlan $plan
            ($decision.entries | Where-Object id -eq 'tooling.pester').affected | Should Be $true
            ($decision.entries | Where-Object id -eq 'selftest.commands').affected | Should Be $false
        }
    }

    It 'accumulates production/shared mappings and keeps File Operations test-only changes narrow' {
        $testOnly = Get-RSValidationImpactDecision -Manifest $manifest `
            -WorkspaceSnapshot (New-RSImpactSnapshot @([pscustomobject]@{ path = 'RedSalamander/SelfTest/FileOperations/case.cpp'; old_path = $null })) `
            -ValidationPlan $plan
        ($testOnly.entries | Where-Object id -eq 'selftest.file-operations').affected | Should Be $true
        ($testOnly.entries | Where-Object id -eq 'selftest.commands').affected | Should Be $false

        $product = Get-RSValidationImpactDecision -Manifest $manifest `
            -WorkspaceSnapshot (New-RSImpactSnapshot @([pscustomobject]@{ path = 'RedSalamander/FolderWindow.FileOperations.cpp'; old_path = $null })) `
            -ValidationPlan $plan
        ($product.entries | Where-Object id -eq 'selftest.file-operations').affected | Should Be $true
        ($product.entries | Where-Object id -eq 'selftest.commands').affected | Should Be $true
    }

    It 'uses both rename sides and widens an unknown relevant path to Full' {
        $decision = Get-RSValidationImpactDecision -Manifest $manifest `
            -WorkspaceSnapshot (New-RSImpactSnapshot @([pscustomobject]@{ path = 'Unknown/new.cpp'; old_path = 'RedSalamander/SelfTest/FileOperations/old.cpp' })) `
            -ValidationPlan $plan
        $decision.changed_paths | Should Be @('RedSalamander/SelfTest/FileOperations/old.cpp', 'Unknown/new.cpp')
        $decision.advisory_build_action | Should Be 'full'
        @($decision.entries | Where-Object affected).Count | Should Be $decision.entries.Count
    }

    It 'selects every canonical Full entry for unknown inputs and shared test support' {
        $full = @(Get-RSTestRunPlan -Suite Full -RepoRoot $repoRoot -Platform x64 -Configuration Debug -RedSalamanderExePath (Join-Path $repoRoot '.build/x64/Debug/RedSalamander.exe'))
        $fullPlan = New-RSValidationPlan -RepoRoot $repoRoot -RequestedSuite Full -Entries $full -ValidationMode Affected
        foreach ($path in @('Unknown/new.cpp', 'Tests/TestSupport/Support.Tests.TestSandbox.h', 'Tools/Run-AllTests.ps1')) {
            $decision = Get-RSValidationImpactDecision -Manifest $manifest -WorkspaceSnapshot (New-RSImpactSnapshot @([pscustomobject]@{path=$path;old_path=$null})) -ValidationPlan $fullPlan
            @($decision.entries | Where-Object affected).Count | Should Be $full.Count
        }
    }

    It 'selects the complete native consumer closure for the vendored RedMonitor client and lock' {
        $full = @(Get-RSTestRunPlan -Suite Full -RepoRoot $repoRoot -Platform x64 -Configuration Debug -RedSalamanderExePath (Join-Path $repoRoot '.build/x64/Debug/RedSalamander.exe'))
        $fullPlan = New-RSValidationPlan -RepoRoot $repoRoot -RequestedSuite Full -Entries $full -ValidationMode Affected
        foreach ($case in @(
            @{ Path = 'ThirdParty/RedMonitor/include/RedMonitor/Client.h'; Rule = 'build.redmonitor-client' },
            @{ Path = 'Dependencies/RedMonitor.lock.json'; Rule = 'build.redmonitor-lock' }
        )) {
            $decision = Get-RSValidationImpactDecision -Manifest $manifest -WorkspaceSnapshot (New-RSImpactSnapshot @([pscustomobject]@{path=$case.Path;old_path=$null})) -ValidationPlan $fullPlan
            ($decision.matched_impact_rules -contains $case.Rule) | Should Be $true
            $decision.advisory_build_action | Should Be 'full'
            @($decision.entries | Where-Object affected).Count | Should Be $full.Count
        }
        @($manifest.rules | Where-Object { $_.id -in @('monitor.product', 'scoped.monitor') }).Count | Should Be 0
    }

    It 'selects SQLite viewer tests without unrelated standalone scenarios' {
        $full = @(Get-RSTestRunPlan -Suite Full -RepoRoot $repoRoot -Platform x64 -Configuration Debug -RedSalamanderExePath (Join-Path $repoRoot '.build/x64/Debug/RedSalamander.exe'))
        $fullPlan = New-RSValidationPlan -RepoRoot $repoRoot -RequestedSuite Full -Entries $full -ValidationMode Affected
        foreach ($path in @('Tests/ViewerSqliteTests/ViewerSqlite.Tests.Runner.cpp', 'Plugins/ViewerSqlite/ViewerSqlite.cpp')) {
            $decision = Get-RSValidationImpactDecision -Manifest $manifest -WorkspaceSnapshot (New-RSImpactSnapshot @([pscustomobject]@{path=$path;old_path=$null})) -ValidationPlan $fullPlan
            $affectedIds = @($decision.entries | Where-Object affected | ForEach-Object id)
            $affectedIds.Count | Should BeGreaterThan 0
            @($affectedIds | Where-Object {$_ -notlike 'viewer-sqlite.*'}).Count | Should Be 0
        }
    }

    It 'treats documentation-only changes as no-build with no affected runtime entries' {
        $decision = Get-RSValidationImpactDecision -Manifest $manifest `
            -WorkspaceSnapshot (New-RSImpactSnapshot @([pscustomobject]@{ path = 'Specs/Testing/example.md'; old_path = $null })) `
            -ValidationPlan $plan
        $decision.advisory_build_action | Should Be 'none'
        @($decision.entries | Where-Object affected).Count | Should Be 0
    }

    It 'keeps fragment and plan Markdown native-inert while unknown Changes input still widens' {
        $paths = @('Changes/2026-11-09-example.md', 'Specs/Plans/WIP/README.md', 'Specs/Plans/WIP/example.md')
        $decision = Get-RSValidationImpactDecision -Manifest $manifest `
            -WorkspaceSnapshot (New-RSImpactSnapshot @($paths | ForEach-Object { [pscustomobject]@{path=$_;old_path=$null} })) `
            -ValidationPlan $plan
        $decision.advisory_build_action | Should Be 'none'
        @($decision.entries | Where-Object affected).Count | Should Be 0
        ($decision.matched_impact_rules -contains 'documentation.changes') | Should Be $true
        foreach ($path in @('Changes/generator.ps1', 'Changes/data.json', 'Changes/nested/runtime.md')) {
            $unknown = Get-RSValidationImpactDecision -Manifest $manifest `
                -WorkspaceSnapshot (New-RSImpactSnapshot @([pscustomobject]@{path=$path;old_path=$null})) `
                -ValidationPlan $plan
            $unknown.advisory_build_action | Should Be 'full'
            @($unknown.entries | Where-Object affected).Count | Should Be $plan.entries.Count
        }
    }

    It 'selects executable Specs inputs and keeps only Markdown prose inert' {
        $cases = @(
            @{ Path = 'Specs/Build/BuildReceipt.schema.json'; Rule = 'specs.build-contracts'; Build = 'none'; Affected = 'tooling.pester' },
            @{ Path = 'Specs/Testing/ValidationPlan.schema.json'; Rule = 'specs.validation-contracts'; Build = 'none'; Affected = 'tooling.pester' },
            @{ Path = 'Specs/SettingsStore.schema.json'; Rule = 'specs.settings-contract'; Build = 'targeted'; Affected = $null },
            @{ Path = 'Specs/Themes/Dracula.theme.json5'; Rule = 'specs.theme-runtime'; Build = 'none'; Affected = 'selftest.commands' },
            @{ Path = 'Specs/Testing/FolderViewPerfBudgets.json5'; Rule = 'specs.performance-data'; Build = 'none'; Affected = 'tooling.pester' },
            @{ Path = 'Specs/Terminal/TerminalEngineFixtureCorpus.v1.json'; Rule = 'specs.terminal-corpus'; Build = 'none'; Affected = 'tooling.pester' },
            @{ Path = 'Specs/Terminal/TerminalEngineEvidenceCorpus/positive.collection.json'; Rule = 'specs.terminal-evidence-corpus'; Build = 'none'; Affected = 'tooling.pester' },
            @{ Path = 'Specs/NormativeConsistency.json'; Rule = 'specs.governance-map'; Build = 'none'; Affected = 'tooling.pester' }
        )
        foreach ($case in $cases) {
            $decision = Get-RSValidationImpactDecision -Manifest $manifest `
                -WorkspaceSnapshot (New-RSImpactSnapshot @([pscustomobject]@{ path = $case.Path; old_path = $null })) `
                -ValidationPlan $plan
            ($decision.matched_impact_rules -contains $case.Rule) | Should Be $true
            $decision.advisory_build_action | Should Be $case.Build
            if ($null -ne $case.Affected) {
                ($decision.entries | Where-Object id -eq $case.Affected).affected | Should Be $true
            }
        }

        $unknown = Get-RSValidationImpactDecision -Manifest $manifest `
            -WorkspaceSnapshot (New-RSImpactSnapshot @([pscustomobject]@{ path = 'Specs/New/ConsumedMap.json'; old_path = $null })) `
            -ValidationPlan $plan
        $unknown.matched_impact_rules | Should Be @('specs.executable-json-fallback')
        $unknown.advisory_build_action | Should Be 'full'
        @($unknown.entries | Where-Object affected).Count | Should Be $unknown.entries.Count
    }

    It 'governs every tracked project import and runtime graph input' {
        $tracked = @(& git -C $repoRoot ls-files '*.sln' '*.slnx' '*.vcxproj' '*.props' '*.targets' 'RuntimeDependencies.props')
        $governance = Test-RSValidationImpactGovernance -Manifest $manifest -TrackedPaths $tracked -RepoRoot $repoRoot
        $governance.IsValid | Should Be $true
        $governance.UnmappedBuildInputs.Count | Should Be 0
        (Test-RSValidationImpactGovernance -Manifest $manifest `
            -TrackedPaths @($tracked + 'new/Unmapped.vcxproj') -RepoRoot $repoRoot).IsValid | Should Be $false
    }

    It 'routes current product UI tests through their live consumer project' {
        $rules = @(Get-RSValidationImpactRulesForPath -Manifest $manifest -Path 'Tests/ProductUiTests/ProductUi.Tests.Runner.cpp')
        $owned = @($rules | Where-Object derived_closure_source -eq 'Tests/ProductUiTests/ProductUiTests.vcxproj')
        $owned.Count | Should Be 1
        (Test-Path -LiteralPath (Join-Path $repoRoot $owned[0].derived_closure_source) -PathType Leaf) | Should Be $true
        @($manifest.rules | Where-Object derived_closure_source -in @(
                'Common/DxUi/DxUi.vcxproj', 'Tests/DxUiTests/DxUiTests.vcxproj')).Count | Should Be 0
    }

    It 'keeps current UI adapters and the exact library pin covered by product UI tests' {
        $entries = @(
            (New-RSTestRunPlanEntry -Id 'standalone.product-ui' -Name ProductUiTests -Kind Executable -Path (Join-Path $repoRoot 'ProductUiTests.exe'))
            (New-RSTestRunPlanEntry -Id 'selftest.commands' -Name Commands -Kind SelfTest -Path (Join-Path $repoRoot 'app.exe'))
            (New-RSTestRunPlanEntry -Id 'selftest.file-operations' -Name FileOps -Kind SelfTest -Path (Join-Path $repoRoot 'app.exe'))
        )
        $productPlan = New-RSValidationPlan -RepoRoot $repoRoot -RequestedSuite Full -Entries $entries -ValidationMode Affected
        foreach ($path in @('Common/ViewerDxUiTheme.h', 'Dependencies/DxUi.lock.json', 'Tests/ProductUiTests/ProductUi.Tests.Runner.cpp')) {
            (Test-Path -LiteralPath (Join-Path $repoRoot $path) -PathType Leaf) | Should Be $true
            $decision = Get-RSValidationImpactDecision -Manifest $manifest `
                -WorkspaceSnapshot (New-RSImpactSnapshot @([pscustomobject]@{ path = $path; old_path = $null })) `
                -ValidationPlan $productPlan
            ($decision.entries | Where-Object id -eq 'standalone.product-ui').affected | Should Be $true
            if ($path -eq 'Tests/ProductUiTests/ProductUi.Tests.Runner.cpp') {
                ($decision.entries | Where-Object id -eq 'selftest.commands').affected | Should Be $false
                ($decision.entries | Where-Object id -eq 'selftest.file-operations').affected | Should Be $false
            } else {
                $decision.advisory_build_action | Should Be 'full'
                @($decision.entries | Where-Object affected).Count | Should Be $entries.Count
            }
        }
    }

    It 'emits deterministic schema-valid shadow decisions and never claims reuse' {
        $snapshot = New-RSImpactSnapshot @([pscustomobject]@{ path = 'Tools/Tests/new.Tests.ps1'; old_path = $null })
        $estimates = @{ 'selftest.commands' = 3000L; 'selftest.file-operations' = 2000L; 'tooling.pester' = 1000L }
        $first = Get-RSValidationImpactDecision -Manifest $manifest -WorkspaceSnapshot $snapshot `
            -ValidationPlan $plan -EstimatedDurationMs $estimates
        $second = Get-RSValidationImpactDecision -Manifest $manifest -WorkspaceSnapshot $snapshot `
            -ValidationPlan $plan -EstimatedDurationMs $estimates
        $first.decision_digest | Should Be $second.decision_digest
        @($first.entries | Where-Object decision -ne 'EXECUTE').Count | Should Be 0
        ($first.entries | Where-Object id -eq 'selftest.commands').estimated_duration_ms | Should Be 3000
        $first.estimated_fresh_duration_ms | Should Be 6000
        $first.estimated_executed_duration_ms | Should Be 6000
        $first.estimated_avoided_duration_ms | Should Be 0
        ($first | ConvertTo-Json -Depth 30 | Test-Json -SchemaFile (Join-Path $repoRoot 'Specs\Testing\ValidationPlanDecision.schema.json')) | Should Be $true
    }

    It 'freezes the phase-4 mode gate and runner ordering contract' {
        $matrix = @(Get-RSValidationModeGateMatrix)
        ($matrix.mode -join ',') | Should Be 'Fresh,ExplainPlan,Resume,Affected,CI,Release'
        @($matrix | Where-Object permits_reuse).Count | Should Be 1
        ($matrix | Where-Object mode -eq Fresh).public | Should Be $true
        ($matrix | Where-Object mode -eq ExplainPlan).public | Should Be $true
        ($matrix | Where-Object mode -in @('Resume', 'Affected') | Where-Object public).Count | Should Be 2
        ($matrix | Where-Object mode -eq Resume).permits_reuse | Should Be $true
        ($matrix | Where-Object mode -eq Affected).permits_reuse | Should Be $false
        $runner = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools\Run-AllTests.ps1') -Raw
        $runner.IndexOf('if ($ExplainPlan)') | Should BeLessThan $runner.IndexOf('$artifactOperationLock = $null')
        $runner | Should Match 'Fresh still executes all'
    }
}
