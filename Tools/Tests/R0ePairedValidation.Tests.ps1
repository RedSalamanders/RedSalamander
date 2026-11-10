$module = Join-Path $PSScriptRoot '..\Modules\Testing\R0ePairedValidation.psm1'
Import-Module $module -Force

function New-R0eFixture([string]$Side='candidate', [bool]$Red=$false) {
    $runId='20261106T104022Z-60312-10c85b8711394035839f567d80335c88'
    $status=if($Red){'failed'}else{'passed'}; $elapsed=if($Red){2100000}else{900000}
    $summary=@{schema='red-salamander.run-all-tests.v1';run_id=$runId;suite='FileOps';case_filter='FileOps_ProviderCapabilityMatrix';
        platform='x64';configuration='Release';timeout_scale=1.0;fail_fast=$false;exit_code=$(if($Red){1}else{2});
        passed=$(if($Red){2}else{3});failed=$(if($Red){1}else{0});skipped=0;total=3;
        quarantine=@{is_valid=$true;has_blocking_entries=$false;active_count=0;invalid_count=0;active_entries=@();invalid_entries=@();errors=@();repair_attempt_count=0}}
    $native=@{suite='FileOperations';timeout_scale=1.0;repeat_count=1;fail_fast=$false;passed=$summary.passed;failed=$summary.failed;skipped=0;cases=@(
        @{name='Setup';status='passed';repeat_index=1},@{name='FileOps_ProviderCapabilityMatrix';status=$status;repeat_index=1;
            reason='R0e capability/admission containment must remain complete, provider-call-free, and within the build-specific ceiling.'},
        @{name='Cleanup_RestorePluginConfig';status='passed';repeat_index=1})}
    $names=@('RouteContainmentUs','CapabilityParseWorkUs','AdmissionWorkUs','RejectionChecksUs','CapabilityParseCount','AdmissionDecisionCount','RejectedUncontainedCount','ProviderCallCount')
    $values=@(4096,10,20,30,4096,4096,1,0);$metrics=@(for($i=0;$i -lt $names.Count;$i++){
        [pscustomobject]@{timestamp='2026-11-06T10:40:22Z';process=1234;thread=4321;metric="FileOps.SelfTest.R0e.$($names[$i])";
            detail=$(if($i -eq 0){'4096-capability-admission-decisions'}else{'value'});hr=0;durationUs=$(if($i -eq 0){$elapsed}else{0});
            value0=$values[$i];value1=$(if($i -eq 0){4096}else{0});value=$(if($i -eq 0){$elapsed}else{$values[$i]});
            unit=$(if($i -eq 0){'us'}else{'count'});scenario='SelfTest';build='Release';branch='';commit='';
            machineHash='machine';runId='2026-11-06_104022'}})
    return @{Summary=$summary;Native=$native;Metrics=$metrics;Side=$Side;RunId=$runId;Platform='x64';RunnerExit=$summary.exit_code}
}
function Assert-R0eThrows([scriptblock]$Action, [string]$Expected) {
    $caught=$null
    try { & $Action | Out-Null } catch { $caught=$_.Exception.Message }
    ($null -ne $caught) | Should Be $true
    ($caught -like "*$Expected*") | Should Be $true
}
Describe 'R0e paired evidence reducer' {
    It 'accepts a green candidate and recognizes only reviewed timing-only baseline RED' {
        $green=New-R0eFixture; (Test-RSR0eSample @green).red | Should Be $false
        $red=New-R0eFixture 'baseline' $true
        (Test-RSR0eSample @red).red | Should Be $true
    }
    It 'rejects RED candidates, unrelated failures, missing and duplicate metrics and changed counts' {
        $red=New-R0eFixture 'candidate' $true; Assert-R0eThrows { Test-RSR0eSample @red } 'Baseline is not'
        $red=New-R0eFixture 'baseline' $true; $red.Native.cases[1].reason='other failure'; Assert-R0eThrows { Test-RSR0eSample @red } 'Baseline is not'
        $fixture=New-R0eFixture; $fixture.Metrics=@($fixture.Metrics|Where-Object metric -notlike '*ProviderCallCount'); Assert-R0eThrows { Test-RSR0eSample @fixture } 'R0e'
        $fixture=New-R0eFixture; $fixture.Metrics+=@($fixture.Metrics[0]); Assert-R0eThrows { Test-RSR0eSample @fixture } 'R0e'
        $fixture=New-R0eFixture; $fixture.Metrics[4].value0=4095; Assert-R0eThrows { Test-RSR0eSample @fixture } 'R0e'
    }
    It 'rejects unclean quarantine and malformed or mismatched native profile' {
        $fixture=New-R0eFixture; $fixture.Summary.quarantine.has_blocking_entries=$true; Assert-R0eThrows { Test-RSR0eSample @fixture } 'quarantine'
        $fixture=New-R0eFixture; $fixture.Summary.timeout_scale=2; Assert-R0eThrows { Test-RSR0eSample @fixture } 'profile'
        $fixture=New-R0eFixture; $fixture.Native.cases[2].status='failed'; Assert-R0eThrows { Test-RSR0eSample @fixture } 'setup/cleanup'
        $fixture=New-R0eFixture; $fixture.Native.timeout_scale=2; Assert-R0eThrows { Test-RSR0eSample @fixture } 'profile mismatch'
        $fixture=New-R0eFixture; $fixture.Native.repeat_count=2; Assert-R0eThrows { Test-RSR0eSample @fixture } 'profile mismatch'
        $fixture=New-R0eFixture; $fixture.Native.fail_fast=$true; Assert-R0eThrows { Test-RSR0eSample @fixture } 'profile mismatch'
        $fixture=New-R0eFixture; $fixture.Native.cases[1].repeat_index=2; Assert-R0eThrows { Test-RSR0eSample @fixture } 'repeat index'
    }
    It 'rejects unreviewed workload text' { Assert-R0eThrows { Get-RSR0eCaseIdentity 'case SelfTestState::Step::FileOps_ProviderCapabilityMatrix:' } 'R0e' }
}
Describe 'R0e paired comparison' {
    It 'keeps an evaluated timing-only RED baseline non-passing and rejects regressions and duplicate runs' {
        $base=@();$candidate=@()
        for($i=1;$i -le 5;$i++){
            $f=New-R0eFixture 'baseline' ($i -eq 5)
            $f.RunId="20261106T104022Z-$i-aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"; $f.Summary.run_id=$f.RunId
            $base+=,(Test-RSR0eSample @f)
            $f=New-R0eFixture
            $f.RunId="20261106T104022Z-$i-bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb"; $f.Summary.run_id=$f.RunId
            $candidate+=,(Test-RSR0eSample @f)
        }
        (Get-RSR0eAssessment -Baseline $base -Candidate $candidate).outcome | Should Be 'NOT_EVALUATED'
        $candidate[4].run_id=$base[0].run_id
        Assert-R0eThrows { Get-RSR0eAssessment -Baseline $base -Candidate $candidate } 'identities are inconsistent'
        $candidate[4].run_id='20261106T104022Z-5-bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb'
        $base[4].red=$false; $base[4].metrics.RouteContainmentUs=900000
        $candidate[4].metrics.RouteContainmentUs=1400000
        Assert-R0eThrows { Get-RSR0eAssessment -Baseline $base -Candidate $candidate } 'matched p95'
    }
}

Describe 'R0e sample profile and timing boundaries' {
    It 'admits the last microsecond below two seconds and rejects the exact two-second candidate' {
        $f=New-R0eFixture; $f.Metrics[0].durationUs=1999999; $f.Metrics[0].value=1999999; (Test-RSR0eSample @f).red | Should Be $false
        $f=New-R0eFixture; $f.Metrics[0].durationUs=2000000; $f.Metrics[0].value=2000000; Assert-R0eThrows { Test-RSR0eSample @f } 'R0e'
        $f=New-R0eFixture 'baseline' $true; $f.Metrics[0].durationUs=2000000; $f.Metrics[0].value=2000000; (Test-RSR0eSample @f).red | Should Be $true
    }
    It 'accepts instrumented RED only when its reported witness matches raw metrics' {
        $f=New-R0eFixture 'baseline' $true
        $f.Native.cases[1].reason='R0e capability/admission containment must remain complete, provider-call-free, and within the build-specific ceiling (parses=4096/4096, decisions=4096, rejected=1, providerCalls=0, elapsed=2100000 us, ceiling=2000000 us, parseWork=10 us, admissionWork=20 us, rejectionChecks=30 us).'
        (Test-RSR0eSample @f).red | Should Be $true
        $f.Metrics[2].value0=21; Assert-R0eThrows { Test-RSR0eSample @f } 'value row'
    }
    It 'rejects a RED reason with mismatched measured elapsed or any changed predicate' {
        $f=New-R0eFixture 'baseline' $true
        $f.Native.cases[1].reason='R0e capability/admission containment must remain complete, provider-call-free, and within the build-specific ceiling (parses=4096/4096, decisions=4096, rejected=1, providerCalls=0, elapsed=2099999 us, ceiling=2000000 us, parseWork=10 us, admissionWork=20 us, rejectionChecks=30 us).'
        Assert-R0eThrows { Test-RSR0eSample @f } 'RED diagnostic disagrees'
        $f=New-R0eFixture 'baseline' $true; $f.Metrics[7].value0=1; Assert-R0eThrows { Test-RSR0eSample @f } 'R0e'
        $f=New-R0eFixture 'baseline' $true; $f.Metrics[6].value0=0; Assert-R0eThrows { Test-RSR0eSample @f } 'R0e'
        $f=New-R0eFixture 'baseline' $true; $f.Metrics[5].value0=4095; Assert-R0eThrows { Test-RSR0eSample @f } 'R0e'
        $f=New-R0eFixture 'baseline' $true; $f.Native.cases[1].reason='R0e capability/admission containment must remain complete, provider-call-free, and within the build-specific ceiling. Extra failure.'; Assert-R0eThrows { Test-RSR0eSample @f } 'Baseline is not'
    }
    It 'accepts canonical floating unit scales and rejects invalid scale types and values' {
        $fixture=New-R0eFixture
        $fixture.Summary.timeout_scale.GetType().Name | Should Be 'Double'
        $fixture.Native.timeout_scale.GetType().Name | Should Be 'Double'
        (Test-RSR0eSample @fixture).red | Should Be $false
        foreach($invalid in @($null,'1',$true,[double]::NaN,[double]::PositiveInfinity,2.0)) {
            $fixture=New-R0eFixture; $fixture.Summary.timeout_scale=$invalid
            Assert-R0eThrows { Test-RSR0eSample @fixture } 'R0e'
            $fixture=New-R0eFixture; $fixture.Native.timeout_scale=$invalid
            Assert-R0eThrows { Test-RSR0eSample @fixture } 'R0e'
        }
    }
    It 'rejects absent, null or coercible but nonnumeric raw metric facts' {
        $f=New-R0eFixture; $f.Metrics[7].PSObject.Properties.Remove('value0'); Assert-R0eThrows { Test-RSR0eSample @f } 'value0'
        $f=New-R0eFixture; $f.Metrics[7].value0=$null; Assert-R0eThrows { Test-RSR0eSample @f } 'R0e'
        $f=New-R0eFixture; $f.Metrics[7].value0='0'; Assert-R0eThrows { Test-RSR0eSample @f } 'R0e'
        $f=New-R0eFixture; $f.Metrics[7].value0=$false; Assert-R0eThrows { Test-RSR0eSample @f } 'R0e'
        $f=New-R0eFixture; $f.Metrics[0].durationUs=$null; Assert-R0eThrows { Test-RSR0eSample @f } 'R0e'
        $f=New-R0eFixture; $f.Metrics[0].PSObject.Properties.Remove('hr'); Assert-R0eThrows { Test-RSR0eSample @f } 'hr'
    }
    It 'rejects mismatched machine, case identity, counts, profile and quarantine metadata' {
        $f=New-R0eFixture; $f.Metrics[2].machineHash='other'; Assert-R0eThrows { Test-RSR0eSample @f } 'R0e'
        $f=New-R0eFixture; $f.Metrics[2].process=9999; Assert-R0eThrows { Test-RSR0eSample @f } 'R0e'
        $f=New-R0eFixture; $f.Metrics[2].thread=9999; Assert-R0eThrows { Test-RSR0eSample @f } 'R0e'
        $f=New-R0eFixture; $f.Metrics[2].runId='2026-11-06_104023'; Assert-R0eThrows { Test-RSR0eSample @f } 'R0e'
        $f=New-R0eFixture; $f.Metrics[2].scenario='other'; Assert-R0eThrows { Test-RSR0eSample @f } 'R0e'
        $f=New-R0eFixture; $f.Metrics[2].build='Debug'; Assert-R0eThrows { Test-RSR0eSample @f } 'R0e'
        $f=New-R0eFixture; $f.Metrics[2].commit='unexpected'; Assert-R0eThrows { Test-RSR0eSample @f } 'R0e'
        $f=New-R0eFixture; $f.Metrics[2].branch='unexpected'; Assert-R0eThrows { Test-RSR0eSample @f } 'R0e'
        $f=New-R0eFixture; $f.Metrics[0].value0=4095; Assert-R0eThrows { Test-RSR0eSample @f } 'R0e'
        $f=New-R0eFixture; $f.Metrics[0].value1=4095; Assert-R0eThrows { Test-RSR0eSample @f } 'route timing'
        $f=New-R0eFixture; $f.Metrics[0].unit='count'; Assert-R0eThrows { Test-RSR0eSample @f } 'route timing'
        $f=New-R0eFixture; $f.Native.cases[2].name='Setup'; Assert-R0eThrows { Test-RSR0eSample @f } 'duplicate'
        $f=New-R0eFixture; $f.Summary.run_id='20261106T104022Z-1-aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'; Assert-R0eThrows { Test-RSR0eSample @f } 'profile'
        $f=New-R0eFixture; $f.Summary.configuration='Debug'; Assert-R0eThrows { Test-RSR0eSample @f } 'profile'
        $f=New-R0eFixture; $f.Summary.platform='ARM64'; Assert-R0eThrows { Test-RSR0eSample @f } 'profile'
        $f=New-R0eFixture; $f.Summary.fail_fast=$true; Assert-R0eThrows { Test-RSR0eSample @f } 'profile'
        $f=New-R0eFixture; $f.Summary.total=4; Assert-R0eThrows { Test-RSR0eSample @f } 'R0e'
        $f=New-R0eFixture; $f.Summary.quarantine.active_count=$null; Assert-R0eThrows { Test-RSR0eSample @f } 'quarantine'
        $f=New-R0eFixture; $f.Summary.quarantine.repair_attempt_count=1; Assert-R0eThrows { Test-RSR0eSample @f } 'quarantine'
    }
}

Describe 'Reviewed R0e source identity' {
    It 'accepts the current exact case and rejects workload, ceiling, and diagnostic mutations' {
        $root=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
        $source=[IO.File]::ReadAllText((Join-Path $root 'RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.Tests.Phases05_06.cpp'))
        (Get-RSR0eCaseIdentity $source).full | Should Be 'A3541BB88E72E6BC41EAD4C77D37F9136A39EF123C8AAB780ECEF63039F020CE'
        $mutated=$source.Replace("kR0ePerfIterations = 4'096u", "kR0ePerfIterations = 4'095u")
        $mutated | Should Not Be $source
        Assert-R0eThrows { Get-RSR0eCaseIdentity $mutated } 'R0e'
        $mutated=$source.Replace("kR0eSelfTestCeilingUs = 2'000'000u", "kR0eSelfTestCeilingUs = 2'000'001u")
        $mutated | Should Not Be $source
        Assert-R0eThrows { Get-RSR0eCaseIdentity $mutated } 'R0e'
        $mutated=$source.Replace('parseWork={} us', 'parseWorkX={} us')
        $mutated | Should Not Be $source
        Assert-R0eThrows { Get-RSR0eCaseIdentity $mutated } 'R0e'
    }
}

function New-R0eArchivedFixture {
    $repo=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
    $root=Assert-RSTestSandboxContainedPath -TestRoot $env:REDSALAMANDER_TEST_ROOT -Path $env:REDSALAMANDER_TEST_ROOT
    $id='{0}-{1}-{2}' -f [DateTime]::UtcNow.ToString('yyyyMMddTHHmmssZ'),$PID,[Guid]::NewGuid().ToString('N')
    $f=New-R0eFixture; $f.RunId=$id; $f.Summary.run_id=$id
    $run=Join-Path $root "runs/$id"; $artifact=Join-Path $run 'artifacts/selftest/last_run'
    foreach($directory in @('fileops','perf')) { [void][IO.Directory]::CreateDirectory((Join-Path $artifact $directory)) }
    [IO.File]::WriteAllText((Join-Path $artifact 'fileops/results.json'),($f.Native|ConvertTo-Json -Depth 10))
    [IO.File]::WriteAllText((Join-Path $artifact 'fileops/trace.txt'),'fixture trace')
    [IO.File]::WriteAllLines((Join-Path $artifact 'perf/perf_metrics.jsonl'),[string[]]@($f.Metrics|ForEach-Object { $_|ConvertTo-Json -Compress }))
    $closed=Save-RSTestSandboxRunEvidence -TestRoot $root -RunId $id
    $summaryPath=Join-Path $closed.ArtifactRoot 'run-all-tests-results.json'
    [IO.File]::WriteAllText($summaryPath,($f.Summary|ConvertTo-Json -Depth 10))
    return @{root=$root;id=$id;artifact=$closed.ArtifactRoot;summaryPath=$summaryPath}
}

Describe 'R0e real archive closure' {
    BeforeAll { Import-Module (Join-Path $PSScriptRoot '../Modules/Testing/TestSandbox.psm1') -Force }
    It 'rejects a nested sandbox root before creating fixture files' {
        $oldRoot=$env:REDSALAMANDER_TEST_ROOT
        $invalid=Join-Path $oldRoot ('invalid-r0e-root-'+[Guid]::NewGuid().ToString('N'))
        try {
            $env:REDSALAMANDER_TEST_ROOT=$invalid
            Assert-R0eThrows { New-R0eArchivedFixture } 'Unauthorized TestSandbox root'
            (Test-Path -LiteralPath $invalid) | Should Be $false
        } finally { $env:REDSALAMANDER_TEST_ROOT=$oldRoot }
    }
    It 'accepts inventoried native artifacts and the separately observed post-closure summary' {
        $fixture=New-R0eArchivedFixture
        $archive=Read-RSR0eArchive -TestRoot $fixture.root -RunId $fixture.id
        $archive.files['artifacts\selftest\last_run\run-all-tests-results.json'].provenance | Should Be 'post-archive-observed'
        $summary=$archive.summary_text | ConvertFrom-Json -AsHashtable
        $sample=Test-RSR0eSample -Summary $summary -Native $archive.native -Metrics $archive.metrics -Side candidate -RunId $fixture.id -Platform x64 -RunnerExit 2
        $sample.red | Should Be $false
        $sample.process | Should Be 1234
        $sample.thread | Should Be 4321
        $sample.metric_run_id | Should Be '2026-11-06_104022'
    }
    It 'rejects corrupted and missing inventoried artifacts' {
        $fixture=New-R0eArchivedFixture
        [IO.File]::AppendAllText((Join-Path $fixture.artifact 'fileops/trace.txt'),'corrupt')
        Assert-R0eThrows { Read-RSR0eArchive -TestRoot $fixture.root -RunId $fixture.id } 'archived bytes do not match inventory'
        $fixture=New-R0eArchivedFixture
        [IO.File]::Delete((Join-Path $fixture.artifact 'fileops/trace.txt'))
        Assert-R0eThrows { Read-RSR0eArchive -TestRoot $fixture.root -RunId $fixture.id } 'R0e archived file missing'
    }
    It 'rejects malformed, wrong-identity and wrong-profile observed summaries' {
        $fixture=New-R0eArchivedFixture
        [IO.File]::WriteAllText($fixture.summaryPath,'{')
        $archive=Read-RSR0eArchive -TestRoot $fixture.root -RunId $fixture.id
        Assert-R0eThrows { $archive.summary_text | ConvertFrom-Json -AsHashtable -ErrorAction Stop } 'JSON'
        $fixture=New-R0eArchivedFixture
        $archive=Read-RSR0eArchive -TestRoot $fixture.root -RunId $fixture.id
        $summary=$archive.summary_text | ConvertFrom-Json -AsHashtable
        $summary.run_id='20261106T104022Z-1-aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa'
        Assert-R0eThrows { Test-RSR0eSample -Summary $summary -Native $archive.native -Metrics $archive.metrics -Side candidate -RunId $fixture.id -Platform x64 -RunnerExit 2 } 'profile mismatch'
        $summary.run_id=$fixture.id; $summary.timeout_scale=2
        Assert-R0eThrows { Test-RSR0eSample -Summary $summary -Native $archive.native -Metrics $archive.metrics -Side candidate -RunId $fixture.id -Platform x64 -RunnerExit 2 } 'profile mismatch'
    }
}

$nativePlatform=if ([System.Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture.ToString() -eq 'Arm64') { 'ARM64' } else { 'x64' }
Describe 'R0e command preflight' {
    It 'rejects a failed engine git status even when its HEAD matches' {
        $path=Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'Tools/Invoke-R0ePairedValidation.ps1'
        $tokens=$null; $errors=$null
        $ast=[Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$errors)
        @($errors).Count | Should Be 0
        $definition=@($ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq 'Require-CleanHead'},$true))
        $definition.Count | Should Be 1
        . ([scriptblock]::Create($definition[0].Extent.Text))
        function git { if ($args -contains 'rev-parse') { $global:LASTEXITCODE=0; 'a' * 40 } else { $global:LASTEXITCODE=1 } }
        Assert-R0eThrows { Require-CleanHead 'fixture-engine' ('a' * 40) } 'checkout is dirty or status failed'
    }
    It 'previews a clean hosted invocation without creating a cohort or changing environment' {
        $repo=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
        $command=Join-Path $repo 'Tools/Invoke-R0ePairedValidation.ps1'
        $engine=New-RSTestSandboxScratchDirectory -RepoRoot $repo -Harness R0ePester -Case EnginePreflight
        $pin=(Get-Content (Join-Path $repo '.ci/project.json') -Raw | ConvertFrom-Json).engine.ref
        $candidate='a' * 40; $baseline='b' * 40
        $cohortDir=Join-Path $env:REDSALAMANDER_TEST_ROOT 'evidence/r0e-paired'
        $before=@(Get-ChildItem -LiteralPath $cohortDir -Name -ErrorAction SilentlyContinue | Sort-Object)
        $workBefore=@(Get-ChildItem -LiteralPath (Join-Path $repo '.build') -Filter 'r0e-*' -Name -ErrorAction SilentlyContinue | Sort-Object)
        $oldActions=$env:GITHUB_ACTIONS; $oldSha=$env:GITHUB_SHA; $oldTests=$env:RSBuildEnableTests; $oldDebug=$env:RSBuildCompilerDebugInformation; $oldRoot=$env:REDSALAMANDER_TEST_ROOT
        function git {
            if ($args -contains 'rev-parse') { $global:LASTEXITCODE=0; if ($args[1] -eq $engine) { $pin } else { $candidate } }
            elseif ($args -contains 'status') { $global:LASTEXITCODE=0 }
            else { throw 'Unexpected Git operation in WhatIf.' }
        }
        try {
            $env:GITHUB_ACTIONS='true'; $env:GITHUB_SHA=$candidate
            & $command -BaselineHead $baseline -CandidateHead $candidate -Platform $nativePlatform -BuildNumber 9 -TestRoot $oldRoot -EngineRoot $engine -WhatIf
            $env:RSBuildEnableTests | Should Be $oldTests
            $env:RSBuildCompilerDebugInformation | Should Be $oldDebug
            $env:REDSALAMANDER_TEST_ROOT | Should Be $oldRoot
            (@(Get-ChildItem -LiteralPath $cohortDir -Name -ErrorAction SilentlyContinue | Sort-Object) -join '|') | Should Be ($before -join '|')
            (@(Get-ChildItem -LiteralPath (Join-Path $repo '.build') -Filter 'r0e-*' -Name -ErrorAction SilentlyContinue | Sort-Object) -join '|') | Should Be ($workBefore -join '|')
        } finally { $env:GITHUB_ACTIONS=$oldActions; $env:GITHUB_SHA=$oldSha }
    }
}

Describe 'R0e command identity admission' {
    It 'rejects unavailable hosted identity before creating a cohort or changing environment' {
        $repo=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
        $command=Join-Path $repo 'Tools/Invoke-R0ePairedValidation.ps1'
        $root=$env:REDSALAMANDER_TEST_ROOT
        $cohortDir=Join-Path $root 'evidence/r0e-paired'
        $before=@(Get-ChildItem -LiteralPath $cohortDir -Name -ErrorAction SilentlyContinue | Sort-Object)
        $oldActions=$env:GITHUB_ACTIONS; $oldSha=$env:GITHUB_SHA; $oldTests=$env:RSBuildEnableTests; $oldDebug=$env:RSBuildCompilerDebugInformation
        try {
            $env:GITHUB_ACTIONS='false'; $env:GITHUB_SHA='a' * 40
            Assert-R0eThrows { & $command -BaselineHead ('b' * 40) -CandidateHead ('a' * 40) -Platform $nativePlatform -BuildNumber 9 -TestRoot $root -EngineRoot $root } 'Hosted source identity is unavailable'
            $env:RSBuildEnableTests | Should Be $oldTests
            $env:RSBuildCompilerDebugInformation | Should Be $oldDebug
            $env:REDSALAMANDER_TEST_ROOT | Should Be $root
            (@(Get-ChildItem -LiteralPath $cohortDir -Name -ErrorAction SilentlyContinue | Sort-Object) -join '|') | Should Be ($before -join '|')
        } finally { $env:GITHUB_ACTIONS=$oldActions; $env:GITHUB_SHA=$oldSha }
    }
}

Describe 'R0e incomplete cohort retention' {
    It 'preserves run identity and log on archive failure and restores environment and location' {
        $repo=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
        $scratch=New-RSTestSandboxScratchDirectory -RepoRoot $repo -Harness R0ePester -Case ('PartialCommand-'+[guid]::NewGuid().ToString('N'))
        $script:fixtureRepo=Join-Path $scratch 'repo'; $script:fixtureEngine=Join-Path $scratch 'engine'
        foreach($dir in @($script:fixtureEngine,(Join-Path $script:fixtureRepo '.ci'),(Join-Path $script:fixtureRepo '.build'))) { [void][IO.Directory]::CreateDirectory($dir) }
        $script:fixtureCandidate='a'*40; $script:fixtureBaseline='b'*40; $script:fixtureEngineHead='c'*40; $script:partialCalls=@()
        [IO.File]::WriteAllText((Join-Path $script:fixtureRepo '.ci/project.json'),('{"engine":{"ref":"'+$script:fixtureEngineHead+'"}}'))
        $paths=@('Dependencies/DxUi.lock.json','vcpkg.json','vcpkg-tool.json','RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.Tests.Phases05_06.cpp')
        foreach($name in $paths) { $p=Join-Path $script:fixtureRepo $name; [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($p)); if($name -like '*.cpp'){[IO.File]::Copy((Join-Path $repo $name),$p)}else{[IO.File]::WriteAllText($p,'fixture')} }
        function git {
            $global:LASTEXITCODE=0
            if($args -contains 'rev-parse'){if($args[1] -eq $script:fixtureEngine){$script:fixtureEngineHead}elseif($args[1] -eq $script:fixtureRepo){$script:fixtureCandidate}else{$script:fixtureBaseline};return}
            if($args -contains 'status' -or $args -contains 'cat-file'){return}
            if($args -contains 'worktree'){$target=$args[5];foreach($name in $paths){$p=Join-Path $target $name;[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($p));[IO.File]::Copy((Join-Path $script:fixtureRepo $name),$p)};return}
            throw 'Unexpected fixture git call.'
        }
        $script:originalRepo=$repo
        $source=[IO.File]::ReadAllText((Join-Path $repo 'Tools/Invoke-R0ePairedValidation.ps1'))
        $source=$source.Replace('$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot ''..''))','$root=$script:fixtureRepo')
        foreach($name in @('Testing/R0ePairedValidation.psm1','Testing/TestSandbox.psm1','Build/BuildEvidence.psm1','Build/VcpkgInstallSafety.psm1')){$source=$source.Replace("Import-Module (Join-Path `$root 'Tools/Modules/$name')","Import-Module (Join-Path `$script:originalRepo 'Tools/Modules/$name')")}
        $marker='$cohortId=''r0e-''+[Guid]::NewGuid().ToString(''N'')'; $source.Contains($marker) | Should Be $true
        $mock=@(
            'function Read-VerifiedReceipt { return [pscustomobject]@{receipt_id=''fixture'';toolchain_identity_digest=''same'';vcpkg_identity_digest=''same'';ghostty_runtime_identity=''same'';runtime_dependency_digest=''same''} }',
            'function Initialize-RSVcpkgToolContext { return [pscustomobject]@{Executable=''fixture'';RegistryRoot=''fixture'';ToolRoot=''fixture''} }',
            'function Get-RSVcpkgBuildEnvironment { return @{} }',
            'function Invoke-R0eChild([string]$Repo,[string]$EntryPath,[string[]]$Arguments,[string]$LogPath) { $script:partialCalls+=,$EntryPath; if($EntryPath -like ''*Run-AllTests.ps1''){[IO.File]::WriteAllText($LogPath,"Run id: 20261106T104022Z-60312-10c85b8711394035839f567d80335c88`n");return 1};[IO.File]::WriteAllText($LogPath,''fixture child'');return 0 }',
            'function Read-RSR0eArchive { throw ''Injected archive validation failure.'' }'
        ) -join "`n"
        $source=$source.Replace($marker,$mock+"`n"+$marker).Replace('exit $exitCode','return $exitCode')
        $action=[scriptblock]::Create($source)
        $oldActions=$env:GITHUB_ACTIONS; $oldSha=$env:GITHUB_SHA; $oldTests=$env:RSBuildEnableTests; $oldDebug=$env:RSBuildCompilerDebugInformation; $oldRoot=$env:REDSALAMANDER_TEST_ROOT; $oldLocation=(Get-Location).Path
        $cohortDir=Join-Path $oldRoot 'evidence/r0e-paired'; $before=@(Get-ChildItem -LiteralPath $cohortDir -Name -ErrorAction SilentlyContinue)
        try {
            $env:GITHUB_ACTIONS='true';$env:GITHUB_SHA=$script:fixtureCandidate
            $output=@(& $action -BaselineHead $script:fixtureBaseline -CandidateHead $script:fixtureCandidate -Platform $nativePlatform -BuildNumber 9 -TestRoot $oldRoot -EngineRoot $script:fixtureEngine -Confirm:$false 2>&1)
            $output[-1] | Should Be 1
            $new=@(Get-ChildItem -LiteralPath $cohortDir -Name | Where-Object { $_ -notin $before });$new.Count | Should Be 1
            $cohort=Join-Path $cohortDir $new[0];$terminal=Get-Content (Join-Path $cohort 'terminal.json') -Raw | ConvertFrom-Json
            $terminal.outcome | Should Be 'INCOMPLETE';$terminal.error | Should Be 'Injected archive validation failure.'
            $terminal.current_run_id | Should Be '20261106T104022Z-60312-10c85b8711394035839f567d80335c88'
            $identified=Get-Content (Join-Path $cohort 'baseline-01-identified.json') -Raw | ConvertFrom-Json
            $identified.run_id | Should Be $terminal.current_run_id;(Test-Path -LiteralPath $identified.log) | Should Be $true
            [IO.File]::ReadAllText($identified.log) | Should Match 'Run id:'
            @($script:partialCalls | Where-Object { $_ -like '*Run-AllTests.ps1' }).Count | Should Be 1
            ([string]$env:RSBuildEnableTests) | Should Be ([string]$oldTests);([string]$env:RSBuildCompilerDebugInformation) | Should Be ([string]$oldDebug);$env:REDSALAMANDER_TEST_ROOT | Should Be $oldRoot;(Get-Location).Path | Should Be $oldLocation
        } finally {$env:GITHUB_ACTIONS=$oldActions;$env:GITHUB_SHA=$oldSha}
    }
}
