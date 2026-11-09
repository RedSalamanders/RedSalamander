Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$repoRoot=Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestRunSummary.psm1') -Force
Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/ValidationFingerprint.psm1')
Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestSandbox.psm1')
Describe 'Exact canonical runner summary handoff' {
    It 'binds the new schema and transport module into runner dependency identity' {
        $closure=@(Get-RSValidationDependencyClosure -RepoRoot $repoRoot -RootRelativePath 'Tools/Run-AllTests.ps1')
        ($closure -contains 'Specs/Testing/RunResultReference.schema.json') | Should Be $true
        ($closure -contains 'Tools/Modules/Testing/TestRunSummary.psm1') | Should Be $true
    }
    BeforeEach {
        $testRoot = Get-RSTestSandboxRoot -RepoRoot $repoRoot
        $fixture=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness tools-pester -Case ('runner-result-'+[guid]::NewGuid().ToString('N'))
        $summaryPath=Join-Path $fixture 'run-all-tests-results.json'
        $referencePath=Join-Path $fixture 'reference.json'
        $summary=[ordered]@{
            schema='red-salamander.run-all-tests.v2';run_id=('20261109T120000Z-1-'+[guid]::NewGuid().ToString('N'))
            validation_mode='Fresh';requested_gate='CI';coverage_scope='complete'
            snapshot_id=('a'*64);build_receipt_id=('b'*64);plan_digest=('c'*64);coverage_plan_digest=('d'*64);decision_digest=('e'*64)
            repository_verdict='PASSED';change_verdict='PASSED'
            counts=@{executed=1;reused=0;provisional=0;promoted=1;invalidated=0;incomplete=0}
            case_counts=@{passed=1;failed=0;skipped=0;total=1}
            duration_ms=1;estimated_avoided_duration_ms=0;commands_broad_seal_state='absent';entries=@()
        }
        $schema=Join-Path $repoRoot 'Specs/Testing/RunAllTestsSummary.schema.json'
        Write-RSAtomicCanonicalJsonFile -LiteralPath $summaryPath -InputObject $summary -SchemaPath $schema -AllowedRoot $fixture -CreateNew
        Write-RSRunResultReference -Path $referencePath -SummaryPath $summaryPath -RepoRoot $repoRoot -TestRoot $testRoot -Platform x64 -Configuration Release
    }
    It 'selects exactly the emitted run regardless of unrelated archive summaries' {
        Copy-Item -LiteralPath $summaryPath -Destination (Join-Path $fixture 'other-summary.json')
        $result=Read-RSRunResultReference -Path $referencePath -RepoRoot $repoRoot -TestRoot $testRoot -Platform x64 -Configuration Release
        $result.Path | Should Be $summaryPath
        $result.Summary.run_id | Should Be $summary.run_id
    }
    It 'refuses overwrite of an existing invocation reference' {
        $threw=$false
        try { Write-RSRunResultReference -Path $referencePath -SummaryPath $summaryPath -RepoRoot $repoRoot -TestRoot $testRoot -Platform x64 -Configuration Release }
        catch [IO.IOException] {$threw=$true}
        $threw | Should Be $true
    }
    It 'refuses missing references, wrong architecture, configuration and gate' {
        foreach($arguments in @(
            @{Path=(Join-Path $fixture 'missing.json');Platform='x64';Configuration='Release';RequestedGate='CI'},
            @{Path=$referencePath;Platform='ARM64';Configuration='Release';RequestedGate='CI'},
            @{Path=$referencePath;Platform='x64';Configuration='Debug';RequestedGate='CI'},
            @{Path=$referencePath;Platform='x64';Configuration='Release';RequestedGate='Full'}
        )) {
            $threw=$false
            try { Read-RSRunResultReference @arguments -RepoRoot $repoRoot -TestRoot $testRoot | Out-Null }
            catch [System.Management.Automation.RuntimeException] {$threw=$true}
            $threw | Should Be $true
        }
    }
    It 'refuses changed bytes and an internally consistent wrong run reference' {
        $summary.duration_ms=2
        Write-RSAtomicCanonicalJsonFile -LiteralPath $summaryPath -InputObject $summary -SchemaPath $schema -AllowedRoot $fixture
        $threw=$false
        try { Read-RSRunResultReference -Path $referencePath -RepoRoot $repoRoot -TestRoot $testRoot -Platform x64 -Configuration Release }
        catch [System.Management.Automation.RuntimeException] {$threw=$true}
        $threw | Should Be $true
        $record=Get-Content $referencePath -Raw|ConvertFrom-Json
        $record.summary_sha256=(Get-FileHash -LiteralPath $summaryPath).Hash.ToLowerInvariant()
        $record.run_id='20261109T120000Z-2-'+[guid]::NewGuid().ToString('N')
        Write-RSAtomicCanonicalJsonFile -LiteralPath $referencePath -InputObject $record -AllowedRoot $fixture -SchemaPath (Join-Path $repoRoot 'Specs/Testing/RunResultReference.schema.json')
        $threw=$false
        try { Read-RSRunResultReference -Path $referencePath -RepoRoot $repoRoot -TestRoot $testRoot -Platform x64 -Configuration Release }
        catch [System.Management.Automation.RuntimeException] {$threw=$true}
        $threw | Should Be $true
    }
    It 'refuses a validly hashed incomplete or unsuccessful summary' {
        foreach($variant in @('empty','all-skipped','no-execution','reused','inconsistent-counts','failed','filtered','resume','incomplete','invalidated')) {
            $summary.case_counts.total=1;$summary.case_counts.passed=1;$summary.case_counts.skipped=0
            $summary.counts.executed=1;$summary.counts.reused=0
            $summary.repository_verdict='PASSED';$summary.coverage_scope='complete';$summary.validation_mode='Fresh'
            $summary.counts.incomplete=0;$summary.counts.invalidated=0
            switch($variant) {
                empty {$summary.case_counts.total=0}
                all-skipped {$summary.case_counts.passed=0;$summary.case_counts.skipped=1}
                no-execution {$summary.counts.executed=0}
                reused {$summary.counts.reused=1}
                inconsistent-counts {$summary.case_counts.total=2}
                failed {$summary.repository_verdict='FAILED'}
                filtered {$summary.coverage_scope='filtered'}
                resume {$summary.validation_mode='Resume'}
                incomplete {$summary.counts.incomplete=1}
                invalidated {$summary.counts.invalidated=1}
            }
            Write-RSAtomicCanonicalJsonFile -LiteralPath $summaryPath -InputObject $summary -SchemaPath $schema -AllowedRoot $fixture
            $ref=Join-Path $fixture ($variant+'.json')
            Write-RSRunResultReference -Path $ref -SummaryPath $summaryPath -RepoRoot $repoRoot -TestRoot $testRoot -Platform x64 -Configuration Release
            $threw=$false
            try { Read-RSRunResultReference -Path $ref -RepoRoot $repoRoot -TestRoot $testRoot -Platform x64 -Configuration Release | Out-Null }
            catch [System.Management.Automation.RuntimeException] {$threw=$true}
            $threw | Should Be $true
        }
    }
    It 'refuses escaping paths even with a well-formed reference' {
        $record=Get-Content $referencePath -Raw|ConvertFrom-Json
        $record.summary_path='../outside.json'
        Write-RSAtomicCanonicalJsonFile -LiteralPath $referencePath -InputObject $record -AllowedRoot $fixture -SchemaPath (Join-Path $repoRoot 'Specs/Testing/RunResultReference.schema.json')
        $threw=$false
        try { Read-RSRunResultReference -Path $referencePath -RepoRoot $repoRoot -TestRoot $testRoot -Platform x64 -Configuration Release | Out-Null }
        catch [System.Management.Automation.RuntimeException] {$threw=$true}
        $threw | Should Be $true
    }
}
