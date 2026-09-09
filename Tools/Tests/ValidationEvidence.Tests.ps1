$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'Tools\Modules\Testing\ValidationFingerprint.psm1') -Force
Import-Module (Join-Path $repoRoot 'Tools\Modules\Testing\TestRunPlan.psm1') -Force
Import-Module (Join-Path $repoRoot 'Tools\Modules\Testing\ValidationEvidence.psm1') -Force
# ValidationEvidence imports ValidationFingerprint into module scope. Import it
# last as well so fixture construction can compute externally visible digests.
Import-Module (Join-Path $repoRoot 'Tools\Modules\Testing\ValidationFingerprint.psm1') -Force

function Test-RSEvidenceActionThrows([scriptblock]$Action) {
    try { & $Action | Out-Null; return $false } catch { return $true }
}

function Get-RSEvidenceActionErrorMessage([scriptblock]$Action) {
    try { & $Action | Out-Null; return $null } catch { return $_.Exception.Message }
}

function Update-RSEvidenceDecisionDigest([object]$Decision) {
    $projection = [ordered]@{}
    foreach ($property in $Decision.PSObject.Properties) {
        if ($property.Name -ne 'decision_digest') { $projection[$property.Name] = $property.Value }
    }
    $Decision.decision_digest = Get-RSCanonicalJsonDigest -InputObject $projection
    return $Decision
}

function New-RSEvidenceFixture([string]$Root, [string]$RunId = '20260813T120000Z-42-0123456789abcdef0123456789abcdef') {
    $execution = New-RSTestRunPlanEntry -Id 'tooling.pester' -Name 'Tooling' -Kind Pester `
        -Path (Join-Path $repoRoot 'Tools\Tests')
    $plan = New-RSValidationPlan -RepoRoot $repoRoot -RequestedSuite Full -Entries @($execution)
    $entry = $plan.entries[0]
    $snapshot = [pscustomobject]@{ snapshot_id = 'a' * 64 }
    $fingerprint = 'b' * 64
    $decision = [pscustomobject][ordered]@{
        schema = 'red-salamander.validation-plan-decision.v1'; canonicalization = 'rs-canonical-json-v1'
        decision_digest = 'c' * 64; planner_mode = 'shadow'; snapshot_id = $snapshot.snapshot_id
        plan_digest = $plan.plan_digest; changed_paths = @(); matched_impact_rules = @(); advisory_build_action = 'none'
        entries = @([pscustomobject][ordered]@{
                id = $entry.id; ordinal = 0L; affected = $false; matched_impact_tags = @(); decision = 'EXECUTE'
                reason_codes = @('FORCED'); expected_fingerprint = $fingerprint; origin_run_id = $null; origin_evidence_digest = $null
            })
        estimated_fresh_duration_ms = 0L; estimated_executed_duration_ms = 0L; estimated_avoided_duration_ms = 0L
    }
    $decision = Update-RSEvidenceDecisionDigest -Decision $decision
    $receipt = [pscustomobject][ordered]@{
        receipt_id = 'd' * 64
        outputs = @([pscustomobject][ordered]@{ id = 'artifact.tooling'; sha256 = 'e' * 64 })
    }
    $run = New-RSValidationEvidenceRun -EvidenceRoot $Root -RunId $RunId -Plan $plan `
        -Decision $decision -BuildReceipt $receipt -SchemaRoot (Join-Path $repoRoot 'Specs\Testing')
    return [pscustomobject]@{
        Run = $run; Plan = $plan; Decision = $decision; Entry = $entry
        Snapshot = $snapshot; Fingerprint = $fingerprint; Receipt = $receipt
    }
}

function New-RSEvidencePassingResult {
    return [pscustomobject]@{
        ExitCode = 0; Passed = 1; Failed = 0; Skipped = 0; Total = 1
        Pending = 0; Inconclusive = 0; Classification = 'PASSED'
    }
}

function New-RSEvidenceTerminalResult([object]$Fixture, [object]$Result) {
    return New-RSValidationTerminalResult -EntryContract $Fixture.Entry -Result $Result `
        -ResultParsed $true -CoverageValid $true -RequiredArtifactsPresent $true
}

function Publish-RSEvidenceFixturePass([object]$Fixture, [AllowNull()][object]$FingerprintComponents = $null) {
    Set-RSValidationEntryCheckpoint -Run $Fixture.Run -EntryId $Fixture.Entry.id -Status running
    $result = New-RSEvidencePassingResult
    return Publish-RSValidationEntryEvidence -Run $Fixture.Run -EntryContract $Fixture.Entry `
        -EntryFingerprint $Fixture.Fingerprint `
        -EntryFingerprintComponents $FingerprintComponents `
        -TerminalResult (New-RSEvidenceTerminalResult -Fixture $Fixture -Result $result) `
        -PreSnapshot $Fixture.Snapshot -PostSnapshot $Fixture.Snapshot -BuildReceipt $Fixture.Receipt `
        -StartedUtc ([datetime]'2026-08-13T12:00:00Z') -EndedUtc ([datetime]'2026-08-13T12:00:01Z')
}

function New-RSEvidenceFingerprintComponents {
    $components = [ordered]@{}
    $names = @(
        'build_artifact', 'logical_input', 'shared_contract', 'runner', 'environment',
        'arguments', 'coverage', 'outcome_policy', 'capability', 'artifact_epoch', 'source_input')
    for ($index = 0; $index -lt $names.Count; $index++) {
        $components[$names[$index]] = ('{0:x}' -f ($index + 1)) * 64
    }
    return [pscustomobject]$components
}

Describe 'Operation Startrail run-directory resolution' {
    It 'accepts only the exact direct child whose leaf matches canonical run state' {
        $fixture = New-RSEvidenceFixture -Root (Join-Path $TestDrive 'run-resolver')
        $schemaRoot = Join-Path $repoRoot 'Specs\Testing'
        $byId = Resolve-RSValidationRunDirectory -EvidenceRoot (Join-Path $TestDrive 'run-resolver') `
            -RunReference $fixture.Run.State.run_id -SchemaRoot $schemaRoot
        $byPath = Resolve-RSValidationRunDirectory -EvidenceRoot (Join-Path $TestDrive 'run-resolver') `
            -RunReference $fixture.Run.RunDirectory -SchemaRoot $schemaRoot
        $byId.RunDirectory | Should Be $fixture.Run.RunDirectory
        $byPath.State.run_id | Should Be $fixture.Run.State.run_id

        (Test-RSEvidenceActionThrows {
                Resolve-RSValidationRunDirectory -EvidenceRoot (Join-Path $TestDrive 'run-resolver') `
                    -RunReference ($fixture.Run.RunDirectory + '\.') -SchemaRoot $schemaRoot
            }) | Should Be $true
        (Test-RSEvidenceActionThrows {
                Resolve-RSValidationRunDirectory -EvidenceRoot (Join-Path $TestDrive 'run-resolver') `
                    -RunReference $fixture.Run.RunDirectory.ToUpperInvariant() -SchemaRoot $schemaRoot
            }) | Should Be $true
    }

    It 'rejects nested, copied-renamed, and junction run directories' {
        $evidenceRoot = Join-Path $TestDrive 'run-resolver-attacks'
        $fixture = New-RSEvidenceFixture -Root $evidenceRoot
        $schemaRoot = Join-Path $repoRoot 'Specs\Testing'
        $runsRoot = Join-Path $evidenceRoot 'runs'

        $nestedParent = Join-Path $runsRoot 'nested'
        [void](New-Item -ItemType Directory -Path $nestedParent)
        $nested = Join-Path $nestedParent $fixture.Run.State.run_id
        Copy-Item -LiteralPath $fixture.Run.RunDirectory -Destination $nested -Recurse
        (Test-RSEvidenceActionThrows {
                Resolve-RSValidationRunDirectory -EvidenceRoot $evidenceRoot -RunReference $nested -SchemaRoot $schemaRoot
            }) | Should Be $true

        $renamedId = '20260813T120001Z-42-1123456789abcdef0123456789abcdef'
        $renamed = Join-Path $runsRoot $renamedId
        Copy-Item -LiteralPath $fixture.Run.RunDirectory -Destination $renamed -Recurse
        (Test-RSEvidenceActionThrows {
                Resolve-RSValidationRunDirectory -EvidenceRoot $evidenceRoot -RunReference $renamed -SchemaRoot $schemaRoot
            }) | Should Be $true

        $junctionId = '20260813T120002Z-42-2123456789abcdef0123456789abcdef'
        $junction = Join-Path $runsRoot $junctionId
        [void](New-Item -ItemType Junction -Path $junction -Target $fixture.Run.RunDirectory)
        (Test-RSEvidenceActionThrows {
                Resolve-RSValidationRunDirectory -EvidenceRoot $evidenceRoot -RunReference $junction -SchemaRoot $schemaRoot
            }) | Should Be $true
    }
}

Describe 'Operation Startrail result-artifact publication' {
    It 'rejects aggregate entry and run quota crossings before publication' {
        $fixture = New-RSEvidenceFixture -Root (Join-Path $TestDrive 'artifact-quotas')
        $first = Join-Path $TestDrive 'first-40m.bin'
        $second = Join-Path $TestDrive 'second-40m.bin'
        foreach ($path in @($first, $second)) {
            $stream = [IO.FileStream]::new($path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
            try { $stream.SetLength(40MB) } finally { $stream.Dispose() }
        }
        (Test-RSEvidenceActionThrows {
                Publish-RSValidationResultArtifacts -Run $fixture.Run -EntryId $fixture.Entry.id `
                    -CandidatePath @($first, $second)
            }) | Should Be $true

        $small = Join-Path $TestDrive 'run-quota.bin'
        [IO.File]::WriteAllBytes($small, [byte[]]::new(16))
        $fixture.Run.ResultArtifactBytes = 1GB - 8
        (Test-RSEvidenceActionThrows {
                Publish-RSValidationResultArtifacts -Run $fixture.Run -EntryId $fixture.Entry.id `
                    -CandidatePath @($small)
            }) | Should Be $true
        @(Get-ChildItem -LiteralPath (Join-Path $fixture.Run.RunDirectory 'entries') -Recurse -File).Count | Should Be 0
    }

    It 'publishes atomic content bindings and rejects source mutation during preservation' {
        $fixture = New-RSEvidenceFixture -Root (Join-Path $TestDrive 'artifact-source-mutation')
        $source = Join-Path $TestDrive 'mutable-result.json'
        [IO.File]::WriteAllText($source, '{"passed":true}', [Text.UTF8Encoding]::new($false))
        $hook = {
            param([string]$SourcePath, [string]$TemporaryPath)
            [IO.File]::AppendAllText($SourcePath, 'changed', [Text.UTF8Encoding]::new($false))
        }
        (Test-RSEvidenceActionThrows {
                Publish-RSValidationResultArtifacts -Run $fixture.Run -EntryId $fixture.Entry.id `
                    -CandidatePath @($source) -AfterCopyTestHook $hook
            }) | Should Be $true
        $resultRoot = Join-Path $fixture.Run.RunDirectory "entries\$($fixture.Entry.id)\results"
        @(Get-ChildItem -LiteralPath $resultRoot -File -ErrorAction SilentlyContinue).Count | Should Be 0

        [IO.File]::WriteAllText($source, '{"passed":true}', [Text.UTF8Encoding]::new($false))
        $record = @(Publish-RSValidationResultArtifacts -Run $fixture.Run -EntryId $fixture.Entry.id `
                -CandidatePath @($source))
        $record.Count | Should Be 1
        $record[0].size_bytes | Should Be (Get-Item $source).Length
        $record[0].sha256 | Should Be ((Get-FileHash $source -Algorithm SHA256).Hash.ToLowerInvariant())
        (Get-FileHash (Join-Path $fixture.Run.RunDirectory $record[0].path) -Algorithm SHA256).Hash.ToLowerInvariant() | Should Be $record[0].sha256
    }

    It 'rejects destination truncation before and after promotion' {
        foreach ($phase in @('before', 'after')) {
            $root = Join-Path $TestDrive "artifact-tamper-$phase"
            $fixture = New-RSEvidenceFixture -Root $root
            $source = Join-Path $root 'result.json'
            [IO.File]::WriteAllText($source, '{"passed":true}', [Text.UTF8Encoding]::new($false))
            $records = @(Publish-RSValidationResultArtifacts -Run $fixture.Run -EntryId $fixture.Entry.id `
                    -CandidatePath @($source))
            $destination = Join-Path $fixture.Run.RunDirectory $records[0].path
            Set-RSValidationEntryCheckpoint -Run $fixture.Run -EntryId $fixture.Entry.id -Status running
            $result = New-RSEvidencePassingResult
            $terminal = New-RSEvidenceTerminalResult -Fixture $fixture -Result $result
            if ($phase -eq 'before') {
                [IO.File]::WriteAllBytes($destination, [byte[]]::new(1))
                (Test-RSEvidenceActionThrows {
                        Publish-RSValidationEntryEvidence -Run $fixture.Run -EntryContract $fixture.Entry `
                            -EntryFingerprint $fixture.Fingerprint -TerminalResult $terminal `
                            -PreSnapshot $fixture.Snapshot -PostSnapshot $fixture.Snapshot -BuildReceipt $fixture.Receipt `
                            -StartedUtc ([datetime]'2026-08-13T12:00:00Z') -EndedUtc ([datetime]'2026-08-13T12:00:01Z') `
                            -ResultArtifact $records -RequiredArtifact @($records.path)
                    }) | Should Be $true
                continue
            }
            $published = Publish-RSValidationEntryEvidence -Run $fixture.Run -EntryContract $fixture.Entry `
                -EntryFingerprint $fixture.Fingerprint -TerminalResult $terminal `
                -PreSnapshot $fixture.Snapshot -PostSnapshot $fixture.Snapshot -BuildReceipt $fixture.Receipt `
                -StartedUtc ([datetime]'2026-08-13T12:00:00Z') -EndedUtc ([datetime]'2026-08-13T12:00:01Z') `
                -ResultArtifact $records -RequiredArtifact @($records.path)
            (Complete-RSValidationEvidenceRun -Run $fixture.Run) | Should Be 'passed'
            [IO.File]::WriteAllBytes($destination, [byte[]]::new(1))
            (Test-RSEvidenceActionThrows {
                    Read-RSVerifiedValidationPromotionEvidence -OriginRunDirectory $fixture.Run.RunDirectory `
                        -EntryId $fixture.Entry.id -PromotionDigest $published.Promotion.promotion_digest `
                        -SchemaRoot (Join-Path $repoRoot 'Specs\Testing') -RequireTerminalGreen
                }) | Should Be $true
        }
    }
}

Describe 'Normalized validation terminal results' {
    BeforeAll {
        $selfTestExecution = New-RSTestRunPlanEntry -Id 'selftest.example' -Name 'Example' -Kind SelfTest `
            -Path (Join-Path $repoRoot 'RedSalamander.exe') -JsonName 'example'
        $selfTestContract = (New-RSValidationPlan -RepoRoot $repoRoot -RequestedSuite Full `
                -Entries @($selfTestExecution)).entries[0]
    }

    It 'accepts one complete passing selftest result' {
        $result = [pscustomobject]@{
            ExitCode = 0; Passed = 1; Failed = 0; Skipped = 0; Total = 1
            Classification = 'PASSED'; Cases = @([pscustomobject]@{ name = 'case.one'; status = 'passed' })
        }
        $terminal = New-RSValidationTerminalResult -EntryContract $selfTestContract -Result $result `
            -ResultParsed $true -CoverageValid $true -RequiredArtifactsPresent $true
        $terminal.status | Should Be 'PASSED'
        @($terminal.reason_codes).Count | Should Be 0
        ($terminal | ConvertTo-Json -Depth 20 | Test-Json `
                -SchemaFile (Join-Path $repoRoot 'Specs\Testing\ValidationTerminalResult.schema.json')) | Should Be $true
    }

    It 'fails closed for missing parse, no cases, zero Pester discovery, and missing artifacts' {
        $emptySelfTest = [pscustomobject]@{
            ExitCode = 0; Passed = 0; Failed = 0; Skipped = 0; Total = 0
            Classification = 'PASSED'; Cases = @()
        }
        $terminal = New-RSValidationTerminalResult -EntryContract $selfTestContract -Result $emptySelfTest `
            -ResultParsed $false -CoverageValid $false -RequiredArtifactsPresent $false
        $terminal.status | Should Be 'FAILED'
        (@($terminal.reason_codes) -contains 'RESULT_NOT_PARSED') | Should Be $true
        (@($terminal.reason_codes) -contains 'NO_CASE_RESULTS') | Should Be $true
        (@($terminal.reason_codes) -contains 'NO_TESTS_DISCOVERED') | Should Be $true
        (@($terminal.reason_codes) -contains 'COVERAGE_INVALID') | Should Be $true
        (@($terminal.reason_codes) -contains 'REQUIRED_ARTIFACT_MISSING') | Should Be $true

        $pesterFixture = New-RSEvidenceFixture -Root (Join-Path $TestDrive 'zero-pester')
        $zeroPester = [pscustomobject]@{ ExitCode = 0; Passed = 0; Failed = 0; Skipped = 0; Total = 0; Classification = 'PASSED' }
        $pesterTerminal = New-RSValidationTerminalResult -EntryContract $pesterFixture.Entry -Result $zeroPester `
            -ResultParsed $true -CoverageValid $true -RequiredArtifactsPresent $true
        $pesterTerminal.status | Should Be 'FAILED'
        (@($pesterTerminal.reason_codes) -contains 'NO_TESTS_DISCOVERED') | Should Be $true
    }

    It 'rejects all-skipped, unallowed-skip, and aggregate count mismatches' {
        $skipped = [pscustomobject]@{ name = 'case.skip'; status = 'skipped'; skip_class = 'unexpected' }
        $result = [pscustomobject]@{
            ExitCode = 0; Passed = 0; Failed = 0; Skipped = 1; Total = 1
            Classification = 'PASSED'; Cases = @($skipped)
        }
        $terminal = New-RSValidationTerminalResult -EntryContract $selfTestContract -Result $result `
            -ResultParsed $true -CoverageValid $true -RequiredArtifactsPresent $true
        (@($terminal.reason_codes) -contains 'ALL_TESTS_SKIPPED') | Should Be $true
        (@($terminal.reason_codes) -contains 'UNALLOWED_SKIP') | Should Be $true
        (@($terminal.reason_codes) -contains 'FORBIDDEN_SKIP_CLASS') | Should Be $true

        $mismatched = [pscustomobject]@{
            ExitCode = 0; Passed = 2; Failed = 0; Skipped = 0; Total = 2
            Classification = 'PASSED'; Cases = @([pscustomobject]@{ name = 'case.one'; status = 'passed' })
        }
        $mismatchTerminal = New-RSValidationTerminalResult -EntryContract $selfTestContract -Result $mismatched `
            -ResultParsed $true -CoverageValid $true -RequiredArtifactsPresent $true
        $mismatchTerminal.status | Should Be 'FAILED'
        (@($mismatchTerminal.reason_codes) -contains 'COUNT_MISMATCH') | Should Be $true
    }
}

Describe 'Operation Startrail provisional evidence and promotion barriers' {
    It 'normalizes explicitly passed empty terminal-artifact collections' {
        $fixture = New-RSEvidenceFixture -Root (Join-Path $TestDrive 'empty-terminal-artifacts')
        Set-RSValidationEntryCheckpoint -Run $fixture.Run -EntryId $fixture.Entry.id -Status running
        # A pipeline function with no output can bind through a typed array
        # parameter as a single empty-string sentinel in strict-mode callers.
        $emptyArtifacts = @('')
        $result = New-RSEvidencePassingResult
        $published = Publish-RSValidationEntryEvidence -Run $fixture.Run -EntryContract $fixture.Entry `
            -EntryFingerprint $fixture.Fingerprint `
            -TerminalResult (New-RSEvidenceTerminalResult -Fixture $fixture -Result $result) `
            -PreSnapshot $fixture.Snapshot -PostSnapshot $fixture.Snapshot -BuildReceipt $fixture.Receipt `
            -StartedUtc ([datetime]'2026-08-13T12:00:00Z') -EndedUtc ([datetime]'2026-08-13T12:00:01Z') `
            -ResultArtifact $emptyArtifacts -RequiredArtifact $emptyArtifacts
        @($published.Evidence.result_artifacts).Count | Should Be 0
        $published.Evidence.status | Should Be 'provisional-passed'
        $fixture.Run.State.entries[0].status | Should Be 'promoted'
    }

    It 'persists every state transition and promotes only after a matching post-entry barrier' {
        $fixture = New-RSEvidenceFixture -Root (Join-Path $TestDrive 'evidence')
        Set-RSValidationEntryCheckpoint -Run $fixture.Run -EntryId $fixture.Entry.id -Status running
        $storedRunning = Get-Content (Join-Path $fixture.Run.RunDirectory 'run-state.json') -Raw | ConvertFrom-Json
        $storedRunning.entries[0].status | Should Be 'running'

        $artifact = Join-Path $TestDrive 'tooling-result.json'
        [IO.File]::WriteAllText($artifact, '{}', [Text.UTF8Encoding]::new($false))
        $preserved = @(Publish-RSValidationResultArtifacts -Run $fixture.Run `
                -EntryId $fixture.Entry.id -CandidatePath @($artifact))
        $start = [datetime]'2026-08-13T12:00:00Z'
        $result = New-RSEvidencePassingResult
        $published = Publish-RSValidationEntryEvidence -Run $fixture.Run -EntryContract $fixture.Entry `
            -EntryFingerprint $fixture.Fingerprint `
            -TerminalResult (New-RSEvidenceTerminalResult -Fixture $fixture -Result $result) `
            -PreSnapshot $fixture.Snapshot -PostSnapshot $fixture.Snapshot -BuildReceipt $fixture.Receipt `
            -StartedUtc $start -EndedUtc $start.AddSeconds(1) -ResultArtifact $preserved `
            -RequiredArtifact @($preserved.path)
        $published.Evidence.status | Should Be 'provisional-passed'
        $published.Promotion | Should Not BeNullOrEmpty
        $fixture.Run.State.entries[0].status | Should Be 'promoted'
        (Complete-RSValidationEvidenceRun -Run $fixture.Run) | Should Be 'passed'
        (Get-Content $published.EvidencePath -Raw | Test-Json -SchemaFile (Join-Path $repoRoot 'Specs\Testing\ValidationEntryEvidence.schema.json')) | Should Be $true
        (Get-Content $published.PromotionPath -Raw | Test-Json -SchemaFile (Join-Path $repoRoot 'Specs\Testing\ValidationPromotion.schema.json')) | Should Be $true
    }

    It 'does not promote failures, source mutation, or missing required terminal artifacts' {
        $cases = @(
            [pscustomobject]@{ Name = 'failure'; Result = [pscustomobject]@{ ExitCode = 1; Passed = 0; Failed = 1; Skipped = 0; Total = 1; Classification = 'REGRESSION' }; Post = 'a' * 64; Required = @() },
            [pscustomobject]@{ Name = 'source'; Result = (New-RSEvidencePassingResult); Post = 'f' * 64; Required = @() },
            [pscustomobject]@{ Name = 'artifact'; Result = (New-RSEvidencePassingResult); Post = 'a' * 64; Required = @('results/missing.json') }
        )
        $index = 0
        foreach ($case in $cases) {
            ++$index
            $runId = "20260813T12000${index}Z-42-0123456789abcdef0123456789abcde${index}"
            $fixture = New-RSEvidenceFixture -Root (Join-Path $TestDrive $case.Name) -RunId $runId
            Set-RSValidationEntryCheckpoint -Run $fixture.Run -EntryId $fixture.Entry.id -Status running
            $post = [pscustomobject]@{ snapshot_id = $case.Post }
            $published = Publish-RSValidationEntryEvidence -Run $fixture.Run -EntryContract $fixture.Entry `
                -EntryFingerprint $fixture.Fingerprint `
                -TerminalResult (New-RSEvidenceTerminalResult -Fixture $fixture -Result $case.Result) `
                -PreSnapshot $fixture.Snapshot `
                -PostSnapshot $post -BuildReceipt $fixture.Receipt -StartedUtc ([datetime]'2026-08-13T12:00:00Z') `
                -EndedUtc ([datetime]'2026-08-13T12:00:01Z') -RequiredArtifact $case.Required
            $published.Promotion | Should BeNullOrEmpty
            $fixture.Run.State.entries[0].status | Should Be 'failed'
            (Complete-RSValidationEvidenceRun -Run $fixture.Run) | Should Be 'failed'
        }
    }

    It 'repairs a dead running checkpoint as interrupted and never promoted' {
        $fixture = New-RSEvidenceFixture -Root (Join-Path $TestDrive 'interrupted')
        Set-RSValidationEntryCheckpoint -Run $fixture.Run -EntryId $fixture.Entry.id -Status running
        (Repair-RSValidationInterruptedRunState -Run $fixture.Run) | Should Be $true
        $fixture.Run.State.status | Should Be 'interrupted'
        $fixture.Run.State.entries[0].status | Should Be 'interrupted'
        $fixture.Run.State.entries[0].promotion_digest | Should BeNullOrEmpty
    }

    It 'blocks readers during mutation and switches atomically to an immutable epoch decision' {
        $fixture = New-RSEvidenceFixture -Root (Join-Path $TestDrive 'epochs')
        $fixture.Entry.artifact_read_ids = @('artifact.profile')
        $fixture.Run.State.entries[0].status = 'promoted'
        $fixture.Run.State.entries[0].promotion_digest = 'f' * 64
        $nextReceipt = [pscustomobject]@{ receipt_id = '1' * 64 }
        $initialDecisionBytes = [IO.File]::ReadAllBytes($fixture.Run.DecisionPath)
        (Start-RSValidationArtifactMutation -Run $fixture.Run `
            -Plan ([pscustomobject]@{ entries = @($fixture.Entry) }) `
            -WrittenArtifactId @('artifact.profile')) | Should Be 1
        $fixture.Run.State.entries[0].status | Should Be 'invalidated'
        $fixture.Run.State.entries[0].reason_codes | Should Be @('ARTIFACT_EPOCH_CHANGED')
        $fixture.Run.State.mutation_state | Should Be 'mutating'
        (Test-RSValidationArtifactReadBarrier -Run $fixture.Run -BuildReceipt $fixture.Receipt -ExpectedEpoch 1) | Should Be $false

        $nextDecision = $fixture.Decision | ConvertTo-Json -Depth 30 | ConvertFrom-Json
        $nextDecision.entries[0].expected_fingerprint = '9' * 64
        $nextDecision = Update-RSEvidenceDecisionDigest -Decision $nextDecision
        (Complete-RSValidationArtifactMutation -Run $fixture.Run -Decision $nextDecision `
            -BuildReceipt $nextReceipt -MutatorEntryId $fixture.Entry.id -ExpectedEpoch 1) | Should Be 1
        $fixture.Run.State.mutation_state | Should Be 'stable'
        $fixture.Run.State.decision_digest | Should Be $nextDecision.decision_digest
        (Test-RSValidationArtifactReadBarrier -Run $fixture.Run -BuildReceipt $nextReceipt -ExpectedEpoch 1) | Should Be $true
        (Test-RSValidationArtifactReadBarrier -Run $fixture.Run -BuildReceipt $fixture.Receipt -ExpectedEpoch 0) | Should Be $false
        @((Get-ChildItem -LiteralPath (Join-Path $fixture.Run.RunDirectory 'decisions') -File)).Count | Should Be 2
        [Convert]::ToHexString([IO.File]::ReadAllBytes((Join-Path $fixture.Run.RunDirectory "decisions\$($fixture.Decision.decision_digest).json"))) |
            Should Be ([Convert]::ToHexString($initialDecisionBytes))
        Test-Path -LiteralPath (Join-Path $fixture.Run.RunDirectory 'validation-plan-decision.json') | Should Be $false
    }

    It 'leaves a crashed mutation unreadable and refuses replacement of an immutable decision' {
        $fixture = New-RSEvidenceFixture -Root (Join-Path $TestDrive 'epoch-crash')
        [void](Start-RSValidationArtifactMutation -Run $fixture.Run `
                -Plan $fixture.Plan -WrittenArtifactId @('artifact.tooling'))
        (Repair-RSValidationInterruptedRunState -Run $fixture.Run) | Should Be $true
        $fixture.Run.State.mutation_state | Should Be 'interrupted'
        (Test-RSValidationArtifactReadBarrier -Run $fixture.Run -BuildReceipt $fixture.Receipt -ExpectedEpoch 1) | Should Be $false

        [IO.File]::WriteAllText($fixture.Run.DecisionPath, '{}', [Text.UTF8Encoding]::new($false))
        (Test-RSEvidenceActionThrows {
                Publish-RSValidationDecisionRecord -RunDirectory $fixture.Run.RunDirectory `
                    -Decision $fixture.Decision -SchemaRoot $fixture.Run.SchemaRoot
            }) | Should Be $true
    }

    It 'accepts only exact promoted origin evidence and materializes direct resume provenance' {
        $origin = New-RSEvidenceFixture -Root (Join-Path $TestDrive 'resume-origin')
        $origin.Entry.cache_policy = 'ExactArtifact'
        $published = Publish-RSEvidenceFixturePass -Fixture $origin
        (Complete-RSValidationEvidenceRun -Run $origin.Run) | Should Be 'passed'
        $originPlan = Get-Content (Join-Path $origin.Run.RunDirectory 'validation-plan.json') -Raw | ConvertFrom-Json
        $originPlan.entries[0].cache_policy = 'ExactArtifact'
        $originDecision = Read-RSValidationDecisionRecord -RunDirectory $origin.Run.RunDirectory `
            -DecisionDigest $origin.Run.State.decision_digest -SchemaRoot $origin.Run.SchemaRoot
        $candidates = @(Get-RSValidationResumeCandidates -OriginRunDirectory $origin.Run.RunDirectory `
            -CurrentPlan $originPlan -CurrentDecision $originDecision -CurrentSnapshot $origin.Snapshot `
            -BuildReceipt $origin.Receipt -SchemaRoot (Join-Path $repoRoot 'Specs\Testing'))
        $candidates.Count | Should Be 1
        $candidates[0].reusable | Should Be $true
        $candidates[0].reason_code | Should Be 'REUSED_EXACT'

        $resumeDecision = New-RSValidationResumeDecision -FreshDecision $originDecision -Candidate $candidates
        $resumeDecision.entries[0].decision | Should Be 'REUSE'
        ($resumeDecision | ConvertTo-Json -Depth 30 | Test-Json -SchemaFile (Join-Path $repoRoot 'Specs\Testing\ValidationPlanDecision.schema.json')) | Should Be $true
        $resumeRun = New-RSValidationEvidenceRun -EvidenceRoot (Join-Path $TestDrive 'resume-current') `
            -RunId '20260813T130000Z-43-1123456789abcdef0123456789abcdef' -Plan $originPlan `
            -Decision $resumeDecision -BuildReceipt $origin.Receipt -SchemaRoot (Join-Path $repoRoot 'Specs\Testing') `
            -ValidationMode Resume -CampaignId $origin.Run.State.campaign_id -CampaignSequence 2 -ParentRunId $origin.Run.State.run_id
        $resumeRun.State.entries[0].status | Should Be 'reused'
        $resumeRun.State.entries[0].origin_evidence_digest | Should Be $published.Evidence.evidence_digest
        (Complete-RSValidationEvidenceRun -Run $resumeRun) | Should Be 'passed'
    }

    It 'builds a schema-valid aggregate v2 with explicit verdict axes and promotion accounting' {
        $fixture = New-RSEvidenceFixture -Root (Join-Path $TestDrive 'aggregate-v2')
        $published = Publish-RSEvidenceFixturePass -Fixture $fixture
        [void](Complete-RSValidationEvidenceRun -Run $fixture.Run)
        $plan = Get-Content (Join-Path $fixture.Run.RunDirectory 'validation-plan.json') -Raw | ConvertFrom-Json
        $decision = Read-RSValidationDecisionRecord -RunDirectory $fixture.Run.RunDirectory `
            -DecisionDigest $fixture.Run.State.decision_digest -SchemaRoot $fixture.Run.SchemaRoot
        $legacy = [pscustomobject]@{ exit_code = 0; passed = 3; failed = 0; skipped = 1; total = 4; duration_ms = 1250 }
        $summary = New-RSValidationRunSummaryV2 -Run $fixture.Run -Plan $plan -Decision $decision `
            -LegacySummary $legacy -EvidenceRoot $fixture.Run.EvidenceRoot `
            -SchemaRoot (Join-Path $repoRoot 'Specs\Testing') -CoverageScope complete
        $summary.schema | Should Be 'red-salamander.run-all-tests.v2'
        $summary.repository_verdict | Should Be 'PASSED'
        $summary.change_verdict | Should Be 'PASSED'
        $summary.counts.executed | Should Be 1
        $summary.counts.promoted | Should Be 1
        ($summary | ConvertTo-Json -Depth 40 | Test-Json -SchemaFile (Join-Path $repoRoot 'Specs\Testing\RunAllTestsSummary.schema.json')) | Should Be $true
        (Convert-RSTestRunSummaryToGitHubStepSummary -Summary $summary) | Should Match 'Repository verdict: \*\*PASSED\*\*'

        $filtered = New-RSValidationRunSummaryV2 -Run $fixture.Run -Plan $plan -Decision $decision `
            -LegacySummary $legacy -EvidenceRoot $fixture.Run.EvidenceRoot `
            -SchemaRoot (Join-Path $repoRoot 'Specs\Testing') -CoverageScope filtered
        $filtered.repository_verdict | Should Be 'NOT_EVALUATED'
        $filtered.change_verdict | Should Be 'PASSED'

        $fixture.Run.State.entries[0].promotion_digest = $null
        $nullPromotion = New-RSValidationRunSummaryV2 -Run $fixture.Run -Plan $plan -Decision $decision `
            -LegacySummary $legacy -EvidenceRoot $fixture.Run.EvidenceRoot `
            -SchemaRoot (Join-Path $repoRoot 'Specs\Testing') -CoverageScope complete
        $nullPromotion.repository_verdict | Should Be 'BLOCKED'

        $fixture.Run.State.entries[0].promotion_digest = $published.Promotion.promotion_digest
        $wrongModePlan = $plan.PSObject.Copy()
        $wrongModePlan.validation_mode = 'Resume'
        $wrongMode = New-RSValidationRunSummaryV2 -Run $fixture.Run -Plan $wrongModePlan -Decision $decision `
            -LegacySummary $legacy -EvidenceRoot $fixture.Run.EvidenceRoot `
            -SchemaRoot (Join-Path $repoRoot 'Specs\Testing') -CoverageScope complete
        $wrongMode.repository_verdict | Should Be 'BLOCKED'
    }

    It 'blocks aggregation when the bound terminal-result record is corrupt' {
        $fixture = New-RSEvidenceFixture -Root (Join-Path $TestDrive 'aggregate-terminal-corrupt')
        $published = Publish-RSEvidenceFixturePass -Fixture $fixture
        [void](Complete-RSValidationEvidenceRun -Run $fixture.Run)
        [IO.File]::WriteAllText($published.TerminalResultPath, '{', [Text.UTF8Encoding]::new($false))
        $plan = Get-Content (Join-Path $fixture.Run.RunDirectory 'validation-plan.json') -Raw | ConvertFrom-Json
        $decision = Read-RSValidationDecisionRecord -RunDirectory $fixture.Run.RunDirectory `
            -DecisionDigest $fixture.Run.State.decision_digest -SchemaRoot $fixture.Run.SchemaRoot
        $legacy = [pscustomobject]@{ exit_code = 0; passed = 1; failed = 0; skipped = 0; total = 1; duration_ms = 1 }
        $summary = New-RSValidationRunSummaryV2 -Run $fixture.Run -Plan $plan -Decision $decision `
            -LegacySummary $legacy -EvidenceRoot $fixture.Run.EvidenceRoot `
            -SchemaRoot (Join-Path $repoRoot 'Specs\Testing') -CoverageScope complete
        $summary.repository_verdict | Should Be 'BLOCKED'
    }

    It 'uses bounded promoted-history medians only as advisory duration estimates' {
        $root = Join-Path $TestDrive 'duration-history'
        $durations = @(1, 3, 5)
        for ($index = 0; $index -lt $durations.Count; ++$index) {
            $runId = "20260813T14000$($index + 1)Z-50-2123456789abcdef0123456789abcde$($index + 1)"
            $fixture = New-RSEvidenceFixture -Root $root -RunId $runId
            $start = [datetime]'2026-08-13T14:00:00Z'
            $result = New-RSEvidencePassingResult
            [void](Publish-RSValidationEntryEvidence -Run $fixture.Run -EntryContract $fixture.Entry `
                -EntryFingerprint $fixture.Fingerprint `
                -TerminalResult (New-RSEvidenceTerminalResult -Fixture $fixture -Result $result) `
                -PreSnapshot $fixture.Snapshot -PostSnapshot $fixture.Snapshot -BuildReceipt $fixture.Receipt `
                -StartedUtc $start -EndedUtc $start.AddSeconds($durations[$index]))
            [void](Complete-RSValidationEvidenceRun -Run $fixture.Run)
        }

        $corrupt = Join-Path (Join-Path $root 'runs') '20260813T150000Z-51-3123456789abcdef0123456789abcdef'
        [void](New-Item -ItemType Directory -Path $corrupt -Force)
        [IO.File]::WriteAllText((Join-Path $corrupt 'run-state.json'), '{', [Text.UTF8Encoding]::new($false))

        $estimates = Get-RSValidationDurationEstimates -EvidenceRoot $root `
            -SchemaRoot (Join-Path $repoRoot 'Specs\Testing') -EntryId @('tooling.pester', 'unknown.entry')
        $estimates['tooling.pester'] | Should Be 3000
        $estimates['unknown.entry'] | Should Be 0
    }

    It 'executes instead of reusing on fingerprint, receipt, non-cacheable, or corrupt evidence mismatch' {
        $origin = New-RSEvidenceFixture -Root (Join-Path $TestDrive 'resume-reject')
        $origin.Entry.cache_policy = 'ExactArtifact'
        $published = Publish-RSEvidenceFixturePass -Fixture $origin
        [void](Complete-RSValidationEvidenceRun -Run $origin.Run)
        $plan = Get-Content (Join-Path $origin.Run.RunDirectory 'validation-plan.json') -Raw | ConvertFrom-Json
        $plan.entries[0].cache_policy = 'ExactArtifact'
        $decision = Read-RSValidationDecisionRecord -RunDirectory $origin.Run.RunDirectory `
            -DecisionDigest $origin.Run.State.decision_digest -SchemaRoot $origin.Run.SchemaRoot

        $changedDecision = $decision.PSObject.Copy()
        $changedDecision.entries = @($decision.entries | ForEach-Object { $copy = $_.PSObject.Copy(); $copy.expected_fingerprint = '1' * 64; $copy })
        (Get-RSValidationResumeCandidates -OriginRunDirectory $origin.Run.RunDirectory -CurrentPlan $plan `
            -CurrentDecision $changedDecision -CurrentSnapshot $origin.Snapshot -BuildReceipt $origin.Receipt `
            -SchemaRoot (Join-Path $repoRoot 'Specs\Testing')).reason_code | Should Be 'SOURCE_INPUT_CHANGED'
        $changedReceipt = $origin.Receipt.PSObject.Copy(); $changedReceipt.receipt_id = '2' * 64
        (Get-RSValidationResumeCandidates -OriginRunDirectory $origin.Run.RunDirectory -CurrentPlan $plan `
            -CurrentDecision $decision -CurrentSnapshot $origin.Snapshot -BuildReceipt $changedReceipt `
            -SchemaRoot (Join-Path $repoRoot 'Specs\Testing')).reason_code | Should Be 'BUILD_ARTIFACT_CHANGED'
        $plan.entries[0].cache_policy = 'Never'
        (Get-RSValidationResumeCandidates -OriginRunDirectory $origin.Run.RunDirectory -CurrentPlan $plan `
            -CurrentDecision $decision -CurrentSnapshot $origin.Snapshot -BuildReceipt $origin.Receipt `
            -SchemaRoot (Join-Path $repoRoot 'Specs\Testing')).reason_code | Should Be 'NON_CACHEABLE'
        $plan.entries[0].cache_policy = 'ExactArtifact'
        [IO.File]::WriteAllText($published.PromotionPath, '{', [Text.UTF8Encoding]::new($false))
        (Get-RSValidationResumeCandidates -OriginRunDirectory $origin.Run.RunDirectory -CurrentPlan $plan `
            -CurrentDecision $decision -CurrentSnapshot $origin.Snapshot -BuildReceipt $origin.Receipt `
            -SchemaRoot (Join-Path $repoRoot 'Specs\Testing')).reason_code | Should Be 'CORRUPT_EVIDENCE'
    }

    It 'refuses a nested prune target reached through an ancestor junction without changing external data' {
        $fixture = New-RSEvidenceFixture -Root (Join-Path $TestDrive 'prune')
        $outside = Join-Path $TestDrive 'outside'
        $sentinel = Join-Path $outside 'sentinel.txt'
        $runsRoot = Join-Path $fixture.Run.EvidenceRoot 'runs'
        $link = Join-Path $runsRoot 'link'
        $runId = '20260813T130000Z-42-1123456789abcdef0123456789abcdef'
        [void](New-Item -ItemType Directory -Path $outside)
        [IO.File]::WriteAllText($sentinel, 'preserve-prune-target')
        [void](New-Item -ItemType Directory -Path (Join-Path $outside $runId))
        [void](New-Item -ItemType Junction -Path $link -Target $outside)
        try {
            $craftedPlan = @([pscustomobject]@{
                    LiteralPath = Join-Path $link $runId
                    RunId = $runId
                    DirectoryIdentity = 'attacker-controlled'
                })
            $message = Get-RSEvidenceActionErrorMessage {
                Remove-RSValidationEvidenceRuns -EvidenceRoot $fixture.Run.EvidenceRoot `
                    -PrunePlan $craftedPlan -Confirm:$false
            }
            $message | Should Match 'direct child'
            (Get-Content -LiteralPath $sentinel -Raw) | Should Be 'preserve-prune-target'
        }
        finally {
            Remove-Item -LiteralPath $link -Force
        }
    }

    It 'selects stable component-specific invalidation reasons before aggregate fallback' {
        $origin = New-RSEvidenceFixture -Root (Join-Path $TestDrive 'resume-components')
        $origin.Entry.cache_policy = 'ExactArtifact'
        $originComponents = New-RSEvidenceFingerprintComponents
        [void](Publish-RSEvidenceFixturePass -Fixture $origin -FingerprintComponents $originComponents)
        [void](Complete-RSValidationEvidenceRun -Run $origin.Run)
        $plan = Get-Content (Join-Path $origin.Run.RunDirectory 'validation-plan.json') -Raw | ConvertFrom-Json
        $plan.entries[0].cache_policy = 'ExactArtifact'
        $decision = Read-RSValidationDecisionRecord -RunDirectory $origin.Run.RunDirectory `
            -DecisionDigest $origin.Run.State.decision_digest -SchemaRoot $origin.Run.SchemaRoot
        $changedDecision = $decision.PSObject.Copy()
        $changedDecision.entries = @($decision.entries | ForEach-Object {
                $copy = $_.PSObject.Copy(); $copy.expected_fingerprint = 'f' * 64; $copy
            })

        $reasonByComponent = [ordered]@{
            build_artifact = 'BUILD_ARTIFACT_CHANGED'; logical_input = 'LOGICAL_INPUT_CHANGED'
            shared_contract = 'SHARED_CONTRACT_CHANGED'; runner = 'RUNNER_CHANGED'
            environment = 'ENVIRONMENT_CHANGED'; arguments = 'ARGUMENTS_CHANGED'
            coverage = 'COVERAGE_CHANGED'; outcome_policy = 'OUTCOME_POLICY_CHANGED'
            capability = 'CAPABILITY_CHANGED'; artifact_epoch = 'ARTIFACT_EPOCH_CHANGED'
            source_input = 'SOURCE_INPUT_CHANGED'
        }
        foreach ($componentName in $reasonByComponent.Keys) {
            $changedComponents = $originComponents.PSObject.Copy()
            $changedComponents.$componentName = '0' * 64
            $componentMap = @{ 'tooling.pester' = [pscustomobject]@{
                    version = 'validation-entry-components.v1'; digests = $changedComponents
                } }
            (Get-RSValidationResumeCandidates -OriginRunDirectory $origin.Run.RunDirectory -CurrentPlan $plan `
                -CurrentDecision $changedDecision -CurrentSnapshot $origin.Snapshot -BuildReceipt $origin.Receipt `
                -CurrentFingerprintComponentsById $componentMap `
                -SchemaRoot (Join-Path $repoRoot 'Specs\Testing')).reason_code | Should Be $reasonByComponent[$componentName]
        }

        $simultaneous = $originComponents.PSObject.Copy()
        $simultaneous.source_input = '0' * 64
        $simultaneous.build_artifact = '0' * 64
        $componentMap = @{ 'tooling.pester' = [pscustomobject]@{
                version = 'validation-entry-components.v1'; digests = $simultaneous
            } }
        (Get-RSValidationResumeCandidates -OriginRunDirectory $origin.Run.RunDirectory -CurrentPlan $plan `
            -CurrentDecision $changedDecision -CurrentSnapshot $origin.Snapshot -BuildReceipt $origin.Receipt `
            -CurrentFingerprintComponentsById $componentMap `
            -SchemaRoot (Join-Path $repoRoot 'Specs\Testing')).reason_code | Should Be 'BUILD_ARTIFACT_CHANGED'

        $future = $originComponents.PSObject.Copy()
        $future | Add-Member -NotePropertyName future_component -NotePropertyValue ('0' * 64)
        $componentMap = @{ 'tooling.pester' = [pscustomobject]@{
                version = 'validation-entry-components.v1'; digests = $future
            } }
        (Get-RSValidationResumeCandidates -OriginRunDirectory $origin.Run.RunDirectory -CurrentPlan $plan `
            -CurrentDecision $changedDecision -CurrentSnapshot $origin.Snapshot -BuildReceipt $origin.Receipt `
            -CurrentFingerprintComponentsById $componentMap `
            -SchemaRoot (Join-Path $repoRoot 'Specs\Testing')).reason_code | Should Be 'SCHEMA_MISMATCH'
    }

    It 'refuses prune apply when a planned run directory identity has changed' {
        $root = Join-Path $TestDrive 'prune-identity'
        $older = New-RSEvidenceFixture -Root $root `
            -RunId '20260813T120000Z-42-0123456789abcdef0123456789abcdef'
        [void](Publish-RSEvidenceFixturePass -Fixture $older)
        [void](Complete-RSValidationEvidenceRun -Run $older.Run)
        $newer = New-RSEvidenceFixture -Root $root `
            -RunId '20260813T130000Z-43-1123456789abcdef0123456789abcdef'
        [void](Publish-RSEvidenceFixturePass -Fixture $newer)
        [void](Complete-RSValidationEvidenceRun -Run $newer.Run)

        $plan = @(Get-RSValidationEvidencePrunePlan -EvidenceRoot $root -MaximumRuns 1)
        $plan.Count | Should Be 1
        $plan[0].RunId | Should Be $older.Run.State.run_id
        $plannedPath = $plan[0].LiteralPath
        Remove-Item -LiteralPath $plannedPath -Recurse -Force
        [void](New-Item -ItemType Directory -Path $plannedPath)
        $sentinel = Join-Path $plannedPath 'replacement-sentinel.txt'
        [IO.File]::WriteAllText($sentinel, 'replacement-must-survive')

        $message = Get-RSEvidenceActionErrorMessage {
            Remove-RSValidationEvidenceRuns -EvidenceRoot $root -PrunePlan $plan -Confirm:$false
        }
        $message | Should Match 'identity changed'
        (Get-Content -LiteralPath $sentinel -Raw) | Should Be 'replacement-must-survive'
    }
}
