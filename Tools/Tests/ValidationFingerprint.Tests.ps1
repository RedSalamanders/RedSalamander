$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $PSScriptRoot 'TestSupport.psm1') -Force
Import-Module (Join-Path $repoRoot 'Tools\Modules\Testing\ValidationFingerprint.psm1') -Force

function Test-RSActionThrows {
    param([Parameter(Mandatory = $true)][scriptblock]$Action)

    $threw = $false
    try { & $Action | Out-Null } catch { $threw = $true }
    return $threw
}

$schemaPaths = @(
    'Specs\Build\BuildReceipt.schema.json',
    'Specs\Build\PortableBuildAttestation.schema.json',
    'Specs\Testing\WorkspaceSnapshot.schema.json',
    'Specs\Testing\ValidationPlan.schema.json',
    'Specs\Testing\ValidationEntryFingerprints.schema.json',
    'Specs\Testing\ValidationPlanDecision.schema.json',
    'Specs\Testing\ValidationRunState.schema.json',
    'Specs\Testing\ValidationTerminalResult.schema.json',
    'Specs\Testing\ValidationEntryEvidence.schema.json',
    'Specs\Testing\ValidationPromotion.schema.json',
    'Specs\Testing\ValidationImpact.schema.json',
    'Specs\Testing\RunAllTestsSummary.schema.json'
)

Describe 'Operation Startrail schema foundation' {
    It 'uses one reviewed schema draft and closes every root object' {
        foreach ($relativePath in $schemaPaths) {
            $schema = Get-Content -LiteralPath (Join-Path $repoRoot $relativePath) -Raw | ConvertFrom-Json
            $schema.'$schema' | Should Be 'http://json-schema.org/draft-07/schema#'
            $schema.additionalProperties | Should Be $false
        }
    }

    It 'rejects missing canonicalization, unknown states, invalid paths, and cache policies' {
        $digest = 'a' * 64
        $snapshotSchema = Join-Path $repoRoot 'Specs\Testing\WorkspaceSnapshot.schema.json'
        $validSnapshot = [ordered]@{
            schema = 'red-salamander.workspace-snapshot.v1'
            canonicalization = 'rs-canonical-json-v1'
            snapshot_id = $digest
            repo_identity = $digest
            git_head = 'b' * 40
            base = $null
            files = @([ordered]@{
                    path = 'src/file.cpp'
                    status = 'modified'
                    base_sha256 = $digest
                    base_mode = '100644'
                    index_sha256 = $digest
                    index_mode = '100644'
                    worktree_sha256 = $digest
                    worktree_size_bytes = 10
                    old_path = $null
                })
            acquisition_before = $digest
            acquisition_after = $digest
            acquisition_attempts = 1
            git_discovery_ms = 2
            content_hashing_ms = 2
            canonicalization_ms = 1
            computation_ms = 5
            created_utc = '2026-08-13T12:00:00Z'
        }
        $validJson = $validSnapshot | ConvertTo-Json -Depth 20 -Compress
        (Test-Json -Json $validJson -SchemaFile $snapshotSchema) | Should Be $true

        $missingCanonicalization = [ordered]@{}
        foreach ($key in $validSnapshot.Keys) {
            if ($key -ne 'canonicalization') { $missingCanonicalization[$key] = $validSnapshot[$key] }
        }
        (Test-Json -Json ($missingCanonicalization | ConvertTo-Json -Depth 20 -Compress) -SchemaFile $snapshotSchema -ErrorAction SilentlyContinue) | Should Be $false

        $validSnapshot.files[0].status = 'mystery'
        (Test-Json -Json ($validSnapshot | ConvertTo-Json -Depth 20 -Compress) -SchemaFile $snapshotSchema -ErrorAction SilentlyContinue) | Should Be $false
        $validSnapshot.files[0].status = 'modified'
        $validSnapshot.files[0].path = '../escape.cpp'
        (Test-Json -Json ($validSnapshot | ConvertTo-Json -Depth 20 -Compress) -SchemaFile $snapshotSchema -ErrorAction SilentlyContinue) | Should Be $false

        $planSchema = Get-Content -LiteralPath (Join-Path $repoRoot 'Specs\Testing\ValidationPlan.schema.json') -Raw
        $planSchema | Should Match '"cache_policy"'
        $planSchema | Should Match '"LogicalSameCampaign"'
        $planSchema | Should Not Match '"TrustWhateverExists"'
    }
}

Describe 'Operation Startrail canonical JSON' {
    It 'normalizes strings, sorts object keys ordinally, and preserves array order' {
        $decomposed = "e$([char]0x0301)"
        $record = [ordered]@{ z = @($true, $null, 2); a = $decomposed; quote = '"\' }
        $json = ConvertTo-RSCanonicalJson -InputObject $record
        $json | Should Be '{"a":"é","quote":"\"\\","z":[true,null,2]}'
    }

    It 'produces one stable lowercase SHA-256 identity' {
        $first = Get-RSCanonicalJsonDigest -InputObject ([ordered]@{ b = 2; a = 1 })
        $second = Get-RSCanonicalJsonDigest -InputObject ([ordered]@{ a = 1; b = 2 })
        $first | Should Be $second
        $first | Should Match '^[0-9a-f]{64}$'
    }

    It 'emits byte-identical canonical fixtures in a clean supported pwsh process' {
        $record = [ordered]@{ z = @($true, $null, 2); a = "e$([char]0x0301)" }
        $expected = [Convert]::ToBase64String((ConvertTo-RSCanonicalJsonBytes -InputObject $record))
        $modulePath = Join-Path $repoRoot 'Tools\Modules\Testing\ValidationFingerprint.psm1'
        $scriptPath = Join-Path $TestDrive 'canonical-child.ps1'
        $scriptText = @'
param([string]$ModulePath)
Import-Module $ModulePath -Force
$record = [ordered]@{ z = @($true, $null, 2); a = "e$([char]0x0301)" }
[Convert]::ToBase64String((ConvertTo-RSCanonicalJsonBytes -InputObject $record))
'@
        [IO.File]::WriteAllText($scriptPath, $scriptText, [Text.UTF8Encoding]::new($false))
        $actual = & pwsh.exe -NoProfile -File $scriptPath -ModulePath $modulePath
        $LASTEXITCODE | Should Be 0
        $actual | Should Be $expected
    }

    It 'rejects floating point and date coercion' {
        (Test-RSActionThrows { ConvertTo-RSCanonicalJson -InputObject 1.5 }) | Should Be $true
        (Test-RSActionThrows { ConvertTo-RSCanonicalJson -InputObject (Get-Date) }) | Should Be $true
    }

    It 'rejects invalid identifiers, paths, aliases, duplicate plan ids, and provenance cycles' {
        (Test-RSActionThrows { Assert-RSValidationIdentifier -Id 'Upper.Case' }) | Should Be $true
        (Test-RSValidationRelativePath -Path 'src/ok.cpp') | Should Be $true
        (Test-RSValidationRelativePath -Path '../escape.cpp') | Should Be $false
        (Test-RSValidationRelativePath -Path 'C:/absolute.cpp') | Should Be $false
        (Test-RSActionThrows { Assert-RSValidationPathSet -Path @('src/File.cpp', 'src/file.cpp') }) | Should Be $true
        (Test-RSActionThrows { Assert-RSValidationPlanSemantics -Plan ([pscustomobject]@{ entries = @([pscustomobject]@{ id = 'entry.one' }, [pscustomobject]@{ id = 'entry.one' }) }) }) | Should Be $true
        (Test-RSActionThrows { Assert-RSValidationProvenance -Record @(
                [pscustomobject]@{ run_id = 'one'; origin_run_id = 'two' },
                [pscustomobject]@{ run_id = 'two'; origin_run_id = 'one' }) }) | Should Be $true
    }

    It 'freezes the complete reason-code enum' {
        $codes = @(Get-RSValidationReasonCodes)
        $codes.Count | Should Be 26
        ($codes -contains 'ORIGIN_NOT_PROMOTED') | Should Be $true
        ($codes -contains 'MONOLITHIC_LOGICAL_CANDIDATE_NOT_FULL') | Should Be $true
    }

    It 'rejects Startrail evidence on Windows PowerShell 5.1 before evidence IO' {
        $modulePath = Join-Path $repoRoot 'Tools\Modules\Testing\ValidationFingerprint.psm1'
        $output = & powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Import-Module '$modulePath' -Force; Assert-RSStartrailPowerShellHost" 2>&1
        $LASTEXITCODE | Should Not Be 0
        ($output -join "`n") | Should Match 'requires PowerShell 7\.4 or newer'
    }
}

Describe 'Operation Startrail dependency closure' {
    It 'derives recursive script, module, and schema inputs without unrelated extras' {
        $root = Join-Path $TestDrive 'dependency-closure'
        [void](New-Item -ItemType Directory -Path (Join-Path $root 'Tools\Modules') -Force)
        [IO.File]::WriteAllText((Join-Path $root 'Tools\Root.ps1'),
            "Import-Module (Join-Path `$PSScriptRoot 'Modules\A.psm1')`n",
            [Text.UTF8Encoding]::new($false))
        [IO.File]::WriteAllText((Join-Path $root 'Tools\Modules\A.psm1'),
            "`$schema = Join-Path `$PSScriptRoot 'schema.json'`n",
            [Text.UTF8Encoding]::new($false))
        [IO.File]::WriteAllText((Join-Path $root 'Tools\Modules\B.psm1'), '', [Text.UTF8Encoding]::new($false))
        [IO.File]::WriteAllText((Join-Path $root 'Tools\Modules\schema.json'), '{}', [Text.UTF8Encoding]::new($false))
        & git -C $root init --quiet
        $closure = @(Get-RSValidationDependencyClosure -RepoRoot $root -RootRelativePath 'Tools/Root.ps1')
        $closure | Should Be @('Tools/Modules/A.psm1', 'Tools/Modules/schema.json', 'Tools/Root.ps1')

        [IO.File]::AppendAllText((Join-Path $root 'Tools\Modules\A.psm1'),
            "Import-Module (Join-Path `$PSScriptRoot 'B.psm1')`n",
            [Text.UTF8Encoding]::new($false))
        @(Get-RSValidationDependencyClosure -RepoRoot $root -RootRelativePath 'Tools/Root.ps1') | Should Be @(
            'Tools/Modules/A.psm1', 'Tools/Modules/B.psm1', 'Tools/Modules/schema.json', 'Tools/Root.ps1')
    }

    It 'binds the current runner summary module and aggregate schema transitively' {
        $closure = @(Get-RSValidationDependencyClosure -RepoRoot $repoRoot -RootRelativePath 'Tools/Run-AllTests.ps1')
        ($closure -contains 'Tools/Modules/Testing/TestRunSummary.psm1') | Should Be $true
        ($closure -contains 'Specs/Testing/RunAllTestsSummary.schema.json') | Should Be $true
        $runnerSource = Get-Content (Join-Path $repoRoot 'Tools\Run-AllTests.ps1') -Raw
        $runnerSource | Should Match 'Get-RSValidationDependencyClosure'
        $runnerSource | Should Not Match '\$runnerDependencyPaths\s*=\s*@\(\s*[''"]'
    }
}

Describe 'Operation Startrail workspace snapshots' {
    It 'keeps the CI managed tool checkout outside product source identity' {
        $fixture = New-RSValidationGitFixture -Root (Join-Path $TestDrive 'managed build checkout')
        [IO.File]::WriteAllText((Join-Path $fixture.Root '.gitignore'), ".build/`n", [Text.UTF8Encoding]::new($false))
        $before = Get-RSWorkspaceSnapshot -RepoRoot $fixture.Root
        $toolRoot = Join-Path $fixture.Root '.build\vcpkg-tool'
        [void](New-Item -ItemType Directory -Path $toolRoot -Force)
        & git -C $toolRoot init --quiet
        $LASTEXITCODE | Should Be 0
        [IO.File]::WriteAllText((Join-Path $toolRoot 'bootstrap-vcpkg.bat'), 'managed tool', [Text.UTF8Encoding]::new($false))
        (Get-RSWorkspaceSnapshot -RepoRoot $fixture.Root).snapshot_id | Should Be $before.snapshot_id

        # A checkout at the old CI location must still fail closed, not become an
        # unreviewed exclusion from the source-integrity boundary.
        $unmanagedRoot = Join-Path $fixture.Root 'vcpkg'
        [void](New-Item -ItemType Directory -Path $unmanagedRoot -Force)
        & git -C $unmanagedRoot init --quiet
        $LASTEXITCODE | Should Be 0
        [IO.File]::WriteAllText((Join-Path $unmanagedRoot 'unexpected.txt'), 'unmanaged tool', [Text.UTF8Encoding]::new($false))
        (Test-RSActionThrows { Get-RSWorkspaceSnapshot -RepoRoot $fixture.Root }) | Should Be $true
    }

    It 'preserves distinct base index and worktree identities for every Git mutation class' {
        $fixture = New-RSValidationGitFixture -Root (Join-Path $TestDrive 'workspace identities')
        $dirtyPath = Join-Path $fixture.Root $fixture.DirtyPath
        [IO.File]::WriteAllText($dirtyPath, "index-content`n", [Text.UTF8Encoding]::new($false))
        & git -C $fixture.Root add -- $fixture.DirtyPath
        [IO.File]::WriteAllText($dirtyPath, "worktree-content`n", [Text.UTF8Encoding]::new($false))

        $first = Get-RSWorkspaceSnapshot -RepoRoot $fixture.Root
        $second = Get-RSWorkspaceSnapshot -RepoRoot $fixture.Root
        $first.snapshot_id | Should Be $second.snapshot_id
        $first.acquisition_before | Should Be $first.acquisition_after
        $first.computation_ms | Should BeGreaterThan -1
        $first.git_discovery_ms | Should BeGreaterThan -1
        $first.content_hashing_ms | Should BeGreaterThan -1
        $first.canonicalization_ms | Should BeGreaterThan -1
        ($first | ConvertTo-Json -Depth 20 | Test-Json -SchemaFile (Join-Path $repoRoot 'Specs\Testing\WorkspaceSnapshot.schema.json')) | Should Be $true

        $dirty = $first.files | Where-Object path -eq $fixture.DirtyPath | Select-Object -First 1
        $dirty.status | Should Be 'modified'
        $dirty.base_sha256 | Should Not Be $dirty.index_sha256
        $dirty.index_sha256 | Should Not Be $dirty.worktree_sha256
        $dirty.worktree_sha256 | Should Be ((Get-FileHash -LiteralPath $dirtyPath -Algorithm SHA256).Hash.ToLowerInvariant())
        ($first.files | Where-Object path -eq $fixture.DeletedPath).status | Should Be 'deleted'
        $rename = $first.files | Where-Object path -eq $fixture.RenamedNewPath | Select-Object -First 1
        $rename.status | Should Be 'renamed'
        $rename.old_path | Should Be $fixture.RenamedOldPath
        ($first.files | Where-Object path -eq $fixture.StagedPath).status | Should Be 'added'
        ($first.files | Where-Object path -eq $fixture.UntrackedPath).status | Should Be 'untracked'

        (Get-Item -LiteralPath $dirtyPath).LastWriteTimeUtc = [datetime]::UtcNow.AddHours(-2)
        (Get-RSWorkspaceSnapshot -RepoRoot $fixture.Root).snapshot_id | Should Be $first.snapshot_id
        (Get-RSWorkspaceSnapshot -RepoRoot $fixture.Root.ToUpperInvariant()).snapshot_id | Should Be $first.snapshot_id
        $excludedArchive = Join-Path $fixture.Root 'Specs\TestRuns\ignored.json'
        [void](New-Item -ItemType Directory -Path (Split-Path -Parent $excludedArchive) -Force)
        [IO.File]::WriteAllText($excludedArchive, '{}', [Text.UTF8Encoding]::new($false))
        [IO.File]::WriteAllText((Join-Path $fixture.Root 'ignored.log'), 'ignored', [Text.UTF8Encoding]::new($false))
        (Get-RSWorkspaceSnapshot -RepoRoot $fixture.Root).snapshot_id | Should Be $first.snapshot_id

        [IO.File]::WriteAllText($dirtyPath, "worktree-content`r`n", [Text.UTF8Encoding]::new($false))
        (Get-RSWorkspaceSnapshot -RepoRoot $fixture.Root).snapshot_id | Should Not Be $first.snapshot_id
    }

    It 'changes identity for committed and pre/post worktree mutations' {
        $fixture = New-RSValidationGitFixture -Root (Join-Path $TestDrive 'workspace committed')
        & git -C $fixture.Root reset --hard --quiet HEAD
        & git -C $fixture.Root clean -fdq
        $clean = Get-RSWorkspaceSnapshot -RepoRoot $fixture.Root
        [IO.File]::WriteAllText((Join-Path $fixture.Root 'dirty.txt'), "post`n", [Text.UTF8Encoding]::new($false))
        $post = Get-RSWorkspaceSnapshot -RepoRoot $fixture.Root
        $post.snapshot_id | Should Not Be $clean.snapshot_id
        & git -C $fixture.Root add -- dirty.txt
        & git -C $fixture.Root commit --quiet -m 'committed mutation'
        $committed = Get-RSWorkspaceSnapshot -RepoRoot $fixture.Root
        $committed.snapshot_id | Should Not Be $clean.snapshot_id
        @($committed.files).Count | Should Be 0
    }

    It 'rejects reparse inputs and unsafe impact-base arguments without fetching' {
        $fixture = New-RSValidationGitFixture -Root (Join-Path $TestDrive 'workspace reparse')
        & git -C $fixture.Root reset --hard --quiet HEAD
        & git -C $fixture.Root clean -fdq
        $link = Join-Path $fixture.Root 'reparse.txt'
        [void](New-Item -ItemType SymbolicLink -Path $link -Target (Join-Path $fixture.Root 'dirty.txt'))
        (Test-RSActionThrows { Get-RSWorkspaceSnapshot -RepoRoot $fixture.Root }) | Should Be $true
        Remove-Item -LiteralPath $link
        (Test-RSActionThrows { Get-RSWorkspaceSnapshot -RepoRoot $fixture.Root -ImpactBase '--help' }) | Should Be $true
    }

    It 'uses the local merge base and includes committed divergence without fetching' {
        $fixture = New-RSValidationGitFixture -Root (Join-Path $TestDrive 'workspace merge base')
        & git -C $fixture.Root reset --hard --quiet HEAD
        & git -C $fixture.Root clean -fdq
        $base = (& git -C $fixture.Root rev-parse HEAD).Trim()
        & git -C $fixture.Root branch impact-base
        [IO.File]::WriteAllText((Join-Path $fixture.Root 'dirty.txt'), "branch-change`n", [Text.UTF8Encoding]::new($false))
        & git -C $fixture.Root add -- dirty.txt
        & git -C $fixture.Root commit --quiet -m 'branch divergence'
        $snapshot = Get-RSWorkspaceSnapshot -RepoRoot $fixture.Root -ImpactBase impact-base
        $snapshot.base | Should Be $base
        ($snapshot.files | Where-Object path -eq 'dirty.txt').status | Should Be 'modified'
    }

    It 'preserves the included side of staged and unstaged boundary-crossing renames' {
        $root = Join-Path $TestDrive 'workspace rename boundaries'
        [void](New-Item -ItemType Directory -Path $root)
        [void](New-Item -ItemType Directory -Path (Join-Path $root '.build'))
        foreach ($path in @(
                'staged-to-log.h', 'unstaged-to-log.h',
                'staged-to-build.h', 'unstaged-to-build.h',
                'excluded-source.log', '.build/excluded-source.h')) {
            $fullPath = Join-Path $root $path
            [void](New-Item -ItemType Directory -Path (Split-Path -Parent $fullPath) -Force)
            [IO.File]::WriteAllText($fullPath, "$path`n", [Text.UTF8Encoding]::new($false))
        }
        & git -C $root init --quiet
        & git -C $root config user.email 'startrail-boundary@invalid.example'
        & git -C $root config user.name 'Startrail Boundary Fixture'
        & git -C $root add -f -- .
        & git -C $root commit --quiet -m 'boundary base'

        & git -C $root mv -- staged-to-log.h staged-to-log.log
        & git -C $root mv -- staged-to-build.h .build/staged-to-build.h
        & git -C $root mv -- excluded-source.log included-from-log.h
        & git -C $root mv -- .build/excluded-source.h included-from-build.h
        Move-Item -LiteralPath (Join-Path $root 'unstaged-to-log.h') -Destination (Join-Path $root 'unstaged-to-log.log')
        Move-Item -LiteralPath (Join-Path $root 'unstaged-to-build.h') -Destination (Join-Path $root '.build\unstaged-to-build.h')

        $snapshot = Get-RSWorkspaceSnapshot -RepoRoot $root
        foreach ($deleted in @('staged-to-log.h', 'unstaged-to-log.h', 'staged-to-build.h', 'unstaged-to-build.h')) {
            $row = $snapshot.files | Where-Object path -eq $deleted | Select-Object -First 1
            $row.status | Should Be 'deleted'
            $row.old_path | Should Be $null
        }
        foreach ($added in @('included-from-log.h', 'included-from-build.h')) {
            $row = $snapshot.files | Where-Object path -eq $added | Select-Object -First 1
            $row.status | Should Be 'added'
            $row.old_path | Should Be $null
        }
        @($snapshot.files | Where-Object path -match '(?:\.log$|^\.build/)').Count | Should Be 0
    }
}

Describe 'Operation Startrail entry fingerprints' {
    It 'binds the complete entry, plan, workspace, receipt, epoch, runner, environment, and quarantine inputs' {
        $entry = [pscustomobject][ordered]@{
            id = 'commands.selftest'; display_name = 'Commands'; kind = 'SelfTest'; path = '.build/x64/Debug/RedSalamander.exe'
            arguments = @('--commands-selftest'); working_directory = '.'; result_contract_id = 'result.commands'
            impact_tags = @('test.commands'); input_sets = @('source.commands'); artifact_ids = @('artifact.profile')
            artifact_read_ids = @('artifact.profile'); artifact_write_ids = @(); produces_artifact_ids = @()
            invalidates_receipt_scopes = @(); cache_policy = 'LogicalSameCampaign'; environment_inputs = @('env.interactive-desktop')
            coverage_contract = [pscustomobject][ordered]@{ suite = 'commands.selftest'; family = 'Commands'; exact_cases = @(); repeat = 1; order_identity = 'commands.selftest.default' }
            outcome_contract = [pscustomobject][ordered]@{ pass_required_cases = @('selected'); static_skip_cases = @(); capability_bound_skip_cases = @('declared-capability'); forbidden_skip_classes = @('unclassified') }
            order_contract = 'broad-same-process'; resource_classes = @('artifact-read-only', 'interactive-desktop')
            estimated_duration_key = 'duration.commands.selftest'; requires_interactive_desktop = $true
        }
        $coveragePlanDigest = 'a' * 64
        $workspace = [pscustomobject]@{ snapshot_id = 'b' * 64 }
        $receipt = [pscustomobject][ordered]@{
            receipt_id = 'c' * 64; source_snapshot_id = 'b' * 64; platform = 'x64'; configuration = 'Debug'
            target = 'Build'; tests_enabled = $true; toolchain_identity_digest = 'd' * 64; vcpkg_identity_digest = 'e' * 64
            ghostty_runtime_identity = 'f' * 64; project_graph_digest = '1' * 64; runtime_dependency_digest = '2' * 64
            outputs = @([pscustomobject][ordered]@{ id = 'artifact.app'; role = 'executable'; relative_path = '.build/x64/Debug/RedSalamander.exe'; size_bytes = 1; sha256 = '3' * 64; runtime_closure = @() })
        }
        $arguments = @{
            EntryContract = $entry; CoveragePlanDigest = $coveragePlanDigest; WorkspaceSnapshot = $workspace; BuildReceipt = $receipt
            InputSetDigests = @{ 'source.commands' = '4' * 64 }; RunnerDependencyDigests = @{ 'tools/run-all-tests.ps1' = '5' * 64 }
            CapabilityInputs = @{ 'env.interactive-desktop' = $true }; QuarantineDigest = '6' * 64
            ClassificationMode = 'enabled'; ArtifactEpoch = 0
        }
        $baseline = Get-RSValidationEntryFingerprint @arguments
        (Get-RSValidationEntryFingerprint @arguments).Fingerprint | Should Be $baseline.Fingerprint
        $baseline.ComponentDigestVersion | Should Be 'validation-entry-components.v1'
        @($baseline.ComponentDigests.PSObject.Properties).Count | Should Be 11
        $baseline.ComputationMs | Should BeGreaterThan -1

        foreach ($mutation in @(
                @{ Name = 'CoveragePlanDigest'; Value = '7' * 64; Component = 'coverage' },
                @{ Name = 'WorkspaceSnapshot'; Value = [pscustomobject]@{ snapshot_id = '8' * 64 }; Component = 'source_input' },
                @{ Name = 'InputSetDigests'; Value = @{ 'source.commands' = '9' * 64 }; Component = 'logical_input' },
                @{ Name = 'RunnerDependencyDigests'; Value = @{ 'tools/run-all-tests.ps1' = '9' * 64 }; Component = 'runner' },
                @{ Name = 'CapabilityInputs'; Value = @{ 'env.interactive-desktop' = $false }; Component = 'capability' },
                @{ Name = 'QuarantineDigest'; Value = '0' * 64; Component = 'environment' },
                @{ Name = 'ClassificationMode'; Value = 'disabled'; Component = 'environment' },
                @{ Name = 'ArtifactEpoch'; Value = 1; Component = 'artifact_epoch' }
            )) {
            $changed = @{} + $arguments
            $changed[$mutation.Name] = $mutation.Value
            $changedFingerprint = Get-RSValidationEntryFingerprint @changed
            $changedFingerprint.Fingerprint | Should Not Be $baseline.Fingerprint
            @($changedFingerprint.ComponentDigests.PSObject.Properties | Where-Object {
                    $_.Value -ne $baseline.ComponentDigests.PSObject.Properties[$_.Name].Value
                }).Name | Should Be @($mutation.Component)
        }
        $changedReceipt = $receipt.PSObject.Copy()
        $changedReceipt.receipt_id = '9' * 64
        $changed = @{} + $arguments
        $changed.BuildReceipt = $changedReceipt
        $receiptFingerprint = Get-RSValidationEntryFingerprint @changed
        @($receiptFingerprint.ComponentDigests.PSObject.Properties | Where-Object {
                $_.Value -ne $baseline.ComponentDigests.PSObject.Properties[$_.Name].Value
            }).Name | Should Be @('build_artifact')
        $changedEntry = $entry.PSObject.Copy()
        $changedEntry.arguments = @('--commands-selftest', '--selftest-repeat=2')
        $changed = @{} + $arguments
        $changed.EntryContract = $changedEntry
        $argumentFingerprint = Get-RSValidationEntryFingerprint @changed
        @($argumentFingerprint.ComponentDigests.PSObject.Properties | Where-Object {
                $_.Value -ne $baseline.ComponentDigests.PSObject.Properties[$_.Name].Value
            }).Name | Should Be @('arguments')

        $changedEntry = $entry.PSObject.Copy()
        $changedEntry.display_name = 'Changed shared contract'
        $changed = @{} + $arguments; $changed.EntryContract = $changedEntry
        (Get-RSValidationEntryFingerprint @changed).ComponentDigests.shared_contract | Should Not Be $baseline.ComponentDigests.shared_contract

        $changedEntry = $entry.PSObject.Copy()
        $changedEntry.outcome_contract = $entry.outcome_contract.PSObject.Copy()
        $changedEntry.outcome_contract.pass_required_cases = @('different')
        $changed = @{} + $arguments; $changed.EntryContract = $changedEntry
        (Get-RSValidationEntryFingerprint @changed).ComponentDigests.outcome_policy | Should Not Be $baseline.ComponentDigests.outcome_policy
    }
}

Describe 'Operation Startrail atomic records and fixtures' {
    It 'publishes schema-valid UTF-8 LF records for absent and present destinations' {
        $root = Join-Path $TestDrive 'records'
        [void](New-Item -ItemType Directory -Path $root)
        $path = Join-Path $root 'snapshot.json'
        $schema = Join-Path $repoRoot 'Specs\Testing\WorkspaceSnapshot.schema.json'
        $digest = 'a' * 64
        $record = [ordered]@{
            schema = 'red-salamander.workspace-snapshot.v1'
            canonicalization = 'rs-canonical-json-v1'
            snapshot_id = $digest
            repo_identity = $digest
            git_head = 'b' * 40
            base = $null
            files = @()
            acquisition_before = $digest
            acquisition_after = $digest
            acquisition_attempts = 1
            git_discovery_ms = 2
            content_hashing_ms = 2
            canonicalization_ms = 1
            computation_ms = 5
            created_utc = '2026-08-13T12:00:00Z'
        }
        Write-RSAtomicCanonicalJsonFile -LiteralPath $path -InputObject $record -SchemaPath $schema -AllowedRoot $root
        $firstBytes = [IO.File]::ReadAllBytes($path)
        $firstBytes[0] | Should Not Be 0xef
        $firstBytes[$firstBytes.Length - 1] | Should Be 10
        (Read-RSValidationJsonFile -LiteralPath $path -SchemaPath $schema -AllowedRoot $root).snapshot_id | Should Be $digest
        $record.created_utc = '2026-08-13T12:01:00Z'
        Write-RSAtomicCanonicalJsonFile -LiteralPath $path -InputObject $record -SchemaPath $schema -AllowedRoot $root
        (Read-RSValidationJsonFile -LiteralPath $path -SchemaPath $schema -AllowedRoot $root).created_utc | Should Be '2026-08-13T12:01:00Z'
        @(Get-ChildItem -LiteralPath $root -Filter '*.tmp').Count | Should Be 0
    }

    It 'creates immutable records once without replacing existing bytes' {
        $root = Join-Path $TestDrive 'create-only-records'
        [void](New-Item -ItemType Directory -Path $root)
        $path = Join-Path $root 'snapshot.json'
        $schema = Join-Path $repoRoot 'Specs\Testing\WorkspaceSnapshot.schema.json'
        $digest = 'a' * 64
        $record = [ordered]@{
            schema = 'red-salamander.workspace-snapshot.v1'; canonicalization = 'rs-canonical-json-v1'
            snapshot_id = $digest; repo_identity = $digest; git_head = 'b' * 40; base = $null; files = @()
            acquisition_before = $digest; acquisition_after = $digest; acquisition_attempts = 1
            git_discovery_ms = 2; content_hashing_ms = 2; canonicalization_ms = 1; computation_ms = 5
            created_utc = '2026-08-13T12:00:00Z'
        }
        Write-RSAtomicCanonicalJsonFile -LiteralPath $path -InputObject $record -SchemaPath $schema `
            -AllowedRoot $root -CreateNew
        $originalBytes = [IO.File]::ReadAllBytes($path)
        $record.created_utc = '2026-08-13T12:01:00Z'
        (Test-RSActionThrows {
                Write-RSAtomicCanonicalJsonFile -LiteralPath $path -InputObject $record -SchemaPath $schema `
                    -AllowedRoot $root -CreateNew
            }) | Should Be $true
        [Convert]::ToHexString([IO.File]::ReadAllBytes($path)) | Should Be ([Convert]::ToHexString($originalBytes))
        @(Get-ChildItem -LiteralPath $root -Filter '*.tmp').Count | Should Be 0
    }

    It 'rejects duplicate JSON keys before deserialization' {
        $root = Join-Path $TestDrive 'duplicates'
        [void](New-Item -ItemType Directory -Path $root)
        $path = Join-Path $root 'duplicate.json'
        [IO.File]::WriteAllText($path, "{`"schema`":`"a`",`"schema`":`"b`"}`n", [Text.UTF8Encoding]::new($false))
        $schema = Join-Path $repoRoot 'Specs\Testing\WorkspaceSnapshot.schema.json'
        (Test-RSActionThrows { Read-RSValidationJsonFile -LiteralPath $path -SchemaPath $schema -AllowedRoot $root }) | Should Be $true
    }

    It 'accepts only exact canonical bytes, key order, escapes, and NFC strings' {
        $root = Join-Path $TestDrive 'canonical-reader'
        [void](New-Item -ItemType Directory -Path $root)
        $schema = Join-Path $root 'record.schema.json'
        $schemaText = '{"$schema":"http://json-schema.org/draft-07/schema#","additionalProperties":false,"properties":{"a":{"type":"string"},"b":{"type":"integer"}},"required":["a","b"],"type":"object"}'
        [IO.File]::WriteAllText($schema, $schemaText, [Text.UTF8Encoding]::new($false))
        $path = Join-Path $root 'record.json'

        [IO.File]::WriteAllText($path, "{`"a`":`"é`",`"b`":1}`n", [Text.UTF8Encoding]::new($false))
        (Read-RSValidationJsonFile -LiteralPath $path -SchemaPath $schema -AllowedRoot $root).a | Should Be 'é'

        foreach ($nonCanonical in @(
                "{`n  `"a`": `"é`",`n  `"b`": 1`n}`n",
                "{`"b`":1,`"a`":`"é`"}`n",
                "{`"a`":`"\u00e9`",`"b`":1}`n",
                "{`"a`":`"e$([char]0x0301)`",`"b`":1}`n")) {
            [IO.File]::WriteAllText($path, $nonCanonical, [Text.UTF8Encoding]::new($false))
            (Test-RSActionThrows {
                    Read-RSValidationJsonFile -LiteralPath $path -SchemaPath $schema -AllowedRoot $root
                }) | Should Be $true
        }
    }

    It 'enforces node and string bounds during streaming validation' {
        $root = Join-Path $TestDrive 'streaming-bounds'
        [void](New-Item -ItemType Directory -Path $root)
        $schema = Join-Path $root 'record.schema.json'
        [IO.File]::WriteAllText($schema,
            '{"$schema":"http://json-schema.org/draft-07/schema#"}',
            [Text.UTF8Encoding]::new($false))
        $path = Join-Path $root 'record.json'

        [IO.File]::WriteAllText($path, ('[' + ((@('null') * 100000) -join ',') + "]`n"),
            [Text.UTF8Encoding]::new($false))
        (Test-RSActionThrows {
                Read-RSValidationJsonFile -LiteralPath $path -SchemaPath $schema -AllowedRoot $root
            }) | Should Be $true

        [IO.File]::WriteAllText($path, ('"' + ('a' * (1MB + 1)) + '"' + "`n"),
            [Text.UTF8Encoding]::new($false))
        (Test-RSActionThrows {
                Read-RSValidationJsonFile -LiteralPath $path -SchemaPath $schema -AllowedRoot $root
            }) | Should Be $true
    }

    It 'rejects record paths outside the allowed root' {
        $root = Join-Path $TestDrive 'bounded-root'
        [void](New-Item -ItemType Directory -Path $root)
        $outside = Join-Path $TestDrive 'outside.json'
        $schema = Join-Path $repoRoot 'Specs\Testing\WorkspaceSnapshot.schema.json'
        (Test-RSActionThrows { Write-RSAtomicCanonicalJsonFile -LiteralPath $outside -InputObject ([ordered]@{}) -SchemaPath $schema -AllowedRoot $root }) | Should Be $true
    }

    It 'creates the reusable dirty staged untracked delete rename and space fixture' {
        $fixture = New-RSValidationGitFixture -Root (Join-Path $TestDrive 'git fixture')
        $status = @(& git -C $fixture.Root status --porcelain=v1 --untracked-files=all)
        ($status -join "`n") | Should Match 'dirty\.txt'
        ($status -join "`n") | Should Match 'staged\.txt'
        ($status -join "`n") | Should Match 'untracked\.txt'
        ($status -join "`n") | Should Match 'delete\.txt'
        ($status -join "`n") | Should Match 'rename\.txt -> renamed\.txt'
        $fixture.SpacePath | Should Be 'path with spaces/tracked.txt'
    }
}
