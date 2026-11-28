#Requires -Version 7.4
<#
.SYNOPSIS
Collect paired native x64 Release small-file Copy measurements.
.DESCRIPTION
Uses the exact noninteractive FileOps_SmallFileCopyCost case on one host.
An ABBA or BAAB block contains four fresh processes. Each observes a copy of
2,000 one-KiB files through the native Local route and the Dummy bridge.
Keeps the source, verified build receipts, invocation archives and raw rows.
The report records the measured comparison for performance review.
.PARAMETER BaselineRevision
Local Git revision of the baseline source.
.PARAMETER CandidateRevision
Local Git revision of the candidate source. Defaults to HEAD.
.PARAMETER Blocks
Number of four-observation blocks.
.PARAMETER Seed
Schedule seed. The default picks and records a random seed.
.PARAMETER CalibrationAA
Collect both labels using the same source revision.
.PARAMETER TestRoot
Existing marked X:\RedSalamander.Perf root. Defaults to REDSALAMANDER_TEST_ROOT.
.PARAMETER OutputDirectory
Optional output parent in the marked root's evidence/paired-copy directory.
.NOTES
Requires native x64 Windows, the exact clean CI engine, the pinned native toolchain
and a quiet host. Creates retained worktrees and binaries under the checkout's
.build and evidence under the marked test root. Cold restores can require network.
Use -WhatIf to inspect the source and test-root preconditions before build or execution.
.EXAMPLE
./performance-paired.ps1 -BaselineRevision main~1 -CandidateRevision main -TestRoot Z:\RedSalamander.Perf -WhatIf
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Low')]
param(
    [Parameter(Mandatory)][string]$BaselineRevision,
    [string]$CandidateRevision = 'HEAD',
    [ValidateRange(5, 20)][int]$Blocks = 5,
    [ValidateRange(-1, [int]::MaxValue)][int]$Seed = -1,
    [switch]$CalibrationAA,
    [string]$TestRoot = $env:REDSALAMANDER_TEST_ROOT,
    [string]$OutputDirectory = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (!$IsWindows -or [Runtime.InteropServices.RuntimeInformation]::OSArchitecture -ne 'X64' -or
    [Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture -ne 'X64') {
    throw 'Use native x64 Windows PowerShell 7 for the small-file Copy comparison.'
}
$root = $PSScriptRoot
$engineRoot = Join-Path $root '.ci-engine'
Import-Module (Join-Path $engineRoot 'Tools/Modules/Ci.Policy.psm1')
Import-Module (Join-Path $root 'Tools/Modules/Build/BuildEvidence.psm1')
Import-Module (Join-Path $root 'Tools/Modules/Testing/TestSandbox.psm1')
Import-Module (Join-Path $root 'Tools/Modules/Testing/R0ePairedValidation.psm1')

$caseName = 'FileOps_SmallFileCopyCost'
$metricName = 'FileOps.SelfTest.SmallFileCopy'
$profile = '.build/x64/Release'
$hostPath = Join-Path $PSHOME 'pwsh.exe'

function Get-CleanHead([string]$Repository) {
    $head = Resolve-CiCommit $Repository 'HEAD'
    $status = @(git -C $Repository status --porcelain=v1)
    if ($LASTEXITCODE -or $status.Count) { throw "Expected a clean checkout at $Repository." }
    return $head
}

function Get-CopyStudyReceipt([string]$Repository, [string]$Source) {
    if ((Get-CleanHead $Repository) -cne $Source) { throw 'The study worktree source moved.' }
    $receipt = Read-RSBuildReceipt -BuildOutputDir (Join-Path $Repository $profile) -SchemaPath (Join-Path $Repository 'Specs/Build/BuildReceipt.schema.json')
    $snapshot = Get-RSBuildSourceSnapshot -RepoRoot $Repository
    if (!$receipt -or $receipt.git_head -cne $Source -or $receipt.target -cne 'solution' -or !$receipt.tests_enabled -or
        $receipt.official_release -or 'build-number=1' -cnotin $receipt.build_arguments -or
        !(Test-RSBuildReceiptForArtifactUse -Receipt $receipt -SourceSnapshot $snapshot -Platform x64 -Configuration Release -RepoRoot $Repository -RequiredRelativePath "$profile/RedSalamander.exe")) {
        throw "The test-enabled x64 Release receipt is incomplete or incompatible at $Repository."
    }
    return $receipt
}

function Invoke-StudyCommand([string]$Repository, [string]$Entry, [string[]]$Arguments, [string]$LogPath) {
    Push-Location -LiteralPath $Repository
    try {
        & $hostPath -NoProfile -NonInteractive -File (Join-Path $Repository $Entry) @Arguments *> $LogPath
        return $LASTEXITCODE
    }
    finally {
        Pop-Location
    }
}

function Get-CopyObservation($Archive, [string]$Id, [string]$Tag, [string]$Repository, [int]$RunnerExit, [DateTimeOffset]$Started, [DateTimeOffset]$Ended) {
    $summary = $Archive.summary_text | ConvertFrom-Json -AsHashtable
    if ($summary.schema -cne 'red-salamander.run-all-tests.v1' -or $summary.run_id -cne $Id -or
        $summary.test_root -ine $TestRoot -or $summary.platform -cne 'x64' -or $summary.configuration -cne 'Release' -or
        $summary.suite -cne 'FileOps' -or $summary.case_filter -cne $caseName -or $summary.exit_code -ne $RunnerExit -or
        $RunnerExit -notin @(0, 2) -or $summary.timeout_scale -ne 1 -or $summary.quarantine.has_blocking_entries -or
        $summary.passed -ne 3 -or $summary.total -ne 3 -or $summary.failed -ne 0 -or
        $summary.skipped -ne 0 -or $summary.executable -ine (Join-Path $Repository "$profile/RedSalamander.exe") -or
        [DateTimeOffset]$summary.run_started_utc -lt $Started -or [DateTimeOffset]$summary.run_ended_utc -gt $Ended -or
        [DateTimeOffset]$summary.run_ended_utc -lt [DateTimeOffset]$summary.run_started_utc) {
        throw 'The runner summary does not bind a passed sample of the selected case.'
    }
    $native = $Archive.native
    $expectedCases = @('Setup', $caseName, 'Cleanup_RestorePluginConfig')
    if ($native.suite -cne 'FileOperations' -or $native.repeat_count -ne 1 -or @($native.cases).Count -ne $expectedCases.Count) {
        throw 'The case report has an unexpected suite or selection.'
    }
    foreach ($name in $expectedCases) {
        if (@($native.cases | Where-Object { $_.name -ceq $name -and $_.status -ceq 'passed' -and $_.repeat_index -eq 1 }).Count -ne 1) {
            throw "Expected one passing copy-study case named $name."
        }
    }
    if ($Archive.metrics.Count -ne 2) { throw 'The case must produce exactly two route observations.' }
    $raw = [ordered]@{}
    foreach ($row in $Archive.metrics) {
        if ($row.detail -cnotmatch "^route=(?<route>native|bridge) files=2000 repeat=0 tag=$([Regex]::Escape($Tag)) concurrency=(?<concurrency>[0-9]+)$") {
            throw 'The copy-study row has an unexpected workload, label or route.'
        }
        $route = $Matches.route
        foreach ($field in 'hr', 'value0', 'value1', 'value2', 'process', 'thread') {
            if ($row.$field -isnot [int] -and $row.$field -isnot [long]) { throw "Invalid numeric observation field $field." }
        }
        if ($row.process -le 0 -or $row.thread -le 0 -or $raw.Contains($route) -or $row.scenario -cne 'SelfTest' -or $row.build -cne 'Release' -or
            $row.hr -ne 0 -or $row.value0 -le 0 -or $row.value1 -ne 2000 -or
            [string]::IsNullOrWhiteSpace($row.machineHash) -or [string]::IsNullOrWhiteSpace($row.runId) -or
            ($route -ceq 'native' -and $row.value2 -ne 0) -or ($route -ceq 'bridge' -and $row.value2 -le 0)) {
            throw "The $route observation does not prove the expected completed copy."
        }
        $raw[$route] = $row
    }
    foreach ($field in 'machineHash', 'runId', 'process', 'thread') {
        if ($raw.native.$field -cne $raw.bridge.$field) { throw "The copy routes disagree on $field." }
    }
    return $raw
}

$driver = Get-CleanHead $root
$engine = Get-CleanHead $engineRoot
$baseline = Resolve-CiCommit $root $BaselineRevision
$candidate = Resolve-CiCommit $root $CandidateRevision
if (($baseline -ceq $candidate) -ne [bool]$CalibrationAA) {
    throw 'Use distinct revisions for a comparison and one revision for CalibrationAA.'
}
foreach ($revision in @($driver, $baseline, $candidate) | Select-Object -Unique) {
    $manifest = git -C $root show "$($revision):.ci/project.json" | ConvertFrom-Json
    if ($LASTEXITCODE -or $manifest.repository -cne 'RedSalamanders/RedSalamander' -or $manifest.engine.ref -cne $engine) {
        throw "Revision $revision must use the same reviewed RedSalamander engine."
    }
}
$inputs = [ordered]@{}
foreach ($path in 'RedSalamander/SelfTest/FileOperations/FolderWindow.FileOperations.Tests.SmallFileCost.cpp',
    'Tools/Modules/Testing/TestSuitePlan.psm1', 'Dependencies/DxUi.lock.json', 'Dependencies/RedMonitor.lock.json',
    'vcpkg.json', 'vcpkg-tool.json', 'vcpkg-configuration.json') {
    $old = Invoke-CiGit $root @('rev-parse', '--verify', "$($baseline):$path") -AllowFailure
    $new = Invoke-CiGit $root @('rev-parse', '--verify', "$($candidate):$path") -AllowFailure
    if ($old.exitCode -and $new.exitCode -and $path -ceq 'vcpkg-configuration.json') { continue }
    if ($old.exitCode -or $new.exitCode -or $old.stdout.Trim() -cne $new.stdout.Trim()) {
        throw "Paired input must match at $path."
    }
    $inputs[$path] = $old.stdout.Trim()
}
if ([string]::IsNullOrWhiteSpace($TestRoot)) { throw 'Supply the existing marked test root with -TestRoot or REDSALAMANDER_TEST_ROOT.' }
$TestRoot = Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path $TestRoot
$evidenceParent = Join-Path $TestRoot 'evidence/paired-copy'
if ($OutputDirectory) {
    $parent = Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path $OutputDirectory -RequireDescendant
    if ($parent -ine $evidenceParent -and !$parent.StartsWith("$evidenceParent\", [StringComparison]::OrdinalIgnoreCase)) {
        throw 'OutputDirectory must remain in the paired-copy evidence directory.'
    }
    $evidenceParent = $parent
}
if (!$PSCmdlet.ShouldProcess("$baseline and $candidate", 'Collect the paired small-file Copy study')) { return }
if ($env:CI_NATIVE_BUILDS_DISABLED -eq 'true') { throw 'Native benchmark execution is forbidden in compiler-free checks.' }

$studyId = [Guid]::NewGuid().ToString('N')
$buildParent = Assert-CiContainedPath -Root $root -Path (Join-Path $root ".build/paired-copy/$studyId")
$evidence = Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path (Join-Path $evidenceParent $studyId) -RequireDescendant
if ((Test-Path -LiteralPath $buildParent) -or (Test-Path -LiteralPath $evidence)) { throw 'Use a fresh study identity.' }
[void][IO.Directory]::CreateDirectory($buildParent)
[void][IO.Directory]::CreateDirectory($evidence)
if ($Seed -eq -1) { $Seed = [Security.Cryptography.RandomNumberGenerator]::GetInt32([int]::MaxValue) }
$random = [Random]::new($Seed)
$schedule = @(1..$Blocks | ForEach-Object { if ($random.Next(2) -eq 0) { 'ABBA' } else { 'BAAB' } })
$report = [ordered]@{
    schema = 'red-salamander.paired-copy.v1'
    status = 'incomplete'
    driver = $driver
    engine = $engine
    baseline = $baseline
    candidate = $candidate
    calibrationAA = [bool]$CalibrationAA
    platform = 'x64'
    configuration = 'Release'
    case = $caseName
    fileCount = 2000
    seed = $Seed
    schedule = $schedule
    inputs = $inputs
    builds = [ordered]@{}
    observations = [Collections.Generic.List[object]]::new()
    comparisons = [Collections.Generic.List[object]]::new()
}
$studyPath = Join-Path $evidence 'study.json'
$envValues = @{
    'REDSALAMANDER_FILEOPS_SMALLFILE_COUNT' = '2000'
    'REDSALAMANDER_FILEOPS_SMALLFILE_REPEATS' = '1'
    'REDSALAMANDER_FILEOPS_SMALLFILE_TAG' = $null
}
$previousEnv = @{}
foreach ($name in $envValues.Keys) { $previousEnv[$name] = [Environment]::GetEnvironmentVariable($name, 'Process') }

try {
    $trees = [ordered]@{}
    $sources = [ordered]@{ A = $baseline; B = $candidate }
    foreach ($label in $sources.Keys) {
        $source = $sources[$label]
        $tree = Assert-CiContainedPath -Root $buildParent -Path (Join-Path $buildParent $label)
        $trees[$label] = $tree
        git -C $root worktree add --detach --no-checkout $tree $source *> (Join-Path $evidence "$label-checkout.log")
        if ($LASTEXITCODE) { throw "Unable to allocate study worktree $label." }
        git -C $tree sparse-checkout set --no-cone '/*' '!/Specs/Plans/WIP/Notes_Scratch.md' *>> (Join-Path $evidence "$label-checkout.log")
        if ($LASTEXITCODE) { throw "Unable to configure study worktree $label." }
        git -C $tree read-tree --reset -u $source *>> (Join-Path $evidence "$label-checkout.log")
        if ($LASTEXITCODE) { throw "Unable to materialize study worktree $label." }
        git clone --quiet --no-tags --no-checkout --shared $engineRoot (Join-Path $tree '.ci-engine') *>> (Join-Path $evidence "$label-checkout.log")
        if ($LASTEXITCODE) { throw "Unable to bind the engine for $label." }
        git -C (Join-Path $tree '.ci-engine') checkout --detach $engine *>> (Join-Path $evidence "$label-checkout.log")
        if ($LASTEXITCODE) { throw "Unable to select the pinned engine for $label." }
        if ((Invoke-StudyCommand $tree 'restore.ps1' @('-Platform','x64','-Configuration','Release') (Join-Path $evidence "$label-restore.log")) -ne 0) {
            throw "Pinned dependency restore failed for $label."
        }
        if ((Invoke-StudyCommand $tree 'build.ps1' @('-Platform','x64','-Configuration','Release','-EnableTests','-BuildNumber','1') (Join-Path $evidence "$label-build.log")) -ne 0) {
            throw "Test-enabled Release build failed for $label."
        }
        $report.builds[$label] = Get-CopyStudyReceipt $tree $source
    }
    foreach ($id in 'toolchain_identity_digest', 'vcpkg_identity_digest', 'ghostty_runtime_identity') {
        if ($report.builds.A.$id -cne $report.builds.B.$id) { throw "The comparison changes input identity $id." }
    }

    foreach ($name in $envValues.Keys) { [Environment]::SetEnvironmentVariable($name, $envValues[$name], 'Process') }
    $machine = $null
    for ($block = 0; $block -lt $schedule.Count; ++$block) {
        $blockRows = [Collections.Generic.List[object]]::new()
        for ($round = 0; $round -lt 4; ++$round) {
            $label = [string]$schedule[$block][$round]
            $tree = $trees[$label]
            $source = $sources[$label]
            $receipt = Get-CopyStudyReceipt $tree $source
            if ($receipt.receipt_id -cne $report.builds[$label].receipt_id) { throw "The receipt for $label changed before the sample." }
            $tag = '{0}-{1:D2}-{2}' -f $label, ($block + 1), ($round + 1)
            $env:REDSALAMANDER_FILEOPS_SMALLFILE_TAG = $tag
            $log = Join-Path $evidence "$tag.log"
            $id = '{0:yyyyMMddTHHmmssZ}-{1}-{2}' -f [DateTime]::UtcNow, $PID, [Guid]::NewGuid().ToString('N')
            $record = [ordered]@{ block = $block + 1; round = $round + 1; side = $label; source = $source; receipt = $receipt.receipt_id; runId = $id; log = $log }
            $report.observations.Add($record)
            $started = [DateTimeOffset]::UtcNow
            $sampleArguments = @(
                '-Suite', 'FileOps', '-CaseFilter', $caseName, '-NonInteractive', '-SkipBuild',
                '-Platform', 'x64', '-Configuration', 'Release', '-BuildNumber', '1',
                '-TimeoutMultiplier', '1', '-SelfTestRepeat', '1', '-RequireCleanQuarantine', '-TestRoot', $TestRoot, '-RunId', $id
            )
            $record.runnerExit = Invoke-StudyCommand $tree 'Tools/Run-AllTests.ps1' $sampleArguments $log
            $record.logSha256 = (Get-FileHash -LiteralPath $log -Algorithm SHA256).Hash.ToLowerInvariant()
            $ended = [DateTimeOffset]::UtcNow
            if ($record.runnerExit -notin @(0, 2)) { throw "The copy-study sample failed. See $log." }
            $archive = Read-RSR0eArchive -TestRoot $TestRoot -RunId $id -MetricName $metricName
            $record.intentSha256 = $archive.intent_sha256
            $record.artifacts = [ordered]@{}
            $samplePath = Assert-RSTestSandboxContainedPath -TestRoot $TestRoot -Path (Join-Path $evidence $tag) -RequireDescendant
            if (Test-Path -LiteralPath $samplePath) { throw 'The sample evidence directory already exists.' }
            [void][IO.Directory]::CreateDirectory($samplePath)
            foreach ($archivedFile in $archive.files.Values) {
                $leaf = [IO.Path]::GetFileName($archivedFile.path)
                $copy = Join-Path $samplePath $leaf
                [IO.File]::Copy($archivedFile.path, $copy, $false)
                $digest = (Get-FileHash -LiteralPath $copy -Algorithm SHA256).Hash.ToLowerInvariant()
                if ($digest -ine $archivedFile.sha256) { throw "Retained sample bytes changed at $leaf." }
                $record.artifacts["$tag/$leaf"] = $digest
            }
            $record.rows = Get-CopyObservation $archive $id $tag $tree $record.runnerExit $started $ended
            $postReceipt = Get-CopyStudyReceipt $tree $source
            if ($postReceipt.receipt_id -cne $receipt.receipt_id) { throw 'The receipt changed during the sample.' }
            if ($null -eq $machine) { $machine = $record.rows.native.machineHash }
            if ($record.rows.native.machineHash -cne $machine) { throw 'The study sampled more than one machine.' }
            $blockRows.Add($record)
        }
        $ratio = [ordered]@{ block = $block + 1; order = $schedule[$block] }
        foreach ($route in 'native', 'bridge') {
            $before = $blockRows | Where-Object side -CEQ 'A' | ForEach-Object { [double]$_.rows[$route].value0 } | Measure-Object -Average
            $after = $blockRows | Where-Object side -CEQ 'B' | ForEach-Object { [double]$_.rows[$route].value0 } | Measure-Object -Average
            $ratio[$route] = [ordered]@{ baselineMicroseconds = $before.Average; candidateMicroseconds = $after.Average; candidateToBaseline = $after.Average / $before.Average }
        }
        $report.comparisons.Add($ratio)
    }
    $report.status = 'collected'
}
catch {
    $report.failure = $_.Exception.Message
    throw
}
finally {
    foreach ($name in $previousEnv.Keys) { [Environment]::SetEnvironmentVariable($name, $previousEnv[$name], 'Process') }
    $report | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $studyPath -Encoding utf8NoBOM
    Write-Host "Paired Copy study evidence retained at $studyPath"
}
