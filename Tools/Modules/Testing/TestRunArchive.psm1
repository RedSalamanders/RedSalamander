# Private test-evidence module; import through its reviewed callers.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# The checked-in archive contract and its 2 MiB / 5 MiB limits were introduced
# by this commit. Older untouched evidence is grandfathered; any archive added
# or modified from this commit onward belongs to the governed inventory.
$script:RSTestRunArchiveContractCommit = '8b9e36191835cc71b5ee0dfeea4dddd5d942dd5a'

function ConvertTo-RSRepoRelativePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $root = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\', '/')
    $fullPath = if ([IO.Path]::IsPathRooted($Path)) {
        [IO.Path]::GetFullPath($Path)
    } else {
        [IO.Path]::GetFullPath((Join-Path $root $Path))
    }
    if (-not $fullPath.StartsWith($root + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Archive path is outside the repository: $Path"
    }

    return $fullPath.Substring($root.Length + 1)
}

function New-RSTestRunArchiveViolation {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Kind,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    return [pscustomobject]@{
        Kind = $Kind
        Path = $Path
        Message = $Message
    }
}

function Get-RSNearestRankPercentile {
    param(
        [Parameter(Mandatory = $true)]
        [long[]]$Samples,

        [Parameter(Mandatory = $true)]
        [ValidateSet(50, 95)]
        [int]$Percentile
    )

    if ($Samples.Count -eq 0) {
        throw 'Nearest-rank percentile requires at least one sample.'
    }
    $ordered = @($Samples | Sort-Object)
    $index = [Math]::Ceiling($ordered.Count * ($Percentile / 100.0)) - 1
    return [long]$ordered[[Math]::Max(0, [int]$index)]
}

function Read-RSTerminalVtUpgradeEvidence {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$RunPath
    )

    $root = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\', '/')
    $relativeRun = ConvertTo-RSRepoRelativePath -RepoRoot $root -Path $RunPath
    $parts = $relativeRun -split '[\\/]'
    if ($parts.Count -ne 5 -or $parts[0] -ne 'Specs' -or $parts[1] -ne 'TestRuns' -or $parts[3] -ne 'Terminal') {
        throw "Terminal VT evidence path must be Specs/TestRuns/<MachineHash>/Terminal/<RunId>: $RunPath"
    }
    $runRoot = Join-Path $root $relativeRun
    $resultsPath = Join-Path $runRoot 'results.json'
    if (-not (Test-Path -LiteralPath $resultsPath -PathType Leaf)) {
        throw "Terminal VT evidence is missing results.json: $relativeRun"
    }
    $schemaPath = Join-Path $root 'Specs\Terminal\TerminalVtUpgradeEvidence.schema.json'
    if (-not (Test-Path -LiteralPath $schemaPath -PathType Leaf)) {
        throw "Terminal VT evidence schema is missing: $schemaPath"
    }

    $json = Get-Content -LiteralPath $resultsPath -Raw
    if (-not (Test-Json -Json $json -SchemaFile $schemaPath -ErrorAction Stop)) {
        throw "Terminal VT evidence does not satisfy TerminalVtUpgradeEvidence.schema.json: $relativeRun"
    }
    return [pscustomobject]@{
        RelativeRun = ($relativeRun -replace '\\', '/')
        RunRoot = $runRoot
        Results = ($json | ConvertFrom-Json -ErrorAction Stop)
    }
}

function Get-RSTestRunArchiveViolations {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$Paths,

        [long]$MaxFileBytes = 2MB,

        [long]$MaxRunBytes = 5MB
    )

    $root = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\', '/')
    $runKeys = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $explicitArchivePaths = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $violations = [System.Collections.Generic.List[object]]::new()

    foreach ($path in $Paths | Sort-Object -Unique) {
        $relativePath = ConvertTo-RSRepoRelativePath -RepoRoot $root -Path $path
        $fullPath = Join-Path $root $relativePath
        if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
            continue
        }

        $parts = $relativePath -split '[\\/]'
        if ($parts.Count -lt 5 -or $parts[0] -ne 'Specs' -or $parts[1] -ne 'TestRuns') {
            continue
        }

        $fileInfo = Get-Item -LiteralPath $fullPath
        if ($fileInfo.Length -gt $MaxFileBytes) {
            $violations.Add([pscustomobject]@{
                Kind = 'FileSize'
                Path = $relativePath
                Message = "Checked-in TestRuns file is $($fileInfo.Length) bytes; maximum is $MaxFileBytes bytes."
            })
        }

        $runKey = ($parts[0..4] -join '/')
        [void]$runKeys.Add($runKey)
        [void]$explicitArchivePaths.Add(($relativePath -replace '\\', '/'))

        if ($fileInfo.Name -eq 'perf_metrics.jsonl') {
            $reader = [IO.File]::OpenText($fullPath)
            try {
                $lineNumber = 0
                while ($null -ne ($line = $reader.ReadLine())) {
                    ++$lineNumber
                    if ([string]::IsNullOrWhiteSpace($line)) {
                        continue
                    }
                    try {
                        $metric = $line | ConvertFrom-Json -ErrorAction Stop
                    } catch {
                        $violations.Add((New-RSTestRunArchiveViolation `
                            -Kind 'PerfJson' `
                            -Path $relativePath `
                            -Message "Performance metric row $lineNumber is not valid JSON."))
                        continue
                    }

                    $profile = $parts[2]
                    $machineHashProperty = if ($null -ne $metric) {
                        $metric.PSObject.Properties['machineHash']
                    } else {
                        $null
                    }
                    $machineHash = if ($null -ne $machineHashProperty) {
                        [string]$machineHashProperty.Value
                    } else {
                        ''
                    }
                    if ($profile -match '^[0-9a-fA-F]{8,40}$' -and -not [string]::IsNullOrWhiteSpace($machineHash) -and
                        -not [string]::Equals($profile, $machineHash, [StringComparison]::OrdinalIgnoreCase)) {
                        $violations.Add((New-RSTestRunArchiveViolation `
                            -Kind 'MachineProfile' `
                            -Path $relativePath `
                            -Message "Archive profile '$profile' does not match embedded machineHash '$machineHash' at metric row $lineNumber."))
                    }
                }
            } finally {
                $reader.Dispose()
            }
        }
    }

    $isExactGitWorktreeRoot = $false
    $gitTopLevel = & git -C $root rev-parse --show-toplevel 2>$null
    if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace([string]$gitTopLevel)) {
        $resolvedGitTopLevel = [IO.Path]::GetFullPath(([string]$gitTopLevel).Trim()).TrimEnd('\', '/')
        $isExactGitWorktreeRoot = [string]::Equals($resolvedGitTopLevel, $root, [StringComparison]::OrdinalIgnoreCase)
    }

    foreach ($runKey in $runKeys) {
        $runRoot = Join-Path $root ($runKey -replace '/', [IO.Path]::DirectorySeparatorChar)
        if (-not (Test-Path -LiteralPath $runRoot -PathType Container)) {
            continue
        }

        $runParts = $runKey -split '/'
        $resultsPath = Join-Path $runRoot 'results.json'
        $isTerminalVtEvidence = $false
        if ($runParts[3] -eq 'Terminal' -and (Test-Path -LiteralPath $resultsPath -PathType Leaf)) {
            try {
                $evidenceProbe = Get-Content -LiteralPath $resultsPath -Raw | ConvertFrom-Json -ErrorAction Stop
                $isTerminalVtEvidence = $evidenceProbe.schema -ceq 'red-salamander.terminal-vt-upgrade-evidence.v1' -or
                    $evidenceProbe.corpus.formatId -ceq 'red-salamander-terminal-vt-upgrade-corpus'
            } catch {
                # A malformed generic Terminal result remains owned by its own
                # area/version validator. VT evidence is identified by schema
                # or corpus identity instead of claiming every Terminal run.
            }
        }
        if ($isTerminalVtEvidence) {
            $requiredArtifacts = @('results.json', 'trace.txt', 'perf\perf_metrics.jsonl')
            foreach ($requiredArtifact in $requiredArtifacts) {
                $requiredPath = Join-Path $runRoot $requiredArtifact
                if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
                    $violations.Add((New-RSTestRunArchiveViolation `
                        -Kind 'RequiredArtifact' `
                        -Path $runKey `
                        -Message "Terminal VT evidence is missing '$requiredArtifact'."))
                }
            }

            if (Test-Path -LiteralPath $resultsPath -PathType Leaf) {
                try {
                    $resultsJson = Get-Content -LiteralPath $resultsPath -Raw
                    $results = $resultsJson | ConvertFrom-Json -ErrorAction Stop
                    $schemaPath = Join-Path $root 'Specs\Terminal\TerminalVtUpgradeEvidence.schema.json'
                    if (-not (Test-Path -LiteralPath $schemaPath -PathType Leaf) -or
                        -not (Test-Json -Json $resultsJson -SchemaFile $schemaPath -ErrorAction Stop)) {
                        $violations.Add((New-RSTestRunArchiveViolation `
                            -Kind 'EvidenceSchema' `
                            -Path "$runKey/results.json" `
                            -Message 'Terminal evidence does not satisfy TerminalVtUpgradeEvidence.schema.json.'))
                    } else {
                        if ($results.runId -cne $runParts[4]) {
                            $violations.Add((New-RSTestRunArchiveViolation `
                                -Kind 'EvidenceBinding' `
                                -Path "$runKey/results.json" `
                                -Message "Embedded runId '$($results.runId)' does not match archive run '$($runParts[4])'."))
                        }
                        if ($results.machineHash -cne $runParts[2]) {
                            $violations.Add((New-RSTestRunArchiveViolation `
                                -Kind 'EvidenceBinding' `
                                -Path "$runKey/results.json" `
                                -Message "Embedded machineHash '$($results.machineHash)' does not match archive profile '$($runParts[2])'."))
                        }

                        $perfPath = Join-Path $runRoot 'perf\perf_metrics.jsonl'
                        if (Test-Path -LiteralPath $perfPath -PathType Leaf) {
                            $metricRows = @(Get-Content -LiteralPath $perfPath |
                                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                                    ForEach-Object { $_ | ConvertFrom-Json -ErrorAction Stop })
                            $metricMap = @{
                                vtWrite = 'terminal.vt_upgrade.vt_write_us'
                                snapshot = 'terminal.vt_upgrade.snapshot_us'
                                hostedRender = 'terminal.vt_upgrade.hosted_render_us'
                            }
                            foreach ($measurementName in $metricMap.Keys) {
                                $rows = @($metricRows | Where-Object metric -CEQ $metricMap[$measurementName])
                                $expectedSamples = @($results.measurements.$measurementName.samples | ForEach-Object { [long]$_ })
                                $actualSamples = @($rows | ForEach-Object { [long]$_.value })
                                $rowsBound = @($rows | Where-Object {
                                        $_.machineHash -ceq $results.machineHash -and
                                        $_.runId -ceq $results.runId -and
                                        $_.build -ceq $results.configuration -and
                                        $_.commit -ceq $results.repositoryCommit -and
                                        $_.scenario -ceq 'terminal-vt-upgrade-corpus-v1'
                                    }).Count -eq $rows.Count
                                if ($rows.Count -ne $results.measurements.sampleCount -or
                                    (Compare-Object -ReferenceObject $expectedSamples -DifferenceObject $actualSamples -SyncWindow 0) -or
                                    -not $rowsBound) {
                                    $violations.Add((New-RSTestRunArchiveViolation `
                                        -Kind 'EvidenceBinding' `
                                        -Path "$runKey/perf/perf_metrics.jsonl" `
                                        -Message "Metric '$($metricMap[$measurementName])' does not exactly bind to the results samples and run identity."))
                                }
                            }
                        }
                    }
                } catch {
                    $violations.Add((New-RSTestRunArchiveViolation `
                        -Kind 'EvidenceJson' `
                        -Path "$runKey/results.json" `
                        -Message "Terminal evidence could not be parsed or validated: $($_.Exception.Message)"))
                }
            }
        }

        $aggregatePaths = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        if ($isExactGitWorktreeRoot) {
            $trackedPaths = & git -C $root ls-files --cached -- $runKey 2>$null
            if ($LASTEXITCODE -ne 0) {
                throw "git ls-files --cached -- $runKey failed while sizing the checked-in TestRuns run."
            }
            foreach ($trackedPath in $trackedPaths) {
                if (-not [string]::IsNullOrWhiteSpace($trackedPath)) {
                    [void]$aggregatePaths.Add($trackedPath.Trim())
                }
            }
            foreach ($explicitPath in $explicitArchivePaths) {
                if ($explicitPath.StartsWith($runKey + '/', [StringComparison]::OrdinalIgnoreCase)) {
                    [void]$aggregatePaths.Add($explicitPath)
                }
            }
        } else {
            foreach ($runFile in Get-ChildItem -LiteralPath $runRoot -File -Recurse -Force) {
                [void]$aggregatePaths.Add((ConvertTo-RSRepoRelativePath -RepoRoot $root -Path $runFile.FullName))
            }
        }

        [long]$runBytes = 0L
        foreach ($aggregatePath in $aggregatePaths) {
            $aggregateFullPath = Join-Path $root ($aggregatePath -replace '/', [IO.Path]::DirectorySeparatorChar)
            if (Test-Path -LiteralPath $aggregateFullPath -PathType Leaf) {
                $runBytes += [long](Get-Item -LiteralPath $aggregateFullPath).Length
            }
        }

        if ($runBytes -gt $MaxRunBytes) {
            $violations.Add([pscustomobject]@{
                Kind = 'RunSize'
                Path = $runKey
                Message = "Checked-in TestRuns run is $runBytes bytes; maximum is $MaxRunBytes bytes."
            })
        }
    }

    return @($violations)
}

function Compare-RSTerminalVtUpgradeEvidence {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$BaselineRunPath,

        [Parameter(Mandatory = $true)]
        [string]$CandidateRunPath
    )

    $violations = [System.Collections.Generic.List[object]]::new()
    try {
        $baseline = Read-RSTerminalVtUpgradeEvidence -RepoRoot $RepoRoot -RunPath $BaselineRunPath
        $candidate = Read-RSTerminalVtUpgradeEvidence -RepoRoot $RepoRoot -RunPath $CandidateRunPath
    } catch {
        $violations.Add((New-RSTestRunArchiveViolation `
            -Kind 'TerminalVtComparison' `
            -Path "$BaselineRunPath -> $CandidateRunPath" `
            -Message $_.Exception.Message))
        return [pscustomobject]@{
            Outcome = 'fail'
            BaselineRun = $BaselineRunPath
            CandidateRun = $CandidateRunPath
            Measurements = @()
            Violations = @($violations)
        }
    }

    $baselineResults = $baseline.Results
    $candidateResults = $candidate.Results
    $pairPath = "$($baseline.RelativeRun) -> $($candidate.RelativeRun)"
    $addViolation = {
        param([string]$Message)
        $violations.Add((New-RSTestRunArchiveViolation `
            -Kind 'TerminalVtComparison' `
            -Path $pairPath `
            -Message $Message))
    }

    foreach ($property in @('repositoryCommit', 'repositoryBranch', 'machineHash', 'configuration')) {
        if ($baselineResults.$property -cne $candidateResults.$property) {
            & $addViolation "Pair identity '$property' differs between baseline and candidate."
        }
    }
    foreach ($property in @('windowsBuild', 'cpu', 'architecture', 'compiler', 'windowsSdk', 'zig', 'vcpkgCommit')) {
        if ($baselineResults.environment.$property -cne $candidateResults.environment.$property) {
            & $addViolation "Environment identity '$property' differs between baseline and candidate."
        }
    }
    foreach ($property in @('formatId', 'version', 'sha256', 'inputByteCount')) {
        if ($baselineResults.corpus.$property -cne $candidateResults.corpus.$property) {
            & $addViolation "Corpus identity '$property' differs between baseline and candidate."
        }
    }
    if ($baselineResults.runtimeIdentitySha256 -ceq $candidateResults.runtimeIdentitySha256) {
        & $addViolation 'Baseline and candidate runtime identities must be distinct.'
    }
    if ($baselineResults.buildReceiptId -ceq $candidateResults.buildReceiptId) {
        & $addViolation 'Baseline and candidate build receipt identities must be distinct.'
    }

    $baselineControl = $baselineResults.behavior.withoutExpectedDeltas
    $candidateControl = $candidateResults.behavior.withoutExpectedDeltas
    if ($null -eq $baselineControl -or $null -eq $candidateControl -or
        $baselineControl.ptyOutputSha256 -cne $candidateControl.ptyOutputSha256 -or
        $baselineControl.snapshotSha256 -cne $candidateControl.snapshotSha256) {
        & $addViolation 'Control behavior digests differ after removing the reviewed expected delta.'
    }
    if ($candidateResults.behavior.ptyOutputSha256 -cne $candidateControl.ptyOutputSha256 -or
        $candidateResults.behavior.snapshotSha256 -cne $candidateControl.snapshotSha256) {
        & $addViolation 'Candidate behavior contains a delta beyond the reviewed DCS high-byte preservation change.'
    }
    $baselineDelta = @($baselineResults.behavior.expectedDeltas)
    $candidateDelta = @($candidateResults.behavior.expectedDeltas)
    if ($baselineDelta.Count -ne 1 -or $candidateDelta.Count -ne 1 -or
        $baselineDelta[0].id -cne $candidateDelta[0].id -or
        $baselineDelta[0].disposition -cne $candidateDelta[0].disposition) {
        & $addViolation 'Expected-delta declarations do not match the single reviewed corpus delta.'
    }

    $measurementRows = [System.Collections.Generic.List[object]]::new()
    foreach ($measurementName in @('vtWrite', 'snapshot', 'hostedRender')) {
        $baselineMeasurement = $baselineResults.measurements.$measurementName
        $candidateMeasurement = $candidateResults.measurements.$measurementName
        $baselineSamples = @($baselineMeasurement.samples | ForEach-Object { [long]$_ })
        $candidateSamples = @($candidateMeasurement.samples | ForEach-Object { [long]$_ })
        $baselineP50 = Get-RSNearestRankPercentile -Samples $baselineSamples -Percentile 50
        $baselineP95 = Get-RSNearestRankPercentile -Samples $baselineSamples -Percentile 95
        $candidateP50 = Get-RSNearestRankPercentile -Samples $candidateSamples -Percentile 50
        $candidateP95 = Get-RSNearestRankPercentile -Samples $candidateSamples -Percentile 95
        $maximumP95 = [long][Math]::Ceiling([decimal]$baselineP95 * [decimal]'1.10')
        $baselineSelfMaximum = [long][Math]::Ceiling([decimal]$baselineP95 * [decimal]'1.10')
        $candidateSelfMaximum = [long][Math]::Ceiling([decimal]$candidateP95 * [decimal]'1.10')

        if ($baselineSamples.Count -ne 200 -or $candidateSamples.Count -ne 200) {
            & $addViolation "Measurement '$measurementName' does not contain exactly 200 samples in both runs."
        }
        if ([long]$baselineMeasurement.p50 -ne $baselineP50 -or [long]$baselineMeasurement.p95 -ne $baselineP95 -or
            [long]$candidateMeasurement.p50 -ne $candidateP50 -or [long]$candidateMeasurement.p95 -ne $candidateP95) {
            & $addViolation "Measurement '$measurementName' recorded percentiles do not match nearest-rank recomputation."
        }
        if ([long]$baselineMeasurement.candidateMaximumP95 -ne $baselineSelfMaximum -or
            [long]$candidateMeasurement.candidateMaximumP95 -ne $candidateSelfMaximum) {
            & $addViolation "Measurement '$measurementName' local candidateMaximumP95 is not p95 x1.10 rounded up."
        }
        if ($candidateP95 -gt $maximumP95) {
            & $addViolation "Measurement '$measurementName' candidate p95 $candidateP95 exceeds baseline ceiling $maximumP95."
        }
        $measurementRows.Add([pscustomobject]@{
            Name = $measurementName
            BaselineP50 = $baselineP50
            CandidateP50 = $candidateP50
            BaselineP95 = $baselineP95
            CandidateP95 = $candidateP95
            CandidateMaximumP95 = $maximumP95
            Passed = $candidateP95 -le $maximumP95
        })
    }

    return [pscustomobject]@{
        Outcome = if ($violations.Count -eq 0) { 'pass' } else { 'fail' }
        BaselineRun = $baseline.RelativeRun
        CandidateRun = $candidate.RelativeRun
        Measurements = @($measurementRows)
        Violations = @($violations)
    }
}

function Get-RSChangedTestRunArchivePaths {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $paths = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $root = [IO.Path]::GetFullPath($RepoRoot)

    $hasOriginMaster = $false
    & git -C $root rev-parse --verify origin/master *> $null
    if ($LASTEXITCODE -eq 0) {
        $hasOriginMaster = $true
    }

    $commands = [System.Collections.Generic.List[object]]::new()
    if ($hasOriginMaster) {
        $commands.Add(@('diff', '--name-only', '--diff-filter=AM', 'origin/master...HEAD', '--', 'Specs/TestRuns'))
    }
    $commands.Add(@('diff', '--name-only', '--diff-filter=AM', '--', 'Specs/TestRuns'))
    $commands.Add(@('diff', '--cached', '--name-only', '--diff-filter=AM', '--', 'Specs/TestRuns'))

    foreach ($arguments in $commands) {
        $output = & git -C $root @arguments 2>$null
        if ($LASTEXITCODE -ne 0) {
            throw "git $($arguments -join ' ') failed while locating changed TestRuns artifacts."
        }
        foreach ($path in $output) {
            if (-not [string]::IsNullOrWhiteSpace($path)) {
                [void]$paths.Add($path.Trim())
            }
        }
    }

    return @($paths | Sort-Object)
}

function Get-RSTestRunArchiveInventoryPaths {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot
    )

    $root = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\', '/')
    $gitTopLevel = & git -C $root rev-parse --show-toplevel 2>$null
    if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace([string]$gitTopLevel)) {
        $resolvedGitTopLevel = [IO.Path]::GetFullPath(([string]$gitTopLevel).Trim()).TrimEnd('\', '/')
        if ([string]::Equals($resolvedGitTopLevel, $root, [StringComparison]::OrdinalIgnoreCase)) {
            $hasContractCommit = $false
            & git -C $root rev-parse --verify "$script:RSTestRunArchiveContractCommit^{commit}" *> $null
            if ($LASTEXITCODE -eq 0) {
                & git -C $root merge-base --is-ancestor $script:RSTestRunArchiveContractCommit HEAD *> $null
                $hasContractCommit = $LASTEXITCODE -eq 0
            }

            if ($hasContractCommit) {
                $paths = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                $contractRange = "$script:RSTestRunArchiveContractCommit^..HEAD"
                $commands = @(
                    @('diff', '--name-only', '--diff-filter=AM', $contractRange, '--', 'Specs/TestRuns'),
                    @('diff', '--name-only', '--diff-filter=AM', '--', 'Specs/TestRuns'),
                    @('diff', '--cached', '--name-only', '--diff-filter=AM', '--', 'Specs/TestRuns')
                )
                foreach ($arguments in $commands) {
                    $output = & git -C $root @arguments 2>$null
                    if ($LASTEXITCODE -ne 0) {
                        throw "git $($arguments -join ' ') failed while enumerating the governed TestRuns inventory."
                    }
                    foreach ($path in $output) {
                        if (-not [string]::IsNullOrWhiteSpace($path)) {
                            [void]$paths.Add($path.Trim())
                        }
                    }
                }
                $trackedPaths = @($paths)
            } else {
                $trackedPaths = & git -C $root ls-files --cached -- Specs/TestRuns 2>$null
                if ($LASTEXITCODE -ne 0) {
                    throw 'git ls-files --cached -- Specs/TestRuns failed while enumerating the checked-in TestRuns inventory.'
                }
            }

            return @($trackedPaths |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                    ForEach-Object { $_.Trim() } |
                    Where-Object { Test-Path -LiteralPath (Join-Path $root $_) -PathType Leaf } |
                    Sort-Object -Unique)
        }
    }

    $testRunsRoot = Join-Path $root 'Specs\TestRuns'
    if (-not (Test-Path -LiteralPath $testRunsRoot -PathType Container)) {
        return @()
    }

    return @(Get-ChildItem -LiteralPath $testRunsRoot -File -Recurse -Force |
            ForEach-Object { ConvertTo-RSRepoRelativePath -RepoRoot $root -Path $_.FullName } |
            Sort-Object -Unique)
}

function Get-RSTestRunArchivePathsForRuns {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$RunPath
    )

    $root = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\', '/')
    $paths = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($candidate in $RunPath) {
        $relativeRun = ConvertTo-RSRepoRelativePath -RepoRoot $root -Path $candidate
        $parts = $relativeRun -split '[\\/]'
        if ($parts.Count -lt 5 -or $parts[0] -ne 'Specs' -or $parts[1] -ne 'TestRuns') {
            throw "Run path is not under Specs/TestRuns/<Profile>/<Area>/<RunId>: $candidate"
        }

        $fullRun = Join-Path $root $relativeRun
        if (-not (Test-Path -LiteralPath $fullRun -PathType Container)) {
            throw "TestRuns run path does not exist: $candidate"
        }

        foreach ($file in Get-ChildItem -LiteralPath $fullRun -File -Recurse -Force) {
            [void]$paths.Add((ConvertTo-RSRepoRelativePath -RepoRoot $root -Path $file.FullName))
        }
    }

    return @($paths | Sort-Object)
}

Export-ModuleMember -Function `
    Get-RSTestRunArchiveViolations, `
    Get-RSChangedTestRunArchivePaths, `
    Get-RSTestRunArchiveInventoryPaths, `
    Get-RSTestRunArchivePathsForRuns, `
    Compare-RSTerminalVtUpgradeEvidence
