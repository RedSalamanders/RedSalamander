Set-StrictMode -Version Latest

$validationFingerprintModule = Join-Path $PSScriptRoot 'ValidationFingerprint.psm1'
Import-Module $validationFingerprintModule -Force -Scope Local -ErrorAction Stop

function Get-RSTestPlanSortedIds {
    param([string[]]$Value = @())

    $items = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($item in $Value) {
        Assert-RSValidationIdentifier -Id $item
        if (-not $items.Add($item)) { throw "Duplicate test-plan identifier '$item'." }
    }
    $result = [string[]]@($items)
    [Array]::Sort($result, [StringComparer]::Ordinal)
    return $result
}

function Get-RSTestRunPlanEntryPolicy {
    param(
        [Parameter(Mandatory = $true)][string]$Id,
        [Parameter(Mandatory = $true)][string]$Kind,
        [string[]]$Arguments = @(),
        [bool]$RequiresInteractiveDesktop = $false
    )

    $family = $Id.Split('.')[0]
    $impactTags = @("test.$family")
    $inputSets = @("source.$family")
    $artifactIds = if ($Kind -in @('Pester', 'PowerShellScript')) { @() } else { @("artifact.$Id") }
    $artifactReads = @($artifactIds)
    $artifactWrites = @()
    $produces = @()
    $invalidates = @()
    $cachePolicy = if ($Kind -in @('Executable', 'CppUnitTest')) { 'ExactArtifact' } else { 'Never' }
    $environment = @('env.profile', 'env.windows-11')
    $coverageFamily = $null
    $exactCases = @()
    $order = 'isolated'
    $resources = @()

    if ($Id -like 'selftest.*') {
        $impactTags = @('test.selftest', "test.$($Id.Split('.')[1])")
        $inputSets = @('source.app', "source.$($Id.Split('.')[1])")
        $artifactIds = @('artifact.redsalamander-runtime')
        $artifactReads = @($artifactIds)
        $cachePolicy = 'ExactArtifact'
        $coverageFamily = $Id.Split('.')[1]
        $order = 'broad-same-process'
        $resources += 'stateful'
        if ($Id -in @('selftest.compare-directories', 'selftest.file-operations')) {
            $environment += @('env.network', 'env.test-resources')
            $resources += 'network'
        }
        $commandsFamilyArgument = @($Arguments | Where-Object { $_ -like '--selftest-family=*' } | Select-Object -First 1)
        if ($Id -eq 'selftest.commands' -and $commandsFamilyArgument.Count -eq 1) {
            $commandsFamily = $commandsFamilyArgument[0].Substring('--selftest-family='.Length)
            Assert-RSValidationIdentifier -Id $commandsFamily
            $coverageFamily = "commands.$commandsFamily"
            $order = 'family-process'
            # Commands families still consume the monolithic RedSalamander image. They are
            # focused iteration coverage, not independently reusable Full evidence.
            $cachePolicy = 'Never'
        }
    } elseif ($Id -like 'dxui.*') {
        $impactTags = @('test.dxui')
        $inputSets = @('source.dxui')
        $artifactIds = @('artifact.dxui-tests-runtime')
        $artifactReads = @($artifactIds)
        $cachePolicy = 'ExactArtifact'
        $coverageFamily = if ($Id -eq 'dxui.all') { $null } else { $Id.Substring(5) }
        $order = if ($Id -eq 'dxui.all') { 'broad-same-process' } else { 'family-process' }
    } elseif ($Id -like 'viewer-pe.*') {
        $impactTags = @('test.viewer-pe')
        $inputSets = @('source.viewer-pe')
        $artifactIds = @('artifact.viewer-pe-tests-runtime')
        $artifactReads = @($artifactIds)
        $cachePolicy = 'ExactArtifact'
        if ($Id -ne 'viewer-pe.all') { $exactCases = @($Arguments | Where-Object { $_ -notmatch '^-' }) }
    } elseif ($Id -like 'viewer-sqlite.*') {
        $impactTags = @('test.viewer-sqlite')
        $inputSets = @('source.viewer-sqlite')
        $artifactIds = @('artifact.viewer-sqlite-tests-runtime')
        $artifactReads = @($artifactIds)
        $cachePolicy = 'ExactArtifact'
    } elseif ($Id -eq 'tooling.pester-build-toolchain') {
        $impactTags = @('test.tooling', 'test.build')
        $inputSets = @('source.tooling', 'source.build-graph')
        $artifactIds = @('artifact.profile')
        $artifactReads = @('artifact.profile')
        $artifactWrites = @('artifact.profile')
        $produces = @('artifact.profile')
        $invalidates = @('receipt.profile')
        # Fresh execution opens a verified post-build artifact epoch. Exact Resume
        # may reuse that green result only while the complete post-entry receipt
        # remains byte-identical, avoiding another mutation of a matching profile.
        $cachePolicy = 'ExactArtifact'
        $resources += 'artifact-mutating'
    } elseif ($Id -eq 'tooling.pester') {
        $impactTags = @('test.tooling')
        $inputSets = @('source.tooling')
    }

    if ($Id -like 'perf.*') {
        $cachePolicy = 'Never'
        $impactTags += 'test.performance'
        $resources += @('performance-exclusive', 'cpu-heavy')
    } elseif ($Kind -in @('CppUnitTest', 'Pester')) {
        $resources += 'cpu-heavy'
    }
    if ($Id -like '*curl*') {
        $cachePolicy = 'Never'
        $environment += 'env.network'
        $resources += 'network'
    }
    if ($RequiresInteractiveDesktop) {
        $environment += 'env.interactive-desktop'
        $resources += 'interactive-desktop'
    }
    if ($artifactWrites.Count -eq 0) { $resources += 'artifact-read-only' }

    $repeat = 1L
    $repeatArgument = @($Arguments | Where-Object { $_ -match '^--selftest-repeat=([0-9]+)$' } | Select-Object -First 1)
    if ($repeatArgument.Count -eq 1) { $repeat = [int64]($repeatArgument[0] -replace '^--selftest-repeat=', '') }
    $shuffleArgument = @($Arguments | Where-Object { $_ -like '--selftest-shuffle=*' } | Select-Object -First 1)
    $orderIdentity = if ($shuffleArgument.Count -eq 1) { "$Id.$($shuffleArgument[0].Substring(19))" } else { "$Id.default" }

    [pscustomobject]@{
        ImpactTags = Get-RSTestPlanSortedIds $impactTags
        InputSets = Get-RSTestPlanSortedIds $inputSets
        ArtifactIds = Get-RSTestPlanSortedIds $artifactIds
        ArtifactReadIds = Get-RSTestPlanSortedIds $artifactReads
        ArtifactWriteIds = Get-RSTestPlanSortedIds $artifactWrites
        ProducesArtifactIds = Get-RSTestPlanSortedIds $produces
        InvalidatesReceiptScopes = Get-RSTestPlanSortedIds $invalidates
        CachePolicy = $cachePolicy
        EnvironmentInputs = Get-RSTestPlanSortedIds $environment
        CoverageContract = [pscustomobject][ordered]@{
            suite = $Id
            family = $coverageFamily
            exact_cases = @($exactCases | Sort-Object -Unique)
            repeat = $repeat
            order_identity = $orderIdentity
        }
        OutcomeContract = [pscustomobject][ordered]@{
            pass_required_cases = @('selected')
            static_skip_cases = @()
            capability_bound_skip_cases = @('declared-capability')
            forbidden_skip_classes = @('unclassified', 'unexpected')
        }
        OrderContract = $order
        ResourceClasses = @($resources | Sort-Object -Unique)
        EstimatedDurationKey = "duration.$Id"
    }
}

function New-RSTestRunPlanEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Id,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [ValidateSet('SelfTest', 'Executable', 'CppUnitTest', 'Pester', 'PowerShellScript')]
        [string]$Kind,

        [Parameter(Mandatory = $true)]
        [string]$Path,

        [string[]]$Arguments = @(),

        [string]$WorkingDirectory = '',

        [string]$JsonName = '',

        [switch]$RequiresInteractiveDesktop
    )

    Assert-RSValidationIdentifier -Id $Id
    $policy = Get-RSTestRunPlanEntryPolicy -Id $Id -Kind $Kind -Arguments $Arguments `
        -RequiresInteractiveDesktop ([bool]$RequiresInteractiveDesktop)

    [pscustomobject]@{
        Id = $Id
        Name = $Name
        Kind = $Kind
        Path = $Path
        Arguments = @($Arguments)
        WorkingDirectory = $WorkingDirectory
        JsonName = $JsonName
        RequiresInteractiveDesktop = [bool]$RequiresInteractiveDesktop
        ImpactTags = @($policy.ImpactTags)
        InputSets = @($policy.InputSets)
        ArtifactIds = @($policy.ArtifactIds)
        ArtifactReadIds = @($policy.ArtifactReadIds)
        ArtifactWriteIds = @($policy.ArtifactWriteIds)
        ProducesArtifactIds = @($policy.ProducesArtifactIds)
        InvalidatesReceiptScopes = @($policy.InvalidatesReceiptScopes)
        CachePolicy = $policy.CachePolicy
        EnvironmentInputs = @($policy.EnvironmentInputs)
        CoverageContract = $policy.CoverageContract
        OutcomeContract = $policy.OutcomeContract
        OrderContract = $policy.OrderContract
        ResourceClasses = @($policy.ResourceClasses)
        EstimatedDurationKey = $policy.EstimatedDurationKey
    }
}

function Assert-RSTestRunPlanEntries {
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Entries)

    $ids = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($entry in $Entries) {
        $id = [string]$entry.Id
        Assert-RSValidationIdentifier -Id $id
        if (-not $ids.Add($id)) { throw "Duplicate test-plan entry ID '$id'." }
        foreach ($property in @(
                'ImpactTags', 'InputSets', 'ArtifactIds', 'ArtifactReadIds', 'ArtifactWriteIds',
                'ProducesArtifactIds', 'InvalidatesReceiptScopes', 'CachePolicy', 'EnvironmentInputs',
                'CoverageContract', 'OutcomeContract', 'OrderContract', 'ResourceClasses', 'EstimatedDurationKey')) {
            if ($null -eq $entry.PSObject.Properties[$property]) {
                throw "Test-plan entry '$id' is missing metadata '$property'."
            }
        }
    }
    return $true
}

function ConvertTo-RSTestPlanPortablePath {
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [AllowEmptyString()][string]$Path
    )

    if ([string]::IsNullOrEmpty($Path)) { return '.' }
    $root = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\')
    $full = [IO.Path]::GetFullPath($Path)
    if ($full.Equals($root, [StringComparison]::OrdinalIgnoreCase)) { return '.' }
    if (-not $full.StartsWith($root + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw "Validation plan path must remain beneath the repository: $full"
    }
    return [IO.Path]::GetRelativePath($root, $full).Replace('\', '/')
}

function ConvertTo-RSValidationPlanEntry {
    param(
        [Parameter(Mandatory = $true)][object]$Entry,
        [Parameter(Mandatory = $true)][string]$RepoRoot
    )

    [pscustomobject][ordered]@{
        id = $Entry.Id
        display_name = $Entry.Name
        kind = $Entry.Kind
        path = ConvertTo-RSTestPlanPortablePath -RepoRoot $RepoRoot -Path $Entry.Path
        arguments = @($Entry.Arguments)
        working_directory = ConvertTo-RSTestPlanPortablePath -RepoRoot $RepoRoot -Path $Entry.WorkingDirectory
        result_contract_id = if ([string]::IsNullOrWhiteSpace([string]$Entry.JsonName)) { 'result.process-exit' } else { "result.$($Entry.JsonName)" }
        impact_tags = @($Entry.ImpactTags)
        input_sets = @($Entry.InputSets)
        artifact_ids = @($Entry.ArtifactIds)
        artifact_read_ids = @($Entry.ArtifactReadIds)
        artifact_write_ids = @($Entry.ArtifactWriteIds)
        produces_artifact_ids = @($Entry.ProducesArtifactIds)
        invalidates_receipt_scopes = @($Entry.InvalidatesReceiptScopes)
        cache_policy = $Entry.CachePolicy
        environment_inputs = @($Entry.EnvironmentInputs)
        coverage_contract = $Entry.CoverageContract
        outcome_contract = $Entry.OutcomeContract
        order_contract = $Entry.OrderContract
        resource_classes = @($Entry.ResourceClasses)
        estimated_duration_key = $Entry.EstimatedDurationKey
        requires_interactive_desktop = [bool]$Entry.RequiresInteractiveDesktop
    }
}

function New-RSValidationPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string]$RequestedSuite,
        [Parameter(Mandatory = $true)][object[]]$Entries,
        [ValidateSet('Fresh', 'Resume', 'Affected')][string]$ValidationMode = 'Fresh',
        [ValidateSet('complete', 'filtered', 'family', 'capability-reduced')][string]$CoverageScope = 'complete'
    )

    [void](Assert-RSTestRunPlanEntries -Entries $Entries)
    $portableEntries = @($Entries | ForEach-Object { ConvertTo-RSValidationPlanEntry -Entry $_ -RepoRoot $RepoRoot })
    $plan = [ordered]@{
        schema = 'red-salamander.validation-plan.v1'
        canonicalization = 'rs-canonical-json-v1'
        plan_digest = ''
        coverage_plan_digest = ''
        requested_suite = $RequestedSuite
        validation_mode = $ValidationMode
        coverage_scope = $CoverageScope
        entries = $portableEntries
    }
    $digestEntries = @($portableEntries | ForEach-Object {
            $projection = [ordered]@{}
            foreach ($property in $_.PSObject.Properties) {
                if ($property.Name -ne 'display_name') { $projection[$property.Name] = $property.Value }
            }
            [pscustomobject]$projection
        })
    $plan.coverage_plan_digest = Get-RSCanonicalJsonDigest -InputObject ([ordered]@{
            schema = 'red-salamander.validation-coverage-plan.v1'
            canonicalization = $plan.canonicalization
            requested_suite = $plan.requested_suite
            coverage_scope = $plan.coverage_scope
            entries = $digestEntries
        })
    $plan.plan_digest = Get-RSCanonicalJsonDigest -InputObject ([ordered]@{
            schema = $plan.schema
            canonicalization = $plan.canonicalization
            coverage_plan_digest = $plan.coverage_plan_digest
            validation_mode = $plan.validation_mode
        })
    Assert-RSValidationPlanSemantics -Plan ([pscustomobject]$plan)
    return [pscustomobject]$plan
}

function Get-RSBuildOutputPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$Platform,

        [Parameter(Mandatory = $true)]
        [string]$Configuration,

        [Parameter(Mandatory = $true)]
        [string]$FileName
    )

    return (Join-Path $RepoRoot (".build\{0}\{1}\{2}" -f $Platform, $Configuration, $FileName))
}

function ConvertTo-RSFullPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return [System.IO.Path]::GetFullPath($Path)
}

function New-RSTestRunId {
    $timestamp = (Get-Date).ToUniversalTime().ToString('yyyyMMddTHHmmssZ')
    $guid = [guid]::NewGuid().ToString('N')
    return "$timestamp-$PID-$guid"
}

function ConvertFrom-RSTestRunId {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RunId
    )

    $match = [regex]::Match($RunId, '^(?<timestamp>\d{8}T\d{6}Z)-(?<pid>\d+)-(?<guid>[0-9a-fA-F]{32})$')
    $format = 'Runner'
    if (-not $match.Success) {
        # Standalone native harnesses use TestSupport's crash-tolerant fallback
        # <harness>-<pid>-<tick> form when no runner-owned id is supplied.
        $match = [regex]::Match($RunId, '^(?<prefix>[A-Za-z0-9][A-Za-z0-9._-]*?)-(?<pid>\d+)-(?<tick>\d+)$')
        $format = 'HarnessFallback'
    }
    if (-not $match.Success) {
        return $null
    }

    $processId = 0L
    if (-not [int64]::TryParse($match.Groups['pid'].Value, [ref]$processId)) {
        return $null
    }

    [pscustomobject]@{
        RunId = $RunId
        Format = $format
        Timestamp = $match.Groups['timestamp'].Value
        ProcessId = $processId
        Guid = $match.Groups['guid'].Value
        Prefix = $match.Groups['prefix'].Value
        TickCount = $match.Groups['tick'].Value
    }
}

function Get-RSLiveProcessSnapshots {
    $processes = @([System.Diagnostics.Process]::GetProcesses())
    $snapshots = @()
    try {
        $snapshots = @($processes | ForEach-Object {
                $startTimeUtc = $null
                try {
                    $startTimeUtc = $_.StartTime.ToUniversalTime()
                } catch {
                    # Access to protected process metadata can be denied. A null
                    # start time keeps the conservative live-PID behavior.
                }

                [pscustomobject]@{
                    ProcessId = [int64]$_.Id
                    StartTimeUtc = $startTimeUtc
                }
            })
    } finally {
        foreach ($process in $processes) {
            $process.Dispose()
        }
    }

    $missingStartIds = @($snapshots | Where-Object { $null -eq $_.StartTimeUtc } | ForEach-Object { [int64]$_.ProcessId })
    if ($missingStartIds.Count -gt 0) {
        try {
            foreach ($process in @(Get-CimInstance Win32_Process -Property ProcessId, CreationDate -ErrorAction Stop)) {
                if ($null -eq $process.CreationDate -or [int64]$process.ProcessId -notin $missingStartIds) {
                    continue
                }

                $snapshot = $snapshots | Where-Object { $_.ProcessId -eq [int64]$process.ProcessId } | Select-Object -First 1
                if ($null -ne $snapshot) {
                    $snapshot.StartTimeUtc = ([datetime]$process.CreationDate).ToUniversalTime()
                }
            }
        } catch {
            # Keep null start times conservative when process metadata is unavailable.
        }
    }

    return $snapshots
}

Export-ModuleMember -Function @(
    'New-RSTestRunPlanEntry',
    'Assert-RSTestRunPlanEntries',
    'ConvertTo-RSValidationPlanEntry',
    'New-RSValidationPlan',
    'Get-RSBuildOutputPath',
    'ConvertTo-RSFullPath',
    'New-RSTestRunId',
    'ConvertFrom-RSTestRunId',
    'Get-RSLiveProcessSnapshots'
)
