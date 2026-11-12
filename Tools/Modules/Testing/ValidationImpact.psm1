Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ValidationFingerprint.psm1') -Scope Local -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot '../Build/BuildProjectSelection.psm1') -Scope Local -ErrorAction Stop

# Cache only immutable pattern translations. Paths, source bytes and validation outcomes
# are always evaluated again. Miss insertion is serialized; hits and matches are thread-safe.
$script:RSImpactCache = [Collections.Concurrent.ConcurrentDictionary[string, regex]]::new([StringComparer]::Ordinal)
$script:RSImpactGate = [object]::new()
function Test-RSValidationImpactPattern {
    param([Parameter(Mandatory = $true)][string]$Pattern, [Parameter(Mandatory = $true)][string]$Path)
    $key = $Pattern
    $matcher = $null
    if (-not $script:RSImpactCache.TryGetValue($key, [ref]$matcher)) {
        [Threading.Monitor]::Enter($script:RSImpactGate)
        try {
            if (-not $script:RSImpactCache.TryGetValue($key, [ref]$matcher)) {
                $builder = [Text.StringBuilder]::new('^')
                for ($index = 0; $index -lt $Pattern.Length; $index++) {
                    $character = $Pattern[$index]
                    if ($character -eq '*' -and $index + 1 -lt $Pattern.Length -and $Pattern[$index + 1] -eq '*') {
                        $index++
                        if ($index + 1 -lt $Pattern.Length -and $Pattern[$index + 1] -eq '/') {
                            $index++
                            [void]$builder.Append('(?:.*/)?')
                        } else {
                            [void]$builder.Append('.*')
                        }
                    } elseif ($character -eq '*') {
                        [void]$builder.Append('[^/]*')
                    } elseif ($character -eq '?') {
                        [void]$builder.Append('[^/]')
                    } else {
                        [void]$builder.Append([regex]::Escape([string]$character))
                    }
                }
                [void]$builder.Append('$')
                $expression = $builder.ToString()
                $matcher = [regex]::new($expression, ([Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::CultureInvariant))
                # Eviction changes cost only. Captured regex instances remain valid for active readers.
                if ($script:RSImpactCache.Count -ge 1024) { $script:RSImpactCache.Clear() }
                [void]$script:RSImpactCache.TryAdd($key, $matcher)
            }
        } finally { [Threading.Monitor]::Exit($script:RSImpactGate) }
    }
    return $matcher.IsMatch($Path)
}

function Read-RSValidationImpactManifest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$SchemaPath,
        [Parameter(Mandatory = $true)][string]$RepoRoot
    )

    $manifest = Read-RSValidationJsonFile -LiteralPath $Path -SchemaPath $SchemaPath -AllowedRoot $RepoRoot
    $ids = @($manifest.rules | ForEach-Object { [string]$_.id })
    if (@($ids | Sort-Object -Unique).Count -ne $ids.Count) { throw 'Validation impact rule IDs must be unique.' }
    foreach ($rule in $manifest.rules) {
        Assert-RSValidationIdentifier -Id ([string]$rule.id)
        Assert-RSValidationPathSet -Path @([string]$rule.pattern)
        if (-not [string]::IsNullOrEmpty([string]$rule.derived_closure_source)) {
            Assert-RSValidationPathSet -Path @([string]$rule.derived_closure_source)
        }
    }
    return $manifest
}

function Get-RSValidationImpactRulesForPath {
    param([Parameter(Mandatory = $true)][object]$Manifest, [Parameter(Mandatory = $true)][string]$Path)

    Assert-RSValidationPathSet -Path @($Path)
    $matches = @($Manifest.rules | Where-Object { Test-RSValidationImpactPattern -Pattern $_.pattern -Path $Path })
    $specific = @($matches | Where-Object { -not [bool]$_.broad_fallback })
    if ($specific.Count -gt 0) { return $specific }
    $fallback = @($matches | Where-Object { [bool]$_.broad_fallback })
    if ($fallback.Count -eq 0) { throw "No validation impact fallback covers '$Path'." }
    # Fallback precedence is reviewed manifest order. This lets a scoped
    # conservative fallback (for example executable Specs data) win before the
    # repository-wide unknown-path fallback without suppressing specific rules.
    return @($fallback | Select-Object -First 1)
}

function Test-RSValidationImpactGovernance {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Manifest,
        [Parameter(Mandatory = $true)][string[]]$TrackedPaths,
        [Parameter(Mandatory = $true)][string]$RepoRoot
    )

    $unmapped = @()
    $solutionPath = Join-Path $RepoRoot 'RedSalamander.sln'
    $solutionText = if (Test-Path -LiteralPath $solutionPath -PathType Leaf) {
        Get-Content -LiteralPath $solutionPath -Raw
    } else { '' }
    $graphPaths = @($TrackedPaths | Where-Object { $_ -match '(?i)(\.vcxproj$|\.props$|\.targets$)' })
    $graphText = @{}
    foreach ($graphPath in $graphPaths) {
        $fullPath = Join-Path $RepoRoot $graphPath
        $graphText[$graphPath] = if (Test-Path -LiteralPath $fullPath -PathType Leaf) {
            Get-Content -LiteralPath $fullPath -Raw
        } else { '' }
    }
    foreach ($path in @($TrackedPaths | Sort-Object -Unique)) {
        $normalized = $path.Replace('\', '/')
        if ($normalized -notmatch '(?i)(\.sln$|\.slnx$|\.vcxproj$|\.props$|\.targets$|^RuntimeDependencies\.props$)') { continue }
        $specificBuildRule = @($Manifest.rules | Where-Object {
                -not [bool]$_.broad_fallback -and $_.classification -eq 'build' -and
                (Test-RSValidationImpactPattern -Pattern $_.pattern -Path $normalized)
            })
        $mapped = ($specificBuildRule.Count -gt 0)
        if ($mapped -and $normalized.EndsWith('.vcxproj', [StringComparison]::OrdinalIgnoreCase)) {
            $solutionSpelling = $normalized.Replace('/', '\')
            $exactReviewed = @($specificBuildRule | Where-Object {
                    $_.pattern -notmatch '[*?]' -and $_.pattern -eq $normalized -and
                    $_.derived_closure_source -eq $normalized
                }).Count -gt 0
            $mapped = $solutionText.Contains($solutionSpelling, [StringComparison]::OrdinalIgnoreCase) -or $exactReviewed
        } elseif ($mapped -and $normalized -match '(?i)(\.props$|\.targets$)' -and
                  (Split-Path $normalized -Leaf) -notlike 'Directory.Build.*') {
            $leaf = Split-Path $normalized -Leaf
            $referenced = $false
            foreach ($candidate in $graphPaths) {
                if ($candidate -eq $path) { continue }
                if ($graphText[$candidate].Contains($leaf, [StringComparison]::OrdinalIgnoreCase)) {
                    $referenced = $true
                    break
                }
            }
            $mapped = $referenced
        }
        if (-not $mapped) { $unmapped += $normalized }
    }
    return [pscustomobject]@{ IsValid = ($unmapped.Count -eq 0); UnmappedBuildInputs = @($unmapped) }
}

function Get-RSValidationImpactDecision {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$Manifest,
        [Parameter(Mandatory = $true)][object]$WorkspaceSnapshot,
        [Parameter(Mandatory = $true)][object]$ValidationPlan,
        [hashtable]$EntryFingerprints = @{},
        [hashtable]$EstimatedDurationMs = @{},
        [ValidateSet('shadow', 'enforced')][string]$PlannerMode = 'shadow'
    )

    $changed = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($file in @($WorkspaceSnapshot.files)) {
        [void]$changed.Add([string]$file.path)
        if (-not [string]::IsNullOrEmpty([string]$file.old_path)) { [void]$changed.Add([string]$file.old_path) }
    }
    $changedPaths = [string[]]@($changed)
    [Array]::Sort($changedPaths, [StringComparer]::Ordinal)
    $matchedRules = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $impactTags = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $buildRank = @{ none = 0; targeted = 1; full = 2; blocked = 3 }
    $buildAction = 'none'
    foreach ($path in $changedPaths) {
        foreach ($rule in @(Get-RSValidationImpactRulesForPath -Manifest $Manifest -Path $path)) {
            [void]$matchedRules.Add([string]$rule.id)
            foreach ($tag in @($rule.impact_tags)) { [void]$impactTags.Add([string]$tag) }
            if ($buildRank[[string]$rule.build_requirement] -gt $buildRank[$buildAction]) {
                $buildAction = [string]$rule.build_requirement
            }
        }
    }
    $entries = @()
    for ($ordinal = 0; $ordinal -lt @($ValidationPlan.entries).Count; $ordinal++) {
        $entry = $ValidationPlan.entries[$ordinal]
        $matchedEntryTags = @($entry.impact_tags | Where-Object { $impactTags.Contains([string]$_) } | Sort-Object -Unique)
        $affected = ($matchedEntryTags.Count -gt 0)
        $expected = if ($EntryFingerprints.ContainsKey([string]$entry.id)) {
            [string]$EntryFingerprints[[string]$entry.id]
        } else {
            Get-RSCanonicalJsonDigest -InputObject ([ordered]@{
                    phase = 'impact-shadow.v1'; plan_digest = $ValidationPlan.plan_digest
                    snapshot_id = $WorkspaceSnapshot.snapshot_id; entry = $entry
                })
        }
        $entries += [pscustomobject][ordered]@{
            id = $entry.id
            ordinal = [int64]$ordinal
            affected = $affected
            matched_impact_tags = $matchedEntryTags
            decision = 'EXECUTE'
            reason_codes = @($(if ($affected) { 'SOURCE_INPUT_CHANGED' } else { 'FORCED' }))
            expected_fingerprint = $expected
            origin_run_id = $null
            origin_evidence_digest = $null
            estimated_duration_ms = if ($EstimatedDurationMs.ContainsKey([string]$entry.id)) { [int64]$EstimatedDurationMs[[string]$entry.id] } else { [int64]0 }
        }
    }
    $ruleIds = [string[]]@($matchedRules)
    [Array]::Sort($ruleIds, [StringComparer]::Ordinal)
    $freshDurationMeasure = $entries | Measure-Object estimated_duration_ms -Sum
    $executedDurationMeasure = $entries | Where-Object decision -eq EXECUTE | Measure-Object estimated_duration_ms -Sum
    $freshDuration = if ($null -eq $freshDurationMeasure -or $null -eq $freshDurationMeasure.Sum) { 0L } else { [int64]$freshDurationMeasure.Sum }
    $executedDuration = if ($null -eq $executedDurationMeasure -or $null -eq $executedDurationMeasure.Sum) { 0L } else { [int64]$executedDurationMeasure.Sum }
    $record = [ordered]@{
        schema = 'red-salamander.validation-plan-decision.v1'
        canonicalization = 'rs-canonical-json-v1'
        decision_digest = '0' * 64
        planner_mode = $PlannerMode
        snapshot_id = $WorkspaceSnapshot.snapshot_id
        plan_digest = $ValidationPlan.plan_digest
        changed_paths = $changedPaths
        matched_impact_rules = $ruleIds
        advisory_build_action = $buildAction
        entries = $entries
        estimated_fresh_duration_ms = $freshDuration
        estimated_executed_duration_ms = $executedDuration
        estimated_avoided_duration_ms = [int64]0
    }
    $projection = [ordered]@{}
    foreach ($property in $record.Keys) { if ($property -ne 'decision_digest') { $projection[$property] = $record[$property] } }
    $record.decision_digest = Get-RSCanonicalJsonDigest -InputObject $projection
    return [pscustomobject]$record
}

function Get-RSValidationAffectedBuildSelection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Decision,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$ExecutionEntries,
        [Parameter(Mandatory)][string]$SolutionPath
    )

    $result = [pscustomobject]@{
        UseSelectedBuild = $false
        Reason = 'The governed impact decision requires the existing full build.'
        Selection = $null
        EntryIds = @()
    }
    if ($Decision.advisory_build_action -cne 'targeted') { return $result }
    $known = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($entry in $ExecutionEntries) {
        if (-not $known.Add([string]$entry.Id)) { throw 'Affected build entries have duplicate identities.' }
    }
    $decisionIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $selectedIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($entry in $Decision.entries) {
        if (-not $decisionIds.Add([string]$entry.id)) { throw 'Affected build decision has duplicate identities.' }
        if ($entry.affected) { [void]$selectedIds.Add([string]$entry.id) }
    }
    if (-not $known.SetEquals($decisionIds)) { throw 'Affected build decision does not describe the requested plan.' }
    $names = [Collections.Generic.List[string]]::new()
    foreach ($entry in $ExecutionEntries) {
        if (-not $selectedIds.Contains([string]$entry.Id)) { continue }
        if (@($entry.ArtifactWriteIds).Count) {
            $result.Reason = 'An affected artifact writer requires the full profile and re-attestation barrier.'
            return $result
        }
        if ($entry.Kind -in @('SelfTest', 'Executable', 'CppUnitTest')) {
            if ($null -eq $entry.PSObject.Properties['BuildProjectNames'] -or -not @($entry.BuildProjectNames).Count) {
                $result.Reason = "Native entry '$($entry.Id)' has no reviewed build-project binding."
                return $result
            }
            foreach ($name in $entry.BuildProjectNames) { $names.Add([string]$name) }
        } elseif (@($entry.ArtifactReadIds).Count) {
            $result.Reason = "Script entry '$($entry.Id)' reads native artifacts outside an explicit build-project binding."
            return $result
        }
    }
    if (-not $names.Count) {
        $result.Reason = 'No affected native project union is available; retain the existing build/receipt gate.'
        return $result
    }
    try {
        $result.Selection = Get-RSBuildProjectUnion -SolutionPath $SolutionPath -ProjectNames $names.ToArray()
    } catch {
        $result.Reason = "The affected project binding could not be resolved; retain the full build. $($_.Exception.Message)"
        return $result
    }
    $result.UseSelectedBuild = $true
    $result.Reason = 'Every affected native entry has an explicit project binding; evaluate its complete solution/runtime union.'
    $result.EntryIds = @($ExecutionEntries | Where-Object { $selectedIds.Contains([string]$_.Id) } | ForEach-Object { $_.Id })
    return $result
}

function Get-RSValidationModeGateMatrix {
    [CmdletBinding()]
    param()

    return @(
        [pscustomobject][ordered]@{ mode = 'Fresh'; public = $true; executes_all = $true; permits_reuse = $false; repository_verdict = 'evaluated-when-complete'; activation_phase = 0L },
        [pscustomobject][ordered]@{ mode = 'ExplainPlan'; public = $true; executes_all = $false; permits_reuse = $false; repository_verdict = 'not-evaluated'; activation_phase = 4L },
        [pscustomobject][ordered]@{ mode = 'Resume'; public = $true; executes_all = $false; permits_reuse = $true; repository_verdict = 'evaluated-when-complete'; activation_phase = 6L },
        [pscustomobject][ordered]@{ mode = 'Affected'; public = $true; executes_all = $false; permits_reuse = $false; repository_verdict = 'not-evaluated'; activation_phase = 9L },
        [pscustomobject][ordered]@{ mode = 'CI'; public = $true; executes_all = $true; permits_reuse = $false; repository_verdict = 'evaluated-when-complete'; activation_phase = 0L },
        [pscustomobject][ordered]@{ mode = 'Release'; public = $true; executes_all = $true; permits_reuse = $false; repository_verdict = 'evaluated-when-complete'; activation_phase = 0L }
    )
}

Export-ModuleMember -Function @(
    'Read-RSValidationImpactManifest',
    'Get-RSValidationImpactRulesForPath',
    'Test-RSValidationImpactGovernance',
    'Get-RSValidationImpactDecision',
    'Get-RSValidationAffectedBuildSelection',
    'Get-RSValidationModeGateMatrix'
)
