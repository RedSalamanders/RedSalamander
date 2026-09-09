Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ValidationFingerprint.psm1') -Scope Local -ErrorAction Stop

function Test-RSValidationImpactPattern {
    param([Parameter(Mandatory = $true)][string]$Pattern, [Parameter(Mandatory = $true)][string]$Path)
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
    return [regex]::IsMatch($Path, $builder.ToString(),
        [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::CultureInvariant)
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
    'Get-RSValidationModeGateMatrix'
)
