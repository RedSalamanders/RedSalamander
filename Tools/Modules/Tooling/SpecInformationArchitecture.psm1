# Private tooling-governance module; import through its reviewed callers.
Set-StrictMode -Version Latest

function ConvertTo-RSForwardSlashPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    return $Path.Replace('\', '/')
}

function Get-RSRepositoryRelativePath {
    param(
        [Parameter(Mandatory = $true)][string]$RepositoryRoot,
        [Parameter(Mandatory = $true)][string]$Path
    )

    $root = [IO.Path]::GetFullPath($RepositoryRoot).TrimEnd('\', '/')
    $full = [IO.Path]::GetFullPath($Path)
    $prefix = $root + [IO.Path]::DirectorySeparatorChar
    if (-not $full.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Path escapes the repository root: $Path"
    }

    return ConvertTo-RSForwardSlashPath -Path $full.Substring($prefix.Length)
}

function Get-RSGitOutput {
    param(
        [Parameter(Mandatory = $true)][string]$RepositoryRoot,
        [Parameter(Mandatory = $true)][string[]]$Arguments,
        [switch]$AllowFailure
    )

    $output = @(& git -C $RepositoryRoot @Arguments 2>$null)
    $exitCode = $LASTEXITCODE
    if ($exitCode -ne 0 -and -not $AllowFailure) {
        throw "git $($Arguments -join ' ') failed with exit code $exitCode."
    }

    return [pscustomobject]@{
        ExitCode = $exitCode
        Output = @($output)
    }
}

function Get-RSInventoryPaths {
    param([Parameter(Mandatory = $true)][string]$RepositoryRoot)

    $result = Get-RSGitOutput -RepositoryRoot $RepositoryRoot -Arguments @(
        'ls-files', '--cached', '--others', '--exclude-standard'
    )

    return @($result.Output |
        ForEach-Object { ConvertTo-RSForwardSlashPath -Path ([string]$_) } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique)
}

function Get-RSMarkdownCodeMask {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Text)

    $lines = [regex]::Split($Text, '(?<=\n)')
    $maskedLines = [System.Collections.Generic.List[string]]::new()
    $inFence = $false
    $fenceCharacter = [char]0
    $fenceLength = 0

    foreach ($line in $lines) {
        $content = $line.TrimEnd("`r", "`n")
        $ending = $line.Substring($content.Length)
        $fenceMatch = [regex]::Match($content, '^\s*(?<fence>`{3,}|~{3,})')
        if ($fenceMatch.Success) {
            $run = $fenceMatch.Groups['fence'].Value
            $character = $run[0]
            if (-not $inFence) {
                $inFence = $true
                $fenceCharacter = $character
                $fenceLength = $run.Length
            } elseif ($character -eq $fenceCharacter -and $run.Length -ge $fenceLength) {
                $inFence = $false
                $fenceCharacter = [char]0
                $fenceLength = 0
            }
            $maskedLines.Add((' ' * $content.Length) + $ending)
            continue
        }

        if ($inFence) {
            $maskedLines.Add((' ' * $content.Length) + $ending)
            continue
        }

        $characters = $content.ToCharArray()
        $index = 0
        while ($index -lt $characters.Length) {
            if ($characters[$index] -ne '`') {
                $index++
                continue
            }

            $delimiterStart = $index
            while ($index -lt $characters.Length -and $characters[$index] -eq '`') {
                $index++
            }
            $delimiterLength = $index - $delimiterStart
            $closingStart = -1
            $search = $index
            while ($search -lt $characters.Length) {
                if ($characters[$search] -ne '`') {
                    $search++
                    continue
                }
                $runStart = $search
                while ($search -lt $characters.Length -and $characters[$search] -eq '`') {
                    $search++
                }
                if (($search - $runStart) -eq $delimiterLength) {
                    $closingStart = $runStart
                    break
                }
            }

            if ($closingStart -lt 0) {
                $index = $delimiterStart + $delimiterLength
                continue
            }

            $closingEnd = $closingStart + $delimiterLength
            for ($maskIndex = $delimiterStart; $maskIndex -lt $closingEnd; $maskIndex++) {
                $characters[$maskIndex] = ' '
            }
            $index = $closingEnd
        }

        $maskedLines.Add((-join $characters) + $ending)
    }

    return -join $maskedLines
}

function ConvertTo-RSGitHubHeadingSlug {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Heading)

    $value = $Heading.Trim().ToLowerInvariant()
    $value = [regex]::Replace($value, '<[^>]+>', '')
    $value = [regex]::Replace($value, '!\[([^\]]*)\]\([^)]*\)', '$1')
    $value = [regex]::Replace($value, '\[([^\]]+)\]\([^)]*\)', '$1')
    $value = $value.Replace('`', '').Replace('*', '').Replace('_', '')
    $value = [regex]::Replace($value, '[^\p{L}\p{Nd}\s-]', '')
    return [regex]::Replace($value, '\s+', '-')
}

function Get-RSMarkdownHeadings {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Text)

    $masked = Get-RSMarkdownCodeMask -Text $Text
    $originalLines = [regex]::Split($Text, '\r?\n')
    $maskedLines = [regex]::Split($masked, '\r?\n')
    $slugCounts = @{}
    $headings = [System.Collections.Generic.List[object]]::new()

    for ($index = 0; $index -lt $maskedLines.Count; $index++) {
        $match = [regex]::Match($maskedLines[$index], '^\s{0,3}(?<marks>#{1,6})\s+(?<title>.+?)\s*#*\s*$')
        if (-not $match.Success) { continue }

        $originalMatch = [regex]::Match($originalLines[$index], '^\s{0,3}(?<marks>#{1,6})\s+(?<title>.+?)\s*#*\s*$')
        if (-not $originalMatch.Success) { continue }
        $title = $originalMatch.Groups['title'].Value
        $baseSlug = ConvertTo-RSGitHubHeadingSlug -Heading $title
        if ($slugCounts.ContainsKey($baseSlug)) {
            $slugCounts[$baseSlug] = [int]$slugCounts[$baseSlug] + 1
            $slug = "$baseSlug-$($slugCounts[$baseSlug])"
        } else {
            $slugCounts[$baseSlug] = 0
            $slug = $baseSlug
        }

        $headings.Add([pscustomobject]@{
            Level = $originalMatch.Groups['marks'].Value.Length
            Title = $title
            Slug = $slug
            Line = $index + 1
        })
    }

    return @($headings)
}

function Get-RSMarkdownLinks {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Text)

    $masked = Get-RSMarkdownCodeMask -Text $Text
    $links = [System.Collections.Generic.List[object]]::new()
    $line = 1
    $index = 0
    while ($index -lt $masked.Length) {
        if ($masked[$index] -eq "`n") {
            $line++
            $index++
            continue
        }
        if ($masked[$index] -ne ']' -or ($index + 1) -ge $masked.Length -or $masked[$index + 1] -ne '(') {
            $index++
            continue
        }

        $labelStart = $index - 1
        while ($labelStart -ge 0 -and $masked[$labelStart] -ne '[' -and $masked[$labelStart] -ne "`n") {
            $labelStart--
        }
        if ($labelStart -lt 0 -or $masked[$labelStart] -ne '[' -or
            ($labelStart -gt 0 -and $masked[$labelStart - 1] -eq '!')) {
            $index += 2
            continue
        }

        $targetStart = $index + 2
        $cursor = $targetStart
        $depth = 1
        $escaped = $false
        while ($cursor -lt $masked.Length -and $depth -gt 0) {
            $character = $masked[$cursor]
            if ($escaped) {
                $escaped = $false
            } elseif ($character -eq '\') {
                $escaped = $true
            } elseif ($character -eq '(') {
                $depth++
            } elseif ($character -eq ')') {
                $depth--
            }
            $cursor++
        }
        if ($depth -ne 0) {
            $index += 2
            continue
        }

        $rawTarget = $Text.Substring($targetStart, ($cursor - 1) - $targetStart).Trim()
        if ($rawTarget.StartsWith('<') -and $rawTarget.Contains('>')) {
            $rawTarget = $rawTarget.Substring(1, $rawTarget.IndexOf('>') - 1)
        } elseif ($rawTarget -match '^(?<path>\S+)(?:\s+["''].*)?$') {
            $rawTarget = $matches['path']
        }
        $label = $Text.Substring($labelStart + 1, $index - $labelStart - 1)
        $links.Add([pscustomobject]@{
            Label = $label
            Target = $rawTarget
            Line = $line
        })
        $index = $cursor
    }

    return @($links)
}

function Get-RSSpecAuthorityClass {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$RelativePath)

    $path = ConvertTo-RSForwardSlashPath -Path $RelativePath
    if ($path -eq 'Specs/Plans/WIP/README.md') { return 'ActivePlanIndex' }
    if ($path -eq 'Specs/Plans/WIP/Notes_Scratch.md') { return 'HumanManagedNote' }
    if ($path -like 'Specs/Plans/WIP/*.md') { return 'ActivePlan' }
    if ($path -like 'Specs/Plans/Done/*' -or $path -like 'Specs/Reviews/*' -or $path -like 'plans/*') { return 'Historical' }
    if ($path -in @(
            'Specs/UI/UI_PreferencesDialog_MigrationHistory.md',
            'Specs/FileSystem/FileSystem_FtpScpSftpPerformanceComparison.md')) { return 'Historical' }
    if ($path -like 'Specs/TestRuns/*') { return 'Evidence' }
    if ($path -like 'Specs/Mockups/*') { return 'Prototype' }
    if ($path -eq 'Specs/Themes/THIRD-PARTY-NOTICES.md') { return 'Legal' }
    if ($path -like 'Specs/Themes/*') { return 'RuntimeData' }
    if ($path -like 'Specs/Build/*.schema.json' -or $path -like 'Specs/Testing/*.schema.json') { return 'MachineContract' }
    if ($path -like 'Specs/Terminal/*.json' -or $path -like 'Specs/Terminal/TerminalEngineEvidenceCorpus/*') { return 'HistoricalData' }
    if ($path -eq 'Specs/SettingsStore.schema.json' -or $path -eq 'Specs/Testing/FolderViewPerfBudgets.json5') { return 'MachineContract' }
    if ($path -eq 'Specs/README.md') { return 'NormativeCatalog' }
    if ($path -match '^Specs/[^/]+/[^/]+\.md$') { return 'Normative' }
    if ($path -like 'Specs/*') { return 'SpecificationArtifact' }
    if ($path -match '\.md$') { return 'ExternalDocumentation' }
    return 'Other'
}

function Get-RSSpecDomain {
    param([Parameter(Mandatory = $true)][string]$RelativePath)

    $path = ConvertTo-RSForwardSlashPath -Path $RelativePath
    if ($path -eq 'Specs/README.md') { return 'Root' }
    if ($path -match '^Specs/(?<domain>[^/]+)/') { return $matches['domain'] }
    return 'External'
}

function Get-RSMarkdownMarkers {
    param([Parameter(Mandatory = $true)][string]$Text)

    $masked = Get-RSMarkdownCodeMask -Text $Text
    return [pscustomobject]@{
        Checkboxes = ([regex]::Matches($masked, '(?m)^\s*[-*]\s+\[[ xX]\]')).Count
        UncheckedCheckboxes = ([regex]::Matches($masked, '(?m)^\s*[-*]\s+\[ \]')).Count
        Todo = ([regex]::Matches($masked, '(?im)\b(?:TODO|FIXME)\b')).Count
        FuturePlanHeadings = ([regex]::Matches(
                $masked,
                '(?im)^#{1,6}\s+.*\b(?:Future Enhancements?|Future Extensions?|Planned Features?|Implementation Plan|Proposed Plan|Roadmap|Remaining Work|Current Bugs\s*/\s*Follow-ups)\b'
            )).Count
        HistoryHeadings = ([regex]::Matches($masked, '(?im)^#{1,6}\s+.*\b(?:History|Migration History|Recent .* Improvements|Completed .* Actions)\b')).Count
    }
}

function Read-RSConsistencyLedger {
    param(
        [Parameter(Mandatory = $true)][string]$RepositoryRoot,
        [string]$LedgerPath
    )

    if ([string]::IsNullOrWhiteSpace($LedgerPath)) {
        $LedgerPath = 'Specs/NormativeConsistency.json'
    }
    $fullPath = if ([IO.Path]::IsPathRooted($LedgerPath)) { $LedgerPath } else { Join-Path $RepositoryRoot $LedgerPath }
    if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
        return [pscustomobject]@{
            Path = Get-RSRepositoryRelativePath -RepositoryRoot $RepositoryRoot -Path $fullPath
            Exists = $false
            Documents = @()
            DocumentMap = @{}
        }
    }

    $document = Get-Content -LiteralPath $fullPath -Raw | ConvertFrom-Json
    $map = @{}
    foreach ($entry in @($document.documents)) {
        $normalized = ConvertTo-RSForwardSlashPath -Path ([string]$entry.path)
        if ($map.ContainsKey($normalized)) {
            throw "Consistency ledger contains duplicate path '$normalized'."
        }
        $map[$normalized] = $entry
    }
    return [pscustomobject]@{
        Path = Get-RSRepositoryRelativePath -RepositoryRoot $RepositoryRoot -Path $fullPath
        Exists = $true
        SchemaVersion = $document.schemaVersion
        ReviewedAtCommit = $document.reviewedAtCommit
        Documents = @($document.documents)
        DocumentMap = $map
    }
}

function Get-RSWipIndexEntries {
    param([Parameter(Mandatory = $true)][string]$IndexText)

    $startMarker = '<!-- spec-ia:wip-index:start -->'
    $endMarker = '<!-- spec-ia:wip-index:end -->'
    $start = $IndexText.IndexOf($startMarker, [StringComparison]::Ordinal)
    $end = $IndexText.IndexOf($endMarker, [StringComparison]::Ordinal)
    if ($start -lt 0 -or $end -lt 0 -or $end -le $start) {
        return @()
    }

    $section = $IndexText.Substring($start + $startMarker.Length, $end - ($start + $startMarker.Length))
    $entries = [System.Collections.Generic.List[object]]::new()
    foreach ($line in [regex]::Split($section, '\r?\n')) {
        if ($line -notmatch '^\s*\|') { continue }
        $cells = @($line.Trim().Trim('|').Split('|') | ForEach-Object { $_.Trim() })
        if ($cells.Count -lt 3) { continue }
        $status = $cells[0].ToUpperInvariant()
        if ($status -notin @('ACTIVE', 'HOLD', 'DECISION', 'RETIRED')) { continue }
        $fileMatch = [regex]::Match($line, '(?:\]\(|`)(?<path>[^)`]+\.md)')
        if (-not $fileMatch.Success) { continue }
        $path = ConvertTo-RSForwardSlashPath -Path $fileMatch.Groups['path'].Value
        if ($path.StartsWith('./')) { $path = $path.Substring(2) }
        if (-not $path.StartsWith('Specs/')) { $path = "Specs/Plans/WIP/$path" }
        $entries.Add([pscustomobject]@{
            Status = $status
            Path = $path
            Line = $line
        })
    }
    return @($entries)
}

function Test-RSPathInsideRoot {
    param(
        [Parameter(Mandatory = $true)][string]$RepositoryRoot,
        [Parameter(Mandatory = $true)][string]$Candidate
    )

    $root = [IO.Path]::GetFullPath($RepositoryRoot).TrimEnd('\', '/')
    $full = [IO.Path]::GetFullPath($Candidate)
    return $full.Equals($root, [StringComparison]::OrdinalIgnoreCase) -or
        $full.StartsWith($root + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)
}

function Test-RSProtectedDoneHistory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RepositoryRoot,
        [string]$BaseCommit = '6aecfde8e',
        [string[]]$AllowedAdditionalPaths = @(
            'Specs/Plans/Done/Operation_Atlas_SpecificationAuthorityAndWipCleanup_2026-08-02.md',
            'Specs/Plans/Done/CodeReview_July21ToAugust4_IndependentFindingsAndRemediation_2026-08-04.md',
            'Specs/Plans/Done/CodeReview_August8_CorrectnessDataSafetyAndSimplification_2026-08-08.md',
            'Specs/Plans/Done/CodeReview_July21ToAugust4_ResidualDecisionsAndEvidence_2026-08-04.md',
            'Specs/Plans/Done/Operation_Emberglass_48Hour_RepositoryDeepReview_2026-07-22.md',
            'Specs/Plans/Done/Operation_Floodgate_CloudParityPerfProofFollowups_2026-06-18.md',
            'Specs/Plans/Done/Operation_Cinderstar_UnifiedRepositoryRiskRemediation_2026-08-09.md',
            'Specs/Plans/Done/Operation_Chordforge_TerminalCommandExperience_2026-08-11.md',
            'Specs/Plans/Done/Operation_Startrail_TrustedIncrementalValidationAndResumableFullEvidence_2026-08-13.md',
            'Specs/Plans/Done/CodeReview_CrossDomain_Remediation_2026-08-10.md',
            'Specs/Plans/Done/Operation_TerminalHostFileOpsDefectFixes_2026-08-19.md',
            'Specs/Plans/Done/Operation_ReviewFollowup_FOTerminalFocus_2026-08-19.md',
            'Specs/Plans/Done/Operation_FileOperations_GlobalBehaviorDecisionReview_2026-08-16.md',
            'Specs/Plans/Done/Operation_FolderView_FocusAndSelectionStateModel_2026-08-20.md',
            'Specs/Plans/Done/Operation_FolderView_FocusSelectionReviewRemediation_2026-08-21.md',
            'Specs/Plans/Done/AdvisorPlan_Index_2026-06-11.md',
            'Specs/Plans/Done/AdvisorPlan_001-test-all-script_2026-06-11.md',
            'Specs/Plans/Done/AdvisorPlan_002-vcpkg-dup-and-docs-refresh_2026-06-11.md',
            'Specs/Plans/Done/AdvisorPlan_003-secret-wipe-hygiene_2026-06-11.md',
            'Specs/Plans/Done/AdvisorPlan_004-plugin-capabilities-contract-test_2026-06-11.md',
            'Specs/Plans/Done/AdvisorPlan_005-plugin-factory-dedup_2026-06-11.md',
            'Specs/Plans/Done/AdvisorPlan_006-settings-schema-parser-tests_2026-06-11.md',
            'Specs/Plans/Done/AdvisorPlan_007-crash-quarantine-tests_2026-06-11.md',
            'Specs/Plans/Done/AdvisorPlan_008-dwrite-setmax-guard_2026-06-11.md',
            'Specs/Plans/Done/AdvisorPlan_009-icon-pipeline-investigation_2026-06-11.md',
            'Specs/Plans/Done/AdvisorPlan_010-watcher-incremental-refresh-investigation_2026-06-11.md',
            'Specs/Plans/Done/AdvisorPlan_011-ci-coverage-gaps_2026-06-11.md',
            'Specs/Plans/Done/AdvisorPlan_012-format-autocommit-race_2026-06-11.md',
            'Specs/Plans/Done/AdvisorPlan_013-undo-file-operations-spike_2026-06-11.md',
            'Specs/Plans/Done/AdvisorPlan_014-viewerimgraw-export-uaf-snapshot-pixels_2026-06-16.md',
            'Specs/Plans/Done/AdvisorPlan_015-viewerimgraw-wic-coinit_2026-06-16.md',
            'Specs/Plans/Done/AdvisorPlan_016-viewerimgraw-atomic-export_2026-06-16.md',
            'Specs/Plans/Done/AdvisorPlan_017-async-workitem-dispatch-hardening_2026-06-16.md',
            'Specs/Plans/Done/AdvisorPlan_018-viewerweb-script-injection_2026-06-16.md',
            'Specs/Plans/Done/AdvisorPlan_019-viewerweb-async-load-data-race_2026-06-16.md',
            'Specs/Plans/Done/AdvisorPlan_020-viewertext-hex-short-read_2026-06-16.md',
            'Specs/Plans/Done/AdvisorPlan_021-viewerpe-no-ui-thread-join_2026-06-16.md',
            'Specs/Plans/Done/AdvisorPlan_022-viewervlc-dll-search-path_2026-06-16.md',
            'Specs/Plans/Done/AdvisorPlan_023-viewervlc-detach-hwnd-before-teardown_2026-06-16.md',
            'Specs/Plans/Done/AdvisorPlan_024-viewertext-glyph-geometry_2026-06-16.md',
            'Specs/Plans/Done/AdvisorPlan_025-viewerspace-large-tree-reliability_2026-06-16.md',
            'Specs/Plans/Done/AdvisorPlan_026-viewersqlite-perf-and-headers_2026-06-16.md',
            'Specs/Plans/Done/AdvisorPlan_027-viewerimgraw-correctness-bundle_2026-06-16.md',
            'Specs/Plans/Done/AdvisorPlan_028-viewertext-reliability-and-deadcode_2026-06-16.md',
            'Specs/Plans/Done/AdvisorPlan_029-viewer-hardcoded-strings-localization_2026-06-16.md',
            'Specs/Plans/Done/AdvisorPlan_030-viewer-scaffolding-consolidation_2026-06-16.md',
            'Specs/Plans/Done/CI_FormatAutocommitRace_2026-08-04.md',
            'Specs/Plans/Done/Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md',
            'Specs/Plans/Done/Product_WhimFilesGapAnalysisAndImprovementPlan_2026-07-08.md',
            'Specs/Plans/Done/DxUi_Uia_ContinuationBaton_2026-06-29.md',
            'Specs/Plans/Done/FolderView_WarpDrive_ContinuationBaton_2026-06-29.md',
            'Specs/Plans/Done/Terminal_LibGhosttyVt_UpgradeAndContinuousReadiness_2026-08-25.md',
            'Specs/Plans/Done/UI_MenuInformationArchitectureAndConsistency_2026-08-28.md',
            'Specs/Plans/Done/UI_MenuContextEncodingReviewRemediation_2026-08-29.md',
            'Specs/Plans/Done/UI_FileOperationsPopupParallelProgressPendingPrompt_2026-08-28.md',
            'Specs/Plans/Done/Operation_FileOperations_PostCloseoutArchitectureDebt_2026-08-28.md',
            'Specs/Plans/Done/Operation_FileOperations_ReliabilityFirstSuccessorPlan_2026-08-29.md',
            'Specs/Plans/Done/Operation_FileOperations_ReliabilityFirst_ExecutionRecord_2026-09-05.md',
            'Specs/Plans/Done/Operation_Beeline_SameEndpointMoveIsRename_2026-09-07.md',
            'Specs/Plans/Done/FileOperations_BeelineResidualsAndProviderRoutes_2026-09-08.md',
            'Specs/Plans/Done/FileOperations_DiscoveryScopeAndLeafProgress_2026-09-13.md',
            'Specs/Plans/Done/FileOperations_DirectApiDiscoveryContract_2026-09-14.md',
            'Specs/Plans/Done/FileOperations_ProviderDiscoveryParity_2026-09-14.md',
            'Specs/Plans/Done/FileOperations_S3MarkerAndResiduals_2026-09-15.md',
            'Specs/Plans/Done/DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md',
            'Specs/Plans/Done/DxUi_SharedLibrary/GapAnalysis_2026-09-09.md',
            'Specs/Plans/Done/DxUi_SharedLibrary/RetirementDispositions_2026-09-12.md'
        )
    )

    $tree = Get-RSGitOutput -RepositoryRoot $RepositoryRoot -Arguments @(
        'ls-tree', '-r', $BaseCommit, '--', 'Specs/Plans/Done'
    ) -AllowFailure
    if ($tree.ExitCode -ne 0) {
        return [pscustomobject]@{
            Available = $false
            BaseCommit = $BaseCommit
            BaselineCount = 0
            Modified = @()
            Missing = @()
            Unexpected = @()
            IsValid = $true
        }
    }

    $baseline = @{}
    foreach ($line in $tree.Output) {
        if ($line -match '^\d+\s+blob\s+(?<blob>[0-9a-f]+)\t(?<path>.+)$') {
            $baseline[(ConvertTo-RSForwardSlashPath -Path $matches['path'])] = $matches['blob']
        }
    }

    $currentPaths = @(Get-RSInventoryPaths -RepositoryRoot $RepositoryRoot | Where-Object { $_ -like 'Specs/Plans/Done/*' })
    $missing = [System.Collections.Generic.List[string]]::new()
    $modified = [System.Collections.Generic.List[string]]::new()
    foreach ($entry in $baseline.GetEnumerator()) {
        $fullPath = Join-Path $RepositoryRoot $entry.Key
        if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) {
            $missing.Add($entry.Key)
            continue
        }
        $hash = Get-RSGitOutput -RepositoryRoot $RepositoryRoot -Arguments @('hash-object', '--', $entry.Key)
        if ($hash.Output.Count -ne 1 -or $hash.Output[0] -cne $entry.Value) {
            $modified.Add($entry.Key)
        }
    }
    $unexpected = @($currentPaths | Where-Object {
            -not $baseline.ContainsKey($_) -and $_ -notin $AllowedAdditionalPaths
        })

    return [pscustomobject]@{
        Available = $true
        BaseCommit = $BaseCommit
        BaselineCount = $baseline.Count
        Modified = @($modified)
        Missing = @($missing)
        Unexpected = @($unexpected)
        AllowedAdditional = @($currentPaths | Where-Object { $_ -in $AllowedAdditionalPaths })
        IsValid = ($modified.Count -eq 0 -and $missing.Count -eq 0 -and $unexpected.Count -eq 0)
    }
}

function Get-RSSpecInventory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RepositoryRoot,
        [string]$Root = 'Specs',
        [string]$ConsistencyLedgerPath = 'Specs/NormativeConsistency.json',
        [string]$ProtectedDoneBaseCommit = '6aecfde8e'
    )

    $repositoryFull = [IO.Path]::GetFullPath($RepositoryRoot)
    $headResult = Get-RSGitOutput -RepositoryRoot $repositoryFull -Arguments @('rev-parse', 'HEAD') -AllowFailure
    $head = if ($headResult.ExitCode -eq 0 -and $headResult.Output.Count -gt 0) { $headResult.Output[0] } else { $null }
    $allPaths = @(Get-RSInventoryPaths -RepositoryRoot $repositoryFull)
    $markdownPaths = @($allPaths | Where-Object {
            $_ -match '\.md$' -and
            ($_ -notlike 'Specs/TestRuns/*' -or $_ -eq 'Specs/TestRuns/README.md') -and
            ($_ -like "$Root/*" -or $_ -like 'plans/*' -or $_ -like 'docs/*' -or $_ -eq 'README.md' -or $_ -eq 'AGENTS.md' -or $_ -eq 'CLAUDE.md')
        })
    $ledger = Read-RSConsistencyLedger -RepositoryRoot $repositoryFull -LedgerPath $ConsistencyLedgerPath
    $files = [System.Collections.Generic.List[object]]::new()
    $analysisMap = @{}

    foreach ($relativePath in $markdownPaths) {
        $fullPath = Join-Path $repositoryFull $relativePath
        if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) { continue }
        $text = Get-Content -LiteralPath $fullPath -Raw
        $headings = @(Get-RSMarkdownHeadings -Text $text)
        $links = @(Get-RSMarkdownLinks -Text $text)
        $authorityClass = Get-RSSpecAuthorityClass -RelativePath $relativePath
        $markers = Get-RSMarkdownMarkers -Text $text
        $consistency = if ($ledger.DocumentMap.ContainsKey($relativePath)) { $ledger.DocumentMap[$relativePath] } else { $null }
        $record = [pscustomobject]@{
            Path = $relativePath
            IsInScope = $relativePath -like "$Root/*"
            Bytes = (Get-Item -LiteralPath $fullPath).Length
            LineCount = ([regex]::Matches($text, "`n")).Count + 1
            FirstHeading = if ($headings.Count -gt 0) { $headings[0].Title } else { $null }
            AuthorityClass = $authorityClass
            Domain = Get-RSSpecDomain -RelativePath $relativePath
            Headings = $headings
            RawLinks = $links
            Markers = $markers
            Consistency = $consistency
            WipStatus = $null
            InboundLinks = @()
            OutboundLinks = @()
        }
        $files.Add($record)
        $analysisMap[$relativePath] = $record
    }

    $findings = [System.Collections.Generic.List[object]]::new()
    $resolvedLinks = [System.Collections.Generic.List[object]]::new()
    foreach ($file in $files) {
        $sourceDirectory = Split-Path -Parent (Join-Path $repositoryFull $file.Path)
        foreach ($link in $file.RawLinks) {
            $rawTarget = [Uri]::UnescapeDataString(([string]$link.Target).Replace('\', '/'))
            if ([string]::IsNullOrWhiteSpace($rawTarget)) { continue }
            if ($rawTarget -match '^(?i:https?|mailto):') { continue }
            $fragment = $null
            $pathPart = $rawTarget
            $hashIndex = $rawTarget.IndexOf('#')
            if ($hashIndex -ge 0) {
                $fragment = $rawTarget.Substring($hashIndex + 1).ToLowerInvariant()
                $pathPart = $rawTarget.Substring(0, $hashIndex)
            }
            $candidate = if ([string]::IsNullOrWhiteSpace($pathPart)) {
                Join-Path $repositoryFull $file.Path
            } else {
                Join-Path $sourceDirectory $pathPart
            }
            if (-not (Test-RSPathInsideRoot -RepositoryRoot $repositoryFull -Candidate $candidate)) {
                $findings.Add([pscustomobject]@{
                    Code = 'LINK_ROOT_ESCAPE'; Severity = 'Error'; Path = $file.Path; Line = $link.Line
                    Message = "Local link escapes the repository root: $($link.Target)"
                })
                continue
            }
            $targetFull = [IO.Path]::GetFullPath($candidate)
            $targetRelative = Get-RSRepositoryRelativePath -RepositoryRoot $repositoryFull -Path $targetFull
            $targetExists = Test-Path -LiteralPath $targetFull
            $targetRecord = if ($analysisMap.ContainsKey($targetRelative)) { $analysisMap[$targetRelative] } else { $null }
            $anchorExists = $true
            if (-not [string]::IsNullOrWhiteSpace($fragment) -and $targetExists -and $null -ne $targetRecord) {
                $anchorExists = @($targetRecord.Headings | Where-Object { $_.Slug -ceq $fragment }).Count -gt 0
            }
            $resolved = [pscustomobject]@{
                Source = $file.Path; Line = $link.Line; Target = $targetRelative; Fragment = $fragment
                TargetExists = $targetExists; AnchorExists = $anchorExists
            }
            $resolvedLinks.Add($resolved)
            $file.OutboundLinks += $resolved
            if ($null -ne $targetRecord) { $targetRecord.InboundLinks += $resolved }
            if (-not $targetExists) {
                $severity = if ($file.AuthorityClass -in @('Historical', 'HistoricalData', 'Evidence', 'Prototype', 'Legal', 'ExternalDocumentation', 'HumanManagedNote')) { 'Information' } else { 'Error' }
                $findings.Add([pscustomobject]@{
                    Code = 'LINK_MISSING_TARGET'; Severity = $severity; Path = $file.Path; Line = $link.Line
                    Message = "Local link target does not exist: $($link.Target)"
                })
            } elseif (-not $anchorExists) {
                $severity = if ($file.AuthorityClass -in @('Historical', 'HistoricalData', 'Evidence', 'Prototype', 'Legal', 'ExternalDocumentation', 'HumanManagedNote')) { 'Information' } else { 'Error' }
                $findings.Add([pscustomobject]@{
                    Code = 'LINK_MISSING_ANCHOR'; Severity = $severity; Path = $file.Path; Line = $link.Line
                    Message = "Local link heading anchor does not exist: $($link.Target)"
                })
            }
        }
    }

    $wipIndexPath = 'Specs/Plans/WIP/README.md'
    $wipEntries = @()
    if ($analysisMap.ContainsKey($wipIndexPath)) {
        $indexText = Get-Content -LiteralPath (Join-Path $repositoryFull $wipIndexPath) -Raw
        $wipEntries = @(Get-RSWipIndexEntries -IndexText $indexText)
    }
    $entryGroups = @($wipEntries | Group-Object Path)
    $directWip = @($files | Where-Object { $_.AuthorityClass -eq 'ActivePlan' -and (Split-Path -Parent $_.Path).Replace('\', '/') -eq 'Specs/Plans/WIP' })
    foreach ($file in $directWip) {
        $matches = @($wipEntries | Where-Object { $_.Path -ceq $file.Path })
        if ($matches.Count -eq 1) {
            $file.WipStatus = $matches[0].Status
        } elseif ($matches.Count -eq 0) {
            $findings.Add([pscustomobject]@{
                Code = 'WIP_UNINDEXED'; Severity = 'Error'; Path = $file.Path; Line = 1
                Message = 'Direct WIP Markdown file is not present in the machine-readable WIP index.'
            })
        } else {
            $findings.Add([pscustomobject]@{
                Code = 'WIP_DUPLICATE_INDEX'; Severity = 'Error'; Path = $file.Path; Line = 1
                Message = "Direct WIP Markdown file is indexed $($matches.Count) times."
            })
        }
    }
    foreach ($entry in $wipEntries) {
        if (-not $analysisMap.ContainsKey($entry.Path)) {
            $findings.Add([pscustomobject]@{
                Code = 'WIP_STALE_INDEX'; Severity = 'Error'; Path = $wipIndexPath; Line = 1
                Message = "WIP index points to missing path '$($entry.Path)'."
            })
        }
    }

    $normative = @($files | Where-Object { $_.AuthorityClass -in @('Normative', 'NormativeCatalog') })
    foreach ($file in $normative) {
        if ($null -eq $file.Consistency) {
            $findings.Add([pscustomobject]@{
                Code = 'NORMATIVE_MISSING_CONSISTENCY'; Severity = 'Error'; Path = $file.Path; Line = 1
                Message = 'Normative file has no consistency-ledger record.'
            })
        } else {
            $status = [string]$file.Consistency.status
            if ($status -ne 'VERIFIED_MATCH') {
                $findings.Add([pscustomobject]@{
                    Code = 'NORMATIVE_UNRESOLVED_CONSISTENCY'; Severity = 'Error'; Path = $file.Path; Line = 1
                    Message = "Normative consistency status is '$status', not VERIFIED_MATCH."
                })
            }
            if (@($file.Consistency.implementationAnchors).Count -eq 0 -or @($file.Consistency.testAnchors).Count -eq 0) {
                $findings.Add([pscustomobject]@{
                    Code = 'NORMATIVE_MISSING_EVIDENCE'; Severity = 'Error'; Path = $file.Path; Line = 1
                    Message = 'Normative consistency record requires implementation and test anchors.'
                })
            }
            foreach ($anchor in @($file.Consistency.implementationAnchors) + @($file.Consistency.testAnchors) + @($file.Consistency.siblingSpecs)) {
                $anchorText = [string]$anchor
                $anchorPath = ($anchorText -split '#', 2)[0]
                $anchorPath = $anchorPath -replace ':\d+$', ''
                if ([string]::IsNullOrWhiteSpace($anchorPath) -or
                    -not (Test-Path -LiteralPath (Join-Path $repositoryFull $anchorPath) -PathType Leaf)) {
                    $findings.Add([pscustomobject]@{
                        Code = 'NORMATIVE_MISSING_ANCHOR_PATH'; Severity = 'Error'; Path = $file.Path; Line = 1
                        Message = "Consistency evidence anchor does not resolve to a file: $anchorText"
                    })
                }
            }
            if ([string]::IsNullOrWhiteSpace([string]$file.Consistency.contractOwner) -or
                [string]::IsNullOrWhiteSpace([string]$file.Consistency.reviewerNote) -or
                @($file.Consistency.reviewedSections).Count -eq 0) {
                $findings.Add([pscustomobject]@{
                    Code = 'NORMATIVE_INCOMPLETE_REVIEW'; Severity = 'Error'; Path = $file.Path; Line = 1
                    Message = 'Consistency row requires a contract owner, reviewed sections, and reviewer note.'
                })
            }
        }

        $exceptions = if ($null -ne $file.Consistency -and
            @($file.Consistency.PSObject.Properties.Name) -contains 'markerExceptions' -and
            $null -ne $file.Consistency.markerExceptions) {
            @($file.Consistency.markerExceptions)
        } else { @() }
        if (($file.Markers.Checkboxes -gt 0 -or $file.Markers.FuturePlanHeadings -gt 0) -and @($exceptions).Count -eq 0) {
            $findings.Add([pscustomobject]@{
                Code = 'NORMATIVE_PLAN_MARKER'; Severity = 'Error'; Path = $file.Path; Line = 1
                Message = "Normative prose contains $($file.Markers.Checkboxes) checkbox(es) and $($file.Markers.FuturePlanHeadings) future/plan heading(s) without a reviewed exception."
            })
        }
    }

    foreach ($file in $directWip) {
        $text = Get-Content -LiteralPath (Join-Path $repositoryFull $file.Path) -Raw
        if ($text -match '(?im)^\s*(?:this\s+)?plan\s+(?:is\s+)?(?:the\s+)?(?:authoritative|source of truth)|plan\s+overrides\s+(?:the\s+)?(?:spec|contract)') {
            $findings.Add([pscustomobject]@{
                Code = 'WIP_AUTHORITY_CLAIM'; Severity = 'Error'; Path = $file.Path; Line = 1
                Message = 'WIP plan claims authority over current contracts.'
            })
        }
        if ($file.WipStatus -eq 'RETIRED') {
            $required = @('Status', 'Retirement date', 'Reason', 'Authoritative', 'git show')
            $missingFields = @($required | Where-Object { $text -notmatch [regex]::Escape($_) })
            if ($file.LineCount -gt 40 -or $file.Bytes -gt 4096 -or $missingFields.Count -gt 0) {
                $findings.Add([pscustomobject]@{
                    Code = 'TOMBSTONE_INVALID'; Severity = 'Error'; Path = $file.Path; Line = 1
                    Message = "Retired tombstone violates size/field contract; missing: $($missingFields -join ', ')."
                })
            }
        }
    }

    $done = Test-RSProtectedDoneHistory -RepositoryRoot $repositoryFull -BaseCommit $ProtectedDoneBaseCommit
    if (-not $done.IsValid) {
        $findings.Add([pscustomobject]@{
            Code = 'DONE_HISTORY_CHANGED'; Severity = 'Error'; Path = 'Specs/Plans/Done'; Line = 1
            Message = "Protected Done history changed. Modified=$($done.Modified.Count), missing=$($done.Missing.Count), unexpected=$($done.Unexpected.Count)."
        })
    }

    $blocking = @($findings | Where-Object { $_.Severity -eq 'Error' })
    return [pscustomobject]@{
        SchemaVersion = 1
        RepositoryRoot = $repositoryFull
        RepositoryCommit = $head
        GeneratedAtUtc = [DateTime]::UtcNow.ToString('o')
        Root = $Root
        ConsistencyLedger = $ledger
        Files = @($files)
        Links = @($resolvedLinks)
        WipIndexEntries = @($wipEntries)
        ProtectedDone = $done
        Findings = @($findings)
        BlockingFindings = $blocking
        Summary = [pscustomobject]@{
            TrackedAndUntrackedPaths = $allPaths.Count
            TestRunPaths = @($allPaths | Where-Object { $_ -like 'Specs/TestRuns/*' }).Count
            Files = $files.Count
            InScopeFiles = @($files | Where-Object IsInScope).Count
            NormativeFiles = $normative.Count
            DirectWipFiles = $directWip.Count
            WipIndexEntries = $wipEntries.Count
            Findings = $findings.Count
            BlockingFindings = $blocking.Count
        }
    }
}

Export-ModuleMember -Function @(
    'ConvertTo-RSGitHubHeadingSlug',
    'Get-RSMarkdownCodeMask',
    'Get-RSMarkdownHeadings',
    'Get-RSMarkdownLinks',
    'Get-RSSpecAuthorityClass',
    'Get-RSWipIndexEntries',
    'Get-RSSpecInventory',
    'Test-RSProtectedDoneHistory'
)
