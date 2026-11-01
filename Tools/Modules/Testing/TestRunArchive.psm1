# Private test-evidence module; import through its reviewed callers.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# The checked-in archive contract and its 2 MiB / 5 MiB limits were introduced
# by this commit. Older untouched evidence is grandfathered; any archive added
# or modified from this commit onward belongs to the governed inventory.
$script:RSTestRunArchiveContractCommit = '8b9e36191835cc71b5ee0dfeea4dddd5d942dd5a'
# Parentless snapshot of only pre-contract path/mode/blob triples still byte-identical
# at the already-published source import. It carries no DualTail commit ancestry.
$script:RSTestRunArchiveImportedBaselineCommit = '4a9489e9cb2d1190528b8ec35971fba885fd7792'
$script:RSTestRunArchiveRelocationSourceCommit = '6b6f5646be297fbc232d137f1df6a00c5cb74461'
$script:RSTestRunArchiveRelocationBaselineCommit = '8b9e36191835cc71b5ee0dfeea4dddd5d942dd5a'
$script:RSTestRunArchiveRelocationPathCount = 1394
$script:RSTestRunArchiveCurrentRelocationPathCount = 1012

function Get-RSTestRunArchiveWorktreeBlobMap {
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Path
    )

    $root = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\', '/')
    $result = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)
    $batchPaths = [Collections.Generic.List[string]]::new()
    $fallbackPaths = [Collections.Generic.List[string]]::new()
    foreach ($pathValue in $Path) {
        if ([string]::IsNullOrWhiteSpace($pathValue) -or $pathValue -match '[\r\n]') {
            $fallbackPaths.Add($pathValue)
        } else {
            $batchPaths.Add($pathValue)
        }
    }

    if ($batchPaths.Count -gt 0) {
        $hashRows = @($batchPaths.ToArray() | & git -C $root hash-object --stdin-paths 2>$null)
        if ($LASTEXITCODE -ne 0 -or $hashRows.Count -ne $batchPaths.Count) {
            throw 'Git worktree hash batch failed or returned an incomplete result.'
        }
        for ($index = 0; $index -lt $batchPaths.Count; $index++) {
            $hash = ([string]$hashRows[$index]).Trim()
            if ($hash -notmatch '^[0-9a-fA-F]{40,64}$') {
                throw "Git worktree hash batch returned an invalid object id for '$($batchPaths[$index])'."
            }
            if (-not $result.TryAdd($batchPaths[$index], $hash.ToLowerInvariant())) {
                throw "Git worktree hash batch contained a duplicate path: $($batchPaths[$index])"
            }
        }
    }

    # Git's stdin-path format is newline-delimited. Retain exact attribute semantics
    # for the exceptional legal pathname by hashing it through the per-path form.
    foreach ($pathValue in $fallbackPaths) {
        if ([string]::IsNullOrWhiteSpace($pathValue)) { continue }
        $hashRow = & git -C $root hash-object "--path=$pathValue" -- $pathValue 2>$null
        if ($LASTEXITCODE -ne 0 -or ([string]$hashRow).Trim() -notmatch '^[0-9a-fA-F]{40,64}$') {
            throw "Git worktree hash failed for '$pathValue'."
        }
        if (-not $result.TryAdd($pathValue, ([string]$hashRow).Trim().ToLowerInvariant())) {
            throw "Git worktree hash batch contained a duplicate path: $pathValue"
        }
    }

    return ,$result
}

function Get-RSTestRunArchiveIndexStageMap {
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Path
    )

    $targetByHex = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)
    foreach ($pathValue in $Path) {
        $hex = [Convert]::ToHexString([Text.Encoding]::UTF8.GetBytes($pathValue))
        if (-not $targetByHex.TryAdd($hex, $pathValue)) {
            throw "Index snapshot request contains a duplicate path: $pathValue"
        }
    }
    $result = [Collections.Generic.Dictionary[string, Collections.Generic.List[object]]]::new([StringComparer]::Ordinal)
    if ($targetByHex.Count -eq 0) { return ,$result }

    $start = [Diagnostics.ProcessStartInfo]::new()
    $start.FileName = 'git.exe'
    $start.UseShellExecute = $false
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $start.CreateNoWindow = $true
    foreach ($value in @('-C', [IO.Path]::GetFullPath($RepoRoot), 'ls-files', '--stage', '-z', '--', 'Specs/TestRuns')) {
        [void]$start.ArgumentList.Add($value)
    }
    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $start
    $memory = [IO.MemoryStream]::new()
    try {
        if (-not $process.Start()) { throw 'Unable to start the staged-index snapshot.' }
        $copyTask = $process.StandardOutput.BaseStream.CopyToAsync($memory)
        $errorTask = $process.StandardError.ReadToEndAsync()
        $null = $copyTask.GetAwaiter().GetResult()
        $errorText = $errorTask.GetAwaiter().GetResult()
        $process.WaitForExit()
        if ($process.ExitCode -ne 0) {
            throw "Git staged-index snapshot failed with exit $($process.ExitCode): $errorText"
        }

        $bytes = $memory.ToArray()
        $offset = 0
        while ($offset -lt $bytes.Length) {
            $recordEnd = [Array]::IndexOf($bytes, [byte]0, $offset)
            if ($recordEnd -lt 0) { throw 'Git staged-index snapshot omitted a NUL record terminator.' }
            $tab = [Array]::IndexOf($bytes, [byte]9, $offset, $recordEnd - $offset)
            if ($tab -lt 0) { throw 'Git staged-index snapshot omitted a path separator.' }
            $header = [Text.Encoding]::ASCII.GetString($bytes, $offset, $tab - $offset)
            $match = [regex]::Match($header, '^(?<mode>[0-7]{6}) (?<blob>[0-9a-fA-F]{40,64}) (?<stage>[0-3])$')
            if (-not $match.Success) { throw "Malformed Git staged-index identity '$header'." }
            $pathHex = [Convert]::ToHexString($bytes, $tab + 1, $recordEnd - $tab - 1)
            if ($targetByHex.ContainsKey($pathHex)) {
                $pathValue = $targetByHex[$pathHex]
                if (-not $result.ContainsKey($pathValue)) {
                    $result.Add($pathValue, [Collections.Generic.List[object]]::new())
                }
                $result[$pathValue].Add([pscustomobject]@{
                        Mode = $match.Groups['mode'].Value
                        Blob = $match.Groups['blob'].Value.ToLowerInvariant()
                        Stage = [int]$match.Groups['stage'].Value
                    })
            }
            $offset = $recordEnd + 1
        }
    } finally {
        $memory.Dispose()
        $process.Dispose()
    }
    return ,$result
}

function Get-RSTestRunArchiveGrandfatheredRelocationState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [string]$ManifestPath = 'Specs/Testing/ProtectedPathRelocations.json',
        [string]$ExpectedSourceCommit = $script:RSTestRunArchiveRelocationSourceCommit,
        [string]$ExpectedBaselineCommit = $script:RSTestRunArchiveRelocationBaselineCommit,
        [int]$ExpectedPathCount = $script:RSTestRunArchiveRelocationPathCount,
        [int]$ExpectedCurrentPathCount = $script:RSTestRunArchiveCurrentRelocationPathCount
    )

    $root = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\', '/')
    $continuationMappings = [ordered]@{
        '2026-07-01_223000_folderview_warpdrive_origin_commands_recheck' = 'c04'
        '2026-07-02_124500_folderview_warpdrive_tracecleanup_tsf_crash_baton' = 'c05'
        '2026-07-02_215000_folderview_warpdrive_commands_green_baton' = 'c06'
        '2026-07-03_083500_folderview_warpdrive_compare_green_viewerpe_red_baton' = 'c07'
        '2026-07-03_093804_folderview_warpdrive_full_red_after_convergence_baton' = 'c08'
        '2026-07-03_100049_folderview_warpdrive_current_baton' = 'c1'
        '2026-07-03_105945_folderview_warpdrive_commands_race_baton' = 'c2'
        '2026-07-03_114630_folderview_warpdrive_preferences_keyboard_baton' = 'c3'
        '2026-07-03_132317_folderview_warpdrive_fullsuite_compare_monitor_baton' = 'c09'
        '2026-07-03_142650_folderview_warpdrive_compare_cache_crash_baton' = 'c10'
        '2026-07-03_150535_folderview_warpdrive_worker_join_commands_status_baton' = 'c11'
        '2026-07-03_155547_folderview_warpdrive_nav_focus_dpi_baton' = 'c12'
        '2026-07-03_161500_folderview_warpdrive_commands_broad_crash' = 'c13'
        '2026-07-03_162914_folderview_warpdrive_speedlimit_focus_baton' = 'c14'
        '2026-07-03_181403_folderview_warpdrive_plugin_config_baton' = 'c15'
        '2026-07-03_202302_folderview_warpdrive_full_closeout_red' = 'c16'
        '2026-07-03_205410_folderview_warpdrive_compare_green_viewerpe_hud_baton' = 'c17'
        '2026-07-03_214957_folderview_warpdrive_full_red_dxui_viewerpe_baton' = 'c18'
        '2026-07-03_215953_folderview_warpdrive_dxui_pointer_root_baton' = 'c19'
        '2026-07-03_225627_folderview_warpdrive_full_compare_ntfs_baton' = 'c20'
        '2026-07-03_235801_folderview_warpdrive_full_red_compare_commands_baton' = 'c21'
    }
    $currentMappings = @(
        [pscustomobject]@{ OldPrefix='Specs/TestRuns/4cb089111a23/Validation/2026-10-26_172200_CiNativeStandaloneFailure'; NewPrefix='Specs/TestRuns/4cb089111a23/Validation/v01'; Class='test-run-current-contract' },
        [pscustomobject]@{ OldPrefix='Specs/TestRuns/4cb089111a23/Validation/2026-09-23_082913_release-pr'; NewPrefix='Specs/TestRuns/4cb089111a23/Validation/v02'; Class='test-run-current-contract' },
        [pscustomobject]@{ OldPrefix='Specs/TestRuns/4cb089111a23/Validation/2026-09-23_094710_footer-wrap-release'; NewPrefix='Specs/TestRuns/4cb089111a23/Validation/v03'; Class='test-run-current-contract' },
        [pscustomobject]@{ OldPrefix='Specs/TestRuns/4cb089111a23/Validation/2026-09-23_footer-summary-release'; NewPrefix='Specs/TestRuns/4cb089111a23/Validation/v04'; Class='test-run-current-contract' },
        [pscustomobject]@{ OldPrefix='Specs/TestRuns/4cb089111a23/Validation/2026-09-27_recycle-skip-gallery'; NewPrefix='Specs/TestRuns/4cb089111a23/Validation/v05'; Class='test-run-current-contract' },
        [pscustomobject]@{ OldPrefix='Specs/TestRuns/4cb089111a23/Validation/2026-10-01_fileops-safety-remediation'; NewPrefix='Specs/TestRuns/4cb089111a23/Validation/v06'; Class='test-run-current-contract' },
        [pscustomobject]@{ OldPrefix='Specs/TestRuns/7d3a1247382a/Validation/space_debug_prepush_native_captures_20261009T201507Z'; NewPrefix='Specs/TestRuns/7d3a1247382a/Validation/v01'; Class='test-run-current-contract' },
        [pscustomobject]@{ OldPrefix='Specs/TestRuns/7d3a1247382a/Validation/space_debug_viewerspace_noninteractive_children_20261009T201507Z'; NewPrefix='Specs/TestRuns/7d3a1247382a/Validation/v02'; Class='test-run-current-contract' },
        [pscustomobject]@{ OldPrefix='Specs/TestRuns/7d3a1247382a/Validation/space_prepush_debug_failure_reports_20261009T201507Z'; NewPrefix='Specs/TestRuns/7d3a1247382a/Validation/v03'; Class='test-run-current-contract' },
        [pscustomobject]@{ OldPrefix='Specs/TestRuns/7d3a1247382a/Validation/space_prepush_release_failure_reports_20261009T213050Z'; NewPrefix='Specs/TestRuns/7d3a1247382a/Validation/v04'; Class='test-run-current-contract' },
        [pscustomobject]@{ OldPrefix='Specs/TestRuns/7d3a1247382a/Validation/space_prepush_release_viewer_diagnostics_20261009T213050Z'; NewPrefix='Specs/TestRuns/7d3a1247382a/Validation/v05'; Class='test-run-current-contract' },
        [pscustomobject]@{ OldPrefix='Specs/TestRuns/7d3a1247382a/Validation/space_release_prepush_native_captures_20261009T213050Z'; NewPrefix='Specs/TestRuns/7d3a1247382a/Validation/v06'; Class='test-run-current-contract' },
        [pscustomobject]@{ OldPrefix='Specs/TestRuns/7d3a1247382a/Commands/viewer_space_same_instance_queued_path_debug_passed_20261009T195007Z'; NewPrefix='Specs/TestRuns/7d3a1247382a/Commands/m01'; Class='test-run-current-contract' },
        [pscustomobject]@{ OldPrefix='Specs/TestRuns/7d3a1247382a/FileOps/2026-10-14_160135-i37-accounting-pr'; NewPrefix='Specs/TestRuns/7d3a1247382a/FileOps/f01'; Class='test-run-current-contract' }
    )
    $continuationTargetRoots = @($continuationMappings.Values | ForEach-Object { "Specs/TestRuns/4cb089111a23/Continuation/$_" })
    $targetRoots = $continuationTargetRoots + @($currentMappings | ForEach-Object { $_.NewPrefix })
    $targetExists = @($targetRoots | Where-Object {
            Test-Path -LiteralPath (Join-Path $root ($_ -replace '/', [IO.Path]::DirectorySeparatorChar)) -PathType Container
        }).Count -gt 0
    $manifestFullPath = Join-Path $root ($ManifestPath -replace '/', [IO.Path]::DirectorySeparatorChar)
    if (-not (Test-Path -LiteralPath $manifestFullPath -PathType Leaf)) {
        return [pscustomobject]@{
            Available = $false
            IsValid = -not $targetExists
            Paths = @()
            Errors = if ($targetExists) { @('Relocated continuation directories exist without their pinned manifest.') } else { @() }
        }
    }

    $errors = [System.Collections.Generic.List[string]]::new()
    try {
        $manifest = Get-Content -LiteralPath $manifestFullPath -Raw | ConvertFrom-Json -ErrorAction Stop
    } catch {
        return [pscustomobject]@{
            Available = $true; IsValid = $false; Paths = @()
            Errors = @("Protected path-relocation manifest is invalid JSON: $($_.Exception.Message)")
        }
    }
    if ($manifest.schema -cne 'red-salamander.immutable-path-relocations.v1' -or
        [string]$manifest.source_commit -cne $ExpectedSourceCommit) {
        $errors.Add('Protected path-relocation manifest schema or source commit is not the reviewed identity.')
    }
    $entries = @($manifest.entries | Where-Object { $_.class -ceq 'test-run-grandfathered' })
    $currentEntries = @($manifest.entries | Where-Object { $_.class -ceq 'test-run-current-contract' })
    Import-Module (Join-Path $PSScriptRoot 'ValidationFingerprint.psm1') -Scope Local -ErrorAction Stop
    $blobSha256ByOid = Get-RSValidationGitBlobContentSha256Map `
        -RepoRoot $root `
        -BlobId @(@($entries) + @($currentEntries) | ForEach-Object { [string]$_.blob })
    if ($entries.Count -ne $ExpectedPathCount) {
        $errors.Add("Expected exactly $ExpectedPathCount grandfathered continuation paths; found $($entries.Count).")
    }
    if ($currentEntries.Count -ne $ExpectedCurrentPathCount) {
        $errors.Add("Expected exactly $ExpectedCurrentPathCount current-contract relocation paths; found $($currentEntries.Count).")
    }

    $resolveRelocatedPath = {
        param([string]$Path)
        foreach ($mapping in $continuationMappings.GetEnumerator()) {
            $oldPrefix = "Specs/TestRuns/4cb089111a23/Continuation/$($mapping.Key)"
            if ($Path.StartsWith($oldPrefix + '/', [StringComparison]::Ordinal)) {
                $newPath = $Path.Replace($oldPrefix, "Specs/TestRuns/4cb089111a23/Continuation/$($mapping.Value)")
                if ($Path.EndsWith('commands_cmd_pane_find_dialog_restored_combined_view_state_grid_doubleClick_activates_selection_after_sort_cycles_and_search_rerun_repro_20260703.log', [StringComparison]::Ordinal)) {
                    $newPath = $newPath.Substring(0, $newPath.LastIndexOf('/') + 1) + 'find-dialog-restored.log'
                }
                if ($Path.EndsWith('runall_commands_cmd_pane_navigationView_path_doubleClick_after_focusAddressBar_tab_traversal_focused_after_nav_menu_waits_20260703.run-all.json', [StringComparison]::Ordinal)) {
                    $newPath = $newPath.Substring(0, $newPath.LastIndexOf('/') + 1) + 'nav-menu-waits.run-all.json'
                }
                return [pscustomobject]@{ Path=$newPath; Class='test-run-grandfathered' }
            }
        }
        foreach ($mapping in $currentMappings) {
            if (-not $Path.StartsWith($mapping.OldPrefix + '/', [StringComparison]::Ordinal)) { continue }
            $newPath = $Path.Replace($mapping.OldPrefix, $mapping.NewPrefix)
            if ($mapping.OldPrefix.EndsWith('CiNativeStandaloneFailure', [StringComparison]::Ordinal)) {
                $runAlias = if ($Path.Contains('/x64-artifact/')) { 'r1' } else { 'r2' }
                $newPath = $newPath.Replace('/x64-artifact/', '/x/').Replace('/arm64-artifact/', '/a/')
                $parts = $newPath.Split('/')
                for ($index = 1; $index -lt $parts.Length; $index++) {
                    if (($parts[$index] -ceq 'runs' -or $parts[$index] -ceq 'completed-runs') -and $index + 1 -lt $parts.Length) {
                        $parts[$index + 1] = $runAlias
                        $index++
                    }
                }
                $newPath = $parts -join '/'
            }
            return [pscustomobject]@{ Path=$newPath; Class=$mapping.Class }
        }
        return $null
    }

    foreach ($commit in @($ExpectedSourceCommit, $ExpectedBaselineCommit)) {
        & git -C $root rev-parse --verify "$commit^{commit}" *> $null
        if ($LASTEXITCODE -ne 0) {
            $errors.Add("Required immutable relocation commit is unavailable: $commit")
        }
        if ($commit -ceq $ExpectedSourceCommit) {
            & git -C $root merge-base --is-ancestor $commit HEAD *> $null
            if ($LASTEXITCODE -ne 0) {
                $errors.Add("Pinned relocation source commit is not an ancestor of HEAD: $commit")
            }
        }
    }

    $sourceTree = @{}
    $baselineTree = @{}
    foreach ($spec in @(
            [pscustomobject]@{ Commit = $ExpectedSourceCommit; Destination = $sourceTree },
            [pscustomobject]@{ Commit = $ExpectedBaselineCommit; Destination = $baselineTree }
        )) {
        $output = & git -C $root ls-tree -r --full-tree $spec.Commit -- Specs/TestRuns 2>$null
        if ($LASTEXITCODE -ne 0) {
            $errors.Add("Cannot read immutable archive identities from $($spec.Commit).")
            continue
        }
        foreach ($line in $output) {
            if ($line -match '^(?<mode>[0-7]{6}) blob (?<blob>[0-9a-f]+)\t(?<path>.+)$') {
                $spec.Destination[([string]$matches['path'] -replace '\\', '/')] = [pscustomobject]@{
                    Mode = $matches['mode']; Blob = $matches['blob']
                }
            }
        }
    }

    # Hash only paths whose source identity and exact destination are already
    # selected by the immutable map. This prevents a malformed manifest from
    # asking Git to read a path outside the repository.
    $approvedWorktreePaths = [Collections.Generic.List[string]]::new()
    $approvedWorktreePathSet = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($entry in @($entries) + @($currentEntries)) {
        $oldPath = [string]$entry.old_path -replace '\\', '/'
        $newPath = [string]$entry.new_path -replace '\\', '/'
        $expected = & $resolveRelocatedPath $oldPath
        if ($null -eq $expected -or $expected.Path -cne $newPath -or -not $sourceTree.ContainsKey($oldPath)) { continue }
        $newFullPath = Join-Path $root ($newPath -replace '/', [IO.Path]::DirectorySeparatorChar)
        if (-not (Test-Path -LiteralPath $newFullPath -PathType Leaf)) { continue }
        $file = Get-Item -LiteralPath $newFullPath -Force
        if ($file.Attributes -band [IO.FileAttributes]::ReparsePoint) { continue }
        if ($approvedWorktreePathSet.Add($newPath)) { $approvedWorktreePaths.Add($newPath) }
    }
    $worktreeBlobByPath = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)
    $worktreeHashError = $null
    try {
        $worktreeBlobByPath = Get-RSTestRunArchiveWorktreeBlobMap -RepoRoot $root -Path $approvedWorktreePaths.ToArray()
    } catch {
        $worktreeHashError = $_.Exception.Message
    }
    $indexRowsByPath = [Collections.Generic.Dictionary[string, Collections.Generic.List[object]]]::new([StringComparer]::Ordinal)
    $indexSnapshotError = $null
    try {
        $indexRowsByPath = Get-RSTestRunArchiveIndexStageMap -RepoRoot $root -Path $approvedWorktreePaths.ToArray()
    } catch {
        $indexSnapshotError = $_.Exception.Message
    }

    $oldSeen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $newSeen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $expectedTargets = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $currentOldSeen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $currentNewSeen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $expectedCurrentTargets = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($entry in $entries) {
        $oldPath = [string]$entry.old_path -replace '\\', '/'
        $newPath = [string]$entry.new_path -replace '\\', '/'
        if (-not $oldSeen.Add($oldPath) -or -not $newSeen.Add($newPath)) {
            $errors.Add("Duplicate old or new grandfathered relocation identity: $oldPath -> $newPath")
            continue
        }
        $expectedNewPath = $null
        foreach ($mapping in $continuationMappings.GetEnumerator()) {
            $oldPrefix = "Specs/TestRuns/4cb089111a23/Continuation/$($mapping.Key)"
            if ($oldPath.StartsWith($oldPrefix + '/', [StringComparison]::Ordinal)) {
                $expectedNewPath = $oldPath.Replace($oldPrefix, "Specs/TestRuns/4cb089111a23/Continuation/$($mapping.Value)")
                if ($oldPath.EndsWith('commands_cmd_pane_find_dialog_restored_combined_view_state_grid_doubleClick_activates_selection_after_sort_cycles_and_search_rerun_repro_20260703.log', [StringComparison]::Ordinal)) {
                    $expectedNewPath = $expectedNewPath.Substring(0, $expectedNewPath.LastIndexOf('/') + 1) + 'find-dialog-restored.log'
                }
                if ($oldPath.EndsWith('runall_commands_cmd_pane_navigationView_path_doubleClick_after_focusAddressBar_tab_traversal_focused_after_nav_menu_waits_20260703.run-all.json', [StringComparison]::Ordinal)) {
                    $expectedNewPath = $expectedNewPath.Substring(0, $expectedNewPath.LastIndexOf('/') + 1) + 'nav-menu-waits.run-all.json'
                }
                break
            }
        }
        if ($null -eq $expectedNewPath -or $newPath -cne $expectedNewPath) {
            $errors.Add("Grandfathered relocation destination does not match the reviewed continuation mapping: $oldPath -> $newPath")
            continue
        }
        [void]$expectedTargets.Add($newPath)
        if ([string]$entry.baseline_commit -cne $ExpectedBaselineCommit) {
            $errors.Add("Grandfathered relocation baseline identity is invalid: $oldPath")
        }
        foreach ($tree in @($sourceTree, $baselineTree)) {
            $identity = $tree[$oldPath]
            if ($null -eq $identity -or [string]$identity.Mode -cne [string]$entry.mode -or
                [string]$identity.Blob -cne [string]$entry.blob) {
                $errors.Add("Grandfathered relocation differs from its pinned source/baseline mode and blob: $oldPath")
                break
            }
        }
        if (Test-Path -LiteralPath (Join-Path $root ($oldPath -replace '/', [IO.Path]::DirectorySeparatorChar))) {
            $errors.Add("Old grandfathered path remains present: $oldPath")
        }
        $newFullPath = Join-Path $root ($newPath -replace '/', [IO.Path]::DirectorySeparatorChar)
        if (-not (Test-Path -LiteralPath $newFullPath -PathType Leaf)) {
            $errors.Add("Mapped grandfathered path is missing: $newPath")
            continue
        }
        $file = Get-Item -LiteralPath $newFullPath -Force
        if ($file.Attributes -band [IO.FileAttributes]::ReparsePoint) {
            $errors.Add("Relocated grandfathered file is a reparse point: $newPath")
            continue
        }
        if ($null -ne $worktreeHashError -or -not $worktreeBlobByPath.ContainsKey($newPath) -or
            $worktreeBlobByPath[$newPath] -cne [string]$entry.blob) {
            $errors.Add("Relocated grandfathered file no longer produces its pinned Git blob: $newPath")
        }
        $blobSha256 = $blobSha256ByOid[[string]$entry.blob]
        if ($blobSha256 -cne [string]$entry.blob_content_sha256) {
            $errors.Add("Relocated grandfathered Git blob differs from its canonical SHA-256: $newPath")
        }
        if ($null -ne $indexSnapshotError) {
            $errors.Add("Cannot inspect the staged identity of relocated grandfathered file: $newPath")
        } else {
            $indexRows = if ($indexRowsByPath.ContainsKey($newPath)) { @($indexRowsByPath[$newPath]) } else { @() }
            if ($indexRows.Count -eq 0) { continue }
            if ($indexRows.Count -ne 1) {
                $errors.Add("Staged mode/blob differs from the pinned grandfathered identity: $newPath")
            } elseif ($indexRows[0].Stage -ne 0) {
                $errors.Add("Staged identity is malformed for relocated grandfathered file: $newPath")
            } elseif ($indexRows[0].Mode -cne [string]$entry.mode -or $indexRows[0].Blob -cne [string]$entry.blob) {
                $errors.Add("Staged mode/blob differs from the pinned grandfathered identity: $newPath")
            }
        }
    }
    foreach ($entry in $currentEntries) {
        $oldPath = [string]$entry.old_path -replace '\\', '/'
        $newPath = [string]$entry.new_path -replace '\\', '/'
        if (-not $currentOldSeen.Add($oldPath) -or -not $currentNewSeen.Add($newPath)) {
            $errors.Add("Duplicate current-contract relocation identity: $oldPath -> $newPath")
            continue
        }
        $expected = & $resolveRelocatedPath $oldPath
        if ($null -eq $expected -or $expected.Class -cne 'test-run-current-contract' -or $newPath -cne $expected.Path) {
            $errors.Add("Current-contract Validation destination does not match the reviewed component map: $oldPath -> $newPath")
            continue
        }
        [void]$expectedCurrentTargets.Add($newPath)
        $sourceIdentity = $sourceTree[$oldPath]
        if ($null -eq $sourceIdentity -or [string]$entry.mode -cne [string]$sourceIdentity.Mode -or
            [string]$entry.blob -cne [string]$sourceIdentity.Blob) {
            $errors.Add("Current-contract Validation relocation differs from its pinned source mode/blob: $oldPath")
        }
        if (Test-Path -LiteralPath (Join-Path $root ($oldPath -replace '/', [IO.Path]::DirectorySeparatorChar))) {
            $errors.Add("Old current-contract Validation path remains present: $oldPath")
        }
        $newFullPath = Join-Path $root ($newPath -replace '/', [IO.Path]::DirectorySeparatorChar)
        if (-not (Test-Path -LiteralPath $newFullPath -PathType Leaf)) {
            $errors.Add("Mapped current-contract Validation path is missing: $newPath")
            continue
        }
        $file = Get-Item -LiteralPath $newFullPath -Force
        if ($file.Attributes -band [IO.FileAttributes]::ReparsePoint) {
            $errors.Add("Relocated current-contract Validation file is a reparse point: $newPath")
            continue
        }
        if ($null -ne $worktreeHashError -or -not $worktreeBlobByPath.ContainsKey($newPath) -or
            $worktreeBlobByPath[$newPath] -cne [string]$entry.blob) {
            $errors.Add("Current-contract Validation file no longer produces its pinned Git blob: $newPath")
        }
        $blobSha256 = $blobSha256ByOid[[string]$entry.blob]
        if ($blobSha256 -cne [string]$entry.blob_content_sha256) {
            $errors.Add("Current-contract Validation Git blob differs from its canonical SHA-256: $newPath")
        }
        if ($null -ne $indexSnapshotError) {
            $errors.Add("Cannot inspect the staged identity of current-contract Validation file: $newPath")
        } else {
            $indexRows = if ($indexRowsByPath.ContainsKey($newPath)) { @($indexRowsByPath[$newPath]) } else { @() }
            if ($indexRows.Count -eq 0) { continue }
            if ($indexRows.Count -ne 1) {
                $errors.Add("Staged mode/blob differs from the pinned current-contract Validation identity: $newPath")
            } elseif ($indexRows[0].Stage -ne 0) {
                $errors.Add("Staged identity is malformed for current-contract Validation file: $newPath")
            } elseif ($indexRows[0].Mode -cne [string]$entry.mode -or $indexRows[0].Blob -cne [string]$entry.blob) {
                $errors.Add("Staged mode/blob differs from the pinned current-contract Validation identity: $newPath")
            }
        }
    }

    foreach ($commitTree in @($sourceTree, $baselineTree)) {
        $continuationEntries = @($commitTree.Keys | Where-Object {
                $matchesRoot = $false
                foreach ($rootName in $continuationMappings.Keys) {
                    if ($_.StartsWith("Specs/TestRuns/4cb089111a23/Continuation/$rootName/", [StringComparison]::Ordinal)) {
                        $matchesRoot = $true
                        break
                    }
                }
                $matchesRoot
            })
        if ($continuationEntries.Count -ne $ExpectedPathCount -or
            @($continuationEntries | Where-Object { -not $oldSeen.Contains($_) }).Count -gt 0) {
            $errors.Add('Pinned source or archive baseline contains missing/unmapped continuation identities.')
        }
    }
    $sourceCurrentEntries = @($sourceTree.Keys | Where-Object {
            $null -ne (& $resolveRelocatedPath $_) -and (& $resolveRelocatedPath $_).Class -ceq 'test-run-current-contract'
        })
    if ($sourceCurrentEntries.Count -ne $ExpectedCurrentPathCount -or
        @($sourceCurrentEntries | Where-Object { -not $currentOldSeen.Contains($_) }).Count -gt 0) {
        $errors.Add('Pinned source contains missing/unmapped current-contract relocation identities.')
    }

    $currentTargets = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($targetRoot in $continuationTargetRoots) {
        $directory = Join-Path $root ($targetRoot -replace '/', [IO.Path]::DirectorySeparatorChar)
        if (-not (Test-Path -LiteralPath $directory -PathType Container)) { continue }
        foreach ($item in Get-ChildItem -LiteralPath $directory -Force -Recurse) {
            if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
                $errors.Add("Reparse point exists inside relocated continuation tree: $($item.FullName)")
                continue
            }
            if (-not $item.PSIsContainer) {
                $relative = ConvertTo-RSRepoRelativePath -RepoRoot $root -Path $item.FullName
                [void]$currentTargets.Add(($relative -replace '\\', '/'))
            }
        }
    }
    if ($currentTargets.Count -ne $expectedTargets.Count -or
        @($currentTargets | Where-Object { -not $expectedTargets.Contains($_) }).Count -gt 0) {
        $errors.Add('Relocated continuation trees contain missing files or unreviewed siblings.')
    }

    $currentValidationTargets = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($validationPrefix in @($currentMappings | ForEach-Object { $_.NewPrefix })) {
        $validationDirectory = Join-Path $root ($validationPrefix -replace '/', [IO.Path]::DirectorySeparatorChar)
        if (-not (Test-Path -LiteralPath $validationDirectory -PathType Container)) { continue }
        foreach ($item in Get-ChildItem -LiteralPath $validationDirectory -Force -Recurse) {
            if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
                $errors.Add("Reparse point exists inside relocated Validation packet: $($item.FullName)")
                continue
            }
            if (-not $item.PSIsContainer) {
                $relative = ConvertTo-RSRepoRelativePath -RepoRoot $root -Path $item.FullName
                [void]$currentValidationTargets.Add(($relative -replace '\\', '/'))
            }
        }
    }
    if ($currentValidationTargets.Count -ne $expectedCurrentTargets.Count -or
        @($currentValidationTargets | Where-Object { -not $expectedCurrentTargets.Contains($_) }).Count -gt 0) {
        $errors.Add('Relocated Validation packet contains missing files or unreviewed siblings.')
    }

    return [pscustomobject]@{
        Available = $true
        IsValid = ($errors.Count -eq 0)
        Paths = if ($errors.Count -eq 0) { @($expectedTargets | Sort-Object) } else { @() }
        Errors = @($errors)
    }
}

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

function Get-RSTestRunArchiveValidatedGrandfatheredRelocationPaths {
    param([Parameter(Mandatory = $true)][string]$RepoRoot)

    $state = Get-RSTestRunArchiveGrandfatheredRelocationState -RepoRoot $RepoRoot
    if (-not $state.IsValid) {
        throw "Pinned grandfathered TestRuns relocation validation failed: $($state.Errors -join '; ')"
    }
    return @($state.Paths)
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
        [string]$CandidateRunPath,

        [AllowEmptyCollection()]
        [string[]]$ZigToolchainTransition = @()
    )

    if ($ZigToolchainTransition.Count -notin @(0, 2) -or
        @($ZigToolchainTransition | Where-Object { $_ -cnotmatch '^[0-9]+[.][0-9]+[.][0-9]+(?:[-+][0-9A-Za-z.-]+)?$' }).Count -ne 0 -or
        ($ZigToolchainTransition.Count -eq 2 -and $ZigToolchainTransition[0] -ceq $ZigToolchainTransition[1])) {
        throw 'ZigToolchainTransition requires exactly two distinct reviewed versions in baseline,candidate order.'
    }
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
    foreach ($property in @('windowsBuild', 'cpu', 'architecture', 'compiler', 'windowsSdk', 'vcpkgCommit')) {
        if ($baselineResults.environment.$property -cne $candidateResults.environment.$property) {
            & $addViolation "Environment identity '$property' differs between baseline and candidate."
        }
    }
    if ($ZigToolchainTransition.Count -eq 0) {
        if ($baselineResults.environment.zig -cne $candidateResults.environment.zig) {
            & $addViolation "Environment identity 'zig' differs between baseline and candidate."
        }
    } elseif ($baselineResults.environment.zig -cne $ZigToolchainTransition[0] -or
        $candidateResults.environment.zig -cne $ZigToolchainTransition[1]) {
        & $addViolation 'Zig versions do not match the exact reviewed baseline,candidate toolchain transition.'
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
        ZigToolchainTransition = [string[]]$ZigToolchainTransition
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

    $baseRef = $null
    foreach ($candidate in @('origin/main', 'origin/master')) {
        & git -C $root rev-parse --verify "$candidate^{commit}" *> $null
        if ($LASTEXITCODE -eq 0) {
            $baseRef = $candidate
            break
        }
    }

    $commands = [System.Collections.Generic.List[object]]::new()
    if ($null -ne $baseRef) {
        $commands.Add(@('diff', '--no-renames', '--name-only', '--diff-filter=AM', "$baseRef...HEAD", '--', 'Specs/TestRuns'))
    }
    $commands.Add(@('diff', '--no-renames', '--name-only', '--diff-filter=AM', '--', 'Specs/TestRuns'))
    $commands.Add(@('diff', '--no-renames', '--cached', '--name-only', '--diff-filter=AM', '--', 'Specs/TestRuns'))

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

    $grandfatheredPaths = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($relocatedPath in Get-RSTestRunArchiveValidatedGrandfatheredRelocationPaths -RepoRoot $root) {
        [void]$grandfatheredPaths.Add($relocatedPath)
    }
    return @($paths | Where-Object { -not $grandfatheredPaths.Contains($_) } | Sort-Object)
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

            $contractRange = $null
            if ($hasContractCommit) {
                $contractRange = "$script:RSTestRunArchiveContractCommit^..HEAD"
            } else {
                & git -C $root rev-parse --verify "$script:RSTestRunArchiveImportedBaselineCommit^{commit}" *> $null
                if ($LASTEXITCODE -eq 0) {
                    & git -C $root merge-base --is-ancestor $script:RSTestRunArchiveImportedBaselineCommit HEAD *> $null
                    if ($LASTEXITCODE -eq 0) {
                        $contractRange = "$script:RSTestRunArchiveImportedBaselineCommit..HEAD"
                    }
                }
            }

            if ($null -ne $contractRange) {
                $paths = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                $commands = @(
                    @('diff', '--no-renames', '--name-only', '--diff-filter=AM', $contractRange, '--', 'Specs/TestRuns'),
                    @('diff', '--no-renames', '--name-only', '--diff-filter=AM', '--', 'Specs/TestRuns'),
                    @('diff', '--no-renames', '--cached', '--name-only', '--diff-filter=AM', '--', 'Specs/TestRuns')
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

            $grandfatheredPaths = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($relocatedPath in Get-RSTestRunArchiveValidatedGrandfatheredRelocationPaths -RepoRoot $root) {
                [void]$grandfatheredPaths.Add($relocatedPath)
            }
            return @($trackedPaths |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                    ForEach-Object { $_.Trim() } |
                    Where-Object { -not $grandfatheredPaths.Contains($_) } |
                    Where-Object { Test-Path -LiteralPath (Join-Path $root $_) -PathType Leaf } |
                    Sort-Object -Unique)
        }
    }

    $testRunsRoot = Join-Path $root 'Specs\TestRuns'
    if (-not (Test-Path -LiteralPath $testRunsRoot -PathType Container)) {
        return @()
    }

    $grandfatheredPaths = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($relocatedPath in Get-RSTestRunArchiveValidatedGrandfatheredRelocationPaths -RepoRoot $root) {
        [void]$grandfatheredPaths.Add($relocatedPath)
    }
    return @(Get-ChildItem -LiteralPath $testRunsRoot -File -Recurse -Force |
            ForEach-Object { ConvertTo-RSRepoRelativePath -RepoRoot $root -Path $_.FullName } |
            Where-Object { -not $grandfatheredPaths.Contains($_) } |
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
