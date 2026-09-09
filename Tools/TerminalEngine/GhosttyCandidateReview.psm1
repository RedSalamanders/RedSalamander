Set-StrictMode -Version Latest

$validationModule = Join-Path (Split-Path -Parent $PSScriptRoot) 'Modules\Testing\ValidationFingerprint.psm1'
$lifecycleModule = Join-Path $PSScriptRoot 'GhosttyRuntimeLifecycle.psm1'
Import-Module $validationModule -ErrorAction Stop
Import-Module $lifecycleModule -ErrorAction Stop

$script:OfficialRepository = 'https://github.com/ghostty-org/ghostty.git'
$script:ReviewPathFilter = @(
    'build.zig',
    'build.zig.zon',
    'include/ghostty/vt',
    'src/build/GhosttyLibVt.zig',
    'src/lib_vt.zig',
    'src/terminal')

function Get-RSGhosttyCandidateReviewUtcText {
    param([Parameter(Mandatory = $true)][DateTime]$Value)

    return $Value.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ', [Globalization.CultureInfo]::InvariantCulture)
}

function Invoke-RSGhosttyCandidateGit {
    param(
        [Parameter(Mandatory = $true)][string]$RepositoryPath,
        [Parameter(Mandatory = $true)][string[]]$Arguments,
        [Parameter(Mandatory = $true)][string]$Context,
        [switch]$AllowFailure
    )

    $output = @(& git.exe -C $RepositoryPath @Arguments 2>&1)
    $exitCode = $LASTEXITCODE
    if (-not $AllowFailure -and $exitCode -ne 0) {
        throw "${Context} failed with git exit ${exitCode}: $($output -join "`n")"
    }
    return [pscustomobject]@{ ExitCode = $exitCode; Lines = [string[]]$output }
}

function Get-RSGhosttyCandidatePathCategory {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    $path = $Path.Replace('\', '/')
    if ($path.StartsWith('include/ghostty/vt/', [StringComparison]::Ordinal)) { return 'public-abi' }
    if ($path -in @('build.zig', 'build.zig.zon', 'src/build/GhosttyLibVt.zig', 'src/lib_vt.zig')) { return 'build-toolchain' }
    if ($path.StartsWith('src/terminal/kitty/', [StringComparison]::Ordinal)) { return 'kitty-graphics' }
    if ($path.StartsWith('src/terminal/snapshot/', [StringComparison]::Ordinal)) { return 'snapshot-serialization' }
    if ($path -match '(?i)(?:^|/)(?:osc|clipboard|paste)(?:[_.\/]|$)') { return 'clipboard-security' }
    if ($path -match '(?i)(?:alloc|capacity|page|grapheme|ring|memory)') { return 'allocation-memory' }
    if ($path.StartsWith('src/terminal/', [StringComparison]::Ordinal)) { return 'parser-state' }
    if ($path -match '(?i)(?:license|copying|notice)') { return 'licenses' }
    if ($path -match '(?i)(?:^|/)(?:test|tests)(?:/|$)') { return 'tests' }
    return 'irrelevant-to-lib-vt'
}

function Get-RSGhosttyCandidateChangeInventory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$BaseCommit,
        [Parameter(Mandatory = $true)][string]$CandidateCommit
    )

    $rangeArguments = @('diff', '--name-status', $BaseCommit, $CandidateCommit, '--') + $script:ReviewPathFilter
    $statusLines = (Invoke-RSGhosttyCandidateGit -RepositoryPath $SourcePath -Arguments $rangeArguments -Context 'Candidate name/status inventory').Lines
    $numstatArguments = @('diff', '--numstat', $BaseCommit, $CandidateCommit, '--') + $script:ReviewPathFilter
    $numstatLines = (Invoke-RSGhosttyCandidateGit -RepositoryPath $SourcePath -Arguments $numstatArguments -Context 'Candidate line inventory').Lines

    $lineStats = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
    $insertions = 0
    $deletions = 0
    foreach ($line in $numstatLines) {
        $parts = $line -split "`t"
        if ($parts.Count -lt 3 -or $parts[0] -cnotmatch '^\d+$' -or $parts[1] -cnotmatch '^\d+$') {
            continue
        }
        $path = [string]$parts[$parts.Count - 1]
        $added = [int]$parts[0]
        $removed = [int]$parts[1]
        $insertions += $added
        $deletions += $removed
        $lineStats[$path] = [pscustomobject]@{ Insertions = $added; Deletions = $removed }
    }

    $entries = @($statusLines | ForEach-Object {
            $parts = $_ -split "`t"
            if ($parts.Count -lt 2) { throw "Malformed candidate name/status line: $_" }
            $path = [string]$parts[$parts.Count - 1]
            $stats = if ($lineStats.ContainsKey($path)) { $lineStats[$path] } else { [pscustomobject]@{ Insertions = 0; Deletions = 0 } }
            [ordered]@{
                path = $path
                status = [string]$parts[0]
                category = Get-RSGhosttyCandidatePathCategory -Path $path
                insertions = [int]$stats.Insertions
                deletions = [int]$stats.Deletions
            }
        } | Sort-Object { $_.path })
    $categoryCounts = @($entries | Group-Object { $_.category } | Sort-Object Name | ForEach-Object {
            [ordered]@{ category = $_.Name; count = $_.Count }
        })
    return [ordered]@{
        fileCount = $entries.Count
        insertions = $insertions
        deletions = $deletions
        categoryCounts = $categoryCounts
        entries = $entries
    }
}

function Get-RSGhosttyCandidateBlobText {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$Commit,
        [Parameter(Mandatory = $true)][string]$Path
    )

    $result = Invoke-RSGhosttyCandidateGit -RepositoryPath $SourcePath -Arguments @('show', "${Commit}:$Path") -Context "Read $Path at $Commit"
    return ($result.Lines -join "`n")
}

function Get-RSGhosttyCandidatePublicModel {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$Commit
    )

    $headerPaths = @((Invoke-RSGhosttyCandidateGit -RepositoryPath $SourcePath -Arguments @('ls-tree', '-r', '--name-only', $Commit, '--', 'include/ghostty/vt') -Context "List public headers at $Commit").Lines |
            Where-Object { $_.EndsWith('.h', [StringComparison]::Ordinal) } |
            Sort-Object -Unique -CaseSensitive)
    $declarations = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)
    $enumValues = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::Ordinal)
    foreach ($path in $headerPaths) {
        $text = Get-RSGhosttyCandidateBlobText -SourcePath $SourcePath -Commit $Commit -Path $path
        foreach ($match in [regex]::Matches($text, '(?s)GHOSTTY_API\s+.*?;')) {
            $signature = [regex]::Replace($match.Value, '\s+', ' ').Trim()
            $nameMatch = [regex]::Match($signature, '\b(?<name>ghostty_[a-z0-9_]+)\s*\(')
            if (-not $nameMatch.Success) { continue }
            $name = $nameMatch.Groups['name'].Value
            if ($declarations.ContainsKey($name) -and $declarations[$name] -cne $signature) {
                throw "Public declaration '$name' is ambiguous at $Commit."
            }
            $declarations[$name] = $signature
        }
        foreach ($match in [regex]::Matches($text, '(?m)^\s*(?<name>GHOSTTY_[A-Z0-9_]+)\s*=\s*(?<value>[^,\r\n/]+)')) {
            $name = $match.Groups['name'].Value
            $value = $match.Groups['value'].Value.Trim()
            if ($enumValues.ContainsKey($name) -and $enumValues[$name] -cne $value) {
                throw "Public enum value '$name' is ambiguous at $Commit."
            }
            $enumValues[$name] = $value
        }
    }
    return [pscustomobject]@{ HeaderPaths = $headerPaths; Declarations = $declarations; EnumValues = $enumValues }
}

function Get-RSGhosttyCandidateMapDelta {
    param(
        [Parameter(Mandatory = $true)][Collections.Generic.Dictionary[string, string]]$Base,
        [Parameter(Mandatory = $true)][Collections.Generic.Dictionary[string, string]]$Candidate,
        [ValidateSet('signature', 'value')][string]$ValueName
    )

    $added = @($Candidate.Keys | Where-Object { -not $Base.ContainsKey($_) } | Sort-Object -CaseSensitive | ForEach-Object {
            [ordered]@{ name = $_; $ValueName = $Candidate[$_] }
        })
    $removed = @($Base.Keys | Where-Object { -not $Candidate.ContainsKey($_) } | Sort-Object -CaseSensitive | ForEach-Object {
            [ordered]@{ name = $_; $ValueName = $Base[$_] }
        })
    $changed = @($Base.Keys | Where-Object { $Candidate.ContainsKey($_) -and $Base[$_] -cne $Candidate[$_] } | Sort-Object -CaseSensitive | ForEach-Object {
            [ordered]@{ name = $_; before = $Base[$_]; after = $Candidate[$_] }
        })
    return [ordered]@{
        baseCount = $Base.Count
        candidateCount = $Candidate.Count
        added = $added
        removed = $removed
        changed = $changed
    }
}

function Get-RSGhosttyCandidateAbiReview {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$BaseCommit,
        [Parameter(Mandatory = $true)][string]$CandidateCommit,
        [Parameter(Mandatory = $true)][string[]]$LoaderRequiredExports
    )

    $base = Get-RSGhosttyCandidatePublicModel -SourcePath $SourcePath -Commit $BaseCommit
    $candidate = Get-RSGhosttyCandidatePublicModel -SourcePath $SourcePath -Commit $CandidateCommit
    $headerChanges = @((Invoke-RSGhosttyCandidateGit -RepositoryPath $SourcePath -Arguments @('diff', '--name-only', $BaseCommit, $CandidateCommit, '--', 'include/ghostty/vt') -Context 'Public header change list').Lines | Sort-Object -Unique -CaseSensitive)
    $replacements = @{ ghostty_terminal_mode_get = 'ghostty_terminal_get' }
    $loaderEntries = @($LoaderRequiredExports | Sort-Object -CaseSensitive | ForEach-Object {
            $replacement = if ($replacements.ContainsKey($_)) { [string]$replacements[$_] } else { $null }
            [ordered]@{
                name = $_
                baseDeclared = $base.Declarations.ContainsKey($_)
                candidateDeclared = $candidate.Declarations.ContainsKey($_)
                replacement = $replacement
            }
        })
    return [ordered]@{
        headerChanges = $headerChanges
        declarations = Get-RSGhosttyCandidateMapDelta -Base $base.Declarations -Candidate $candidate.Declarations -ValueName signature
        enumValues = Get-RSGhosttyCandidateMapDelta -Base $base.EnumValues -Candidate $candidate.EnumValues -ValueName value
        loaderRequiredExports = [ordered]@{ count = $loaderEntries.Count; entries = $loaderEntries }
    }
}

function Get-RSGhosttyCandidateSecurityReview {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$BaseCommit,
        [Parameter(Mandatory = $true)][string]$CandidateCommit
    )

    $range = "${BaseCommit}..${CandidateCommit}"
    $queries = [ordered]@{
        'allocation-pointer-bounds' = 'alloc|capacity|overflow|underflow|pointer|slice|usize|range|bounds?'
        'osc-dcs-apc' = 'osc|dcs|apc'
        'clipboard-paste' = 'clipboard|paste|5522'
        'kitty-image' = 'kitty|image|placement|pixel|png'
        'snapshot-serialization' = 'snapshot|encode|decode|crc32'
    }
    $sets = [ordered]@{}
    foreach ($entry in $queries.GetEnumerator()) {
        $hashes = @((Invoke-RSGhosttyCandidateGit -RepositoryPath $SourcePath -Arguments (@('log', '--format=%H', '--regexp-ignore-case', "-G$($entry.Value)", $range, '--') + $script:ReviewPathFilter) -Context "Security query $($entry.Key)").Lines | Sort-Object -Unique -CaseSensitive)
        $sets[$entry.Key] = [Collections.Generic.HashSet[string]]::new([string[]]$hashes, [StringComparer]::Ordinal)
    }
    $logLines = (Invoke-RSGhosttyCandidateGit -RepositoryPath $SourcePath -Arguments (@('log', '--format=%H%x09%aI%x09%s', $range, '--') + $script:ReviewPathFilter) -Context 'Candidate security commit inventory').Lines
    $commits = @($logLines | ForEach-Object {
            $parts = $_ -split "`t", 3
            if ($parts.Count -ne 3) { throw "Malformed candidate commit inventory line: $_" }
            $hash = [string]$parts[0]
            $categories = @($sets.Keys | Where-Object { $sets[$_].Contains($hash) } | Sort-Object -CaseSensitive)
            if ($categories.Count -eq 0) { $categories = @('general-terminal') }
            $paths = @((Invoke-RSGhosttyCandidateGit -RepositoryPath $SourcePath -Arguments (@('show', '--format=', '--name-only', $hash, '--') + $script:ReviewPathFilter) -Context "Candidate commit paths $hash").Lines |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique -CaseSensitive)
            [ordered]@{
                commit = $hash
                date = [string]$parts[1]
                subject = [string]$parts[2]
                url = "https://github.com/ghostty-org/ghostty/commit/$hash"
                categories = [string[]]$categories
                paths = [string[]]$paths
            }
        })
    $categoryCounts = @($sets.Keys | Sort-Object -CaseSensitive | ForEach-Object {
            [ordered]@{ category = $_; commitCount = $sets[$_].Count }
        })
    $findings = @(
        [ordered]@{
            id = 'terminal-mode-query-replacement'
            commits = @('cfc19e8053b96dbc9a7d7994b84ec6eef7eb17de')
            upstreamPaths = @('include/ghostty/vt/terminal.h', 'src/terminal/c/terminal.zig')
            productPaths = @('Plugins/Terminal/Terminal.cpp', 'Plugins/Terminal/TerminalRuntimeLoader.cpp', 'Plugins/Terminal/TerminalRuntimeLoader.h')
            disposition = 'adapter-change-required'
            summary = 'The candidate removes mode_get/mode_set and requires the sized terminal_get data query; the loader and all mode reads require an explicit ABI adaptation.'
        },
        [ordered]@{
            id = 'clipboard-sized-request-and-bounds'
            commits = @('9f2aa93e825735220c44159a2f4856af3ea6e79c', '70f0065759428c0594c7c5befcb20104ff7ab615', 'bc2f7d7d2fa589a3abf9e4f6696e0a7f7c204e4f')
            upstreamPaths = @('include/ghostty/vt/terminal.h', 'src/terminal/osc.zig', 'src/terminal/stream_handler.zig')
            productPaths = @('Plugins/Terminal/Terminal.cpp', 'Plugins/Terminal/Terminal.h')
            disposition = 'adapter-change-required'
            summary = 'Clipboard callbacks, Kitty OSC 5522 transactions, reply ownership, and the configurable write ceiling require size-gated fields, prompt denial for unsupported transactions, and the existing 256 KiB product bound.'
        },
        [ordered]@{
            id = 'kitty-pending-animation-and-placement-lifecycle'
            commits = @('6760c6482be2df6da273239d684086394ccac29a', 'd8b920e504670a7603c912e9219668e2b358b909', '73903f76aa39f1c8ae67e42b61d9712b7b78be2e', 'b5e86a42844e8be67af67c7d612a058ed8cd340c')
            upstreamPaths = @('include/ghostty/vt/kitty_graphics.h', 'src/terminal/kitty/graphics_exec.zig', 'src/terminal/kitty/graphics_storage.zig')
            productPaths = @('Plugins/Terminal/Terminal.cpp', 'Plugins/Terminal/Terminal.h')
            disposition = 'adapter-change-required'
            summary = 'Pending payloads, animation-frame generation, eviction, and pin release change cache state transitions and require explicit pending/ready/replaced/deleted handling.'
        },
        [ordered]@{
            id = 'kitty-resource-and-path-hardening'
            commits = @('590d669c4a72eb9cb990bf0162071c2b9eb0f7ad', 'f766f303a7d39b9f41fdda1b6d71b329f410d45f', 'ec04900ab957c7637584e2002aeb6f24c985bd70', 'e5840bb9bacdaae78c42baa4b276f52eaa91fdbe')
            upstreamPaths = @('src/terminal/kitty/graphics_command.zig', 'src/terminal/kitty/graphics_exec.zig', 'src/terminal/kitty/graphics_image.zig')
            productPaths = @('Plugins/Terminal/Terminal.cpp', 'Tools/TerminalEngine/Patches/Ghostty-WindowsPrivateRuntime.patch')
            disposition = 'upstream-hardening-reviewed'
            summary = 'Primary commits tighten PNG allocations, shared-memory ranges, file paths, and placement geometry. They are reviewed as upstream hardening/correctness changes and are not independently labeled vulnerabilities.'
        },
        [ordered]@{
            id = 'osc-grapheme-and-dcs-parser-bounds'
            commits = @('46767b521358200bfe3f268f365ccd2f218db558', '4e817e79a1d7e3fe7393297e3c8f1269abb6523a')
            upstreamPaths = @('src/terminal/Parser.zig', 'src/terminal/Stream.zig', 'src/terminal/osc.zig')
            productPaths = @('Plugins/Terminal/Terminal.cpp', 'Tests/PluginContractTests/PluginContractTests.cpp')
            disposition = 'regression-coverage-required'
            summary = 'The interval bounds OSC/grapheme allocations and changes DCS high-byte handling; deterministic oversized-control and high-byte DCS regressions are required.'
        },
        [ordered]@{
            id = 'snapshot-decode-ownership-and-bounds'
            commits = @('418b5d18054a015f25efc30592a412a0ac5c6c68', 'b5290e74c42c2fc5f891eb20f08d0ac0c7c634f8', '9a5279db682832274442ba2471d277b2ca9b4fa4')
            upstreamPaths = @('include/ghostty/vt/snapshot.h', 'src/terminal/snapshot/reader.zig', 'src/terminal/snapshot/snapshot.zig')
            productPaths = @('Plugins/Terminal/Terminal.cpp', 'Tests/PluginContractTests/PluginContractTests.cpp')
            disposition = 'regression-coverage-required'
            summary = 'Snapshot decoding changes bounds and reference release behavior; malformed/truncated snapshot and continuation coverage must remain fail-closed.'
        })
    return [ordered]@{
        claimPolicy = 'primary-source-review-without-independent-vulnerability-labels'
        reviewedCommitCount = $commits.Count
        categoryCounts = $categoryCounts
        commits = $commits
        findings = $findings
    }
}

function Get-RSGhosttyCandidateOfficialVerification {
    param(
        [Parameter(Mandatory = $true)][string]$CandidateCommit,
        [string]$VerificationFixturePath
    )

    if (-not [string]::IsNullOrWhiteSpace($VerificationFixturePath)) {
        $response = Get-Content -LiteralPath $VerificationFixturePath -Raw -ErrorAction Stop |
            ConvertFrom-Json -Depth 30 -NoEnumerate -DateKind String -ErrorAction Stop
    }
    else {
        $response = Invoke-RestMethod -Uri "https://api.github.com/repos/ghostty-org/ghostty/commits/$CandidateCommit" `
            -Headers @{ 'User-Agent' = 'RedSalamander-Ghostty-Candidate-Review/1' } -TimeoutSec 30 -ErrorAction Stop
    }
    if ([string]$response.sha -cne $CandidateCommit -or [string]$response.commit.tree.sha -cnotmatch '^[0-9a-f]{40}$') {
        throw 'Official Ghostty candidate verification returned a substituted or malformed identity.'
    }
    if (-not [bool]$response.commit.verification.verified) {
        throw "Official Ghostty candidate verification is not valid: $([string]$response.commit.verification.reason)"
    }
    return [ordered]@{
        verified = $true
        reason = [string]$response.commit.verification.reason
        commitUrl = [string]$response.html_url
        authorDate = [string]$response.commit.author.date
        committerDate = [string]$response.commit.committer.date
    }
}

function New-RSGhosttyRuntimeCandidateReview {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string]$CandidateCommit,
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$ArchivePath,
        [Parameter(Mandatory = $true)][string]$RepeatArchivePath,
        [Parameter(Mandatory = $true)][string]$ObservationReportPath,
        [string]$VerificationFixturePath,
        [DateTime]$GeneratedUtc = [DateTime]::UtcNow
    )

    $repoRoot = [IO.Path]::GetFullPath($RepoRoot).TrimEnd('\')
    $buildRoot = Join-Path $repoRoot '.build'
    $candidateRoot = Join-Path $buildRoot "TerminalEngineUpgrade\$CandidateCommit"
    if ($CandidateCommit -cnotmatch '^[0-9a-f]{40}$') { throw 'Candidate commit must be one lowercase full SHA.' }
    foreach ($path in @($SourcePath, $ArchivePath, $RepeatArchivePath)) {
        $resolved = [IO.Path]::GetFullPath($path)
        if (-not $resolved.StartsWith($candidateRoot.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw "Candidate input must remain below '$candidateRoot': $resolved"
        }
    }
    $sourcePath = [IO.Path]::GetFullPath($SourcePath)
    $archivePath = [IO.Path]::GetFullPath($ArchivePath)
    $repeatArchivePath = [IO.Path]::GetFullPath($RepeatArchivePath)
    $observationReportPath = [IO.Path]::GetFullPath($ObservationReportPath)
    if (-not $observationReportPath.StartsWith($buildRoot.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Candidate observation report must remain below the repository .build root.'
    }

    $lockPath = Join-Path $repoRoot 'External\TerminalEngine\GhosttyRuntimeLock.v1.json'
    $lockSchema = Join-Path $repoRoot 'Specs\Terminal\GhosttyRuntimeLock.schema.json'
    $lock = Read-RSGhosttyRuntimeLock -LiteralPath $lockPath -SchemaPath $lockSchema -RepoRoot $repoRoot
    $baseCommit = [string]$lock.upstream.commit
    $remote = [string]((Invoke-RSGhosttyCandidateGit -RepositoryPath $sourcePath -Arguments @('remote', 'get-url', 'origin') -Context 'Candidate origin').Lines | Select-Object -First 1)
    if ($remote -cne $script:OfficialRepository) { throw "Candidate origin is not the official Ghostty repository: $remote" }
    $head = [string]((Invoke-RSGhosttyCandidateGit -RepositoryPath $sourcePath -Arguments @('rev-parse', 'HEAD') -Context 'Candidate HEAD').Lines | Select-Object -First 1)
    $tree = [string]((Invoke-RSGhosttyCandidateGit -RepositoryPath $sourcePath -Arguments @('rev-parse', "$CandidateCommit`^{tree}") -Context 'Candidate tree').Lines | Select-Object -First 1)
    if ($head -cne $CandidateCommit -or $tree -cnotmatch '^[0-9a-f]{40}$') { throw 'Candidate source is not detached at the exact requested commit/tree.' }
    $status = @((Invoke-RSGhosttyCandidateGit -RepositoryPath $sourcePath -Arguments @('status', '--porcelain=v1') -Context 'Candidate source status').Lines)
    if ($status.Count -ne 0) { throw 'Candidate source must be pristine before review.' }
    $ancestry = Invoke-RSGhosttyCandidateGit -RepositoryPath $sourcePath -Arguments @('merge-base', '--is-ancestor', $baseCommit, $CandidateCommit) -Context 'Candidate ancestry' -AllowFailure
    if ($ancestry.ExitCode -ne 0) { throw 'Candidate is not a descendant of the canonical lock commit.' }
    $countsText = [string]((Invoke-RSGhosttyCandidateGit -RepositoryPath $sourcePath -Arguments @('rev-list', '--left-right', '--count', "${baseCommit}...${CandidateCommit}") -Context 'Candidate ancestry counts').Lines | Select-Object -First 1)
    $counts = $countsText -split '\s+'
    if ($counts.Count -ne 2 -or $counts[0] -cnotmatch '^\d+$' -or $counts[1] -cnotmatch '^\d+$') { throw "Malformed ancestry counts: $countsText" }
    $behindCount = [int]$counts[0]
    $aheadCount = [int]$counts[1]
    if ($behindCount -ne 0) { throw "Candidate history is behind/diverged by $behindCount commits." }

    $zon = Get-RSGhosttyCandidateBlobText -SourcePath $sourcePath -Commit $CandidateCommit -Path 'build.zig.zon'
    $versionMatch = [regex]::Match($zon, '(?m)^\s*\.version\s*=\s*"(?<version>[^"]+)"')
    if (-not $versionMatch.Success) { throw 'Candidate build.zig.zon has no version.' }
    $version = $versionMatch.Groups['version'].Value
    $archiveSha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $archivePath).Hash.ToLowerInvariant()
    $repeatArchiveSha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $repeatArchivePath).Hash.ToLowerInvariant()
    if ($archiveSha256 -cne $repeatArchiveSha256) { throw 'Repeated official candidate archive downloads do not match.' }
    $officialVerification = Get-RSGhosttyCandidateOfficialVerification -CandidateCommit $CandidateCommit -VerificationFixturePath $VerificationFixturePath

    $observationSchema = Join-Path $repoRoot 'Specs\Terminal\GhosttyRuntimeStatus.schema.json'
    $observation = Read-RSValidationJsonFile -LiteralPath $observationReportPath -SchemaPath $observationSchema -AllowedRoot $buildRoot
    if ([string]$observation.mode -cne 'Candidate' -or [string]$observation.observation.candidateRequestedCommit -cne $CandidateCommit -or
        [string]$observation.observation.commit -cne $CandidateCommit -or [string]$observation.observation.tree -cne $tree) {
        throw 'Candidate observation report does not bind the exact requested commit/tree.'
    }

    $inventory = Get-RSGhosttyCandidateChangeInventory -SourcePath $sourcePath -BaseCommit $baseCommit -CandidateCommit $CandidateCommit
    $abi = Get-RSGhosttyCandidateAbiReview -SourcePath $sourcePath -BaseCommit $baseCommit -CandidateCommit $CandidateCommit -LoaderRequiredExports ([string[]]$lock.abi.loaderRequiredExports)
    $security = Get-RSGhosttyCandidateSecurityReview -SourcePath $sourcePath -BaseCommit $baseCommit -CandidateCommit $CandidateCommit
    $generatedText = Get-RSGhosttyCandidateReviewUtcText $GeneratedUtc
    $archiveRelative = [IO.Path]::GetRelativePath($repoRoot, $archivePath).Replace('\', '/')
    $archiveUrl = "https://github.com/ghostty-org/ghostty/archive/$CandidateCommit.tar.gz"
    $review = [ordered]@{
        schemaVersion = 1
        generatedUtc = $generatedText
        source = [ordered]@{
            repository = $script:OfficialRepository
            baseCommit = $baseCommit
            candidateCommit = $CandidateCommit
            tree = $tree
            version = $version
            archive = [ordered]@{ path = $archiveRelative; url = $archiveUrl; sha256 = $archiveSha256; repeatSha256 = $repeatArchiveSha256 }
            officialVerification = $officialVerification
        }
        ancestry = [ordered]@{ isDescendant = $true; aheadCount = $aheadCount; behindCount = $behindCount }
        pathFilter = [string[]]$script:ReviewPathFilter
        inventory = $inventory
        abi = $abi
        security = $security
    }
    [void](New-Item -ItemType Directory -Path $candidateRoot -Force)
    $reviewPath = Join-Path $candidateRoot 'candidate-review.json'
    $reviewSchemaPath = Join-Path $repoRoot 'Specs\Terminal\GhosttyRuntimeCandidateReview.schema.json'
    $reviewSchemaErrors = @()
    $reviewJson = ConvertTo-RSCanonicalJson -InputObject $review
    if (-not (Test-Json -Json $reviewJson -SchemaFile $reviewSchemaPath -ErrorVariable reviewSchemaErrors -ErrorAction SilentlyContinue)) {
        throw "Candidate review does not satisfy its schema: $(@($reviewSchemaErrors | ForEach-Object { $_.Exception.Message }) -join ' | ')"
    }
    Write-RSAtomicCanonicalJsonFile -LiteralPath $reviewPath -InputObject $review `
        -SchemaPath $reviewSchemaPath -AllowedRoot $candidateRoot

    $sourceCommit = [string]((Invoke-RSGhosttyCandidateGit -RepositoryPath $repoRoot -Arguments @('rev-parse', 'HEAD') -Context 'Repository source binding').Lines | Select-Object -First 1)
    $observationRelative = [IO.Path]::GetRelativePath($repoRoot, $observationReportPath).Replace('\', '/')
    $reviewRelative = [IO.Path]::GetRelativePath($repoRoot, $reviewPath).Replace('\', '/')
    $binding = [ordered]@{
        sourceBinding = [ordered]@{ repositoryCommit = $sourceCommit }
        canonicalLock = [ordered]@{
            path = 'External/TerminalEngine/GhosttyRuntimeLock.v1.json'
            sha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $lockPath).Hash.ToLowerInvariant()
            commit = $baseCommit
            tree = [string]$lock.upstream.tree
        }
        candidate = [ordered]@{
            repository = $script:OfficialRepository
            commit = $CandidateCommit
            tree = $tree
            version = $version
            archivePath = $archiveRelative
            archiveUrl = $archiveUrl
            archiveSha256 = $archiveSha256
        }
        ancestry = [ordered]@{ aheadCount = $aheadCount; behindCount = $behindCount }
        observation = [ordered]@{ path = $observationRelative; sha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $observationReportPath).Hash.ToLowerInvariant() }
        review = [ordered]@{ path = $reviewRelative; sha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $reviewPath).Hash.ToLowerInvariant() }
        toolchain = [ordered]@{
            zigVersion = [string]$lock.toolchain.zigVersion
            archives = @($lock.toolchain.archives | ForEach-Object {
                    [ordered]@{ hostPlatform = [string]$_.hostPlatform; fileName = [string]$_.fileName; sha256 = [string]$_.sha256; executableSha256 = [string]$_.executableSha256 }
                })
        }
        dependency = [ordered]@{
            id = [string]$lock.dependencies.id
            sourceArchiveSha256 = [string]$lock.dependencies.sourceArchive.sha256
            contentArchiveSha256 = [string]$lock.dependencies.contentArchive.sha256
            licenseSha256 = [string]$lock.dependencies.licenseSha256
        }
        overlay = [ordered]@{
            path = [string]$lock.overlay.path
            normalizedTextSha256 = [string]$lock.overlay.normalizedTextSha256
            concernIds = [string[]]$lock.overlay.concernIds
        }
    }
    $descriptor = [ordered]@{
        schemaVersion = 1
        generatedUtc = $generatedText
        sourceBinding = $binding.sourceBinding
        canonicalLock = $binding.canonicalLock
        candidate = $binding.candidate
        ancestry = $binding.ancestry
        observation = $binding.observation
        review = $binding.review
        toolchain = $binding.toolchain
        dependency = $binding.dependency
        overlay = $binding.overlay
        inputDigest = Get-RSCanonicalJsonDigest -InputObject $binding
    }
    $descriptorPath = Join-Path $candidateRoot 'candidate-descriptor.json'
    $descriptorSchemaPath = Join-Path $repoRoot 'Specs\Terminal\GhosttyRuntimeCandidateDescriptor.schema.json'
    $descriptorSchemaErrors = @()
    $descriptorJson = ConvertTo-RSCanonicalJson -InputObject $descriptor
    if (-not (Test-Json -Json $descriptorJson -SchemaFile $descriptorSchemaPath -ErrorVariable descriptorSchemaErrors -ErrorAction SilentlyContinue)) {
        throw "Candidate descriptor does not satisfy its schema: $(@($descriptorSchemaErrors | ForEach-Object { $_.Exception.Message }) -join ' | ')"
    }
    Write-RSAtomicCanonicalJsonFile -LiteralPath $descriptorPath -InputObject $descriptor `
        -SchemaPath $descriptorSchemaPath -AllowedRoot $candidateRoot
    return [pscustomobject]@{
        CandidateCommit = $CandidateCommit
        Tree = $tree
        Version = $version
        ArchiveSha256 = $archiveSha256
        AheadCount = $aheadCount
        BehindCount = $behindCount
        FileCount = [int]$inventory.fileCount
        DeclarationBaseCount = [int]$abi.declarations.baseCount
        DeclarationCandidateCount = [int]$abi.declarations.candidateCount
        SecurityCommitCount = [int]$security.reviewedCommitCount
        ReviewPath = $reviewPath
        DescriptorPath = $descriptorPath
    }
}

Export-ModuleMember -Function @(
    'Get-RSGhosttyCandidatePathCategory',
    'Get-RSGhosttyCandidateChangeInventory',
    'Get-RSGhosttyCandidateAbiReview',
    'Get-RSGhosttyCandidateSecurityReview',
    'New-RSGhosttyRuntimeCandidateReview')
