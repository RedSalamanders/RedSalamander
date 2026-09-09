Set-StrictMode -Version Latest

$repoToolsRoot = Split-Path -Parent $PSScriptRoot
$validationModule = Join-Path $repoToolsRoot 'Modules\Testing\ValidationFingerprint.psm1'
$lifecycleModule = Join-Path $PSScriptRoot 'GhosttyRuntimeLifecycle.psm1'
$textIdentityModule = Join-Path $PSScriptRoot 'TerminalEngineTextIdentity.psm1'
Import-Module $validationModule -ErrorAction Stop
Import-Module $lifecycleModule -ErrorAction Stop
Import-Module $textIdentityModule -ErrorAction Stop

function Invoke-RSGhosttyCandidateOverlayGit {
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

function Assert-RSGhosttyCandidateOverlayPath {
    param(
        [Parameter(Mandatory = $true)][string]$CandidateRoot,
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Field,
        [switch]$Leaf,
        [switch]$Container
    )

    $candidateFull = [IO.Path]::GetFullPath($CandidateRoot).TrimEnd('\')
    $full = [IO.Path]::GetFullPath($Path)
    if (-not $full.StartsWith($candidateFull + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw "$Field must remain below the exact candidate root '$candidateFull'."
    }
    if ($Leaf -and -not (Test-Path -LiteralPath $full -PathType Leaf)) {
        throw "$Field is not a file: $full"
    }
    if ($Container -and -not (Test-Path -LiteralPath $full -PathType Container)) {
        throw "$Field is not a directory: $full"
    }
    return $full
}

function Get-RSGhosttyCandidateOverlayAbiReview {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$PristineSourcePath,
        [Parameter(Mandatory = $true)][string]$AppliedSourcePath
    )

    $pristineTerminalHeader = Get-Content -LiteralPath (Join-Path $PristineSourcePath 'include\ghostty\vt\terminal.h') -Raw
    $appliedTerminalHeader = Get-Content -LiteralPath (Join-Path $AppliedSourcePath 'include\ghostty\vt\terminal.h') -Raw
    $appliedTerminalZig = Get-Content -LiteralPath (Join-Path $AppliedSourcePath 'src\terminal\c\terminal.zig') -Raw
    $pristineKittyHeader = Get-Content -LiteralPath (Join-Path $PristineSourcePath 'include\ghostty\vt\kitty_graphics.h') -Raw
    $appliedKittyHeader = Get-Content -LiteralPath (Join-Path $AppliedSourcePath 'include\ghostty\vt\kitty_graphics.h') -Raw
    $appliedKittyZig = Get-Content -LiteralPath (Join-Path $AppliedSourcePath 'src\terminal\c\kitty_graphics.zig') -Raw

    $terminalOptionPattern = '(?m)^\s*GHOSTTY_TERMINAL_OPT_(?!MAX_VALUE)(?<name>[A-Z0-9_]+)\s*=\s*(?<value>\d+)\s*,'
    $pristineTerminalMatches = @([regex]::Matches($pristineTerminalHeader, $terminalOptionPattern))
    if ($pristineTerminalMatches.Count -eq 0) { throw 'Candidate public terminal option values were not found.' }
    $publicTerminalNames = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($match in $pristineTerminalMatches) {
        [void]$publicTerminalNames.Add($match.Groups['name'].Value)
    }
    $publicTerminalValues = @($pristineTerminalMatches | ForEach-Object { [int]$_.Groups['value'].Value })
    $publicTerminalMaximum = [int](($publicTerminalValues | Measure-Object -Maximum).Maximum)

    $privateTerminalOptions = @([regex]::Matches($appliedTerminalHeader, $terminalOptionPattern) |
        Where-Object { -not $publicTerminalNames.Contains($_.Groups['name'].Value) } |
        ForEach-Object {
            $suffix = $_.Groups['name'].Value
            $cName = "GHOSTTY_TERMINAL_OPT_$suffix"
            $zigName = $suffix.ToLowerInvariant()
            $cValue = [int]$_.Groups['value'].Value
            $zigMatch = [regex]::Match(
                $appliedTerminalZig,
                "(?m)^\s*$([regex]::Escape($zigName))\s*=\s*(?<value>\d+)\s*,")
            if (-not $zigMatch.Success) {
                throw "Private terminal option '$cName' is missing from the Zig candidate overlay."
            }
            $zigValue = [int]$zigMatch.Groups['value'].Value
            if ($cValue -ne $zigValue) {
                throw "Private terminal option C/Zig mismatch for '$cName': $cValue/$zigValue."
            }
            if ($cValue -le $publicTerminalMaximum) {
                throw "Private terminal option '$cName' value $cValue collides with the public range through $publicTerminalMaximum."
            }
            [ordered]@{ cName = $cName; zigName = $zigName; cValue = $cValue; zigValue = $zigValue }
        } | Sort-Object -Property @{ Expression = { $_.cValue } }, @{ Expression = { $_.cName } })
    if ($privateTerminalOptions.Count -eq 0) {
        throw 'The candidate overlay does not add a private terminal option.'
    }
    if (@($privateTerminalOptions.cName | Sort-Object -Unique -CaseSensitive).Count -ne $privateTerminalOptions.Count -or
        @($privateTerminalOptions.cValue | Sort-Object -Unique).Count -ne $privateTerminalOptions.Count) {
        throw 'Private terminal option names and values must be unique.'
    }

    $publicKittyValues = @([regex]::Matches(
            $pristineKittyHeader,
            '(?m)^\s*GHOSTTY_KITTY_GRAPHICS_DATA_(?!MAX_VALUE)(?<name>[A-Z0-9_]+)\s*=\s*(?<value>\d+)\s*,'
        ) | ForEach-Object { [int]$_.Groups['value'].Value })
    if ($publicKittyValues.Count -eq 0) { throw 'Candidate public Kitty data values were not found.' }
    $publicKittyMaximum = [int](($publicKittyValues | Measure-Object -Maximum).Maximum)
    $privateDataNames = [ordered]@{
        'GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_COUNT' = 'placement_count'
        'GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_ALLOCATED_BYTES' = 'placement_allocated_bytes'
        'GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_COUNT_LIMIT' = 'placement_count_limit'
        'GHOSTTY_KITTY_GRAPHICS_DATA_PLACEMENT_BYTES_LIMIT' = 'placement_bytes_limit'
    }
    $privateData = @($privateDataNames.GetEnumerator() | ForEach-Object {
            $cMatch = [regex]::Match($appliedKittyHeader, "(?m)^\s*$([regex]::Escape($_.Key))\s*=\s*(?<value>\d+)\s*,")
            $zigMatch = [regex]::Match($appliedKittyZig, "(?m)^\s*$([regex]::Escape($_.Value))\s*=\s*(?<value>\d+)\s*,")
            if (-not $cMatch.Success -or -not $zigMatch.Success) {
                throw "Private Kitty data value '$($_.Key)' is missing from the C or Zig candidate overlay."
            }
            $cValue = [int]$cMatch.Groups['value'].Value
            $zigValue = [int]$zigMatch.Groups['value'].Value
            if ($cValue -ne $zigValue) { throw "Private Kitty data C/Zig mismatch for '$($_.Key)': $cValue/$zigValue." }
            if ($cValue -le $publicKittyMaximum) { throw "Private Kitty data value $cValue collides with the public range through $publicKittyMaximum." }
            [ordered]@{ cName = [string]$_.Key; zigName = [string]$_.Value; cValue = $cValue; zigValue = $zigValue }
        })
    if (@($privateData.cValue | Sort-Object -Unique).Count -ne $privateData.Count) {
        throw 'Private Kitty data values are not unique.'
    }

    return [ordered]@{
        publicTerminalOptionMaximum = $publicTerminalMaximum
        privateTerminalOptions = $privateTerminalOptions
        publicKittyDataMaximum = $publicKittyMaximum
        privateKittyDataValues = $privateData
    }
}

function New-RSGhosttyRuntimeCandidateOverlayReview {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RepoRoot,
        [Parameter(Mandatory = $true)][string]$CandidateCommit,
        [Parameter(Mandatory = $true)][string]$PristineSourcePath,
        [Parameter(Mandatory = $true)][string]$AppliedSourcePath,
        [Parameter(Mandatory = $true)][string]$PatchPath,
        [Parameter(Mandatory = $true)][string]$DescriptorPath,
        [Parameter(Mandatory = $true)][string]$ConcernReviewPath,
        [DateTime]$GeneratedUtc = [DateTime]::UtcNow
    )

    if ($CandidateCommit -cnotmatch '^[0-9a-f]{40}$') { throw 'CandidateCommit must be a lowercase full SHA.' }
    $repoRoot = [IO.Path]::GetFullPath($RepoRoot)
    $candidateRoot = Join-Path $repoRoot ".build\TerminalEngineUpgrade\$CandidateCommit"
    $pristineSourcePath = Assert-RSGhosttyCandidateOverlayPath -CandidateRoot $candidateRoot -Path $PristineSourcePath -Field 'PristineSourcePath' -Container
    $appliedSourcePath = Assert-RSGhosttyCandidateOverlayPath -CandidateRoot $candidateRoot -Path $AppliedSourcePath -Field 'AppliedSourcePath' -Container
    $patchPath = Assert-RSGhosttyCandidateOverlayPath -CandidateRoot $candidateRoot -Path $PatchPath -Field 'PatchPath' -Leaf
    $descriptorPath = Assert-RSGhosttyCandidateOverlayPath -CandidateRoot $candidateRoot -Path $DescriptorPath -Field 'DescriptorPath' -Leaf
    $concernReviewPath = Assert-RSGhosttyCandidateOverlayPath -CandidateRoot $candidateRoot -Path $ConcernReviewPath -Field 'ConcernReviewPath' -Leaf

    $pristineHead = ((Invoke-RSGhosttyCandidateOverlayGit -RepositoryPath $pristineSourcePath -Arguments @('rev-parse', 'HEAD') -Context 'Resolve pristine candidate HEAD').Lines | Select-Object -First 1)
    $appliedHead = ((Invoke-RSGhosttyCandidateOverlayGit -RepositoryPath $appliedSourcePath -Arguments @('rev-parse', 'HEAD') -Context 'Resolve applied candidate HEAD').Lines | Select-Object -First 1)
    if ($pristineHead -cne $CandidateCommit -or $appliedHead -cne $CandidateCommit) {
        throw "Candidate overlay sources must both be based at exact commit $CandidateCommit."
    }
    $pristineStatus = (Invoke-RSGhosttyCandidateOverlayGit -RepositoryPath $pristineSourcePath -Arguments @('status', '--porcelain=v1', '--untracked-files=no') -Context 'Inspect pristine candidate status').Lines
    if ($pristineStatus.Count -ne 0) { throw "Pristine candidate source has tracked changes: $($pristineStatus -join '; ')" }
    $candidateTree = ((Invoke-RSGhosttyCandidateOverlayGit -RepositoryPath $pristineSourcePath -Arguments @('rev-parse', 'HEAD^{tree}') -Context 'Resolve candidate tree').Lines | Select-Object -First 1)

    $applyCheck = Invoke-RSGhosttyCandidateOverlayGit -RepositoryPath $pristineSourcePath -Arguments @('apply', '--check', $patchPath) -Context 'Candidate overlay apply check' -AllowFailure
    if ($applyCheck.ExitCode -ne 0) { throw "Candidate overlay does not apply cleanly: $($applyCheck.Lines -join "`n")" }
    $diffCheck = Invoke-RSGhosttyCandidateOverlayGit -RepositoryPath $appliedSourcePath -Arguments @('diff', '--check', 'HEAD') -Context 'Applied candidate diff check' -AllowFailure
    if ($diffCheck.ExitCode -ne 0) { throw "Applied candidate diff check failed: $($diffCheck.Lines -join "`n")" }

    $numstatLines = (Invoke-RSGhosttyCandidateOverlayGit -RepositoryPath $appliedSourcePath -Arguments @('diff', '--numstat', 'HEAD', '--') -Context 'Candidate overlay line inventory').Lines
    $files = @($numstatLines | ForEach-Object {
            $parts = $_ -split "`t"
            if ($parts.Count -ne 3 -or $parts[0] -cnotmatch '^\d+$' -or $parts[1] -cnotmatch '^\d+$') {
                throw "Malformed candidate overlay numstat line: $_"
            }
            [ordered]@{ path = [string]$parts[2]; insertions = [int]$parts[0]; deletions = [int]$parts[1] }
        } | Sort-Object { $_.path })
    $insertions = [int](($files.insertions | Measure-Object -Sum).Sum)
    $deletions = [int](($files.deletions | Measure-Object -Sum).Sum)
    if ($files.Count -gt 16 -or ($insertions + $deletions) -gt 2500) {
        throw "Candidate overlay exceeds review limits: $($files.Count) files, $insertions insertions, $deletions deletions."
    }

    $patchIdentity = Get-RSTerminalLfTextIdentity -LiteralPath $patchPath
    $diffArguments = @('diff', '--no-ext-diff', '--full-index', 'HEAD', '--') + [string[]]@($files.path)
    $regeneratedLines = (Invoke-RSGhosttyCandidateOverlayGit -RepositoryPath $appliedSourcePath -Arguments $diffArguments -Context 'Regenerate candidate overlay').Lines
    $regeneratedIdentity = Get-RSTerminalLfStringIdentity -Text (($regeneratedLines -join "`n") + "`n")
    if ($regeneratedIdentity.Sha256 -cne $patchIdentity.Sha256) {
        throw "Regenerated candidate overlay SHA-256 '$($regeneratedIdentity.Sha256)' does not match '$($patchIdentity.Sha256)'."
    }

    $lockPath = Join-Path $repoRoot 'External\TerminalEngine\GhosttyRuntimeLock.v1.json'
    $lockSchemaPath = Join-Path $repoRoot 'Specs\Terminal\GhosttyRuntimeLock.schema.json'
    $lock = Read-RSGhosttyRuntimeLock -RepoRoot $repoRoot -LiteralPath $lockPath -SchemaPath $lockSchemaPath
    $reviewInput = Get-Content -LiteralPath $concernReviewPath -Raw | ConvertFrom-Json -Depth 100 -DateKind String
    if ([string]$reviewInput.candidateCommit -cne $CandidateCommit) { throw 'Concern review candidateCommit does not match CandidateCommit.' }
    $concerns = @($reviewInput.concerns)
    if ($concerns.Count -lt 1 -or $concerns.Count -gt 5) { throw 'Concern review must contain one through five candidate concerns.' }
    $concernIds = [string[]]@($concerns | ForEach-Object { [string]$_.id })
    if (@($concernIds | Sort-Object -Unique -CaseSensitive).Count -ne $concernIds.Count) { throw 'Candidate concern IDs must be unique.' }
    $sortedConcernIds = [string[]]@($concernIds | Sort-Object -CaseSensitive)
    if (($concernIds -join "`n") -cne ($sortedConcernIds -join "`n")) { throw 'Candidate concern IDs must be ordinal sorted.' }
    foreach ($concern in $concerns) {
        if ([string]::IsNullOrWhiteSpace([string]$concern.summary) -or @($concern.sourceEvidence).Count -eq 0 -or @($concern.focusedTests).Count -eq 0) {
            throw "Candidate concern '$($concern.id)' requires summary, sourceEvidence, and focusedTests."
        }
    }

    $dispositions = @($reviewInput.currentConcernDispositions)
    $currentIds = [string[]]@($lock.overlay.concernIds)
    $dispositionIds = [string[]]@($dispositions | ForEach-Object { [string]$_.id })
    $missingDispositions = @($currentIds | Where-Object { $_ -cnotin $dispositionIds })
    $extraDispositions = @($dispositionIds | Where-Object { $_ -cnotin $currentIds })
    if ($dispositions.Count -ne $currentIds.Count -or $missingDispositions.Count -ne 0 -or $extraDispositions.Count -ne 0) {
        throw "Current concern dispositions must cover the lock exactly; missing=$($missingDispositions -join ',') extra=$($extraDispositions -join ',')."
    }
    foreach ($disposition in $dispositions) {
        $mappedIds = [string[]]@($disposition.candidateConcernIds)
        if (@($mappedIds | Where-Object { $_ -cnotin $concernIds }).Count -ne 0) { throw "Disposition '$($disposition.id)' references an unknown candidate concern." }
        if ([string]$disposition.disposition -ceq 'removed' -and $mappedIds.Count -ne 0) { throw "Removed concern '$($disposition.id)' must not map to a candidate concern." }
        if ([string]$disposition.disposition -cne 'removed' -and $mappedIds.Count -eq 0) { throw "Retained/replaced concern '$($disposition.id)' must map to a candidate concern." }
        if ([string]::IsNullOrWhiteSpace([string]$disposition.rationale)) { throw "Disposition '$($disposition.id)' requires a rationale." }
    }

    $abi = Get-RSGhosttyCandidateOverlayAbiReview -PristineSourcePath $pristineSourcePath -AppliedSourcePath $appliedSourcePath
    $descriptorSchemaPath = Join-Path $repoRoot 'Specs\Terminal\GhosttyRuntimeCandidateDescriptor.schema.json'
    $descriptorBackupPath = Join-Path $candidateRoot 'candidate-descriptor.p4.json'
    if (-not (Test-Path -LiteralPath $descriptorBackupPath -PathType Leaf)) {
        $initialDescriptor = Get-Content -LiteralPath $descriptorPath -Raw | ConvertFrom-Json -Depth 100 -DateKind String
        Write-RSAtomicCanonicalJsonFile -LiteralPath $descriptorBackupPath -InputObject $initialDescriptor -SchemaPath $descriptorSchemaPath -AllowedRoot $candidateRoot
    }
    $baseDescriptor = Get-Content -LiteralPath $descriptorBackupPath -Raw | ConvertFrom-Json -Depth 100 -DateKind String
    $baseDescriptorSha256 = (Get-FileHash -LiteralPath $descriptorBackupPath -Algorithm SHA256).Hash.ToLowerInvariant()

    $patchRelative = [IO.Path]::GetRelativePath($repoRoot, $patchPath).Replace('\', '/')
    $review = [ordered]@{
        schemaVersion = 2
        generatedUtc = $GeneratedUtc.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ', [Globalization.CultureInfo]::InvariantCulture)
        candidateCommit = $CandidateCommit
        candidateTree = $candidateTree
        baseDescriptorSha256 = $baseDescriptorSha256
        patch = [ordered]@{
            path = $patchRelative
            normalizedTextSha256 = $patchIdentity.Sha256
            fileCount = $files.Count
            insertions = $insertions
            deletions = $deletions
            files = $files
        }
        concerns = @($concerns | ForEach-Object {
                [ordered]@{ id = [string]$_.id; summary = [string]$_.summary; sourceEvidence = [string[]]$_.sourceEvidence; focusedTests = [string[]]$_.focusedTests }
            })
        currentConcernDispositions = @($dispositions | ForEach-Object {
                [ordered]@{ id = [string]$_.id; disposition = [string]$_.disposition; candidateConcernIds = [string[]]$_.candidateConcernIds; rationale = [string]$_.rationale }
            })
        abi = $abi
        verification = [ordered]@{
            pristineSourceClean = $true
            applyCheckPassed = $true
            appliedDiffCheckPassed = $true
            regeneratedPatchSha256 = $regeneratedIdentity.Sha256
            regeneratedPatchMatches = $true
        }
    }
    $reviewSchemaPath = Join-Path $repoRoot 'Specs\Terminal\GhosttyRuntimeCandidateOverlayReview.schema.json'
    $reviewPath = Join-Path $candidateRoot 'candidate-overlay-review.json'
    Write-RSAtomicCanonicalJsonFile -LiteralPath $reviewPath -InputObject $review -SchemaPath $reviewSchemaPath -AllowedRoot $candidateRoot

    $binding = [ordered]@{
        sourceBinding = $baseDescriptor.sourceBinding
        canonicalLock = $baseDescriptor.canonicalLock
        candidate = $baseDescriptor.candidate
        ancestry = $baseDescriptor.ancestry
        observation = $baseDescriptor.observation
        review = $baseDescriptor.review
        toolchain = $baseDescriptor.toolchain
        dependency = $baseDescriptor.dependency
        overlay = [ordered]@{ path = $patchRelative; normalizedTextSha256 = $patchIdentity.Sha256; concernIds = $concernIds }
    }
    $updatedDescriptor = [ordered]@{
        schemaVersion = 1
        generatedUtc = $review.generatedUtc
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
    Write-RSAtomicCanonicalJsonFile -LiteralPath $descriptorPath -InputObject $updatedDescriptor -SchemaPath $descriptorSchemaPath -AllowedRoot $candidateRoot

    return [pscustomobject]@{
        CandidateCommit = $CandidateCommit
        PatchSha256 = $patchIdentity.Sha256
        FileCount = $files.Count
        Insertions = $insertions
        Deletions = $deletions
        ConcernCount = $concernIds.Count
        ReviewPath = $reviewPath
        ReviewSha256 = (Get-FileHash -LiteralPath $reviewPath -Algorithm SHA256).Hash.ToLowerInvariant()
        DescriptorPath = $descriptorPath
        DescriptorSha256 = (Get-FileHash -LiteralPath $descriptorPath -Algorithm SHA256).Hash.ToLowerInvariant()
        InputDigest = [string]$updatedDescriptor.inputDigest
    }
}

Export-ModuleMember -Function @(
    'Get-RSGhosttyCandidateOverlayAbiReview',
    'New-RSGhosttyRuntimeCandidateOverlayReview')
