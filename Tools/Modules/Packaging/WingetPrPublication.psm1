# Private packaging module; import through its reviewed callers.
Set-StrictMode -Version Latest

function Get-RSSha256Hex {
    param(
        [Parameter(Mandatory = $true)]
        [byte[]]$Bytes
    )

    return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($Bytes)).ToLowerInvariant()
}

function ConvertTo-RSWingetCanonicalManifest {
    param(
        [AllowEmptyString()]
        [string]$Content = ''
    )

    # `wingetcreate submit` re-serializes manifests before committing them: it prepends a
    # "# Created using wingetcreate" comment, writes CRLF, re-indents sequences, and reorders keys.
    # Byte-exact hashes therefore never match. This form keeps every scalar line (values, list
    # items, block-scalar lines) and discards only comments, blank lines, indentation, list markers,
    # line endings, and order, so the generated and committed manifests hash identically while any
    # changed URL, hash, version, alias, or date still differs.
    $lines = @()
    foreach ($rawLine in ($Content -replace "`r`n", "`n" -replace "`r", "`n") -split "`n") {
        $line = $rawLine.TrimEnd()
        if ($line.TrimStart().StartsWith('#')) {
            continue
        }
        $line = [regex]::Replace($line, '^\s*(?:-\s+)?', '')
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }
        $lines += $line
    }

    return (@($lines | Sort-Object -CaseSensitive) -join "`n")
}

function Get-RSWingetManifestContentHash {
    param(
        [Parameter(Mandatory = $true)]
        [byte[]]$Bytes
    )

    $content = [Text.UTF8Encoding]::new($false).GetString($Bytes)
    if ($content.Length -gt 0 -and $content[0] -eq [char]0xFEFF) {
        $content = $content.Substring(1)
    }
    $canonical = ConvertTo-RSWingetCanonicalManifest -Content $content
    return Get-RSSha256Hex -Bytes ([Text.UTF8Encoding]::new($false).GetBytes($canonical))
}

function New-RSWingetPublicationIdentity {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PackageIdentifier,

        [Parameter(Mandatory = $true)]
        [string]$Version,

        [Parameter(Mandatory = $true)]
        [string]$ManifestRoot,

        [Parameter(Mandatory = $true)]
        [string]$PullRequestTitle
    )

    if ($PackageIdentifier -notmatch '^[A-Za-z0-9][A-Za-z0-9.-]{0,127}$') {
        throw "Invalid Winget package identifier '$PackageIdentifier'."
    }
    if ($Version -notmatch '^\d+\.\d+\.\d+$') {
        throw "Invalid Winget package version '$Version'."
    }
    if (-not (Test-Path -LiteralPath $ManifestRoot -PathType Container)) {
        throw "Winget manifest root not found: $ManifestRoot"
    }

    $requiredNames = @(
        "$PackageIdentifier.installer.yaml",
        "$PackageIdentifier.locale.en-US.yaml",
        "$PackageIdentifier.yaml"
    )
    $manifestFiles = @(Get-ChildItem -LiteralPath $ManifestRoot -File)
    $actualNames = @($manifestFiles.Name | Sort-Object -CaseSensitive)
    $expectedNames = @($requiredNames | Sort-Object -CaseSensitive)
    if (($actualNames -join "`n") -cne ($expectedNames -join "`n")) {
        throw "Winget manifest root must contain exactly: $($expectedNames -join ', ')."
    }

    $packageSegments = @($PackageIdentifier -split '\.')
    $firstLetter = $packageSegments[0].Substring(0, 1).ToLowerInvariant()
    $manifestPrefix = "manifests/$firstLetter/$($packageSegments -join '/')/$Version"
    $manifestHashes = [ordered]@{}
    foreach ($fileName in $expectedNames) {
        $localPath = Join-Path $ManifestRoot $fileName
        $content = [IO.File]::ReadAllText($localPath)
        if ($content -notmatch "(?m)^PackageIdentifier:\s*$([regex]::Escape($PackageIdentifier))\s*$") {
            throw "Winget manifest '$fileName' does not declare PackageIdentifier '$PackageIdentifier'."
        }
        if ($content -notmatch "(?m)^PackageVersion:\s*$([regex]::Escape($Version))\s*$") {
            throw "Winget manifest '$fileName' does not declare PackageVersion '$Version'."
        }

        $remotePath = "$manifestPrefix/$fileName"
        $manifestHashes[$remotePath] = Get-RSWingetManifestContentHash -Bytes ([IO.File]::ReadAllBytes($localPath))
    }

    $identityLines = @(
        "package=$PackageIdentifier",
        "version=$Version"
    )
    foreach ($entry in $manifestHashes.GetEnumerator()) {
        $identityLines += "$($entry.Key)=$($entry.Value)"
    }
    $identityBytes = [Text.UTF8Encoding]::new($false).GetBytes($identityLines -join "`n")
    $publicationId = Get-RSSha256Hex -Bytes $identityBytes
    $initialTitle = "$PullRequestTitle [RS-WINGET:$publicationId]"
    $bodyMarker = "<!-- RS-WINGET-PUBLICATION:$publicationId -->"
    $guidPattern = '[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}'
    $branchPattern = '^' + [regex]::Escape("$PackageIdentifier-$Version-") + $guidPattern + '$'

    return [pscustomobject]@{
        PackageIdentifier = $PackageIdentifier
        Version = $Version
        PullRequestTitle = $PullRequestTitle
        InitialTitle = $initialTitle
        BodyMarker = $bodyMarker
        PublicationId = $publicationId
        ManifestPrefix = $manifestPrefix
        ManifestHashes = $manifestHashes
        BranchPattern = $branchPattern
    }
}

function Get-RSRepositoryOperation {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Repository,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (-not $Repository.ContainsKey($Name) -or $Repository[$Name] -isnot [scriptblock]) {
        throw "Winget publication repository operation '$Name' is required."
    }

    return $Repository[$Name]
}

function Get-RSPagedResults {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Operation,

        [Parameter(Mandatory = $true)]
        [int]$MaximumPages,

        [int]$PageSize = 50,

        [object[]]$LeadingArguments = @()
    )

    $results = @()
    for ($page = 1; $page -le $MaximumPages; $page++) {
        $arguments = @($LeadingArguments) + @($page, $PageSize)
        # Invoke-RestMethod emits a JSON array as one object, so enumerate the page before counting it.
        $pageResults = @(& $Operation @arguments | ForEach-Object { $_ })
        $results += $pageResults
        if ($pageResults.Count -lt $PageSize) {
            break
        }
    }

    return @($results)
}

function Test-RSStringSetEqual {
    param(
        [string[]]$Left,
        [string[]]$Right
    )

    $leftSorted = @($Left | Sort-Object -CaseSensitive -Unique)
    $rightSorted = @($Right | Sort-Object -CaseSensitive -Unique)
    return ($leftSorted.Count -eq $Left.Count -and
        $rightSorted.Count -eq $Right.Count -and
        ($leftSorted -join "`n") -ceq ($rightSorted -join "`n"))
}

function Get-RSWingetObjectString {
    param(
        [object]$Object,

        [Parameter(Mandatory = $true)]
        [string[]]$Path
    )

    $current = $Object
    foreach ($segment in $Path) {
        if ($null -eq $current) {
            return ''
        }
        $property = $current.PSObject.Properties[$segment]
        if ($null -eq $property) {
            return ''
        }
        $current = $property.Value
    }
    if ($null -eq $current) {
        return ''
    }
    return [string]$current
}

function Test-RSWingetPullRequestCandidate {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$PullRequest,

        [Parameter(Mandatory = $true)]
        [pscustomobject]$Identity,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedAuthor
    )

    # Ownership and conflict checks need the changed-file list, and that list costs one request per
    # pull request. Only the token owner's pull requests and those naming the package in their title,
    # body, or head branch can be this publication or an exact-version conflict; every WingetCreate,
    # Komac, and bot submission carries the package identifier in at least its head branch.
    $author = Get-RSWingetObjectString -Object $PullRequest -Path @('user', 'login')
    if ([string]::Equals($author, $ExpectedAuthor, [StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }
    foreach ($path in @(@('title'), @('body'), @('head', 'ref'))) {
        $text = Get-RSWingetObjectString -Object $PullRequest -Path $path
        if ($text.IndexOf([string]$Identity.PackageIdentifier, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
            return $true
        }
    }
    return $false
}

function Get-RSWingetPublicationState {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Identity,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedAuthor,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedHeadRepository,

        [Parameter(Mandatory = $true)]
        [hashtable]$Repository,

        [int]$MaximumPullRequestPages = 2,

        [int]$MaximumFilePages = 2
    )

    $listPullRequests = Get-RSRepositoryOperation -Repository $Repository -Name 'ListPullRequests'
    $listPullRequestFiles = Get-RSRepositoryOperation -Repository $Repository -Name 'ListPullRequestFiles'
    $getManifestContentHash = Get-RSRepositoryOperation -Repository $Repository -Name 'GetManifestContentHash'

    try {
        $pullRequests = @(Get-RSPagedResults `
                -Operation $listPullRequests `
                -MaximumPages $MaximumPullRequestPages)
    }
    catch [System.Exception] {
        return [pscustomobject]@{ Kind = 'Pending'; Reason = "The upstream pull-request list is not yet readable: $($_.Exception.Message)"; PullRequest = $null }
    }

    $expectedPaths = @($Identity.ManifestHashes.Keys)
    $candidates = @()
    $scanIncomplete = $false
    $scanFailure = $null
    foreach ($pullRequest in $pullRequests) {
        $hasMarker = (($pullRequest.title -ceq $Identity.InitialTitle) -or
            ([string]$pullRequest.body).Contains($Identity.BodyMarker, [StringComparison]::Ordinal))
        if (-not $hasMarker -and
            -not (Test-RSWingetPullRequestCandidate -PullRequest $pullRequest -Identity $Identity -ExpectedAuthor $ExpectedAuthor)) {
            continue
        }
        try {
            $files = @(Get-RSPagedResults `
                    -Operation $listPullRequestFiles `
                    -MaximumPages $MaximumFilePages `
                    -LeadingArguments @([int]$pullRequest.number))
        }
        catch [System.Exception] {
            $scanIncomplete = $true
            if ($null -eq $scanFailure) {
                $scanFailure = "pull request #$($pullRequest.number): $($_.Exception.Message)"
            }
            if ($hasMarker) {
                $candidates += [pscustomobject]@{
                    PullRequest = $pullRequest
                    Files = @()
                    FileListPending = $true
                    FileListFailure = [string]$_.Exception.Message
                    HasMarker = $true
                }
            }
            continue
        }

        $fileNames = @($files | ForEach-Object { [string]$_.filename })
        $hasManifestPath = @($fileNames | Where-Object {
                $_ -ceq $Identity.ManifestPrefix -or
                $_.StartsWith("$($Identity.ManifestPrefix)/", [StringComparison]::Ordinal)
            }).Count -gt 0
        if ($hasMarker -or $hasManifestPath) {
            $candidates += [pscustomobject]@{
                PullRequest = $pullRequest
                Files = $files
                FileListPending = $false
                FileListFailure = $null
                HasMarker = $hasMarker
            }
        }
    }

    if ($candidates.Count -gt 1) {
        return [pscustomobject]@{ Kind = 'Conflict'; Reason = "Multiple pull requests target $($Identity.PackageIdentifier) $($Identity.Version)."; PullRequest = $null }
    }
    if ($candidates.Count -eq 0) {
        if ($scanIncomplete) {
            return [pscustomobject]@{ Kind = 'Pending'; Reason = "The exact-version pull-request scan is incomplete: $scanFailure"; PullRequest = $null }
        }
        return [pscustomobject]@{ Kind = 'None'; Reason = 'No exact-version pull request exists.'; PullRequest = $null }
    }

    $candidate = $candidates[0]
    $pullRequest = $candidate.PullRequest
    if ([string]$pullRequest.state -cne 'open') {
        return [pscustomobject]@{ Kind = 'Conflict'; Reason = "Pull request #$($pullRequest.number) is not open."; PullRequest = $pullRequest }
    }

    $authorMatches = [string]::Equals([string]$pullRequest.user.login, $ExpectedAuthor, [StringComparison]::OrdinalIgnoreCase)
    $repositoryMatches = [string]::Equals([string]$pullRequest.head.repo.full_name, $ExpectedHeadRepository, [StringComparison]::OrdinalIgnoreCase)
    $branchMatches = ([string]$pullRequest.head.ref -cmatch $Identity.BranchPattern)
    if (-not $authorMatches -or -not $repositoryMatches -or -not $branchMatches -or -not $candidate.HasMarker) {
        return [pscustomobject]@{ Kind = 'Conflict'; Reason = "Pull request #$($pullRequest.number) is not owned by this publication identity."; PullRequest = $pullRequest }
    }
    if ($candidate.FileListPending) {
        return [pscustomobject]@{ Kind = 'Pending'; Reason = "Pull request #$($pullRequest.number) files are not yet visible: $($candidate.FileListFailure)"; PullRequest = $pullRequest }
    }

    $fileNames = @($candidate.Files | ForEach-Object { [string]$_.filename })
    if (-not (Test-RSStringSetEqual -Left $fileNames -Right $expectedPaths)) {
        return [pscustomobject]@{ Kind = 'Conflict'; Reason = "Pull request #$($pullRequest.number) does not contain the exact manifest file set."; PullRequest = $pullRequest }
    }
    foreach ($file in $candidate.Files) {
        if ([string]$file.status -notin @('added', 'modified')) {
            return [pscustomobject]@{ Kind = 'Conflict'; Reason = "Pull request #$($pullRequest.number) has an invalid manifest file status."; PullRequest = $pullRequest }
        }
    }

    foreach ($path in $expectedPaths) {
        try {
            $actualHash = [string](& $getManifestContentHash `
                    ([string]$pullRequest.head.repo.full_name) `
                    $path `
                    ([string]$pullRequest.head.sha))
        }
        catch [System.Exception] {
            # Name the path and the upstream error: a 404 that never clears and a 403 rate limit need
            # different fixes, and the bounded-retry failure is the only place this reason surfaces.
            return [pscustomobject]@{ Kind = 'Pending'; Reason = "Pull request #$($pullRequest.number) manifest content is not yet visible: '$path' at $([string]$pullRequest.head.sha): $($_.Exception.Message)"; PullRequest = $pullRequest }
        }
        if ($actualHash -cne [string]$Identity.ManifestHashes[$path]) {
            return [pscustomobject]@{ Kind = 'Conflict'; Reason = "Pull request #$($pullRequest.number) canonical manifest content does not match '$path'."; PullRequest = $pullRequest }
        }
    }

    return [pscustomobject]@{ Kind = 'Owned'; Reason = "Pull request #$($pullRequest.number) is an exact automation-owned publication."; PullRequest = $pullRequest }
}

function Get-RSWingetSubmitDiagnostics {
    param(
        [object]$SubmitResult,

        [ValidateRange(1, 200)]
        [int]$MaximumLines = 40,

        [ValidateRange(80, 4000)]
        [int]$MaximumLineLength = 500
    )

    if ($null -eq $SubmitResult) {
        return @()
    }
    $outputProperty = $SubmitResult.PSObject.Properties['Output']
    if ($null -eq $outputProperty -or $null -eq $outputProperty.Value) {
        return @()
    }

    # WingetCreate output is diagnostic only: it is never parsed for the PR URL, and any
    # token-shaped value is redacted before the text can reach a log or an exception.
    $tokenShape = '\b(?:gh[pousr]_[A-Za-z0-9]{20,}|github_pat_[A-Za-z0-9_]{20,})\b'
    $credentialHeaderShape = '(?i)\b(bearer|token|password)(\s*[:=]\s*|\s+)[A-Za-z0-9._\-/+=]{16,}'
    $lines = @()
    foreach ($item in @($outputProperty.Value)) {
        $line = ([string]$item).Trim()
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }
        $line = [regex]::Replace($line, $tokenShape, '***')
        $line = [regex]::Replace($line, $credentialHeaderShape, '$1$2***')
        if ($line.Length -gt $MaximumLineLength) {
            $line = $line.Substring(0, $MaximumLineLength) + ' ...'
        }
        $lines += $line
    }

    if ($lines.Count -gt $MaximumLines) {
        $omitted = $lines.Count - $MaximumLines
        $lines = @("... ($omitted earlier line(s) omitted)") + @($lines[$omitted..($lines.Count - 1)])
    }

    return @($lines)
}

function Set-RSWingetPullRequestMetadata {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Identity,

        [Parameter(Mandatory = $true)]
        [pscustomobject]$PullRequest,

        [Parameter(Mandatory = $true)]
        [string]$Body,

        [Parameter(Mandatory = $true)]
        [hashtable]$Repository
    )

    $patchPullRequest = Get-RSRepositoryOperation -Repository $Repository -Name 'PatchPullRequest'
    $desiredBody = "$($Identity.BodyMarker)`n`n$($Body.Trim())"
    $updated = & $patchPullRequest ([int]$PullRequest.number) $Identity.PullRequestTitle $desiredBody
    if ($null -eq $updated -or
        [int]$updated.number -ne [int]$PullRequest.number -or
        [string]$updated.state -cne 'open' -or
        [string]$updated.title -cne $Identity.PullRequestTitle -or
        -not ([string]$updated.body).Contains($Identity.BodyMarker, [StringComparison]::Ordinal) -or
        [string]::IsNullOrWhiteSpace([string]$updated.html_url)) {
        throw "Winget pull request #$($PullRequest.number) metadata update could not be verified."
    }

    $expectedUrl = "https://github.com/microsoft/winget-pkgs/pull/$($PullRequest.number)"
    if ([string]$updated.html_url -cne $expectedUrl) {
        throw "Winget pull request #$($PullRequest.number) returned an unexpected URL."
    }

    return [pscustomobject]@{
        PullRequestNumber = [int]$updated.number
        Url = [string]$updated.html_url
    }
}

function Invoke-RSWingetPublication {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Identity,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedAuthor,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedHeadRepository,

        [Parameter(Mandatory = $true)]
        [string]$PullRequestBody,

        [Parameter(Mandatory = $true)]
        [hashtable]$Repository,

        [Parameter(Mandatory = $true)]
        [scriptblock]$Submit,

        # A state check now costs a few requests, so the waits carry the eventual-consistency window
        # (15, 30, 45, 60, 60 s ≈ 3.5 min over six attempts) instead of the scan's own duration.
        [scriptblock]$Wait = { param([int]$Attempt) Start-Sleep -Seconds ([Math]::Min(15 * $Attempt, 60)) },

        [ValidateRange(1, 20)]
        [int]$MaximumVisibilityAttempts = 6
    )

    $state = Get-RSWingetPublicationState `
        -Identity $Identity `
        -ExpectedAuthor $ExpectedAuthor `
        -ExpectedHeadRepository $ExpectedHeadRepository `
        -Repository $Repository
    if ($state.Kind -ceq 'Conflict') {
        throw $state.Reason
    }
    if ($state.Kind -ceq 'Owned') {
        $result = Set-RSWingetPullRequestMetadata `
            -Identity $Identity `
            -PullRequest $state.PullRequest `
            -Body $PullRequestBody `
            -Repository $Repository
        return [pscustomobject]@{
            PullRequestNumber = $result.PullRequestNumber
            Url = $result.Url
            Submitted = $false
        }
    }

    $submitted = $false
    $submitFailure = $null
    $submitDiagnostics = @()
    if ($state.Kind -ceq 'None') {
        $submitted = $true
        try {
            $submitResult = & $Submit $Identity.InitialTitle $Identity $Repository
            $submitDiagnostics = @(Get-RSWingetSubmitDiagnostics -SubmitResult $submitResult)
            if ($null -eq $submitResult -or [int]$submitResult.ExitCode -ne 0) {
                $exitCode = if ($null -eq $submitResult) { 'unknown' } else { [string]$submitResult.ExitCode }
                $submitFailure = "wingetcreate submit failed with exit code $exitCode."
            }
        }
        catch [System.Exception] {
            $submitFailure = "wingetcreate submit did not return cleanly: $($_.Exception.Message)"
        }
    }

    for ($attempt = 1; $attempt -le $MaximumVisibilityAttempts; $attempt++) {
        $state = Get-RSWingetPublicationState `
            -Identity $Identity `
            -ExpectedAuthor $ExpectedAuthor `
            -ExpectedHeadRepository $ExpectedHeadRepository `
            -Repository $Repository
        if ($state.Kind -ceq 'Conflict') {
            throw $state.Reason
        }
        if ($state.Kind -ceq 'Owned') {
            $result = Set-RSWingetPullRequestMetadata `
                -Identity $Identity `
                -PullRequest $state.PullRequest `
                -Body $PullRequestBody `
                -Repository $Repository
            return [pscustomobject]@{
                PullRequestNumber = $result.PullRequestNumber
                Url = $result.Url
                Submitted = $submitted
            }
        }
        if ($attempt -lt $MaximumVisibilityAttempts) {
            & $Wait $attempt $Repository | Out-Null
        }
    }

    $failureSuffix = if ([string]::IsNullOrWhiteSpace([string]$submitFailure)) { '' } else { " $submitFailure" }
    $submitSuffix = if ($submitted) { '' } else { ' Submission never ran because the exact-version state stayed unreadable.' }
    $stateSuffix = if ([string]::IsNullOrWhiteSpace([string]$state.Reason)) { '' } else { " Last state: $($state.Kind) - $($state.Reason)" }
    # A bare exit code hides the WingetCreate error that actually blocked the release, so the
    # captured (redacted, bounded) output travels with the failure.
    $diagnosticsSuffix = if ($submitDiagnostics.Count -eq 0) { '' } else { "`nWingetCreate output (token-shaped values redacted):`n" + ($submitDiagnostics -join "`n") }
    throw "No exact automation-owned Winget pull request became visible after $MaximumVisibilityAttempts direct-list attempts.$failureSuffix$submitSuffix$stateSuffix$diagnosticsSuffix"
}

Export-ModuleMember -Function @(
    'ConvertTo-RSWingetCanonicalManifest',
    'Get-RSWingetManifestContentHash',
    'Get-RSWingetPublicationState',
    'Invoke-RSWingetPublication',
    'New-RSWingetPublicationIdentity',
    'Set-RSWingetPullRequestMetadata'
)
