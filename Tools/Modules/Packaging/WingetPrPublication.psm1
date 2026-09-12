# Private packaging module; import through its reviewed callers.
Set-StrictMode -Version Latest

function Get-RSSha256Hex {
    param(
        [Parameter(Mandatory = $true)]
        [byte[]]$Bytes
    )

    return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($Bytes)).ToLowerInvariant()
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
        $manifestHashes[$remotePath] = Get-RSSha256Hex -Bytes ([IO.File]::ReadAllBytes($localPath))
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
    $getFileSha256 = Get-RSRepositoryOperation -Repository $Repository -Name 'GetFileSha256'

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
        return [pscustomobject]@{ Kind = 'Pending'; Reason = "Pull request #$($pullRequest.number) files are not yet visible."; PullRequest = $pullRequest }
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
            $actualHash = [string](& $getFileSha256 `
                    ([string]$pullRequest.head.repo.full_name) `
                    $path `
                    ([string]$pullRequest.head.sha))
        }
        catch [System.Exception] {
            return [pscustomobject]@{ Kind = 'Pending'; Reason = "Pull request #$($pullRequest.number) manifest content is not yet visible."; PullRequest = $pullRequest }
        }
        if ($actualHash -cne [string]$Identity.ManifestHashes[$path]) {
            return [pscustomobject]@{ Kind = 'Conflict'; Reason = "Pull request #$($pullRequest.number) manifest hash does not match '$path'."; PullRequest = $pullRequest }
        }
    }

    return [pscustomobject]@{ Kind = 'Owned'; Reason = "Pull request #$($pullRequest.number) is an exact automation-owned publication."; PullRequest = $pullRequest }
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

        [scriptblock]$Wait = { param([int]$Attempt) Start-Sleep -Seconds ([Math]::Min($Attempt, 5)) },

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
    if ($state.Kind -ceq 'None') {
        $submitted = $true
        try {
            $submitResult = & $Submit $Identity.InitialTitle $Identity $Repository
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
    throw "No exact automation-owned Winget pull request became visible after $MaximumVisibilityAttempts direct-list attempts.$failureSuffix$submitSuffix$stateSuffix"
}

Export-ModuleMember -Function @(
    'Get-RSWingetPublicationState',
    'Invoke-RSWingetPublication',
    'New-RSWingetPublicationIdentity',
    'Set-RSWingetPullRequestMetadata'
)
