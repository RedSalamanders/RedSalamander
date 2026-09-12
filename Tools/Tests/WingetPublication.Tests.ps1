Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$testRunPlanModule = Join-Path $repoRoot 'Tools\Modules\Testing\TestRunPlan.psm1'
$publicationModule = Join-Path $repoRoot 'Tools\Modules\Packaging\WingetPrPublication.psm1'
Import-Module $testRunPlanModule -Force -ErrorAction Stop
Import-Module $publicationModule -Force -ErrorAction Stop

function New-RSWingetPublicationTestManifestRoot {
    param(
        [string]$Version = '7.0.42'
    )

    $root = New-RSTestSandboxScratchDirectory `
        -RepoRoot $repoRoot `
        -Harness 'tools-pester' `
        -Case "winget-publication-$([Guid]::NewGuid().ToString('N'))"
    $packageIdentifier = 'RedSalamanders.RedSalamander'
    $contents = [ordered]@{
        "$packageIdentifier.installer.yaml" = @"
PackageIdentifier: $packageIdentifier
PackageVersion: $Version
InstallerType: zip
"@
        "$packageIdentifier.locale.en-US.yaml" = @"
PackageIdentifier: $packageIdentifier
PackageVersion: $Version
PackageLocale: en-US
"@
        "$packageIdentifier.yaml" = @"
PackageIdentifier: $packageIdentifier
PackageVersion: $Version
DefaultLocale: en-US
"@
    }
    foreach ($entry in $contents.GetEnumerator()) {
        [IO.File]::WriteAllText(
            (Join-Path $root $entry.Key),
            $entry.Value.Replace("`r`n", "`n"),
            [Text.UTF8Encoding]::new($false))
    }

    return $root
}

function New-RSWingetPublicationTestRepository {
    $state = @{
        PullRequests = [Collections.ArrayList]::new()
        FilesByNumber = @{}
        HashesByRefAndPath = @{}
        HiddenListCalls = 0
        ListCalls = 0
        ListFailure = $null
        PatchCalls = 0
        SubmitCalls = 0
        WaitCalls = 0
        NextNumber = 100
    }

    # The comma keeps these arrays unenumerated, which is how Invoke-RestMethod returns a JSON array.
    $listPullRequests = {
        param([int]$Page, [int]$PageSize)
        $state.ListCalls++
        if ($state.ListFailure) {
            throw $state.ListFailure
        }
        if ($Page -ne 1) {
            return ,@()
        }
        if ($state.HiddenListCalls -gt 0) {
            $state.HiddenListCalls--
            return ,@()
        }
        return ,@($state.PullRequests)
    }.GetNewClosure()
    $listPullRequestFiles = {
        param([int]$Number, [int]$Page, [int]$PageSize)
        if ($Page -ne 1) {
            return ,@()
        }
        return ,@($state.FilesByNumber[[string]$Number])
    }.GetNewClosure()
    $getFileSha256 = {
        param([string]$RepositoryFullName, [string]$Path, [string]$Ref)
        $key = "$Ref|$Path"
        if (-not $state.HashesByRefAndPath.ContainsKey($key)) {
            throw "Remote file is not visible: $key"
        }
        return $state.HashesByRefAndPath[$key]
    }.GetNewClosure()
    $patchPullRequest = {
        param([int]$Number, [string]$Title, [string]$Body)
        $state.PatchCalls++
        $pullRequest = @($state.PullRequests | Where-Object { [int]$_.number -eq $Number })[0]
        $pullRequest.title = $Title
        $pullRequest.body = $Body
        return $pullRequest
    }.GetNewClosure()

    return [pscustomobject]@{
        State = $state
        Operations = @{
            _TestState = $state
            ListPullRequests = $listPullRequests
            ListPullRequestFiles = $listPullRequestFiles
            GetFileSha256 = $getFileSha256
            PatchPullRequest = $patchPullRequest
        }
    }
}

function Add-RSWingetPublicationTestPullRequest {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$State,

        [Parameter(Mandatory = $true)]
        [pscustomobject]$Identity,

        [string]$Author = 'rs-automation',
        [string]$HeadRepository = 'rs-automation/winget-pkgs',
        [string]$PullRequestState = 'open',
        [ValidateSet('Title', 'Body', 'None')]
        [string]$MarkerLocation = 'Title',
        [switch]$HashMismatch,
        [switch]$ExtraFile
    )

    $number = [int]$State.NextNumber
    $State.NextNumber++
    $guidSuffix = $number.ToString('000000000000')
    $headRef = "$($Identity.PackageIdentifier)-$($Identity.Version)-11111111-1111-1111-1111-$guidSuffix"
    $headSha = "test-head-$number"
    $title = if ($MarkerLocation -ceq 'Title') { $Identity.InitialTitle } else { $Identity.PullRequestTitle }
    $body = if ($MarkerLocation -ceq 'Body') { "$($Identity.BodyMarker)`n`nExisting body" } else { 'WingetCreate body' }
    $pullRequest = [pscustomobject]@{
        number = $number
        state = $PullRequestState
        title = $title
        body = $body
        html_url = "https://github.com/microsoft/winget-pkgs/pull/$number"
        user = [pscustomobject]@{ login = $Author }
        head = [pscustomobject]@{
            ref = $headRef
            sha = $headSha
            repo = [pscustomobject]@{ full_name = $HeadRepository }
        }
    }
    [void]$State.PullRequests.Add($pullRequest)

    $files = @()
    foreach ($entry in $Identity.ManifestHashes.GetEnumerator()) {
        $files += [pscustomobject]@{ filename = $entry.Key; status = 'added' }
        $hash = if ($HashMismatch -and $files.Count -eq 1) { '0' * 64 } else { $entry.Value }
        $State.HashesByRefAndPath["$headSha|$($entry.Key)"] = $hash
    }
    if ($ExtraFile) {
        $files += [pscustomobject]@{ filename = 'README.md'; status = 'modified' }
    }
    $State.FilesByNumber[[string]$number] = $files
    return $pullRequest
}

Describe 'Resumable Winget pull-request publication' {
    BeforeEach {
        $script:publicationTestRoots = @()
        $manifestRoot = New-RSWingetPublicationTestManifestRoot
        $script:publicationTestRoots += $manifestRoot
        $identity = New-RSWingetPublicationIdentity `
            -PackageIdentifier 'RedSalamanders.RedSalamander' `
            -Version '7.0.42' `
            -ManifestRoot $manifestRoot `
            -PullRequestTitle 'Update: RedSalamanders.RedSalamander to 7.0.42'
        $testRepository = New-RSWingetPublicationTestRepository
        $expectedAuthor = 'rs-automation'
        $expectedHeadRepository = 'rs-automation/winget-pkgs'
        $body = 'Automated validation evidence.'
        $noWait = {
            param([int]$Attempt, [hashtable]$Repository)
            $Repository['_TestState']['WaitCalls']++
        }
    }

    AfterEach {
        foreach ($root in $script:publicationTestRoots) {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'creates only when the exact package-version state is empty' {
        $submit = {
            param([string]$InitialTitle, [pscustomobject]$PublicationIdentity, [hashtable]$Repository)
            $state = $Repository['_TestState']
            $state['SubmitCalls']++
            Add-RSWingetPublicationTestPullRequest `
                -State $state `
                -Identity $PublicationIdentity | Out-Null
            return [pscustomobject]@{ ExitCode = 0; Output = @('created') }
        }

        $result = Invoke-RSWingetPublication `
            -Identity $identity `
            -ExpectedAuthor $expectedAuthor `
            -ExpectedHeadRepository $expectedHeadRepository `
            -PullRequestBody $body `
            -Repository $testRepository.Operations `
            -Submit $submit `
            -Wait $noWait

        $result.Submitted | Should Be $true
        $result.Url | Should Be 'https://github.com/microsoft/winget-pkgs/pull/100'
        $testRepository.State.SubmitCalls | Should Be 1
        $testRepository.State.PatchCalls | Should Be 1
    }

    It 'resumes an exact automation-owned pull request without submitting again' {
        Add-RSWingetPublicationTestPullRequest `
            -State $testRepository.State `
            -Identity $identity | Out-Null
        $submit = { throw 'submit must not run' }

        $result = Invoke-RSWingetPublication `
            -Identity $identity `
            -ExpectedAuthor $expectedAuthor `
            -ExpectedHeadRepository $expectedHeadRepository `
            -PullRequestBody $body `
            -Repository $testRepository.Operations `
            -Submit $submit `
            -Wait $noWait

        $result.Submitted | Should Be $false
        $testRepository.State.PatchCalls | Should Be 1
        $testRepository.State.PullRequests[0].body | Should Match ([regex]::Escape($identity.BodyMarker))
    }

    It 'recovers when submission creates the PR and then throws' {
        $submit = {
            param([string]$InitialTitle, [pscustomobject]$PublicationIdentity, [hashtable]$Repository)
            $state = $Repository['_TestState']
            $state['SubmitCalls']++
            Add-RSWingetPublicationTestPullRequest `
                -State $state `
                -Identity $PublicationIdentity | Out-Null
            throw 'simulated post-create crash'
        }

        $result = Invoke-RSWingetPublication `
            -Identity $identity `
            -ExpectedAuthor $expectedAuthor `
            -ExpectedHeadRepository $expectedHeadRepository `
            -PullRequestBody $body `
            -Repository $testRepository.Operations `
            -Submit $submit `
            -Wait $noWait

        $result.Url | Should Be 'https://github.com/microsoft/winget-pkgs/pull/100'
        $testRepository.State.PullRequests.Count | Should Be 1
    }

    It 'recovers from malformed WingetCreate output through direct listing' {
        $submit = {
            param([string]$InitialTitle, [pscustomobject]$PublicationIdentity, [hashtable]$Repository)
            $state = $Repository['_TestState']
            $state['SubmitCalls']++
            Add-RSWingetPublicationTestPullRequest `
                -State $state `
                -Identity $PublicationIdentity | Out-Null
            return [pscustomobject]@{ ExitCode = 0; Output = @('no pull request URL here') }
        }

        $result = Invoke-RSWingetPublication `
            -Identity $identity `
            -ExpectedAuthor $expectedAuthor `
            -ExpectedHeadRepository $expectedHeadRepository `
            -PullRequestBody $body `
            -Repository $testRepository.Operations `
            -Submit $submit `
            -Wait $noWait

        $result.PullRequestNumber | Should Be 100
        $testRepository.State.PatchCalls | Should Be 1
    }

    It 'stops without mutation on an external exact-version pull request' {
        Add-RSWingetPublicationTestPullRequest `
            -State $testRepository.State `
            -Identity $identity `
            -Author 'external-user' `
            -HeadRepository 'external-user/winget-pkgs' `
            -MarkerLocation None | Out-Null
        $submit = { throw 'submit must not run' }

        { Invoke-RSWingetPublication `
                -Identity $identity `
                -ExpectedAuthor $expectedAuthor `
                -ExpectedHeadRepository $expectedHeadRepository `
                -PullRequestBody $body `
                -Repository $testRepository.Operations `
                -Submit $submit `
                -Wait $noWait } |
            Should Throw 'is not owned by this publication identity'
        $testRepository.State.PatchCalls | Should Be 0
    }

    It 'stops without mutation on duplicate exact-version pull requests' {
        Add-RSWingetPublicationTestPullRequest -State $testRepository.State -Identity $identity | Out-Null
        Add-RSWingetPublicationTestPullRequest -State $testRepository.State -Identity $identity | Out-Null
        $submit = { throw 'submit must not run' }

        { Invoke-RSWingetPublication `
                -Identity $identity `
                -ExpectedAuthor $expectedAuthor `
                -ExpectedHeadRepository $expectedHeadRepository `
                -PullRequestBody $body `
                -Repository $testRepository.Operations `
                -Submit $submit `
                -Wait $noWait } |
            Should Throw 'Multiple pull requests target RedSalamanders.RedSalamander 7.0.42.'
        $testRepository.State.PatchCalls | Should Be 0
    }

    It 'stops without mutation on a closed exact-version pull request' {
        Add-RSWingetPublicationTestPullRequest `
            -State $testRepository.State `
            -Identity $identity `
            -PullRequestState closed | Out-Null
        $submit = { throw 'submit must not run' }

        { Invoke-RSWingetPublication `
                -Identity $identity `
                -ExpectedAuthor $expectedAuthor `
                -ExpectedHeadRepository $expectedHeadRepository `
                -PullRequestBody $body `
                -Repository $testRepository.Operations `
                -Submit $submit `
                -Wait $noWait } |
            Should Throw 'Pull request #100 is not open.'
        $testRepository.State.PatchCalls | Should Be 0
    }

    It 'stops when exact manifest content or file scope differs' {
        Add-RSWingetPublicationTestPullRequest `
            -State $testRepository.State `
            -Identity $identity `
            -HashMismatch | Out-Null
        $submit = { throw 'submit must not run' }

        { Invoke-RSWingetPublication `
                -Identity $identity `
                -ExpectedAuthor $expectedAuthor `
                -ExpectedHeadRepository $expectedHeadRepository `
                -PullRequestBody $body `
                -Repository $testRepository.Operations `
                -Submit $submit `
                -Wait $noWait } |
            Should Throw 'manifest hash does not match'

        $testRepository = New-RSWingetPublicationTestRepository
        Add-RSWingetPublicationTestPullRequest `
            -State $testRepository.State `
            -Identity $identity `
            -ExtraFile | Out-Null
        { Invoke-RSWingetPublication `
                -Identity $identity `
                -ExpectedAuthor $expectedAuthor `
                -ExpectedHeadRepository $expectedHeadRepository `
                -PullRequestBody $body `
                -Repository $testRepository.Operations `
                -Submit $submit `
                -Wait $noWait } |
            Should Throw 'does not contain the exact manifest file set'
    }

    It 'uses bounded direct-list retries for eventual visibility' {
        $testRepository.State.HiddenListCalls = 2
        $submit = {
            param([string]$InitialTitle, [pscustomobject]$PublicationIdentity, [hashtable]$Repository)
            $state = $Repository['_TestState']
            $state['SubmitCalls']++
            Add-RSWingetPublicationTestPullRequest `
                -State $state `
                -Identity $PublicationIdentity | Out-Null
            return [pscustomobject]@{ ExitCode = 0; Output = @() }
        }

        $result = Invoke-RSWingetPublication `
            -Identity $identity `
            -ExpectedAuthor $expectedAuthor `
            -ExpectedHeadRepository $expectedHeadRepository `
            -PullRequestBody $body `
            -Repository $testRepository.Operations `
            -Submit $submit `
            -Wait $noWait `
            -MaximumVisibilityAttempts 3

        $result.PullRequestNumber | Should Be 100
        $testRepository.State.ListCalls | Should Be 3
        $testRepository.State.WaitCalls | Should Be 1

        $emptyRepository = New-RSWingetPublicationTestRepository
        $emptyWait = {
            param([int]$Attempt, [hashtable]$Repository)
            $Repository['_TestState']['WaitCalls']++
        }
        $failedSubmit = {
            param([string]$InitialTitle)
            return [pscustomobject]@{ ExitCode = 9; Output = @() }
        }
        { Invoke-RSWingetPublication `
                -Identity $identity `
                -ExpectedAuthor $expectedAuthor `
                -ExpectedHeadRepository $expectedHeadRepository `
                -PullRequestBody $body `
                -Repository $emptyRepository.Operations `
                -Submit $failedSubmit `
                -Wait $emptyWait `
                -MaximumVisibilityAttempts 3 } |
            Should Throw 'after 3 direct-list attempts'
        $emptyRepository.State.ListCalls | Should Be 4
        $emptyRepository.State.WaitCalls | Should Be 2
    }

    It 'names the blocking reason when the upstream list never becomes readable' {
        $blockedRepository = New-RSWingetPublicationTestRepository
        $blockedRepository.State.ListFailure = 'Response status code does not indicate success: 403 (Forbidden).'
        $blockedWait = {
            param([int]$Attempt, [hashtable]$Repository)
            $Repository['_TestState']['WaitCalls']++
        }
        $submit = {
            param([string]$InitialTitle, [pscustomobject]$PublicationIdentity, [hashtable]$Repository)
            $Repository['_TestState']['SubmitCalls']++
            return [pscustomobject]@{ ExitCode = 0; Output = @() }
        }

        $failure = $null
        try {
            Invoke-RSWingetPublication `
                -Identity $identity `
                -ExpectedAuthor $expectedAuthor `
                -ExpectedHeadRepository $expectedHeadRepository `
                -PullRequestBody $body `
                -Repository $blockedRepository.Operations `
                -Submit $submit `
                -Wait $blockedWait `
                -MaximumVisibilityAttempts 2
        }
        catch {
            $failure = [string]$_.Exception.Message
        }

        $failure | Should Match 'after 2 direct-list attempts'
        $failure | Should Match 'Submission never ran'
        $failure | Should Match 'Last state: Pending'
        $failure | Should Match '403 \(Forbidden\)'
        $blockedRepository.State.SubmitCalls | Should Be 0
    }

    It 'repairs idempotently and serialized attempts converge on one PR' {
        $submit = {
            param([string]$InitialTitle, [pscustomobject]$PublicationIdentity, [hashtable]$Repository)
            $state = $Repository['_TestState']
            $state['SubmitCalls']++
            Add-RSWingetPublicationTestPullRequest `
                -State $state `
                -Identity $PublicationIdentity | Out-Null
            return [pscustomobject]@{ ExitCode = 0; Output = @() }
        }

        $first = Invoke-RSWingetPublication `
            -Identity $identity `
            -ExpectedAuthor $expectedAuthor `
            -ExpectedHeadRepository $expectedHeadRepository `
            -PullRequestBody $body `
            -Repository $testRepository.Operations `
            -Submit $submit `
            -Wait $noWait
        $second = Invoke-RSWingetPublication `
            -Identity $identity `
            -ExpectedAuthor $expectedAuthor `
            -ExpectedHeadRepository $expectedHeadRepository `
            -PullRequestBody $body `
            -Repository $testRepository.Operations `
            -Submit $submit `
            -Wait $noWait

        $first.Url | Should Be $second.Url
        $testRepository.State.PullRequests.Count | Should Be 1
        $testRepository.State.SubmitCalls | Should Be 1
        $testRepository.State.PatchCalls | Should Be 2
    }
}
