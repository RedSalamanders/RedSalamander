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
        FileListCalls = 0
        ListFailure = $null
        FileListFailure = $null
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
        $state.FileListCalls++
        if ($state.FileListFailure) {
            throw $state.FileListFailure
        }
        if ($Page -ne 1) {
            return ,@()
        }
        return ,@($state.FilesByNumber[[string]$Number])
    }.GetNewClosure()
    $getManifestContentHash = {
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
            GetManifestContentHash = $getManifestContentHash
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

function Add-RSWingetPublicationTestUnrelatedPullRequest {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$State,

        [string]$PackageIdentifier = 'Contoso.Widget',
        [string]$Version = '1.2.3',
        [string]$Author = 'contoso-bot',
        [string]$Title = 'Update: Contoso.Widget to 1.2.3',
        [string]$Body = 'Automated Contoso.Widget update.'
    )

    $number = [int]$State.NextNumber
    $State.NextNumber++
    $guidSuffix = $number.ToString('000000000000')
    $headSha = "test-head-$number"
    $segments = @($PackageIdentifier -split '\.')
    $manifestPrefix = "manifests/$($segments[0].Substring(0, 1).ToLowerInvariant())/$($segments -join '/')/$Version"
    $pullRequest = [pscustomobject]@{
        number = $number
        state = 'open'
        title = $Title
        body = $Body
        html_url = "https://github.com/microsoft/winget-pkgs/pull/$number"
        user = [pscustomobject]@{ login = $Author }
        head = [pscustomobject]@{
            ref = "$PackageIdentifier-$Version-22222222-2222-2222-2222-$guidSuffix"
            sha = $headSha
            repo = [pscustomobject]@{ full_name = "$Author/winget-pkgs" }
        }
    }
    [void]$State.PullRequests.Add($pullRequest)
    $State.FilesByNumber[[string]$number] = @(
        [pscustomobject]@{ filename = "$manifestPrefix/$PackageIdentifier.installer.yaml"; status = 'added' },
        [pscustomobject]@{ filename = "$manifestPrefix/$PackageIdentifier.locale.en-US.yaml"; status = 'added' },
        [pscustomobject]@{ filename = "$manifestPrefix/$PackageIdentifier.yaml"; status = 'added' }
    )
    return $pullRequest
}

Describe 'Canonical Winget manifest content' {
    BeforeAll {
        # Shape of Installer/winget/templates output: two-space list indentation, ReleaseDate before
        # Installers, LF endings, and a schema comment.
        $generated = @(
            '# yaml-language-server: $schema=https://aka.ms/winget-manifest.installer.1.12.0.schema.json',
            '',
            'PackageIdentifier: RedSalamanders.RedSalamander',
            'PackageVersion: 7.0.59',
            'Platform:',
            '  - Windows.Desktop',
            'NestedInstallerFiles:',
            '  - RelativeFilePath: RedLauncher.exe',
            '    PortableCommandAlias: RedSalamander',
            '  - RelativeFilePath: red.exe',
            '    PortableCommandAlias: red',
            'Description: |-',
            '  Features include:',
            '  - Dual-pane file management',
            'ReleaseDate: 2026-09-16',
            'Installers:',
            '  - Architecture: x64',
            '    InstallerUrl: https://github.com/RedSalamanders/RedSalamander/releases/download/v7.0.59/RedSalamander-7.0.59-x64-Portable.zip',
            '    InstallerSha256: 7F29F0E7D8842C77970F3947A0851132BB493361D20B9E04E628AB42D4597EE2',
            'ManifestType: installer',
            'ManifestVersion: 1.12.0'
        ) -join "`n"
        # Shape wingetcreate 1.12.8.0 committed for winget-pkgs#436389: a creation comment, CRLF,
        # zero-indent list markers, and ReleaseDate moved to the end.
        $committed = @(
            '# Created using wingetcreate 1.12.8.0',
            '# yaml-language-server: $schema=https://aka.ms/winget-manifest.installer.1.12.0.schema.json',
            '',
            'PackageIdentifier: RedSalamanders.RedSalamander',
            'PackageVersion: 7.0.59',
            'Platform:',
            '- Windows.Desktop',
            'NestedInstallerFiles:',
            '- RelativeFilePath: RedLauncher.exe',
            '  PortableCommandAlias: RedSalamander',
            '- RelativeFilePath: red.exe',
            '  PortableCommandAlias: red',
            'Description: |-',
            '  Features include:',
            '  - Dual-pane file management',
            'Installers:',
            '- Architecture: x64',
            '  InstallerUrl: https://github.com/RedSalamanders/RedSalamander/releases/download/v7.0.59/RedSalamander-7.0.59-x64-Portable.zip',
            '  InstallerSha256: 7F29F0E7D8842C77970F3947A0851132BB493361D20B9E04E628AB42D4597EE2',
            'ManifestType: installer',
            'ManifestVersion: 1.12.0',
            'ReleaseDate: 2026-09-16',
            ''
        ) -join "`r`n"
        $utf8 = [Text.UTF8Encoding]::new($false)
    }

    It 'hashes the generated manifest and its wingetcreate re-serialization identically' {
        $generatedHash = Get-RSWingetManifestContentHash -Bytes $utf8.GetBytes($generated)
        $committedHash = Get-RSWingetManifestContentHash -Bytes $utf8.GetBytes($committed)
        $committedHash | Should Be $generatedHash
        $generatedHash | Should Match '^[0-9a-f]{64}$'

        $withBom = [byte[]](@(0xEF, 0xBB, 0xBF) + $utf8.GetBytes($generated))
        (Get-RSWingetManifestContentHash -Bytes $withBom) | Should Be $generatedHash
    }

    It 'keeps every scalar line so a changed value, date, alias, or missing line still differs' {
        $baseline = Get-RSWingetManifestContentHash -Bytes $utf8.GetBytes($generated)
        $variants = @(
            $generated.Replace('7F29F0E7', '7F29F0E8'),
            $generated.Replace('ReleaseDate: 2026-09-16', 'ReleaseDate: 2026-09-17'),
            $generated.Replace('PortableCommandAlias: red', 'PortableCommandAlias: rs'),
            $generated.Replace("`n  - Dual-pane file management", ''),
            $generated.Replace('v7.0.59/', 'v7.0.60/')
        )
        foreach ($variant in $variants) {
            (Get-RSWingetManifestContentHash -Bytes $utf8.GetBytes($variant)) | Should Not Be $baseline
        }
    }

    It 'canonicalizes to sorted scalar lines without comments, markers, or blank lines' {
        $canonical = ConvertTo-RSWingetCanonicalManifest -Content $committed
        $lines = @($canonical -split "`n")
        $lines | Should Not Match '^#'
        $lines | Should Not Match '^\s'
        $lines | Should Not Match '^- '
        $lines | Should Not Match '^$'
        $canonical | Should Not Match "`r"
        ($lines -join "`n") | Should Be (@($lines | Sort-Object -CaseSensitive) -join "`n")
        $lines -contains 'RelativeFilePath: red.exe' | Should Be $true
        $lines -contains 'Dual-pane file management' | Should Be $true
        (ConvertTo-RSWingetCanonicalManifest -Content '') | Should Be ''
    }

    It 'binds the publication identity to canonical content rather than bytes' {
        $lfRoot = New-RSWingetPublicationTestManifestRoot
        $crlfRoot = New-RSWingetPublicationTestManifestRoot
        try {
            foreach ($file in Get-ChildItem -LiteralPath $crlfRoot -File) {
                $text = [IO.File]::ReadAllText($file.FullName)
                [IO.File]::WriteAllText($file.FullName, "# Created using wingetcreate 1.12.8.0`r`n" + $text.Replace("`n", "`r`n"), $utf8)
            }
            $lfIdentity = New-RSWingetPublicationIdentity `
                -PackageIdentifier 'RedSalamanders.RedSalamander' `
                -Version '7.0.42' `
                -ManifestRoot $lfRoot `
                -PullRequestTitle 'Update: RedSalamanders.RedSalamander to 7.0.42'
            $crlfIdentity = New-RSWingetPublicationIdentity `
                -PackageIdentifier 'RedSalamanders.RedSalamander' `
                -Version '7.0.42' `
                -ManifestRoot $crlfRoot `
                -PullRequestTitle 'Update: RedSalamanders.RedSalamander to 7.0.42'

            $crlfIdentity.PublicationId | Should Be $lfIdentity.PublicationId
            $crlfIdentity.InitialTitle | Should Be $lfIdentity.InitialTitle
        }
        finally {
            Remove-Item -LiteralPath $lfRoot, $crlfRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
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
            Should Throw 'canonical manifest content does not match'

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

    It 'reads changed files only for candidate pull requests' {
        Add-RSWingetPublicationTestUnrelatedPullRequest -State $testRepository.State | Out-Null
        Add-RSWingetPublicationTestUnrelatedPullRequest `
            -State $testRepository.State `
            -PackageIdentifier 'Fabrikam.Tool' `
            -Version '4.5.6' `
            -Author 'fabrikam-release' `
            -Title 'New package: Fabrikam.Tool version 4.5.6' `
            -Body 'Fresh package.' | Out-Null
        Add-RSWingetPublicationTestUnrelatedPullRequest `
            -State $testRepository.State `
            -PackageIdentifier 'RedSalamanders.OtherApp' `
            -Version '7.0.42' `
            -Author 'someone-else' `
            -Title 'Update: RedSalamanders.OtherApp to 7.0.42' `
            -Body 'Same publisher, different package.' | Out-Null
        Add-RSWingetPublicationTestPullRequest -State $testRepository.State -Identity $identity | Out-Null
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
        $result.PullRequestNumber | Should Be 103
        $testRepository.State.FileListCalls | Should Be 1
        $testRepository.State.PatchCalls | Should Be 1
    }

    It 'creates past unrelated pull requests without reading their files' {
        Add-RSWingetPublicationTestUnrelatedPullRequest -State $testRepository.State | Out-Null
        Add-RSWingetPublicationTestUnrelatedPullRequest `
            -State $testRepository.State `
            -PackageIdentifier 'Fabrikam.Tool' `
            -Version '4.5.6' `
            -Author 'fabrikam-release' `
            -Title 'New package: Fabrikam.Tool version 4.5.6' `
            -Body 'Fresh package.' | Out-Null
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
        $result.PullRequestNumber | Should Be 102
        $testRepository.State.SubmitCalls | Should Be 1
        $testRepository.State.ListCalls | Should Be 2
        $testRepository.State.FileListCalls | Should Be 1
    }

    It 'still stops on an external pull request that names the package only in its head branch' {
        $external = Add-RSWingetPublicationTestPullRequest `
            -State $testRepository.State `
            -Identity $identity `
            -Author 'external-user' `
            -HeadRepository 'external-user/winget-pkgs' `
            -MarkerLocation None
        $external.title = 'Automatic package update'
        $external.body = 'Nothing named here.'
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
        $testRepository.State.FileListCalls | Should Be 1
        $testRepository.State.PatchCalls | Should Be 0
    }

    It 'still inspects the token owner''s pull requests when nothing else names the package' {
        $renamed = Add-RSWingetPublicationTestPullRequest `
            -State $testRepository.State `
            -Identity $identity `
            -MarkerLocation None
        $renamed.title = 'Renamed by a maintainer'
        $renamed.body = 'Marker removed.'
        $renamed.head.ref = 'maintainer-rename'
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
        $testRepository.State.FileListCalls | Should Be 1
        $testRepository.State.PatchCalls | Should Be 0
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

    It 'surfaces redacted WingetCreate output when submission fails and nothing becomes visible' {
        $emptyRepository = New-RSWingetPublicationTestRepository
        $emptyWait = {
            param([int]$Attempt, [hashtable]$Repository)
            $Repository['_TestState']['WaitCalls']++
        }
        $classicToken = 'ghp_' + ('A1b2C3d4E5' * 4)
        $fineGrainedToken = 'github_pat_' + ('Z9y8X7w6V5' * 3) + '_' + ('q1W2e3R4t5' * 6)
        $failedSubmit = {
            param([string]$InitialTitle)
            return [pscustomobject]@{
                ExitCode = 1
                Output = @(
                    'Telemetry Settings',
                    '------------------',
                    '',
                    "   Checking access with Authorization: Bearer $classicToken   ",
                    "token=$fineGrainedToken",
                    'password: 0123456789abcdefghij token expired',
                    'Error: Unable to sync your fork with the upstream repository. Please update your fork and try again.'
                )
            }
        }.GetNewClosure()

        $failure = $null
        try {
            Invoke-RSWingetPublication `
                -Identity $identity `
                -ExpectedAuthor $expectedAuthor `
                -ExpectedHeadRepository $expectedHeadRepository `
                -PullRequestBody $body `
                -Repository $emptyRepository.Operations `
                -Submit $failedSubmit `
                -Wait $emptyWait `
                -MaximumVisibilityAttempts 2
        }
        catch {
            $failure = [string]$_.Exception.Message
        }

        $failure | Should Match 'wingetcreate submit failed with exit code 1\.'
        $failure | Should Match 'Last state: None'
        $failure | Should Match 'WingetCreate output \(token-shaped values redacted\):'
        $failure | Should Match '(?m)^Error: Unable to sync your fork with the upstream repository\.'
        $failure | Should Match '(?m)^Checking access with Authorization: Bearer \*\*\*$'
        $failure | Should Match '(?m)^token=\*\*\*$'
        $failure | Should Match '(?m)^password: \*\*\* token expired$'
        $failure | Should Not Match '0123456789abcdefghij'
        $failure | Should Not Match 'ghp_'
        $failure | Should Not Match 'github_pat_'
        $failure | Should Not Match "(?m)^\s*$"
    }

    It 'bounds WingetCreate output and keeps its last lines when submission succeeds but nothing becomes visible' {
        $emptyRepository = New-RSWingetPublicationTestRepository
        $emptyWait = {
            param([int]$Attempt, [hashtable]$Repository)
            $Repository['_TestState']['WaitCalls']++
        }
        $noisySubmit = {
            param([string]$InitialTitle)
            $lines = @(1..60 | ForEach-Object { "progress line $_" })
            $lines += 'x' * 700
            return [pscustomobject]@{ ExitCode = 0; Output = $lines }
        }

        $failure = $null
        try {
            Invoke-RSWingetPublication `
                -Identity $identity `
                -ExpectedAuthor $expectedAuthor `
                -ExpectedHeadRepository $expectedHeadRepository `
                -PullRequestBody $body `
                -Repository $emptyRepository.Operations `
                -Submit $noisySubmit `
                -Wait $emptyWait `
                -MaximumVisibilityAttempts 2
        }
        catch {
            $failure = [string]$_.Exception.Message
        }

        $failure | Should Match 'after 2 direct-list attempts\. Last state: None'
        $failure | Should Not Match 'wingetcreate submit failed'
        $failure | Should Match '\(21 earlier line\(s\) omitted\)'
        $failure | Should Not Match '(?m)^progress line 21$'
        $failure | Should Match '(?m)^progress line 22$'
        $failure | Should Match '(?m)^progress line 60$'
        $failure | Should Match "(?m)^x{500} \.\.\.$"
        $failure | Should Not Match 'x{501}'
    }

    It 'omits the WingetCreate output section when submission produced none' {
        $emptyRepository = New-RSWingetPublicationTestRepository
        $emptyWait = {
            param([int]$Attempt, [hashtable]$Repository)
            $Repository['_TestState']['WaitCalls']++
        }
        $silentSubmit = {
            param([string]$InitialTitle)
            return [pscustomobject]@{ ExitCode = 1 }
        }

        $failure = $null
        try {
            Invoke-RSWingetPublication `
                -Identity $identity `
                -ExpectedAuthor $expectedAuthor `
                -ExpectedHeadRepository $expectedHeadRepository `
                -PullRequestBody $body `
                -Repository $emptyRepository.Operations `
                -Submit $silentSubmit `
                -Wait $emptyWait `
                -MaximumVisibilityAttempts 2
        }
        catch {
            $failure = [string]$_.Exception.Message
        }

        $failure | Should Match 'wingetcreate submit failed with exit code 1\.'
        $failure | Should Not Match 'WingetCreate output'
    }

    It 'names the path, commit, and upstream error when manifest content never becomes visible' {
        $created = Add-RSWingetPublicationTestPullRequest -State $testRepository.State -Identity $identity
        $missingPath = @($identity.ManifestHashes.Keys)[0]
        $testRepository.State.HashesByRefAndPath.Remove("$($created.head.sha)|$missingPath")
        $submit = { throw 'submit must not run' }

        $failure = $null
        try {
            Invoke-RSWingetPublication `
                -Identity $identity `
                -ExpectedAuthor $expectedAuthor `
                -ExpectedHeadRepository $expectedHeadRepository `
                -PullRequestBody $body `
                -Repository $testRepository.Operations `
                -Submit $submit `
                -Wait $noWait `
                -MaximumVisibilityAttempts 2
        }
        catch {
            $failure = [string]$_.Exception.Message
        }

        $failure | Should Match 'after 2 direct-list attempts'
        $failure | Should Match "Last state: Pending - Pull request #100 manifest content is not yet visible: '$([regex]::Escape($missingPath))' at test-head-100: Remote file is not visible: test-head-100\|"
        $testRepository.State.PatchCalls | Should Be 0
    }

    It 'names the upstream error when an owned pull request''s file list never becomes visible' {
        Add-RSWingetPublicationTestPullRequest -State $testRepository.State -Identity $identity | Out-Null
        $testRepository.State.FileListFailure = 'Response status code does not indicate success: 403 (rate limited).'
        $submit = { throw 'submit must not run' }

        $failure = $null
        try {
            Invoke-RSWingetPublication `
                -Identity $identity `
                -ExpectedAuthor $expectedAuthor `
                -ExpectedHeadRepository $expectedHeadRepository `
                -PullRequestBody $body `
                -Repository $testRepository.Operations `
                -Submit $submit `
                -Wait $noWait `
                -MaximumVisibilityAttempts 2
        }
        catch {
            $failure = [string]$_.Exception.Message
        }

        $failure | Should Match 'Last state: Pending - Pull request #100 files are not yet visible: Response status code does not indicate success: 403 \(rate limited\)\.'
        $failure | Should Match 'Submission never ran'
        $testRepository.State.PatchCalls | Should Be 0
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
