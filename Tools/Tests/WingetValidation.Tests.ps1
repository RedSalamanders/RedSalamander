Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$testRunPlanModule = Join-Path $repoRoot 'Tools\Modules\Testing\TestRunPlan.psm1'
$helperScript = Join-Path $repoRoot 'Installer\winget\WingetValidation.ps1'
Import-Module $testRunPlanModule -Force -ErrorAction Stop

Describe 'Winget validation helper' {
    BeforeAll {
        . $helperScript
    }

    function New-RSTemporaryWingetManifestRoot {
        return (New-RSTestSandboxScratchDirectory `
                -RepoRoot $repoRoot `
                -Harness 'tools-pester' `
                -Case "winget-validation-$([Guid]::NewGuid().ToString('N'))")
    }

    function New-RSFakeWingetCommand {
        param(
            [Parameter(Mandatory = $true)]
            [string]$Root,

            [Parameter(Mandatory = $true)]
            [string]$VersionText,

            [Parameter(Mandatory = $true)]
            [string[]]$ValidationOutput,

            [Parameter(Mandatory = $true)]
            [int]$ValidationExitCode
        )

        $scriptPath = Join-Path $Root 'winget.cmd'
        $outputCommands = @($ValidationOutput | ForEach-Object { "echo $_" }) -join "`r`n"
        Set-Content -Path $scriptPath -Encoding ASCII -Value @"
@echo off
if "%1"=="--version" (
  echo $VersionText
  exit /b 0
)
if "%1"=="validate" (
  $outputCommands
  exit /b $ValidationExitCode
)
exit /b 99
"@
        return $scriptPath
    }

    It 'allows only the known winget 1.11 schema-header warning for 1.12 manifests' {
        $output = @(
            'Manifest validation succeeded with warnings.'
            'Manifest Warning: The schema header URL does not match the expected pattern. Value: # yaml-language-server: $schema=https://aka.ms/winget-manifest.installer.1.12.0.schema.json Line: 1, Column: 25 File: RedSalamanders.RedSalamander.installer.yaml'
            'Manifest Warning: The schema header URL does not match the expected pattern. Value: # yaml-language-server: $schema=https://aka.ms/winget-manifest.defaultLocale.1.12.0.schema.json Line: 1, Column: 25 File: RedSalamanders.RedSalamander.locale.en-US.yaml'
            'Manifest Warning: The schema header URL does not match the expected pattern. Value: # yaml-language-server: $schema=https://aka.ms/winget-manifest.version.1.12.0.schema.json Line: 1, Column: 25 File: RedSalamanders.RedSalamander.yaml'
        )

        Test-RSLegacyWingetSchemaHeaderWarning -WingetVersion 'v1.11.510' -OutputLines $output | Should Be $true
    }

    It 'rejects unrelated warnings' {
        $output = @(
            'Manifest validation succeeded with warnings.'
            'Manifest Warning: Field usage requires verified publishers. [Icons]'
        )

        Test-RSLegacyWingetSchemaHeaderWarning -WingetVersion 'v1.11.510' -OutputLines $output | Should Be $false
    }

    It 'does not suppress warnings from current winget validators' {
        $output = @(
            'Manifest validation succeeded with warnings.'
            'Manifest Warning: The schema header URL does not match the expected pattern. Value: # yaml-language-server: $schema=https://aka.ms/winget-manifest.installer.1.12.0.schema.json'
        )

        Test-RSLegacyWingetSchemaHeaderWarning -WingetVersion 'v1.28.240' -OutputLines $output | Should Be $false
    }

    It 'fails clearly when the winget executable is unavailable' {
        $tempRoot = New-RSTemporaryWingetManifestRoot
        try {
            $missingWinget = Join-Path $tempRoot 'missing-winget.exe'

            { Invoke-RSWingetManifestValidation -ManifestPath $tempRoot -WingetCommand $missingWinget } |
                Should Throw "winget executable not found: $missingWinget"
        } finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'propagates real validation failures with the winget exit code' {
        $tempRoot = New-RSTemporaryWingetManifestRoot
        try {
            $fakeWinget = New-RSFakeWingetCommand `
                -Root $tempRoot `
                -VersionText 'v1.12.0' `
                -ValidationOutput @('Manifest validation failed.', 'Manifest Error: PackageVersion is required.') `
                -ValidationExitCode 23

            { Invoke-RSWingetManifestValidation -ManifestPath $tempRoot -WingetCommand $fakeWinget } |
                Should Throw 'winget validate failed with exit code 23.'
        } finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'does not suppress unrelated validation failures from legacy winget validators' {
        $tempRoot = New-RSTemporaryWingetManifestRoot
        try {
            $fakeWinget = New-RSFakeWingetCommand `
                -Root $tempRoot `
                -VersionText 'v1.11.510' `
                -ValidationOutput @('Manifest validation failed.', 'Manifest Error: InstallerSha256 is invalid.') `
                -ValidationExitCode 7

            { Invoke-RSWingetManifestValidation -ManifestPath $tempRoot -WingetCommand $fakeWinget } |
                Should Throw 'winget validate failed with exit code 7.'
        } finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

Describe 'Winget release workflow' {
    BeforeAll {
        $workflowPath = Join-Path $repoRoot '.github\workflows\winget-release.yml'
        $workflow = Get-Content -Path $workflowPath -Raw
        $publicationModulePath = Join-Path $repoRoot 'Tools\Modules\Packaging\WingetPrPublication.psm1'
        $publicationModule = Get-Content -LiteralPath $publicationModulePath -Raw
    }

    It 'enables local manifest installs before testing the generated manifest' {
        $enableIndex = $workflow.IndexOf('winget settings --enable LocalManifestFiles')
        $installMatch = [regex]::Match($workflow, '(?ms)winget install\s+`\r?\n\s+--manifest \$env:WINGET_MANIFEST_ROOT')

        if ($enableIndex -lt 0) {
            throw 'winget-release.yml must enable LocalManifestFiles before winget install --manifest.'
        }

        if (-not $installMatch.Success) {
            throw 'winget-release.yml must install the generated winget-manifest.'
        }

        ($enableIndex -lt $installMatch.Index) | Should Be $true
    }

    It 'pins winget-create by its banner version and cleans up only an installed package' {
        $workflow | Should Match 'WingetCreateCLI\\s\+\(\?<version>'
        $workflow | Should Match '\\\+\[0-9a-fA-F\]\{7,40\}'
        $workflow | Should Match ([regex]::Escape("`$versionMatch.Groups['version'].Value"))
        $workflow | Should Not Match '\$resolvedVersion = \(& \$resolvedPath --version\)'

        $installIndex = $workflow.IndexOf('- name: Install winget-create')
        $installSection = $workflow.Substring($installIndex)
        $installSection = $installSection.Substring(0, $installSection.IndexOf('- name: Validate Winget manifest'))
        $installSection | Should Match ([regex]::Escape('$global:LASTEXITCODE = 0'))
        ($installSection.IndexOf('$global:LASTEXITCODE = 0') -gt $installSection.IndexOf('--version 2>&1')) | Should Be $true
        $installSection.TrimEnd() | Should Match '(?m)^\s+exit 0$'

        $cleanupIndex = $workflow.IndexOf('Cleanup Winget install test')
        $cleanupSection = $workflow.Substring($cleanupIndex)
        $cleanupSection = $cleanupSection.Substring(0, $cleanupSection.IndexOf('- name: Submit to Winget'))
        $cleanupSection | Should Match 'winget list'
        $cleanupSection | Should Match '\$noApplicationsFound = -1978335212'
        $cleanupSection | Should Match 'nothing to clean up'
        $cleanupSection | Should Match 'winget list failed with exit code'
        ($cleanupSection.IndexOf('winget list') -lt $cleanupSection.IndexOf('winget uninstall')) | Should Be $true
    }

    It 'uses one guarded manifest root and exact versioned portable inputs end to end' {
        $workflow | Should Match '(?m)^\s+WINGET_MANIFEST_ROOT: \.build\\AppPackages\\winget-manifest\r?$'
        $workflow | Should Match 'OutputDir = \$env:WINGET_MANIFEST_ROOT'
        $workflow | Should Match 'Get-ChildItem \$env:WINGET_MANIFEST_ROOT'
        $workflow | Should Match 'winget validate --manifest \$env:WINGET_MANIFEST_ROOT'
        $workflow | Should Match 'winget install[\s\S]+?--manifest \$env:WINGET_MANIFEST_ROOT'
        $workflow | Should Match 'Join-Path \$env:GITHUB_WORKSPACE \$env:WINGET_MANIFEST_ROOT'
        $workflow | Should Match '(?s)WINGETCREATE_PATH submit.+?\$env:WINGET_MANIFEST_ROOT'
        $workflow | Should Match '\$outFile = \$assetName'
        $workflow | Should Match 'RedSalamander-\$version-x64-Portable\.zip'
        $workflow | Should Match 'RedSalamander-\$version-ARM64-Portable\.zip'
        $workflow | Should Not Match 'RedSalamander-(x64|ARM64)\.zip'

        $buildScript = Get-Content -LiteralPath (Join-Path $repoRoot 'build.ps1') -Raw
        $buildScript | Should Match 'RedSalamander-\$packageVersion-x64-Portable\.zip'
        $buildScript | Should Match 'RedSalamander-\$packageVersion-ARM64-Portable\.zip'
        $buildScript | Should Not Match 'Get-ChildItem.+?RedSalamander-\*-x64-Portable\.zip'
    }

    It 'submits a detailed checklist-aligned pull request after validation and install testing' {
        $workflow | Should Match ([regex]::Escape('$pullRequestTitle = "Update: RedSalamanders.RedSalamander to $version"'))
        $workflow | Should Not Match 'search/issues|in:title'
        $workflow | Should Match 'Import-Module \$env:PUBLICATION_MODULE -Force'
        $workflow | Should Match 'Invoke-RSWingetPublication'
        $workflow | Should Match '\$env:WINGETCREATE_PATH submit[\s\S]{0,260}--prtitle \$InitialTitle'
        $publicationModule | Should Match 'RS-WINGET-PUBLICATION:'
        $publicationModule | Should Match 'ExpectedAuthor'
        $publicationModule | Should Match 'ExpectedHeadRepository'
        $publicationModule | Should Match 'ManifestHashes'
        $publicationModule | Should Match 'BranchPattern'
        $publicationModule | Should Match "Kind = 'Conflict'"
        $publicationModule | Should Match "Kind = 'Owned'"
        $publicationModule | Should Match 'PatchPullRequest'
        $workflow | Should Match '## 🧪 Automated validation'
        $workflow | Should Match 'winget validate --manifest \.build\\AppPackages\\winget-manifest'
        $workflow | Should Match 'winget install --manifest \.build\\AppPackages\\winget-manifest'
        $workflow | Should Match 'winget list --exact --id RedSalamanders\.RedSalamander'
        $workflow | Should Match 'Manifest conforms to the \[1\.12 schema\]'
        $workflow | Should Match 'pull_request_url=\$\(\$publication\.Url\)'
    }

    It 'keeps event data inert and verifies one pinned publisher before exposing its token' {
        $workflow | Should Match 'REQUESTED_VERSION:\s*\$\{\{ inputs\.version \|\| github\.event\.inputs\.version \}\}'
        $workflow | Should Match 'RELEASE_TAG:\s*\$\{\{ github\.event\.release\.tag_name \}\}'
        $workflow | Should Not Match '\$rawVersion\s*=\s*"\$\{\{'
        $workflow | Should Match 'Microsoft\.WingetCreate --version \$expectedVersion'
        $workflow | Should Match '\$expectedVersion = "1\.12\.8\.0"'
        $workflow | Should Match 'WINGET_CREATE_GITHUB_TOKEN:\s*\$\{\{ secrets\.WINGET_TOKEN \}\}'
        $workflow | Should Not Match '--token \$env:'
        $workflow | Should Not Match '\$output\s*\|\s*ForEach-Object\s*\{\s*Write-Host'

        $verifyIndex = $workflow.IndexOf('Resolved winget-create version')
        $secretIndex = $workflow.IndexOf('WINGET_CREATE_GITHUB_TOKEN:')
        ($verifyIndex -ge 0 -and $verifyIndex -lt $secretIndex) | Should Be $true
    }

    It 'names the token scopes and redacted WingetCreate output instead of a bare submit exit code' {
        # Scope names are safe to log and make an unsupported token shape visible before submission.
        $workflow | Should Match 'api\.github\.com/user" -ResponseHeadersVariable tokenResponseHeaders'
        $workflow | Should Match "-ieq 'X-OAuth-Scopes'"
        $workflow | Should Match "notcontains 'public_repo'"
        $workflow | Should Not Match 'Write-Host[^\n]*\$env:WINGET_CREATE_GITHUB_TOKEN'
        $workflow | Should Not Match 'Write-Host[^\n]*\$tokenResponseHeaders\b'

        # The raw output stays out of the workflow; the module redacts it and bounds it before it can
        # travel with the failure that stops publication.
        $publicationModule | Should Match 'function Get-RSWingetSubmitDiagnostics'
        $publicationModule | Should Match 'gh\[pousr\]_'
        $publicationModule | Should Match 'github_pat_'
        $publicationModule | Should Match 'WingetCreate output \(token-shaped values redacted\)'
        $publicationModule | Should Match '\$submitDiagnostics = @\(Get-RSWingetSubmitDiagnostics -SubmitResult \$submitResult\)'
        $publicationModule | Should Not Match "'Get-RSWingetSubmitDiagnostics'"
    }

    It 'treats upstream master as the authority for an already published version' {
        # The PR list covers only recently updated PRs; a merged publication drops out within hours
        # and a fresh submission then opens an empty PR that upstream closes.
        $workflow | Should Match 'GetUpstreamManifest = \$getUpstreamManifest'
        $workflow | Should Match 'api\.github\.com/repos/microsoft/winget-pkgs/git/ref/heads/master'
        $workflow | Should Match 'contents/\$encodedPrefix\?ref=\$ref'
        $workflow | Should Match 'if \(\$status -ne 404\) \{ throw \}'
        $workflow | Should Match '"published=\$\(\[bool\]\$publication\.Published\)" >> \$env:GITHUB_OUTPUT'
        $workflow | Should Match 'steps\.winget_submission\.outputs\.published'
        $publicationModule | Should Match "-Name 'GetUpstreamManifest'"
        $publicationModule | Should Match "Kind = 'Published'"
        $publicationModule | Should Match 'function New-RSWingetPublishedResult'
        $publicationModule | Should Match 'the changed-file list is empty'
    }

    It 'reads changed files only for candidate pull requests' {
        # One changed-file request per listed upstream PR turned each state check into ~100 calls; the
        # scan now inspects only the token owner's PRs and PRs naming the package or carrying the marker.
        $publicationModule | Should Match 'function Test-RSWingetPullRequestCandidate'
        $publicationModule | Should Match "@\(@\('title'\), @\('body'\), @\('head', 'ref'\)\)"
        $publicationModule | Should Match '(?s)-not \$hasMarker -and\s+-not \(Test-RSWingetPullRequestCandidate[^\n]*\)\) \{\s+continue'
        $publicationModule | Should Not Match "'Test-RSWingetPullRequestCandidate'"

        # The waits, not the scan, now carry GitHub's eventual-consistency window for a fresh PR.
        $workflow | Should Match 'Start-Sleep -Seconds \(\[Math\]::Min\(15 \* \$Attempt, 60\)\)'
        $publicationModule | Should Match 'Start-Sleep -Seconds \(\[Math\]::Min\(15 \* \$Attempt, 60\)\)'
        $publicationModule | Should Match "manifest content is not yet visible: '\`$path' at"
        $publicationModule | Should Match 'files are not yet visible: \$\(\$candidate\.FileListFailure\)'
    }

    It 'compares canonical manifest content and stamps a deterministic ReleaseDate so reruns resume' {
        # wingetcreate re-serializes the committed manifests, so byte hashes never match the generated files.
        $workflow | Should Match 'GetManifestContentHash = \$getManifestContentHash'
        $workflow | Should Match '(?s)\$getManifestContentHash = \{.+?return Get-RSWingetManifestContentHash -Bytes \$bytes'
        $workflow | Should Not Match 'GetFileSha256'
        $workflow | Should Not Match 'ToHexString\(\[Security\.Cryptography\.SHA256\]::HashData\(\$bytes\)\)'
        $publicationModule | Should Match 'function ConvertTo-RSWingetCanonicalManifest'
        $publicationModule | Should Match 'function Get-RSWingetManifestContentHash'
        $publicationModule | Should Match "'Get-RSWingetManifestContentHash'"
        $publicationModule | Should Match '\$manifestHashes\[\$remotePath\] = Get-RSWingetManifestContentHash -Bytes'
        $publicationModule | Should Match "-Name 'GetManifestContentHash'"
        $publicationModule | Should Not Match 'GetFileSha256'

        # The identity marker embeds the manifest hashes, so ReleaseDate must not drift with the rerun day.
        $workflow | Should Match '\$releaseDate = Get-RSReleaseDate \$release\.published_at'
        $workflow | Should Match '"release_date=\$releaseDate" >> \$env:GITHUB_OUTPUT'
        $workflow | Should Match 'RELEASE_DATE:\s*\$\{\{ steps\.portable\.outputs\.release_date \}\}'
        $workflow | Should Match '\(Get-Command \$generator\)\.Parameters\.ContainsKey\(''ReleaseDate''\)'
        $workflow | Should Match '\$generatorArguments\.ReleaseDate = \$env:RELEASE_DATE'
        $workflow | Should Match '& \$generator @generatorArguments'

        $generator = Get-Content -LiteralPath (Join-Path $repoRoot 'Installer\winget\generate-manifest.ps1') -Raw
        $generator | Should Match '(?m)^\s+\[string\]\$ReleaseDate = ''''\r?$'
        $generator | Should Match 'function Resolve-WingetReleaseDate'
        $generator | Should Match "TryParseExact\(\s*\`$CandidateDate,\s*'yyyy-MM-dd'"
        $generator | Should Match '\$ReleaseDate = Resolve-WingetReleaseDate -CandidateDate \$ReleaseDate'
        $generator | Should Not Match '\$ReleaseDate = Get-Date -Format "yyyy-MM-dd"'
        $generator | Should Match '(?m)^\.PARAMETER ReleaseDate\r?$'
    }

    It 'derives the release date from the deserialized published_at value the GitHub API returns' {
        # Invoke-RestMethod turns "2026-09-17T21:07:11Z" into a [datetime] (Kind Utc); the v7.0.60
        # publication stopped because the workflow matched that object's culture text against the
        # ISO pattern. Execute the workflow's own function against every shape it can receive.
        $functionMatch = [regex]::Match($workflow, '(?ms)^ {10}function Get-RSReleaseDate\(\[object\]\$PublishedAt\) \{.*?^ {10}\}\r?$')
        $functionMatch.Success | Should Be $true
        . ([scriptblock]::Create($functionMatch.Value))

        (Get-RSReleaseDate ([datetime]::new(2026, 9, 17, 21, 7, 11, [DateTimeKind]::Utc))) | Should Be '2026-09-17'
        (Get-RSReleaseDate ([datetime]::new(2026, 9, 17, 23, 30, 0, [DateTimeKind]::Unspecified))) | Should Be '2026-09-17'
        $lateLocal = [datetime]::new(2026, 9, 17, 23, 30, 0, [DateTimeKind]::Utc).ToLocalTime()
        (Get-RSReleaseDate $lateLocal) | Should Be '2026-09-17'
        (Get-RSReleaseDate '2026-09-17T21:07:11Z') | Should Be '2026-09-17'
        (Get-RSReleaseDate '2026-09-17T23:30:00+02:00') | Should Be '2026-09-17'
        (Get-RSReleaseDate '2026-09-18T01:30:00+02:00') | Should Be '2026-09-17'
        { Get-RSReleaseDate '' } | Should Throw 'did not report a published_at timestamp'
        { Get-RSReleaseDate $null } | Should Throw 'did not report a published_at timestamp'
        { Get-RSReleaseDate 'yesterday' } | Should Throw 'not a recognizable timestamp'
    }

    It 'supports a serialized direct release handoff' {
        $workflow | Should Match '(?ms)workflow_call:\s+inputs:\s+version:'
        $workflow | Should Match "(?ms)concurrency:\s+group:\s+winget-RedSalamanders\.RedSalamander-\$\{\{ github\.event_name == 'release'.*?format\('v\{0\}', inputs\.version \|\| github\.event\.inputs\.version\) \}\}"
        $workflow | Should Match 'cancel-in-progress:\s*false'
        $workflow | Should Match 'ref:\s*\$\{\{ github\.workflow_sha \}\}'
        $workflow | Should Match 'Tools/Modules/Packaging/WingetPrPublication\.psm1'
    }
}

Describe 'Winget manifest template' {
    BeforeAll {
        $installerTemplatePath = Join-Path $repoRoot 'Installer\winget\templates\RedSalamanders.RedSalamander.installer.yaml'
        $installerTemplate = Get-Content -Path $installerTemplatePath -Raw
        $msixManifestPath = Join-Path $repoRoot 'Installer\msix\Package.appxmanifest'
        $msixManifest = Get-Content -Path $msixManifestPath -Raw
        $msixProjectPath = Join-Path $repoRoot 'Installer\msix\RedSalamanderInstaller.wapproj'
        $msixProject = Get-Content -Path $msixProjectPath -Raw
        $directoryBuildPropsPath = Join-Path $repoRoot 'Directory.Build.props'
        $directoryBuildProps = Get-Content -Path $directoryBuildPropsPath -Raw
        $zipScriptPath = Join-Path $repoRoot 'Installer\zip\build-zip.ps1'
        $zipScript = Get-Content -Path $zipScriptPath -Raw
        $minimumOsHeaderPath = Join-Path $repoRoot 'Common\MinimumOsVersion.h'
        $minimumOsHeader = Get-Content -Path $minimumOsHeaderPath -Raw
        $minimumOsSourcePath = Join-Path $repoRoot 'Common\MinimumOsVersion.cpp'
        $minimumOsSource = Get-Content -Path $minimumOsSourcePath -Raw
    }

    It 'declares the dependency-free launcher aliases as portable alias targets on the Windows 11 build 22000.2600 floor' {
        $installerTemplate | Should Match '(?m)^MinimumOSVersion:\s+10\.0\.22000\.2600\r?$'
        $installerTemplate | Should Not Match '10\.0\.26100\.0'
        $installerTemplate | Should Not Match '10\.0\.19041\.0'
        $msixManifest | Should Match 'MinVersion="10\.0\.22000\.2600"'
        $msixProject | Should Match '<TargetPlatformMinVersion>10\.0\.22000\.2600</TargetPlatformMinVersion>'
        $directoryBuildProps | Should Match '<WindowsTargetPlatformVersion>10\.0\.26100\.0</WindowsTargetPlatformVersion>'
        $directoryBuildProps | Should Match 'NTDDI_VERSION=NTDDI_WIN11_GE'
        $minimumOsHeader | Should Match 'kMinimumWindowsBuildNumber\s*=\s*22000'
        $minimumOsHeader | Should Match 'kMinimumWindowsBuildRevision\s*=\s*2600'
        $minimumOsSource | Should Match 'TryGetWindowsBuildRevision'
        $minimumOsSource | Should Match 'CurrentVersion'
        $zipScript | Should Match 'Windows 11 build 22000\.2600'
        [regex]::Matches($installerTemplate, '(?m)^NestedInstallerFiles:\r?$').Count | Should Be 1
        [regex]::Matches($installerTemplate, '(?m)^\s+- RelativeFilePath: RedLauncher\.exe\r?\n\s+PortableCommandAlias: RedSalamander\r?$').Count | Should Be 1
        [regex]::Matches($installerTemplate, '(?m)^\s+- RelativeFilePath: red\.exe\r?\n\s+PortableCommandAlias: red\r?$').Count | Should Be 1
        $installerTemplate | Should Match '(?ms)^Commands:\s*\r?\n\s+- RedSalamander\r?\n\s+- red\r?\n'
    }

    It 'marks the launcher executables as launch files in installation metadata' {
        $installerTemplate | Should Match '(?ms)^InstallationMetadata:\s*\r?\n\s+Files:\s*\r?\n\s+- RelativeFilePath: RedLauncher\.exe\r?\n\s+FileType: launch\r?\n\s+DisplayName: RedSalamander'
        $installerTemplate | Should Match '(?ms)^InstallationMetadata:.*\r?\n\s+- RelativeFilePath: red\.exe\r?\n\s+FileType: launch\r?\n\s+DisplayName: red'
    }

    It 'copies RedLauncher.exe and its red.exe alias into the portable ZIP root' {
        $zipScript | Should Match 'Copy-Item \(Join-Path \$BuildOutputDir "RedLauncher\.exe"\) \$TempDir'
        $zipScript | Should Match 'Copy-Item \(Join-Path \$BuildOutputDir "red\.exe"\) \$TempDir'
        $zipScript | Should Not Match 'RedLauncherConsole\.exe'
    }

    It 'preserves package subdirectories for plugins, language satellites, and themes' {
        $zipScript | Should Match '\$LangSource = Join-Path \$BuildOutputDir "Lang"'
        $zipScript | Should Match 'Copy-Item \$LangSource -Destination \$LangDest -Recurse -Force'

        $catchAllInclude = 'Content Include="..\..\.build\$(Platform)\$(Configuration)\**\*"'
        $catchAllIndex = $msixProject.IndexOf($catchAllInclude, [System.StringComparison]::Ordinal)
        $catchAllIndex | Should Not Be -1
        $catchAllEnd = $msixProject.IndexOf('>', $catchAllIndex)
        $catchAllLine = $msixProject.Substring($catchAllIndex, $catchAllEnd - $catchAllIndex)

        foreach ($folder in @('Plugins', 'Lang', 'Themes')) {
            $include = 'Content Include="..\..\.build\$(Platform)\$(Configuration)\{0}\**\*"' -f $folder
            $exclude = '..\..\.build\$(Platform)\$(Configuration)\{0}\**\*' -f $folder
            $packagePath = '<PackagePath>{0}\%(RecursiveDir)%(Filename)%(Extension)</PackagePath>' -f $folder
            $link = '<Link>{0}\%(RecursiveDir)%(Filename)%(Extension)</Link>' -f $folder
            $targetPath = '<TargetPath>{0}\%(RecursiveDir)%(Filename)%(Extension)</TargetPath>' -f $folder

            $msixProject.Contains($include) | Should Be $true
            $catchAllLine.Contains($exclude) | Should Be $true
            $msixProject.Contains($packagePath) | Should Be $true
            $msixProject.Contains($link) | Should Be $true
            $msixProject.Contains($targetPath) | Should Be $true
        }
    }
}

Describe 'RedLauncher project' {
    BeforeAll {
        $launcherProjectPath = Join-Path $repoRoot 'RedLauncher\RedLauncher.vcxproj'
        $consoleLauncherProjectPath = Join-Path $repoRoot 'RedLauncher\RedLauncherConsole.vcxproj'
        $launcherManifestPath = Join-Path $repoRoot 'RedLauncher\res\exe.manifest'
        $launcherSourcePath = Join-Path $repoRoot 'RedLauncher\Main.cpp'
        $redSalamanderSourcePath = Join-Path $repoRoot 'RedSalamander\RedSalamander.cpp'
        $monitorSourcePath = Join-Path $repoRoot 'RedSalamanderMonitor\RedSalamanderMonitor.cpp'
        $configureSourcePath = Join-Path $repoRoot 'RedConfigure\Main.cpp'
        $searchServiceSourcePath = Join-Path $repoRoot 'RedSalamanderSearchService\Main.cpp'
        $solutionPath = Join-Path $repoRoot 'RedSalamander.sln'
        $solution = Get-Content -Path $solutionPath -Raw
        $buildScriptPath = Join-Path $repoRoot 'build.ps1'
        $buildScript = Get-Content -Path $buildScriptPath -Raw
    }

    It 'is a first-party static-CRT detached console executable project for the WinGet alias' {
        Test-Path $launcherProjectPath | Should Be $true
        Test-Path $consoleLauncherProjectPath | Should Be $false
        Test-Path $launcherManifestPath | Should Be $true

        $launcherProject = Get-Content -Path $launcherProjectPath -Raw
        $launcherManifest = Get-Content -Path $launcherManifestPath -Raw
        $launcherProject | Should Match '<ClCompile Include="Main\.cpp" />'
        $launcherProject | Should Match '<SubSystem>Console</SubSystem>'
        $launcherProject | Should Match '<AdditionalManifestFiles>\$\(ProjectDir\)res\\exe\.manifest</AdditionalManifestFiles>'
        $launcherManifest | Should Match 'consoleAllocationPolicy'
        $launcherManifest | Should Match '>detached<'
        $launcherProject | Should Not Match 'REDLAUNCHER_GUI'
        $launcherProject | Should Not Match 'REDLAUNCHER_CONSOLE'
        $launcherProject | Should Match '<RuntimeLibrary>MultiThreaded</RuntimeLibrary>'
        $launcherProject | Should Match '<RuntimeLibrary>MultiThreadedDebug</RuntimeLibrary>'
        $launcherProject | Should Match '<Target Name="CopyRedLauncherAlias"'
        $launcherProject | Should Match 'SourceFiles="\$\(TargetPath\)"'
        $launcherProject | Should Match 'DestinationFiles="\$\(OutDir\)red\.exe"'
        $launcherProject | Should Not Match '<ProjectReference'
        $launcherProject | Should Not Match 'Common\.vcxproj'
        $solution | Should Not Match 'RedLauncherConsole'
        $buildScript | Should Not Match 'RedLauncherConsole'
    }

    It 'resolves the WinGet alias symlink before launching RedSalamander.exe' {
        Test-Path $launcherSourcePath | Should Be $true

        $launcherSource = Get-Content -Path $launcherSourcePath -Raw
        $launcherSource | Should Match 'GetFinalPathNameByHandleW'
        $launcherSource | Should Match 'RedSalamander\.exe'
        $launcherSource | Should Match 'CreateProcessW'
        $launcherSource | Should Match 'RtlGetVersion'
        $launcherSource | Should Match 'TryGetWindowsBuildRevision'
        $launcherSource | Should Match 'kMinimumWindowsBuildNumber\s*=\s*22000'
        $launcherSource | Should Match 'kMinimumWindowsBuildRevision\s*=\s*2600'
        $launcherSource | Should Match 'Windows 11 build 22000\.2600'
        (Get-Content -Path $redSalamanderSourcePath -Raw) | Should Match 'EnsureCurrentWindowsVersionSupported'
        (Get-Content -Path $monitorSourcePath -Raw) | Should Match 'EnsureCurrentWindowsVersionSupported'
        (Get-Content -Path $configureSourcePath -Raw) | Should Match 'EnsureCurrentWindowsVersionSupported'
        (Get-Content -Path $searchServiceSourcePath -Raw) | Should Match 'EnsureCurrentWindowsVersionSupported'
    }

    It 'waits only for foreground self-test invocations' {
        Test-Path $launcherSourcePath | Should Be $true

        $launcherSource = Get-Content -Path $launcherSourcePath -Raw
        $launcherSource | Should Match 'ShouldWaitForTargetExit'
        $launcherSource | Should Match 'wmain'
        $launcherSource | Should Not Match 'wWinMain'
        $launcherSource | Should Match '--selftest'
        $launcherSource | Should Match '--compare-selftest'
        $launcherSource | Should Match '--commands-selftest'
        $launcherSource | Should Match '--fileops-selftest'
        $launcherSource | Should Match '--selftest-list-cases'
        $launcherSource | Should Not Match 'argc\s*>\s*1'
        $launcherSource | Should Not Match '--etw'
        $launcherSource | Should Not Match '--perf'
    }

}

Describe 'VC runtime ZIP helper' {
    BeforeAll {
        . (Join-Path $repoRoot 'Installer\zip\VcRuntime.ps1')
    }

    It 'selects the latest non-OneCore MSVC redist directory for the requested architecture' {
        $caseId = ([Guid]::NewGuid().ToString('N')).Substring(0, 8)
        $tempRoot = New-RSTestSandboxScratchDirectory `
            -RepoRoot $repoRoot `
            -Harness 'tools-pester' `
            -Case "vc-redist-$caseId"
        try {
            $older = Join-Path $tempRoot 'VS\VC\Redist\MSVC\14.50.10000\x64\Microsoft.VC145.CRT'
            $newer = Join-Path $tempRoot 'VS\VC\Redist\MSVC\14.51.20000\x64\Microsoft.VC145.CRT'
            $oneCore = Join-Path $tempRoot 'VS\VC\Redist\MSVC\14.99.99999\onecore\x64\Microsoft.VC145.CRT'
            New-Item -ItemType Directory -Path $older, $newer, $oneCore -Force | Out-Null

            Get-RSVcRuntimeRedistDirectory -Platform x64 -SearchRoots @($tempRoot) | Should Be $newer
        } finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'copies app-local MSVC runtime DLLs and requires the core launch dependencies' {
        $caseId = ([Guid]::NewGuid().ToString('N')).Substring(0, 8)
        $tempRoot = New-RSTestSandboxScratchDirectory `
            -RepoRoot $repoRoot `
            -Harness 'tools-pester' `
            -Case "vc-copy-$caseId"
        $dest = Join-Path $tempRoot 'package'
        try {
            $redist = Join-Path $tempRoot 'VS\VC\Redist\MSVC\14.51.20000\arm64\Microsoft.VC145.CRT'
            New-Item -ItemType Directory -Path $redist, $dest -Force | Out-Null
            foreach ($dllName in @('msvcp140.dll', 'msvcp140_atomic_wait.dll', 'vcruntime140.dll', 'vcruntime140_1.dll', 'vccorlib140.dll')) {
                Set-Content -Path (Join-Path $redist $dllName) -Value $dllName -Encoding ASCII
            }

            $copied = Copy-RSVcRuntimeDependencies -Platform ARM64 -DestinationDir $dest -SearchRoots @($tempRoot)

            ($copied -contains 'msvcp140.dll') | Should Be $true
            ($copied -contains 'msvcp140_atomic_wait.dll') | Should Be $true
            ($copied -contains 'vcruntime140.dll') | Should Be $true
            ($copied -contains 'vcruntime140_1.dll') | Should Be $true
            Test-Path (Join-Path $dest 'vccorlib140.dll') | Should Be $true
        } finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'fails clearly when a required runtime DLL is missing' {
        $caseId = ([Guid]::NewGuid().ToString('N')).Substring(0, 8)
        $tempRoot = New-RSTestSandboxScratchDirectory `
            -RepoRoot $repoRoot `
            -Harness 'tools-pester' `
            -Case "vc-missing-$caseId"
        $dest = Join-Path $tempRoot 'package'
        try {
            $redist = Join-Path $tempRoot 'VS\VC\Redist\MSVC\14.51.20000\x64\Microsoft.VC145.CRT'
            New-Item -ItemType Directory -Path $redist, $dest -Force | Out-Null
            foreach ($dllName in @('msvcp140.dll', 'vcruntime140.dll', 'vcruntime140_1.dll')) {
                Set-Content -Path (Join-Path $redist $dllName) -Value $dllName -Encoding ASCII
            }

            { Copy-RSVcRuntimeDependencies -Platform x64 -DestinationDir $dest -SearchRoots @($tempRoot) } |
                Should Throw "Required MSVC runtime DLL 'msvcp140_atomic_wait.dll'"
        } finally {
            Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}
