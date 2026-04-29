Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$helperScript = Join-Path $repoRoot 'Tools\Versioning.ps1'

function New-RSTestVersionRepo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root
    )

    $commonDir = Join-Path $Root 'Common'
    New-Item -ItemType Directory -Path $commonDir -Force | Out-Null

    @'
// WARNING: minimal test header
#define VERSINFO_MAJOR 7
#define VERSINFO_MINORA 0
#define VERSINFO_MINORB 0
'@ | Set-Content -Path (Join-Path $commonDir 'Version.h') -Encoding ASCII
}

Describe 'Versioning helper' {
    BeforeAll {
        . $helperScript
    }

    It 'reuses the saved local build number across ordinary local builds' {
        $testRepo = Join-Path $TestDrive 'ReuseSavedLocalBuildNumber'
        New-RSTestVersionRepo -Root $testRepo

        $firstContext = Get-RSVersionContext -RepoRoot $testRepo -Configuration Debug -Platform x64
        Save-RSVersionContext -RepoRoot $testRepo -VersionContext $firstContext | Out-Null

        $statePath = Get-RSVersionStatePath -RepoRoot $testRepo
        (Get-Item $statePath).LastWriteTimeUtc = [DateTime]::UtcNow.AddDays(-1)

        $secondContext = Get-RSVersionContext -RepoRoot $testRepo -Configuration Release -Platform ARM64

        $secondContext.BuildNumber | Should Be $firstContext.BuildNumber
    }

    It 'allocates a new local build number when the saved context is missing' {
        $testRepo = Join-Path $TestDrive 'MissingSavedContextGetsNextNumber'
        New-RSTestVersionRepo -Root $testRepo

        $firstContext = Get-RSVersionContext -RepoRoot $testRepo -Configuration Debug -Platform x64
        Save-RSVersionContext -RepoRoot $testRepo -VersionContext $firstContext | Out-Null

        Remove-Item -LiteralPath (Get-RSVersionStatePath -RepoRoot $testRepo) -Force

        $secondContext = Get-RSVersionContext -RepoRoot $testRepo -Configuration Debug -Platform x64

        $secondContext.BuildNumber | Should Be ($firstContext.BuildNumber + 1)
    }
}
