<#
.SYNOPSIS
Runs an explicitly selected sealed Terminal Gate0 or Round4 test profile.

.DESCRIPTION
Validates every historical test file against the checked-in profile manifest
before invoking Pester. These profiles are archival and intentionally excluded
from ordinary CI and Full test runs.

.PARAMETER Profile
Selects the Gate0 or Round4 historical profile.

.PARAMETER PassThru
Returns a normalized result object for workflow automation.

.OUTPUTS
With -PassThru, a PSCustomObject containing the profile name, file count, and
Pester result counts. Without -PassThru, writes only a concise host summary.

.EXAMPLE
./Tools/TerminalEngine/Invoke-HistoricalTerminalTests.ps1 -Profile Round4

.NOTES
Requires Pester. Executes only the SHA-256-authenticated files declared by
HistoricalTerminalProfiles.json and may create the fixtures owned by those
historical tests. Throws on missing/drifting files, Pester failures, or recorded
Round4 count drift. It is not an ordinary CI/Full entrypoint.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet('Gate0', 'Round4')]
    [string]$Profile,

    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$manifestPath = Join-Path $PSScriptRoot 'HistoricalTerminalProfiles.json'
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
$profileDefinition = $manifest.profiles.PSObject.Properties[$Profile]
if ($null -eq $profileDefinition) {
    throw "Historical Terminal profile '$Profile' is not defined in $manifestPath."
}

$testDefinitions = @($manifest.historicalTests | Where-Object {
        @($_.profiles) -contains $Profile
    })
$expectedFileCount = [int]$profileDefinition.Value.expectedFileCount
if ($testDefinitions.Count -ne $expectedFileCount) {
    throw "Historical Terminal profile '$Profile' expected $expectedFileCount files but declares $($testDefinitions.Count)."
}

$testPaths = @()
foreach ($definition in $testDefinitions) {
    $testPath = Join-Path $repoRoot ([string]$definition.path).Replace('/', '\')
    if (-not (Test-Path -LiteralPath $testPath -PathType Leaf)) {
        throw "Historical Terminal profile '$Profile' is missing '$($definition.path)'."
    }

    $actualSha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $testPath).Hash.ToLowerInvariant()
    if ($actualSha256 -cne [string]$definition.sha256) {
        throw "Historical Terminal profile '$Profile' identity drifted for '$($definition.path)'."
    }
    $testPaths += $testPath
}

$pester = Get-Command Invoke-Pester -ErrorAction Stop
if ($pester.Parameters.ContainsKey('Path')) {
    $result = Invoke-Pester -Path $testPaths -PassThru
}
elseif ($pester.Parameters.ContainsKey('Script')) {
    $result = Invoke-Pester -Script $testPaths -PassThru
}
else {
    throw 'Invoke-Pester exposes neither -Path nor -Script.'
}

function Get-RSHistoricalPesterCount {
    param([Parameter(Mandatory)][string[]]$Names)

    foreach ($name in $Names) {
        $property = $result.PSObject.Properties[$name]
        if ($null -ne $property) {
            return [int]$property.Value
        }
    }
    return 0
}

$summary = [pscustomobject]@{
    Profile = $Profile
    TestFileCount = $testPaths.Count
    Passed = Get-RSHistoricalPesterCount -Names @('PassedCount', 'Passed')
    Failed = Get-RSHistoricalPesterCount -Names @('FailedCount', 'Failed')
    Skipped = Get-RSHistoricalPesterCount -Names @('SkippedCount', 'Skipped')
    Pending = Get-RSHistoricalPesterCount -Names @('PendingCount', 'Pending')
    Inconclusive = Get-RSHistoricalPesterCount -Names @('InconclusiveCount', 'Inconclusive')
}

if ($summary.Failed -ne 0) {
    throw "Historical Terminal profile '$Profile' failed $($summary.Failed) test(s)."
}

$expectedPassedProperty = $profileDefinition.Value.PSObject.Properties['expectedPassedCount']
if ($null -ne $expectedPassedProperty -and
    ($summary.Passed -ne [int]$expectedPassedProperty.Value -or
        $summary.Skipped -ne 0 -or
        $summary.Pending -ne 0 -or
        $summary.Inconclusive -ne 0)) {
    throw ("Historical Terminal profile '{0}' did not reproduce its recorded result: " +
        'passed={1} skipped={2} pending={3} inconclusive={4}.') -f `
        $Profile, $summary.Passed, $summary.Skipped, $summary.Pending, $summary.Inconclusive
}

if ($PassThru) {
    Write-Output $summary
}
else {
    Write-Host ("Historical Terminal {0}: files={1} passed={2}" -f `
            $summary.Profile, $summary.TestFileCount, $summary.Passed)
}
