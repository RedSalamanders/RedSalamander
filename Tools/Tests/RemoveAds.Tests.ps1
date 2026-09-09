Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'Tools\Modules\Testing\TestRunPlan.psm1') -Force -ErrorAction Stop

$canonicalScript = Join-Path $repoRoot 'Tools\remove-ads.ps1'
$windowsPowerShell = (Get-Command powershell.exe -ErrorAction Stop).Source

function New-RSAdsScratch {
    return (New-RSTestSandboxScratchDirectory `
            -RepoRoot $repoRoot `
            -Harness 'tools-pester' `
            -Case ('remove-ads-' + [guid]::NewGuid().ToString('N')))
}

function Invoke-RSAdsProcess {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScriptPath,

        [string[]]$Arguments = @()
    )

    $output = @(& $windowsPowerShell -NoProfile -ExecutionPolicy Bypass -File $ScriptPath @Arguments 2>&1)
    return [pscustomobject]@{
        ExitCode = $LASTEXITCODE
        Output = ($output | ForEach-Object { $_.ToString() }) -join [Environment]::NewLine
    }
}

function Invoke-RSAdsDeniedRemovalProcess {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $previousScript = [Environment]::GetEnvironmentVariable('RS_REMOVE_ADS_TEST_SCRIPT', 'Process')
    $previousPath = [Environment]::GetEnvironmentVariable('RS_REMOVE_ADS_TEST_PATH', 'Process')
    try {
        [Environment]::SetEnvironmentVariable('RS_REMOVE_ADS_TEST_SCRIPT', $canonicalScript, 'Process')
        [Environment]::SetEnvironmentVariable('RS_REMOVE_ADS_TEST_PATH', $Path, 'Process')
        $command = @'
function global:Remove-Item {
    [CmdletBinding()]
    param(
        [string]$LiteralPath,
        [string]$Stream,
        [switch]$Force
    )
    if (-not [string]::IsNullOrEmpty($Stream)) {
        throw [UnauthorizedAccessException]::new('Injected inaccessible-stream failure.')
    }
    Microsoft.PowerShell.Management\Remove-Item @PSBoundParameters
}
& $env:RS_REMOVE_ADS_TEST_SCRIPT -Path $env:RS_REMOVE_ADS_TEST_PATH
exit $LASTEXITCODE
'@
        $output = @(& $windowsPowerShell -NoProfile -ExecutionPolicy Bypass -Command $command 2>&1)
        return [pscustomobject]@{
            ExitCode = $LASTEXITCODE
            Output = ($output | ForEach-Object { $_.ToString() }) -join [Environment]::NewLine
        }
    } finally {
        [Environment]::SetEnvironmentVariable('RS_REMOVE_ADS_TEST_SCRIPT', $previousScript, 'Process')
        [Environment]::SetEnvironmentVariable('RS_REMOVE_ADS_TEST_PATH', $previousPath, 'Process')
    }
}

function Set-RSAdsFixtureStream {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [string]$Value = 'fixture'
    )

    Set-Content -LiteralPath $Path -Stream $Name -Value $Value -Encoding UTF8 -ErrorAction Stop
}

function Get-RSAdsFixtureStreamNames {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return @(
        Get-Item -LiteralPath $Path -Stream * -Force -ErrorAction Stop |
            Where-Object { $_.Stream -ne ':$DATA' } |
            ForEach-Object { $_.Stream }
    )
}

Describe 'ADS removal tools' {
    BeforeAll {
        $probeRoot = New-RSAdsScratch
        $probeFile = Join-Path $probeRoot 'probe.txt'
        Set-Content -LiteralPath $probeFile -Value 'primary' -Encoding UTF8
        $script:adsSupported = $true
        try {
            Set-RSAdsFixtureStream -Path $probeFile -Name 'RedSalamander.Probe'
            $script:adsSupported = (Get-RSAdsFixtureStreamNames -Path $probeFile) -contains 'RedSalamander.Probe'
        } catch {
            $script:adsSupported = $false
            Write-Warning "Named-stream runtime cases are skipped because this filesystem/provider does not support ADS: $($_.Exception.Message)"
        }
    }

    It 'discovers the canonical command syntax without duplicate common parameters' {
        $canonicalSyntax = (Get-Command $canonicalScript -Syntax | Out-String)
        $canonicalSyntax | Should Match '\[-WhatIf\]'
        ([regex]::Matches($canonicalSyntax, '\[-WhatIf\]')).Count | Should Be 1
    }

    It 'removes a named stream and then all remaining streams with truthful counts' -Skip:(-not $script:adsSupported) {
        $scratch = New-RSAdsScratch
        $file = Join-Path $scratch 'named.txt'
        Set-Content -LiteralPath $file -Value 'primary' -Encoding UTF8
        Set-RSAdsFixtureStream -Path $file -Name 'Zone.Identifier'
        Set-RSAdsFixtureStream -Path $file -Name 'RedSalamander.Metadata'

        $named = Invoke-RSAdsProcess -ScriptPath $canonicalScript -Arguments @(
            '-Path', $file, '-StreamName', 'zone.identifier')
        $named.ExitCode | Should Be 0
        $named.Output | Should Match 'Matched: 1\. Would remove: 0\. Removed: 1\. Skipped: 0\. Failed: 0\.'
        ((Get-RSAdsFixtureStreamNames -Path $file) -contains 'Zone.Identifier') | Should Be $false
        ((Get-RSAdsFixtureStreamNames -Path $file) -contains 'RedSalamander.Metadata') | Should Be $true

        $all = Invoke-RSAdsProcess -ScriptPath $canonicalScript -Arguments @('-Path', $file)
        $all.ExitCode | Should Be 0
        $all.Output | Should Match 'Matched: 1\. Would remove: 0\. Removed: 1\. Skipped: 0\. Failed: 0\.'
        @(Get-RSAdsFixtureStreamNames -Path $file).Count | Should Be 0
    }

    It 'honors recursion and stream-name filtering' -Skip:(-not $script:adsSupported) {
        $scratch = New-RSAdsScratch
        $nested = Join-Path $scratch 'nested'
        New-Item -ItemType Directory -Path $nested | Out-Null
        $rootFile = Join-Path $scratch 'root.txt'
        $nestedFile = Join-Path $nested 'nested.txt'
        Set-Content -LiteralPath $rootFile -Value 'root' -Encoding UTF8
        Set-Content -LiteralPath $nestedFile -Value 'nested' -Encoding UTF8
        Set-RSAdsFixtureStream -Path $rootFile -Name 'Zone.Identifier'
        Set-RSAdsFixtureStream -Path $nestedFile -Name 'Zone.Identifier'
        Set-RSAdsFixtureStream -Path $nestedFile -Name 'Keep.Stream'

        $flat = Invoke-RSAdsProcess -ScriptPath $canonicalScript -Arguments @(
            '-Path', $scratch, '-StreamName', 'Zone.Identifier')
        $flat.ExitCode | Should Be 0
        $flat.Output | Should Match 'Matched: 1\.'
        ((Get-RSAdsFixtureStreamNames -Path $nestedFile) -contains 'Zone.Identifier') | Should Be $true

        $recursive = Invoke-RSAdsProcess -ScriptPath $canonicalScript -Arguments @(
            '-Path', $scratch, '-Recurse', '-StreamName', 'Zone.Identifier')
        $recursive.ExitCode | Should Be 0
        $recursive.Output | Should Match 'Matched: 1\. Would remove: 0\. Removed: 1\. Skipped: 0\. Failed: 0\.'
        ((Get-RSAdsFixtureStreamNames -Path $nestedFile) -contains 'Zone.Identifier') | Should Be $false
        ((Get-RSAdsFixtureStreamNames -Path $nestedFile) -contains 'Keep.Stream') | Should Be $true
    }

    It 'previews matches without modifying streams' -Skip:(-not $script:adsSupported) {
        $scratch = New-RSAdsScratch
        $file = Join-Path $scratch 'preview.txt'
        Set-Content -LiteralPath $file -Value 'primary' -Encoding UTF8
        Set-RSAdsFixtureStream -Path $file -Name 'Zone.Identifier'

        $preview = Invoke-RSAdsProcess -ScriptPath $canonicalScript -Arguments @(
            '-Path', $file, '-WhatIf')
        $preview.ExitCode | Should Be 0
        $preview.Output | Should Match 'Matched: 1\. Would remove: 1\. Removed: 0\. Skipped: 0\. Failed: 0\.'
        ((Get-RSAdsFixtureStreamNames -Path $file) -contains 'Zone.Identifier') | Should Be $true
    }

    It 'reports missing paths as failures and exits nonzero' {
        $missing = Join-Path (New-RSAdsScratch) 'missing'
        $result = Invoke-RSAdsProcess -ScriptPath $canonicalScript -Arguments @('-Path', $missing)
        $result.ExitCode | Should Be 1
        $result.Output | Should Match 'Cannot resolve ADS scan path'
        $result.Output | Should Match 'Matched: 0\. Would remove: 0\. Removed: 0\. Skipped: 0\. Failed: 1\.'
    }

    It 'surfaces inaccessible-file failures and preserves the stream' -Skip:(-not $script:adsSupported) {
        $scratch = New-RSAdsScratch
        $file = Join-Path $scratch 'locked.txt'
        Set-Content -LiteralPath $file -Value 'primary' -Encoding UTF8
        Set-RSAdsFixtureStream -Path $file -Name 'Zone.Identifier'

        $result = Invoke-RSAdsDeniedRemovalProcess -Path $file

        $result.ExitCode | Should Be 1
        $result.Output | Should Match 'Cannot (enumerate streams for|remove)'
        $result.Output | Should Match 'Failed: 1\.'
        ((Get-RSAdsFixtureStreamNames -Path $file) -contains 'Zone.Identifier') | Should Be $true
    }

}
