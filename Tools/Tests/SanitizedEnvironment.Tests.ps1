Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$helperScript = Join-Path $repoRoot 'Tools\SanitizedEnvironment.ps1'

Describe 'SanitizedEnvironment helper' {
    BeforeAll {
        . $helperScript
    }

    It 'creates a child ProcessStartInfo without mutating the caller Path' {
        $originalPath = [System.Environment]::GetEnvironmentVariable('Path', 'Process')
        try {
            [System.Environment]::SetEnvironmentVariable('Path', 'C:\Temp\One;C:\Temp\Two', 'Process')

            $null = New-RSProcessStartInfo `
                -FilePath 'cmd.exe' `
                -Arguments @('/d', '/c', 'exit', '0')

            [System.Environment]::GetEnvironmentVariable('Path', 'Process') | Should Be 'C:\Temp\One;C:\Temp\Two'
        }
        finally {
            [System.Environment]::SetEnvironmentVariable('Path', $originalPath, 'Process')
        }
    }

    It 'creates a child environment with a canonical Path entry and extra variables' {
        $psi = New-RSProcessStartInfo `
            -FilePath 'cmd.exe' `
            -Arguments @('/d', '/c', 'exit', '0') `
            -AdditionalEnvironment @{ 'RS_TEST_FLAG' = '1' }

        $pathEntry = $psi.Environment.Keys | Where-Object { $_ -ceq 'Path' }
        $duplicatePathAlias = $psi.Environment.Keys | Where-Object { $_ -ceq 'PATH' }

        $pathEntry | Should Be 'Path'
        $duplicatePathAlias | Should BeNullOrEmpty
        $psi.Environment['Path'] | Should Not BeNullOrEmpty
        $psi.Environment['RS_TEST_FLAG'] | Should Be '1'
    }

    It 'preserves the effective process Path order while merging duplicate aliases' {
        $processEnvironment = [System.Collections.Specialized.OrderedDictionary]::new([System.StringComparer]::Ordinal)
        $processEnvironment.Add('PATH', 'C:\UpperOnly;C:\Shared')
        $processEnvironment.Add('Path', 'C:\LowerOnly;C:\Shared')
        $processEnvironment.Add('RS_TEST_FLAG', '1')

        $normalized = Get-RSNormalizedEnvironmentMap `
            -ProcessEnvironment $processEnvironment `
            -ProcessPathValue 'C:\EffectiveFirst;C:\Shared'

        $pathEntry = $normalized.Keys | Where-Object { $_ -ceq 'Path' }
        $duplicatePathAlias = $normalized.Keys | Where-Object { $_ -ceq 'PATH' }

        $pathEntry | Should Be 'Path'
        $duplicatePathAlias | Should BeNullOrEmpty
        $normalized['Path'] | Should Be 'C:\EffectiveFirst;C:\Shared;C:\UpperOnly;C:\LowerOnly'
        $normalized['RS_TEST_FLAG'] | Should Be '1'
    }

    It 'canonicalizes uppercase PATH overrides for child processes' {
        $psi = New-RSProcessStartInfo `
            -FilePath 'cmd.exe' `
            -Arguments @('/d', '/c', 'exit', '0') `
            -AdditionalEnvironment @{ 'PATH' = 'C:\Override' }

        $pathEntry = $psi.Environment.Keys | Where-Object { $_ -ceq 'Path' }
        $duplicatePathAlias = $psi.Environment.Keys | Where-Object { $_ -ceq 'PATH' }

        $pathEntry | Should Be 'Path'
        $duplicatePathAlias | Should BeNullOrEmpty
        $psi.Environment['Path'] | Should Be 'C:\Override'
    }

    It 'launches direct child processes with a sanitized Path key' {
        $scriptPath = Join-Path $TestDrive 'CaptureEnvironment.ps1'
        $outputPath = Join-Path $TestDrive 'environment.json'
        @'
param(
    [string]$OutputPath
)

$keys = @([System.Environment]::GetEnvironmentVariables('Process').Keys | Where-Object { $_ -ieq 'Path' })
$payload = [pscustomobject]@{
    keys = $keys
    path = [System.Environment]::GetEnvironmentVariable('Path', 'Process')
}
$payload | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
'@ | Set-Content -Path $scriptPath -Encoding ASCII

        $exitCode = Invoke-RSProcess `
            -FilePath 'powershell.exe' `
            -Arguments @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $scriptPath, $outputPath) `
            -WorkingDirectory $TestDrive `
            -AdditionalEnvironment @{ 'PATH' = 'C:\DirectChild' }

        $payload = Get-Content -Path $outputPath -Raw | ConvertFrom-Json

        $exitCode | Should Be 0
        @($payload.keys) | Should Be @('Path')
        $payload.path | Should Be 'C:\DirectChild'
    }
}
