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
}
