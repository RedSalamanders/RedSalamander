Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot (
        'Tools\TerminalEngine\TerminalRuntimeImportPolicy.psm1')) -Force

Describe 'Terminal runtime actual-import policy' {
    It 'accepts only the exact sorted module set from the parsed PE identity' {
        $identity = [pscustomobject]@{
            semantic = [pscustomobject]@{
                imports = @(
                    [pscustomobject]@{ module = 'kernel32.dll'; symbol = 'A' },
                    [pscustomobject]@{ module = 'kernel32.dll'; symbol = 'B' },
                    [pscustomobject]@{ module = 'ntdll.dll'; symbol = 'C' })
            }
        }
        $modules = Get-RSTerminalRuntimeImportModules -Identity $identity
        ($modules -join "`n") | Should Be "kernel32.dll`nntdll.dll"
        (Test-RSTerminalRuntimeImportModules `
                -Identity $identity `
                -ExpectedModules @('kernel32.dll', 'ntdll.dll')) |
            Should Be $true
        (Test-RSTerminalRuntimeImportModules `
                -Identity $identity `
                -ExpectedModules @('ntdll.dll', 'kernel32.dll')) |
            Should Be $false
        (Test-RSTerminalRuntimeImportModules `
                -Identity $identity `
                -ExpectedModules @('kernel32.dll')) |
            Should Be $false

        $identity.semantic.imports +=
            [pscustomobject]@{ module = 'unexpected.dll'; symbol = 'D' }
        (Test-RSTerminalRuntimeImportModules `
                -Identity $identity `
                -ExpectedModules @('kernel32.dll', 'ntdll.dll')) |
            Should Be $false
    }
}
