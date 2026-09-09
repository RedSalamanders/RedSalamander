$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$modulePath = Join-Path $repoRoot 'Tools\TerminalEngine\TerminalRuntimeExportIdentity.psm1'
Import-Module $modulePath -Force

Describe 'Terminal runtime exact export-set identity' {
    It 'is order-independent and rejects missing, additional, or duplicate exports' {
        $exact = Get-RSTerminalExportSetIdentity -ExportNames @('zeta', 'alpha')
        $reordered = Get-RSTerminalExportSetIdentity -ExportNames @('alpha', 'zeta')
        $missing = Get-RSTerminalExportSetIdentity -ExportNames @('alpha')
        $extra = Get-RSTerminalExportSetIdentity -ExportNames @('alpha', 'zeta', 'extra')

        $exact.count | Should Be 2
        $exact.sha256 | Should Be $reordered.sha256
        $missing.sha256 | Should Not Be $exact.sha256
        $extra.sha256 | Should Not Be $exact.sha256
        $duplicateRejected = $false
        try {
            Get-RSTerminalExportSetIdentity -ExportNames @('alpha', 'alpha') | Out-Null
        }
        catch [System.Management.Automation.RuntimeException] {
            $duplicateRejected = $true
            $_.Exception.Message | Should Match 'duplicate'
        }
        $duplicateRejected | Should Be $true
    }
}
