Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot (
        'Tools\TerminalEngine\TerminalEngineTextIdentity.psm1')) -Force

Describe 'Terminal-engine semantic text identity' {
    It 'gives CRLF, CR, and LF inputs one UTF-8 LF identity' {
        $lf = Join-Path $TestDrive 'lf.txt'
        $crlf = Join-Path $TestDrive 'crlf.txt'
        $cr = Join-Path $TestDrive 'cr.txt'
        [IO.File]::WriteAllText(
            $lf,
            "alpha`nbeta`n",
            [Text.UTF8Encoding]::new($false))
        [IO.File]::WriteAllText(
            $crlf,
            "alpha`r`nbeta`r`n",
            [Text.UTF8Encoding]::new($false))
        [IO.File]::WriteAllText(
            $cr,
            "alpha`rbeta`r",
            [Text.UTF8Encoding]::new($false))

        $lfIdentity = Get-RSTerminalLfTextIdentity -LiteralPath $lf
        foreach ($candidate in @($crlf, $cr)) {
            $identity = Get-RSTerminalLfTextIdentity -LiteralPath $candidate
            $identity.Text | Should Be $lfIdentity.Text
            $identity.Sha256 | Should Be $lfIdentity.Sha256
            ([Convert]::ToHexString($identity.Utf8Bytes)) |
                Should Be ([Convert]::ToHexString($lfIdentity.Utf8Bytes))
        }
        $lfIdentity.Sha256 | Should Match '^[0-9a-f]{64}$'
    }

    It 'rejects a substituted license identity while accepting line-ending changes' {
        $expected = Get-RSTerminalLfStringIdentity -Text "license`ntext`n"
        $crlf = Get-RSTerminalLfStringIdentity -Text "license`r`ntext`r`n"
        $substituted = Get-RSTerminalLfStringIdentity -Text "different`nlicense`n"

        $crlf.Sha256 | Should Be $expected.Sha256
        $substituted.Sha256 | Should Not Be $expected.Sha256
    }
}
