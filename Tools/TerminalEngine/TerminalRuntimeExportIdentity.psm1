Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-RSTerminalExportSetIdentity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [string[]]$ExportNames
    )

    if ($ExportNames.Count -eq 0 -or
        @($ExportNames | Where-Object { [string]::IsNullOrEmpty($_) }).Count -ne 0) {
        throw 'Terminal runtime export names must be a non-empty set of named exports.'
    }

    $sorted = @($ExportNames | Sort-Object -CaseSensitive -Unique)
    if ($sorted.Count -ne $ExportNames.Count) {
        throw 'Terminal runtime export names contain a duplicate.'
    }

    $canonicalBytes = [Text.Encoding]::UTF8.GetBytes(($sorted -join "`n") + "`n")
    $sha256 = [Convert]::ToHexString(
        [Security.Cryptography.SHA256]::HashData($canonicalBytes)).ToLowerInvariant()
    return [pscustomobject]@{
        count = [uint32]$sorted.Count
        sha256 = $sha256
        names = $sorted
    }
}

Export-ModuleMember -Function Get-RSTerminalExportSetIdentity
