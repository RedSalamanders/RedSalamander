Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function ConvertTo-RSTerminalLfTextIdentity {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Text,

        [string]$Path = ''
    )

    # This helper is intentionally Terminal-build-specific: its identity is
    # UTF-8 without BOM after CRLF/CR are mapped to LF, independent of checkout
    # line-ending policy.
    $normalizedText = $Text.Replace("`r`n", "`n").Replace("`r", "`n")
    $utf8Bytes = [Text.UTF8Encoding]::new($false).GetBytes($normalizedText)
    $sha256 = [Convert]::ToHexString(
        [Security.Cryptography.SHA256]::HashData($utf8Bytes)).ToLowerInvariant()
    return [pscustomobject]@{
        Path = $Path
        Text = $normalizedText
        Utf8Bytes = $utf8Bytes
        Sha256 = $sha256
    }
}

function Get-RSTerminalLfTextIdentity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$LiteralPath
    )

    $resolved = (Resolve-Path -LiteralPath $LiteralPath -ErrorAction Stop).Path
    return ConvertTo-RSTerminalLfTextIdentity `
        -Text ([IO.File]::ReadAllText($resolved)) `
        -Path $resolved
}

function Get-RSTerminalLfStringIdentity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Text
    )

    return ConvertTo-RSTerminalLfTextIdentity -Text $Text
}

Export-ModuleMember -Function @(
    'Get-RSTerminalLfStringIdentity',
    'Get-RSTerminalLfTextIdentity')
