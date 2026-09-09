Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-RSTerminalRuntimeImportModules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [object]$Identity
    )

    return [string[]]@(
        $Identity.semantic.imports |
            ForEach-Object { [string]$_.module } |
            Sort-Object -CaseSensitive -Unique)
}

function Test-RSTerminalRuntimeImportModules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [object]$Identity,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [string[]]$ExpectedModules
    )

    $actualModules = Get-RSTerminalRuntimeImportModules -Identity $Identity
    return ($actualModules -join "`0") -ceq ($ExpectedModules -join "`0")
}

Export-ModuleMember -Function @(
    'Get-RSTerminalRuntimeImportModules',
    'Test-RSTerminalRuntimeImportModules')
