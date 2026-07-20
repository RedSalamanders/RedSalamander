#Requires -Version 5.1
<#
.SYNOPSIS
    Removes NTFS alternate data streams (ADS) from files.

.DESCRIPTION
    Scans files under the specified path and removes all alternate data streams
    (e.g. Zone.Identifier, SmartScreen, mark-of-the-web) that Windows attaches
    to files downloaded from the internet or copied from network shares.

.PARAMETER Path
    Root path to scan. Defaults to the current directory.

.PARAMETER Recurse
    Scan subdirectories recursively.

.PARAMETER StreamName
    Remove only streams with this name (e.g. "Zone.Identifier").
    When omitted, ALL alternate streams are removed.

.PARAMETER WhatIf
    Show what would be done without making any changes.

.EXAMPLE
    .\remove-ads.ps1 -Path . -Recurse
    Removes all ADS from every file under the current directory.

.EXAMPLE
    .\remove-ads.ps1 -Path C:\Downloads -Recurse -StreamName Zone.Identifier
    Removes only the Zone.Identifier stream (unblocks downloaded files).

.EXAMPLE
    .\remove-ads.ps1 -Path . -Recurse -WhatIf
    Preview what would be removed without making changes.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Position = 0)]
    [string] $Path = '.',

    [switch] $Recurse,

    [string] $StreamName,

    [switch] $WhatIf
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$resolvedPath = Resolve-Path -LiteralPath $Path

$getChildItemArgs = @{
    LiteralPath = $resolvedPath
    File        = $true
    Recurse     = $Recurse.IsPresent
    Force       = $true   # include hidden/system files
    ErrorAction = 'SilentlyContinue'
}

$removed  = 0
$skipped  = 0
$errors   = 0

Get-ChildItem @getChildItemArgs | ForEach-Object {
    $file = $_

    try {
        $getStreamArgs = @{ LiteralPath = $file.FullName }
        if ($StreamName) {
            $getStreamArgs['Stream'] = $StreamName
        }

        $streams = Get-Item @getStreamArgs -Stream * -ErrorAction SilentlyContinue |
                   Where-Object { $_.Stream -ne ':$DATA' }   # skip the primary stream

        if ($StreamName) {
            $streams = $streams | Where-Object { $_.Stream -eq $StreamName }
        }

        foreach ($stream in $streams) {
            $displayName = "$($file.FullName):$($stream.Stream)"

            if ($PSCmdlet.ShouldProcess($displayName, 'Remove-Item')) {
                Remove-Item -LiteralPath $file.FullName -Stream $stream.Stream -ErrorAction Stop
                Write-Verbose "Removed: $displayName"
                $removed++
            } else {
                $skipped++
            }
        }
    } catch {
        Write-Warning "Error processing '$($file.FullName)': $_"
        $errors++
    }
}

$action = if ($WhatIf) { 'Would remove' } else { 'Removed' }
Write-Host "$action $removed stream(s). Skipped: $skipped. Errors: $errors."
