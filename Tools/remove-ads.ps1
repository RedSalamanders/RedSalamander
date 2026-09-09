<#
.SYNOPSIS
    Removes NTFS alternate data streams (ADS) from files.

.DESCRIPTION
    Canonical repository command for removing named alternate data streams.
    It reports matched, previewed, removed, skipped, and failed streams
    separately and exits nonzero after any resolution, enumeration, or removal
    failure.

.PARAMETER Path
    File or directory to scan. Defaults to the current directory.

.PARAMETER Recurse
    Scan subdirectories recursively when Path names a directory.

.PARAMETER StreamName
    Remove only the named stream, such as Zone.Identifier. When omitted, remove
    every named stream while preserving the primary :$DATA stream.

.OUTPUTS
    One summary string reporting matched, previewed, removed, skipped, and failed stream counts.

.NOTES
    Prerequisites: an NTFS path and permission to enumerate/remove its named streams. Side effects: removes matching alternate data streams only after ShouldProcess approval; -WhatIf is read-only. Exit code is 1 after any resolution, enumeration, or removal failure and 0 otherwise. This is the sole supported ADS-removal command.

.EXAMPLE
    .\Tools\remove-ads.ps1 -Path . -Recurse
    Removes every named stream below the current directory.

.EXAMPLE
    .\Tools\remove-ads.ps1 -Path C:\Downloads -Recurse -StreamName Zone.Identifier
    Removes only Zone.Identifier streams.

.EXAMPLE
    .\Tools\remove-ads.ps1 -Path . -Recurse -WhatIf
    Reports matching streams without changing them.
#>
#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [Parameter(Position = 0)]
    [string] $Path = '.',

    [switch] $Recurse,

    [string] $StreamName
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$matched = 0
$wouldRemove = 0
$removed = 0
$skipped = 0
$failed = 0

function Write-RSAdsSummary {
    Write-Output "Matched: $matched. Would remove: $wouldRemove. Removed: $removed. Skipped: $skipped. Failed: $failed."
}

try {
    $resolvedPath = Resolve-Path -LiteralPath $Path -ErrorAction Stop
    $rootItem = Get-Item -LiteralPath $resolvedPath.Path -Force -ErrorAction Stop
} catch {
    Write-Warning "Cannot resolve ADS scan path '$Path': $($_.Exception.Message)"
    $failed++
    Write-RSAdsSummary
    exit 1
}

try {
    $files = if ($rootItem.PSIsContainer) {
        @(Get-ChildItem `
            -LiteralPath $rootItem.FullName `
            -File `
            -Recurse:$Recurse.IsPresent `
            -Force `
            -ErrorAction Stop)
    } else {
        @($rootItem)
    }
} catch {
    Write-Warning "Cannot enumerate ADS scan path '$($rootItem.FullName)': $($_.Exception.Message)"
    $failed++
    Write-RSAdsSummary
    exit 1
}

foreach ($file in $files) {
    try {
        $streams = @(
            Get-Item -LiteralPath $file.FullName -Stream * -Force -ErrorAction Stop |
                Where-Object {
                    $_.Stream -ne ':$DATA' -and
                    ([string]::IsNullOrEmpty($StreamName) -or
                     [string]::Equals($_.Stream, $StreamName, [StringComparison]::OrdinalIgnoreCase))
                }
        )
    } catch {
        Write-Warning "Cannot enumerate streams for '$($file.FullName)': $($_.Exception.Message)"
        $failed++
        continue
    }

    foreach ($stream in $streams) {
        $matched++
        $displayName = "$($file.FullName):$($stream.Stream)"
        if (-not $PSCmdlet.ShouldProcess($displayName, 'Remove alternate data stream')) {
            if ([bool]$WhatIfPreference) {
                $wouldRemove++
            } else {
                $skipped++
            }
            continue
        }

        try {
            Remove-Item -LiteralPath $file.FullName -Stream $stream.Stream -ErrorAction Stop
            Write-Verbose "Removed: $displayName"
            $removed++
        } catch {
            Write-Warning "Cannot remove '$displayName': $($_.Exception.Message)"
            $failed++
        }
    }
}

Write-RSAdsSummary
if ($failed -ne 0) {
    exit 1
}
exit 0
