<#
.SYNOPSIS
    Validates and displays the checked-in RedSalamander Tools inventory.

.DESCRIPTION
    Reconciles Tools/tool-inventory.json with tracked and non-ignored files beneath
    Tools. Every direct child of Tools requires an explicit manifest entry; reviewed
    directory rules may classify conventional files in deeper subdirectories.

    The command can emit the complete inventory as PowerShell objects, JSON, or a
    Markdown report. Use -FailOnFindings in CI and review workflows to reject stale,
    unclassified, ambiguous, or malformed inventory metadata.

.PARAMETER RepoRoot
    Repository root containing the Tools directory. Defaults to the parent of this
    script's directory.

.PARAMETER ManifestPath
    Repository-relative or absolute path to the inventory manifest.

.PARAMETER Format
    Output representation: Markdown, Json, or Object. Markdown is the default.

.PARAMETER FailOnFindings
    Throws when inventory validation reports any blocking finding.

.EXAMPLE
    .\Tools\Get-ToolInventory.ps1 -FailOnFindings

    Validates the working tree and prints a Markdown inventory.

.EXAMPLE
    .\Tools\Get-ToolInventory.ps1 -Format Object

    Returns the structured inventory for further PowerShell filtering.

.OUTPUTS
    System.String for Markdown and JSON formats; PSCustomObject for Object format.

.NOTES
    Tooling inventory contract: Specs/Testing/Testing_ToolingGovernance.md.
    Update Tools/tool-inventory.json in the same change as any tool lifecycle,
    ownership, consumer, side-effect, output, test, replacement, or path change.
#>

[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent $PSScriptRoot),

    [string]$ManifestPath = 'Tools/tool-inventory.json',

    [ValidateSet('Markdown', 'Json', 'Object')]
    [string]$Format = 'Markdown',

    [Alias('Validate')]
    [switch]$FailOnFindings
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$modulePath = Join-Path $PSScriptRoot 'Modules\Tooling\ToolInventory.psm1'
Import-Module $modulePath -Force

$inventory = Get-RSToolInventory -RepositoryRoot $RepoRoot -ManifestPath $ManifestPath
if ($FailOnFindings -and $inventory.Summary.BlockingFindings -ne 0) {
    $summary = @($inventory.Findings | Select-Object -First 10 | ForEach-Object {
            "$($_.Code): $($_.Path)"
        }) -join '; '
    throw "Tools inventory validation failed with $($inventory.Summary.BlockingFindings) finding(s): $summary"
}

switch ($Format) {
    'Object' {
        $inventory
    }
    'Json' {
        $inventory | ConvertTo-Json -Depth 12
    }
    default {
        ConvertTo-RSToolInventoryMarkdown -Inventory $inventory
    }
}
