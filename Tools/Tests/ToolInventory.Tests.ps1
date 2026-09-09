Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$modulePath = Join-Path $repoRoot 'Tools\Modules\Tooling\ToolInventory.psm1'
$commandPath = Join-Path $repoRoot 'Tools\Get-ToolInventory.ps1'

function Assert-RSToolInventoryEqual {
    param(
        [AllowNull()][object]$Actual,
        [AllowNull()][object]$Expected,
        [Parameter(Mandatory = $true)][string]$Message
    )

    if ($Actual -ne $Expected) {
        throw "$Message Expected '$Expected' but got '$Actual'."
    }
}

function New-RSToolInventoryFixtureManifest {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [object[]]$Entries = @(),
        [object[]]$Rules = @()
    )

    [ordered]@{
        schemaVersion = 3
        entries = @($Entries)
        rules = @($Rules)
    } | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $Path -Encoding utf8
}

function New-RSToolInventoryFixtureMetadata {
    param(
        [string]$Path,
        [string]$Id,
        [string]$PathPattern
    )

    $metadata = [ordered]@{
        kind = 'command'
        visibility = 'internal'
        lifecycle = 'active'
        purpose = 'Exercise a focused tool inventory fixture.'
        usage = 'Used only by ToolInventory.Tests.ps1.'
        usedBy = @('ToolInventory.Tests.ps1')
        manualEntryPoints = @()
        prerequisites = @('Repository working tree')
        sideEffects = @()
        outputs = @('Fixture result')
        ownerSpec = 'Specs/Testing/Testing_ToolingGovernance.md'
        tests = @('Tools/Tests/ToolInventory.Tests.ps1')
        replacement = $null
        supportPolicy = $null
        namingPolicy = $null
    }
    if (-not [string]::IsNullOrWhiteSpace($Path)) {
        $metadata.path = $Path
    }
    if (-not [string]::IsNullOrWhiteSpace($Id)) {
        $metadata.id = $Id
    }
    if (-not [string]::IsNullOrWhiteSpace($PathPattern)) {
        $metadata.pathPattern = $PathPattern
    }
    return [pscustomobject]$metadata
}

Describe 'Tools inventory contract' {
    BeforeAll {
        Import-Module $modulePath -Force
    }

    It 'classifies every tracked or non-ignored file in the current Tools tree' {
        $inventory = Get-RSToolInventory -RepositoryRoot $repoRoot

        Assert-RSToolInventoryEqual -Actual $inventory.Summary.BlockingFindings -Expected 0 -Message 'The checked-in tool inventory must be valid.'
        Assert-RSToolInventoryEqual -Actual $inventory.Summary.ClassifiedFiles -Expected $inventory.Summary.DiscoveredFiles -Message 'Every discovered Tools file must be classified.'
        Assert-RSToolInventoryEqual -Actual (@($inventory.Files | Where-Object { $_.Path -match '^Tools/[^/]+$' -and $_.ClassificationSource -ne 'entry' }).Count) -Expected 0 -Message 'Every direct child of Tools must use an explicit entry.'
    }

    It 'rejects an unclassified top-level file even when a broad rule would match it' {
        $manifestPath = Join-Path $TestDrive 'top-level-rule.json'
        $rule = New-RSToolInventoryFixtureMetadata -Id 'all-powershell' -PathPattern '^Tools/.+\.ps1$'
        New-RSToolInventoryFixtureManifest -Path $manifestPath -Rules @($rule)

        $inventory = Get-RSToolInventory `
            -RepositoryRoot $repoRoot `
            -ManifestPath $manifestPath `
            -FilePaths @('Tools/New-UndocumentedCommand.ps1')

        Assert-RSToolInventoryEqual -Actual (@($inventory.Findings | Where-Object Code -eq 'TOP_LEVEL_REQUIRES_EXPLICIT_ENTRY').Count) -Expected 1 -Message 'A directory rule must not admit a new top-level tool.'
    }

    It 'rejects an unclassified nested file outside reviewed directory rules' {
        $manifestPath = Join-Path $TestDrive 'nested-unclassified.json'
        New-RSToolInventoryFixtureManifest -Path $manifestPath

        $inventory = Get-RSToolInventory `
            -RepositoryRoot $repoRoot `
            -ManifestPath $manifestPath `
            -FilePaths @('Tools/Modules/Unknown.psm1')

        Assert-RSToolInventoryEqual -Actual (@($inventory.Findings | Where-Object Code -eq 'UNCLASSIFIED_PATH').Count) -Expected 1 -Message 'A nested tool must have an entry or reviewed rule.'
    }

    It 'fails when a derived module consumer is missing from usedBy' {
        $fixtureRoot = Join-Path $TestDrive 'derived-consumers'
        $toolsRoot = Join-Path $fixtureRoot 'Tools'
        [void](New-Item -ItemType Directory -Path (Join-Path $toolsRoot 'Modules') -Force)
        $moduleRelative = 'Tools/Modules/Dependency.psm1'
        $consumerRelative = 'Tools/Consumer.ps1'
        Set-Content -LiteralPath (Join-Path $fixtureRoot $moduleRelative) -Value 'function Get-Fixture { $true }' -Encoding utf8
        Set-Content -LiteralPath (Join-Path $fixtureRoot $consumerRelative) `
            -Value "Import-Module (Join-Path `$PSScriptRoot 'Modules\Dependency.psm1')" -Encoding utf8
        & git -C $fixtureRoot init --quiet

        $moduleEntry = New-RSToolInventoryFixtureMetadata -Path $moduleRelative
        $moduleEntry.kind = 'module'
        $moduleEntry | Add-Member -NotePropertyName consumerClosure -NotePropertyValue 'derived-references'
        $moduleEntry.usedBy = @('Someone else')
        $consumerEntry = New-RSToolInventoryFixtureMetadata -Path $consumerRelative
        $manifestPath = Join-Path $TestDrive 'derived-consumers.json'
        New-RSToolInventoryFixtureManifest -Path $manifestPath -Entries @($moduleEntry, $consumerEntry)

        $inventory = Get-RSToolInventory -RepositoryRoot $fixtureRoot -ManifestPath $manifestPath `
            -FilePaths @($moduleRelative, $consumerRelative)
        Assert-RSToolInventoryEqual -Actual (@($inventory.Findings | Where-Object Code -eq 'MISSING_DERIVED_CONSUMER').Count) `
            -Expected 1 -Message 'An undeclared real import edge must fail governance.'

        $moduleEntry.usedBy = @($consumerRelative)
        New-RSToolInventoryFixtureManifest -Path $manifestPath -Entries @($moduleEntry, $consumerEntry)
        $reconciled = Get-RSToolInventory -RepositoryRoot $fixtureRoot -ManifestPath $manifestPath `
            -FilePaths @($moduleRelative, $consumerRelative)
        Assert-RSToolInventoryEqual -Actual $reconciled.Summary.BlockingFindings -Expected 0 `
            -Message 'A reconciled derived consumer edge must pass governance.'
    }

    It 'rejects stale explicit paths' {
        $manifestPath = Join-Path $TestDrive 'stale-entry.json'
        $entry = New-RSToolInventoryFixtureMetadata -Path 'Tools/Missing.ps1'
        New-RSToolInventoryFixtureManifest -Path $manifestPath -Entries @($entry)

        $inventory = Get-RSToolInventory `
            -RepositoryRoot $repoRoot `
            -ManifestPath $manifestPath `
            -FilePaths @()

        Assert-RSToolInventoryEqual -Actual (@($inventory.Findings | Where-Object Code -eq 'STALE_ENTRY').Count) -Expected 1 -Message 'A removed or renamed file must not leave stale inventory metadata.'
    }

    It 'ignores tracked files deleted from the working tree during an in-progress move' {
        $fixtureRoot = Join-Path $TestDrive 'deleted-working-tree-file'
        $toolsRoot = Join-Path $fixtureRoot 'Tools'
        New-Item -ItemType Directory -Path $toolsRoot -Force | Out-Null

        $keptEntry = New-RSToolInventoryFixtureMetadata -Path 'Tools/Kept.ps1'
        $manifestEntry = New-RSToolInventoryFixtureMetadata -Path 'Tools/tool-inventory.json'
        $manifestPath = Join-Path $toolsRoot 'tool-inventory.json'
        New-RSToolInventoryFixtureManifest -Path $manifestPath -Entries @($keptEntry, $manifestEntry)
        Set-Content -LiteralPath (Join-Path $toolsRoot 'Kept.ps1') -Value '# kept' -Encoding utf8
        $deletedPath = Join-Path $toolsRoot 'Deleted.ps1'
        Set-Content -LiteralPath $deletedPath -Value '# deleted after indexing' -Encoding utf8

        & git -C $fixtureRoot init --quiet
        if ($LASTEXITCODE -ne 0) {
            throw 'Failed to initialize the deleted-working-tree-file Git fixture.'
        }
        & git -C $fixtureRoot add -- Tools
        if ($LASTEXITCODE -ne 0) {
            throw 'Failed to index the deleted-working-tree-file Git fixture.'
        }
        Remove-Item -LiteralPath $deletedPath -Force

        $inventory = Get-RSToolInventory -RepositoryRoot $fixtureRoot

        Assert-RSToolInventoryEqual -Actual $inventory.Summary.DiscoveredFiles -Expected 2 -Message 'Deleted tracked files must not remain in the working-tree inventory.'
        Assert-RSToolInventoryEqual -Actual $inventory.Summary.BlockingFindings -Expected 0 -Message 'An unstaged deletion from a path move must not create a phantom unclassified file.'
    }

    It 'requires an explicit naming policy for reviewed public non-Verb-Noun entrypoints' {
        $manifestPath = Join-Path $TestDrive 'legacy-name.json'
        $entry = New-RSToolInventoryFixtureMetadata -Path 'Tools/AnalyzeThing.ps1'
        $entry.visibility = 'public'
        New-RSToolInventoryFixtureManifest -Path $manifestPath -Entries @($entry)

        $inventory = Get-RSToolInventory `
            -RepositoryRoot $repoRoot `
            -ManifestPath $manifestPath `
            -FilePaths @('Tools/AnalyzeThing.ps1')

        Assert-RSToolInventoryEqual -Actual (@($inventory.Findings | Where-Object Code -eq 'INVALID_METADATA').Count) -Expected 1 -Message 'A reviewed naming exception must be explicit.'

        $entry.namingPolicy = 'Reviewed legacy path retained for an existing consumer.'
        New-RSToolInventoryFixtureManifest -Path $manifestPath -Entries @($entry)
        $reviewedInventory = Get-RSToolInventory `
            -RepositoryRoot $repoRoot `
            -ManifestPath $manifestPath `
            -FilePaths @('Tools/AnalyzeThing.ps1')

        Assert-RSToolInventoryEqual -Actual $reviewedInventory.Summary.BlockingFindings -Expected 0 -Message 'A noncanonical public path with an explicit reviewed policy should be classified.'
        Assert-RSToolInventoryEqual -Actual $reviewedInventory.Files[0].NamingPolicy -Expected $entry.namingPolicy -Message 'Resolved inventory records must retain naming-policy metadata.'
    }

    It 'requires explicit prerequisites metadata on entries and directory rules' {
        $manifestPath = Join-Path $TestDrive 'missing-prerequisites.json'
        $entry = New-RSToolInventoryFixtureMetadata -Path 'Tools/Fixture.ps1'
        $entry.PSObject.Properties.Remove('prerequisites')
        New-RSToolInventoryFixtureManifest -Path $manifestPath -Entries @($entry)

        $inventory = Get-RSToolInventory `
            -RepositoryRoot $repoRoot `
            -ManifestPath $manifestPath `
            -FilePaths @('Tools/Fixture.ps1')

        Assert-RSToolInventoryEqual -Actual (@($inventory.Findings | Where-Object {
                    $_.Code -eq 'INVALID_METADATA' -and $_.Message -match "prerequisites"
                }).Count) -Expected 1 -Message 'Every effective record must declare its prerequisites array, including an explicit empty array.'
    }

    It 'requires retention policy for compatibility records' {
        $manifestPath = Join-Path $TestDrive 'missing-support-policy.json'
        $entry = New-RSToolInventoryFixtureMetadata -Path 'Tools/Legacy.ps1'
        $entry.lifecycle = 'compatibility'
        $entry.replacement = 'Tools/Legacy.ps1'
        New-RSToolInventoryFixtureManifest -Path $manifestPath -Entries @($entry)

        $inventory = Get-RSToolInventory `
            -RepositoryRoot $repoRoot `
            -ManifestPath $manifestPath `
            -FilePaths @('Tools/Legacy.ps1')

        Assert-RSToolInventoryEqual -Actual (@($inventory.Findings | Where-Object {
                    $_.Code -eq 'INVALID_METADATA' -and $_.Message -match "supportPolicy"
                }).Count) -Expected 1 -Message 'Compatibility records must state their retention or removal policy.'

        $entry.lifecycle = 'compatibility'
        $entry.supportPolicy = 'Retain until an explicit reviewed lifecycle decision authorizes removal.'
        New-RSToolInventoryFixtureManifest -Path $manifestPath -Entries @($entry)
        $reviewedInventory = Get-RSToolInventory `
            -RepositoryRoot $repoRoot `
            -ManifestPath $manifestPath `
            -FilePaths @('Tools/Legacy.ps1')

        Assert-RSToolInventoryEqual -Actual $reviewedInventory.Summary.BlockingFindings -Expected 0 -Message 'A compatibility lifecycle record is valid when supportPolicy is present.'
        Assert-RSToolInventoryEqual -Actual $reviewedInventory.Files[0].SupportPolicy -Expected $entry.supportPolicy -Message 'Resolved inventory records must retain support policy metadata.'
    }

    It 'requires every non-null replacement to be an exact existing Tools path' {
        $manifestPath = Join-Path $TestDrive 'replacement-path.json'
        $entry = New-RSToolInventoryFixtureMetadata -Path 'Tools/Legacy.ps1'
        $entry.replacement = 'Tools/Current.ps1 for current use'
        $current = New-RSToolInventoryFixtureMetadata -Path 'Tools/Current.ps1'
        New-RSToolInventoryFixtureManifest -Path $manifestPath -Entries @($entry, $current)

        $inventory = Get-RSToolInventory `
            -RepositoryRoot $repoRoot `
            -ManifestPath $manifestPath `
            -FilePaths @('Tools/Legacy.ps1', 'Tools/Current.ps1')

        Assert-RSToolInventoryEqual -Actual (@($inventory.Findings | Where-Object Code -eq 'INVALID_REPLACEMENT').Count) -Expected 1 -Message 'Replacement prose must not be accepted as a canonical path.'

        $entry.replacement = 'Tools/Current.ps1'
        New-RSToolInventoryFixtureManifest -Path $manifestPath -Entries @($entry, $current)
        $reviewedInventory = Get-RSToolInventory `
            -RepositoryRoot $repoRoot `
            -ManifestPath $manifestPath `
            -FilePaths @('Tools/Legacy.ps1', 'Tools/Current.ps1')

        Assert-RSToolInventoryEqual -Actual $reviewedInventory.Summary.BlockingFindings -Expected 0 -Message 'An existing exact repository-relative Tools path is a valid replacement.'
    }

    It 'rejects historical entries without a reviewed manual entry point' {
        $manifestPath = Join-Path $TestDrive 'orphaned-historical.json'
        $entry = New-RSToolInventoryFixtureMetadata -Path 'Tools/HistoricalArtifact.psm1'
        $entry.kind = 'module'
        $entry.lifecycle = 'historical'
        New-RSToolInventoryFixtureManifest -Path $manifestPath -Entries @($entry)

        $inventory = Get-RSToolInventory `
            -RepositoryRoot $repoRoot `
            -ManifestPath $manifestPath `
            -FilePaths @('Tools/HistoricalArtifact.psm1')

        Assert-RSToolInventoryEqual -Actual (@($inventory.Findings | Where-Object Code -eq 'ORPHANED_HISTORICAL_ENTRY').Count) -Expected 1 -Message 'An orphaned historical entry must fail inventory validation.'
    }

    It 'rejects unreviewed and missing historical manual entry-point paths' {
        $manifestPath = Join-Path $TestDrive 'invalid-historical-root.json'
        $entry = New-RSToolInventoryFixtureMetadata -Path 'Tools/HistoricalArtifact.psm1'
        $entry.kind = 'module'
        $entry.lifecycle = 'historical'
        $entry.manualEntryPoints = @('Tools/Unreviewed-HistoricalRoot.ps1')
        New-RSToolInventoryFixtureManifest -Path $manifestPath -Entries @($entry)

        $invalidInventory = Get-RSToolInventory `
            -RepositoryRoot $repoRoot `
            -ManifestPath $manifestPath `
            -FilePaths @('Tools/HistoricalArtifact.psm1')

        Assert-RSToolInventoryEqual -Actual (@($invalidInventory.Findings | Where-Object Code -eq 'INVALID_MANUAL_ENTRY_POINT').Count) -Expected 1 -Message 'Only the three reviewed historical roots may retain a historical entry.'

        $entry.manualEntryPoints = @('Tools/Evaluate-TerminalEngines.ps1')
        New-RSToolInventoryFixtureManifest -Path $manifestPath -Entries @($entry)
        $missingInventory = Get-RSToolInventory `
            -RepositoryRoot $repoRoot `
            -ManifestPath $manifestPath `
            -FilePaths @('Tools/HistoricalArtifact.psm1')

        Assert-RSToolInventoryEqual -Actual (@($missingInventory.Findings | Where-Object Code -eq 'MISSING_MANUAL_ENTRY_POINT').Count) -Expected 1 -Message 'A reviewed root name must still resolve to an existing Tools file.'
    }

    It 'rejects a reviewed manual root whose inventory entry is not a historical command' {
        $manifestPath = Join-Path $TestDrive 'nonhistorical-root.json'
        $entry = New-RSToolInventoryFixtureMetadata -Path 'Tools/HistoricalArtifact.psm1'
        $entry.kind = 'module'
        $entry.lifecycle = 'historical'
        $entry.manualEntryPoints = @('Tools/Evaluate-TerminalEngines.ps1')
        $root = New-RSToolInventoryFixtureMetadata -Path 'Tools/Evaluate-TerminalEngines.ps1'
        New-RSToolInventoryFixtureManifest -Path $manifestPath -Entries @($entry, $root)

        $inventory = Get-RSToolInventory `
            -RepositoryRoot $repoRoot `
            -ManifestPath $manifestPath `
            -FilePaths @('Tools/HistoricalArtifact.psm1', 'Tools/Evaluate-TerminalEngines.ps1')

        Assert-RSToolInventoryEqual -Actual (@($inventory.Findings | Where-Object Code -eq 'NONHISTORICAL_MANUAL_ENTRY_POINT').Count) -Expected 1 -Message 'A current command cannot justify retaining historical files.'
    }

    It 'allows reviewed manual roots to retain themselves and exposes the resolved closure' {
        $manifestPath = Join-Path $TestDrive 'historical-self-root.json'
        $root = New-RSToolInventoryFixtureMetadata -Path 'Tools/Evaluate-TerminalEngines.ps1'
        $root.lifecycle = 'historical'
        $root.manualEntryPoints = @('Tools/Evaluate-TerminalEngines.ps1')
        New-RSToolInventoryFixtureManifest -Path $manifestPath -Entries @($root)

        $inventory = Get-RSToolInventory `
            -RepositoryRoot $repoRoot `
            -ManifestPath $manifestPath `
            -FilePaths @('Tools/Evaluate-TerminalEngines.ps1')

        Assert-RSToolInventoryEqual -Actual $inventory.Summary.BlockingFindings -Expected 0 -Message 'A reviewed historical command may name itself as its manual entry point.'
        Assert-RSToolInventoryEqual -Actual $inventory.Files[0].ManualEntryPoints[0] -Expected 'Tools/Evaluate-TerminalEngines.ps1' -Message 'Resolved inventory records must expose manual entry-point metadata.'
    }

    It 'uses a reviewed nested rule when exactly one rule matches' {
        $manifestPath = Join-Path $TestDrive 'nested-rule.json'
        $rule = New-RSToolInventoryFixtureMetadata -Id 'pester' -PathPattern '^Tools/Tests/[^/]+\.Tests\.ps1$'
        $rule.kind = 'test'
        New-RSToolInventoryFixtureManifest -Path $manifestPath -Rules @($rule)

        $inventory = Get-RSToolInventory `
            -RepositoryRoot $repoRoot `
            -ManifestPath $manifestPath `
            -FilePaths @('Tools/Tests/Example.Tests.ps1')

        Assert-RSToolInventoryEqual -Actual $inventory.Summary.BlockingFindings -Expected 0 -Message 'A conventional nested Pester test should match its reviewed rule.'
        Assert-RSToolInventoryEqual -Actual $inventory.Files[0].ClassificationSource -Expected 'rule:pester' -Message 'The resolved file should retain its rule provenance.'
    }

    It 'renders a useful Markdown file inventory' {
        $inventory = Get-RSToolInventory -RepositoryRoot $repoRoot
        $markdown = ConvertTo-RSToolInventoryMarkdown -Inventory $inventory

        if ($markdown -notmatch '(?m)^## Files\r?$' -or
            $markdown -notmatch 'Tools/Get-ToolInventory\.ps1' -or
            $markdown -notmatch '(?m)^- Blocking findings: 0\r?$') {
            throw 'Markdown output must include summary, file table, and the public inventory command.'
        }
    }

    It 'documents the command contract and supports Validate as an alias' {
        $source = Get-Content -LiteralPath $commandPath -Raw
        if ($source -notmatch '(?ms)^<#.*?\.NOTES\s+.+?#>') {
            throw 'Get-ToolInventory.ps1 must contain a non-empty .NOTES help section.'
        }
        if ($source -notmatch "\[Alias\('Validate'\)\]\s*\r?\n\s*\[switch\]\`$FailOnFindings") {
            throw 'Get-ToolInventory.ps1 must expose Validate as an alias for FailOnFindings.'
        }

        $inventory = & $commandPath -Validate -Format Object
        Assert-RSToolInventoryEqual -Actual $inventory.Summary.BlockingFindings -Expected 0 -Message 'The Validate alias must execute strict inventory validation.'
    }
}
