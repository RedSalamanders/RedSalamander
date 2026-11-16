Import-Module (Join-Path $PSScriptRoot 'TestSupport.psm1') -Force -ErrorAction Stop
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
        Invoke-RSTestFixtureGit -Arguments @('-C'; $fixtureRoot; 'init'; '--quiet')

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

        Invoke-RSTestFixtureGit -Arguments @('-C'; $fixtureRoot; 'init'; '--quiet')
        if ($LASTEXITCODE -ne 0) {
            throw 'Failed to initialize the deleted-working-tree-file Git fixture.'
        }
        Invoke-RSTestFixtureGit -Arguments @('-C'; $fixtureRoot; 'add'; '--'; 'Tools')
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

Describe 'Engine-free tool inventory metadata transport' {
    It 'keeps bootstrap independent and separates successful paths from native diagnostics' {
        $fixture=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('inventory-native-'+[guid]::NewGuid().ToString('N'))
        $copy=Join-Path $fixture 'ToolBootstrap.psm1'
        Copy-Item -LiteralPath (Join-Path $repoRoot 'Tools/Modules/Tooling/ToolInventory.psm1') -Destination $copy
        $module=Import-Module $copy -PassThru -Scope Local -Prefix ToolBootstrap
        $savedConsoleEncoding=[Console]::OutputEncoding
        try {
            [Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
            (Test-Path (Join-Path $fixture '.ci-engine'))|Should Be $false
            $value='first é雪'+[char]13+[char]10+'second'
            $result=& $module {param($root,$value)Invoke-RSToolInventoryGitCapture -Root $root -Arguments @('-c',('fixture.payload='+$value),'config','--get','fixture.payload')} $fixture $value
            $result.ExitCode|Should Be 0
            ($result.Output -join '|')|Should Be 'first é雪|second'
            $result.ErrorOutput.Count|Should Be 0
            [Console]::OutputEncoding=[Text.Encoding]::GetEncoding(437)
            $legacy=& $module {param($root)Invoke-RSToolInventoryGitCapture -Root $root -Arguments @('-c','fixture.payload=é雪','config','--get','fixture.payload')} $fixture
            $legacy.Output[0]|Should Be 'é雪'
            [Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
            $null=& $module {param($root)Invoke-RSToolInventoryGitCapture -Root $root -Arguments @('init','--quiet')} $fixture
            $null=& $module {param($root)Invoke-RSToolInventoryGitCapture -Root $root -Arguments @('config','core.quotepath','false')} $fixture
            [void][IO.Directory]::CreateDirectory((Join-Path $fixture 'Tools'))
            $path='Tools/café-Δ-雪.ps1'
            [IO.File]::WriteAllText((Join-Path $fixture $path),'# owned fixture',[Text.UTF8Encoding]::new($false))
            foreach($encoding in @([Text.Encoding]::GetEncoding(437),[Text.UTF8Encoding]::new($false))) {
                [Console]::OutputEncoding=$encoding
                $paths=@(& $module {param($root)Get-RSToolRepositoryFilePaths -RepositoryRoot $root} $fixture)
                ($paths -ccontains $path)|Should Be $true
                [Console]::OutputEncoding.CodePage|Should Be $encoding.CodePage
            }
            $result=& $module {param($root)Invoke-RSToolInventoryGitCapture -Root $root -Arguments @('rev-parse','--verify','missing-object')} $fixture
            ($result.ExitCode -ne 0)|Should Be $true
            $result.Output.Count|Should Be 0
            ($result.ErrorOutput.Count -gt 0)|Should Be $true
            $failed=$false
            try {& $module {param($root)Get-RSToolRepositoryFilePaths -RepositoryRoot $root} (Join-Path $fixture 'missing')}
            catch {$failed=$_.Exception.Message.StartsWith('Failed to enumerate tracked and non-ignored Tools files from ')}
            $failed|Should Be $true
        } finally {[Console]::OutputEncoding=$savedConsoleEncoding;Remove-Module $module}
    }
}

Describe 'Guarded scheduled Git fixtures' {
    BeforeEach {
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Build/ProcessStreaming.psm1') -ErrorAction Stop
        $starterPath=Join-Path $repoRoot 'Tools/Modules/Build/SanitizedEnvironment.psm1'
        $module=@(Get-Module -All|Where-Object Path -ceq $starterPath)
        if($module.Count -ne 1){throw 'Expected the canonical fixture starter.'}
        $script:fixtureGitAudit=[pscustomobject]@{Starter=$module[0].ExportedFunctions['Start-RSContainedProcess'];Rows=[Collections.Generic.List[object]]::new()}
        $audit=$script:fixtureGitAudit
        Mock Start-RSContainedProcess -ModuleName ProcessStreaming ({
            param($ProcessStartInfo,[switch]$DelegateArtifactOperation)
            if($ProcessStartInfo.UseShellExecute -or -not $ProcessStartInfo.CreateNoWindow -or -not $ProcessStartInfo.RedirectStandardOutput -or
                -not $ProcessStartInfo.RedirectStandardError -or $DelegateArtifactOperation -or $ProcessStartInfo.StandardOutputEncoding.CodePage -ne 65001){throw 'Refusing unsafe scheduled fixture child before launch.'}
            $owned=& $audit.Starter -ProcessStartInfo $ProcessStartInfo
            $audit.Rows.Add(@{pid=$owned.Id;contained=$owned.IsInContainmentJob;argv=@($ProcessStartInfo.ArgumentList);cwd=$ProcessStartInfo.WorkingDirectory;noWindow=$ProcessStartInfo.CreateNoWindow})
            return $owned
        }.GetNewClosure())
        $script:fixtureGitRoot=Join-Path $TestDrive ([guid]::NewGuid().ToString('N'));[void][IO.Directory]::CreateDirectory($script:fixtureGitRoot)
    }
    AfterEach {
        foreach($row in $script:fixtureGitAudit.Rows){$row.contained|Should Be $true;(Get-Process -Id $row.pid -ErrorAction SilentlyContinue)|Should BeNullOrEmpty}
        $script:fixtureGitAudit.Rows.Count|Should BeGreaterThan 0
    }
    It 'preserves literal argv, caller CWD, empty stdout and Unicode under both console encodings' {
        $saved=[Console]::OutputEncoding
        Push-Location $script:fixtureGitRoot
        try {
            Invoke-RSTestFixtureGit -Arguments @('-C'; $script:fixtureGitRoot; 'init'; '--quiet')
            $LASTEXITCODE|Should Be 0
            foreach($encoding in @([Text.Encoding]::GetEncoding(437),[Text.UTF8Encoding]::new($false))){
                [Console]::OutputEncoding=$encoding
                $nameAndValue=@('fixture.literal','literal * é雪')
                $empty=@(Invoke-RSTestFixtureGit -Arguments @('-C'; $script:fixtureGitRoot; 'config'; '--local'; '--'; $nameAndValue))
                $empty.Count|Should Be 0;$LASTEXITCODE|Should Be 0
                $value=Invoke-RSTestFixtureGit -Arguments @('-C'; $script:fixtureGitRoot; 'config'; '--get'; 'fixture.literal')
                $value|Should Be 'literal * é雪';$LASTEXITCODE|Should Be 0
                [Console]::OutputEncoding.CodePage|Should Be $encoding.CodePage
            }
            $multiple=@(Invoke-RSTestFixtureGit -Arguments @('-C'; $script:fixtureGitRoot; 'config'; '--add'; 'fixture.literal'; 'second * é雪'))
            $multiple.Count|Should Be 0
            $lines=@(Invoke-RSTestFixtureGit -Arguments @('-C'; $script:fixtureGitRoot; 'config'; '--get-all'; 'fixture.literal'))
            $LASTEXITCODE|Should Be 0;$lines.Count|Should Be 2
            foreach($line in $lines){($line -is [string])|Should Be $true}
            ($lines -join '|')|Should Be 'literal * é雪|second * é雪'
            Invoke-RSTestFixtureGit -Arguments @('-C'; $script:fixtureGitRoot; 'config'; '--add'; 'fixture.blank'; '')
            $LASTEXITCODE|Should Be 0
            $blank=@(Invoke-RSTestFixtureGit -Arguments @('-C'; $script:fixtureGitRoot; 'config'; '--get-all'; 'fixture.blank'))
            $LASTEXITCODE|Should Be 0;$blank.Count|Should Be 1
            ($blank[0] -is [string])|Should Be $true;$blank[0]|Should Be ''
            $script:fixtureGitAudit.Rows.Count|Should Be 9
            foreach($row in $script:fixtureGitAudit.Rows){$row.cwd|Should Be $script:fixtureGitRoot}
            ($script:fixtureGitAudit.Rows[1].argv -join '|')|Should Be ('-C|'+$script:fixtureGitRoot+'|config|--local|--|fixture.literal|literal * é雪')
        } finally {Pop-Location;[Console]::OutputEncoding=$saved}
    }
    It 'preserves nonzero native exit and stderr redirection without converting errors to stdout' {
        $merged=@(Invoke-RSTestFixtureGit -Arguments @('-C'; $script:fixtureGitRoot; 'rev-parse'; 'HEAD') 2>&1)
        $LASTEXITCODE|Should Be 128
        ($merged -join "`n")|Should Match 'not a git repository'
        $suppressed=@(Invoke-RSTestFixtureGit -Arguments @('-C'; $script:fixtureGitRoot; 'rev-parse'; 'HEAD') 2>$null)
        $LASTEXITCODE|Should Be 128;$suppressed.Count|Should Be 0
        $script:fixtureGitAudit.Rows.Count|Should Be 2
    }
    It 'retains both streams on a partial-output failure and keeps only owned fixture writes' {
        Invoke-RSTestFixtureGit -Arguments @('-C'; $script:fixtureGitRoot; 'init'; '--quiet')
        Invoke-RSTestFixtureGit -Arguments @('-C'; $script:fixtureGitRoot; '-c'; 'user.name=Fixture'; '-c'; 'user.email=fixture@example.invalid'; '-c'; 'commit.gpgsign=false'; 'commit'; '--quiet'; '--allow-empty'; '-m'; 'fixture')
        $LASTEXITCODE|Should Be 0
        $result=@(Invoke-RSTestFixtureGit -Arguments @('-C'; $script:fixtureGitRoot; 'rev-parse'; 'HEAD'; 'HEAD:missing') 2>&1)
        $LASTEXITCODE|Should Be 128
        @($result|Where-Object {[string]$_ -match '^[0-9a-f]{40}$'}).Count|Should Be 1
        ($result -join "`n")|Should Match 'fatal:'
        $script:fixtureGitAudit.Rows.Count|Should Be 3
    }
}
