Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'Tools\Modules\Tooling\SpecInformationArchitecture.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $repoRoot 'Tools\Modules\Testing\TestRunPlan.psm1') -Force -ErrorAction Stop

$testRunContext = New-RSTestRunContext -RepoRoot $repoRoot
$fixtureParent = New-RSTestSandboxScratchDirectory `
    -RepoRoot $repoRoot `
    -Harness 'tools-pester' `
    -Case 'spec-information-architecture'

function Set-RSFixtureFile {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$RelativePath,
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Content
    )

    $path = Join-Path $Root $RelativePath
    New-Item -ItemType Directory -Path (Split-Path -Parent $path) -Force | Out-Null
    Set-Content -LiteralPath $path -Value $Content -Encoding utf8
}

function New-RSSpecFixtureRepository {
    $root = Join-Path $fixtureParent ([Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $root -Force | Out-Null
    & git -C $root init --quiet
    if ($LASTEXITCODE -ne 0) {
        throw 'Failed to initialize the specification information architecture fixture repository.'
    }
    & git -C $root config core.longpaths true
    if ($LASTEXITCODE -ne 0) {
        throw 'Failed to enable long-path support for the specification information architecture fixture repository.'
    }
    & git -C $root config user.email 'spec-ia-tests@example.invalid'
    & git -C $root config user.name 'Spec IA Tests'
    & git -C $root config core.autocrlf false

    Set-RSFixtureFile -Root $root -RelativePath 'src/example.cpp' -Content 'int Example() { return 1; }'
    Set-RSFixtureFile -Root $root -RelativePath 'Tests/example.ps1' -Content "'example'"
    Set-RSFixtureFile -Root $root -RelativePath 'Specs/Core/Core_Example.md' -Content @'
# Example contract

Current behavior is implemented by `src/example.cpp`.
'@
    Set-RSFixtureFile -Root $root -RelativePath 'Specs/NormativeConsistency.json' -Content @'
{
  "schemaVersion": 1,
  "reviewedAtCommit": "fixture",
  "documents": [
    {
      "path": "Specs/Core/Core_Example.md",
      "status": "VERIFIED_MATCH",
      "contractOwner": "Example",
      "implementationAnchors": ["src/example.cpp#Example"],
      "testAnchors": ["Tests/example.ps1#example"],
      "siblingSpecs": [],
      "reviewedSections": ["all"],
      "reviewerNote": "Fixture contract matches its implementation and test.",
      "discrepancyIds": []
    }
  ]
}
'@
    Set-RSFixtureFile -Root $root -RelativePath 'Specs/Plans/WIP/Plan.md' -Content @'
# Fixture plan

- **Status:** ACTIVE
- **Priority:** P2
- **Planned at:** fixture
- **Ownership:** unfinished fixture work only
- **Remaining outcome:** keep the fixture valid
- **Authoritative spec:** `Specs/Core/Core_Example.md`
- **Verification:** run this focused suite
- **Done criteria:** inventory passes
- **STOP:** stop on ambiguity
'@
    Set-RSFixtureFile -Root $root -RelativePath 'Specs/Plans/WIP/README.md' -Content @'
# WIP index

<!-- spec-ia:wip-index:start -->
| Status | Rank | File |
|---|---|---|
| ACTIVE | I0 | [Plan](Plan.md) |
<!-- spec-ia:wip-index:end -->
'@
    Set-RSFixtureFile -Root $root -RelativePath 'Specs/Plans/Done/History.md' -Content "# Historical plan`n"
    & git -C $root add --all
    if ($LASTEXITCODE -ne 0) {
        throw 'Failed to stage the specification information architecture fixture repository.'
    }
    & git -C $root commit --quiet -m 'fixture baseline'
    if ($LASTEXITCODE -ne 0) {
        throw 'Failed to commit the specification information architecture fixture repository.'
    }
    return $root
}

function Get-RSFixtureInventory {
    param([Parameter(Mandatory = $true)][string]$Root)

    return Get-RSSpecInventory `
        -RepositoryRoot $Root `
        -Root Specs `
        -ConsistencyLedgerPath 'Specs/NormativeConsistency.json' `
        -ProtectedDoneBaseCommit 'HEAD'
}

function Get-RSFindingCount {
    param(
        [Parameter(Mandatory = $true)]$Inventory,
        [Parameter(Mandatory = $true)][string]$Code,
        [string]$Severity
    )

    return @($Inventory.Findings | Where-Object {
            $_.Code -eq $Code -and ([string]::IsNullOrWhiteSpace($Severity) -or $_.Severity -eq $Severity)
        }).Count
}

Describe 'Specification information architecture' {
    It 'accepts a valid normative document with an evidence-backed consistency row' {
        $root = New-RSSpecFixtureRepository
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'NORMATIVE_MISSING_CONSISTENCY') | Should Be 0
        (Get-RSFindingCount -Inventory $inventory -Code 'NORMATIVE_MISSING_EVIDENCE') | Should Be 0
    }

    It 'accepts a valid active WIP plan indexed exactly once' {
        $root = New-RSSpecFixtureRepository
        Set-RSFixtureFile -Root $root -RelativePath 'Specs/Plans/WIP/Notes_Scratch.md' -Content "# Human-managed note`n"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'WIP_UNINDEXED') | Should Be 0
        $inventory.Summary.DirectWipFiles | Should Be 1
        @($inventory.WipIndexEntries).Count | Should Be 1
        $inventory.WipIndexEntries[0].Status | Should Be 'ACTIVE'
        ($inventory.Files | Where-Object Path -eq 'Specs/Plans/WIP/Notes_Scratch.md').AuthorityClass | Should Be 'HumanManagedNote'
    }

    It 'accepts a compact retired tombstone with the required retrieval fields' {
        $root = New-RSSpecFixtureRepository
        Set-RSFixtureFile -Root $root -RelativePath 'Specs/Plans/WIP/Retired.md' -Content @'
# Retired fixture

- **Status:** RETIRED
- **Retirement date:** 2026-08-04
- **Reason:** Superseded by the example owner.
- **Authoritative spec:** `Specs/Core/Core_Example.md`
- **Historical retrieval:** `git show HEAD:Specs/Plans/WIP/Retired.md`
'@
        $indexPath = Join-Path $root 'Specs/Plans/WIP/README.md'
        $index = Get-Content -LiteralPath $indexPath -Raw
        $index = $index.Replace('| ACTIVE | I0 | [Plan](Plan.md) |', "| ACTIVE | I0 | [Plan](Plan.md) |`n| RETIRED | R0 | [Retired](Retired.md) |")
        Set-Content -LiteralPath $indexPath -Value $index -Encoding utf8
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'TOMBSTONE_INVALID') | Should Be 0
    }

    It 'rejects a duplicate WIP index entry' {
        $root = New-RSSpecFixtureRepository
        $indexPath = Join-Path $root 'Specs/Plans/WIP/README.md'
        $index = Get-Content -LiteralPath $indexPath -Raw
        $index = $index.Replace('| ACTIVE | I0 | [Plan](Plan.md) |', "| ACTIVE | I0 | [Plan](Plan.md) |`n| HOLD | H0 | [Plan duplicate](Plan.md) |")
        Set-Content -LiteralPath $indexPath -Value $index -Encoding utf8
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'WIP_DUPLICATE_INDEX') | Should Be 1
    }

    It 'rejects an unindexed direct WIP file' {
        $root = New-RSSpecFixtureRepository
        Set-RSFixtureFile -Root $root -RelativePath 'Specs/Plans/WIP/Unindexed.md' -Content "# Unindexed`n"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'WIP_UNINDEXED') | Should Be 1
    }

    It 'rejects a stale WIP index path' {
        $root = New-RSSpecFixtureRepository
        $indexPath = Join-Path $root 'Specs/Plans/WIP/README.md'
        $index = Get-Content -LiteralPath $indexPath -Raw
        $index = $index.Replace('| ACTIVE | I0 | [Plan](Plan.md) |', "| ACTIVE | I0 | [Plan](Plan.md) |`n| ACTIVE | I1 | [Missing](Missing.md) |")
        Set-Content -LiteralPath $indexPath -Value $index -Encoding utf8
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'WIP_STALE_INDEX') | Should Be 1
    }

    It 'rejects an unchecked checkbox in normative prose' {
        $root = New-RSSpecFixtureRepository
        Add-Content -LiteralPath (Join-Path $root 'Specs/Core/Core_Example.md') -Value "`n- [ ] implement later"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'NORMATIVE_PLAN_MARKER') | Should Be 1
    }

    It 'rejects future or implementation-plan headings in normative prose' {
        $root = New-RSSpecFixtureRepository
        Add-Content -LiteralPath (Join-Path $root 'Specs/Core/Core_Example.md') -Value "`n## Future Enhancements`n"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'NORMATIVE_PLAN_MARKER') | Should Be 1
    }

    It 'resolves a local relative link and heading anchor' {
        $root = New-RSSpecFixtureRepository
        Add-Content -LiteralPath (Join-Path $root 'Specs/Core/Core_Example.md') -Value "`n## Exact Anchor`n[Self](Core_Example.md#exact-anchor)"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'LINK_MISSING_TARGET' -Severity 'Error') | Should Be 0
        (Get-RSFindingCount -Inventory $inventory -Code 'LINK_MISSING_ANCHOR' -Severity 'Error') | Should Be 0
    }

    It 'uses GitHub duplicate-heading suffixes' {
        $headings = @(Get-RSMarkdownHeadings -Text "# Same`n## Same`n## Same")

        $headings[0].Slug | Should Be 'same'
        $headings[1].Slug | Should Be 'same-1'
        $headings[2].Slug | Should Be 'same-2'
    }

    It 'rejects missing local targets and heading anchors' {
        $root = New-RSSpecFixtureRepository
        Add-Content -LiteralPath (Join-Path $root 'Specs/Core/Core_Example.md') -Value "`n[Missing](Missing.md)`n[Bad anchor](Core_Example.md#not-present)"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'LINK_MISSING_TARGET' -Severity 'Error') | Should Be 1
        (Get-RSFindingCount -Inventory $inventory -Code 'LINK_MISSING_ANCHOR' -Severity 'Error') | Should Be 1
    }

    It 'ignores fake links inside fenced Cpp PowerShell and Markdown examples' {
        $text = @'
# Example

```cpp
auto sample = "](";
[fake](missing-cpp.md)
```

~~~powershell
$sample = ']('
[fake](missing-pwsh.md)
~~~

````markdown
[fake](missing-markdown.md)
````
'@
        @(Get-RSMarkdownLinks -Text $text).Count | Should Be 0
    }

    It 'ignores a fake link inside inline code' {
        $text = '# Example `call [fake](missing.md)`'

        @(Get-RSMarkdownLinks -Text $text).Count | Should Be 0
    }

    It 'rejects a repository-root escape attempt' {
        $root = New-RSSpecFixtureRepository
        Add-Content -LiteralPath (Join-Path $root 'Specs/Core/Core_Example.md') -Value "`n[Escape](../../../outside.md)"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'LINK_ROOT_ESCAPE') | Should Be 1
    }

    It 'reports a stale link from immutable Done as informational' {
        $root = New-RSSpecFixtureRepository
        Add-Content -LiteralPath (Join-Path $root 'Specs/Plans/Done/History.md') -Value "`n[Old](missing.md)"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'LINK_MISSING_TARGET' -Severity 'Information') | Should Be 1
        (Get-RSFindingCount -Inventory $inventory -Code 'LINK_MISSING_TARGET' -Severity 'Error') | Should Be 0
    }

    It 'rejects a modified protected Done file' {
        $root = New-RSSpecFixtureRepository
        Add-Content -LiteralPath (Join-Path $root 'Specs/Plans/Done/History.md') -Value "changed"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        @($inventory.ProtectedDone.Modified).Count | Should Be 1
    }

    It 'admits the reviewed completed menu plan path without weakening Done protection' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/UI_MenuInformationArchitectureAndConsistency_2026-08-28.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed menu plan`n"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        @($inventory.ProtectedDone.Modified).Count | Should Be 0
    }

    It 'admits the reviewed completed menu remediation path without weakening Done protection' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/UI_MenuContextEncodingReviewRemediation_2026-08-29.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed menu remediation plan`n"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        @($inventory.ProtectedDone.Modified).Count | Should Be 0
    }

    It 'admits the reviewed retired File Operations popup plan without weakening Done protection' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/UI_FileOperationsPopupParallelProgressPendingPrompt_2026-08-28.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Retired File Operations popup plan`n"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        @($inventory.ProtectedDone.Modified).Count | Should Be 0
    }

    It 'admits the reviewed retired File Operations architecture-debt plan without weakening Done protection' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/Operation_FileOperations_PostCloseoutArchitectureDebt_2026-08-28.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Retired File Operations architecture-debt plan`n"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        @($inventory.ProtectedDone.Modified).Count | Should Be 0
    }

    It 'admits the retired File Operations reliability core plan without weakening Done protection' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/Operation_FileOperations_ReliabilityFirstSuccessorPlan_2026-08-29.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Retired File Operations reliability core plan`n"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        @($inventory.ProtectedDone.Modified).Count | Should Be 0
    }

    It 'admits the retired File Operations execution record without weakening Done protection' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/Operation_FileOperations_ReliabilityFirst_ExecutionRecord_2026-09-05.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Retired File Operations execution record`n"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        @($inventory.ProtectedDone.Modified).Count | Should Be 0
    }

    It 'admits only the reviewed completed Beeline residual record' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/FileOperations_BeelineResidualsAndProviderRoutes_2026-09-08.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed Beeline residual record`n"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        @($inventory.ProtectedDone.Modified).Count | Should Be 0

        $unreviewedPath = 'Specs/Plans/Done/FileOperations_BeelineResidualsAndProviderRoutes_2026-09-09.md'
        Set-RSFixtureFile -Root $root -RelativePath $unreviewedPath -Content "# Unreviewed sibling`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        @($inventory.ProtectedDone.Unexpected).Count | Should Be 1
        (@($inventory.ProtectedDone.Unexpected) -contains $unreviewedPath) | Should Be $true
    }

    It 'admits only the reviewed completed discovery-scope record' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/FileOperations_DiscoveryScopeAndLeafProgress_2026-09-13.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed discovery-scope record`n"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        @($inventory.ProtectedDone.Modified).Count | Should Be 0

        $unreviewedPath = 'Specs/Plans/Done/FileOperations_DiscoveryScopeAndLeafProgress_2026-09-14.md'
        Set-RSFixtureFile -Root $root -RelativePath $unreviewedPath -Content "# Unreviewed sibling`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        @($inventory.ProtectedDone.Unexpected).Count | Should Be 1
        (@($inventory.ProtectedDone.Unexpected) -contains $unreviewedPath) | Should Be $true
    }

    It 'admits only the reviewed completed direct-API discovery record' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/FileOperations_DirectApiDiscoveryContract_2026-09-14.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed direct-API discovery record`n"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        @($inventory.ProtectedDone.Modified).Count | Should Be 0

        $unreviewedPath = 'Specs/Plans/Done/FileOperations_DirectApiDiscoveryContract_2026-09-15.md'
        Set-RSFixtureFile -Root $root -RelativePath $unreviewedPath -Content "# Unreviewed sibling`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        @($inventory.ProtectedDone.Unexpected).Count | Should Be 1
        (@($inventory.ProtectedDone.Unexpected) -contains $unreviewedPath) | Should Be $true
    }

    It 'admits only the reviewed completed provider discovery parity record' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/FileOperations_ProviderDiscoveryParity_2026-09-14.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed provider discovery parity record`n"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        @($inventory.ProtectedDone.Modified).Count | Should Be 0

        $unreviewedPath = 'Specs/Plans/Done/FileOperations_ProviderDiscoveryParity_2026-09-15.md'
        Set-RSFixtureFile -Root $root -RelativePath $unreviewedPath -Content "# Unreviewed sibling`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        @($inventory.ProtectedDone.Unexpected).Count | Should Be 1
        (@($inventory.ProtectedDone.Unexpected) -contains $unreviewedPath) | Should Be $true
    }

    It 'admits only the reviewed completed S3 marker and residuals record' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/FileOperations_S3MarkerAndResiduals_2026-09-15.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed S3 marker and residuals record`n"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        @($inventory.ProtectedDone.Modified).Count | Should Be 0

        $unreviewedPath = 'Specs/Plans/Done/FileOperations_S3MarkerAndResiduals_2026-09-16.md'
        Set-RSFixtureFile -Root $root -RelativePath $unreviewedPath -Content "# Unreviewed sibling`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        @($inventory.ProtectedDone.Unexpected).Count | Should Be 1
        (@($inventory.ProtectedDone.Unexpected) -contains $unreviewedPath) | Should Be $true
    }

    It 'admits only the reviewed Terminal and File Operations follow-up record' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/Operation_ReviewFollowup_FOTerminalFocus_2026-08-19.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed Terminal and File Operations follow-up`n"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        @($inventory.ProtectedDone.Modified).Count | Should Be 0

        $unreviewedPath = 'Specs/Plans/Done/Operation_ReviewFollowup_FOTerminalFocus_2026-08-20.md'
        Set-RSFixtureFile -Root $root -RelativePath $unreviewedPath -Content "# Unreviewed sibling`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        @($inventory.ProtectedDone.Unexpected).Count | Should Be 1
        (@($inventory.ProtectedDone.Unexpected) -contains $unreviewedPath) | Should Be $true
    }

    It 'admits only the reviewed DxUi adoption closeout and supporting records' {
        foreach ($relativePath in @(
            'Specs/Plans/Done/DxUi_SharedLibraryAdoptionAndReleasePlan_2026-09-09.md',
            'Specs/Plans/Done/DxUi_SharedLibrary/GapAnalysis_2026-09-09.md',
            'Specs/Plans/Done/DxUi_SharedLibrary/RetirementDispositions_2026-09-12.md'
        )) {
            $root = New-RSSpecFixtureRepository
            Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Reviewed DxUi adoption record`n"
            $inventory = Get-RSFixtureInventory -Root $root
            (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
            (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
            @($inventory.ProtectedDone.Modified).Count | Should Be 0

            $unreviewedPath = $relativePath.Replace('.md', '_unreviewed.md')
            Set-RSFixtureFile -Root $root -RelativePath $unreviewedPath -Content "# Unreviewed sibling`n"
            $inventory = Get-RSFixtureInventory -Root $root
            (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
            @($inventory.ProtectedDone.Unexpected).Count | Should Be 1
            (@($inventory.ProtectedDone.Unexpected) -contains $unreviewedPath) | Should Be $true
        }
    }

    It 'rejects a tombstone above the line cap' {
        $root = New-RSSpecFixtureRepository
        $body = @('# Retired', '', '- **Status:** RETIRED', '- **Retirement date:** 2026-08-04', '- **Reason:** superseded', '- **Authoritative spec:** Core', '- `git show HEAD:path`')
        $body += @(1..40 | ForEach-Object { "extra line $_" })
        Set-RSFixtureFile -Root $root -RelativePath 'Specs/Plans/WIP/Retired.md' -Content ($body -join "`n")
        $indexPath = Join-Path $root 'Specs/Plans/WIP/README.md'
        $index = (Get-Content -LiteralPath $indexPath -Raw).Replace('| ACTIVE | I0 | [Plan](Plan.md) |', "| ACTIVE | I0 | [Plan](Plan.md) |`n| RETIRED | R0 | [Retired](Retired.md) |")
        Set-Content -LiteralPath $indexPath -Value $index -Encoding utf8
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'TOMBSTONE_INVALID') | Should Be 1
    }

    It 'classifies protected schema theme TestRuns and Mockups outside normative prose rules' {
        Get-RSSpecAuthorityClass -RelativePath 'Specs/SettingsStore.schema.json' | Should Be 'MachineContract'
        Get-RSSpecAuthorityClass -RelativePath 'Specs/Build/BuildReceipt.schema.json' | Should Be 'MachineContract'
        Get-RSSpecAuthorityClass -RelativePath 'Specs/Testing/ValidationPlan.schema.json' | Should Be 'MachineContract'
        Get-RSSpecAuthorityClass -RelativePath 'Specs/Themes/Default.json5' | Should Be 'RuntimeData'
        Get-RSSpecAuthorityClass -RelativePath 'Specs/TestRuns/run/README.md' | Should Be 'Evidence'
        Get-RSSpecAuthorityClass -RelativePath 'Specs/Mockups/design-qa.md' | Should Be 'Prototype'
        Get-RSSpecAuthorityClass -RelativePath 'Specs/Plans/WIP/Notes_Scratch.md' | Should Be 'HumanManagedNote'
    }

    It 'requires every normative file to have a consistency record and evidence anchors' {
        $root = New-RSSpecFixtureRepository
        Set-RSFixtureFile -Root $root -RelativePath 'Specs/Core/Core_Missing.md' -Content "# Missing ledger row`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'NORMATIVE_MISSING_CONSISTENCY') | Should Be 1

        $ledgerPath = Join-Path $root 'Specs/NormativeConsistency.json'
        $ledger = Get-Content -LiteralPath $ledgerPath -Raw | ConvertFrom-Json
        $ledger.documents[0].implementationAnchors = @()
        $ledger | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $ledgerPath -Encoding utf8
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'NORMATIVE_MISSING_EVIDENCE') | Should Be 1
    }

    It 'rejects UNVERIFIED and unresolved contradiction statuses in strict inventory' {
        $root = New-RSSpecFixtureRepository
        $ledgerPath = Join-Path $root 'Specs/NormativeConsistency.json'
        $ledger = Get-Content -LiteralPath $ledgerPath -Raw | ConvertFrom-Json
        $ledger.documents[0].status = 'CROSS_SPEC_CONTRADICTION'
        $ledger | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $ledgerPath -Encoding utf8
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'NORMATIVE_UNRESOLVED_CONSISTENCY') | Should Be 1
    }

    It 'rejects a WIP plan that claims authority over current contracts' {
        $root = New-RSSpecFixtureRepository
        Add-Content -LiteralPath (Join-Path $root 'Specs/Plans/WIP/Plan.md') -Value "`nThis plan is authoritative."
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'WIP_AUTHORITY_CLAIM') | Should Be 1
    }
}

if (Test-Path -LiteralPath $fixtureParent -PathType Container) {
    $resolvedFixtureParent = [IO.Path]::GetFullPath($fixtureParent)
    $expectedPrefix = [IO.Path]::GetFullPath((Join-Path $testRunContext.RunRoot 'scratch\tools-pester')) + [IO.Path]::DirectorySeparatorChar
    if ($resolvedFixtureParent.StartsWith($expectedPrefix, [StringComparison]::OrdinalIgnoreCase)) {
        Remove-Item -LiteralPath $resolvedFixtureParent -Recurse -Force
    }
}
