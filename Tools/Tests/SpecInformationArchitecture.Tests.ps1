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

    It 'accepts distinct WIP ranks that no archived plan claims' {
        $root = New-RSSpecFixtureRepository
        Set-RSFixtureFile -Root $root -RelativePath 'Specs/Plans/Done/History.md' -Content "# I7 Archived plan`n"
        & git -C $root add --all
        & git -C $root commit --quiet -m 'fixture archived plan'
        # Two distinct files: pointing two rows at Plan.md would add an unrelated WIP_DUPLICATE_INDEX error.
        Set-RSFixtureFile -Root $root -RelativePath 'Specs/Plans/WIP/Second.md' -Content "# Second fixture plan`n"
        $indexPath = Join-Path $root 'Specs/Plans/WIP/README.md'
        $index = Get-Content -LiteralPath $indexPath -Raw
        $index = $index.Replace('| ACTIVE | I0 | [Plan](Plan.md) |', "| ACTIVE | I0 | [Plan](Plan.md) |`n| HOLD | H0 | [Second](Second.md) |")
        Set-Content -LiteralPath $indexPath -Value $index -Encoding utf8
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'WIP_DUPLICATE_RANK') | Should Be 0
        (Get-RSFindingCount -Inventory $inventory -Code 'WIP_RANK_ARCHIVED_ID') | Should Be 0
        (Get-RSFindingCount -Inventory $inventory -Code 'WIP_DUPLICATE_INDEX') | Should Be 0
        @($inventory.BlockingFindings).Count | Should Be 0
        $inventory.WipIndexEntries[0].Rank | Should Be 'I0'
    }

    It 'rejects two WIP index rows that share a rank' {
        $root = New-RSSpecFixtureRepository
        Set-RSFixtureFile -Root $root -RelativePath 'Specs/Plans/WIP/Second.md' -Content "# Second fixture plan`n"
        $indexPath = Join-Path $root 'Specs/Plans/WIP/README.md'
        $index = Get-Content -LiteralPath $indexPath -Raw
        $index = $index.Replace('| ACTIVE | I0 | [Plan](Plan.md) |', "| ACTIVE | I0 | [Plan](Plan.md) |`n| ACTIVE | I0 | [Second](Second.md) |")
        Set-Content -LiteralPath $indexPath -Value $index -Encoding utf8
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'WIP_DUPLICATE_RANK' -Severity 'Error') | Should Be 1
        (Get-RSFindingCount -Inventory $inventory -Code 'WIP_DUPLICATE_INDEX') | Should Be 0
        (Get-RSFindingCount -Inventory $inventory -Code 'WIP_RANK_ARCHIVED_ID') | Should Be 0
        @($inventory.BlockingFindings | Where-Object Code -eq 'WIP_DUPLICATE_RANK').Count | Should Be 1
    }

    It 'rejects a WIP rank that an archived plan title already claims' {
        $root = New-RSSpecFixtureRepository
        Set-RSFixtureFile -Root $root -RelativePath 'Specs/Plans/Done/History.md' -Content "# I0 Archived plan`n"
        & git -C $root add --all
        & git -C $root commit --quiet -m 'fixture archived plan'
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'WIP_RANK_ARCHIVED_ID' -Severity 'Error') | Should Be 1
        (Get-RSFindingCount -Inventory $inventory -Code 'WIP_DUPLICATE_RANK') | Should Be 0
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        $finding = @($inventory.Findings | Where-Object Code -eq 'WIP_RANK_ARCHIVED_ID')[0]
        ($finding.Message -match 'Specs/Plans/Done/History\.md') | Should Be $true
    }

    It 'does not treat an unrelated archived title as a claimed plan ID' {
        $root = New-RSSpecFixtureRepository
        Set-RSFixtureFile -Root $root -RelativePath 'Specs/Plans/Done/History.md' -Content "# Notes about I0 and the I0x series`n"
        & git -C $root add --all
        & git -C $root commit --quiet -m 'fixture archived note'
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'WIP_RANK_ARCHIVED_ID') | Should Be 0
    }

    It 'does not read an ordinary title word ending in a digit as a plan ID' {
        $root = New-RSSpecFixtureRepository
        Set-RSFixtureFile -Root $root -RelativePath 'Specs/Plans/Done/History.md' -Content "# S3 Interactive directory probe`n"
        & git -C $root add --all
        & git -C $root commit --quiet -m 'fixture archived S3 title'
        $indexPath = Join-Path $root 'Specs/Plans/WIP/README.md'
        $index = (Get-Content -LiteralPath $indexPath -Raw).Replace('| ACTIVE | I0 | [Plan](Plan.md) |', '| ACTIVE | S3 | [Plan](Plan.md) |')
        Set-Content -LiteralPath $indexPath -Value $index -Encoding utf8
        $inventory = Get-RSFixtureInventory -Root $root

        $inventory.WipIndexEntries[0].Rank | Should Be 'S3'
        (Get-RSFindingCount -Inventory $inventory -Code 'WIP_RANK_ARCHIVED_ID') | Should Be 0
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

    It 'preserves GitHub identifier and punctuation spacing in heading fragments' {
        ConvertTo-RSGitHubHeadingSlug -Heading 'Application Shell, Startup & Lifecycle' | Should Be 'application-shell-startup--lifecycle'
        ConvertTo-RSGitHubHeadingSlug -Heading 'Cache coherence: the `_version` model' | Should Be 'cache-coherence-the-_version-model'
        ConvertTo-RSGitHubHeadingSlug -Heading 'Two  spaces' | Should Be 'two--spaces'
        ConvertTo-RSGitHubHeadingSlug -Heading 'BR-8 — Preferences Keyboard search input identity' | Should Be 'br-8--preferences-keyboard-search-input-identity'
        ConvertTo-RSGitHubHeadingSlug -Heading 'Identifier foo_bar and __bold__ text' | Should Be 'identifier-foo_bar-and-bold-text'
        ConvertTo-RSGitHubHeadingSlug -Heading 'Literal `_code_` and _emphasis_' | Should Be 'literal-_code_-and-emphasis'
    }

    It 'accepts real links to identifier and ampersand heading fragments' {
        $root = New-RSSpecFixtureRepository
        Add-Content -LiteralPath (Join-Path $root 'Specs/Core/Core_Example.md') -Value @'

## Startup & Lifecycle
[Lifecycle](Core_Example.md#startup--lifecycle)

## The `_version` model
[Version](Core_Example.md#the-_version-model)
'@
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'LINK_MISSING_ANCHOR' -Severity 'Error') | Should Be 0
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

    It 'keeps a missing link visible after line-leading triple-backtick inline code' {
        $text = @'
```literal```

[Missing](MissingAfterInline.md)
'@
        $links = @(Get-RSMarkdownLinks -Text $text)
        $links.Count | Should Be 1
        $links[0].Target | Should Be 'MissingAfterInline.md'
        $links[0].Line | Should Be 3
        $inlineLink = $text.Replace('literal', '[fake](missing-inline.md)')
        @(Get-RSMarkdownLinks -Text $inlineLink).Count | Should Be 1
        (Get-RSMarkdownCodeMask -Text $text).Length | Should Be $text.Length

        $root = New-RSSpecFixtureRepository
        Add-Content -LiteralPath (Join-Path $root 'Specs/Core/Core_Example.md') -Value $text
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'LINK_MISSING_TARGET' -Severity 'Error') | Should Be 1
    }

    It 'keeps a real backtick fence open when a matching run has trailing nonwhitespace' {
        $text = @'
   ```markdown
[Hidden](missing-interior.md)
```not-a-closer
[Still hidden](missing-after-false-closer.md)
# Hidden heading
   ````CLOSING_WHITESPACE
# Visible heading
[Shown](outside.md)
'@
        $text = $text.Replace('CLOSING_WHITESPACE', " `t")
        $links = @(Get-RSMarkdownLinks -Text $text)
        $links.Count | Should Be 1
        $links[0].Target | Should Be 'outside.md'
        $headings = @(Get-RSMarkdownHeadings -Text $text)
        $headings.Count | Should Be 1
        $headings[0].Title | Should Be 'Visible heading'
        (Get-RSMarkdownCodeMask -Text $text).Length | Should Be $text.Length
    }

    It 'allows backticks in tilde fence info and requires a whitespace-only closing tail' {
        $text = @'
~~~example`info`
[Hidden](missing-tilde-interior.md)
~~~not-a-closer
[Still hidden](missing-after-false-tilde-closer.md)
# Hidden heading
~~~CLOSING_WHITESPACE
# Visible heading
[Shown](outside.md)
'@
        $text = $text.Replace('CLOSING_WHITESPACE', " `t")
        $links = @(Get-RSMarkdownLinks -Text $text)
        $links.Count | Should Be 1
        $links[0].Target | Should Be 'outside.md'
        $headings = @(Get-RSMarkdownHeadings -Text $text)
        $headings.Count | Should Be 1
        $headings[0].Title | Should Be 'Visible heading'
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

    It 'admits only the reviewed completed I35 copy/move UX and reliability record' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/FileOperations_CopyMoveUxAndReliabilityReview_2026-10-11.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed I35 record`n"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        @($inventory.ProtectedDone.Modified).Count | Should Be 0

        $unreviewedPath = 'Specs/Plans/Done/FileOperations_CopyMoveUxAndReliabilityReview_2026-10-12.md'
        Set-RSFixtureFile -Root $root -RelativePath $unreviewedPath -Content "# Unreviewed sibling`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        @($inventory.ProtectedDone.Unexpected).Count | Should Be 1
        (@($inventory.ProtectedDone.Unexpected) -contains $unreviewedPath) | Should Be $true
    }

    It 'admits only the reviewed completed I37 folder-copy progress record' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/FileOperations_FolderCopyLiveProgress_2026-10-14.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed I37 record`n"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        @($inventory.ProtectedDone.Modified).Count | Should Be 0

        $unreviewedPath = 'Specs/Plans/Done/FileOperations_FolderCopyLiveProgress_2026-10-15.md'
        Set-RSFixtureFile -Root $root -RelativePath $unreviewedPath -Content "# Unreviewed sibling`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        @($inventory.ProtectedDone.Unexpected).Count | Should Be 1
        (@($inventory.ProtectedDone.Unexpected) -contains $unreviewedPath) | Should Be $true
    }

    It 'admits only the reviewed completed card information hierarchy record' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/FileOperations_CardInformationHierarchy_2026-10-16.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed card hierarchy record`n"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        @($inventory.ProtectedDone.Modified).Count | Should Be 0

        $unreviewedPath = 'Specs/Plans/Done/FileOperations_CardInformationHierarchy_2026-10-17.md'
        Set-RSFixtureFile -Root $root -RelativePath $unreviewedPath -Content "# Unreviewed sibling`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        @($inventory.ProtectedDone.Unexpected).Count | Should Be 1
        (@($inventory.ProtectedDone.Unexpected) -contains $unreviewedPath) | Should Be $true
    }

    It 'admits only the reviewed completed I20 terminal font and plugin numeric record' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/Terminal_FontResolutionAndPluginConfigNumericContract_2026-09-10.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed I20 record`n"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        @($inventory.ProtectedDone.Modified).Count | Should Be 0

        $unreviewedPath = 'Specs/Plans/Done/Terminal_FontResolutionAndPluginConfigNumericContract_2026-09-11.md'
        Set-RSFixtureFile -Root $root -RelativePath $unreviewedPath -Content "# Unreviewed sibling`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        @($inventory.ProtectedDone.Unexpected).Count | Should Be 1
        (@($inventory.ProtectedDone.Unexpected) -contains $unreviewedPath) | Should Be $true
    }

    It 'admits only the reviewed completed I28 release dependency recovery record' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/Build_ReleaseDependencyRecovery_2026-10-02.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed I28 record`n"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        @($inventory.ProtectedDone.Modified).Count | Should Be 0

        $unreviewedPath = 'Specs/Plans/Done/Build_ReleaseDependencyRecovery_2026-10-03.md'
        Set-RSFixtureFile -Root $root -RelativePath $unreviewedPath -Content "# Unreviewed sibling`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        @($inventory.ProtectedDone.Unexpected).Count | Should Be 1
        (@($inventory.ProtectedDone.Unexpected) -contains $unreviewedPath) | Should Be $true
    }

    It 'admits only the reviewed completed I38 dialog keyboard focus record' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/FileOperations_DialogKeyboardFocus_2026-10-17.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed I38 record`n"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        @($inventory.ProtectedDone.Modified).Count | Should Be 0

        $unreviewedPath = 'Specs/Plans/Done/FileOperations_DialogKeyboardFocus_2026-10-18.md'
        Set-RSFixtureFile -Root $root -RelativePath $unreviewedPath -Content "# Unreviewed sibling`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        @($inventory.ProtectedDone.Unexpected).Count | Should Be 1
        (@($inventory.ProtectedDone.Unexpected) -contains $unreviewedPath) | Should Be $true
    }

    It 'admits only the reviewed completed I42 Copy Move selection convergence record' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/FileOperations_SelectionConvergenceAndParallelMove_2026-10-18.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed I42 record`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        $unreviewedPath = 'Specs/Plans/Done/FileOperations_SelectionConvergenceAndParallelMove_2026-10-19.md'
        Set-RSFixtureFile -Root $root -RelativePath $unreviewedPath -Content "# Unreviewed sibling`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        (@($inventory.ProtectedDone.Unexpected) -contains $unreviewedPath) | Should Be $true
    }

    It 'admits the six reviewed October 19 records and rejects every sibling and extra companion' {
        $root = New-RSSpecFixtureRepository
        $relativePaths = @(
            'Specs/Plans/Done/FileOperations_CardCountsAndThroughput_2026-10-19.md',
            'Specs/Plans/Done/FileOperations_DocumentationPrivacy_2026-10-19.md',
            'Specs/Plans/Done/FileOperations_LastNoteWrapping_2026-10-19.md',
            'Specs/Plans/Done/FileOperations_UIReviewFixes_2026-10-19.md',
            'Specs/Plans/Done/ScreenshotSyntheticData_2026-10-19.md',
            'Specs/Plans/Done/StashReview_2026-10-19.md'
        )
        foreach ($relativePath in $relativePaths) {
            Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Reviewed completed record`n"
        }
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        @($inventory.ProtectedDone.Modified).Count | Should Be 0
        foreach ($relativePath in $relativePaths) {
            (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        }

        $unreviewedPaths = @($relativePaths | ForEach-Object { $_.Replace('2026-10-19', '2026-10-20') })
        foreach ($relativePath in $unreviewedPaths) {
            Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Unreviewed sibling`n"
        }
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        @($inventory.ProtectedDone.Unexpected).Count | Should Be $unreviewedPaths.Count
        foreach ($relativePath in $unreviewedPaths) {
            (@($inventory.ProtectedDone.Unexpected) -contains $relativePath) | Should Be $true
        }

        $companionPath = 'Specs/Plans/Done/StashReview_2026-10-19/unreviewed.md'
        Set-RSFixtureFile -Root $root -RelativePath $companionPath -Content "# Unreviewed companion`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        @($inventory.ProtectedDone.Unexpected).Count | Should Be ($unreviewedPaths.Count + 1)
        (@($inventory.ProtectedDone.Unexpected) -contains $companionPath) | Should Be $true
    }

    It 'admits only the reviewed PowerShell tooling closeout record' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/PowerShellTooling_2026-10-28.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed tooling migration`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        foreach ($path in @('Specs/Plans/Done/PowerShellTooling_2026-10-29.md', 'Specs/Plans/Done/PowerShellTooling_2026-10-28/unreviewed.md')) {
            Set-RSFixtureFile -Root $root -RelativePath $path -Content "# Unreviewed companion`n"
        }
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        @($inventory.ProtectedDone.Unexpected).Count | Should Be 2
    }

    It 'admits only the reviewed dependency integration closeout record' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/DependencyUpdateBatch_2026-10-29.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed dependency integration`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        $unreviewed = @(
            'Specs/Plans/Done/DependencyUpdateBatch_2026-10-30.md',
            'Specs/Plans/Done/DependencyUpdateBatch_2026-10-29/unreviewed.md'
        )
        foreach ($path in $unreviewed) {
            Set-RSFixtureFile -Root $root -RelativePath $path -Content "# Unreviewed companion`n"
        }
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        @($inventory.ProtectedDone.Unexpected).Count | Should Be 2
        foreach ($path in $unreviewed) {
            (@($inventory.ProtectedDone.Unexpected) -contains $path) | Should Be $true
        }
    }

    It 'admits only the reviewed agent-skill remediation closeout record' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/AgentSkillsAudit_RedTeam_2026-10-24.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed guidance remediation`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        foreach ($path in @('Specs/Plans/Done/AgentSkillsAudit_RedTeam_2026-10-25.md', 'Specs/Plans/Done/AgentSkillsAudit_RedTeam_2026-10-24/unreviewed.md')) {
            Set-RSFixtureFile -Root $root -RelativePath $path -Content "# Unreviewed companion`n"
        }
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        @($inventory.ProtectedDone.Unexpected).Count | Should Be 2
    }

    It 'admits only the reviewed completed waiting-card record' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/FileOperations_WaitingCards_2026-10-20.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed waiting-card record`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        foreach ($unreviewedPath in @(
            'Specs/Plans/Done/FileOperations_WaitingCards_2026-10-21.md',
            'Specs/Plans/Done/FileOperations_WaitingCards_2026-10-20/unreviewed.md'
        )) {
            Set-RSFixtureFile -Root $root -RelativePath $unreviewedPath -Content "# Unreviewed companion or sibling`n"
        }
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        @($inventory.ProtectedDone.Unexpected).Count | Should Be 2
    }

    It 'admits only the reviewed completed test-run evidence archival record' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/TestRunEvidenceArchival_2026-10-21.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed test-run retention record`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        foreach ($path in @('Specs/Plans/Done/TestRunEvidenceArchival_2026-10-22.md', 'Specs/Plans/Done/TestRunEvidenceArchival_2026-10-21/unreviewed.md')) {
            Set-RSFixtureFile -Root $root -RelativePath $path -Content "# Unreviewed record`n"
        }
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        @($inventory.ProtectedDone.Unexpected).Count | Should Be 2
    }

    It 'admits only the reviewed completed I44 sort paint measurement record' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/UI_ReleaseSortPaintLatency_2026-10-20.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed I44 record`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        $unreviewedPath = 'Specs/Plans/Done/UI_ReleaseSortPaintLatency_2026-10-21.md'
        Set-RSFixtureFile -Root $root -RelativePath $unreviewedPath -Content "# Unreviewed sibling`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        (@($inventory.ProtectedDone.Unexpected) -contains $unreviewedPath) | Should Be $true
    }

    It 'admits only the reviewed completed I45 writer CPU record' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/UI_DebugSortPaintCpu_2026-10-20.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed I45 record`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        $unreviewedPath = 'Specs/Plans/Done/UI_DebugSortPaintCpu_2026-10-21.md'
        Set-RSFixtureFile -Root $root -RelativePath $unreviewedPath -Content "# Unreviewed sibling`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        (@($inventory.ProtectedDone.Unexpected) -contains $unreviewedPath) | Should Be $true
    }

    It 'admits only the reviewed completed I43 sorting row movement record' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/UI_DebugSortingRowMoves_2026-10-20.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed I43 record`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        $unreviewedPath = 'Specs/Plans/Done/UI_DebugSortingRowMoves_2026-10-21.md'
        Set-RSFixtureFile -Root $root -RelativePath $unreviewedPath -Content "# Unreviewed sibling`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        (@($inventory.ProtectedDone.Unexpected) -contains $unreviewedPath) | Should Be $true
    }

    It 'admits only the reviewed completed I41 natural sorting performance record' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/UI_NaturalSortingPerformance_2026-10-17.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed I41 record`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        $unreviewedPath = 'Specs/Plans/Done/UI_NaturalSortingPerformance_2026-10-18.md'
        Set-RSFixtureFile -Root $root -RelativePath $unreviewedPath -Content "# Unreviewed sibling`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        (@($inventory.ProtectedDone.Unexpected) -contains $unreviewedPath) | Should Be $true
    }

    It 'admits only the reviewed completed I40 natural name sorting record' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/UI_NaturalNameSorting_2026-10-17.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed I40 record`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        $unreviewedPath = 'Specs/Plans/Done/UI_NaturalNameSorting_2026-10-18.md'
        Set-RSFixtureFile -Root $root -RelativePath $unreviewedPath -Content "# Unreviewed sibling`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        (@($inventory.ProtectedDone.Unexpected) -contains $unreviewedPath) | Should Be $true
    }

    It 'admits only the reviewed completed I39 breadcrumb natural sorting record' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/Navigation_BreadcrumbNaturalSorting_2026-10-17.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed I39 record`n"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        @($inventory.ProtectedDone.Modified).Count | Should Be 0

        $unreviewedPath = 'Specs/Plans/Done/Navigation_BreadcrumbNaturalSorting_2026-10-18.md'
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

    It 'admits only the exact reviewed I26 File Operations UI plan path' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/UI_FileOperationsReview_2026-09-16.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed I26 File Operations UI review`n"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        @($inventory.ProtectedDone.Modified).Count | Should Be 0

        $unreviewedPath = 'Specs/Plans/Done/UI_FileOperationsReview_2026-09-17.md'
        Set-RSFixtureFile -Root $root -RelativePath $unreviewedPath -Content "# Unreviewed sibling`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        @($inventory.ProtectedDone.Unexpected).Count | Should Be 1
        (@($inventory.ProtectedDone.Unexpected) -contains $unreviewedPath) | Should Be $true
    }

    It 'admits only the exact reviewed I27 File Operations safety remediation path' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/UI_FileOperationsSafetyRemediation_2026-10-01.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Completed I27 safety remediation`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        @($inventory.ProtectedDone.Modified).Count | Should Be 0
        $unreviewedPath = 'Specs/Plans/Done/UI_FileOperationsSafetyRemediation_2026-10-02.md'
        Set-RSFixtureFile -Root $root -RelativePath $unreviewedPath -Content "# Unreviewed sibling`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        (@($inventory.ProtectedDone.Unexpected) -contains $unreviewedPath) | Should Be $true
    }

    It 'admits reviewed I27 gallery companions and rejects an unreviewed image in the same collection' {
        $root = New-RSSpecFixtureRepository
        $collection = 'Specs/Plans/Done/UI_FileOperationsSafetyRemediation_2026-10-01/implementation/2026-10-04-final-gallery/20261003T232201Z-75488-0a7264eaad564dcd8aa1014d5093c463'
        $navigation = "$collection/gallery.html"
        $capture = "$collection/fileops/visual-gallery/01-empty__dark__480dip__fr-FR__display1__144dpi.png"
        Set-RSFixtureFile -Root $root -RelativePath $navigation -Content 'Reviewed navigation path fixture'
        Set-RSFixtureFile -Root $root -RelativePath $capture -Content 'Reviewed capture path fixture, not pixel qualification'
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $navigation) | Should Be $true
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $capture) | Should Be $true
        @($inventory.ProtectedDone.Modified).Count | Should Be 0

        $unreviewedPath = "$collection/fileops/visual-gallery/unreviewed-conflict.png"
        Set-RSFixtureFile -Root $root -RelativePath $unreviewedPath -Content 'Unreviewed sibling path fixture'
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        @($inventory.ProtectedDone.Unexpected).Count | Should Be 1
        (@($inventory.ProtectedDone.Unexpected) -contains $unreviewedPath) | Should Be $true
    }

    It 'admits only the exact reviewed I30 delivered review and implementation record' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/CodeReview_MasterMerges_Delivered_2026-10-10.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Delivered review and implementation history record`n"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        @($inventory.ProtectedDone.Modified).Count | Should Be 0

        $unreviewedPath = 'Specs/Plans/Done/CodeReview_MasterMerges_Delivered_2026-10-11.md'
        Set-RSFixtureFile -Root $root -RelativePath $unreviewedPath -Content "# Unreviewed sibling`n"
        $inventory = Get-RSFixtureInventory -Root $root
        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 1
        @($inventory.ProtectedDone.Unexpected).Count | Should Be 1
        (@($inventory.ProtectedDone.Unexpected) -contains $unreviewedPath) | Should Be $true
    }

    It 'admits only the exact reviewed I30 findings ledger inside its archived directory' {
        $root = New-RSSpecFixtureRepository
        $relativePath = 'Specs/Plans/Done/CodeReview_MasterMerges_2026-10-05/findings-ledger.md'
        Set-RSFixtureFile -Root $root -RelativePath $relativePath -Content "# Archived findings ledger`n"
        $inventory = Get-RSFixtureInventory -Root $root

        (Get-RSFindingCount -Inventory $inventory -Code 'DONE_HISTORY_CHANGED') | Should Be 0
        (@($inventory.ProtectedDone.AllowedAdditional) -contains $relativePath) | Should Be $true
        @($inventory.ProtectedDone.Modified).Count | Should Be 0

        $unreviewedPath = 'Specs/Plans/Done/CodeReview_MasterMerges_2026-10-05/findings-ledger-addendum.md'
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

Describe 'Markdown fence indentation boundaries' {
    It 'keeps unindented prose visible after an opener prefixed by <Name>' -TestCases @(
        @{Name='four spaces'; Indent='    '},
        @{Name='a tab'; Indent="`t"}
    ) {
        param($Name, $Indent)
        $text = $Indent + '```markdown' + "`n# Visible heading`n[Missing](missing-unindented.md)"
        $links = @(Get-RSMarkdownLinks -Text $text)
        $links.Count | Should Be 1
        $links[0].Target | Should Be 'missing-unindented.md'
        $headings = @(Get-RSMarkdownHeadings -Text $text)
        $headings.Count | Should Be 1
        $headings[0].Title | Should Be 'Visible heading'
    }

    It 'keeps a valid fence open across a closer prefixed by <Name>' -TestCases @(
        @{Name='four spaces'; Indent='    '},
        @{Name='a tab'; Indent="`t"}
    ) {
        param($Name, $Indent)
        $text = '```markdown' + "`n" + $Indent + '```' + "`n[Hidden](missing-interior.md)`n# Hidden heading`n" + '```' + "`n[Shown](outside.md)`n# Visible heading"
        $links = @(Get-RSMarkdownLinks -Text $text)
        $links.Count | Should Be 1
        $links[0].Target | Should Be 'outside.md'
        $headings = @(Get-RSMarkdownHeadings -Text $text)
        $headings.Count | Should Be 1
        $headings[0].Title | Should Be 'Visible heading'
    }
}
