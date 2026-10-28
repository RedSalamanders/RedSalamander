Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Drift-prevention contracts for the documentation-completeness work
# (Operation_Codex_DocumentationCompleteness). These guard regression
# classes the documentation audit found: published capture identity/privacy,
# keyboard shortcut/spec drift away from
# the code of record, and settings keys read/written in code but missing from the
# schema. See Specs/Plans/Done/Operation_Codex_DocumentationCompleteness_2026-06-18.md.

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'Tools\Modules\Tooling\SpecInformationArchitecture.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot 'TestSupport.psm1') -Force

Describe 'Agent guidance packages' {
    It 'uses the canonical Markdown parser for balanced targets and inline code' {
        $sample = '[Contract](contract(v1).md) and `[example](missing.md)`'
        $links = @(Get-RSMarkdownLinks -Text $sample)
        $links.Count | Should Be 1
        $links[0].Target | Should Be 'contract(v1).md'
    }

    It 'keeps owned skill and agent references usable without loading other checkouts' {
        $roots = @('.github/skills', '.agents/skills', '.github/agents')
        $documents = [Collections.Generic.List[string]]::new()
        foreach ($relativeRoot in $roots) {
            $pending = [Collections.Generic.Stack[string]]::new()
            $pending.Push((Join-Path $repoRoot $relativeRoot))
            while ($pending.Count) {
                foreach ($entry in Get-ChildItem -LiteralPath $pending.Pop() -Force) {
                    ($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) | Should Be 0
                    if ($entry.PSIsContainer) { $pending.Push($entry.FullName) }
                    elseif ($entry.Extension -ieq '.md') { $documents.Add($entry.FullName) }
                }
            }
        }
        foreach ($document in $documents) {
            $content = [IO.File]::ReadAllText($document, [Text.UTF8Encoding]::new($false, $true))
            [string]::IsNullOrWhiteSpace($content) | Should Be $false
            foreach ($link in Get-RSMarkdownLinks -Text $content) {
                $target = [string]$link.Target
                if ($target -match '^(?:https?:|mailto:|#|//)') { continue }
                $file = [Uri]::UnescapeDataString(($target -split '#', 2)[0])
                $resolved = [IO.Path]::GetFullPath((Join-Path (Split-Path $document) $file))
                $relative = [IO.Path]::GetRelativePath($repoRoot, $resolved)
                [IO.Path]::IsPathRooted($relative) | Should Be $false
                ($relative -eq '..' -or $relative.StartsWith('../') -or $relative.StartsWith('..\')) | Should Be $false
                (Test-Path -LiteralPath $resolved) | Should Be $true
            }
        }
    }
}

Describe 'Specification information architecture integration' {
    # The protected Done history is measured from base commit 4a3a6c9dc3d094cb2269c438de7f3ae24ab5e853, which a squashed or
    # shallow clone (hosted runners check out one commit) does not carry.
    & git -C $repoRoot rev-parse --verify --quiet '4a3a6c9dc3d094cb2269c438de7f3ae24ab5e853^{commit}' *> $null
    $protectedHistoryBaseAvailable = ($LASTEXITCODE -eq 0)
    if (-not $protectedHistoryBaseAvailable) {
        Write-Host 'Skipping the protected Done history contract: base commit 4a3a6c9dc3d094cb2269c438de7f3ae24ab5e853 is not in this (squashed or shallow) clone.'
    }
    It 'keeps every approved post-Terminal Done history blob unchanged' -Skip:(-not $protectedHistoryBaseAvailable) {
        $history = Test-RSProtectedDoneHistory -RepositoryRoot $repoRoot -BaseCommit '4a3a6c9dc3d094cb2269c438de7f3ae24ab5e853'

        $history.Available | Should Be $true
        $history.BaselineCount | Should Be 133
        $history.IsValid | Should Be $true
    }
}

Describe 'Visible native audit contracts' {
    It 'keeps dependency trees outside the product scan and classifies every current residual' {
        $audit = Get-RSText -Path 'Tools\Audit-RemainingWin32UiDependencies.ps1'

        $audit | Should Match '\\vcpkg_installed\\'
        $audit | Should Not Match 'Common\\DxUi\\'
        $audit | Should Match 'Tests\\TestSupport\\TestSupport\.Tests\.DirectedSelfTestInputWarning\.h'
        $audit | Should Not Match 'Tests\\DxUiTests\\'
        { & (Join-Path $repoRoot 'Tools\Audit-ComctlReportSurfaces.ps1') | Out-Null } | Should Not Throw
        { & (Join-Path $repoRoot 'Tools\Audit-VisibleNativeSurfaces.ps1') | Out-Null } | Should Not Throw
        { & (Join-Path $repoRoot 'Tools\Audit-RemainingWin32UiDependencies.ps1') -FailOnFindings | Out-Null } | Should Not Throw
    }
}

Describe 'Keyboard shortcut documentation drift contracts' {
    It 'binds the Alt+arrow folder-view chords to history and selected-name commands in code' {
        $source = Get-RSText -Path 'RedSalamander\ShortcutDefaults.cpp'

        $source | Should Match 'VK_DOWN,\s*ShortcutManager::kModAlt,\s*L"cmd/pane/selection/goToNextSelectedName"'
        $source | Should Match 'VK_UP,\s*ShortcutManager::kModAlt,\s*L"cmd/pane/selection/goToPreviousSelectedName"'
        $source | Should Match 'VK_LEFT,\s*ShortcutManager::kModAlt,\s*L"cmd/pane/historyBack"'
        $source | Should Match 'VK_RIGHT,\s*ShortcutManager::kModAlt,\s*L"cmd/pane/historyForward"'
    }

    It 'documents the Alt+arrow chords consistently in docs/KeyboardShortcuts.md' {
        $doc = Get-RSText -Path 'docs\KeyboardShortcuts.md'

        $doc | Should Match 'Alt\+Left'
        $doc | Should Match 'Alt\+Right'
        $doc | Should Match 'Alt\+Up'
        $doc | Should Match 'Alt\+Down'
        $doc | Should Match 'history back'
        $doc | Should Match 'history forward'
        $doc | Should Match 'previous selected name'
        $doc | Should Match 'next selected name'
    }

    It 'keeps the implemented folderView keyboard spec aligned with the Alt+arrow command ids' {
        $spec = Get-RSText -Path 'Specs\UI\UI_CommandMenuKeyboard.md'

        $spec | Should Match 'cmd/pane/historyBack'
        $spec | Should Match 'cmd/pane/historyForward'
        $spec | Should Match 'cmd/pane/selection/goToPreviousSelectedName'
        $spec | Should Match 'cmd/pane/selection/goToNextSelectedName'
    }
}

Describe 'Settings schema vs code key-coverage contracts' {
    It 'accepts every canonical formatted theme function spelling' {
        $schema = Get-Content -LiteralPath (Join-Path $repoRoot 'Specs\SettingsStore.schema.json') -Raw | ConvertFrom-Json
        $functionPattern = $schema.'$defs'.themeColorSource.anyOf[1].pattern
        $canonicalExpressions = @(
            'ref(palette.base)',
            'lighten(palette.base,0.2)',
            'darken(palette.base,0.2)',
            'alpha(palette.base,0.5)',
            'blend(palette.black,palette.white,0.25)',
            'contrast(palette.base)',
            'perceptualTone(palette.base,60)',
            'ensureContrast(palette.base,palette.white,4.5)',
            'harmonize(palette.base,palette.red,0.25)',
            'systemColor(window)',
            'tone(palette.white,palette.black)',
            'seededRainbow(runtime.seed,0.7,0.9,1,15)',
            'seededChoice(runtime.seed,palette.red,palette.green)'
        )

        foreach ($expression in $canonicalExpressions) {
            [regex]::IsMatch($expression, $functionPattern) | Should Be $true
        }
        $schema.'$defs'.themeColorSource.anyOf[2].enum | Should Be 'systemAccent()'
    }

    It 'reads connections.allowInsecureTlsInAutomation in the settings store' {
        $source = Get-RSText -Path 'Common\Common\SettingsStore.cpp'

        $source | Should Match 'allowInsecureTlsInAutomation'
    }

    It 'declares connections.allowInsecureTlsInAutomation in the settings schema' {
        $schema = Get-RSText -Path 'Specs\SettingsStore.schema.json'

        $schema | Should Match '"allowInsecureTlsInAutomation"'
    }

    It 'declares the Windows Hello connection keys in the settings schema' {
        $schema = Get-RSText -Path 'Specs\SettingsStore.schema.json'

        $schema | Should Match '"bypassWindowsHello"'
        $schema | Should Match '"windowsHelloReauthTimeoutMinute"'
    }

    It 'keeps the settings schema parseable as JSON' {
        $schemaPath = Join-Path $repoRoot 'Specs\SettingsStore.schema.json'

        { Get-Content -LiteralPath $schemaPath -Raw | ConvertFrom-Json | Out-Null } | Should Not Throw
    }

    It 'documents the Monitor filter mask default as 63 (six message types) in the schema' {
        $schema = Get-RSText -Path 'Specs\SettingsStore.schema.json'

        $schema | Should Match '"mask":\s*\{[^}]*"maximum":\s*63[^}]*"default":\s*63'
    }
}

Describe 'File Operations popup documentation drift contracts' {
    It 'tracks split popup cases in source and domain specs without freezing them in coverage policy' {
        $coverage = Get-RSText -Path 'Specs\Testing\Testing_TestCoverage.md'
        $fileOperations = Get-RSText -Path 'Specs\FileSystem\FileSystem_FileOperations.md'
        $caseSource = Get-RSText -Path 'RedSalamander\SelfTest\Commands\Commands.Tests.FileOps.cpp'
        $uxPlan = Get-RSText -Path 'Specs\Plans\Done\UI_FileOperationsPopupUxRefinementPlan_2026-07-07.md'
        $splitCases = @(
            'cmd_pane_fileops_popup_progress_contracts',
            'cmd_pane_fileops_conflict_metadata_uses_single_provider_roundtrip',
            'cmd_pane_fileops_conflict_prompt_metadata_and_actions',
            'cmd_pane_fileops_popup_presentation_settings_and_taskbar',
            'cmd_pane_fileops_completed_group_and_navigation'
        )

        foreach ($caseName in $splitCases) {
            $caseSource | Should Match ([Regex]::Escape($caseName))
            $coverage | Should Not Match ([Regex]::Escape($caseName))
        }
        $coverage | Should Match 'Get-TestInventory\.ps1'
        $coverage | Should Match '\.build/SpecInventory/test-inventory\.json'
        $fileOperations | Should Match 'Get-TestInventory\.ps1'
        $fileOperations | Should Match 'File Operations selftest'
        $caseSource | Should Not Match 'cmd_pane_fileops_conflict_prompt_compacts_actions'
        $fileOperations | Should Not Match 'cmd_pane_fileops_conflict_prompt_compacts_actions'
        $uxPlan | Should Match '\[UI_FileOperationsPopupCodeReviewRemediation_2026-07-10\.md\]\(\.\./Done/UI_FileOperationsPopupCodeReviewRemediation_2026-07-10\.md\)'
        $uxPlan | Should Match 'scoped correctness/architecture review is complete'
        $uxPlan | Should Not Match 'review-remediation evidence gate keep the plan in WIP'
    }
}

Describe 'Historical audit governance contracts' {
    It 'marks every review campaign document as a non-live historical snapshot' {
        $reviewFiles = @(Get-ChildItem -LiteralPath (Join-Path $repoRoot 'Specs\Reviews') -File -Filter '*.md')
        $reviewFiles.Count | Should BeGreaterThan 0

        foreach ($file in $reviewFiles) {
            $text = Get-Content -LiteralPath $file.FullName -Raw
            $text | Should Match 'Historical snapshot — not a live work queue'
            $text | Should Match '2026-07-17'
            $text | Should Match 'f4e0c8c3bed8'
            $text | Should Match 'Specs/Plans/Done/Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15\.md'
            $text | Should Match 'Specs/Plans/WIP/README\.md'
        }
    }

    It 'records that the legacy root plans queue was archived to Done and deleted' {
        Test-Path -LiteralPath (Join-Path $repoRoot 'plans') | Should Be $false

        $index = Get-RSText -Path 'Specs\Plans\Done\AdvisorPlan_Index_2026-06-11.md'
        $index | Should Match 'No remaining action'
        $index | Should Match 'Frozen history: do not resume'
        $index | Should Match 'AdvisorPlan_012-format-autocommit-race_2026-06-11\.md'
        $index | Should Match 'AdvisorPlan_013-undo-file-operations-spike_2026-06-11\.md'
        $index | Should Match 'RETIRED \(archived 2026-08-25; no remaining action; live owner is I9\)'
        $index | Should Match 'routed to WhimFiles G2'
        $index | Should Not Match '(?m)^\s*[-*]\s+\[ \]'

        Get-ChildItem -LiteralPath (Join-Path $repoRoot 'Specs\Plans\Done') -File -Filter 'AdvisorPlan_*.md' | ForEach-Object {
            $archived = Get-Content -LiteralPath $_.FullName -Raw
            $archived | Should Match 'No remaining action'
            $archived | Should Match 'Frozen history: do not resume'
            $archived | Should Not Match '(?m)^\s*[-*]\s+\[ \]'
        }

        $wipIndex = Get-RSText -Path 'Specs\Plans\WIP\README.md'
        $wipIndex | Should Match 'CI_FormatAutocommitRace_2026-08-25\.md'
        $wipIndex | Should Match 'CI_FormatAutocommitRace_2026-08-04\.md'
        $wipIndex | Should Match 'AdvisorPlan_012-format-autocommit-race_2026-06-11\.md'
        $wipIndex | Should Match 'only live owner'
        $wipIndex | Should Match '(?i)do not recreate `plans/`'
        $wipIndex | Should Match 'Operation_FolderView_WarpDrive_RemainingCloseout_2026-08-25\.md'
        $wipIndex | Should Match 'Product_WhimFilesRemainingDecisions_2026-08-25\.md'
        $wipIndex | Should Not Match 'path-preserving shim'
        $wipIndex | Should Not Match 'DxUi_Uia_ContinuationBaton_2026-06-29\.md'
        $wipIndex | Should Not Match 'FolderView_WarpDrive_ContinuationBaton_2026-06-29\.md'

        Test-Path -LiteralPath (Join-Path $repoRoot 'Specs\Plans\WIP\CI_FormatAutocommitRace_2026-08-04.md') | Should Be $false
        Test-Path -LiteralPath (Join-Path $repoRoot 'Specs\Plans\WIP\Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md') | Should Be $false
        Test-Path -LiteralPath (Join-Path $repoRoot 'Specs\Plans\WIP\Product_WhimFilesGapAnalysisAndImprovementPlan_2026-07-08.md') | Should Be $false
        Test-Path -LiteralPath (Join-Path $repoRoot 'Specs\Plans\WIP\DxUi_Uia_ContinuationBaton_2026-06-29.md') | Should Be $false
        Test-Path -LiteralPath (Join-Path $repoRoot 'Specs\Plans\WIP\FolderView_WarpDrive_ContinuationBaton_2026-06-29.md') | Should Be $false

        @(
            'Specs\Plans\Done\CI_FormatAutocommitRace_2026-08-04.md',
            'Specs\Plans\Done\Operation_FolderView_WarpDrive_AnyCircumstancePerformance_2026-06-28.md',
            'Specs\Plans\Done\Product_WhimFilesGapAnalysisAndImprovementPlan_2026-07-08.md',
            'Specs\Plans\Done\DxUi_Uia_ContinuationBaton_2026-06-29.md',
            'Specs\Plans\Done\FolderView_WarpDrive_ContinuationBaton_2026-06-29.md'
        ) | ForEach-Object {
            $archived = Get-Content -LiteralPath (Join-Path $repoRoot $_) -Raw
            $archived | Should Match 'No remaining action'
            $archived | Should Match 'Frozen history: do not resume'
            $archived | Should Not Match '(?m)^\s*[-*]\s+\[ \]'
        }
    }

    It 'resolves reviewed Markdown links and backticked repository paths' {
        $reviewReadmePath = Join-Path $repoRoot 'Specs\Reviews\README.md'
        $reviewReadme = Get-Content -LiteralPath $reviewReadmePath -Raw
        $references = @()
        $references += [regex]::Matches($reviewReadme, '\]\((?<path>[^)#]+\.md)(?:#[^)]+)?\)') |
            ForEach-Object { $_.Groups['path'].Value }
        $references += [regex]::Matches($reviewReadme, '`(?<path>Specs/[A-Za-z0-9_./-]+\.md)`') |
            ForEach-Object { $_.Groups['path'].Value }
        $references.Count | Should BeGreaterThan 0

        foreach ($reference in $references | Select-Object -Unique) {
            $normalized = $reference.Replace('/', [IO.Path]::DirectorySeparatorChar)
            $candidate = if ($normalized.StartsWith("Specs$([IO.Path]::DirectorySeparatorChar)")) {
                Join-Path $repoRoot $normalized
            } else {
                Join-Path (Split-Path -Parent $reviewReadmePath) $normalized
            }
            (Test-Path -LiteralPath $candidate -PathType Leaf) | Should Be $true
        }

        (Test-Path -LiteralPath (Join-Path $repoRoot 'Specs\Plans\WIP\Operation_Curl_Short2xxStagedUploadProof_2026-07-21.md')) |
            Should Be $false
    }

    It 'keeps the completed Observatory ledger internally consistent' {
        $observatory = Get-RSText -Path 'Specs\Plans\Done\Operation_Observatory_WholeRepositoryCodeAuditAndRemediationPlan_2026-07-15.md'
        $agents = Get-RSText -Path 'AGENTS.md'
        $claude = Get-RSText -Path 'CLAUDE.md'

        $observatory | Should Match 'Status \| DONE'
        $observatory | Should Match 'Full closeout run is green or every failure is exactly classified'
        $observatory | Should Not Match 'the canonical CI and Full suites are green'
        $observatory | Should Not Match 'still-open bounded-lifetime'

        $findingRows = @($observatory -split "\r?\n" | Where-Object { $_ -match '^\| OBS-' })
        $findingRows.Count | Should BeGreaterThan 0
        foreach ($row in $findingRows) {
            $row | Should Match '(?i)\b(closed|complete|routed|rejected|covered|historical)\b'
        }

        $agents | Should Not Match '(?m)^- \*\*fmt\*\*:'
        $claude | Should Not Match '(?m)^- \*\*Key Dependencies\*\*:.*\bfmt\b'
    }

    It 'keeps generated Theme Gallery Vite state out of version control' {
        $ignore = Get-RSText -Path 'Specs\Mockups\ThemeGalleryWorkbench\.gitignore'
        $ignore | Should Match '(?m)^\.vite/\r?$'
        $ignore | Should Match '(?m)^dist/\r?$'

        $trackedDist = @(& git -C $repoRoot ls-files -- 'Specs/Mockups/ThemeGalleryWorkbench/dist/**')
        $LASTEXITCODE | Should Be 0
        $trackedDist.Count | Should Be 0
    }

    It 'keeps the vcpkg Dependabot group limited to direct manifest dependencies' {
        $dependabot = Get-RSText -Path '.github\dependabot.yml'
        $manifest = Get-Content -LiteralPath (Join-Path $repoRoot 'vcpkg.json') -Raw | ConvertFrom-Json
        $directDependencies = @($manifest.dependencies | ForEach-Object {
                if ($_ -is [string]) { $_ } else { $_.name }
            })
        $group = [regex]::Match($dependabot, '(?ms)^\s{6}core-stack:\s*\r?\n\s{8}patterns:\s*\r?\n(?<patterns>(?:\s{10}-[^\r\n]+\r?\n)+)')
        $group.Success | Should Be $true
        $patterns = @([regex]::Matches($group.Groups['patterns'].Value, '(?m)^\s+-\s+([^\s#]+)') | ForEach-Object { $_.Groups[1].Value })

        foreach ($pattern in $patterns) {
            ($directDependencies -contains $pattern) | Should Be $true
        }
        ($patterns -contains 'wil') | Should Be $true
        ($patterns -contains 'yyjson') | Should Be $true
    }
}

Describe 'Published File Operations capture provenance' {
    It 'binds all documented count and lifecycle scenes to reviewed native synthetic captures' {
        $images = @(Get-ChildItem -LiteralPath (Join-Path $repoRoot 'docs/res') -Filter 'file-operations-*.png' -File)
        $expectedNames = @('popup','counting','move-totals','file-progress','paused','queued','completed','conflict','wrapped-note','issues') |
            ForEach-Object { 'file-operations-'+$_+'.png' }
        (($images.Name | Sort-Object) -join ',') | Should Be (($expectedNames | Sort-Object) -join ',')
        foreach ($image in $images) {
            $provenancePath = [IO.Path]::ChangeExtension($image.FullName, '.provenance.json')
            (Test-Path -LiteralPath $provenancePath -PathType Leaf) | Should Be $true
            $provenance = Get-Content -LiteralPath $provenancePath -Raw | ConvertFrom-Json
            (Get-FileHash -LiteralPath $image.FullName -Algorithm SHA256).Hash.ToLowerInvariant() | Should Be $provenance.sha256
            $provenance.kind | Should Be 'native-production-renderer-harness-capture'
            $provenance.sourceRunId | Should Match '^\d{8}T\d{6}Z-\d+-[a-f0-9]{32}$'
            $provenance.build.receiptId | Should Match '^[a-f0-9]{64}$'
            $provenance.privacy.dataOrigin | Should Be 'invented harness fixture'
            $provenance.privacy.personalData | Should Be $false
            $provenance.privacy.redactionLabels | Should Be $false
            $provenance.capture.noActivation | Should Be $true
            $provenance.capture.desktopGlobalInput | Should Be $false
            $provenance.capture.wholeFrame | Should Be $true
            $provenance.capture.imageEdits | Should Be 'none'
            $provenance.review.completeFrame | Should Be $true
            foreach ($value in $provenance.syntheticData.PSObject.Properties.Value) {
                $value | Should Not Match '(?i)RedSalamander\.Perf|\.build|[a-z]:[\\/]Users[\\/]'
                if ($value -match '^[a-zA-Z]:[\\/]') {
                    $value | Should Match '^[a-zA-Z]:[\\/](?:Demo|Example|Samples)[\\/]'
                }
            }
        }
    }
}

Describe 'Published product scenario capture provenance' {
    $gallery = Get-Content -LiteralPath (Join-Path $repoRoot 'docs/res/product-scenarios.provenance.json') -Raw | ConvertFrom-Json
    $supplement = Get-Content -LiteralPath (Join-Path $repoRoot 'docs/res/supplemental-scenarios.provenance.json') -Raw | ConvertFrom-Json
    $allCaptures = @($gallery.captures) + @($supplement.captures)

    It 'closes the Terminal, utility-dialog, provider and real-playback illustration gaps' {
        @($supplement.captures).Count | Should Be 21
        foreach ($name in @('change-case','change-attributes','selection-mask','file-filter','make-file-list','create-directory','new-file',
                'archive-pack','archive-unpack','shared-directories','shortcuts-terminal','terminal-embedded','terminal-floating-tabs','terminal-find',
                'provider-ftp','provider-google-drive','provider-onedrive','provider-s3','provider-mtp','viewer-vlc-playback','viewer-vlc-paused')) {
            ($supplement.captures.name -contains $name) | Should Be $true
        }
        $supplement.realAccounts | Should Be $false
        $supplement.terminalShellStarted | Should Be $false
        $supplement.mediaOrigin | Should Be 'generated silent AVI'
    }

    It 'covers Search, Space, actual viewers, plugins and every Preferences category' {
        $gallery.schema | Should Be 'red-salamander.product-documentation-gallery.v1'
        @($gallery.captures).Count | Should Be 47
        @($gallery.captures.name | Select-Object -Unique).Count | Should Be 47
        @($gallery.captures | Where-Object group -eq 'product').Count | Should Be 26
        @($gallery.captures | Where-Object group -eq 'viewer').Count | Should Be 8
        @($gallery.captures | Where-Object group -eq 'space').Count | Should Be 13
        foreach ($name in @('find-files','find-files-empty','plugins','preferences-plugin-mtp','connections','connections-mtp','command-palette',
                'folder-view-brief-detailed','folder-view-extra-detailed','theme-light','theme-dark','theme-high-contrast-app','theme-rainbow',
                'viewer-text','viewer-text-hex','viewer-imgraw','viewer-pe','viewer-web','viewer-json','viewer-markdown','viewer-vlc-unavailable',
                'viewer-space-scanning','viewer-space-inflight','viewer-space-analyst','viewer-space-refreshing','viewer-space-refresh-canceled',
                'viewer-space-refresh-refused','viewer-space-empty','viewer-space-error','viewer-space-explorer','viewer-space-bottom-panel','viewer-space-map-only','viewer-space-excluded','viewer-space-cleanup-review')) {
            ($gallery.captures.name -contains $name) | Should Be $true
        }
        foreach ($page in @('general','panes','viewers','editors','keyboard','mouse','themes','advanced','compare-directories','hot-paths','file-operations','user-menu','monitor')) {
            ($gallery.captures.name -contains ('preferences-'+$page)) | Should Be $true
        }
    }

    It 'binds every whole native frame to its origin receipt and reviewed synthetic data' {
        foreach ($capture in $allCaptures) {
            $path = Join-Path $repoRoot ('docs/res/'+$capture.name+'.png')
            $sidecar = Get-Content -LiteralPath (Join-Path $repoRoot $capture.provenance) -Raw | ConvertFrom-Json
            $digest = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant()
            $digest | Should Be $capture.sha256
            $digest | Should Be $sidecar.sha256
            $sidecar.scenario | Should Be $capture.scenario
            $sidecar.kind | Should Be 'native-production-renderer-harness-capture'
            $sidecar.sourceRunId | Should Match '^\d{8}T\d{6}Z-\d+-[a-f0-9]{32}$'
            (Get-FileHash -LiteralPath (Join-Path $repoRoot $sidecar.source) -Algorithm SHA256).Hash.ToLowerInvariant() | Should Be $digest
            $archive = Split-Path -Parent (Split-Path -Parent (Join-Path $repoRoot $sidecar.source))
            $receiptName = if ($capture.group -eq 'product') { 'app-build-receipt.json' } else { 'viewer-build-receipt.json' }
            $receipt = Get-Content -LiteralPath (Join-Path $archive $receiptName) -Raw | ConvertFrom-Json
            $sidecar.build.receiptId | Should Be $receipt.receipt_id
            $sidecar.build.sourceSnapshotId | Should Be $receipt.source_snapshot_id
            $sidecar.build.executableSha256 | Should Match '^[a-f0-9]{64}$'
            ($receipt.outputs.sha256 -contains $sidecar.build.executableSha256) | Should Be $true
            $sidecar.display.dpi | Should BeGreaterThan 0
            $sidecar.display.pixelWidth | Should BeGreaterThan 0
            $sidecar.display.pixelHeight | Should BeGreaterThan 0
            $sidecar.capture.noActivation | Should Be $true
            $sidecar.capture.desktopGlobalInput | Should Be $false
            $sidecar.capture.wholeFrame | Should Be $true
            $sidecar.capture.imageEdits | Should Be 'none'
            $sidecar.review.completeFrame | Should Be $true
            $sidecar.privacy.dataOrigin | Should Be 'invented harness fixture'
            $sidecar.privacy.personalData | Should Be $false
            $sidecar.privacy.redactionLabels | Should Be $false
            $sidecar.privacy.realAccounts | Should Be $false
            foreach ($root in $sidecar.syntheticData.roots) {
                $root | Should Match '^(?:[A-Z]:[\\/]Demo[\\/]|/?Demo/|/@conn:Demo (?:Drive|OneDrive)/|/Fake Phone/Internal Storage/|/demo-library/|//anonymous@127\.0\.0\.1:[0-9]+/)'
                $root | Should Not Match '(?i)RedSalamander\.Perf|\.build|[\\/]Users[\\/]'
            }
            if ($capture.group -eq 'space') {
                $native = Get-Content -LiteralPath (Join-Path $repoRoot $sidecar.nativeCaptureProvenance) -Raw | ConvertFrom-Json
                $native.valid | Should Be $true
                $native.foreground_preserved | Should Be $true
            }
        }
    }

    It 'links every published scenario from a guide or its scenario table' {
        $guides = (Get-ChildItem -LiteralPath (Join-Path $repoRoot 'docs') -File -Filter '*.md' |
            ForEach-Object { [IO.File]::ReadAllText($_.FullName) }) -join "`n"
        foreach ($capture in $allCaptures) {
            $guides | Should Match ([regex]::Escape('res/'+$capture.name+'.png'))
        }
        $tour = Get-RSText -Path 'docs/VisualTour.md'
        $tour | Should Match 'live\s+accounts/devices'
        $tour | Should Match 'FileOperations\.md'
    }
}
