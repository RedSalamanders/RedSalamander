Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Drift-prevention contracts for the documentation-completeness work
# (Operation_Codex_DocumentationCompleteness). These guard the two regression
# classes the documentation audit found: keyboard shortcut/spec drift away from
# the code of record, and settings keys read/written in code but missing from the
# schema. See Specs/Plans/Done/Operation_Codex_DocumentationCompleteness_2026-06-18.md.

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

function Get-RSText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return Get-Content -LiteralPath (Join-Path $repoRoot $Path) -Raw
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
    It 'tracks the split popup cases and completed remediation under Done' {
        $coverage = Get-RSText -Path 'Specs\Testing\Testing_TestCoverage.md'
        $fileOperations = Get-RSText -Path 'Specs\FileSystem\FileSystem_FileOperations.md'
        $uxPlan = Get-RSText -Path 'Specs\Plans\WIP\UI_FileOperationsPopupUxRefinementPlan_2026-07-07.md'
        $splitCases = @(
            'cmd_pane_fileops_popup_progress_contracts',
            'cmd_pane_fileops_conflict_metadata_uses_single_provider_roundtrip',
            'cmd_pane_fileops_conflict_prompt_metadata_and_actions',
            'cmd_pane_fileops_popup_presentation_settings_and_taskbar',
            'cmd_pane_fileops_completed_group_and_navigation'
        )

        foreach ($caseName in $splitCases) {
            $coverage | Should Match ([Regex]::Escape($caseName))
            $fileOperations | Should Match ([Regex]::Escape($caseName))
        }
        $coverage | Should Not Match 'cmd_pane_fileops_conflict_prompt_compacts_actions'
        $fileOperations | Should Not Match 'cmd_pane_fileops_conflict_prompt_compacts_actions'
        $uxPlan | Should Match '\[UI_FileOperationsPopupCodeReviewRemediation_2026-07-10\.md\]\(\.\./Done/UI_FileOperationsPopupCodeReviewRemediation_2026-07-10\.md\)'
        $uxPlan | Should Match 'review-remediation evidence gate is complete'
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

    It 'marks the legacy root plans index historical and routes its remaining decisions' {
        $plans = Get-RSText -Path 'plans\README.md'

        $plans | Should Match 'Historical snapshot — not a live work queue'
        $plans | Should Match '012-format-autocommit-race[\s\S]{0,300}separate live owner indexed by `Specs/Plans/WIP/README\.md`'
        $plans | Should Match '013-undo-file-operations-spike[\s\S]{0,300}routed to WhimFiles G2'

        $wipIndex = Get-RSText -Path 'Specs\Plans\WIP\README.md'
        $wipIndex | Should Match 'plans/012-format-autocommit-race\.md'
        $wipIndex | Should Match 'Observatory deliberately coordinated but did not absorb'
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
