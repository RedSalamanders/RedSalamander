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
