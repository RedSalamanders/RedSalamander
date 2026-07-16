Describe 'Theme distribution contracts' {
    BeforeAll {
        $repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
        $themesRoot = Join-Path $repoRoot 'Specs\Themes'
        $thirdPartyThemes = @(
            @{ File = 'Dracula.theme.json5'; License = 'Dracula.LICENSE.txt'; Source = 'github.com/dracula/dracula-theme' }
            @{ File = 'CatppuccinLatte.theme.json5'; License = 'Catppuccin.LICENSE.txt'; Source = 'github.com/catppuccin/palette' }
            @{ File = 'CatppuccinFrappe.theme.json5'; License = 'Catppuccin.LICENSE.txt'; Source = 'github.com/catppuccin/palette' }
            @{ File = 'CatppuccinMocha.theme.json5'; License = 'Catppuccin.LICENSE.txt'; Source = 'github.com/catppuccin/palette' }
        )
    }

    It 'keeps pinned source and exact license references beside every third-party theme' {
        Test-Path (Join-Path $themesRoot 'THIRD-PARTY-NOTICES.md') | Should Be $true
        foreach ($theme in $thirdPartyThemes) {
            $themePath = Join-Path $themesRoot $theme.File
            Test-Path $themePath | Should Be $true
            Test-Path (Join-Path $themesRoot ('Licenses\' + $theme.License)) | Should Be $true
            $text = Get-Content -LiteralPath $themePath -Raw
            $text | Should Match 'SPDX-License-Identifier: MIT'
            $text | Should Match ([regex]::Escape($theme.Source))
            $text | Should Match ([regex]::Escape('Licenses/' + $theme.License))
        }
    }

    It 'copies themes, notices, and exact licenses from both application projects' {
        foreach ($project in @('RedSalamander\RedSalamander.vcxproj', 'RedSalamanderMonitor\RedSalamanderMonitor.vcxproj')) {
            $text = Get-Content -LiteralPath (Join-Path $repoRoot $project) -Raw
            $text | Should Match '<ThemeFiles Include="\$\(ThemeSourceDir\)\\\*\.theme\.json5" />'
            $text | Should Match '<ThemeNoticeFiles Include="\$\(ThemeSourceDir\)\\THIRD-PARTY-NOTICES\.md" />'
            $text | Should Match '<ThemeLicenseFiles Include="\$\(ThemeSourceDir\)\\Licenses\\\*\.LICENSE\.txt" />'
            $text | Should Match 'SourceFiles="@\(ThemeNoticeFiles\)"'
            $text | Should Match 'SourceFiles="@\(ThemeLicenseFiles\)"'
        }
    }

    It 'keeps migrated semantic color blocks authored as references or functions' {
        Get-ChildItem -LiteralPath $themesRoot -Filter '*.theme.json5' | ForEach-Object {
            $text = Get-Content -LiteralPath $_.FullName -Raw
            $colors = [regex]::Match($text, '(?s)"colors"\s*:\s*\{(?<body>.*?)\}\s*,?\s*\}').Groups['body'].Value
            $colors | Should Not BeNullOrEmpty
            $colors | Should Not Match '"[^"\r\n]+"\s*:\s*"#[0-9A-Fa-f]{6,8}"'
        }
    }

    It 'keeps every shipped Catppuccin flavor on the approved Lavender Blue Overlay2 mapping' {
        foreach ($file in @('CatppuccinLatte.theme.json5', 'CatppuccinFrappe.theme.json5', 'CatppuccinMocha.theme.json5')) {
            $text = Get-Content -LiteralPath (Join-Path $themesRoot $file) -Raw
            $text | Should Match '"selection"\s*:\s*"blend\(palette\.base,palette\.overlay2,26%\)"'
            $text | Should Match '"app\.accent"\s*:\s*"ref\(palette\.lavender\)"'
            $text | Should Match '"navigation\.accent"\s*:\s*"ref\(palette\.blue\)"'
            $text | Should Match '"folderView\.focusBorder"\s*:\s*"ref\(palette\.lavender\)"'
            $text | Should Not Match '"app\.accent"\s*:\s*"ref\(palette\.mauve\)"'
        }
    }

    It 'keeps Preferences duplicate and reset operations lossless for authored palettes' {
        $text = Get-Content -LiteralPath (Join-Path $repoRoot 'RedSalamander\Preferences.Themes.cpp') -Raw
        $text | Should Match 'def\.palette\s*=\s*sourceDef->palette;'
        $text | Should Match 'def\.colors\s*=\s*sourceDef->colors;'
        $text | Should Match 'def->palette\.clear\(\);\s*\r?\n\s*def->colors\.clear\(\);'
    }

    It 'routes startup hot reload and Preferences preview through one app-theme resolver' {
        $appTheme = Get-Content -LiteralPath (Join-Path $repoRoot 'RedSalamander\AppTheme.cpp') -Raw
        $consumers = @(
            'RedSalamander\RedSalamander.cpp',
            'RedSalamander\SettingsHotReload.cpp',
            'RedSalamander\Preferences.Themes.cpp'
        )
        $consumerText = ($consumers | ForEach-Object { Get-Content -LiteralPath (Join-Path $repoRoot $_) -Raw }) -join "`n"

        $appTheme | Should Match 'AppThemeSelectionResolution\s+ResolveAppThemeSelection\s*\('
        $appTheme | Should Match 'void\s+ApplyAppThemeColorOverrides\s*\('
        $consumerText | Should Match 'ResolveAppThemeSelection\s*\('
        $consumerText | Should Not Match 'void\s+(ApplyThemeOverrides|ApplyAppThemeOverrides|ApplyDialogThemeOverrides)\s*\('
        $consumerText | Should Not Match 'ThemeMode\s+ThemeModeFromThemeId\s*\('
    }

    It 'keeps RedConfigure preview editing on real runtime semantic keys' {
        $paths = @(
            'RedConfigure\Themes\ThemePreviewModel.cpp',
            'RedConfigure\RedConfigureRoot.cpp',
            'RedConfigure\RedConfigureThemeExampleControl.h'
        )
        $text = ($paths | ForEach-Object { Get-Content -LiteralPath (Join-Path $repoRoot $_) -Raw }) -join "`n"
        foreach ($legacyAlias in @(
                'menu.selectionBackground',
                'folderView.itemForeground',
                'folderView.itemForegroundSelected',
                'folderView.warningForeground',
                'dialog.background',
                'progress.fill',
                'diff.addedBackground')) {
            $text | Should Not Match ('L"' + [regex]::Escape($legacyAlias) + '"')
        }
        $text | Should Match 'L"menu\.selectionBg"'
        $text | Should Match 'L"folderView\.textNormal"'
        $text | Should Match 'L"viewer\.diff\.addedBackground"'
    }

    It 'keeps app palette profiles named and removes exact pass-through adapters' {
        $findSource = Get-Content -LiteralPath (Join-Path $repoRoot 'RedSalamander\FindFilesWindow.cpp') -Raw
        $issuesSource = Get-Content -LiteralPath (Join-Path $repoRoot 'RedSalamander\FolderWindow.FileOperations.IssuesPane.cpp') -Raw
        $credentialSource = Get-Content -LiteralPath (Join-Path $repoRoot 'RedSalamander\ConnectionCredentialPromptDialog.cpp') -Raw
        $preferencesHeader = Get-Content -LiteralPath (Join-Path $repoRoot 'RedSalamander\Preferences.Internal.h') -Raw
        $appSources = Get-ChildItem -LiteralPath (Join-Path $repoRoot 'RedSalamander') -File -Include '*.cpp', '*.h' -Recurse |
            ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }
        $appSourceText = $appSources -join "`n"

        $findSource | Should Match 'MakeFolderContentDxPalette\(_theme\)'
        $issuesSource | Should Match 'MakeFolderContentDxPalette\(_theme\)'
        $credentialSource | Should Match 'ThemePalette\s+MakeCredentialPromptDxPalette\('
        $preferencesHeader | Should Not Match 'ThemePalette\s+MakeDxPalette\('
        $appSourceText | Should Not Match 'PrefsUi::MakeDxPalette'
        $appSourceText | Should Not Match 'ThemePalette\s+MakeDxPalette\('
    }

    It 'keeps the public control gallery complete for built-ins, both Rainbow bases, and every shipped theme' {
        $expectedFiles = @(
            'theme-controls-light.png',
            'theme-controls-dark.png',
            'theme-controls-rainbow-light.png',
            'theme-controls-rainbow-dark.png',
            'theme-controls-high-contrast.png'
        )
        Get-ChildItem -LiteralPath $themesRoot -Filter '*.theme.json5' | ForEach-Object {
            $text = Get-Content -LiteralPath $_.FullName -Raw
            $id = [regex]::Match($text, '"id"\s*:\s*"user/(?<slug>[a-z0-9-]+)"').Groups['slug'].Value
            $id | Should Not BeNullOrEmpty
            $expectedFiles += 'theme-controls-' + $id + '.png'
        }

        $galleryRoot = Join-Path $repoRoot 'docs\res'
        $actualFiles = @(Get-ChildItem -LiteralPath $galleryRoot -Filter 'theme-controls-*.png' -File | Select-Object -ExpandProperty Name | Sort-Object)
        $expectedFiles = @($expectedFiles | Sort-Object)
        (Compare-Object -ReferenceObject $expectedFiles -DifferenceObject $actualFiles) | Should BeNullOrEmpty

        $themesDoc = Get-Content -LiteralPath (Join-Path $repoRoot 'docs\Themes.md') -Raw
        foreach ($file in $expectedFiles) {
            $path = Join-Path $galleryRoot $file
            Test-Path -LiteralPath $path | Should Be $true
            $bytes = [System.IO.File]::ReadAllBytes($path)
            $bytes.Length | Should BeGreaterThan 10000
            [System.BitConverter]::ToString($bytes[0..7]) | Should Be '89-50-4E-47-0D-0A-1A-0A'
            $themesDoc | Should Match ([regex]::Escape('res/' + $file))
        }
    }
}
