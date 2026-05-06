Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$helperScript = Join-Path $repoRoot 'Installer\winget\WingetValidation.ps1'

Describe 'Winget validation helper' {
    BeforeAll {
        . $helperScript
    }

    It 'allows only the known winget 1.11 schema-header warning for 1.12 manifests' {
        $output = @(
            'Manifest validation succeeded with warnings.'
            'Manifest Warning: The schema header URL does not match the expected pattern. Value: # yaml-language-server: $schema=https://aka.ms/winget-manifest.installer.1.12.0.schema.json Line: 1, Column: 25 File: RedSalamanders.RedSalamander.installer.yaml'
            'Manifest Warning: The schema header URL does not match the expected pattern. Value: # yaml-language-server: $schema=https://aka.ms/winget-manifest.defaultLocale.1.12.0.schema.json Line: 1, Column: 25 File: RedSalamanders.RedSalamander.locale.en-US.yaml'
            'Manifest Warning: The schema header URL does not match the expected pattern. Value: # yaml-language-server: $schema=https://aka.ms/winget-manifest.version.1.12.0.schema.json Line: 1, Column: 25 File: RedSalamanders.RedSalamander.yaml'
        )

        Test-RSLegacyWingetSchemaHeaderWarning -WingetVersion 'v1.11.510' -OutputLines $output | Should Be $true
    }

    It 'rejects unrelated warnings' {
        $output = @(
            'Manifest validation succeeded with warnings.'
            'Manifest Warning: Field usage requires verified publishers. [Icons]'
        )

        Test-RSLegacyWingetSchemaHeaderWarning -WingetVersion 'v1.11.510' -OutputLines $output | Should Be $false
    }

    It 'does not suppress warnings from current winget validators' {
        $output = @(
            'Manifest validation succeeded with warnings.'
            'Manifest Warning: The schema header URL does not match the expected pattern. Value: # yaml-language-server: $schema=https://aka.ms/winget-manifest.installer.1.12.0.schema.json'
        )

        Test-RSLegacyWingetSchemaHeaderWarning -WingetVersion 'v1.28.240' -OutputLines $output | Should Be $false
    }
}
