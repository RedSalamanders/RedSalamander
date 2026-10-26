Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Describe 'Shared Changes workflow' {
    BeforeAll {
        $script:ChangesRepository = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
        Import-Module (Join-Path $script:ChangesRepository 'Tools/Modules/Testing/TestSandbox.psm1')
    }

    It 'validates this repository through the common command' {
        & (Join-Path $script:ChangesRepository 'Tools/Test-Changelog.ps1')
    }

    It 'passes the shared behavior fixtures in the existing marked sandbox' {
        $scratch = New-RSTestSandboxScratchDirectory -RepoRoot $script:ChangesRepository -Harness 'ChangesWorkflow' -Case 'Behavior'
        & (Join-Path $script:ChangesRepository 'Tools/Test-ChangelogWorkflow.ps1') -ScratchRoot $scratch
    }

    It 'rejects a reparse root without touching its owned target' {
        $scratch = New-RSTestSandboxScratchDirectory -RepoRoot $script:ChangesRepository -Harness 'ChangesWorkflow' -Case 'Reparse'
        $target = Join-Path $scratch 'real'
        [void][IO.Directory]::CreateDirectory((Join-Path $target 'Changes'))
        [IO.File]::WriteAllText((Join-Path $target 'Changes/README.md'), '# Fragments')
        [IO.File]::WriteAllText((Join-Path $target 'Changes/2026-10-21-keep.md'), '- Keep this input.')
        [IO.File]::WriteAllText((Join-Path $target 'CHANGELOG.md'), "# Changelog`n`n## Unreleased`n")
        $before = Get-FileHash -LiteralPath (Join-Path $target 'CHANGELOG.md')
        $alias = Join-Path $scratch 'alias'
        $null = New-RSTestSandboxLink -Path $alias -Target $target -ItemType Junction
        try {
            $message = ''
            try { & (Join-Path $script:ChangesRepository 'Tools/Fold-Changelog.ps1') -Root $alias }
            catch { $message = $_.Exception.Message }
            $message | Should Match 'reparse points'
            (Get-FileHash -LiteralPath (Join-Path $target 'CHANGELOG.md')).Hash | Should Be $before.Hash
            Test-Path -LiteralPath (Join-Path $target 'Changes/2026-10-21-keep.md') | Should Be $true
        } finally { Remove-Item -LiteralPath $alias -Force }
    }
}
