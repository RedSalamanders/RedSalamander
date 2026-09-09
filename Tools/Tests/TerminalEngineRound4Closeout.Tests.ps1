$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$closeoutScript = Join-Path $repoRoot (
    'Tools\Test-TerminalEngineRound4Closeout.ps1')
$frozenWorkflow = Join-Path $repoRoot (
    '.github\workflows\terminal-engine-gate0.yml')
$closeoutWorkflow = Join-Path $repoRoot (
    '.github\workflows\terminal-engine-round4-closeout.yml')

Describe 'TerminalEngine Round 4 no-winner closeout' {
    It 'keeps executable collection frozen and defines a disabled-workflow closeout path' {
        (Get-FileHash -Algorithm SHA256 -LiteralPath $frozenWorkflow).
            Hash.ToLowerInvariant() |
            Should Be (
                'e6f15b6fef9c0d367e12dc7f43cb769106dc982934658f501120728599f32d1c')

        $workflowText = [IO.File]::ReadAllText($closeoutWorkflow)
        $workflowText | Should Match '"on":\s*\r?\n\s+workflow_dispatch:'
        $workflowText | Should Not Match '(?m)^\s+push:'
        $workflowText | Should Match 'disabled_manually'
        $workflowText | Should Match 'forbidden manual-dispatch history'
        $workflowText | Should Match 'active.*run'
        $workflowText | Should Match 'overlapping executable run'
        $workflowText | Should Match 'Test-TerminalEngineRound4Closeout\.ps1'
        ([regex]::Matches(
                $workflowText,
                'Test-TerminalEngineRound4Closeout\.ps1')).Count |
            Should Be 2
        $workflowText | Should Match 'Run-AllTests\.ps1 -Suite Full'
        $workflowText | Should Match '\$tests\.Count -ne 20'
        $workflowText | Should Match '\$passed -ne 239'
        $workflowText | Should Not Match 'upload-artifact'
        $workflowText | Should Not Match 'Evaluate-TerminalEngines\.ps1.+Finalize'
    }

    It 'rejects the superseding product tree before historical closeout effects' {
        $evidenceRoot = Join-Path $TestDrive 'round4-evidence'
        Test-Path -LiteralPath (Join-Path $repoRoot 'Plugins\Terminal') |
            Should Be $true
        [IO.File]::ReadAllText(
            (Join-Path $repoRoot 'RuntimeDependencies.props')) |
            Should Match 'ghostty-vt-x64'

        $caught = $null
        try {
            & $closeoutScript `
                -RepoRoot $repoRoot `
                -EvidenceRoot $evidenceRoot
        }
        catch [System.IO.InvalidDataException] {
            $caught = $_.Exception
        }

        $caught | Should Not BeNullOrEmpty
        $caught.Message | Should Be (
            "Round-4 closeout identity 'RuntimeDependencies' changed.")
        (Test-Path -LiteralPath $evidenceRoot) | Should Be $false
    }
}
