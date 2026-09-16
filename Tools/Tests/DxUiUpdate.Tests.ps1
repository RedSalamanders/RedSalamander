$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'Tools/Modules/Build/DxUiUpdate.psm1') -Force

function New-RSDxUiUpdateFixture {
    param([string]$Root, [string]$Commit)
    [void](New-Item -ItemType Directory -Path (Join-Path $Root 'Dependencies') -Force)
    [ordered]@{ repository='https://github.com/RedSalamanders/DxUi'; commit=$Commit; apiRevision=2; targets=@('DxUi') } |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $Root 'Dependencies/DxUi.lock.json') -NoNewline
}

function New-RSDxUiUpdateRequest {
    param([string]$Current, [string]$Candidate, [string]$Conclusion = 'success')
    return {
        param([string]$Route)
        if ($Route -eq 'commits/main') { return [pscustomobject]@{ sha=$Candidate } }
        if ($Route -eq "compare/$Current...$Candidate") { return [pscustomobject]@{ status='ahead' } }
        if ($Route -like 'actions/workflows/ci.yml/runs*') {
            return [pscustomobject]@{ workflow_runs=@([pscustomobject]@{ head_sha=$Candidate; status='completed'; conclusion=$Conclusion }) }
        }
        throw "Unexpected DxUi update route: $Route"
    }.GetNewClosure()
}

Describe 'RedSalamander DxUi updater' {
    It 'updates a successfully validated candidate and runs the Full validation action by default' {
        $current = 'a' * 40; $candidate = 'b' * 40; $root = Join-Path $TestDrive 'default'
        New-RSDxUiUpdateFixture -Root $root -Commit $current
        $state = [pscustomobject]@{ Ran=$false }
        $result = Invoke-RSDxUiUpdate -RepoRoot $root -Request (New-RSDxUiUpdateRequest -Current $current -Candidate $candidate) -ValidationAction { $state.Ran=$true }
        $result.Updated | Should Be $true; $result.Validated | Should Be $true; $state.Ran | Should Be $true
        ((Get-Content -Raw -LiteralPath (Join-Path $root 'Dependencies/DxUi.lock.json') | ConvertFrom-Json).commit) | Should Be $candidate
    }

    It 'updates without product validation only when explicitly requested' {
        $current = 'c' * 40; $candidate = 'd' * 40; $root = Join-Path $TestDrive 'update-only'
        New-RSDxUiUpdateFixture -Root $root -Commit $current
        $state = [pscustomobject]@{ Ran=$false }
        $result = Invoke-RSDxUiUpdate -RepoRoot $root -UpdateOnly -Request (New-RSDxUiUpdateRequest -Current $current -Candidate $candidate) -ValidationAction { $state.Ran=$true }
        $result.Updated | Should Be $true; $result.Validated | Should Be $false; $state.Ran | Should Be $false
    }

    It 'rejects an upstream candidate without successful completed CI and preserves the lock' {
        $current = 'e' * 40; $candidate = 'f' * 40; $root = Join-Path $TestDrive 'rejected'
        New-RSDxUiUpdateFixture -Root $root -Commit $current
        $failure = $null
        try { Invoke-RSDxUiUpdate -RepoRoot $root -UpdateOnly -Request (New-RSDxUiUpdateRequest -Current $current -Candidate $candidate -Conclusion 'failure') } catch { $failure = $_.Exception.Message }
        $failure | Should Match '^Cannot select a validated DxUi main candidate:'
        ((Get-Content -Raw -LiteralPath (Join-Path $root 'Dependencies/DxUi.lock.json') | ConvertFrom-Json).commit) | Should Be $current
    }
}
