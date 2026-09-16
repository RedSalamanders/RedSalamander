# Product-owned DxUi pin update. The library's ConsumerUpdate module stays read-only;
# this module owns RedSalamander's mutable lock and Full-suite qualification command.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Read-RSDxUiUpdateLock {
    param([Parameter(Mandatory)][string]$LockFile)

    $pin = Get-Content -LiteralPath $LockFile -Raw | ConvertFrom-Json
    if ($pin.repository -cne 'https://github.com/RedSalamanders/DxUi' -or
        $pin.commit -cnotmatch '^[0-9a-f]{40}$' -or $pin.apiRevision -ne 2 -or
        @($pin.targets).Count -ne 1 -or $pin.targets[0] -cne 'DxUi') {
        throw "DxUi lock is not the canonical API-revision-2 DxUi target: $LockFile"
    }
    return $pin
}

function New-RSDxUiUpdateRequest {
    $headers = @{ Accept = 'application/vnd.github+json'; 'X-GitHub-Api-Version' = '2022-11-28' }
    $token = if ($env:GH_TOKEN) { $env:GH_TOKEN } else { $env:GITHUB_TOKEN }
    if ($token) { $headers.Authorization = "Bearer $token" }
    return {
        param([Parameter(Mandatory)][string]$Route)
        Invoke-RestMethod -Uri "https://api.github.com/repos/RedSalamanders/DxUi/$Route" `
            -Headers $headers -TimeoutSec 10 -ErrorAction Stop
    }.GetNewClosure()
}

function Get-RSDxUiValidatedMainCandidate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$LockFile,
        [scriptblock]$Request
    )

    $pin = Read-RSDxUiUpdateLock -LockFile $LockFile
    if (-not $Request) { $Request = New-RSDxUiUpdateRequest }
    try {
        $head = & $Request 'commits/main'
        $candidate = [string]$head.sha
        if ($candidate -cnotmatch '^[0-9a-f]{40}$') { throw 'DxUi main did not return an exact commit.' }
        if ($candidate -ceq $pin.commit) {
            return [pscustomobject]@{ LockFile=$LockFile; Current=$pin.commit; Candidate=$candidate; IsCurrent=$true }
        }

        $comparison = & $Request "compare/$($pin.commit)...$candidate"
        if ($comparison.status -eq 'behind') { throw 'The product DxUi pin is ahead of DxUi main; review the lock manually.' }
        if ($comparison.status -ne 'ahead') { throw 'The product DxUi pin and DxUi main diverge; review the lock manually.' }

        $runs = @((& $Request "actions/workflows/ci.yml/runs?head_sha=$candidate&branch=main&event=push&per_page=1").workflow_runs)
        if ($runs.Count -ne 1 -or $runs[0].head_sha -cne $candidate -or
            $runs[0].status -ne 'completed' -or $runs[0].conclusion -ne 'success') {
            throw "DxUi main $($candidate.Substring(0, 12)) has no successful completed validation; the lock remains unchanged."
        }
        return [pscustomobject]@{ LockFile=$LockFile; Current=$pin.commit; Candidate=$candidate; IsCurrent=$false }
    }
    catch [System.Exception] {
        throw "Cannot select a validated DxUi main candidate: $($_.Exception.Message)"
    }
}

function Set-RSDxUiLockCommit {
    param(
        [Parameter(Mandatory)][string]$LockFile,
        [Parameter(Mandatory)][ValidatePattern('^[0-9a-f]{40}$')][string]$Commit
    )

    $pin = Read-RSDxUiUpdateLock -LockFile $LockFile
    $pin.commit = $Commit
    $json = ($pin | ConvertTo-Json -Depth 5) + [Environment]::NewLine
    $directory = Split-Path -Parent $LockFile
    $temporary = Join-Path $directory ('.' + [IO.Path]::GetFileName($LockFile) + ".$PID.$([guid]::NewGuid().ToString('N')).tmp")
    try {
        [IO.File]::WriteAllText($temporary, $json, [Text.UTF8Encoding]::new($false))
        [IO.File]::Move($temporary, $LockFile, $true)
    }
    finally {
        if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force }
    }
}

function Invoke-RSDxUiUpdate {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$RepoRoot,
        [switch]$UpdateOnly,
        [scriptblock]$Request,
        [scriptblock]$ValidationAction
    )

    $lockFile = Join-Path $RepoRoot 'Dependencies/DxUi.lock.json'
    $candidate = Get-RSDxUiValidatedMainCandidate -LockFile $lockFile -Request $Request
    if ($candidate.IsCurrent) {
        Write-Host "DxUi is already pinned to validated main $($candidate.Current.Substring(0, 12))." -ForegroundColor Green
        return [pscustomobject]@{ Updated=$false; Validated=$false; Candidate=$candidate.Candidate }
    }

    if (-not $PSCmdlet.ShouldProcess($lockFile, "Update DxUi pin $($candidate.Current.Substring(0, 12)) -> $($candidate.Candidate.Substring(0, 12))")) {
        return [pscustomobject]@{ Updated=$false; Validated=$false; Candidate=$candidate.Candidate; WhatIf=$true }
    }

    if (-not $UpdateOnly -and -not $ValidationAction) {
        throw 'A product validation action is required unless -UpdateOnly is selected.'
    }
    Set-RSDxUiLockCommit -LockFile $lockFile -Commit $candidate.Candidate
    Write-Host "DxUi pin updated: $($candidate.Current.Substring(0, 12)) -> $($candidate.Candidate.Substring(0, 12))." -ForegroundColor Yellow

    if ($UpdateOnly) {
        Write-Host 'Product validation skipped by -UpdateOnly; retain external validation evidence with this branch.' -ForegroundColor Yellow
        return [pscustomobject]@{ Updated=$true; Validated=$false; Candidate=$candidate.Candidate }
    }

    Write-Host 'Running RedSalamander Full validation for the updated DxUi pin...' -ForegroundColor Cyan
    & $ValidationAction
    Write-Host 'DxUi pin update and RedSalamander validation passed.' -ForegroundColor Green
    return [pscustomobject]@{ Updated=$true; Validated=$true; Candidate=$candidate.Candidate }
}

Export-ModuleMember -Function Get-RSDxUiValidatedMainCandidate, Invoke-RSDxUiUpdate
