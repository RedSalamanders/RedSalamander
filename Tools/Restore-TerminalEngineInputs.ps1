<#
.SYNOPSIS
Restores or verifies inputs declared by an explicit Terminal-engine lock.

.DESCRIPTION
Compatibility entrypoint for generic Terminal qualification and upgrade work. It verifies the lock and materializes each declared input in the controlled Terminal cache, or performs cache-only verification in offline mode.

.PARAMETER LockFile
Specifies the explicit Terminal-engine evaluation or candidate lock.

.PARAMETER Offline
Forbids downloads and verifies that all locked inputs already exist with exact identities.

.PARAMETER PassThru
Returns one status record per declared input.

.OUTPUTS
With -PassThru, PSCustomObject status records for the locked inputs. Otherwise no pipeline output.

.NOTES
Prerequisites: a valid explicit lock; online restore also requires network access. Side effects: may download exact locked inputs into the Terminal build cache. Exit is nonzero on lock, download, cache, or identity failure. Primary consumers: maintainer-operated candidate qualification and upgrade diagnostics; this is not needed for the normal Ghostty product build.

.EXAMPLE
.\Tools\Restore-TerminalEngineInputs.ps1 -LockFile .\.build\candidate-lock.json -Offline -PassThru
#>
[CmdletBinding()]
param(
    [string]$LockFile = 'External\terminal-engine.lock.json',

    [switch]$Offline,

    [switch]$PassThru
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$lifecycleModule = Join-Path $repoRoot 'Tools\TerminalEngine\TerminalEngineLifecycle.psm1'
Import-Module $lifecycleModule -Force -ErrorAction Stop

$lock = Read-RSTerminalEngineLock -LockFile $LockFile
$records = @(Restore-RSTerminalEngineInputs -Lock $lock -Offline:$Offline)

if ($PassThru) {
    $records
}
else {
    $mode = if ($Offline) { 'offline verification' } else { 'restore' }
    Write-Host "Terminal engine input $mode passed: $($lock.engine.candidateId) $($lock.engine.pin)"
    foreach ($record in $records) {
        Write-Host "  $($record.id): $($record.status)"
    }
}
