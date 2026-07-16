<#
.SYNOPSIS
    Lists or removes legacy RedSalamander test scratch roots that predate REDSALAMANDER_TEST_ROOT.

.DESCRIPTION
    Dry-run is the default. Pass -Apply to remove resolved targets. Each removal still goes through
    ShouldProcess, so callers can combine -Apply with -WhatIf or -Confirm.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [switch]$Apply,

    [string]$LocalAppDataRoot = $env:LOCALAPPDATA,

    [string]$TempRoot = [System.IO.Path]::GetTempPath(),

    [string[]]$DriveRoots = @()
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $scriptRoot 'TestRunPlan.ps1')

if (@($DriveRoots).Count -eq 0) {
    $DriveRoots = @(Get-RSFixedDriveRoots)
}

$plan = @(Get-RSTestSandboxLegacyCleanupPlan `
        -LocalAppDataRoot $LocalAppDataRoot `
        -TempRoot $TempRoot `
        -DriveRoots $DriveRoots)
$targets = @(Resolve-RSTestSandboxCleanupTargets -Plan $plan)

if (-not $Apply) {
    Write-Host "Dry run: pass -Apply to remove the resolved legacy test sandbox targets." -ForegroundColor Yellow
    $targets
    return
}

foreach ($target in $targets) {
    if ($PSCmdlet.ShouldProcess($target.Path, "Remove legacy RedSalamander test artifact ($($target.Category))")) {
        try {
            Remove-Item -LiteralPath $target.Path -Recurse -Force -ErrorAction Stop
            $target | Add-Member -NotePropertyName Status -NotePropertyValue 'Removed' -Force
            $target | Add-Member -NotePropertyName Error -NotePropertyValue $null -Force
        } catch {
            $message = $_.Exception.Message
            Write-Warning "Failed to remove legacy RedSalamander test artifact '$($target.Path)': $message"
            $target | Add-Member -NotePropertyName Status -NotePropertyValue 'Failed' -Force
            $target | Add-Member -NotePropertyName Error -NotePropertyValue $message -Force
        }
    } else {
        $target | Add-Member -NotePropertyName Status -NotePropertyValue 'Skipped' -Force
        $target | Add-Member -NotePropertyName Error -NotePropertyValue $null -Force
    }
}

$targets
