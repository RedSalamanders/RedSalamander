<#
.SYNOPSIS
    Loads the compatibility TestRunPlan command surface.

.DESCRIPTION
    Imports the canonical modular test-plan implementation into the caller's
    global session. This path is retained for sealed historical Pester files.

.EXAMPLE
    . .\Tools\TestRunPlan.ps1

    Loads the established Get-RSTestRunPlan and related helper functions.

.OUTPUTS
    None. The loader publishes module commands into the PowerShell session.

.NOTES
    Prerequisite: Tools/Modules/Testing/TestRunPlan.psm1 and its domain modules.
    Side effects: loads functions globally. Current callers should import the
    canonical module directly; frozen Gate-0/Round-4 tests retain this path.
#>

Import-Module (Join-Path $PSScriptRoot 'Modules\Testing\TestRunPlan.psm1') -Force -Global -ErrorAction Stop
