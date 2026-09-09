<#
.SYNOPSIS
    Loads the compatibility sanitized-process command surface.

.DESCRIPTION
    Imports the canonical sanitized-environment module into the caller's global
    session. This path remains stable for sealed Terminal isolation tooling.

.EXAMPLE
    . .\Tools\SanitizedEnvironment.ps1

    Loads New-RSProcessStartInfo, Invoke-RSProcess, and related commands.

.OUTPUTS
    None. The loader publishes module commands into the PowerShell session.

.NOTES
    Prerequisite: Tools/Modules/Build/SanitizedEnvironment.psm1.
    Side effects: loads functions globally. Current callers should import the
    canonical module directly; sealed historical tooling retains this path.
#>

Import-Module (Join-Path $PSScriptRoot 'Modules\Build\SanitizedEnvironment.psm1') -Force -Global -ErrorAction Stop
