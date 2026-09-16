<#
.SYNOPSIS
Update RedSalamander's DxUi lock to the current successfully validated DxUi main commit.
.DESCRIPTION
Updates only Dependencies/DxUi.lock.json after confirming the exact DxUi main commit has a completed successful
CI validation. By default it then runs the RedSalamander Full validation suite. -UpdateOnly retains the lock change
without running local product validation, for use when equivalent validation was completed in another environment.
.PARAMETER UpdateOnly
Skip RedSalamander validation after changing the lock. This never bypasses DxUi's own successful CI requirement.
.OUTPUTS
Writes the selected commit and whether the lock changed or product validation ran. Changes only
Dependencies/DxUi.lock.json; it does not commit or push the branch.
.NOTES
Requires public GitHub API access. If the candidate is not exactly the current successful DxUi main CI commit, the
lock remains unchanged. A failed local Full validation intentionally leaves the updated lock for diagnosis.
.EXAMPLE
.\Tools\Update-DxUi.ps1
.EXAMPLE
.\Tools\Update-DxUi.ps1 -UpdateOnly
#>
[CmdletBinding(SupportsShouldProcess)]
param([switch]$UpdateOnly)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
Import-Module (Join-Path $PSScriptRoot 'Modules/Build/DxUiUpdate.psm1') -Force -ErrorAction Stop
$validationAction = if ($UpdateOnly) {
    $null
}
else {
    {
        & (Join-Path $repoRoot 'Tools/Run-AllTests.ps1') -Suite Full -ValidationMode Fresh
        if ($LASTEXITCODE -ne 0) { throw "RedSalamander Full validation failed with exit code $LASTEXITCODE." }
    }
}
Invoke-RSDxUiUpdate -RepoRoot $repoRoot -UpdateOnly:$UpdateOnly -ValidationAction $validationAction -WhatIf:$WhatIfPreference
