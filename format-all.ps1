<#
.SYNOPSIS
    Deprecated compatibility wrapper for format.ps1. It will be removed after one release.

.DESCRIPTION
    The repository-root format.ps1 is the format command. It takes clang-format from
    Install-CiFormatter in the pinned .ci-engine and formats the C and C++ files that CI
    formats, or verifies them with -Check. This wrapper keeps the old command name working for
    one release: it runs format.ps1, or format.ps1 -Check when called with -WhatIf. Call
    format.ps1 directly; format-all.ps1 will be removed.

.PARAMETER RepositoryRoot
    Retained for compatibility. format.ps1 always formats its own checkout, so the only accepted
    value is this script's own folder (the default); any other folder is refused before anything
    runs.

.OUTPUTS
    Text status from format.ps1. Without -WhatIf, files that need formatting are rewritten.

.NOTES
    Deprecated; use format.ps1. Requires PowerShell 7.4, Git and the .ci-engine checkout at the
    engine.ref that .ci/project.json pins. format.ps1 may download and verify the pinned
    formatter into .build/shared-format. -WhatIf runs format.ps1 -Check, which writes no source
    and fails, listing every file that needs formatting. A failure of format.ps1 propagates
    unchanged as a terminating error (nonzero exit). Primary consumer: existing local callers of
    the old command name.

.EXAMPLE
    .\format-all.ps1 -WhatIf

    Runs .\format.ps1 -Check.

.EXAMPLE
    .\format.ps1

    The replacement command: formats the files CI formats.
#>
#Requires -Version 7.4
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
param(
    [string]$RepositoryRoot = $PSScriptRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Warning 'format-all.ps1 is deprecated and will be removed after one release; use format.ps1 (format.ps1 -Check verifies).'
$requested = [IO.Path]::TrimEndingDirectorySeparator($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($RepositoryRoot))
if (-not [string]::Equals($requested, [IO.Path]::TrimEndingDirectorySeparator($PSScriptRoot), [StringComparison]::OrdinalIgnoreCase)) {
    throw "format-all.ps1 forwards to format.ps1, which formats only its own checkout. Run format.ps1 in '$requested' instead."
}

$format = Join-Path $PSScriptRoot 'format.ps1'
if ($PSCmdlet.ShouldProcess($PSScriptRoot, 'Format the C and C++ files CI formats with format.ps1')) {
    & $format
} elseif ($WhatIfPreference) {
    # The preview is the read-only check; format.ps1 itself has no ShouldProcess surface.
    $WhatIfPreference = $false
    & $format -Check
}
