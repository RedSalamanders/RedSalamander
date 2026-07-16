Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$helperScript = Join-Path $repoRoot 'Tools\Run-MtpLiveCloseout.ps1'

function Get-RSText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    return Get-Content -LiteralPath (Join-Path $repoRoot $Path) -Raw
}

Describe 'MTP live closeout helper contracts' {
    It 'keeps live mode gated on an approved device and scratch folder' {
        $source = Get-RSText -Path 'Tools\Run-MtpLiveCloseout.ps1'

        $source | Should Match '\[CmdletBinding\(DefaultParameterSetName = ''Live''\)\]'
        $source | Should Match '\[Parameter\(ParameterSetName = ''Live'', Mandatory = \$true\)\]\s*\[ValidateNotNullOrEmpty\(\)\]\s*\[string\]\$Device'
        $source | Should Match '\[Parameter\(ParameterSetName = ''Live'', Mandatory = \$true\)\]\s*\[ValidateNotNullOrEmpty\(\)\]\s*\[string\]\$Scratch'
        $source | Should Match 'REDSALAMANDER_SELFTEST_MTP_DEVICE'
        $source | Should Match 'REDSALAMANDER_SELFTEST_MTP_SCRATCH'
    }

    It 'keeps probe mode read-only by using an impossible requested device and empty scratch value' {
        $source = Get-RSText -Path 'Tools\Run-MtpLiveCloseout.ps1'

        $source | Should Match '\[Parameter\(ParameterSetName = ''Probe'', Mandatory = \$true\)\]\s*\[switch\]\$ProbeNoDevice'
        $source | Should Match '__redsal_mtp_probe_no_such_device__'
        $source | Should Match '\$scratchValue = if \(\$ProbeNoDevice\) \{ '''' \} else \{ \$Scratch \}'
    }

    It 'runs Run-AllTests in a child PowerShell process so exit cannot skip archival work' {
        $source = Get-RSText -Path 'Tools\Run-MtpLiveCloseout.ps1'

        $source | Should Match '\$pwsh = \(Get-Process -Id \$PID\)\.Path'
        $source | Should Match 'Start-Process -FilePath \$pwsh'
        $source | Should Match '''-File'', \$runAll'
        $source | Should Match '''-CaseFilter'', ''mtp_live_device_smoke'''
        $source | Should Match '-RedirectStandardOutput \$stdout'
        $source | Should Match '-RedirectStandardError \$stderr'
    }

    It 'archives continuation evidence and restores caller MTP environment variables' {
        $source = Get-RSText -Path 'Tools\Run-MtpLiveCloseout.ps1'

        $source | Should Match 'pnp-device-probe\.txt'
        $source | Should Match 'cim-wpd-probe\.txt'
        $source | Should Match 'command\.txt'
        $source | Should Match 'mtp-env\.txt'
        $source | Should Match 'run-all-tests-results\.json'
        $source | Should Match 'summary\.md'
        $source | Should Match 'finally\s*\{\s*foreach \(\$entry in \$previousEnv\.GetEnumerator\(\)\)'
        $source | Should Match 'Set-ProcessEnvironmentValue -Name \$entry\.Key -Value \$entry\.Value'
    }
}
