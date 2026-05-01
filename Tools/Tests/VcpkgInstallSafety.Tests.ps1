Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$helperScript = Join-Path $repoRoot 'Tools\VcpkgInstallSafety.ps1'
. $helperScript

function Assert-ThrowsTerminatingError {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock
    )

    $threw = $false
    try {
        $null = & $ScriptBlock
    } catch {
        $threw = $true
    }

    $threw | Should Be $true
}

Describe 'Vcpkg install safety helper' {
    It 'accepts vcpkg triplet leaf names' {
        Assert-RSVcpkgTripletLeafName -Triplet 'x64-windows' | Should Be 'x64-windows'
        Assert-RSVcpkgTripletLeafName -Triplet 'arm64-windows' | Should Be 'arm64-windows'
        Assert-RSVcpkgTripletLeafName -Triplet 'x64-windows-static-md' | Should Be 'x64-windows-static-md'
    }

    It 'rejects triplet values that are paths' {
        Assert-ThrowsTerminatingError { Assert-RSVcpkgTripletLeafName -Triplet '..' }
        Assert-ThrowsTerminatingError { Assert-RSVcpkgTripletLeafName -Triplet '..\..' }
        Assert-ThrowsTerminatingError { Assert-RSVcpkgTripletLeafName -Triplet 'x64\windows' }
        Assert-ThrowsTerminatingError { Assert-RSVcpkgTripletLeafName -Triplet 'x64/windows' }
    }

    It 'resolves child paths under the intended root' {
        $root = Join-Path ([System.IO.Path]::GetTempPath()) 'rs-vcpkg-install-root-test'
        $path = Resolve-RSVcpkgSafeChildPath -Root $root -Child 'x64-windows' -Description 'test triplet'

        $path | Should Be (Join-Path $root 'x64-windows')
    }

    It 'rejects child paths that escape the intended root' {
        $root = Join-Path ([System.IO.Path]::GetTempPath()) 'rs-vcpkg-install-root-test'

        Assert-ThrowsTerminatingError { Resolve-RSVcpkgSafeChildPath -Root $root -Child '..\outside' -Description 'test triplet' }
        Assert-ThrowsTerminatingError { Resolve-RSVcpkgSafeChildPath -Root $root -Child '..\..' -Description 'test triplet' }
    }
}
