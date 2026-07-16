Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$testRunPlanScript = Join-Path $repoRoot 'Tools\TestRunPlan.ps1'
$helperScript = Join-Path $repoRoot 'Tools\VcpkgInstallSafety.ps1'
. $testRunPlanScript
. $helperScript

function New-RSTemporaryVcpkgInstallSafetyRoot {
    return (New-RSTestSandboxScratchDirectory `
            -RepoRoot $repoRoot `
            -Harness 'tools-pester' `
            -Case "vcpkg-install-safety-$([System.Guid]::NewGuid().ToString('N'))")
}

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
        $root = New-RSTemporaryVcpkgInstallSafetyRoot
        try {
            $path = Resolve-RSVcpkgSafeChildPath -Root $root -Child 'x64-windows' -Description 'test triplet'

            $path | Should Be (Join-Path $root 'x64-windows')
        } finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'rejects child paths that escape the intended root' {
        $root = New-RSTemporaryVcpkgInstallSafetyRoot
        try {
            Assert-ThrowsTerminatingError { Resolve-RSVcpkgSafeChildPath -Root $root -Child '..\outside' -Description 'test triplet' }
            Assert-ThrowsTerminatingError { Resolve-RSVcpkgSafeChildPath -Root $root -Child '..\..' -Description 'test triplet' }
        } finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    It 'merges a single source file while StrictMode is active' {
        $root = New-RSTemporaryVcpkgInstallSafetyRoot
        try {
            $src = Join-Path $root 'src\test-triplet'
            $dst = Join-Path $root 'dst\test-triplet'
            New-Item -ItemType Directory -Path (Join-Path $src 'include') -Force | Out-Null
            New-Item -ItemType Directory -Path $dst -Force | Out-Null
            [System.IO.File]::WriteAllText((Join-Path $src 'include\header.h'), '// v1')

            $result = Merge-RSVcpkgTripletSafe -SourcePath $src -DestinationPath $dst -TripletName 'test-triplet'

            $result.FilesTotal | Should Be 1
            $result.FilesCopied | Should Be 1
            Test-Path -LiteralPath (Join-Path $dst 'include\header.h') | Should Be $true
        } finally {
            Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}
