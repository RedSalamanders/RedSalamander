Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/HostedTestDesktop.psm1') -Force

Describe 'Hosted desktop fixture lifecycle' {
    BeforeEach {
        $script:priorGitHubActions = [Environment]::GetEnvironmentVariable('GITHUB_ACTIONS', 'Process')
        $env:GITHUB_ACTIONS = 'true'
        $smallMode = [pscustomobject]@{ Width = 1024; Height = 768; BitsPerPixel = 32; Frequency = 60 }
        $largeMode = [pscustomobject]@{ Width = 1920; Height = 1080; BitsPerPixel = 32; Frequency = 60 }
        $small = [pscustomobject]@{ Mode = $smallMode; Width = 1024; Height = 768; WorkWidth = 1024; WorkHeight = 728; CoordinateScale = 1.0 }
        $large = [pscustomobject]@{ Mode = $largeMode; Width = 1920; Height = 1080; WorkWidth = 1920; WorkHeight = 1040; CoordinateScale = 1.0 }
        $global:rsHostedDesktopFixtureTest = [pscustomobject]@{
            States = [Collections.Generic.Queue[object]]::new()
            Small = $small
            Large = $large
            ActionCalls = 0
            NativeExit = 0
        }
        $global:rsHostedDesktopFixtureTest.States.Enqueue($small)
        $global:rsHostedDesktopFixtureTest.States.Enqueue($large)
        $global:rsHostedDesktopFixtureTest.States.Enqueue($small)
        Mock Get-RSHostedDesktopState { $global:rsHostedDesktopFixtureTest.States.Dequeue() } -ModuleName HostedTestDesktop
        Mock Get-RSHostedDisplayModes {
            [pscustomobject]@{ Width = 1280; Height = 720; BitsPerPixel = 32; Frequency = 60 }
            [pscustomobject]@{ Width = 3840; Height = 2160; BitsPerPixel = 32; Frequency = 60 }
            $global:rsHostedDesktopFixtureTest.Large.Mode
        } -ModuleName HostedTestDesktop
        Mock Set-RSHostedDisplayMode {} -ModuleName HostedTestDesktop
    }
    AfterEach {
        [Environment]::SetEnvironmentVariable('GITHUB_ACTIONS', $script:priorGitHubActions, 'Process')
        Remove-Variable rsHostedDesktopFixtureTest -Scope Global -ErrorAction SilentlyContinue
    }
    It 'selects a sufficient enumerated mode and restores it after successful qualification' {
        Invoke-RSWithHostedTestDesktop { $global:rsHostedDesktopFixtureTest.ActionCalls++ }
        $global:rsHostedDesktopFixtureTest.ActionCalls | Should Be 1
        Assert-MockCalled Set-RSHostedDisplayMode -ModuleName HostedTestDesktop -Scope It -Times 2 -Exactly
        Assert-MockCalled Set-RSHostedDisplayMode -ModuleName HostedTestDesktop -Scope It -Times 1 -Exactly -ParameterFilter { $Mode.Width -eq 1920 -and $Mode.Height -eq 1080 }
        Assert-MockCalled Set-RSHostedDisplayMode -ModuleName HostedTestDesktop -Scope It -Times 1 -Exactly -ParameterFilter { $Mode.Width -eq 1024 -and $Mode.Height -eq 768 }
    }
    It 'restores the display while propagating a failing native action' {
        { Invoke-RSWithHostedTestDesktop { throw 'seeded native failure' } } | Should Throw 'seeded native failure'
        Assert-MockCalled Set-RSHostedDisplayMode -ModuleName HostedTestDesktop -Scope It -Times 1 -Exactly -ParameterFilter { $Mode.Width -eq 1024 }
    }
    It 'propagates a checked nonzero native child exit while restoring the display' {
        $action = {
            $PSNativeCommandUseErrorActionPreference = $false
            & (Get-Process -Id $PID).Path -NoProfile -NonInteractive -Command 'exit 7'
            $global:rsHostedDesktopFixtureTest.NativeExit = $LASTEXITCODE
            if ($LASTEXITCODE -ne 0) { throw "Native fixture failed with exit code $LASTEXITCODE." }
        }
        { Invoke-RSWithHostedTestDesktop -Action $action } | Should Throw 'Native fixture failed with exit code 7'
        $global:rsHostedDesktopFixtureTest.NativeExit | Should Be 7
        Assert-MockCalled Set-RSHostedDisplayMode -ModuleName HostedTestDesktop -Scope It -Times 2 -Exactly
        Assert-MockCalled Set-RSHostedDisplayMode -ModuleName HostedTestDesktop -Scope It -Times 1 -Exactly -ParameterFilter { $Mode.Width -eq 1024 }
    }
    It 'accounts for display coordinate scaling instead of choosing a clipped physical mode' {
        $global:rsHostedDesktopFixtureTest.Small.CoordinateScale = 1.5
        Invoke-RSWithHostedTestDesktop { $global:rsHostedDesktopFixtureTest.ActionCalls++ }
        $global:rsHostedDesktopFixtureTest.ActionCalls | Should Be 1
        Assert-MockCalled Set-RSHostedDisplayMode -ModuleName HostedTestDesktop -Scope It -Times 1 -Exactly -ParameterFilter { $Mode.Width -eq 3840 }
        Assert-MockCalled Set-RSHostedDisplayMode -ModuleName HostedTestDesktop -Scope It -Times 0 -Exactly -ParameterFilter { $Mode.Width -eq 1920 }
    }
    It 'preserves an already sufficient display without any mode change' {
        $global:rsHostedDesktopFixtureTest.States.Clear()
        $global:rsHostedDesktopFixtureTest.States.Enqueue($global:rsHostedDesktopFixtureTest.Large)
        $global:rsHostedDesktopFixtureTest.States.Enqueue($global:rsHostedDesktopFixtureTest.Large)
        Invoke-RSWithHostedTestDesktop { $global:rsHostedDesktopFixtureTest.ActionCalls++ }
        $global:rsHostedDesktopFixtureTest.ActionCalls | Should Be 1
        Assert-MockCalled Set-RSHostedDisplayMode -ModuleName HostedTestDesktop -Scope It -Times 0 -Exactly
    }
    It 'fails before native execution when the driver offers no sufficient mode' {
        Mock Get-RSHostedDisplayModes { $global:rsHostedDesktopFixtureTest.Small.Mode } -ModuleName HostedTestDesktop
        { Invoke-RSWithHostedTestDesktop { $global:rsHostedDesktopFixtureTest.ActionCalls++ } } | Should Throw 'No supported hosted display mode'
        $global:rsHostedDesktopFixtureTest.ActionCalls | Should Be 0
        Assert-MockCalled Set-RSHostedDisplayMode -ModuleName HostedTestDesktop -Scope It -Times 0 -Exactly
    }
    It 'refuses local invocation before reading or changing the desktop' {
        $env:GITHUB_ACTIONS = 'false'
        { Invoke-RSWithHostedTestDesktop { throw 'must not execute' } } | Should Throw 'restricted to GitHub Actions'
        Assert-MockCalled Get-RSHostedDesktopState -ModuleName HostedTestDesktop -Scope It -Times 0 -Exactly
        Assert-MockCalled Set-RSHostedDisplayMode -ModuleName HostedTestDesktop -Scope It -Times 0 -Exactly
    }
    It 'fails qualification when display restoration fails' {
        Mock Set-RSHostedDisplayMode { if ($Mode.Width -eq 1024) { throw 'seeded restoration failure' } } -ModuleName HostedTestDesktop
        { Invoke-RSWithHostedTestDesktop { $global:rsHostedDesktopFixtureTest.ActionCalls++ } } | Should Throw 'seeded restoration failure'
        $global:rsHostedDesktopFixtureTest.ActionCalls | Should Be 1
    }
}
