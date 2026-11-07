Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/HostedTestDesktop.psm1') -Force
Import-Module (Join-Path $repoRoot 'Tools/Tests/TestSupport.psm1')

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
    It 'reads desktop diagnostics without changing the display or invoking qualification' {
        $desktop = [pscustomobject]@{ WindowStation='WinSta0'; WindowStationError=0; ThreadDesktop='Default'; ThreadDesktopError=0; InputDesktop=''; InputDesktopError=5; ForegroundDesktop=''; ForegroundDesktopError=5 }
        $global:rsHostedDesktopFixtureTest.Small | Add-Member -NotePropertyName DesktopContext -NotePropertyValue $desktop
        $result = Get-RSHostedDesktopDiagnostics
        $result.DesktopContext.ThreadDesktop | Should Be 'Default'
        $result.DesktopContext.InputDesktop | Should Be ''
        $result.DesktopContext.InputDesktopError | Should Be 5
        $result.DesktopContext.ForegroundDesktopError | Should Be 5
        ($result.PSObject.Properties.Name -contains 'Mode') | Should Be $false
        $global:rsHostedDesktopFixtureTest.ActionCalls | Should Be 0
        Assert-MockCalled Get-RSHostedDesktopState -ModuleName HostedTestDesktop -Scope It -Times 1 -Exactly
        Assert-MockCalled Get-RSHostedDisplayModes -ModuleName HostedTestDesktop -Scope It -Times 0 -Exactly
        Assert-MockCalled Set-RSHostedDisplayMode -ModuleName HostedTestDesktop -Scope It -Times 0 -Exactly
    }
    It 'refuses local diagnostics before reading the desktop' {
        $env:GITHUB_ACTIONS = 'false'
        $failure = $null
        try { $null = Get-RSHostedDesktopDiagnostics } catch { $failure = $_.Exception.Message }
        ($null -ne $failure) | Should Be $true
        $failure | Should Match 'restricted to GitHub Actions'
        Assert-MockCalled Get-RSHostedDesktopState -ModuleName HostedTestDesktop -Scope It -Times 0 -Exactly
        Assert-MockCalled Set-RSHostedDisplayMode -ModuleName HostedTestDesktop -Scope It -Times 0 -Exactly
    }
}

Describe 'Hosted owned desktop entry lifetime' {
    BeforeEach {
        $script:priorOwnedDesktopEnvironment = @{}
        foreach ($name in @('GITHUB_ACTIONS','RUNNER_ENVIRONMENT','GITHUB_REPOSITORY')) {
            $script:priorOwnedDesktopEnvironment[$name] = [Environment]::GetEnvironmentVariable($name, 'Process')
        }
        $env:GITHUB_ACTIONS='true'; $env:RUNNER_ENVIRONMENT='github-hosted'; $env:GITHUB_REPOSITORY='RedSalamanders/RedSalamander'
        $global:rsOwnedDesktopTest = [pscustomobject]@{ Events=[Collections.Generic.List[string]]::new(); FailCreate=$false; FailScope=$false; FailScopeRetire=$false; FailRestore=$false; ReportRestored=$true }
        $lease = [pscustomobject]@{ Name=('RedSalamander.Test.' + ('a' * 32)); Restored=$false; ForegroundBoundaryReady=$true }
        $lease | Add-Member -MemberType ScriptMethod -Name Dispose -Value {
            $global:rsOwnedDesktopTest.Events.Add('restore')
            if ($global:rsOwnedDesktopTest.FailRestore) { throw 'seeded restore failure' }
            $this.Restored=$global:rsOwnedDesktopTest.ReportRestored
        }
        $scope = [pscustomobject]@{}
        $scope | Add-Member -MemberType ScriptMethod -Name Dispose -Value {
            $global:rsOwnedDesktopTest.Events.Add('retire-process-scope')
            if ($global:rsOwnedDesktopTest.FailScopeRetire) { throw 'seeded process scope retirement failure' }
        }
        $global:rsOwnedDesktopTest | Add-Member -NotePropertyName Lease -NotePropertyValue $lease
        $global:rsOwnedDesktopTest | Add-Member -NotePropertyName Scope -NotePropertyValue $scope
        Mock New-RSHostedOwnedDesktop {
            $global:rsOwnedDesktopTest.Events.Add('create')
            if ($global:rsOwnedDesktopTest.FailCreate) { throw 'seeded desktop creation failure' }
            $global:rsOwnedDesktopTest.Lease
        } -ModuleName HostedTestDesktop
        Mock New-RSHostedDesktopProcessScope {
            $global:rsOwnedDesktopTest.Events.Add('process-scope')
            if ($global:rsOwnedDesktopTest.FailScope) { throw 'seeded process scope creation failure' }
            $global:rsOwnedDesktopTest.Scope
        } -ModuleName HostedTestDesktop
        Mock Set-RSHostedDisplayMode {} -ModuleName HostedTestDesktop
    }
    AfterEach {
        foreach ($name in $script:priorOwnedDesktopEnvironment.Keys) {
            [Environment]::SetEnvironmentVariable($name, $script:priorOwnedDesktopEnvironment[$name], 'Process')
        }
        Remove-Variable rsOwnedDesktopTest -Scope Global -ErrorAction SilentlyContinue
    }
    It 'selects the owned fixture for this project hosted native entry without native calls' {
        Get-RSHostedInputDesktopRequest | Should Be 'owned'
        $global:rsOwnedDesktopTest.Events.Count | Should Be 0
    }
    It 'preserves the original fixture for both sources of a paired comparison' {
        ($null -eq (Get-RSHostedInputDesktopRequest -PairedComparison)) | Should Be $true
        $global:rsOwnedDesktopTest.Events.Count | Should Be 0
    }
    It 'does not select an owned desktop for local execution' {
        $env:GITHUB_ACTIONS='false'
        ($null -eq (Get-RSHostedInputDesktopRequest)) | Should Be $true
        $global:rsOwnedDesktopTest.Events.Count | Should Be 0
    }
    It 'does not select an owned desktop for a self hosted runner' {
        $env:RUNNER_ENVIRONMENT='self-hosted'
        ($null -eq (Get-RSHostedInputDesktopRequest)) | Should Be $true
        $global:rsOwnedDesktopTest.Events.Count | Should Be 0
    }
    It 'does not select an owned desktop for another repository' {
        $env:GITHUB_REPOSITORY='fixture/other'
        ($null -eq (Get-RSHostedInputDesktopRequest)) | Should Be $true
        $global:rsOwnedDesktopTest.Events.Count | Should Be 0
    }
    It 'restores after contained entry drain and preserves the action result' {
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedOwnedDesktop { $global:rsOwnedDesktopTest.Events.Add('entry-drained'); 7 } }
        $r.Succeeded | Should Be $true
        $r.Output[0] | Should Be 7
        ($global:rsOwnedDesktopTest.Events -join ',') | Should Be 'create,process-scope,entry-drained,retire-process-scope,restore'
        Assert-MockCalled Set-RSHostedDisplayMode -ModuleName HostedTestDesktop -Scope It -Times 0 -Exactly
    }
    It 'preserves a failing native entry after restoring the original desktop' {
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedOwnedDesktop { throw 'seeded native entry failure' } }
        $r.Succeeded | Should Be $false
        $r.Error.Exception.Message | Should Match 'seeded native entry failure'
        ($global:rsOwnedDesktopTest.Events -join ',') | Should Be 'create,process-scope,retire-process-scope,restore'
    }
    It 'restores when process scope construction fails' {
        $global:rsOwnedDesktopTest.FailScope=$true
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedOwnedDesktop { throw 'must not execute' } }
        $r.Succeeded | Should Be $false
        $r.Error.Exception.Message | Should Match 'seeded process scope creation failure'
        ($global:rsOwnedDesktopTest.Events -join ',') | Should Be 'create,process-scope,restore'
    }
    It 'does not execute the action after desktop creation fails' {
        $global:rsOwnedDesktopTest.FailCreate=$true
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedOwnedDesktop { throw 'must not execute' } }
        $r.Succeeded | Should Be $false
        $r.Error.Exception.Message | Should Match 'seeded desktop creation failure'
        ($global:rsOwnedDesktopTest.Events -join ',') | Should Be 'create'
    }
    It 'restores and refuses native execution without a positive foreground boundary witness' {
        $global:rsOwnedDesktopTest.Lease.ForegroundBoundaryReady=$false
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedOwnedDesktop { throw 'must not execute' } }
        $r.Succeeded | Should Be $false
        $r.Error.Exception.Message | Should Match 'no verified foreground boundary'
        ($global:rsOwnedDesktopTest.Events -join ',') | Should Be 'create,restore'
        Assert-MockCalled New-RSHostedDesktopProcessScope -ModuleName HostedTestDesktop -Scope It -Times 0 -Exactly
    }
    It 'retains both native and restoration failures' {
        $global:rsOwnedDesktopTest.FailRestore=$true
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedOwnedDesktop { throw 'seeded native entry failure' } }
        $r.Succeeded | Should Be $false
        $r.Error.Exception.Message | Should Match 'seeded native entry failure'
        $r.Error.Exception.Message | Should Match 'seeded restore failure'
    }
    It 'fails an otherwise successful action when restoration fails' {
        $global:rsOwnedDesktopTest.FailRestore=$true
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedOwnedDesktop { 0 } }
        $r.Succeeded | Should Be $false
        $r.Error.Exception.Message | Should Match 'seeded restore failure'
    }
    It 'still restores the input desktop if process scope retirement fails' {
        $global:rsOwnedDesktopTest.FailScopeRetire=$true
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedOwnedDesktop { 0 } }
        $r.Succeeded | Should Be $false
        $r.Error.Exception.Message | Should Match 'seeded process scope retirement failure'
        $global:rsOwnedDesktopTest.Events[$global:rsOwnedDesktopTest.Events.Count-1] | Should Be 'restore'
    }
    It 'rejects a missing positive restoration witness' {
        $global:rsOwnedDesktopTest.ReportRestored=$false
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedOwnedDesktop { 0 } }
        $r.Succeeded | Should Be $false
        $r.Error.Exception.Message | Should Match 'did not restore its original input desktop'
    }
    It 'refuses local invocation before any desktop or native action' {
        $env:GITHUB_ACTIONS='false'
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedOwnedDesktop { throw 'must not execute' } }
        $r.Succeeded | Should Be $false
        $r.Error.Exception.Message | Should Match 'restricted to the project GitHub-hosted runner'
        $global:rsOwnedDesktopTest.Events.Count | Should Be 0
    }
    It 'refuses another repository before any desktop or native action' {
        $env:GITHUB_REPOSITORY='fixture/other'
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedOwnedDesktop { throw 'must not execute' } }
        $r.Succeeded | Should Be $false
        $global:rsOwnedDesktopTest.Events.Count | Should Be 0
    }
    It 'refuses a self-hosted runner before any desktop or native action' {
        $env:RUNNER_ENVIRONMENT='self-hosted'
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedOwnedDesktop { throw 'must not execute' } }
        $r.Succeeded | Should Be $false
        $global:rsOwnedDesktopTest.Events.Count | Should Be 0
    }
}
