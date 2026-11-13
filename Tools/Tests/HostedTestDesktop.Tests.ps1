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

Describe 'Hosted Default desktop entry lifetime' {
    BeforeEach {
        $script:priorDefaultDesktopEnvironment = @{}
        foreach ($name in @('GITHUB_ACTIONS','RUNNER_ENVIRONMENT','GITHUB_REPOSITORY')) {
            $script:priorDefaultDesktopEnvironment[$name] = [Environment]::GetEnvironmentVariable($name, 'Process')
        }
        $env:GITHUB_ACTIONS='true'; $env:RUNNER_ENVIRONMENT='github-hosted'; $env:GITHUB_REPOSITORY='RedSalamanders/RedSalamander'
        $global:rsDefaultDesktopTest = [pscustomobject]@{ Events=[Collections.Generic.List[string]]::new(); FailCreate=$false; FailScope=$false; FailScopeRetire=$false; FailRestore=$false; NestedRestoreFailure=$false; ReportRestored=$true }
        $lease = [pscustomobject]@{ Name='Default'; Restored=$false; ForegroundBoundaryReady=$true }
        $lease | Add-Member -MemberType ScriptMethod -Name Dispose -Value {
            $global:rsDefaultDesktopTest.Events.Add('restore')
            if ($global:rsDefaultDesktopTest.NestedRestoreFailure) {
                throw [InvalidOperationException]::new('Hosted boundary retirement failed.', [TimeoutException]::new('seeded boundary join threshold exceeded'))
            }
            if ($global:rsDefaultDesktopTest.FailRestore) { throw 'seeded restore failure' }
            $this.Restored=$global:rsDefaultDesktopTest.ReportRestored
        }
        $scope = [pscustomobject]@{}
        $scope | Add-Member -MemberType ScriptMethod -Name Dispose -Value {
            $global:rsDefaultDesktopTest.Events.Add('retire-process-scope')
            if ($global:rsDefaultDesktopTest.FailScopeRetire) { throw 'seeded process scope retirement failure' }
        }
        $global:rsDefaultDesktopTest | Add-Member -NotePropertyName Lease -NotePropertyValue $lease
        $global:rsDefaultDesktopTest | Add-Member -NotePropertyName Scope -NotePropertyValue $scope
        Mock New-RSHostedDefaultDesktop {
            $global:rsDefaultDesktopTest.Events.Add('create')
            if ($global:rsDefaultDesktopTest.FailCreate) { throw 'seeded desktop creation failure' }
            $global:rsDefaultDesktopTest.Lease
        } -ModuleName HostedTestDesktop
        Mock New-RSHostedDesktopProcessScope {
            $global:rsDefaultDesktopTest.Events.Add('process-scope')
            if ($global:rsDefaultDesktopTest.FailScope) { throw 'seeded process scope creation failure' }
            $global:rsDefaultDesktopTest.Scope
        } -ModuleName HostedTestDesktop
        Mock Set-RSHostedDisplayMode {} -ModuleName HostedTestDesktop
    }
    AfterEach {
        foreach ($name in $script:priorDefaultDesktopEnvironment.Keys) {
            [Environment]::SetEnvironmentVariable($name, $script:priorDefaultDesktopEnvironment[$name], 'Process')
        }
        Remove-Variable rsDefaultDesktopTest -Scope Global -ErrorAction SilentlyContinue
    }
    It 'selects the rendered Default fixture for this project hosted native entry without native calls' {
        Get-RSHostedInputDesktopRequest | Should Be 'default'
        $global:rsDefaultDesktopTest.Events.Count | Should Be 0
    }
    It 'preserves the original fixture for both sources of a paired comparison' {
        ($null -eq (Get-RSHostedInputDesktopRequest -PairedComparison)) | Should Be $true
        $global:rsDefaultDesktopTest.Events.Count | Should Be 0
    }
    It 'does not select a hosted boundary for local execution' {
        $env:GITHUB_ACTIONS='false'
        ($null -eq (Get-RSHostedInputDesktopRequest)) | Should Be $true
        $global:rsDefaultDesktopTest.Events.Count | Should Be 0
    }
    It 'does not select a hosted boundary for a self hosted runner' {
        $env:RUNNER_ENVIRONMENT='self-hosted'
        ($null -eq (Get-RSHostedInputDesktopRequest)) | Should Be $true
        $global:rsDefaultDesktopTest.Events.Count | Should Be 0
    }
    It 'does not select a hosted boundary for another repository' {
        $env:GITHUB_REPOSITORY='fixture/other'
        ($null -eq (Get-RSHostedInputDesktopRequest)) | Should Be $true
        $global:rsDefaultDesktopTest.Events.Count | Should Be 0
    }
    It 'restores after contained entry drain and preserves the action result' {
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedDefaultDesktop { $global:rsDefaultDesktopTest.Events.Add('entry-drained'); 7 } }
        $r.Succeeded | Should Be $true
        $r.Output[0] | Should Be 7
        ($global:rsDefaultDesktopTest.Events -join ',') | Should Be 'create,process-scope,entry-drained,retire-process-scope,restore'
        Assert-MockCalled Set-RSHostedDisplayMode -ModuleName HostedTestDesktop -Scope It -Times 0 -Exactly
    }
    It 'preserves a failing native entry after restoring the original desktop' {
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedDefaultDesktop { throw 'seeded native entry failure' } }
        $r.Succeeded | Should Be $false
        $r.Error.Exception.Message | Should Match 'seeded native entry failure'
        ($global:rsDefaultDesktopTest.Events -join ',') | Should Be 'create,process-scope,retire-process-scope,restore'
    }
    It 'restores when process scope construction fails' {
        $global:rsDefaultDesktopTest.FailScope=$true
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedDefaultDesktop { throw 'must not execute' } }
        $r.Succeeded | Should Be $false
        $r.Error.Exception.Message | Should Match 'seeded process scope creation failure'
        ($global:rsDefaultDesktopTest.Events -join ',') | Should Be 'create,process-scope,restore'
    }
    It 'does not execute the action after desktop creation fails' {
        $global:rsDefaultDesktopTest.FailCreate=$true
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedDefaultDesktop { throw 'must not execute' } }
        $r.Succeeded | Should Be $false
        $r.Error.Exception.Message | Should Match 'seeded desktop creation failure'
        ($global:rsDefaultDesktopTest.Events -join ',') | Should Be 'create'
    }
    It 'restores and refuses native execution without a positive foreground boundary witness' {
        $global:rsDefaultDesktopTest.Lease.ForegroundBoundaryReady=$false
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedDefaultDesktop { throw 'must not execute' } }
        $r.Succeeded | Should Be $false
        $r.Error.Exception.Message | Should Match 'no verified foreground boundary'
        ($global:rsDefaultDesktopTest.Events -join ',') | Should Be 'create,restore'
        Assert-MockCalled New-RSHostedDesktopProcessScope -ModuleName HostedTestDesktop -Scope It -Times 0 -Exactly
    }
    It 'retains both native and restoration failures' {
        $global:rsDefaultDesktopTest.FailRestore=$true
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedDefaultDesktop { throw 'seeded native entry failure' } }
        $r.Succeeded | Should Be $false
        $r.Error.Exception.Message | Should Match 'seeded native entry failure'
        $r.Error.Exception.Message | Should Match 'seeded restore failure'
    }
    It 'fails an otherwise successful action when restoration fails' {
        $global:rsDefaultDesktopTest.FailRestore=$true
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedDefaultDesktop { 0 } }
        $r.Succeeded | Should Be $false
        $r.Error.Exception.Message | Should Match 'seeded restore failure'
    }
    It 'still restores the input desktop if process scope retirement fails' {
        $global:rsDefaultDesktopTest.FailScopeRetire=$true
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedDefaultDesktop { 0 } }
        $r.Succeeded | Should Be $false
        $r.Error.Exception.Message | Should Match 'seeded process scope retirement failure'
        $global:rsDefaultDesktopTest.Events[$global:rsDefaultDesktopTest.Events.Count-1] | Should Be 'restore'
    }
    It 'rejects a missing positive restoration witness' {
        $global:rsDefaultDesktopTest.ReportRestored=$false
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedDefaultDesktop { 0 } }
        $r.Succeeded | Should Be $false
        $r.Error.Exception.Message | Should Match 'did not restore its original foreground and focus'
    }
    It 'logs the inner retirement cause while retaining failure of a successful native entry' {
        Mock Write-Host {} -ModuleName HostedTestDesktop
        $global:rsDefaultDesktopTest.NestedRestoreFailure=$true
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedDefaultDesktop { 7 } }
        $r.Succeeded | Should Be $false
        $r.Error.Exception.Message | Should Match 'Hosted boundary retirement failed'
        Assert-MockCalled Write-Host -ModuleName HostedTestDesktop -Scope It -Times 1 -Exactly -ParameterFilter {
            "$Object" -match 'Hosted Default desktop retirement exception:' -and
            "$Object" -match 'System.TimeoutException: seeded boundary join threshold exceeded'
        }
    }
    It 'refuses local invocation before any desktop or native action' {
        $env:GITHUB_ACTIONS='false'
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedDefaultDesktop { throw 'must not execute' } }
        $r.Succeeded | Should Be $false
        $r.Error.Exception.Message | Should Match 'restricted to the project GitHub-hosted runner'
        $global:rsDefaultDesktopTest.Events.Count | Should Be 0
    }
    It 'refuses another repository before any desktop or native action' {
        $env:GITHUB_REPOSITORY='fixture/other'
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedDefaultDesktop { throw 'must not execute' } }
        $r.Succeeded | Should Be $false
        $global:rsDefaultDesktopTest.Events.Count | Should Be 0
    }
    It 'refuses a self-hosted runner before any desktop or native action' {
        $env:RUNNER_ENVIRONMENT='self-hosted'
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedDefaultDesktop { throw 'must not execute' } }
        $r.Succeeded | Should Be $false
        $global:rsDefaultDesktopTest.Events.Count | Should Be 0
    }
}

Describe 'Hosted boundary managed drain' {
    It 'keeps disposal pending after its threshold until a stalled managed worker exits' {
        & (Get-Module HostedTestDesktop) { Initialize-RSHostedDesktopNative }
        if (-not ('RedSalamander.Testing.ManagedBoundaryDrainControl' -as [type])) {
            Add-Type -TypeDefinition @'
using System;
using System.Reflection;
using System.Threading;
namespace RedSalamander.Testing {
    public static class ManagedBoundaryDrainControl {
        public static bool Run(Type boundaryType) {
            using (var started = new ManualResetEventSlim(false))
            using (var release = new ManualResetEventSlim(false))
            using (var disposeStarted = new ManualResetEventSlim(false))
            using (var finished = new ManualResetEventSlim(false)) {
                Exception failure = null;
                var worker = new Thread(() => { started.Set(); release.Wait(); }) { IsBackground = true };
                worker.Start();
                Thread disposer = null;
                bool disposerStarted = false;
                try {
                    if (!started.Wait(TimeSpan.FromSeconds(2))) throw new InvalidOperationException("Managed worker did not start.");
                    var ctor = boundaryType.GetConstructor(BindingFlags.Instance | BindingFlags.NonPublic, null,
                        new[] { typeof(Thread), typeof(TimeSpan) }, null);
                    if (ctor == null) throw new InvalidOperationException("Managed boundary constructor missing.");
                    var boundary = (IDisposable)ctor.Invoke(new object[] { worker, TimeSpan.FromMilliseconds(50) });
                    disposer = new Thread(() => {
                        disposeStarted.Set();
                        try { boundary.Dispose(); }
                        catch (TimeoutException error) { failure = error; }
                        catch (ObjectDisposedException error) { failure = error; }
                        catch (InvalidOperationException error) { failure = error; }
                        finally { finished.Set(); }
                    }) { IsBackground = true };
                    disposer.Start();
                    disposerStarted = true;
                    if (!disposeStarted.Wait(TimeSpan.FromSeconds(2))) throw new InvalidOperationException("Managed disposer did not start.");
                    bool returnedWhileWorkerLive = finished.Wait(TimeSpan.FromMilliseconds(250));
                    bool workerStillLive = worker.IsAlive;
                    release.Set();
                    if (!finished.Wait(TimeSpan.FromSeconds(2)) || !disposer.Join(TimeSpan.FromSeconds(2)))
                        throw new InvalidOperationException("Managed disposal did not finish after worker release.");
                    return !returnedWhileWorkerLive && workerStillLive && !worker.IsAlive && failure is TimeoutException;
                } finally { release.Set(); worker.Join(); if (disposerStarted) disposer.Join(); }
            }
        }
    }
}
'@
        }
        $type=[RedSalamander.Testing.HostedDesktopNative].Assembly.GetType('RedSalamander.Testing.HostedForegroundBoundary', $true)
        [RedSalamander.Testing.ManagedBoundaryDrainControl]::Run($type) | Should Be $true
    }
}

Describe 'Approved hosted ARM64 stock preparation' {
    BeforeEach {
        $script:savedActions=$env:GITHUB_ACTIONS; $script:savedRunner=$env:RUNNER_ENVIRONMENT; $script:savedRepo=$env:GITHUB_REPOSITORY
        $env:GITHUB_ACTIONS='true'; $env:RUNNER_ENVIRONMENT='github-hosted'; $env:GITHUB_REPOSITORY='RedSalamanders/RedSalamander'
        $global:stockObservations=[Collections.Generic.Queue[object]]::new()
        $global:stockCloses=[Collections.Generic.List[string]]::new()
        Mock Get-RSHostedStockSurface { $global:stockObservations.Dequeue() } -ModuleName HostedTestDesktop
    }
    AfterEach {
        $env:GITHUB_ACTIONS=$script:savedActions; $env:RUNNER_ENVIRONMENT=$script:savedRunner; $env:GITHUB_REPOSITORY=$script:savedRepo
        Remove-Variable stockObservations,stockCloses -Scope Global -ErrorAction SilentlyContinue
    }
    It 'limits opt-in to native ARM64 hosted project roles' {
        $p=@{Phase='ReleaseTest';Platform='ARM64';QualificationScope='Normal';OsArchitecture='Arm64';ProcessArchitecture='Arm64'}
        Get-RSHostedStockPreparationRequest @p | Should Be '1'
        $p.Phase='Qualification'; $p.QualificationScope='CommandsHandoff'; Get-RSHostedStockPreparationRequest @p | Should Be '1'
        $p.QualificationScope='R0ePaired'; Get-RSHostedStockPreparationRequest @p | Should BeNullOrEmpty
        $p.QualificationScope='DesktopProbe'; Get-RSHostedStockPreparationRequest @p | Should BeNullOrEmpty
        $p.QualificationScope='Normal'; $p.Platform='x64'; Get-RSHostedStockPreparationRequest @p | Should BeNullOrEmpty
        $p.Platform='ARM64'; $p.ProcessArchitecture='X64'; Get-RSHostedStockPreparationRequest @p | Should BeNullOrEmpty
        $p.ProcessArchitecture='Arm64'; $env:RUNNER_ENVIRONMENT='self-hosted'; Get-RSHostedStockPreparationRequest @p | Should BeNullOrEmpty
        $env:RUNNER_ENVIRONMENT='github-hosted'; $env:GITHUB_ACTIONS='false'; Get-RSHostedStockPreparationRequest @p | Should BeNullOrEmpty
        $env:GITHUB_ACTIONS='true'; $env:GITHUB_REPOSITORY='other/repository'; Get-RSHostedStockPreparationRequest @p | Should BeNullOrEmpty
    }
    It 'completes only after the approved account and search retire and shell is observed' {
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestSandbox.psm1') -Force
        $dir=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('hosted-stock-success-'+[guid]::NewGuid().ToString('N'))
        foreach($kind in @('Account','Search','Shell')) {
            $surface=[pscustomobject]@{Kind=$kind;Window=123;Thread=1;Process=2;Session=3;Title=$kind;Class='fixture';Image='fixture';ApplicationId='fixture';ForegroundDesktop='';ForegroundDesktopError=5}
            $surface | Add-Member ScriptMethod RequestClose { $global:stockCloses.Add($this.Kind) }
            $surface | Add-Member ScriptMethod Dispose { }
            $global:stockObservations.Enqueue($surface)
        }
        & (Get-Module HostedTestDesktop) { param($d) Invoke-RSHostedStockPreparationCore -Directory $d -SourceHead ('a'*40) -ObservationMilliseconds 0 } $dir
        ($global:stockCloses -join ',') | Should Be 'Account,Search'
        $completion=Get-Content (Join-Path $dir 'completion.json') -Raw | ConvertFrom-Json
        $completion.completed | Should Be $true
        $completion.last_observation.kind | Should Be 'Shell'
    }
    It 'retains a failing close and never requests a second close' {
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestSandbox.psm1') -Force
        $dir=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('hosted-stock-failure-'+[guid]::NewGuid().ToString('N'))
        $surface=[pscustomobject]@{Kind='Search';Window=123;Thread=1;Process=2;Session=3;Title='Search';Class='fixture';Image='fixture';ApplicationId='fixture';ForegroundDesktop='';ForegroundDesktopError=5}
        $surface | Add-Member ScriptMethod RequestClose { $global:stockCloses.Add($this.Kind); throw 'seeded close failure' }
        $surface | Add-Member ScriptMethod Dispose { }
        $global:stockObservations.Enqueue($surface)
        { & (Get-Module HostedTestDesktop) { param($d) Invoke-RSHostedStockPreparationCore -Directory $d -SourceHead ('a'*40) -ObservationMilliseconds 0 } $dir } | Should Throw 'seeded close failure'
        ($global:stockCloses -join ',') | Should Be 'Search'
        $result=Get-Content (Join-Path $dir 'search-result.json') -Raw | ConvertFrom-Json
        $result.retired | Should Be $false
        $result.error | Should Match 'seeded close failure'
        (Get-Content (Join-Path $dir 'completion.json') -Raw | ConvertFrom-Json).completed | Should Be $false
    }
    It 'observes the captured runner console without asking it to close' {
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestSandbox.psm1') -Force
        $dir=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('hosted-runner-console-'+[guid]::NewGuid().ToString('N'))
        $surface=[pscustomobject]@{Kind='RunnerConsole';Window=123;Thread=1;Process=2;Session=3;Title='fixture';Class='ConsoleWindowClass';Image='fixture';ApplicationId='';ForegroundDesktop='Default';ForegroundDesktopError=0}
        $surface | Add-Member ScriptMethod RequestClose { $global:stockCloses.Add($this.Kind); throw 'Runner console must never be closed.' }
        $surface | Add-Member ScriptMethod Dispose { }
        $global:stockObservations.Enqueue($surface)
        & (Get-Module HostedTestDesktop) { param($d) Invoke-RSHostedStockPreparationCore -Directory $d -SourceHead ('a'*40) -ObservationMilliseconds 0 } $dir
        $global:stockCloses.Count | Should Be 0
        $completion=Get-Content (Join-Path $dir 'completion.json') -Raw | ConvertFrom-Json
        $completion.completed | Should Be $true
        $completion.last_observation.kind | Should Be 'RunnerConsole'
        @(Get-ChildItem -LiteralPath $dir -Filter '*-intent.json').Count | Should Be 0
    }
    It 'refuses a second appearance without sending another close' {
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestSandbox.psm1') -Force
        $dir=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('hosted-stock-repeat-'+[guid]::NewGuid().ToString('N'))
        foreach($kind in @('Search','Search')) {
            $surface=[pscustomobject]@{Kind=$kind;Window=123;Thread=1;Process=2;Session=3;Title=$kind;Class='fixture';Image='fixture';ApplicationId='fixture';ForegroundDesktop='';ForegroundDesktopError=5}
            $surface | Add-Member ScriptMethod RequestClose { $global:stockCloses.Add($this.Kind) }
            $surface | Add-Member ScriptMethod Dispose { }
            $global:stockObservations.Enqueue($surface)
        }
        { & (Get-Module HostedTestDesktop) { param($d) Invoke-RSHostedStockPreparationCore -Directory $d -SourceHead ('a'*40) -ObservationMilliseconds 0 } $dir } | Should Throw 'reappeared'
        ($global:stockCloses -join ',') | Should Be 'Search'
        (Get-Content (Join-Path $dir 'completion.json') -Raw | ConvertFrom-Json).completed | Should Be $false
    }
    It 'refuses to close when its create-only intent already exists' {
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestSandbox.psm1') -Force
        $dir=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('hosted-stock-audit-'+[guid]::NewGuid().ToString('N'))
        [IO.File]::WriteAllText((Join-Path $dir 'search-intent.json'),'seeded previous intent')
        $surface=[pscustomobject]@{Kind='Search';Window=123;Thread=1;Process=2;Session=3;Title='Search';Class='fixture';Image='fixture';ApplicationId='fixture';ForegroundDesktop='';ForegroundDesktopError=5}
        $surface | Add-Member ScriptMethod RequestClose { $global:stockCloses.Add($this.Kind) }
        $surface | Add-Member ScriptMethod Dispose { }
        $global:stockObservations.Enqueue($surface)
        $attempt=Invoke-RSTestBlock { & (Get-Module HostedTestDesktop) { param($d) Invoke-RSHostedStockPreparationCore -Directory $d -SourceHead ('a'*40) -ObservationMilliseconds 0 } $dir }
        $attempt.Succeeded | Should Be $false
        $attempt.Error.Exception.Message | Should Match 'already exists'
        $global:stockCloses.Count | Should Be 0
        [IO.File]::ReadAllText((Join-Path $dir 'search-intent.json')) | Should Be 'seeded previous intent'
        (Get-Content (Join-Path $dir 'completion.json') -Raw | ConvertFrom-Json).completed | Should Be $false
    }
    It 'retains approved one-close intents and refuses an unknown foreground' {
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestSandbox.psm1') -Force
        $dir=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('hosted-stock-'+[guid]::NewGuid().ToString('N'))
        foreach($kind in @('Account','Search','Unknown')) {
            $surface=[pscustomobject]@{Kind=$kind;Window=123;Thread=1;Process=2;Session=3;Title=$kind;Class='fixture';Image='fixture';ApplicationId='fixture';ForegroundDesktop='';ForegroundDesktopError=5}
            $surface | Add-Member ScriptMethod RequestClose { $global:stockCloses.Add($this.Kind) }
            $surface | Add-Member ScriptMethod Dispose { }
            $global:stockObservations.Enqueue($surface)
        }
        { & (Get-Module HostedTestDesktop) { param($d) Invoke-RSHostedStockPreparationCore -Directory $d -SourceHead ('a'*40) -ObservationMilliseconds 0 } $dir } | Should Throw 'unapproved foreground surface'
        ($global:stockCloses -join ',') | Should Be 'Account,Search'
        (Get-Content (Join-Path $dir 'account-result.json') -Raw | ConvertFrom-Json).retired | Should Be $true
        (Get-Content (Join-Path $dir 'search-result.json') -Raw | ConvertFrom-Json).retired | Should Be $true
        $completion=Get-Content (Join-Path $dir 'completion.json') -Raw | ConvertFrom-Json
        $completion.completed | Should Be $false
        $completion.last_observation.kind | Should Be 'Unknown'
        $completion.last_observation.foreground_desktop_error | Should Be 5
    }
}

Describe 'Hosted foreground chord managed contracts' {
    It 'releases only inserted key downs that have not already been released' {
        & (Get-Module HostedTestDesktop) { Initialize-RSHostedDesktopNative }
        $assembly = ('RedSalamander.Testing.HostedDesktopNative' -as [type]).Assembly
        $boundary = $assembly.GetType('RedSalamander.Testing.HostedForegroundBoundary', $true)
        $flags = [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic
        $remaining = $boundary.GetMethod('RemainingOwnedChordKeyUps', $flags)
        $remaining | Should Not BeNullOrEmpty
        foreach ($sample in @(
            @{ downs=0; ups=0; expected='' },
            @{ downs=1; ups=0; expected='18' },
            @{ downs=1; ups=1; expected='' },
            @{ downs=2; ups=0; expected='135,18' },
            @{ downs=2; ups=1; expected='18' },
            @{ downs=2; ups=2; expected='' }
        )) {
            $result = $remaining.Invoke($null, [object[]]@([uint32]$sample.downs, [uint32]$sample.ups))
            (@($result) -join ',') | Should Be $sample.expected
        }
        foreach ($sample in @(@{ downs=3; ups=0 }, @{ downs=1; ups=2 }, @{ downs=0; ups=1 })) {
            $reflectionError = $null
            try { [void]$remaining.Invoke($null, [object[]]@([uint32]$sample.downs, [uint32]$sample.ups)) }
            catch [System.Management.Automation.MethodInvocationException] { $reflectionError = $_.Exception.GetBaseException() }
            ($reflectionError -is [ArgumentOutOfRangeException]) | Should Be $true
        }
    }
    It 'uses the Windows 64-bit INPUT and keyboard layout' {
        & (Get-Module HostedTestDesktop) { Initialize-RSHostedDesktopNative }
        $assembly = ('RedSalamander.Testing.HostedDesktopNative' -as [type]).Assembly
        $boundary = $assembly.GetType('RedSalamander.Testing.HostedForegroundBoundary', $true)
        $keyboard = $boundary.GetNestedType('KeyboardInput', [Reflection.BindingFlags]::NonPublic)
        $input = $boundary.GetNestedType('NativeInput', [Reflection.BindingFlags]::NonPublic)
        $keyboard | Should Not BeNullOrEmpty
        $input | Should Not BeNullOrEmpty
        $sizeOf = [Runtime.InteropServices.Marshal].GetMethod('SizeOf', [type[]]@([type]))
        $sizeOf.Invoke($null, [object[]]@($keyboard)) | Should Be 24
        $sizeOf.Invoke($null, [object[]]@($input)) | Should Be 40
        [Runtime.InteropServices.Marshal]::OffsetOf($keyboard, 'ExtraInfo').ToInt32() | Should Be 16
        [Runtime.InteropServices.Marshal]::OffsetOf($input, 'Type').ToInt32() | Should Be 0
        [Runtime.InteropServices.Marshal]::OffsetOf($input, 'Keyboard').ToInt32() | Should Be 8
    }
}

Describe 'Hosted foreground chord role policy' {
    It 'permits only selected project-hosted native x64 or opted-in ARM64 roles' {
        & (Get-Module HostedTestDesktop) { Initialize-RSHostedDesktopNative }
        $assembly = ('RedSalamander.Testing.HostedDesktopNative' -as [type]).Assembly
        $boundary = $assembly.GetType('RedSalamander.Testing.HostedForegroundBoundary', $true)
        $flags = [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic
        $policy = $boundary.GetMethod('AllowsOwnedBoundaryChord', $flags)
        $policy | Should Not BeNullOrEmpty
        $arm = [Runtime.InteropServices.Architecture]::Arm64
        $x64 = [Runtime.InteropServices.Architecture]::X64
        foreach ($sample in @(
            @{request='default'; stock='1'; actions='true'; runner='github-hosted'; repository='RedSalamanders/RedSalamander'; os=$arm; process=$arm; expected=$true},
            @{request='default'; stock='1'; actions='true'; runner='github-hosted'; repository='RedSalamanders/RedSalamander'; os=$x64; process=$x64; expected=$true},
            @{request='default'; stock=''; actions='true'; runner='github-hosted'; repository='RedSalamanders/RedSalamander'; os=$x64; process=$x64; expected=$true},
            @{request=''; stock='1'; actions='true'; runner='github-hosted'; repository='RedSalamanders/RedSalamander'; os=$arm; process=$arm; expected=$true},
            @{request=''; stock='1'; actions='true'; runner='github-hosted'; repository='RedSalamanders/RedSalamander'; os=$x64; process=$x64; expected=$false},
            @{request='default'; stock='1'; actions='false'; runner='github-hosted'; repository='RedSalamanders/RedSalamander'; os=$x64; process=$x64; expected=$false},
            @{request='default'; stock='1'; actions='true'; runner='self-hosted'; repository='RedSalamanders/RedSalamander'; os=$x64; process=$x64; expected=$false},
            @{request='default'; stock='1'; actions='true'; runner='github-hosted'; repository='other/repo'; os=$x64; process=$x64; expected=$false},
            @{request='default'; stock=''; actions='true'; runner='github-hosted'; repository='RedSalamanders/RedSalamander'; os=$arm; process=$arm; expected=$false},
            @{request='original'; stock='1'; actions='true'; runner='github-hosted'; repository='RedSalamanders/RedSalamander'; os=$x64; process=$x64; expected=$false},
            @{request='1'; stock='1'; actions='true'; runner='github-hosted'; repository='RedSalamanders/RedSalamander'; os=$x64; process=$x64; expected=$false},
            @{request='default'; stock='1'; actions='false'; runner='github-hosted'; repository='RedSalamanders/RedSalamander'; os=$arm; process=$arm; expected=$false},
            @{request='default'; stock='1'; actions='true'; runner='self-hosted'; repository='RedSalamanders/RedSalamander'; os=$arm; process=$arm; expected=$false},
            @{request='default'; stock='1'; actions='true'; runner='github-hosted'; repository='other/repo'; os=$arm; process=$arm; expected=$false},
            @{request='default'; stock='1'; actions='true'; runner='github-hosted'; repository='RedSalamanders/RedSalamander'; os=$x64; process=$arm; expected=$false},
            @{request='default'; stock='1'; actions='true'; runner='github-hosted'; repository='RedSalamanders/RedSalamander'; os=$arm; process=$x64; expected=$false}
        )) {
            $actual = $policy.Invoke($null, [object[]]@($sample.request, $sample.stock, $sample.actions, $sample.runner, $sample.repository, $sample.os, $sample.process))
            $actual | Should Be $sample.expected
        }
    }
}

Describe 'Hosted boundary return policy' {
    It 'waits only for null or owned foreground before the existing deadline' {
        & (Get-Module HostedTestDesktop) { Initialize-RSHostedDesktopNative }
        $boundary = ('RedSalamander.Testing.HostedDesktopNative' -as [type]).Assembly.GetType('RedSalamander.Testing.HostedForegroundBoundary', $true)
        $wait = $boundary.GetMethod('CanWaitForForegroundRestore', [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic)
        foreach ($sample in @(@{current=0;elapsed=0;expected=$true}, @{current=0;elapsed=1499;expected=$true}, @{current=0;elapsed=1500;expected=$false}, @{current=42;elapsed=1499;expected=$true}, @{current=42;elapsed=1500;expected=$false}, @{current=43;elapsed=0;expected=$false}, @{current=43;elapsed=1499;expected=$false})) {
            $wait.Invoke($null, [object[]]@([IntPtr]$sample.current, [IntPtr]42, [long]$sample.elapsed)) | Should Be $sample.expected
        }
        $caught = $null
        try { [void]$wait.Invoke($null, [object[]]@([IntPtr]::Zero, [IntPtr]42, [long]-1)) }
        catch [System.Management.Automation.MethodInvocationException] { $caught = $_.Exception.GetBaseException() }
        ($caught -is [ArgumentOutOfRangeException]) | Should Be $true
    }
    It 'clears the no-activation and tool-window styles while preserving other bits' {
        & (Get-Module HostedTestDesktop) { Initialize-RSHostedDesktopNative }
        $boundary = ('RedSalamander.Testing.HostedDesktopNative' -as [type]).Assembly.GetType('RedSalamander.Testing.HostedForegroundBoundary', $true)
        $style = $boundary.GetMethod('BoundaryReturnStyle', [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic)
        $style.Invoke($null, [object[]]@([long]0x08040080)) | Should Be ([long]0x00040000)
        $style.Invoke($null, [object[]]@([long]0x08000000)) | Should Be ([long]0)
        $style.Invoke($null, [object[]]@([long]0x00040080)) | Should Be ([long]0x00040000)
        $style.Invoke($null, [object[]]@([long]0x00000080)) | Should Be ([long]0)
        $style.Invoke($null, [object[]]@([long]0x00400020)) | Should Be ([long]0x00400020)
    }
    It 'bounds each message wait by the original restoration deadline' {
        & (Get-Module HostedTestDesktop) { Initialize-RSHostedDesktopNative }
        $boundary = ('RedSalamander.Testing.HostedDesktopNative' -as [type]).Assembly.GetType('RedSalamander.Testing.HostedForegroundBoundary', $true)
        $wait = $boundary.GetMethod('RestoreWaitMilliseconds', [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic)
        foreach ($sample in @(@{elapsed=0; expected=10}, @{elapsed=1490; expected=10}, @{elapsed=1499; expected=1}, @{elapsed=1500; expected=0}, @{elapsed=1501; expected=0})) {
            $wait.Invoke($null, [object[]]@([long]$sample.elapsed)) | Should Be ([uint32]$sample.expected)
        }
        $caught = $null
        try { [void]$wait.Invoke($null, [object[]]@([long]-1)) }
        catch [System.Management.Automation.MethodInvocationException] { $caught = $_.Exception.GetBaseException() }
        ($caught -is [ArgumentOutOfRangeException]) | Should Be $true
    }
}

Describe 'Hosted boundary inherited desktop policy' {
    It 'requires positive WinSta0 and Default thread and input identities without errors' {
        & (Get-Module HostedTestDesktop) { Initialize-RSHostedDesktopNative }
        $assembly = ('RedSalamander.Testing.HostedDesktopNative' -as [type]).Assembly
        $boundary = $assembly.GetType('RedSalamander.Testing.HostedForegroundBoundary', $true)
        $contextType = $assembly.GetType('RedSalamander.Testing.DesktopContext', $true)
        $method = $boundary.GetMethod('IsVerifiedInheritedDesktop', ([Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic))
        $method | Should Not BeNullOrEmpty
        $context = [Activator]::CreateInstance($contextType)
        $context.WindowStation='WinSta0'; $context.ThreadDesktop='Default'; $context.InputDesktop='Default'
        $method.Invoke($null, [object[]]@($context)) | Should Be $true
        foreach($property in @('WindowStation','ThreadDesktop','InputDesktop')) {
            $saved=$context.$property
            $context.$property=''; $method.Invoke($null, [object[]]@($context)) | Should Be $false
            $context.$property='Other'; $method.Invoke($null, [object[]]@($context)) | Should Be $false
            $context.$property=$saved
        }
        foreach($property in @('WindowStationError','ThreadDesktopError','InputDesktopError')) {
            $context.$property=5; $method.Invoke($null, [object[]]@($context)) | Should Be $false
            $context.$property=0
        }
        $method.Invoke($null, [object[]]@($null)) | Should Be $false
    }
}

Describe 'Hosted WSL startup guard managed contracts' {
    It 'accepts only the approved console identity, arguments and matching creation instance' {
        & (Get-Module HostedTestDesktop) { Initialize-RSHostedStockSurfaceNative }
        $type=('RedSalamander.Testing.HostedStockSurface' -as [type])
        $flags=[Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic
        $identity=$type.GetMethod('ConsoleIdentityMatches',$flags)
        $arguments=$type.GetMethod('WslArgumentsMatch',$flags)
        $creation=$type.GetMethod('WslCreationMatches',$flags)
        foreach($image in @('C:\Windows\System32\wsl.exe','C:\ProgramData\GitHub\HostedComputeAgent\hosted-compute-agent')) {
            $identity.Invoke($null,[object[]]@($image,'ConsoleWindowClass',$image,'',[int]15703,$image)) | Should Be $true
            $identity.Invoke($null,[object[]]@($image,'Other',$image,'',[int]15703,$image)) | Should Be $false
            $identity.Invoke($null,[object[]]@($image,'ConsoleWindowClass','C:\Other\wsl.exe','',[int]15703,$image)) | Should Be $false
            $identity.Invoke($null,[object[]]@($image,'ConsoleWindowClass',$image,'unexpected.app',[int]0,$image)) | Should Be $false
            $identity.Invoke($null,[object[]]@('other','ConsoleWindowClass',$image,'',[int]15703,$image)) | Should Be $false
            $identity.Invoke($null,[object[]]@($image,'ConsoleWindowClass',$image,'',[int]5,$image)) | Should Be $false
            $identity.Invoke($null,[object[]]@(($image+'.exe'),'ConsoleWindowClass',($image+'.exe'),'', [int]15703,$image)) | Should Be $false
        }
        $image='C:\Windows\System32\wsl.exe'
        $classify=$type.GetMethod('Classify',$flags)
        [string]$agent=Join-Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::CommonApplicationData)) 'GitHub/HostedComputeAgent/hosted-compute-agent'
        $classify.Invoke($null,[object[]]@([IntPtr]::Zero,$agent,'ConsoleWindowClass',$agent,'',[int]15703)) | Should Be 'RunnerConsole'
        $classify.Invoke($null,[object[]]@([IntPtr]::Zero,$agent,'ConsoleWindowClass',($agent+'.exe'),'', [int]15703)) | Should Be 'Unknown'
        $classify.Invoke($null,[object[]]@([IntPtr]::Zero,'other','ConsoleWindowClass',$agent,'',[int]15703)) | Should Be 'Unknown'
        $call=[object[]]::new(2);$call[1]=$image
        $call[0]=[string[]]@($image,'--update','--confirm','--prompt-before-exit');$arguments.Invoke($null,$call) | Should Be $true
        $call[0]=[string[]]@('wsl.exe','--update','--confirm','--prompt-before-exit');$arguments.Invoke($null,$call) | Should Be $true
        $call[0]=[string[]]@($image,'--update','--confirm');$arguments.Invoke($null,$call) | Should Be $false
        $call[0]=[string[]]@($image,'--update','--confirm','--prompt-before-exit','--extra');$arguments.Invoke($null,$call) | Should Be $false
        $call[0]=[string[]]@($image,'--update','--prompt-before-exit','--confirm');$arguments.Invoke($null,$call) | Should Be $false
        $call[0]=[string[]]@('C:\Other\wsl.exe','--update','--confirm','--prompt-before-exit');$arguments.Invoke($null,$call) | Should Be $false
        $creation.Invoke($null,[object[]]@([long]1009,[long]1000)) | Should Be $true
        $creation.Invoke($null,[object[]]@([long]1010,[long]1000)) | Should Be $false
        $creation.Invoke($null,[object[]]@([long]999,[long]1000)) | Should Be $true
        $creation.Invoke($null,[object[]]@([long]1000,[long]1010)) | Should Be $false
        $creation.Invoke($null,[object[]]@([long]0,[long]0)) | Should Be $false
    }
}

Describe 'Hosted WSL startup audit' {
    BeforeEach {
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestSandbox.psm1') -Force
        $global:rsWslSurfaceQueue=[Collections.Generic.Queue[object]]::new()
        $global:rsWslCloseCount=0
        $global:rsWslQuery=[pscustomobject]@{ProcessId=42;CreationDate=[DateTime]::UtcNow;ExecutablePath='C:\Windows\System32\wsl.exe';CommandLine='fixture'}
        Mock Get-RSHostedStockSurface { $global:rsWslSurfaceQueue.Dequeue() } -ModuleName HostedTestDesktop
        Mock Get-RSHostedWslProcess { $global:rsWslQuery } -ModuleName HostedTestDesktop
    }
    AfterEach {Remove-Variable rsWslSurfaceQueue,rsWslCloseCount,rsWslQuery -Scope Global -ErrorAction SilentlyContinue}
    It 'records authorization before the one approved close and then observes shell' {
        $dir=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('wsl-audit-'+[guid]::NewGuid().ToString('N'))
        foreach($kind in @('WslStartup','Shell')) {
            $surface=[pscustomobject]@{Kind=$kind;Window=123;Thread=1;Process=42;Session=3;Title='fixture';Class='fixture';Image='C:\Windows\System32\wsl.exe';ApplicationId='';ApplicationIdStatus=15703;ForegroundDesktop='';ForegroundDesktopError=5;CloseRequested=$false;WslArgumentCount=0;WslArgumentMatch=$false;WslCreationUtcTicks=0;WslCreationMatch=$false}
            $surface | Add-Member ScriptMethod AuthorizeWsl {param($command,$ticks) $this.WslArgumentCount=4;$this.WslArgumentMatch=$true;$this.WslCreationUtcTicks=$ticks;$this.WslCreationMatch=$true}
            $surface | Add-Member ScriptMethod RequestClose {$this.CloseRequested=$true;$global:rsWslCloseCount++}
            $surface | Add-Member ScriptMethod Dispose {}
            $global:rsWslSurfaceQueue.Enqueue($surface)
        }
        & (Get-Module HostedTestDesktop) {param($d) Invoke-RSHostedStockPreparationCore -Directory $d -SourceHead ('a'*40) -ObservationMilliseconds 0} $dir
        $global:rsWslCloseCount | Should Be 1
        $result=Get-Content (Join-Path $dir 'wslstartup-result.json') -Raw|ConvertFrom-Json
        $result.retired | Should Be $true
        $result.argument_count | Should Be 4
        $result.arguments_match | Should Be $true
        $result.creation_match | Should Be $true
        $result.close_requested | Should Be $true
        (Get-Content (Join-Path $dir 'completion.json') -Raw|ConvertFrom-Json).wsl_startup_close_requested | Should Be $true
    }
    It 'retains an argument refusal and never requests a close' {
        $dir=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('wsl-args-refusal-'+[guid]::NewGuid().ToString('N'))
        $surface=[pscustomobject]@{Kind='WslStartup';Window=123;Thread=1;Process=42;Session=3;Title='fixture';Class='fixture';Image='C:\Windows\System32\wsl.exe';ApplicationId='';ApplicationIdStatus=15703;ForegroundDesktop='';ForegroundDesktopError=5;CloseRequested=$false;WslArgumentCount=4;WslArgumentMatch=$false;WslCreationUtcTicks=0;WslCreationMatch=$false}
        $surface | Add-Member ScriptMethod AuthorizeWsl {param($command,$ticks) throw 'WSL startup argument sequence is unavailable or mismatched.'}
        $surface | Add-Member ScriptMethod RequestClose {$this.CloseRequested=$true;$global:rsWslCloseCount++}
        $surface | Add-Member ScriptMethod Dispose {}
        $global:rsWslSurfaceQueue.Enqueue($surface)
        { & (Get-Module HostedTestDesktop) {param($d) Invoke-RSHostedStockPreparationCore -Directory $d -SourceHead ('a'*40) -ObservationMilliseconds 0} $dir } | Should Throw 'WSL startup argument sequence is unavailable or mismatched.'
        $global:rsWslCloseCount | Should Be 0
        $result=Get-Content (Join-Path $dir 'wslstartup-result.json') -Raw|ConvertFrom-Json
        $result.argument_count | Should Be 4
        $result.arguments_match | Should Be $false
        $result.close_requested | Should Be $false
    }
    It 'records a CIM identity refusal and does not request a close' {
        $dir=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('wsl-refusal-'+[guid]::NewGuid().ToString('N'))
        $surface=[pscustomobject]@{Kind='WslStartup';Window=123;Thread=1;Process=42;Session=3;Title='fixture';Class='fixture';Image='C:\Windows\System32\wsl.exe';ApplicationId='';ApplicationIdStatus=15703;ForegroundDesktop='';ForegroundDesktopError=5;CloseRequested=$false;WslArgumentCount=0;WslArgumentMatch=$false;WslCreationUtcTicks=0;WslCreationMatch=$false}
        $surface | Add-Member ScriptMethod RequestClose {$this.CloseRequested=$true;$global:rsWslCloseCount++}
        $surface | Add-Member ScriptMethod Dispose {}
        $global:rsWslSurfaceQueue.Enqueue($surface);$global:rsWslQuery.ProcessId=43
        { & (Get-Module HostedTestDesktop) {param($d) Invoke-RSHostedStockPreparationCore -Directory $d -SourceHead ('a'*40) -ObservationMilliseconds 0} $dir } | Should Throw 'WSL CIM process identity is unavailable or mismatched.'
        $global:rsWslCloseCount | Should Be 0
        $result=Get-Content (Join-Path $dir 'wslstartup-result.json') -Raw|ConvertFrom-Json
        $result.retired | Should Be $false
        $result.arguments_match | Should Be $false
        $result.close_requested | Should Be $false
        (Get-Content (Join-Path $dir 'completion.json') -Raw|ConvertFrom-Json).wsl_startup_close_requested | Should Be $false
    }
}

Describe 'HostedTestDesktop AggregateException and filtered case run verification' {
    It 'preserves both action and restoration exceptions as AggregateException in Invoke-RSWithHostedTestDesktop' {
        $priorGitHubActions = [Environment]::GetEnvironmentVariable('GITHUB_ACTIONS', 'Process')
        $env:GITHUB_ACTIONS = 'true'
        $global:rsHostedAggQueue = [Collections.Generic.Queue[object]]::new()
        try {
            $smallMode = [pscustomobject]@{ Width = 1024; Height = 768; BitsPerPixel = 32; Frequency = 60 }
            $largeMode = [pscustomobject]@{ Width = 1920; Height = 1080; BitsPerPixel = 32; Frequency = 60 }
            $small = [pscustomobject]@{ Mode = $smallMode; Width = 1024; Height = 768; WorkWidth = 1024; WorkHeight = 728; CoordinateScale = 1.0 }
            $large = [pscustomobject]@{ Mode = $largeMode; Width = 1920; Height = 1080; WorkWidth = 1920; WorkHeight = 1040; CoordinateScale = 1.0 }
            $global:rsHostedAggQueue.Enqueue($small)
            $global:rsHostedAggQueue.Enqueue($large)
            $global:rsHostedAggMode = $largeMode
            Mock Get-RSHostedDesktopState { $global:rsHostedAggQueue.Dequeue() } -ModuleName HostedTestDesktop
            Mock Get-RSHostedDisplayModes { $global:rsHostedAggMode } -ModuleName HostedTestDesktop
            Mock Set-RSHostedDisplayMode {
                if ($Mode.Width -eq 1024) { throw [InvalidOperationException]::new('seeded display restore failure') }
            } -ModuleName HostedTestDesktop

            $caught = $null
            try {
                Invoke-RSWithHostedTestDesktop -Action { throw [InvalidOperationException]::new('seeded action failure') }
            } catch {
                $caught = $_.Exception
            }
            ($null -ne $caught) | Should Be $true
            ($caught -is [System.AggregateException]) | Should Be $true
            $caught.InnerExceptions.Count | Should Be 2
            $caught.InnerExceptions[0].Message | Should Match 'seeded action failure'
            $caught.InnerExceptions[1].Message | Should Match 'seeded display restore failure'
        } finally {
            [Environment]::SetEnvironmentVariable('GITHUB_ACTIONS', $priorGitHubActions, 'Process')
            Remove-Variable rsHostedAggQueue, rsHostedAggMode -Scope Global -ErrorAction SilentlyContinue
        }
    }

    It 'validates writer-faithful v1 Commands proof and rejects stale ambiguous or mismatched joins' {
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/HostedDiagnosticEvidence.psm1') -ErrorAction Stop
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestSuitePlan.psm1') -ErrorAction Stop
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestRunPlan.Common.psm1') -ErrorAction Stop
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/ValidationFingerprint.psm1') -ErrorAction Stop
        Mock Read-RSBuildReceipt { [pscustomobject]@{
            receipt_id=('b'*64);repo_identity=('c'*64);source_snapshot_id=('a'*64)
            platform='ARM64';configuration='Debug';target='Build';tests_enabled=$true
            toolchain_identity_digest=('e'*64);vcpkg_identity_digest=('e'*64);ghostty_runtime_identity=('e'*64)
            project_graph_digest=('e'*64);runtime_dependency_digest=('e'*64);outputs=@()
        } } -ModuleName HostedDiagnosticEvidence
        Mock Get-RSBuildSourceSnapshot { [pscustomobject]@{SnapshotId=('a'*64);GitHead=('d'*40)} } -ModuleName HostedDiagnosticEvidence
        Mock Test-RSBuildReceiptForArtifactUse { $true } -ModuleName HostedDiagnosticEvidence
        Mock Assert-RSTestSandboxContainedPath { param($TestRoot,$Path,[switch]$RequireDescendant) [IO.Path]::GetFullPath($Path) } -ModuleName HostedDiagnosticEvidence
        $sandbox = Join-Path $TestDrive 'v1-proof'
        $runId = '20261113T120000Z-1234-' + ('a'*32)
        $run = Join-Path $sandbox "evidence/completed-runs/$runId"
        $last = Join-Path $run 'artifacts/selftest/last_run'
        $pre = Join-Path $run 'artifacts/selftest/previous_run'
        $records = Join-Path $sandbox "evidence/run-archive-records/$runId"
        [void][IO.Directory]::CreateDirectory($pre)
        [void][IO.Directory]::CreateDirectory((Join-Path $last 'commands'))
        [void][IO.Directory]::CreateDirectory((Join-Path $last 'perf'))
        [void][IO.Directory]::CreateDirectory((Join-Path $run 'artifacts/native-process-exits'))
        [void][IO.Directory]::CreateDirectory($records)
        $exe = [IO.Path]::GetFullPath((Join-Path $repoRoot '.build/ARM64/Debug/RedSalamander.exe'))
        $plan = New-RSValidationPlan -RepoRoot $repoRoot -RequestedSuite Commands -ValidationMode Fresh -CoverageScope filtered `
            -Entries @(Get-RSTestRunPlan -Suite Commands -RepoRoot $repoRoot -Platform ARM64 -Configuration Debug -RedSalamanderExePath $exe -CaseFilter terminal_)
        $fixtureReceipt = & (Get-Module HostedDiagnosticEvidence) { Read-RSBuildReceipt -BuildOutputDir 'synthetic' -SchemaPath 'synthetic' }
        $row = Get-RSValidationEntryFingerprint -EntryContract $plan.entries[0] -CoveragePlanDigest $plan.coverage_plan_digest `
            -WorkspaceSnapshot @{snapshot_id=('a'*64)} -BuildReceipt $fixtureReceipt
        $fp = [ordered]@{schema='red-salamander.validation-entry-fingerprints.v1';canonicalization='rs-canonical-json-v1';
            plan_digest=$plan.plan_digest;coverage_plan_digest=$plan.coverage_plan_digest;snapshot_id=('a'*64);build_receipt_id=('b'*64);artifact_epoch=0;
            entries=@(@{id='selftest.commands';fingerprint=$row.Fingerprint;component_digest_version=$row.ComponentDigestVersion;component_digests=$row.ComponentDigests;computation_ms=0})}
        $fp.fingerprints_digest=Get-RSCanonicalJsonDigest $fp
        $fp.computation_ms=0
        $decision=[ordered]@{schema='red-salamander.validation-plan-decision.v1';canonicalization='rs-canonical-json-v1';planner_mode='shadow';
            snapshot_id=('a'*64);plan_digest=$plan.plan_digest;changed_paths=@();matched_impact_rules=@();advisory_build_action='none';
            entries=@(@{id='selftest.commands';ordinal=0;affected=$true;matched_impact_tags=@();decision='EXECUTE';reason_codes=@('FORCED');expected_fingerprint=$row.Fingerprint;origin_run_id=$null;origin_evidence_digest=$null});
            estimated_fresh_duration_ms=0;estimated_executed_duration_ms=0;estimated_avoided_duration_ms=0}
        $decision.decision_digest=Get-RSCanonicalJsonDigest $decision
        function Write-CanonicalProof($Path,$Value,$Schema) {
            Write-RSAtomicCanonicalJsonFile -LiteralPath $Path -InputObject $Value -AllowedRoot $sandbox -SchemaPath (Join-Path $repoRoot "Specs/Testing/$Schema")
        }
        Write-CanonicalProof (Join-Path $pre 'entry-fingerprints.json') $fp 'ValidationEntryFingerprints.schema.json'
        Write-CanonicalProof (Join-Path $pre 'validation-plan.json') $plan 'ValidationPlan.schema.json'
        Write-CanonicalProof (Join-Path $pre 'validation-plan-decision.json') $decision 'ValidationPlanDecision.schema.json'
        $snap=[ordered]@{schema='red-salamander.workspace-snapshot.v1';canonicalization='rs-canonical-json-v1';snapshot_id=('a'*64);repo_identity=('c'*64);git_head=('d'*40);base=('d'*40);files=@();
            acquisition_before=('a'*64);acquisition_after=('a'*64);acquisition_attempts=1;git_discovery_ms=0;content_hashing_ms=0;canonicalization_ms=0;computation_ms=0;created_utc='2026-11-13T12:00:02Z'}
        Write-CanonicalProof (Join-Path $pre 'workspace-initial.json') $snap 'WorkspaceSnapshot.schema.json'
        $snap.created_utc='2026-11-13T12:00:54.500Z'
        Write-CanonicalProof (Join-Path $last 'workspace-after-selftest.commands.json') $snap 'WorkspaceSnapshot.schema.json'
        $snap.created_utc='2026-11-13T12:00:54.800Z'
        Write-CanonicalProof (Join-Path $last 'workspace-final.json') $snap 'WorkspaceSnapshot.schema.json'
        $native=@{run_started_utc='2026-11-13T12:00:06Z';suite='Commands';passed=2;failed=0;skipped=0;cases=@(@{name='terminal_one';status='passed'},@{name='terminal_two';status='passed'})}
        $nativePath=Join-Path $last 'commands/results.json'
        $native|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $nativePath
        $exitPath=Join-Path $run ('artifacts/native-process-exits/'+('f'*32)+'.json')
        @{schema='red-salamander.native-process-exit.v1';process_id=42;executable=$exe;observed_utc='2026-11-13T12:00:54Z';exit_code=0}|ConvertTo-Json|Set-Content -LiteralPath $exitPath
        $metricPath=Join-Path $last 'perf/perf_metrics.jsonl'
        '{"process":42,"timestamp":"2026-11-13T12:00:50Z"}'|Set-Content -LiteralPath $metricPath
        $summary=@{schema='red-salamander.run-all-tests.v1';run_id=$runId;suite='Commands';case_filter='terminal_';platform='ARM64';configuration='Debug';executable=$exe;
            run_started_utc='2026-11-13T12:00:05Z';run_ended_utc='2026-11-13T12:00:55Z';exit_code=2;total=2;passed=2;failed=0;skipped=0;
            quarantine=@{is_valid=$true;has_blocking_entries=$false;active_count=0;active_entries=@();invalid_count=0;invalid_entries=@();errors=@()};
            test_sandbox_audit=@{is_clean=$false;issue_count=1};suites=@(@{suite='Commands';kind='SelfTest';classification='PASSED';exit_code=0;passed=2;failed=0;skipped=0;retry_attempts=@()})}
        $summaryPath=Join-Path $last 'run-all-tests-results.json'
        $summary|ConvertTo-Json -Depth 12|Set-Content -LiteralPath $summaryPath
        function Seal-Proof {
            $inventory=@(Get-ChildItem -LiteralPath $run -File -Recurse|Where-Object FullName -ne $summaryPath|ForEach-Object {
                @{path=[IO.Path]::GetRelativePath($run,$_.FullName);directory=$false;length=$_.Length;sha256=(Get-FileHash -LiteralPath $_.FullName).Hash}
            })
            $intent=@{schema='red-salamander.run-archive.v1';run_id=$runId;disposition='completed';source=(Join-Path $sandbox "runs/$runId");destination=$run;inventory=$inventory}
            $intentPath=Join-Path $records 'intent.json';$intent|ConvertTo-Json -Depth 12|Set-Content -LiteralPath $intentPath
            @{schema='red-salamander.run-archive-complete.v1';source=$intent.source;destination=$run;intent_sha256=(Get-FileHash $intentPath).Hash}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $records 'complete.json')
        }
        $invoke={Assert-RSHostedFilteredCaseRun -ExitCode 2 -Output "Run id: $runId" -TestRoot $sandbox -RepoRoot $repoRoot -Suite Commands -CaseFilter terminal_ -Platform ARM64 -Configuration Debug -InvocationStartedUtc ([DateTimeOffset]'2026-11-13T12:00:00Z') -InvocationEndedUtc ([DateTimeOffset]'2026-11-13T12:01:00Z')}
        Seal-Proof
        (& $invoke).RepositoryVerdict | Should Be 'NOT_EVALUATED'
        # Mutations are resealed where appropriate so semantic joins, not just byte checks, must reject them.
        $original=[IO.File]::ReadAllText($metricPath)
        foreach($control in @(
            @('{"process":43,"timestamp":"2026-11-13T12:00:50Z"}','PID'),
            @('{"process":42,"timestamp":"2026-11-13T12:00:56Z"}','outside the invocation window'),
            @('not json','JSON'), @('','empty'))) {
            [IO.File]::WriteAllText($metricPath,$control[0]);Seal-Proof;{& $invoke}|Should Throw $control[1]
        }
        [IO.File]::WriteAllText($metricPath,$original)
        $exitText=[IO.File]::ReadAllText($exitPath)
        [IO.File]::WriteAllText($exitPath,$exitText.Replace('"exit_code": 0','"exit_code": 9'));Seal-Proof;{& $invoke}|Should Throw 'exit code'
        [IO.File]::WriteAllText($exitPath,$exitText)
        $fpPath=Join-Path $pre 'entry-fingerprints.json';$fpText=[IO.File]::ReadAllText($fpPath)
        foreach($bad in @($fpText.Replace(('a'*64),('9'*64)),$fpText.Replace(('b'*64),('8'*64)))) {
            [IO.File]::WriteAllText($fpPath,$bad);Seal-Proof;{& $invoke}|Should Throw 'source or receipt'
        }
        [IO.File]::WriteAllText($fpPath,$fpText)
        Remove-Item -LiteralPath $fpPath;Seal-Proof;{& $invoke}|Should Throw 'Missing'
        [IO.File]::WriteAllText($fpPath,$fpText)
        Copy-Item -LiteralPath $fpPath -Destination (Join-Path $last 'entry-fingerprints.json');Seal-Proof;{& $invoke}|Should Throw 'ambiguous'
        Remove-Item -LiteralPath (Join-Path $last 'entry-fingerprints.json')
        $planPath=Join-Path $pre 'validation-plan.json';$planText=[IO.File]::ReadAllText($planPath)
        [IO.File]::WriteAllText($planPath,$planText.Replace('terminal_','other_'));Seal-Proof;{& $invoke}|Should Throw 'plan, profile'
        [IO.File]::WriteAllText($planPath,$planText)
        $summary.configuration='Release';$summary|ConvertTo-Json -Depth 12|Set-Content $summaryPath;{& $invoke}|Should Throw 'does not match'
        $summary.configuration='Debug';$summary|ConvertTo-Json -Depth 12|Set-Content $summaryPath
        $postPath=Join-Path $last 'workspace-final.json';$postText=[IO.File]::ReadAllText($postPath)
        [IO.File]::WriteAllText($postPath,$postText.Replace(('a'*64),('9'*64)));Seal-Proof;{& $invoke}|Should Throw 'source identity changed'
        [IO.File]::WriteAllText($postPath,$postText)
        $native.cases[0].name='outside_terminal_suffix';$native|ConvertTo-Json -Depth 8|Set-Content $nativePath;Seal-Proof;{& $invoke}|Should Throw 'unexpected'
        $native.cases[0].name='terminal_one';$native|ConvertTo-Json -Depth 8|Set-Content $nativePath;Seal-Proof
        Mock Test-RSBuildReceiptForArtifactUse { $false } -ModuleName HostedDiagnosticEvidence
        {& $invoke}|Should Throw 'artifact receipt mismatch'
        Mock Test-RSBuildReceiptForArtifactUse { $true } -ModuleName HostedDiagnosticEvidence
        (& $invoke).Passed | Should Be 2
    }

    It 'validates archived filtered case runs in Assert-RSHostedFilteredCaseRun for exit 2 NOT_EVALUATED and rejects failures' {
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestSandbox.psm1') -ErrorAction Stop
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/HostedDiagnosticEvidence.psm1') -ErrorAction Stop
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/ValidationFingerprint.psm1') -ErrorAction Stop
        Mock Read-RSBuildReceipt { [pscustomobject]@{receipt_id=('b'*64)} } -ModuleName HostedDiagnosticEvidence
        Mock Get-RSBuildSourceSnapshot { [pscustomobject]@{SnapshotId=('a'*64)} } -ModuleName HostedDiagnosticEvidence
        Mock Test-RSBuildReceiptForArtifactUse { $true } -ModuleName HostedDiagnosticEvidence
        $sandboxRoot = Join-Path $TestDrive ('filtered-case-' + [guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $sandboxRoot -Force
        Mock Assert-RSTestSandboxContainedPath {
            param($TestRoot, $Path, [switch]$RequireDescendant)
            return [IO.Path]::GetFullPath($Path)
        } -ModuleName HostedDiagnosticEvidence

        $runId = '20261113T120000Z-1234-' + ('a' * 32)
        $runDir = Join-Path $sandboxRoot "evidence/completed-runs/$runId"
        $lastRunDir = Join-Path $runDir 'artifacts/selftest/last_run'
        $recordsDir = Join-Path $sandboxRoot "evidence/run-archive-records/$runId"
        $null = New-Item -ItemType Directory -Path $lastRunDir, $recordsDir -Force

        $startedUtc = [DateTimeOffset]::Parse('2026-11-13T12:00:00Z')
        $endedUtc = [DateTimeOffset]::Parse('2026-11-13T12:01:00Z')
        $v2 = [ordered]@{
            schema = 'red-salamander.run-all-tests.v2'
            run_id = $runId
            validation_mode = 'Fresh'
            requested_gate = 'Commands'
            coverage_scope = 'filtered'
            snapshot_id = ('a'*64); build_receipt_id = ('b'*64)
            plan_digest = ('a'*64); coverage_plan_digest = ('a'*64); decision_digest = ('a'*64)
            duration_ms = 1; estimated_avoided_duration_ms = 0; commands_broad_seal_state = 'absent'
            entries = @(@{
                id='selftest.commands';execution='executed';promotion_state='promoted';evidence_digest=('a'*64)
                promotion_digest=('a'*64);origin_run_id=$null;artifact_epoch=0;artifact_read_ids=@();artifact_write_ids=@()
                produces_artifact_ids=@();cache_policy='ExactArtifact';reason_codes=@();coverage_digest=('a'*64)
                outcome_digest=('a'*64);capability_digest=$null;duration_ms=1
            })
            repository_verdict = 'NOT_EVALUATED'
            change_verdict = 'PASSED'
            counts = [ordered]@{ executed = 1; reused = 0; provisional = 0; promoted = 1; invalidated = 0; incomplete = 0 }
            case_counts = [ordered]@{ passed = 3; failed = 0; skipped = 0; total = 3 }
        }
        $v1 = [ordered]@{
            schema = 'red-salamander.run-all-tests.v1'
            run_id = $runId
            run_started_utc = '2026-11-13T12:00:05.0000000Z'
            run_ended_utc = '2026-11-13T12:00:55.0000000Z'
            suite = 'Commands'
            case_filter = 'terminal_'
            platform = 'ARM64'
            configuration = 'Debug'
            exit_code = 2
            executable = [IO.Path]::GetFullPath((Join-Path $repoRoot '.build/ARM64/Debug/RedSalamander.exe'))
            total = 3
            passed = 3
            failed = 0
            skipped = 0
            quarantine = [ordered]@{
                is_valid = $true; has_blocking_entries = $false
                active_count = 0; active_entries = @()
                invalid_count = 0; invalid_entries = @(); errors = @()
            }
        }
        Write-RSAtomicCanonicalJsonFile -LiteralPath (Join-Path $lastRunDir 'run-all-tests-results.json') -InputObject $v2 -AllowedRoot $sandboxRoot -SchemaPath (Join-Path $repoRoot 'Specs/Testing/RunAllTestsSummary.schema.json')
        [IO.File]::WriteAllText((Join-Path $lastRunDir 'run-all-tests-results-v1.json'), ($v1 | ConvertTo-Json -Depth 10))
        [void][IO.Directory]::CreateDirectory((Join-Path $lastRunDir 'commands'))
        $nativePath = Join-Path $lastRunDir 'commands/results.json'
        $native = @{suite='Commands';passed=3;failed=0;skipped=0;cases=@(
            @{name='terminal_one';status='passed'},@{name='terminal_two';status='passed'},@{name='terminal_three';status='passed'})}
        [IO.File]::WriteAllText($nativePath, ($native | ConvertTo-Json -Depth 10))

        $intentPath = Join-Path $recordsDir 'intent.json'
        $intent = [ordered]@{
            schema = 'red-salamander.run-archive.v1'
            run_id = $runId
            disposition = 'completed'
            source = [IO.Path]::GetFullPath((Join-Path $sandboxRoot "runs/$runId"))
            destination = [IO.Path]::GetFullPath($runDir)
            inventory = @(@{path='artifacts\selftest\last_run\commands\results.json';directory=$false;
                length=(Get-Item $nativePath).Length;sha256=(Get-FileHash $nativePath).Hash})
        }
        [IO.File]::WriteAllText($intentPath, ($intent | ConvertTo-Json -Depth 6))
        $intentHash = (Get-FileHash -LiteralPath $intentPath -Algorithm SHA256).Hash
        $complete = [ordered]@{
            schema = 'red-salamander.run-archive-complete.v1'
            source = [IO.Path]::GetFullPath((Join-Path $sandboxRoot "runs/$runId"))
            destination = [IO.Path]::GetFullPath($runDir)
            intent_sha256 = $intentHash
        }
        [IO.File]::WriteAllText((Join-Path $recordsDir 'complete.json'), ($complete | ConvertTo-Json -Depth 6))

        $outputText = "Header`nRun id: $runId`nDone"
        $verified = Assert-RSHostedFilteredCaseRun `
            -ExitCode 2 -Output $outputText -TestRoot $sandboxRoot -RepoRoot $repoRoot -Suite 'Commands' `
            -CaseFilter 'terminal_' -Platform 'ARM64' -Configuration 'Debug' `
            -InvocationStartedUtc $startedUtc -InvocationEndedUtc $endedUtc
        $verified.RunId | Should Be $runId
        $verified.RepositoryVerdict | Should Be 'NOT_EVALUATED'
        $verified.ChangeVerdict | Should Be 'PASSED'

        # Reject mismatched case filter, failed case count, or exit 1
        { Assert-RSHostedFilteredCaseRun -ExitCode 2 -Output $outputText -TestRoot $sandboxRoot -RepoRoot $repoRoot -Suite 'Commands' -CaseFilter 'other_' -Platform 'ARM64' -Configuration 'Debug' -InvocationStartedUtc $startedUtc -InvocationEndedUtc $endedUtc } | Should Throw 'does not match'
        { Assert-RSHostedFilteredCaseRun -ExitCode 1 -Output $outputText -TestRoot $sandboxRoot -RepoRoot $repoRoot -Suite 'Commands' -CaseFilter 'terminal_' -Platform 'ARM64' -Configuration 'Debug' -InvocationStartedUtc $startedUtc -InvocationEndedUtc $endedUtc } | Should Throw 'failed with exit code 1'
        { Assert-RSHostedFilteredCaseRun -ExitCode 0 -Output $outputText -TestRoot $sandboxRoot -RepoRoot $repoRoot -Suite 'Commands' -CaseFilter 'terminal_' -Platform 'ARM64' -Configuration 'Debug' -InvocationStartedUtc $startedUtc -InvocationEndedUtc $endedUtc } | Should Throw 'failed with exit code 0'

        $v2.case_counts.failed = 1
        Write-RSAtomicCanonicalJsonFile -LiteralPath (Join-Path $lastRunDir 'run-all-tests-results.json') -InputObject $v2 -AllowedRoot $sandboxRoot -SchemaPath (Join-Path $repoRoot 'Specs/Testing/RunAllTestsSummary.schema.json')
        { Assert-RSHostedFilteredCaseRun -ExitCode 2 -Output $outputText -TestRoot $sandboxRoot -RepoRoot $repoRoot -Suite 'Commands' -CaseFilter 'terminal_' -Platform 'ARM64' -Configuration 'Debug' -InvocationStartedUtc $startedUtc -InvocationEndedUtc $endedUtc } | Should Throw 'incomplete, skipped, or failed'

        $v2.case_counts.failed = 0
        $v2.build_receipt_id = 'c'*64
        Write-RSAtomicCanonicalJsonFile -LiteralPath (Join-Path $lastRunDir 'run-all-tests-results.json') -InputObject $v2 -AllowedRoot $sandboxRoot -SchemaPath (Join-Path $repoRoot 'Specs/Testing/RunAllTestsSummary.schema.json')
        { Assert-RSHostedFilteredCaseRun -ExitCode 2 -Output $outputText -TestRoot $sandboxRoot -RepoRoot $repoRoot -Suite 'Commands' -CaseFilter 'terminal_' -Platform 'ARM64' -Configuration 'Debug' -InvocationStartedUtc $startedUtc -InvocationEndedUtc $endedUtc } | Should Throw 'artifact receipt mismatch'
        $v2.build_receipt_id = 'b'*64
        Write-RSAtomicCanonicalJsonFile -LiteralPath (Join-Path $lastRunDir 'run-all-tests-results.json') -InputObject $v2 -AllowedRoot $sandboxRoot -SchemaPath (Join-Path $repoRoot 'Specs/Testing/RunAllTestsSummary.schema.json')
        [IO.File]::AppendAllText($nativePath, ' ')
        { Assert-RSHostedFilteredCaseRun -ExitCode 2 -Output $outputText -TestRoot $sandboxRoot -RepoRoot $repoRoot -Suite 'Commands' -CaseFilter 'terminal_' -Platform 'ARM64' -Configuration 'Debug' -InvocationStartedUtc $startedUtc -InvocationEndedUtc $endedUtc } | Should Throw 'does not match the archive'

        # Re-seal a coherent archive containing an unselected case: valid bytes and
        # aggregate counts alone must not authorize the requested filter.
        $native.cases[0].name = 'unselected_case'
        [IO.File]::WriteAllText($nativePath, ($native | ConvertTo-Json -Depth 10))
        $intent.inventory[0].length = (Get-Item $nativePath).Length
        $intent.inventory[0].sha256 = (Get-FileHash $nativePath).Hash
        [IO.File]::WriteAllText($intentPath, ($intent | ConvertTo-Json -Depth 10))
        $complete.intent_sha256 = (Get-FileHash $intentPath).Hash
        [IO.File]::WriteAllText((Join-Path $recordsDir 'complete.json'), ($complete | ConvertTo-Json -Depth 10))
        { Assert-RSHostedFilteredCaseRun -ExitCode 2 -Output $outputText -TestRoot $sandboxRoot -RepoRoot $repoRoot -Suite 'Commands' -CaseFilter 'terminal_' -Platform 'ARM64' -Configuration 'Debug' -InvocationStartedUtc $startedUtc -InvocationEndedUtc $endedUtc } | Should Throw 'unexpected, duplicated, or not passing'

        $intent.inventory = @()
        [IO.File]::WriteAllText($intentPath, ($intent | ConvertTo-Json -Depth 10))
        $complete.intent_sha256 = (Get-FileHash $intentPath).Hash
        [IO.File]::WriteAllText((Join-Path $recordsDir 'complete.json'), ($complete | ConvertTo-Json -Depth 10))
        { Assert-RSHostedFilteredCaseRun -ExitCode 2 -Output $outputText -TestRoot $sandboxRoot -RepoRoot $repoRoot -Suite 'Commands' -CaseFilter 'terminal_' -Platform 'ARM64' -Configuration 'Debug' -InvocationStartedUtc $startedUtc -InvocationEndedUtc $endedUtc } | Should Throw 'does not match the archive'
    }
}
