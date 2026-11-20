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
        Mock Get-RSHostedForegroundQueueDiagnostics { [pscustomobject]@{ interpretationPermitted=$false; nonAtomic=$true } } -ModuleName HostedTestDesktop
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

Describe 'Hosted foreground queue diagnostic evidence' {
    BeforeAll {
        & (Get-Module HostedTestDesktop) { Initialize-RSHostedDesktopNative }
        $script:queueBoundary = 'RedSalamander.Testing.HostedDesktopNative' -as [type]
        $script:queueBoundary = $script:queueBoundary.Assembly.GetType('RedSalamander.Testing.HostedForegroundBoundary', $true)
        if (-not ('RSHostedForegroundQueueProbe' -as [type])) {
            Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Reflection;
public static class RSHostedForegroundQueueProbe {
    public static object Run(Type boundary, string mode) {
        int identities = 0, classes = 0, ambientError = 0;
        var events = new List<string>();
        var arguments = new List<uint>();
        Func<Tuple<IntPtr,uint,uint>> identity = () => {
            int index = identities++;
            if (index >= 4) throw new InvalidOperationException("Extra identity read.");
            events.Add("identity-" + index);
            ambientError = 999;
            long hwnd = 17; uint thread = 7, process = 42;
            if (mode == "null") { hwnd = 0; thread = 0; process = 0; }
            if (mode == "zero-thread") thread = 0;
            if (mode == "zero-process") process = 0;
            if (mode == "window-" + index) hwnd = 18;
            if (mode == "thread-" + index) thread = 8;
            if (mode == "process-" + index) process = 43;
            return Tuple.Create(new IntPtr(hwnd), thread, process);
        };
        Func<IntPtr,Tuple<string,int>> readClass = window => {
            if (++classes != 1) throw new InvalidOperationException("Extra class read.");
            events.Add("class");
            if (mode == "class-failure" || window == IntPtr.Zero) return Tuple.Create<string,int>(null, 5);
            return Tuple.Create(mode == "non-console" ? "OwnedFixtureClass" : "ConsoleWindowClass", 0);
        };
        Func<uint,object> query = thread => {
            arguments.Add(thread); events.Add("query-" + thread);
            bool success = mode != "explicit-87" || thread == 0;
            return new Dictionary<string,object> {
                { "requestedThread", thread }, { "inputSize", 72u }, { "attempted", true }, { "success", success },
                { "savedError", success ? null : (object)87 },
                { "outputs", success ? new Dictionary<string,object>{{"active","0x11"},{"focus","0x12"}} : null }
            };
        };
        var method = boundary.GetMethod("ObserveForegroundQueueForDiagnostics", BindingFlags.Static | BindingFlags.NonPublic);
        object facts = method.Invoke(null, new object[]{identity, query, readClass});
        return new Dictionary<string,object>{{"facts",facts},{"events",events.ToArray()},{"arguments",arguments.ToArray()},
            {"identities",identities},{"classes",classes},{"ambientError",ambientError}};
    }
}
'@
        }
    }
    It 'uses exactly four identity samples and one in-bracket class read around the explicit and zero queries' {
        $r = [RSHostedForegroundQueueProbe]::Run($script:queueBoundary, 'stable')
        ($r.events -join ',') | Should Be 'identity-0,class,query-7,identity-1,identity-2,query-0,identity-3'
        ($r.arguments -join ',') | Should Be '7,0'
        $r.identities | Should Be 4
        $r.classes | Should Be 1
        $r.facts.fourSampleIdentityStable | Should Be $true
        $r.facts.interpretationPermitted | Should Be $true
        $r.facts.originalClass.isConsoleWindowClass | Should Be $true
        $r.facts.nonAtomic | Should Be $true
        $r.facts.acceptanceAuthority | Should Be $false
        $r.facts.explicitQuery.outputs.focus | Should Be $r.facts.zeroQuery.outputs.focus
        $r.facts.processArchitecture | Should Be ([Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture.ToString())
        $r.facts.osArchitecture | Should Be ([Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString())
        $r.facts.pointerBytes | Should Be ([IntPtr]::Size)
    }
    It 'refuses interpretation for every changed window thread or process sample without retrying' {
        foreach ($field in @('window','thread','process')) {
            foreach ($index in 1..3) {
                $r = [RSHostedForegroundQueueProbe]::Run($script:queueBoundary, ($field+'-'+$index))
                $r.facts.fourSampleIdentityStable | Should Be $false
                $r.facts.interpretationPermitted | Should Be $false
                $r.facts.originalClass.correlated | Should Be $false
                $r.facts.originalClass.isConsoleWindowClass | Should BeNullOrEmpty
                ($r.arguments -join ',') | Should Be '7,0'
                $r.identities | Should Be 4
            }
        }
    }
    It 'retains missing identities without substituting zero for an unavailable explicit thread' {
        foreach ($mode in @('null','zero-thread','zero-process')) {
            $r = [RSHostedForegroundQueueProbe]::Run($script:queueBoundary, $mode)
            $r.facts.interpretationPermitted | Should Be $false
            $r.identities | Should Be 4
            if ($mode -eq 'null') {
                $r.facts.originalClass.attempted | Should Be $false
                $r.facts.originalClass.savedError | Should BeNullOrEmpty
            }
            if ($mode -ne 'zero-process') {
                ($r.arguments -join ',') | Should Be '0'
                $r.facts.explicitQuery.attempted | Should Be $false
                $r.facts.explicitQuery.savedError | Should BeNullOrEmpty
                $r.facts.explicitQuery.outputs | Should BeNullOrEmpty
            }
        }
    }
    It 'preserves the explicit saved87 and failed-output absence when a later observation changes ambient error' {
        $r = [RSHostedForegroundQueueProbe]::Run($script:queueBoundary, 'explicit-87')
        $r.facts.explicitQuery.success | Should Be $false
        $r.facts.explicitQuery.savedError | Should Be 87
        $r.facts.explicitQuery.outputs | Should BeNullOrEmpty
        $r.ambientError | Should Be 999
        $r.facts.zeroQuery.success | Should Be $true
        $r.facts.zeroQuery.savedError | Should BeNullOrEmpty
        ($r.arguments -join ',') | Should Be '7,0'
        $r.facts.acceptanceAuthority | Should Be $false
    }
    It 'reports actual non-console or unavailable class facts without a console assumption' {
        $r = [RSHostedForegroundQueueProbe]::Run($script:queueBoundary, 'non-console')
        $r.facts.originalClass.value | Should Be 'OwnedFixtureClass'
        $r.facts.originalClass.isConsoleWindowClass | Should Be $false
        $r = [RSHostedForegroundQueueProbe]::Run($script:queueBoundary, 'class-failure')
        $r.facts.originalClass.known | Should Be $false
        $r.facts.originalClass.savedError | Should Be 5
        $r.facts.originalClass.value | Should BeNullOrEmpty
        $r.facts.originalClass.isConsoleWindowClass | Should BeNullOrEmpty
    }
    It 'keeps native query errors immediate outputs conditional and the cold sampler free of authority' {
        $source = [IO.File]::ReadAllText((Join-Path $repoRoot 'Tools/Modules/Testing/HostedTestDesktop.psm1'))
        $start = $source.IndexOf('private static Tuple<IntPtr, uint, uint> ReadForegroundIdentityForDiagnostics()')
        $end = $source.IndexOf('// Failure-only observations cannot authorize', $start)
        $cold = $source.Substring($start, $end-$start)
        $guard = { param($text)
            if ($text -notmatch 'bool success = GetGUIThreadInfo\(thread, ref info\);\s*int error = success \? 0 : Marshal.GetLastWin32Error\(\);') { throw 'GUI error must be saved immediately' }
            if ($text -notmatch 'object outputs = success \?') { throw 'Failed GUI output must remain unavailable' }
            if ($text -notmatch 'new GuiThreadInfo \{ Size = size \}') { throw 'Each query needs a fresh initialized structure' }
            if ($text -notmatch 'new StringBuilder\(256\)') { throw 'Class buffer must stay bounded' }
            if ($text -notmatch 'int length = HostedDesktopNative.GetClassNameW\(window, text, text.Capacity\);\s*int error = length > 0 \? 0 : Marshal.GetLastWin32Error\(\);') { throw 'Class error must be saved immediately' }
            if ($text -match 'SetForegroundWindow|SetFocus|AttachThreadInput|SendInput|CreateWindow|ShowWindow|OpenProcess|OpenThread|OpenDesktop|SwitchDesktop|GetWindowText|EnumWindows|SetCursor|\bwhile\s*\(|\bfor\s*\(') { throw 'Cold diagnostic cannot gain native authority or retry' }
        }
        & $guard $cold
        $counterfactuals = @(
            @{ Text=$cold.Replace('int error = success ? 0 : Marshal.GetLastWin32Error();', 'int error = 0;'); Message='GUI error must be saved immediately' }
            @{ Text=$cold.Replace('object outputs = success ?', 'object outputs = true ?'); Message='Failed GUI output must remain unavailable' }
            @{ Text=($cold + ' SetForegroundWindow(window);'); Message='Cold diagnostic cannot gain native authority or retry' }
        )
        foreach ($counterfactual in $counterfactuals) {
            $rejection = $null
            try { & $guard $counterfactual.Text } catch { $rejection = $_.Exception }
            $rejection | Should Not BeNullOrEmpty
            $rejection.GetType().FullName | Should Be 'System.Management.Automation.RuntimeException'
            $rejection.Message | Should Be $counterfactual.Message
        }
        $cold | Should Not Match 'ReadFocus\('
        $queryMethod = $script:queueBoundary.GetMethod('ReadForegroundGuiForDiagnostics', [Reflection.BindingFlags]'Static,NonPublic')
        $guiType = $script:queueBoundary.GetNestedType('GuiThreadInfo', [Reflection.BindingFlags]::NonPublic)
        [Runtime.InteropServices.Marshal]::SizeOf([type]$guiType) | Should Be $(if ([IntPtr]::Size -eq 8) {72} else {48})
        $queryMethod | Should Not BeNullOrEmpty
    }
}

Describe 'Hosted original foreground focus contract' {
    BeforeAll {
        & (Get-Module HostedTestDesktop) { Initialize-RSHostedDesktopNative }
        $script:primaryBoundary = ('RedSalamander.Testing.HostedDesktopNative' -as [type]).Assembly.GetType('RedSalamander.Testing.HostedForegroundBoundary', $true)
        if (-not ('RSHostedPrimaryForegroundProbe' -as [type])) {
            Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
public static class RSHostedPrimaryForegroundProbe {
    static readonly BindingFlags Flags = BindingFlags.Static | BindingFlags.NonPublic;
    static object Call(Type boundary, string name, params object[] args) { return boundary.GetMethod(name, Flags).Invoke(null, args); }
    static object Field(object value, string name) { return value.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(value); }
    static object Capture(Type boundary, string mode, List<uint> queries, List<string> events,
        Func<Tuple<IntPtr,uint,uint>> externalForeground = null) {
        int foregroundReads = 0, originalReads = 0, focusReads = 0;
        Func<Tuple<IntPtr,uint,uint>> foreground = externalForeground ?? (() => {
            int index = foregroundReads++;
            events.Add("foreground-" + index);
            long h = 17; uint t = 7, p = 42;
            if (index > 0 && mode == "changed-foreground") h = 19;
            if (index > 0 && mode == "changed-foreground-thread") t = 8;
            if (index > 0 && mode == "changed-foreground-process") p = 43;
            if (mode == "not-original") h = 51;
            return Tuple.Create(new IntPtr(h),t,p);
        });
        Func<uint,Tuple<IntPtr,IntPtr>> query = requested => {
            uint actual = mode == "explicit-binding-counterfactual" ? 7u : requested;
            queries.Add(actual); events.Add("gui-" + actual);
            if (actual != 0 || mode == "failed-zero") {
                var saved = new Win32Exception(87, "controlled original GUI query failed");
                Marshal.SetLastPInvokeError(999);
                throw saved;
            }
            if (mode == "fatal-zero") throw new OutOfMemoryException("synthetic fatal query control");
            long active = mode == "wrong-active" ? 19 : mode == "null-active" ? 0 : 17;
            long focus = mode == "null-focus" || mode == "null-active" ? 0 : mode == "focus-original" ? 17 : mode == "other-focus" ? 20 : 18;
            return Tuple.Create(new IntPtr(active),new IntPtr(focus));
        };
        Func<IntPtr,Tuple<uint,uint,IntPtr,bool>> identity = window => {
            bool original = window.ToInt64() == 17;
            int index = original ? originalReads++ : focusReads++;
            uint t = 7, p = 42; long root = 17; bool valid = true;
            if (!original && mode == "cross-owner") { t = 8; p = 43; }
            if (original && index > 0 && mode == "changed-original-thread") t = 8;
            if (original && index > 0 && mode == "changed-original-process") p = 43;
            if (!original && index > 0 && mode == "changed-focus-thread") t = 8;
            if (!original && index > 0 && mode == "changed-focus-process") p = 43;
            if (!original && mode == "wrong-root") root = 19;
            if (!original && mode == "unknown-focus") { t = 0; p = 0; valid = false; }
            if (original && mode == "unknown-original") valid = false;
            return Tuple.Create(t,p,new IntPtr(root),valid);
        };
        return Call(boundary,"CaptureOriginalForegroundFocus",new IntPtr(17),7u,42u,foreground,query,identity);
    }
    static Dictionary<string,object> Failure(Exception error) {
        return new Dictionary<string,object>{{"failureType",error == null ? null : error.GetType().FullName},
            {"message",error == null ? null : error.Message},{"nativeError",error is Win32Exception ? (object)((Win32Exception)error).NativeErrorCode : null}};
    }
    public static object CaptureCase(Type boundary, string mode) {
        var queries = new List<uint>(); var events = new List<string>(); object saved = null; Exception error = null;
        try { saved = Capture(boundary,mode,queries,events); }
        catch (TargetInvocationException caught) { error = caught.GetBaseException(); }
        var result = Failure(error); result["queries"] = queries.ToArray(); result["events"] = events.ToArray();
        result["mayWrite"] = saved != null && (bool)Call(boundary,"MayWriteOriginalForeground",saved);
        if (saved != null) {
            result["focus"] = ((IntPtr)Field(saved,"Focus")).ToInt64();
            result["focusThread"] = Field(saved,"FocusThread"); result["focusProcess"] = Field(saved,"FocusProcess");
            result["immutable"] = Array.TrueForAll(saved.GetType().GetFields(BindingFlags.Instance | BindingFlags.NonPublic), f => f.IsInitOnly);
        }
        return result;
    }
    public static object RestoreCase(Type boundary, string mode) {
        var entryQueries = new List<uint>(); var events = new List<string>();
        object saved = Capture(boundary, mode.StartsWith("cross-") ? "cross-owner" : "stable",entryQueries,events);
        if (mode == "empty") {
            var type = saved.GetType();
            saved = Activator.CreateInstance(type, BindingFlags.Instance | BindingFlags.NonPublic, null,
                new object[]{IntPtr.Zero,0u,0u,IntPtr.Zero,0u,0u,IntPtr.Zero},null);
        }
        var queries = new List<uint>(); var writes = new List<long>();
        bool already = mode == "already" || mode == "cross-already" || mode.StartsWith("focus-") || mode == "interference-before-query";
        IntPtr current = new IntPtr(already ? 17 : 51); int waits = 0, validations = 0, foregroundReads = 0, observations = 0;
        bool changedOwner = false; Exception error = null;
        Func<IntPtr> foreground = () => {
            foregroundReads++; IntPtr observed = current;
            if (mode == "interference-before-query" && foregroundReads == 2) current = new IntPtr(68);
            return observed;
        };
        Action<bool> validate = afterWait => {
            validations++;
            if (validations == 2 && mode == "desktop-before-write")
                throw new InvalidOperationException("controlled prewrite Default desktop refusal");
            if (validations == 2 && mode == "interference-before-write") current = new IntPtr(68);
            if (validations == 2 && mode == "same-hwnd-owner-change") changedOwner = true;
            Func<IntPtr,Tuple<uint,uint,IntPtr,bool>> identity = h => Tuple.Create(
                changedOwner && h.ToInt64() == 18 ? 8u : h.ToInt64() == 18 && mode.StartsWith("cross-") ? 8u : 7u,
                h.ToInt64() == 18 && mode.StartsWith("cross-") ? 43u : 42u,new IntPtr(17),true);
            if (!(bool)Call(boundary,"IsSavedOriginalFocusLive",saved,identity))
                throw new InvalidOperationException("controlled saved identity changed");
        };
        Func<object> observe = () => {
            observations++;
            string observationMode = mode.StartsWith("cross-") ? "cross-owner" :
                (mode == "focus-timeout" || (mode == "focus-settle" && observations == 1)) ? "other-focus" : "stable";
            return Capture(boundary, observationMode, queries, events,
                () => Tuple.Create(current, current.ToInt64() == 17 ? 7u : 9u, current.ToInt64() == 17 ? 42u : 99u));
        };
        Func<IntPtr,bool> write = target => { writes.Add(target.ToInt64()); events.Add("write-"+target.ToInt64()); return true; };
        Action<Stopwatch> wait = timer => {
            waits++;
            if (mode == "interference-after-write") current = new IntPtr(68);
            else if (mode.StartsWith("focus-")) current = new IntPtr(17);
            else current = waits == 1 ? IntPtr.Zero : new IntPtr(17);
        };
        Func<Stopwatch,long> elapsed = timer => mode == "foreground-timeout" || mode == "focus-timeout" ? 1500 : 0;
        Func<IntPtr,Exception> foreign = window => new InvalidOperationException("controlled foreign foreground");
        try { Call(boundary,"RestoreOriginalForeground",new IntPtr(51),saved,foreground,validate,observe,write,wait,elapsed,foreign); }
        catch (TargetInvocationException caught) { error = caught.GetBaseException(); }
        var result = Failure(error); result["queries"] = queries.ToArray(); result["writes"] = writes.ToArray();
        result["waits"] = waits; result["observations"] = observations; result["validations"] = validations; result["events"] = events.ToArray();
        result["savedFocus"] = ((IntPtr)Field(saved,"Focus")).ToInt64(); result["savedFocusThread"] = Field(saved,"FocusThread");
        return result;
    }
}
'@
        }
    }
    It 'uses primary zero despite a backend that rejects the HWND-derived thread with87' {
        $r=[RSHostedPrimaryForegroundProbe]::CaptureCase($script:primaryBoundary,'stable')
        $r.failureType | Should BeNullOrEmpty
        ($r.queries -join ',') | Should Be '0'
        $r.focus | Should Be 18
        $r.mayWrite | Should Be $true
        $r.immutable | Should Be $true
        $old=[RSHostedPrimaryForegroundProbe]::CaptureCase($script:primaryBoundary,'explicit-binding-counterfactual')
        $old.failureType | Should Be 'System.ComponentModel.Win32Exception'
        $old.nativeError | Should Be 87
        ($old.queries -join ',') | Should Be '7'
    }
    It 'accepts genuine null focus and a related child without equating focus with foreground' {
        foreach($mode in @('null-focus','focus-original','stable')) {
            $r=[RSHostedPrimaryForegroundProbe]::CaptureCase($script:primaryBoundary,$mode)
            $r.failureType | Should BeNullOrEmpty
            $r.mayWrite | Should Be $true
            ($r.queries -join ',') | Should Be '0'
        }
        [RSHostedPrimaryForegroundProbe]::CaptureCase($script:primaryBoundary,'null-focus').focus | Should Be 0
    }
    It 'refuses each same-HWND owner change wrong root unavailable identity or invalid active witness' {
        foreach($mode in @('changed-foreground','changed-foreground-thread','changed-foreground-process','changed-original-thread','changed-original-process','changed-focus-thread','changed-focus-process','wrong-root','unknown-focus','unknown-original','wrong-active','null-active','not-original')) {
            $r=[RSHostedPrimaryForegroundProbe]::CaptureCase($script:primaryBoundary,$mode)
            $r.failureType | Should Be 'System.InvalidOperationException'
            $r.mayWrite | Should Be $false
            if($mode -in @('not-original','unknown-original')) { @($r.queries).Count | Should Be 0 }
        }
    }
    It 'preserves failed zero87 and fatal query exceptions without manufacturing null focus' {
        $r=[RSHostedPrimaryForegroundProbe]::CaptureCase($script:primaryBoundary,'failed-zero')
        $r.failureType | Should Be 'System.ComponentModel.Win32Exception'
        $r.nativeError | Should Be 87
        $r.mayWrite | Should Be $false
        ($r.queries -join ',') | Should Be '0'
        $r=[RSHostedPrimaryForegroundProbe]::CaptureCase($script:primaryBoundary,'fatal-zero')
        $r.failureType | Should Be 'System.OutOfMemoryException'
    }
    It 'permits related cross-owner observation but refuses write admission before acquisition' {
        $r=[RSHostedPrimaryForegroundProbe]::CaptureCase($script:primaryBoundary,'cross-owner')
        $r.failureType | Should BeNullOrEmpty
        $r.focusThread | Should Be 8
        $r.focusProcess | Should Be 43
        $r.mayWrite | Should Be $false
        $already=[RSHostedPrimaryForegroundProbe]::RestoreCase($script:primaryBoundary,'cross-already')
        $already.failureType | Should BeNullOrEmpty
        @($already.writes).Count | Should Be 0
        $blocked=[RSHostedPrimaryForegroundProbe]::RestoreCase($script:primaryBoundary,'cross-write')
        $blocked.failureType | Should Be 'System.InvalidOperationException'
        @($blocked.writes).Count | Should Be 0
        @($blocked.queries).Count | Should Be 0
    }
    It 'issues one original-HWND request and no GUI queries while owned or null foreground settles' {
        $r=[RSHostedPrimaryForegroundProbe]::RestoreCase($script:primaryBoundary,'settle')
        $r.failureType | Should BeNullOrEmpty
        ($r.writes -join ',') | Should Be '17'
        ($r.queries -join ',') | Should Be '0'
        $r.waits | Should Be 2
        $r.savedFocus | Should Be 18
        $r=[RSHostedPrimaryForegroundProbe]::RestoreCase($script:primaryBoundary,'empty')
        $r.failureType | Should BeNullOrEmpty
        @($r.writes).Count | Should Be 0
        @($r.queries).Count | Should Be 0
    }
    It 'never writes after foreground or same-HWND owner interference and never queries a non-original queue' {
        foreach($mode in @('interference-before-write','same-hwnd-owner-change','interference-before-query')) {
            $r=[RSHostedPrimaryForegroundProbe]::RestoreCase($script:primaryBoundary,$mode)
            $r.failureType | Should Be 'System.InvalidOperationException'
            @($r.writes).Count | Should Be 0
            @($r.queries).Count | Should Be 0
        }
        $r=[RSHostedPrimaryForegroundProbe]::RestoreCase($script:primaryBoundary,'interference-after-write')
        $r.failureType | Should Be 'System.InvalidOperationException'
        ($r.writes -join ',') | Should Be '17'
        @($r.queries).Count | Should Be 0
    }
    It 'keeps both1500ms limits and immutable expected focus while observing a valid focus transition' {
        $r=[RSHostedPrimaryForegroundProbe]::RestoreCase($script:primaryBoundary,'focus-settle')
        $r.failureType | Should BeNullOrEmpty
        @($r.writes).Count | Should Be 0
        ($r.queries -join ',') | Should Be '0,0'
        $r.savedFocus | Should Be 18
        foreach($mode in @('foreground-timeout','focus-timeout')) {
            $r=[RSHostedPrimaryForegroundProbe]::RestoreCase($script:primaryBoundary,$mode)
            $r.failureType | Should Be 'System.InvalidOperationException'
            $r.savedFocus | Should Be 18
        }
    }
    It 'preserves a second prewrite desktop refusal with zero requests and zero GUI queries' {
        $r=[RSHostedPrimaryForegroundProbe]::RestoreCase($script:primaryBoundary,'desktop-before-write')
        $r.failureType | Should Be 'System.InvalidOperationException'
        $r.message | Should Be 'controlled prewrite Default desktop refusal'
        $r.validations | Should Be 2
        @($r.writes).Count | Should Be 0
        @($r.queries).Count | Should Be 0
        $source=[IO.File]::ReadAllText((Join-Path $repoRoot 'Tools/Modules/Testing/HostedTestDesktop.psm1'))
        $guard={param($text)
            if($text -notmatch '!afterWait && HostedDesktopNative.InputDesktopName\(\) != "Default"' -or
               $text -notmatch 'afterWait && !IsVerifiedInheritedDesktop\(HostedDesktopNative.ReadDesktopContext\(0\)\)') {
                throw 'Native prewrite and postwait desktop predicates must remain bound'
            }
        }
        & $guard $source
        $rejection=$null
        try { & $guard ($source.Replace('!afterWait && HostedDesktopNative.InputDesktopName() != "Default"','false')) } catch { $rejection=$_.Exception }
        $rejection.GetType().FullName | Should Be 'System.Management.Automation.RuntimeException'
        $rejection.Message | Should Be 'Native prewrite and postwait desktop predicates must remain bound'
    }
    It 'rejects acceptance-source counterfactuals and retains exact native bindings and owned readiness' {
        $source=[IO.File]::ReadAllText((Join-Path $repoRoot 'Tools/Modules/Testing/HostedTestDesktop.psm1'))
        $guard={param($text)
            $start=$text.IndexOf('private static GuiThreadInfo ReadFocus(')
            $end=$text.IndexOf('// DesktopProbe only:', $start)
            $raw=$text.Substring($start,$end-$start)
            if($text -notmatch 'var gui = query\(0\);' -or $raw -notmatch 'GetGUIThreadInfo\(0, ref info\)' -or $raw -match 'GetGUIThreadInfo\(thread, ref info\)|error\s*==\s*87') {throw 'Primary foreground query required'}
            if($text -notmatch '!MatchesOriginal\(before,' -or $text -notmatch '!MatchesOriginal\(after,' -or $text -notmatch 'gui.Item1 != original') {throw 'Both original brackets and exact active witness required'}
            if($text -notmatch 'saved.FocusThread == saved.OriginalThread && saved.FocusProcess == saved.OriginalProcess') {throw 'Cross-owner write admission forbidden'}
            if($text -match '\bsaved\s*=\s*\(OriginalForegroundFocus\)observeOriginal') {throw 'Expected focus cannot be rebased'}
        }
        & $guard $source
        $mutations=@(
            @{Text=$source.Replace('var gui = query(0);','var gui = query(thread);');Message='Primary foreground query required'}
            @{Text=$source.Replace('!MatchesOriginal(after,','!MatchesOriginal(before,');Message='Both original brackets and exact active witness required'}
            @{Text=$source.Replace('gui.Item1 != original','false');Message='Both original brackets and exact active witness required'}
            @{Text=$source.Replace('saved.FocusThread == saved.OriginalThread && saved.FocusProcess == saved.OriginalProcess','true');Message='Cross-owner write admission forbidden'}
            @{Text=($source+' saved = (OriginalForegroundFocus)observeOriginal();');Message='Expected focus cannot be rebased'}
        )
        foreach($mutation in $mutations){$failure=$null;try{& $guard $mutation.Text}catch{$failure=$_.Exception};$failure.GetType().FullName|Should Be 'System.Management.Automation.RuntimeException';$failure.Message|Should Be $mutation.Message}
        $run=$source.Substring($source.IndexOf('private void Run() {'),$source.IndexOf('public void Dispose()', $source.IndexOf('private void Run() {'))-$source.IndexOf('private void Run() {'))
        $run.IndexOf('if (!MayWriteOriginalForeground(originalFocus))') | Should BeLessThan $run.IndexOf('var window = CreateWindowExW(')
        $source | Should Match 'SetForegroundWindow, WaitOneRestoreMessage, timer => timer.ElapsedMilliseconds'
        $source | Should Match 'GetForegroundWindow\(\) != handle \|\| GetFocus\(\) != handle'
        $source | Should Not Match 'SetFocus\(original|SetFocus\(saved|SetFocus\(gui\.'
    }
}

function Initialize-RSHostedDiagnosticMessageProbe {
    if (-not ('RedSalamander.Testing.HostedDiagnosticMessageProbeException' -as [type])) {
        Add-Type -TypeDefinition @"
namespace RedSalamander.Testing {
    public sealed class HostedDiagnosticMessageProbeException : System.Exception {
        private readonly string label;
        public int Reads { get; private set; }
        public HostedDiagnosticMessageProbeException(string label) : base(label) { this.label = label; }
        public void Reset() { Reads = 0; }
        public override string Message { get { Reads++; return label; } }
    }
}
"@
    }
}
Describe 'Hosted Default desktop entry lifetime' {
    BeforeEach {
        $script:priorDefaultDesktopEnvironment = @{}
        foreach ($name in @('GITHUB_ACTIONS','RUNNER_ENVIRONMENT','GITHUB_REPOSITORY')) {
            $script:priorDefaultDesktopEnvironment[$name] = [Environment]::GetEnvironmentVariable($name, 'Process')
        }
        $env:GITHUB_ACTIONS='true'; $env:RUNNER_ENVIRONMENT='github-hosted'; $env:GITHUB_REPOSITORY='RedSalamanders/RedSalamander'
        $global:rsDefaultDesktopTest = [pscustomobject]@{ Events=[Collections.Generic.List[string]]::new(); FailCreate=$false; FailScope=$false; FailScopeRetire=$false; FailRestore=$false; NestedRestoreFailure=$false; ReportRestored=$true; DiagnosticText=[Collections.Generic.List[string]]::new() }
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
        Mock Write-Host {} -ModuleName HostedTestDesktop
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
    It 'retains the original nested setup error and captures both aggregate branches without native execution' {
        $native=[ComponentModel.Win32Exception]::new(5, 'seeded native setup leaf')
        $setup=[InvalidOperationException]::new('seeded boundary startup', $native)
        $cleanup=[InvalidOperationException]::new('seeded boundary cleanup', [TimeoutException]::new('seeded cleanup timeout'))
        $failure=[AggregateException]::new('seeded setup and cleanup', [Exception[]]@($setup,$cleanup))
        $record=[Management.Automation.ErrorRecord]::new($failure,'seeded-hosted-setup',[Management.Automation.ErrorCategory]::InvalidOperation,$null)
        $global:rsDefaultDesktopTest | Add-Member -NotePropertyName SetupError -NotePropertyValue $record
        Mock New-RSHostedDefaultDesktop {
            $global:rsDefaultDesktopTest.Events.Add('create')
            throw $global:rsDefaultDesktopTest.SetupError
        } -ModuleName HostedTestDesktop
        Mock Write-Host { $global:rsDefaultDesktopTest.DiagnosticText.Add(($Object -join ' ')) } -ModuleName HostedTestDesktop
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedDefaultDesktop { throw 'must not execute' } }
        $r.Succeeded | Should Be $false
        $r.Error.FullyQualifiedErrorId | Should Be 'seeded-hosted-setup'
        ([object]::ReferenceEquals($r.Error.Exception,$failure)) | Should Be $true
        ($global:rsDefaultDesktopTest.Events -join ',') | Should Be 'create'
        $global:rsDefaultDesktopTest.DiagnosticText.Count | Should Be 1
        $diagnostic=$global:rsDefaultDesktopTest.DiagnosticText[0]
        $diagnostic | Should Match 'Hosted Default desktop setup exception:'
        $diagnostic | Should Match 'System.ComponentModel.Win32Exception'
        $diagnostic | Should Match 'NativeErrorCode=5'
        $diagnostic | Should Match ('HResult=0x'+$native.HResult.ToString('X8'))
        $diagnostic | Should Match 'seeded native setup leaf'
        $diagnostic | Should Match 'System.TimeoutException'
        $diagnostic | Should Match 'seeded cleanup timeout'
        Assert-MockCalled New-RSHostedDesktopProcessScope -ModuleName HostedTestDesktop -Scope It -Times 0 -Exactly
        Assert-MockCalled Set-RSHostedDisplayMode -ModuleName HostedTestDesktop -Scope It -Times 0 -Exactly
    }
    It 'preserves the primary setup error when the diagnostic sink fails' {
        $failure=[InvalidOperationException]::new('seeded primary setup error')
        $record=[Management.Automation.ErrorRecord]::new($failure,'seeded-hosted-sink',[Management.Automation.ErrorCategory]::InvalidOperation,$null)
        $global:rsDefaultDesktopTest | Add-Member -NotePropertyName SetupError -NotePropertyValue $record
        Mock New-RSHostedDefaultDesktop {
            $global:rsDefaultDesktopTest.Events.Add('create')
            throw $global:rsDefaultDesktopTest.SetupError
        } -ModuleName HostedTestDesktop
        Mock Write-Host { throw [IO.IOException]::new('seeded diagnostic sink failure') } -ModuleName HostedTestDesktop
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedDefaultDesktop { throw 'must not execute' } }
        $r.Succeeded | Should Be $false
        $r.Error.FullyQualifiedErrorId | Should Be 'seeded-hosted-sink'
        ([object]::ReferenceEquals($r.Error.Exception,$failure)) | Should Be $true
        $r.Error.Exception.Message | Should Be 'seeded primary setup error'
        ($global:rsDefaultDesktopTest.Events -join ',') | Should Be 'create'
        Assert-MockCalled Write-Host -ModuleName HostedTestDesktop -Scope It -Times 1 -Exactly
        Assert-MockCalled New-RSHostedDesktopProcessScope -ModuleName HostedTestDesktop -Scope It -Times 0 -Exactly
    }
    It 'propagates a synthetic fatal diagnostic sink exception without resource exhaustion' {
        $primary=[InvalidOperationException]::new('seeded primary before fatal diagnostic')
        $record=[Management.Automation.ErrorRecord]::new($primary,'seeded-hosted-fatal-primary',[Management.Automation.ErrorCategory]::InvalidOperation,$null)
        $fatal=[AccessViolationException]::new('seeded synthetic fatal diagnostic sink')
        $global:rsDefaultDesktopTest | Add-Member -NotePropertyName SetupError -NotePropertyValue $record
        $global:rsDefaultDesktopTest | Add-Member -NotePropertyName FatalSinkError -NotePropertyValue $fatal
        Mock New-RSHostedDefaultDesktop {
            $global:rsDefaultDesktopTest.Events.Add('create')
            throw $global:rsDefaultDesktopTest.SetupError
        } -ModuleName HostedTestDesktop
        Mock Write-Host { throw $global:rsDefaultDesktopTest.FatalSinkError } -ModuleName HostedTestDesktop
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedDefaultDesktop { throw 'must not execute' } }
        $r.Succeeded | Should Be $false
        ([object]::ReferenceEquals($r.Error.Exception,$fatal)) | Should Be $true
        $r.Error.Exception.Message | Should Be 'seeded synthetic fatal diagnostic sink'
        ($global:rsDefaultDesktopTest.Events -join ',') | Should Be 'create'
        Assert-MockCalled New-RSHostedDesktopProcessScope -ModuleName HostedTestDesktop -Scope It -Times 0 -Exactly
    }
    It 'bounds setup diagnostic depth without traversing an excluded leaf' {
        Mock Write-Host { $global:rsDefaultDesktopTest.DiagnosticText.Add(($Object -join ' ')) } -ModuleName HostedTestDesktop
        $failure=[Exception]::new('unreachable-depth-leaf')
        foreach($index in 1..12){$failure=[InvalidOperationException]::new(('seeded level '+$index),$failure)}
        $record=[Management.Automation.ErrorRecord]::new($failure,'seeded-depth',[Management.Automation.ErrorCategory]::InvalidOperation,$null)
        & (Get-Module HostedTestDesktop) { param($captured) Write-RSHostedSetupFailureDiagnostic -ErrorRecord $captured } $record
        $global:rsDefaultDesktopTest.DiagnosticText.Count | Should Be 1
        $diagnostic=$global:rsDefaultDesktopTest.DiagnosticText[0]
        $diagnostic | Should Match '\[diagnostic-truncated: depth\]'
        $diagnostic | Should Not Match 'unreachable-depth-leaf'
        ($diagnostic.Length -le 16384) | Should Be $true
        $global:rsDefaultDesktopTest.Events.Count | Should Be 0
    }
    It 'bounds setup diagnostic nodes fields and total text with explicit truncation markers' {
        Mock Write-Host { $global:rsDefaultDesktopTest.DiagnosticText.Add(($Object -join ' ')) } -ModuleName HostedTestDesktop
        $children=[Collections.Generic.List[Exception]]::new()
        foreach($index in 1..48){$children.Add([Exception]::new(('x'*5000)+('leaf-'+$index)))}
        $failure=[AggregateException]::new('seeded wide exception', $children)
        $record=[Management.Automation.ErrorRecord]::new($failure,'seeded-width',[Management.Automation.ErrorCategory]::InvalidOperation,$null)
        & (Get-Module HostedTestDesktop) { param($captured) Write-RSHostedSetupFailureDiagnostic -ErrorRecord $captured } $record
        $global:rsDefaultDesktopTest.DiagnosticText.Count | Should Be 1
        $diagnostic=$global:rsDefaultDesktopTest.DiagnosticText[0]
        $diagnostic | Should Match 'diagnostic-truncated:.*field-string'
        $diagnostic | Should Match 'diagnostic-truncated:.*node-count'
        $diagnostic | Should Match 'diagnostic-truncated:.*output-string'
        ($diagnostic.Length -le 16384) | Should Be $true
        $global:rsDefaultDesktopTest.Events.Count | Should Be 0
    }
    It 'does not read message getters below the aggregate depth budget' {
        Initialize-RSHostedDiagnosticMessageProbe
        $probe=[RedSalamander.Testing.HostedDiagnosticMessageProbeException]::new('excluded-deep-leaf')
        $tree=$probe
        for($index=0;$index -lt 12;$index++) { $tree=[AggregateException]::new('composite', [Exception[]]@($tree)) }
        $record=[Management.Automation.ErrorRecord]::new($tree,'DeepAggregate',[Management.Automation.ErrorCategory]::NotSpecified,$null)
        $probe.Reset()
        Mock Write-Host { param($Object) $global:rsDefaultDesktopTest.DiagnosticText.Add([string]$Object) } -ModuleName HostedTestDesktop
        & (Get-Module HostedTestDesktop) { param($errorRecord) Write-RSHostedSetupFailureDiagnostic -ErrorRecord $errorRecord } $record
        $probe.Reads | Should Be 0
        $text=$global:rsDefaultDesktopTest.DiagnosticText -join ''
        $text | Should Match '\[aggregate-summary-omitted\]'
        $text | Should Match 'diagnostic-truncated: depth'
        $text | Should Not Match 'excluded-deep-leaf'
        Assert-MockCalled New-RSHostedDefaultDesktop -ModuleName HostedTestDesktop -Scope It -Times 0 -Exactly
        Assert-MockCalled New-RSHostedDesktopProcessScope -ModuleName HostedTestDesktop -Scope It -Times 0 -Exactly
    }
It 'reads only visited leaf message getters within the aggregate node budget' {
        Initialize-RSHostedDiagnosticMessageProbe
        $probes=@(0..47 | ForEach-Object { [RedSalamander.Testing.HostedDiagnosticMessageProbeException]::new("leaf-$_") })
        $tree=[AggregateException]::new('composite', [Exception[]]$probes)
        $record=[Management.Automation.ErrorRecord]::new($tree,'WideAggregate',[Management.Automation.ErrorCategory]::NotSpecified,$null)
        foreach($probe in $probes) { $probe.Reset() }
        Mock Write-Host { param($Object) $global:rsDefaultDesktopTest.DiagnosticText.Add([string]$Object) } -ModuleName HostedTestDesktop
        & (Get-Module HostedTestDesktop) { param($errorRecord) Write-RSHostedSetupFailureDiagnostic -ErrorRecord $errorRecord } $record
        ($probes[31..47] | Measure-Object Reads -Sum).Sum | Should Be 0
        ($probes[0..30] | Measure-Object Reads -Sum).Sum | Should Be 31
        $text=$global:rsDefaultDesktopTest.DiagnosticText -join ''
        $text | Should Match '\[aggregate-summary-omitted\]'
        $text | Should Match 'diagnostic-truncated: node-count'
        Assert-MockCalled New-RSHostedDefaultDesktop -ModuleName HostedTestDesktop -Scope It -Times 0 -Exactly
        Assert-MockCalled New-RSHostedDesktopProcessScope -ModuleName HostedTestDesktop -Scope It -Times 0 -Exactly
    }
    It 'does not invoke setup diagnostic formatting on the successful path' {
        Mock Write-RSHostedSetupFailureDiagnostic {} -ModuleName HostedTestDesktop
        $r=Invoke-RSTestBlock { Invoke-RSWithHostedDefaultDesktop { 7 } }
        $r.Succeeded | Should Be $true
        $r.Output[0] | Should Be 7
        ($global:rsDefaultDesktopTest.Events -join ',') | Should Be 'create,process-scope,retire-process-scope,restore'
        Assert-MockCalled Write-RSHostedSetupFailureDiagnostic -ModuleName HostedTestDesktop -Scope It -Times 0 -Exactly
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

Describe 'Hosted focus query failure diagnostics' {
    BeforeEach {
        Mock Write-Host {} -ModuleName HostedTestDesktop
        & (Get-Module HostedTestDesktop) { Initialize-RSHostedDesktopNative }
        $script:focusBoundary = ('RedSalamander.Testing.HostedDesktopNative' -as [type]).Assembly.GetType('RedSalamander.Testing.HostedForegroundBoundary', $true)
        $script:focusFlags = [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic
        $script:focusFailure = $script:focusBoundary.GetMethod('FocusReadFailure', $script:focusFlags)
        if (-not ('RedSalamander.Testing.HostedFocusFailureControl' -as [type])) {
            # Managed delegates only: no window, desktop, input or process queries.
            Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Runtime.InteropServices;
namespace RedSalamander.Testing {
    public static class HostedFocusFailureControl {
        public static int Calls;
        public static long Window;
        public static Func<IntPtr, string> ChangedAmbient { get { return Observe; } }
        public static Func<IntPtr, string> Unavailable { get { return Fail; } }
        public static Func<IntPtr, string> Fatal { get { return FailFatally; } }
        private static string Observe(IntPtr window) {
            Calls++; Window = window.ToInt64(); Marshal.SetLastPInvokeError(5);
            return "originalWindowValid=False; observedThread=0; observedPid=0; currentForeground=0x22; foregroundMatches=False";
        }
        private static string Fail(IntPtr window) {
            Calls++; Window = window.ToInt64();
            throw new IOException("private probe exception text");
        }
        private static string FailFatally(IntPtr window) { throw new AccessViolationException("synthetic fatal focus observer"); }
        // Compiler control only: never invoked. The old failed-branch capture shape
        // still requires display-class storage at this method's entry.
        public static Func<IntPtr, string> EagerCaptureForControl(bool failed, uint thread, uint process) {
            if (failed) return window => (thread + process + window.ToInt64()).ToString();
            return null;
        }
    }
}
'@
        }
        [RedSalamander.Testing.HostedFocusFailureControl]::Calls = 0
        [RedSalamander.Testing.HostedFocusFailureControl]::Window = 0
    }
    AfterEach {
        Remove-Variable rsFocusBoundedText -Scope Global -ErrorAction SilentlyContinue
    }
    It 'retains the saved native error and exact context after an observer changes ambient error' {
        $saved = [Runtime.InteropServices.Marshal]::GetLastPInvokeError()
        try {
            foreach ($stage in @('entry-original-focus','verify-original-focus','restore-original-focus')) {
                $failure = $script:focusFailure.Invoke($null, [object[]]@([int]87,[uint32]12,[uint32]72,[IntPtr]0x11,[uint32]34,$stage,'saved-entry-identity',[RedSalamander.Testing.HostedFocusFailureControl]::ChangedAmbient))
                ($failure -is [ComponentModel.Win32Exception]) | Should Be $true
                $failure.NativeErrorCode | Should Be 87
                $failure.Message.StartsWith('Cannot verify the original hosted foreground focus.') | Should Be $true
                $failure.Message | Should Match ('stage='+$stage+'; queryThread=0; expectedOriginalThread=12; requestedPid=34; requestedHwnd=0x11; cbSize=72; pointerBytes='+[IntPtr]::Size)
                $failure.Message | Should Match 'originalAtCapture=\{saved-entry-identity\}'
                $failure.Message | Should Match 'originalWindowValid=False; observedThread=0; observedPid=0; currentForeground=0x22; foregroundMatches=False'
                $failure.Message | Should Match 'nonAtomic=True; failedGuiFields=unspecified'
            }
            [RedSalamander.Testing.HostedFocusFailureControl]::Calls | Should Be 3
            [RedSalamander.Testing.HostedFocusFailureControl]::Window | Should Be 17
        } finally { [Runtime.InteropServices.Marshal]::SetLastPInvokeError($saved) }
    }
    It 'preserves the saved error when observation is unavailable and does not disclose its exception message' {
        $failure = $script:focusFailure.Invoke($null, [object[]]@([int]87,[uint32]12,[uint32]72,[IntPtr]0x11,[uint32]34,'entry-original-focus',$null,[RedSalamander.Testing.HostedFocusFailureControl]::Unavailable))
        $failure.NativeErrorCode | Should Be 87
        $failure.Message | Should Match 'originalAtCapture=\{unavailable\}'
        $failure.Message | Should Match 'unavailable=IOException'
        $failure.Message | Should Not Match 'private probe exception text'
        [RedSalamander.Testing.HostedFocusFailureControl]::Calls | Should Be 1
    }
    It 'does not swallow a synthetic fatal observer exception' {
        $caught = $null
        try { [void]$script:focusFailure.Invoke($null, [object[]]@([int]87,[uint32]12,[uint32]72,[IntPtr]0x11,[uint32]34,'entry-original-focus',$null,[RedSalamander.Testing.HostedFocusFailureControl]::Fatal)) }
        catch [Management.Automation.MethodInvocationException] { $caught = $_.Exception.GetBaseException() }
        ($caught -is [AccessViolationException]) | Should Be $true
        $caught.Message | Should Be 'synthetic fatal focus observer'
    }
    It 'matches the pointer-width GUITHREADINFO layout and native signature without calling user32' {
        $gui = $script:focusBoundary.GetNestedType('GuiThreadInfo', [Reflection.BindingFlags]::NonPublic)
        $sizeOf = [Runtime.InteropServices.Marshal].GetMethod('SizeOf', [type[]]@([type]))
        $expected = if ([IntPtr]::Size -eq 8) { 72 } else { 48 }
        $sizeOf.Invoke($null, [object[]]@($gui)) | Should Be $expected
        foreach ($field in @(@{name='Size';offset=0},@{name='Flags';offset=4},@{name='Active';offset=8},
                @{name='Focus';offset=8+[IntPtr]::Size},@{name='Capture';offset=8+2*[IntPtr]::Size},
                @{name='MenuOwner';offset=8+3*[IntPtr]::Size},@{name='MoveSize';offset=8+4*[IntPtr]::Size},
                @{name='Caret';offset=8+5*[IntPtr]::Size},@{name='CaretBounds';offset=8+6*[IntPtr]::Size})) {
            [Runtime.InteropServices.Marshal]::OffsetOf($gui, $field.name).ToInt32() | Should Be $field.offset
        }
        $method = $script:focusBoundary.GetMethod('GetGUIThreadInfo', $script:focusFlags)
        $method.GetParameters()[0].ParameterType | Should Be ([uint32])
        $method.GetParameters()[1].ParameterType | Should Be ($gui.MakeByRefType())
        $import = @($method.GetCustomAttributes([Runtime.InteropServices.DllImportAttribute], $false))[0]
        $import.SetLastError | Should Be $true
        $import.ExactSpelling | Should Be $true
        $return = @($method.ReturnParameter.GetCustomAttributes([Runtime.InteropServices.MarshalAsAttribute], $false))[0]
        $return.Value | Should Be ([Runtime.InteropServices.UnmanagedType]::Bool)
    }
    It 'captures the original error before only the fixed failure probes and rejects an ordering counterfactual' {
        $source = [IO.File]::ReadAllText((Join-Path $repoRoot 'Tools/Modules/Testing/HostedTestDesktop.psm1'))
        $check = {
            param($text)
            $start = $text.IndexOf('private static GuiThreadInfo ReadFocus(')
            $end = $text.IndexOf('internal static Win32Exception FocusReadFailure(', $start)
            $read = $text.Substring($start, $end-$start)
            if ($read -notmatch 'if\s*\(!GetGUIThreadInfo\(0, ref info\)\)\s*\{\s*int error = Marshal.GetLastWin32Error\(\);\s*throw CaptureFocusReadFailure\(error, thread, size, original, originalProcess, stage, originalAtCapture\);') {
                throw 'the original native error must be captured before failure observations'
            }
        }
        & $check $source
        { & $check ($source.Replace('int error = Marshal.GetLastWin32Error();','int error = 0;')) } | Should Throw 'the original native error must be captured'
        $start = $source.IndexOf('internal static Win32Exception FocusReadFailure(')
        $end = $source.IndexOf('internal static ushort[] RemainingOwnedChordKeyUps(', $start)
        $probe = $source.Substring($start,$end-$start)
        foreach ($api in @('IsWindow','GetWindowThreadProcessId','GetForegroundWindow')) {
            [regex]::Matches($probe, ('\b'+$api+'\(')).Count | Should Be 1
        }
        $probe | Should Not Match 'SetForegroundWindow|SetFocus|SendInput|AttachThreadInput|OpenProcess|OpenThread|EnumWindows|GetWindowText|Sleep|Wait|Retry'
        [regex]::Matches($source, 'ReadOriginalForegroundFocus\(originalThread, original, originalProcess, "(?:entry|verify|restore)-original-focus",').Count | Should Be 3
    }
    It 'keeps observer capture outside the compiled successful query and rejects the old capture shape' {
        $read = $script:focusBoundary.GetMethod('ReadFocus', $script:focusFlags)
        $displayClasses = @($read.GetMethodBody().LocalVariables | Where-Object {
            $_.LocalType.IsDefined([Runtime.CompilerServices.CompilerGeneratedAttribute], $false) -and $_.LocalType.Name -like '*DisplayClass*'
        })
        $displayClasses.Count | Should Be 0
        $control = [RedSalamander.Testing.HostedFocusFailureControl].GetMethod('EagerCaptureForControl')
        @($control.GetMethodBody().LocalVariables | Where-Object {
            $_.LocalType.IsDefined([Runtime.CompilerServices.CompilerGeneratedAttribute], $false) -and $_.LocalType.Name -like '*DisplayClass*'
        }).Count | Should Be 1
        $source = [IO.File]::ReadAllText((Join-Path $repoRoot 'Tools/Modules/Testing/HostedTestDesktop.psm1'))
        $start = $source.IndexOf('private static GuiThreadInfo ReadFocus(')
        $end = $source.IndexOf('internal static Win32Exception FocusReadFailure(', $start)
        $body = $source.Substring($start,$end-$start)
        $guard = { param($text) if ($text -match '=>|\bdelegate\b') { throw 'ReadFocus cannot construct an observer closure' } }
        & $guard $body
        $oldShape = $body.Replace('throw CaptureFocusReadFailure(error, thread, size, original, originalProcess, stage, originalAtCapture);',
            'throw FocusReadFailure(error, thread, size, original, originalProcess, stage, originalAtCapture, window => ObserveFocusFailureWindow(window, thread, originalProcess));')
        ($oldShape -cne $body) | Should Be $true
        { & $guard $oldShape } | Should Throw 'ReadFocus cannot construct an observer closure'
    }
    It 'retains fresh focus facts through the real bounded formatter before clipping a long saved identity' {
        $saved = [Runtime.InteropServices.Marshal]::GetLastPInvokeError()
        $global:rsFocusBoundedText = [Collections.Generic.List[string]]::new()
        Mock Write-Host { param($Object) $global:rsFocusBoundedText.Add([string]$Object) } -ModuleName HostedTestDesktop
        try {
            $failure = $script:focusFailure.Invoke($null, [object[]]@([int]87,[uint32]12,[uint32]72,[IntPtr]0x11,[uint32]34,'entry-original-focus',('old-identity-'+('x'*5000)),[RedSalamander.Testing.HostedFocusFailureControl]::ChangedAmbient))
            $record = [Management.Automation.ErrorRecord]::new($failure,'bounded-focus-context',[Management.Automation.ErrorCategory]::NotSpecified,$null)
            & (Get-Module HostedTestDesktop) { param($captured) Write-RSHostedSetupFailureDiagnostic -ErrorRecord $captured } $record
            $global:rsFocusBoundedText.Count | Should Be 1
            $text = $global:rsFocusBoundedText[0]
            $text | Should Match 'NativeErrorCode=87'
            $text | Should Match 'HResult=0x80004005'
            $text | Should Match 'queryThread=0; expectedOriginalThread=12; requestedPid=34; requestedHwnd=0x11; cbSize=72'
            $text | Should Match 'failureObservation=\{originalWindowValid=False; observedThread=0; observedPid=0; currentForeground=0x22; foregroundMatches=False\}'
            $text | Should Match 'nonAtomic=True; failedGuiFields=unspecified'
            $text | Should Match 'diagnostic-truncated: field-string'
            ($text.Length -le 16384) | Should Be $true
            # Same real formatter, but the rejected order puts the long old identity first.
            $freshStart = $failure.Message.IndexOf('failureObservation={')
            $oldStart = $failure.Message.IndexOf('originalAtCapture={')
            $oldOrder = $failure.Message.Substring(0,$freshStart) + $failure.Message.Substring($oldStart) + '; ' +
                $failure.Message.Substring($freshStart,$oldStart-$freshStart)
            $global:rsFocusBoundedText.Clear()
            $oldFailure = [ComponentModel.Win32Exception]::new(87,$oldOrder)
            $oldRecord = [Management.Automation.ErrorRecord]::new($oldFailure,'rejected-focus-order',[Management.Automation.ErrorCategory]::NotSpecified,$null)
            & (Get-Module HostedTestDesktop) { param($captured) Write-RSHostedSetupFailureDiagnostic -ErrorRecord $captured } $oldRecord
            $global:rsFocusBoundedText.Count | Should Be 1
            $global:rsFocusBoundedText[0] | Should Not Match 'originalWindowValid=False|nonAtomic=True'
            $global:rsFocusBoundedText[0] | Should Match 'NativeErrorCode=87'
            $global:rsFocusBoundedText[0] | Should Match 'diagnostic-truncated: field-string'
        } finally { [Runtime.InteropServices.Marshal]::SetLastPInvokeError($saved) }
    }
    It 'reuses only the matching normalized inline implementation in the same process' {
        $before = 'RedSalamander.Testing.HostedDesktopNative' -as [type]
        & (Get-Module HostedTestDesktop) { Initialize-RSHostedDesktopNative }
        ([object]::ReferenceEquals($before, ('RedSalamander.Testing.HostedDesktopNative' -as [type]))) | Should Be $true
        $field = $before.GetField('ImplementationSourceHash', $script:focusFlags)
        ([string]$field.GetRawConstantValue()) | Should Match '^[a-f0-9]{64}$'
        $field.IsPrivate | Should Be $true
    }
    It 'refuses missing or changed cached implementation identity before native initialization' {
        $tokens=$null; $errors=$null
        $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repoRoot 'Tools/Modules/Testing/HostedTestDesktop.psm1'),[ref]$tokens,[ref]$errors)
        $initializer=@($ast.FindAll({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Initialize-RSHostedDesktopNative'},$true))[0]
        $body=$initializer.Body.Extent.Text.Substring(1,$initializer.Body.Extent.Text.Length-2)
        foreach ($counterfactual in @(
                $body.Replace('using System;', 'using System; // changed implementation identity'),
                $body.Replace('$loaded = ''RedSalamander.Testing.HostedDesktopNative'' -as [type]', '$loaded = [Exception]'))) {
            $attempt=Invoke-RSTestBlock { & ([scriptblock]::Create($counterfactual)) }
            $attempt.Succeeded | Should Be $false
            $attempt.Error.Exception.Message | Should Match 'start a fresh PowerShell process'
        }
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
        $global:stockDiagnostic='instanceAvailable=True; creationMatchesCapture=True; childReads=1; parentReads=0; relationReads=1'
        Mock Get-RSHostedStockSurface {
            $surface=$global:stockObservations.Dequeue()
            $surface | Add-Member NoteProperty CapturedUtc '2026-11-14T00:00:00.0000000Z' -Force
            $surface | Add-Member NoteProperty CreationKnown $true -Force
            $surface | Add-Member NoteProperty CreationFileTime100ns 200 -Force
            $surface | Add-Member NoteProperty CreationError 0 -Force
            $surface | Add-Member ScriptMethod DescribeCapturedProcess { param($observer) $global:stockDiagnostic } -Force
            $surface
        } -ModuleName HostedTestDesktop
    }
    AfterEach {
        $env:GITHUB_ACTIONS=$script:savedActions; $env:RUNNER_ENVIRONMENT=$script:savedRunner; $env:GITHUB_REPOSITORY=$script:savedRepo
        Remove-Variable stockObservations,stockCloses,stockDiagnostic -Scope Global -ErrorAction SilentlyContinue
    }
    It 'limits opt-in to native ARM64 hosted project roles' {
        $p=@{Phase='ReleaseTest';Platform='ARM64';QualificationScope='Normal';OsArchitecture='Arm64';ProcessArchitecture='Arm64'}
        Get-RSHostedStockPreparationRequest @p | Should Be '1'
        $p.Phase='Qualification'; $p.QualificationScope='CommandsHandoff'; Get-RSHostedStockPreparationRequest @p | Should Be '1'
        $p.QualificationScope='CommandsUiAttribution'; Get-RSHostedStockPreparationRequest @p | Should Be '1'
        $p.Platform='x64'; Get-RSHostedStockPreparationRequest @p | Should BeNullOrEmpty
        $p.Platform='ARM64';$p.Phase='Static';Get-RSHostedStockPreparationRequest @p | Should BeNullOrEmpty
        $p.Phase='Qualification'
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
        $completion.first_observation.kind | Should Be 'Account'
        $completion.observation_count | Should Be 3
        $completion.extended_observations | Should Be 1
    }
    It 'preserves the primary Unknown refusal with unavailable enrichment and zero close requests' {
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestSandbox.psm1')
        $dir=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('hosted-unknown-observation-'+[guid]::NewGuid().ToString('N'))
        $global:stockDiagnostic='instanceAvailable=False; unavailable=InvalidOperationException'
        $surface=[pscustomobject]@{Kind='Unknown';Window=123;Thread=1;Process=2;Session=3;Title='fixture';Class='fixture';Image='fixture';ApplicationId='';ForegroundDesktop='Default';ForegroundDesktopError=0}
        $surface | Add-Member ScriptMethod RequestClose { $global:stockCloses.Add($this.Kind);throw 'Unknown must never close.' }
        $surface | Add-Member ScriptMethod Dispose { }
        $global:stockObservations.Enqueue($surface)
        $attempt=Invoke-RSTestBlock { & (Get-Module HostedTestDesktop) {param($d) Invoke-RSHostedStockPreparationCore -Directory $d -SourceHead ('a'*40) -ObservationMilliseconds 0} $dir }
        $attempt.Succeeded | Should Be $false
        $attempt.Error.Exception.Message | Should Be 'An unapproved foreground surface prevents hosted preparation.'
        $completion=Get-Content (Join-Path $dir 'completion.json') -Raw | ConvertFrom-Json
        $completion.error | Should Be 'An unapproved foreground surface prevents hosted preparation.'
        $completion.last_observation.process_diagnostic | Should Be $global:stockDiagnostic
        @($completion.first_observation.PSObject.Properties.Name) -ccontains 'process_diagnostic' | Should Be $false
        $completion.first_observation.creation_file_time_100ns | Should Be 200
        $completion.last_observation.creation_file_time_100ns | Should Be 200
        $completion.observation_count | Should Be 1;$completion.extended_observations | Should Be 1
        $completion.observation_limit | Should Match 'absence is not background-process absence or creator proof'
        $global:stockCloses.Count | Should Be 0
        @(Get-ChildItem $dir -Filter '*-intent.json').Count | Should Be 0
    }
}

Describe 'Hosted console-host delegation precondition' {
    BeforeEach {
        $script:savedActions=$env:GITHUB_ACTIONS; $script:savedRunner=$env:RUNNER_ENVIRONMENT; $script:savedRepo=$env:GITHUB_REPOSITORY
        $env:GITHUB_ACTIONS='true'; $env:RUNNER_ENVIRONMENT='github-hosted'; $env:GITHUB_REPOSITORY='RedSalamanders/RedSalamander'
        # In-memory registry only: no fixture reads or writes the real HKCU.
        $global:consoleHostValues=@{}; $global:consoleHostWrites=[Collections.Generic.List[string]]::new()
        $script:consoleHostRead={param($name) $global:consoleHostValues[$name]}
        $script:consoleHostWrite={param($name,$value) $global:consoleHostWrites.Add($name); $global:consoleHostValues[$name]=$value}
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestSandbox.psm1') -Force
    }
    AfterEach {
        $env:GITHUB_ACTIONS=$script:savedActions; $env:RUNNER_ENVIRONMENT=$script:savedRunner; $env:GITHUB_REPOSITORY=$script:savedRepo
        Remove-Variable consoleHostValues,consoleHostWrites -Scope Global -ErrorAction SilentlyContinue
    }
    It 'refuses local, self-hosted and non-ARM64 invocation before any registry access' {
        $arguments = @{
            Phase = 'ReleaseTest'
            Platform = 'ARM64'
            QualificationScope = 'Normal'
            TestRoot = 'X:\RedSalamander.Perf'
            EvidenceDirectory = 'X:\RedSalamander.Perf\evidence'
            SourceCommit = 'a' * 40
        }
        $env:GITHUB_ACTIONS='false'
        { Set-RSHostedConsoleHostDelegation @arguments } | Should Throw 'hosted native ARM64 test roles'
        $env:GITHUB_ACTIONS='true'; $env:RUNNER_ENVIRONMENT='self-hosted'
        { Set-RSHostedConsoleHostDelegation @arguments } | Should Throw 'hosted native ARM64 test roles'
        $env:RUNNER_ENVIRONMENT='github-hosted'
        $arguments.Platform='x64'
        { Set-RSHostedConsoleHostDelegation @arguments } | Should Throw 'hosted native ARM64 test roles'
    }
    It 'writes both unset values, verifies them and records the prior state create-only' {
        $dir=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('hosted-console-host-'+[guid]::NewGuid().ToString('N'))
        $global:consoleHostValues['DelegationTerminal']='{00000000-0000-0000-0000-000000000000}'
        $audit=& (Get-Module HostedTestDesktop) {param($d,$r,$w) Invoke-RSHostedConsoleHostDelegationCore -Directory $d -Read $r -Write $w -SourceCommit ('a' * 40)} $dir $script:consoleHostRead $script:consoleHostWrite
        @($global:consoleHostWrites) -join ',' | Should Be 'DelegationConsole,DelegationTerminal'
        $global:consoleHostValues['DelegationConsole'] | Should Be '{B23D10C0-E52E-411E-9D5B-C09FDF709C7D}'
        $global:consoleHostValues['DelegationTerminal'] | Should Be '{B23D10C0-E52E-411E-9D5B-C09FDF709C7D}'
        $record=Get-Content (Join-Path $dir 'delegation.json') -Raw | ConvertFrom-Json
        $record.schema | Should Be 'red-salamander.hosted-console-host-delegation.v1'
        $record.source_commit | Should Be ('a' * 40)
        $record.verified | Should Be $true
        $record.prior.DelegationConsole | Should BeNullOrEmpty
        $record.prior.DelegationTerminal | Should Be '{00000000-0000-0000-0000-000000000000}'
        @($record.changed) -join ',' | Should Be 'DelegationConsole,DelegationTerminal'
        $audit.verified | Should Be $true
        $global:consoleHostValues['DelegationConsole'] = $null
        $repeat=Invoke-RSTestBlock { & (Get-Module HostedTestDesktop) {param($d,$r,$w) Invoke-RSHostedConsoleHostDelegationCore -Directory $d -Read $r -Write $w -SourceCommit ('a' * 40)} $dir $script:consoleHostRead $script:consoleHostWrite }
        $repeat.Succeeded | Should Be $false
        $repeat.Error.Exception.Message | Should Match 'already exists'
        $global:consoleHostWrites.Count | Should Be 2
        (Get-Content (Join-Path $dir 'delegation.json') -Raw | ConvertFrom-Json).observed_utc | Should Be $record.observed_utc
    }
    It 'leaves already selected values unchanged regardless of case' {
        $dir=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('hosted-console-host-set-'+[guid]::NewGuid().ToString('N'))
        $global:consoleHostValues['DelegationConsole']='{b23d10c0-e52e-411e-9d5b-c09fdf709c7d}'
        $global:consoleHostValues['DelegationTerminal']='{B23D10C0-E52E-411E-9D5B-C09FDF709C7D}'
        $audit=& (Get-Module HostedTestDesktop) {param($d,$r,$w) Invoke-RSHostedConsoleHostDelegationCore -Directory $d -Read $r -Write $w -SourceCommit ('a' * 40)} $dir $script:consoleHostRead $script:consoleHostWrite
        $global:consoleHostWrites.Count | Should Be 0
        @($audit.changed).Count | Should Be 0
        $audit.verified | Should Be $true
    }
    It 'fails closed and still records the audit when a value does not read back' {
        $dir=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('hosted-console-host-fail-'+[guid]::NewGuid().ToString('N'))
        $ignoredWrite={param($name,$value) $global:consoleHostWrites.Add($name)}
        $attempt=Invoke-RSTestBlock { & (Get-Module HostedTestDesktop) {param($d,$r,$w) Invoke-RSHostedConsoleHostDelegationCore -Directory $d -Read $r -Write $w -SourceCommit ('a' * 40)} $dir $script:consoleHostRead $ignoredWrite }
        $attempt.Succeeded | Should Be $false
        $attempt.Error.Exception.Message | Should Match 'DelegationConsole did not read back'
        $record=Get-Content (Join-Path $dir 'delegation.json') -Raw | ConvertFrom-Json
        $record.verified | Should Be $false
        $record.error | Should Match 'did not read back'
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
    BeforeEach {
        & (Get-Module HostedTestDesktop) { Initialize-RSHostedDesktopNative }
        if (-not ('RedSalamander.Testing.HostedIdentityDiagnosticControl' -as [type])) {
            # Managed-only observers: no windows, desktop queries, activation or input.
            Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
namespace RedSalamander.Testing {
    public static class HostedIdentityDiagnosticControl {
        public static readonly List<long> Seen = new List<long>();
        public static Func<IntPtr, string> Recorder { get { return Observe; } }
        public static Func<IntPtr, string> Unavailable { get { return Fail; } }
        public static Func<IntPtr, string> Fatal { get { return FailFatally; } }
        private static string Observe(IntPtr window) {
            Seen.Add(window.ToInt64());
            return "observed-hwnd=0x" + window.ToInt64().ToString("X");
        }
        private static string Fail(IntPtr window) {
            Seen.Add(window.ToInt64());
            throw new InvalidOperationException("seeded private diagnostic text");
        }
        private static string FailFatally(IntPtr window) { throw new OutOfMemoryException("seeded fatal control"); }
    }
}
'@
        }
        [RedSalamander.Testing.HostedIdentityDiagnosticControl]::Seen.Clear()
    }
    It 'formats window identities without native calls or title/content fields' {
        $native = 'RedSalamander.Testing.HostedDesktopNative' -as [type]
        $format = $native.GetMethod('FormatWindowIdentity', [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic)
        $captured = [DateTimeOffset]::Parse('2026-11-14T00:00:00+00:00', [Globalization.CultureInfo]::InvariantCulture)
        $identity = $format.Invoke($null, [object[]]@([IntPtr]0x123, [uint32]12, [uint32]34, "Example`r`nClass", $true, $captured))
        $identity | Should Match 'hwnd=0x123; thread=12; pid=34; classKnown=True; class="Example\\r\\nClass"; capturedUtc=2026-11-14T00:00:00.0000000\+00:00'
        $identity | Should Not Match 'title|clipboard|\r|\n'
    }
    It 'retains distinct immediate identities and the saved original in the actual refusal builder' {
        $boundary = ('RedSalamander.Testing.HostedDesktopNative' -as [type]).Assembly.GetType('RedSalamander.Testing.HostedForegroundBoundary', $true)
        $refuse = $boundary.GetMethod('ForeignForegroundRefusal', [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic)
        $failure = $refuse.Invoke($null, [object[]]@([IntPtr]0x22, [IntPtr]0x11, [IntPtr]0x33, 'saved-original-identity', [RedSalamander.Testing.HostedIdentityDiagnosticControl]::Recorder))
        ($failure -is [InvalidOperationException]) | Should Be $true
        $failure.Message.StartsWith('A foreign foreground change must not be replaced by the hosted fixture.') | Should Be $true
        $failure.Message | Should Match 'capturedUtc=\d{4}-\d{2}-\d{2}T'
        $failure.Message | Should Match 'current=\{observed-hwnd=0x22\}; originalAtCapture=\{saved-original-identity\}; originalNow=\{observed-hwnd=0x11\}; boundary=\{observed-hwnd=0x33\}'
        ([RedSalamander.Testing.HostedIdentityDiagnosticControl]::Seen -join ',') | Should Be '34,17,51'
    }
    It 'keeps the safety refusal when diagnostic capture fails and omits private exception messages' {
        $boundary = ('RedSalamander.Testing.HostedDesktopNative' -as [type]).Assembly.GetType('RedSalamander.Testing.HostedForegroundBoundary', $true)
        $flags = [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic
        $capture = $boundary.GetMethod('CaptureIdentityForDiagnostics', $flags)
        $saved = $capture.Invoke($null, [object[]]@([IntPtr]0x11, [RedSalamander.Testing.HostedIdentityDiagnosticControl]::Unavailable))
        $failure = $boundary.GetMethod('ForeignForegroundRefusal', $flags).Invoke($null, [object[]]@([IntPtr]0x22, [IntPtr]0x11, [IntPtr]0x33, $saved, [RedSalamander.Testing.HostedIdentityDiagnosticControl]::Unavailable))
        ($failure -is [InvalidOperationException]) | Should Be $true
        $failure.Message.StartsWith('A foreign foreground change must not be replaced by the hosted fixture.') | Should Be $true
        $failure.Message | Should Match 'current=\{hwnd=0x22; capturedUtc=.*unavailable=InvalidOperationException\}'
        $failure.Message | Should Not Match 'seeded private diagnostic text'
        ([RedSalamander.Testing.HostedIdentityDiagnosticControl]::Seen -join ',') | Should Be '17,34,17,51'
    }
    It 'does not swallow a fatal CLR diagnostic failure' {
        $boundary = ('RedSalamander.Testing.HostedDesktopNative' -as [type]).Assembly.GetType('RedSalamander.Testing.HostedForegroundBoundary', $true)
        $refuse = $boundary.GetMethod('ForeignForegroundRefusal', [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic)
        $caught = $null
        try { [void]$refuse.Invoke($null, [object[]]@([IntPtr]0x22, [IntPtr]0x11, [IntPtr]0x33, 'saved', [RedSalamander.Testing.HostedIdentityDiagnosticControl]::Fatal)) }
        catch [System.Management.Automation.MethodInvocationException] { $caught = $_.Exception.GetBaseException() }
        ($caught -is [OutOfMemoryException]) | Should Be $true
    }
    It 'uses the identity builder only on the unchanged foreign foreground refusal branch' {
        $source = Get-Content -LiteralPath (Join-Path $repoRoot 'Tools/Modules/Testing/HostedTestDesktop.psm1') -Raw
        $source | Should Match 'current => ForeignForegroundRefusal\(current, original, handle, originalAtCapture, HostedDesktopNative.DescribeWindowIdentity\)'
        $source | Should Match 'if \(current != original\)\s+throw foreignRefusal\(current\);'
        $source | Should Match 'if \(current != handle\) throw foreignRefusal\(current\);'
        [regex]::Matches($source, 'throw foreignRefusal\(current\)').Count | Should Be 2
        $start = $source.IndexOf('internal static string DescribeWindowIdentity(')
        $end = $source.IndexOf('// ABI/layer exception:', $start)
        $observation = $source.Substring($start, $end - $start)
        $observation | Should Not Match 'SetForegroundWindow|SetFocus|SendInput|GetWindowText|Sleep|Kill|CloseWindow'
    }
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
        Mock Get-RSHostedStockSurface {
            $surface=$global:rsWslSurfaceQueue.Dequeue()
            $surface | Add-Member NoteProperty CapturedUtc '2026-11-14T00:00:00.0000000Z' -Force
            $surface | Add-Member NoteProperty CreationKnown $true -Force
            $surface | Add-Member NoteProperty CreationFileTime100ns 200 -Force
            $surface | Add-Member NoteProperty CreationError 0 -Force
            $surface
        } -ModuleName HostedTestDesktop
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
        foreach ($fixtureTimeout in @(1.0,2.0)) {
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
        $sandbox = Join-Path $TestDrive ('v1-proof-' + $fixtureTimeout)
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
            -Entries @(Get-RSTestRunPlan -Suite Commands -RepoRoot $repoRoot -Platform ARM64 -Configuration Debug -RedSalamanderExePath $exe -CaseFilter terminal_ -TimeoutMultiplier $fixtureTimeout)
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
        $native=@{run_started_utc='2026-11-13T12:00:06Z';suite='Commands';timeout_scale=$fixtureTimeout;passed=2;failed=0;skipped=0;cases=@(@{name='terminal_one';status='passed'},@{name='terminal_two';status='passed'})}
        $nativePath=Join-Path $last 'commands/results.json'
        $native|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $nativePath
        $exitPath=Join-Path $run ('artifacts/native-process-exits/'+('f'*32)+'.json')
        @{schema='red-salamander.native-process-exit.v1';process_id=42;executable=$exe;observed_utc='2026-11-13T12:00:54Z';exit_code=0}|ConvertTo-Json|Set-Content -LiteralPath $exitPath
        $metricPath=Join-Path $last 'perf/perf_metrics.jsonl'
        '{"process":42,"timestamp":"2026-11-13T12:00:50Z"}'|Set-Content -LiteralPath $metricPath
        $summary=@{schema='red-salamander.run-all-tests.v1';run_id=$runId;suite='Commands';case_filter='terminal_';platform='ARM64';configuration='Debug';executable=$exe;
            run_started_utc='2026-11-13T12:00:05Z';run_ended_utc='2026-11-13T12:00:55Z';timeout_scale=$fixtureTimeout;exit_code=2;total=2;passed=2;failed=0;skipped=0;
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
        $invoke={
            param($ExpectedTimeout=$fixtureTimeout)
            $arguments=@{ExitCode=2;Output="Run id: $runId";TestRoot=$sandbox;RepoRoot=$repoRoot;Suite='Commands';CaseFilter='terminal_';Platform='ARM64';Configuration='Debug';
                InvocationStartedUtc=([DateTimeOffset]'2026-11-13T12:00:00Z');InvocationEndedUtc=([DateTimeOffset]'2026-11-13T12:01:00Z')}
            if($null -ne $ExpectedTimeout){$arguments.TimeoutMultiplier=$ExpectedTimeout}
            Assert-RSHostedFilteredCaseRun @arguments
        }
        Seal-Proof
        (& $invoke).RepositoryVerdict | Should Be 'NOT_EVALUATED'
        if($fixtureTimeout -eq 1.0){(& $invoke $null).Passed | Should Be 2}
        else {
            {& $invoke $null}|Should Throw 'trusted invocation'
            {& $invoke 1.0}|Should Throw 'trusted invocation'
            {& $invoke 3.0}|Should Throw 'trusted invocation'
            # Untrusted matching summary/native claims cannot authorize a different archived plan.
            $summary.timeout_scale=3.0;$native.timeout_scale=3.0
            $summary|ConvertTo-Json -Depth 12|Set-Content $summaryPath
            $native|ConvertTo-Json -Depth 8|Set-Content $nativePath;Seal-Proof
            {& $invoke 3.0}|Should Throw 'plan, profile'
            $summary.timeout_scale=2.0;$native.timeout_scale=1.0
            $summary|ConvertTo-Json -Depth 12|Set-Content $summaryPath
            $native|ConvertTo-Json -Depth 8|Set-Content $nativePath;Seal-Proof
            {& $invoke}|Should Throw 'trusted invocation'
            $native.timeout_scale=2.0;$native|ConvertTo-Json -Depth 8|Set-Content $nativePath;Seal-Proof
            foreach($bad in @(0.0,101.0,[double]::NaN,[double]::PositiveInfinity)){
                {& $invoke $bad}|Should Throw 'TimeoutMultiplier'
            }
        }
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
            timeout_scale = 1.0
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
        $native = @{suite='Commands';timeout_scale=1.0;passed=3;failed=0;skipped=0;cases=@(
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

Describe 'Hosted stock captured-instance diagnostics' {
    BeforeEach {
        & (Get-Module HostedTestDesktop) { Initialize-RSHostedStockSurfaceNative }
        $script:stockType='RedSalamander.Testing.HostedStockSurface' -as [type]
        $script:observerType='RedSalamander.Testing.HostedDesktopNative' -as [type]
        $script:nonpublic=[Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic
    }
    function New-HeldObservationFixture {
        $child=[RedSalamander.Testing.HostedProcessFacts]::new()
        $child.Opened=$true;$child.HandleProcess=42;$child.CreationAvailable=$true;$child.CreationTime=200;$child.ParentAvailable=$true;$child.ParentProcess=41
        $parent=[RedSalamander.Testing.HostedProcessFacts]::new()
        $parent.Opened=$true;$parent.HandleProcess=41;$parent.CreationAvailable=$true;$parent.CreationTime=100
        [pscustomobject]@{Lease=[Microsoft.Win32.SafeHandles.SafeProcessHandle]::new([IntPtr]123,$false);Child=$child;Parent=$parent;
            Creation=200;FinalThread=[uint32]7;FinalPid=[uint32]42;RelationError=0;HeldReads=0;ParentReads=0;WindowReads=0;SameLease=$true;LeaseLive=$true}
    }
    function Invoke-HeldObservationFixture($f, $StockSurface=$null) {
        $held=[Func[Microsoft.Win32.SafeHandles.SafeProcessHandle,uint32,bool,RedSalamander.Testing.HostedProcessFacts]] {
            param($lease,$pidValue,$extended)
            $f.HeldReads++;$f.SameLease=$f.SameLease -and [object]::ReferenceEquals($lease,$f.Lease)
            $f.LeaseLive=$f.LeaseLive -and -not $lease.IsClosed
            if($pidValue -ne 42 -or -not $extended){throw 'Wrong captured child query.'};$f.Child
        }.GetNewClosure()
        $parent=[Func[uint32,bool,RedSalamander.Testing.HostedProcessFacts]] {
            param($pidValue,$extended)
            $f.ParentReads++;if($pidValue -ne 41 -or $extended){throw 'Wrong parent query.'};$f.Parent
        }.GetNewClosure()
        $window=[Func[IntPtr,Tuple[uint32,uint32,int]]] {
            param($windowValue)
            $f.WindowReads++;if($windowValue.ToInt64() -ne 291){throw 'Wrong captured window.'}
            $f.LeaseLive=$f.LeaseLive -and -not $f.Lease.IsClosed
            [Tuple[uint32,uint32,int]]::new($f.FinalThread,$f.FinalPid,$f.RelationError)
        }.GetNewClosure()
        $method=$script:observerType.GetMethod('ObserveHeldProcessForWindow',$script:nonpublic)
        $arguments=[object[]]@([IntPtr]291,[uint32]7,[uint32]42,$f.Lease,[long]$f.Creation,$true,$held,$parent,$window)
        if($null -ne $StockSurface) {
            $callback=[Func[IntPtr,uint32,uint32,Microsoft.Win32.SafeHandles.SafeProcessHandle,long,bool,string]] {
                param($windowValue,$threadValue,$pidValue,$lease,$creation,$extended)
                if($windowValue.ToInt64() -ne 291 -or $threadValue -ne 7 -or $pidValue -ne 42 -or -not [object]::ReferenceEquals($lease,$f.Lease)){throw 'Captured identity changed.'}
                $arguments[4]=[long]$creation
                $method.Invoke($null,$arguments)
            }.GetNewClosure()
            return $StockSurface.DescribeCapturedProcess($callback)
        }
        $method.Invoke($null,$arguments)
    }
    function New-FakeCapturedStock($lease) {
        $ctor=$script:stockType.GetConstructors([Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)[0]
        $surface=$ctor.Invoke([object[]]@([IntPtr]291,[uint32]7,[uint32]42,[uint32]3,$lease,'private title','fixture','private path','private app','Unknown'))
        $script:stockType.GetProperty('CreationKnown').SetValue($surface,$true)
        $script:stockType.GetProperty('CreationFileTime100ns').SetValue($surface,[long]200)
        $surface
    }
    It 'borrows the exact capture lease through one child parent and final-window query without disposing it' {
        $f=New-HeldObservationFixture
        try {
            $text=Invoke-HeldObservationFixture $f
            $text | Should Match 'instanceAvailable=True.*parentCreationCompatible=True.*relationStable=True'
            $text | Should Match 'capturedCreationFileTime100ns=200; creationMatchesCapture=True; childReads=1; parentReads=1; relationReads=1'
            $f.SameLease | Should Be $true;$f.LeaseLive | Should Be $true;$f.Lease.IsClosed | Should Be $false
            "$($f.HeldReads)/$($f.ParentReads)/$($f.WindowReads)" | Should Be '1/1/1'
        } finally {$f.Lease.Dispose()}
    }
    It 'rejects changed creation or PID and unavailable baseline identity before any parent query' {
        foreach($mode in @('creation','pid','missing-baseline','creation-denied')) {
            $f=New-HeldObservationFixture
            try {
                if($mode -eq 'creation'){$f.Child.CreationTime=201}
                if($mode -eq 'pid'){$f.Child.HandleProcess=99}
                if($mode -eq 'missing-baseline'){$f.Creation=-1}
                if($mode -eq 'creation-denied'){$f.Child.CreationAvailable=$false;$f.Child.CreationError=5}
                $text=Invoke-HeldObservationFixture $f
                $text | Should Match 'instanceAvailable=False.*parentQueried=False.*creationMatchesCapture=False'
                "$($f.HeldReads)/$($f.ParentReads)/$($f.WindowReads)" | Should Be '1/0/1'
            } finally {$f.Lease.Dispose()}
        }
    }
    It 'retains changed-window denial and newer-parent facts without upgrading them into authority' {
        $f=New-HeldObservationFixture
        try {
            $f.FinalPid=99;$f.Parent.CreationTime=300;$f.Child.AnyJobAvailable=$true;$f.Child.AnyJob=$true
            $f.Child.ImageError=5;$f.Child.SessionError=87
            $text=Invoke-HeldObservationFixture $f
            $text | Should Match 'imageKnown=False; imageError=5'
            $text | Should Match 'sessionKnown=False; session=0; sessionError=87; anyJobKnown=True; anyJob=True; jobError=0; exactTestJob=unavailable'
            $text | Should Match 'parentCreationCompatible=False; relationStable=False; finalPid=99'
            $text | Should Match 'not creator/ownership proof'
        } finally {$f.Lease.Dispose()}
    }
    It 'passes the captured tuple lease and creation through the actual stock bridge without leaking classification fields' {
        $lease=[Microsoft.Win32.SafeHandles.SafeProcessHandle]::new([IntPtr]123,$false)
        $surface=New-FakeCapturedStock $lease
        try {
            $seen=[Collections.Generic.List[string]]::new()
            $callback=[Func[IntPtr,uint32,uint32,Microsoft.Win32.SafeHandles.SafeProcessHandle,long,bool,string]] {
                param($window,$threadValue,$pidValue,$held,$creation,$extended)
                $seen.Add("$($window.ToInt64())/$threadValue/$pidValue/$creation/$extended/$([object]::ReferenceEquals($held,$lease))")
                'safe metadata'
            }.GetNewClosure()
            $surface.DescribeCapturedProcess($callback) | Should Be 'safe metadata'
            ($seen -join ',') | Should Be '291/7/42/200/True/True'
            $script:stockType.GetProperty('CreationFileTime100ns').SetValue($surface,[long]0)
            $surface.DescribeCapturedProcess($callback) | Should Be 'safe metadata'
            $seen[1] | Should Be '291/7/42/-1/True/True'
            $script:stockType.GetProperty('CreationFileTime100ns').SetValue($surface,[long]200)
            $script:stockType.GetProperty('CreationKnown').SetValue($surface,$false)
            $surface.DescribeCapturedProcess($callback) | Should Be 'safe metadata'
            $seen[2] | Should Be '291/7/42/-1/True/True'
            $surface.Kind | Should Be 'Unknown';$surface.CloseRequested | Should Be $false
            $lease.IsClosed | Should Be $false
        } finally {$surface.Dispose()}
    }
    It 'keeps denial errors and untrusted image fields private through the borrowed-lease observer' {
        $f=New-HeldObservationFixture
        try {
            $f.Child.Opened=$false;$f.Child.OpenError=5;$f.Child.CreationAvailable=$false;$f.Child.CreationError=87
            $f.Child.ImageAvailable=$true;$f.Child.ImageKind='private-image.exe';$f.Child.DigestAvailable=$true;$f.Child.ImageDigest='private path'
            $text=Invoke-HeldObservationFixture $f
            $text | Should Match 'instanceAvailable=False; opened=False; openError=5'
            $text | Should Match 'creationKnown=False; creationFileTime100ns=200; creationError=87'
            $text | Should Match 'imageKind=other; pathDigestAlgorithm=SHA256-UTF16LE; pathDigestKnown=False; pathDigest=unavailable'
            $text | Should Not Match 'private-image|private path'
            $f.ParentReads | Should Be 0;$f.WindowReads | Should Be 1
        } finally {$f.Lease.Dispose()}
    }
    It 'rejects zero and missing stock creation through the real bridge and canonical fake-reader observer' {
        $f=New-HeldObservationFixture;$surface=New-FakeCapturedStock $f.Lease
        try {
            foreach($mode in @('zero','missing')) {
                $f.HeldReads=0;$f.ParentReads=0;$f.WindowReads=0
                $script:stockType.GetProperty('CreationKnown').SetValue($surface,($mode -eq 'zero'))
                $script:stockType.GetProperty('CreationFileTime100ns').SetValue($surface,[long]0)
                $text=Invoke-HeldObservationFixture $f $surface
                $text | Should Match 'instanceAvailable=False.*parentQueried=False.*capturedCreationFileTime100ns=-1; creationMatchesCapture=False'
                "$($f.HeldReads)/$($f.ParentReads)/$($f.WindowReads)" | Should Be '1/0/1'
                $surface.Kind | Should Be 'Unknown';$surface.CloseRequested | Should Be $false
            }
        } finally {$surface.Dispose()}
    }
    It 'keeps unavailable diagnostic errors private and preserves direct fatal propagation' {
        $lease=[Microsoft.Win32.SafeHandles.SafeProcessHandle]::new([IntPtr]123,$false)
        $surface=New-FakeCapturedStock $lease
        try {
            foreach($exceptionType in @([InvalidOperationException],[OutOfMemoryException])) {
                $delegateType=[Func[IntPtr,uint32,uint32,Microsoft.Win32.SafeHandles.SafeProcessHandle,long,bool,string]]
                $method=[Reflection.Emit.DynamicMethod]::new('CapturedStockQueryFailure',[string],[Type[]]@([IntPtr],[uint32],[uint32],[Microsoft.Win32.SafeHandles.SafeProcessHandle],[long],[bool]))
                $il=$method.GetILGenerator();$il.Emit([Reflection.Emit.OpCodes]::Newobj,$exceptionType.GetConstructor([Type]::EmptyTypes));$il.Emit([Reflection.Emit.OpCodes]::Throw)
                $callback=$method.CreateDelegate($delegateType)
                if($exceptionType -eq [InvalidOperationException]) {
                    $surface.DescribeCapturedProcess($callback) | Should Be 'instanceAvailable=False; unavailable=InvalidOperationException'
                } else {
                    $caught=$null;try{[void]$surface.DescribeCapturedProcess($callback)}catch{$caught=$_.Exception.GetBaseException()}
                    $caught.GetType() | Should Be ([OutOfMemoryException])
                }
                $surface.CloseRequested | Should Be $false
            }
        } finally {$surface.Dispose()}
    }
    It 'initializes both real managed types and their exact direct delegate in a fresh noninteractive PowerShell' {
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestSandbox.psm1')
        $scratch=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case ('hosted-stock-import-'+[guid]::NewGuid().ToString('N'))
        $childScript=Join-Path $scratch 'import-bridge.ps1'
        [IO.File]::WriteAllText($childScript,@'
param($Repo)
$ErrorActionPreference='Stop'
Import-Module (Join-Path $Repo 'Tools/Modules/Testing/HostedTestDesktop.psm1')
& (Get-Module HostedTestDesktop) { Initialize-RSHostedStockSurfaceNative }
$type=[Func[IntPtr,uint32,uint32,Microsoft.Win32.SafeHandles.SafeProcessHandle,long,bool,string]]
$callback=[RedSalamander.Testing.HostedDesktopNative].GetMethod('DescribeHeldProcess').CreateDelegate($type)
if($callback.Method.Name -cne 'DescribeHeldProcess'){throw 'Wrong managed bridge.'}
$ctor=[RedSalamander.Testing.HostedStockSurface].GetConstructors([Reflection.BindingFlags]::Instance -bor [Reflection.BindingFlags]::NonPublic)[0]
$surface=$ctor.Invoke([object[]]@([IntPtr]0,[uint32]0,[uint32]0,[uint32]0,$null,'','','','','Empty'))
try {
    $result=$surface.DescribeCapturedProcess($callback)
    if($result -cne 'instanceAvailable=False; unavailable=InvalidWindowIdentity'){throw 'Unexpected unavailable identity.'}
} finally {$surface.Dispose()}
'fresh-managed-bridge-passed'
'@,[Text.UTF8Encoding]::new($false))
        $start=[Diagnostics.ProcessStartInfo]::new();$start.FileName=(Get-Process -Id $PID).Path
        $start.UseShellExecute=$false;$start.CreateNoWindow=$true;$start.RedirectStandardOutput=$true;$start.RedirectStandardError=$true
        foreach($arg in @('-NoProfile','-NonInteractive','-File',$childScript,$repoRoot)){[void]$start.ArgumentList.Add($arg)}
        $child=[Diagnostics.Process]::new();$child.StartInfo=$start
        try {
            [void]$child.Start();$stdout=$child.StandardOutput.ReadToEndAsync();$stderr=$child.StandardError.ReadToEndAsync()
            $child.WaitForExit();$output=$stdout.GetAwaiter().GetResult();$errorText=$stderr.GetAwaiter().GetResult()
            [IO.File]::WriteAllText((Join-Path $scratch 'stdout.txt'),$output);[IO.File]::WriteAllText((Join-Path $scratch 'stderr.txt'),$errorText)
            $child.ExitCode | Should Be 0;$output | Should Match 'fresh-managed-bridge-passed';$errorText | Should BeNullOrEmpty
        } finally {$child.Dispose()}
    }
    It 'records bounded synthetic work counts without treating fake timings as native-query performance' {
        Import-Module (Join-Path $repoRoot 'Tools/Modules/Testing/TestSandbox.psm1')
        $scratch=New-RSTestSandboxScratchDirectory -RepoRoot $repoRoot -Harness 'tools-pester' -Case 'hosted-stock-work-counts'
        $rows=[Collections.Generic.List[object]]::new()
        foreach($scenario in @('matching','changed-creation','changed-window','newer-parent','denied')) {
            for($iteration=1;$iteration -le 8;$iteration++) {
                $f=New-HeldObservationFixture
                try {
                    if($scenario -eq 'changed-creation'){$f.Child.CreationTime=201}
                    if($scenario -eq 'changed-window'){$f.FinalPid=99}
                    if($scenario -eq 'newer-parent'){$f.Parent.CreationTime=300}
                    if($scenario -eq 'denied'){$f.Child.Opened=$false;$f.Child.OpenError=5;$f.Child.CreationAvailable=$false}
                    $timer=[Diagnostics.Stopwatch]::StartNew();$text=Invoke-HeldObservationFixture $f;$timer.Stop()
                    $expectedParents=if($scenario -in @('changed-creation','denied')){0}else{1}
                    "$($f.HeldReads)/$($f.ParentReads)/$($f.WindowReads)" | Should Be "1/$expectedParents/1"
                    $f.SameLease | Should Be $true;$f.LeaseLive | Should Be $true
                    $rows.Add([ordered]@{scenario=$scenario;iteration=$iteration;held_child_reads=$f.HeldReads;parent_reads=$f.ParentReads;
                        final_relation_reads=$f.WindowReads;same_lease=$f.SameLease;lease_live=$f.LeaseLive;elapsed_ticks=$timer.ElapsedTicks})
                } finally {$f.Lease.Dispose()}
            }
        }
        [ordered]@{schema='red-salamander.hosted-stock-fake-work.v1';run_id=$env:REDSALAMANDER_TEST_RUN_ID;
            clock_frequency=[Diagnostics.Stopwatch]::Frequency;scope='managed fake readers only; not native-query or pipeline performance';rows=@($rows)} |
            ConvertTo-Json -Depth 7 | Set-Content -LiteralPath (Join-Path $scratch 'work-counts.json') -Encoding utf8NoBOM
    }
}

Describe 'Hosted process attribution diagnostics' {
    BeforeEach {
        & (Get-Module HostedTestDesktop) { Initialize-RSHostedDesktopNative }
        $script:processObserver = ('RedSalamander.Testing.HostedDesktopNative' -as [type]).GetMethod('ObserveProcessForWindow', [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic)
    }
    function New-ProcessObservationFixture {
        $child = [RedSalamander.Testing.HostedProcessFacts]::new()
        $child.Opened = $true; $child.HandleProcess = 42; $child.CreationAvailable = $true; $child.CreationTime = 200
        $child.ParentAvailable = $true; $child.ParentProcess = 41
        $parent = [RedSalamander.Testing.HostedProcessFacts]::new()
        $parent.Opened = $true; $parent.HandleProcess = 41; $parent.CreationAvailable = $true; $parent.CreationTime = 100
        [pscustomobject]@{Child=$child;Parent=$parent;FinalPid=[uint32]42;FinalThread=[uint32]7;RelationError=0;Fatal=$false;Unavailable=$false;
            Reads=[Collections.Generic.List[string]]::new();Windows=[Collections.Generic.List[long]]::new()}
    }
    function Invoke-ProcessObservationFixture($fixture, [bool]$extended=$true) {
        $read = [Func[uint32,bool,RedSalamander.Testing.HostedProcessFacts]] {
            param($pidValue,$full)
            $fixture.Reads.Add("$pidValue/$full")
            if ($fixture.Fatal) { throw [OutOfMemoryException]::new('private fatal text') }
            if ($fixture.Unavailable) { throw [InvalidOperationException]::new('private query text') }
            if ($pidValue -eq 42) { return $fixture.Child }
            if ($pidValue -eq 41 -and -not $full) { return $fixture.Parent }
            throw 'Unexpected process query.'
        }.GetNewClosure()
        $readWindow = [Func[IntPtr,Tuple[uint32,uint32,int]]] {
            param($window)
            $fixture.Windows.Add($window.ToInt64())
            [Tuple[uint32,uint32,int]]::new($fixture.FinalThread,$fixture.FinalPid,$fixture.RelationError)
        }.GetNewClosure()
        $script:processObserver.Invoke($null, [object[]]@([IntPtr]0x123,[uint32]7,[uint32]42,$extended,$read,$readWindow))
    }
    It 'binds exact child creation and only one minimal parent candidate through the actual observer' {
        $f=New-ProcessObservationFixture; $text=Invoke-ProcessObservationFixture $f
        $text | Should Match 'instanceAvailable=True.*handlePid=42.*creationFileTime100ns=200'
        $text | Should Match 'parentQueried=True.*handlePid=41.*creationFileTime100ns=100.*parentCreationCompatible=True'
        $text | Should Match 'relationStable=True; finalPid=42; finalThread=7; relationError=0'
        ($f.Reads -join ',') | Should Be '42/True,41/False'
        ($f.Windows -join ',') | Should Be '291'
        $text | Should Match 'not creator/ownership proof'
    }
    It 'keeps the entry-time query minimal and distinct from later extended observations' {
        $f=New-ProcessObservationFixture; $text=Invoke-ProcessObservationFixture $f $false
        ($f.Reads -join ',') | Should Be '42/False'
        $text | Should Match 'extended=False; instanceAvailable=True'
        $text | Should Not Match 'imageKind|pathDigest|anyJobKnown'
        $source=Get-Content -LiteralPath (Join-Path $repoRoot 'Tools/Modules/Testing/HostedTestDesktop.psm1') -Raw
        $source | Should Match 'originalAtCapture = CaptureIdentityForDiagnostics\(original, HostedDesktopNative.DescribeEntryWindowIdentity\)'
        $source | Should Match '(?s)DescribeEntryWindowIdentity\(IntPtr window\)\s*\{\s*return DescribeWindowIdentityCore\(window, false\);'
    }
    It 'refuses to join a reused parent PID to an older child creation identity' {
        $f=New-ProcessObservationFixture; $f.Parent.CreationTime=300
        (Invoke-ProcessObservationFixture $f) | Should Match 'parentCreationCompatible=False'
    }
    It 'never opens zero or self as a parent candidate' {
        foreach($parent in @(0,42)) {
            $f=New-ProcessObservationFixture;$f.Child.ParentProcess=$parent
            (Invoke-ProcessObservationFixture $f) | Should Match 'parentQueried=False'
            ($f.Reads -join ',') | Should Be '42/True'
        }
    }
    It 'does not query a parent after a mismatched process handle or missing creation identity' {
        foreach($mode in @('wrong-handle','missing-creation','zero-creation')) {
            $f=New-ProcessObservationFixture
            if($mode -eq 'wrong-handle'){$f.Child.HandleProcess=99}
            if($mode -eq 'missing-creation'){$f.Child.CreationAvailable=$false;$f.Child.CreationError=5}
            if($mode -eq 'zero-creation'){$f.Child.CreationTime=0}
            $text=Invoke-ProcessObservationFixture $f
            $text | Should Match 'instanceAvailable=False'
            $text | Should Match 'parentQueried=False'
            ($f.Reads -join ',') | Should Be '42/True'
        }
    }
    It 'retains independent denial and query errors instead of treating missing facts as negative membership' {
        $f=New-ProcessObservationFixture
        $f.Child.Opened=$false;$f.Child.OpenError=5;$f.Child.IdentityError=6;$f.Child.CreationAvailable=$false;$f.Child.CreationError=87
        $f.Child.ImageError=122;$f.Child.SessionError=1168;$f.Child.JobError=50;$f.Child.ParentStatus=-1073741790
        $text=Invoke-ProcessObservationFixture $f
        $text | Should Match 'opened=False; openError=5; handlePid=42; identityError=6; creationKnown=False; creationFileTime100ns=200; creationError=87'
        $text | Should Match 'imageKnown=False; imageError=122'
        $text | Should Match 'sessionKnown=False; session=0; sessionError=1168; anyJobKnown=False; anyJob=False; jobError=50; exactTestJob=unavailable'
        $text | Should Match 'parentNtStatus=0xC0000022'
    }
    It 'records changed or vanished captured HWND relations without another foreground selection' {
        foreach($sample in @(@{pid=99;thread=7;error=0},@{pid=0;thread=0;error=1400})) {
            $f=New-ProcessObservationFixture;$f.FinalPid=$sample.pid;$f.FinalThread=$sample.thread;$f.RelationError=$sample.error
            $text=Invoke-ProcessObservationFixture $f
            $text | Should Match 'relationStable=False'
            $text | Should Match ('relationError='+$sample.error)
            ($f.Windows -join ',') | Should Be '291'
        }
    }
    It 'never upgrades membership in any Job into exact test containment or authority' {
        $f=New-ProcessObservationFixture;$f.Child.AnyJobAvailable=$true;$f.Child.AnyJob=$true
        (Invoke-ProcessObservationFixture $f) | Should Match 'anyJobKnown=True; anyJob=True; jobError=0; exactTestJob=unavailable'
    }
    It 'omits arbitrary image paths names and malformed digests even from an injected backend' {
        $f=New-ProcessObservationFixture;$f.Child.ImageAvailable=$true;$f.Child.DigestAvailable=$true
        $f.Child.ImageKind='C:\private-fixture\account-secret.exe';$f.Child.ImageDigest='private invalid digest'
        $text=Invoke-ProcessObservationFixture $f
        $text | Should Match 'imageKind=other; pathDigestAlgorithm=SHA256-UTF16LE; pathDigestKnown=False; pathDigest=unavailable'
        $text | Should Not Match 'account-secret|private-fixture|private invalid|title|commandLine|clipboard'
    }
    It 'uses SHA256 of exact UTF16LE bytes without exposing the queried path' {
        $native='RedSalamander.Testing.HostedDesktopNative' -as [type]
        $flags=[Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic
        $path='C:\private-fixture\WindowsTerminal.exe'
        $digest=$native.GetMethod('ProcessImageDigest',$flags).Invoke($null,[object[]]@($path))
        $digest | Should Match '^[0-9a-f]{64}$'
        $digest | Should Be ([Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::Unicode.GetBytes($path))).ToLowerInvariant())
        $native.GetMethod('ProcessImageKind',$flags).Invoke($null,[object[]]@($path)) | Should Be 'WindowsTerminal.exe'
        $f=New-ProcessObservationFixture;$f.Child.ImageKind='WindowsTerminal.exe';$f.Child.ImageDigest=$digest;$f.Child.DigestAvailable=$true
        (Invoke-ProcessObservationFixture $f) | Should Not Match 'private-fixture'
        $unpaired='A'+[char]0xd800+'Z'
        $expected=[Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([byte[]]@(0x41,0,0,0xd8,0x5a,0))).ToLowerInvariant()
        $native.GetMethod('ProcessImageDigest',$flags).Invoke($null,[object[]]@($unpaired)) | Should Be $expected
    }
    It 'propagates direct managed reader exceptions without swallowing fatal failures' {
        foreach($exceptionType in @([InvalidOperationException],[OutOfMemoryException])) {
            $method=[Reflection.Emit.DynamicMethod]::new('ProcessObservationFailure',[RedSalamander.Testing.HostedProcessFacts],[Type[]]@([uint32],[bool]))
            $il=$method.GetILGenerator()
            $il.Emit([Reflection.Emit.OpCodes]::Newobj,$exceptionType.GetConstructor([Type]::EmptyTypes))
            $il.Emit([Reflection.Emit.OpCodes]::Throw)
            $read=$method.CreateDelegate([Func[uint32,bool,RedSalamander.Testing.HostedProcessFacts]])
            $readWindow=[Func[IntPtr,Tuple[uint32,uint32,int]]] { param($window) [Tuple[uint32,uint32,int]]::new(7,42,0) }
            $caught=$null
            try{[void]$script:processObserver.Invoke($null,[object[]]@([IntPtr]0x123,[uint32]7,[uint32]42,$true,$read,$readWindow))}catch{$caught=$_.Exception.GetBaseException()}
            $caught | Should Not BeNullOrEmpty
            $caught.GetType() | Should Be $exceptionType
        }
    }
}
