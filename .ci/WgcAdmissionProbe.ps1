[CmdletBinding()]
param([string]$TestRoot='D:\RedSalamander.Perf')
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$env:REDSALAMANDER_TEST_ROOT=$TestRoot
[Environment]::SetEnvironmentVariable('REDSALAMANDER_TEST_RUN_ID',$null,'Process')
Import-Module (Join-Path $repo 'Tools\Modules\Testing\TestSandbox.psm1') -Force
if($env:GITHUB_ACTIONS -ceq 'true' -and $env:RUNNER_ENVIRONMENT -ceq 'github-hosted' -and $env:GITHUB_REPOSITORY -ceq 'RedSalamanders/RedSalamander'){ Initialize-RSTestSandboxRoot -TestRoot $TestRoot -RepoRoot $repo -AllowExternalInitialization | Out-Null }
$taskTemp=New-RSTestSandboxScratchDirectory -RepoRoot $repo -Harness review-tools -Case wgc-desktop-admission-cloud-01
$env:TEMP=$taskTemp;$env:TMP=$taskTemp
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
public static class WgcDesktopAdmissionProbe {
    public static long OriginalForegroundBefore; public static long OriginalForegroundAfter;
    static int lastWidth, lastHeight, lastSizeHr;
    [StructLayout(LayoutKind.Sequential)] struct ItemSize { public int Width, Height; }
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int GetItemSize(IntPtr self,out ItemSize size);
    // Independent API control: creates only inactive alternate desktops, never
    // switches input desktops, injects input, or captures an image. Unlike the
    // hosted lease this is allowed locally and measures capture-item admission.
    public sealed class Desktop : Microsoft.Win32.SafeHandles.SafeHandleZeroOrMinusOneIsInvalid {
        public Desktop() : base(true) {}
        [DllImport("user32.dll", ExactSpelling=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool CloseDesktop(IntPtr h);
        protected override bool ReleaseHandle() { return CloseDesktop(handle); }
    }
    public sealed class Window : Microsoft.Win32.SafeHandles.SafeHandleZeroOrMinusOneIsInvalid {
        public Window() : base(true) {}
        [DllImport("user32.dll", ExactSpelling=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool DestroyWindow(IntPtr h);
        protected override bool ReleaseHandle() { return DestroyWindow(handle); }
    }
    public sealed class HString : Microsoft.Win32.SafeHandles.SafeHandleZeroOrMinusOneIsInvalid {
        public HString() : base(true) {}
        [DllImport("combase.dll", ExactSpelling=true)] static extern int WindowsDeleteString(IntPtr h);
        protected override bool ReleaseHandle() { return WindowsDeleteString(handle)>=0; }
    }
    public sealed class ComOwner : Microsoft.Win32.SafeHandles.SafeHandleZeroOrMinusOneIsInvalid {
        public ComOwner() : base(true) {}
        protected override bool ReleaseHandle() { Marshal.Release(handle); return true; }
    }
    [DllImport("combase.dll", ExactSpelling=true)] static extern int RoInitialize(uint type);
    [DllImport("combase.dll", ExactSpelling=true)] static extern void RoUninitialize();
    [DllImport("combase.dll", CharSet=CharSet.Unicode, ExactSpelling=true)] static extern int WindowsCreateString(string s,uint length,out HString h);
    [DllImport("combase.dll", ExactSpelling=true)] static extern int RoGetActivationFactory(HString name,ref Guid iid,out ComOwner factory);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)] delegate int CreateForWindow(IntPtr self,IntPtr hwnd,ref Guid iid,out ComOwner result);
    [DllImport("user32.dll", CharSet=CharSet.Unicode, ExactSpelling=true, SetLastError=true)] static extern Desktop CreateDesktopW(string name,IntPtr device,IntPtr mode,uint flags,uint access,IntPtr attributes);
    [DllImport("user32.dll", ExactSpelling=true, SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool SetThreadDesktop(Desktop desktop);
    [DllImport("user32.dll", ExactSpelling=true, SetLastError=true)] static extern Desktop OpenInputDesktop(uint flags,bool inherit,uint access);
    [DllImport("user32.dll", CharSet=CharSet.Unicode, ExactSpelling=true, SetLastError=true)] static extern Window CreateWindowExW(uint ex,string cls,string title,uint style,int x,int y,int w,int h,IntPtr parent,IntPtr menu,IntPtr module,IntPtr param);
    [DllImport("user32.dll", ExactSpelling=true)] static extern int GetSystemMetrics(int index);
    [DllImport("user32.dll", ExactSpelling=true)] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll", ExactSpelling=true)] static extern uint GetWindowThreadProcessId(IntPtr window,out uint process);
    [DllImport("user32.dll", CharSet=CharSet.Unicode, ExactSpelling=true)] static extern int GetWindowTextW(IntPtr window,StringBuilder title,int capacity);
    public static uint CurrentForegroundProcess() { uint process;GetWindowThreadProcessId(GetForegroundWindow(),out process);return process; }
    public static string CurrentForegroundTitle() {var title=new StringBuilder(1024);GetWindowTextW(GetForegroundWindow(),title,title.Capacity);return title.ToString();}
    [DllImport("user32.dll", ExactSpelling=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("dwmapi.dll", ExactSpelling=true)] static extern int DwmFlush();
    [DllImport("user32.dll", ExactSpelling=true,SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool AllowSetForegroundWindow(uint process);
    [DllImport("kernel32.dll", ExactSpelling=true)] static extern uint GetCurrentProcessId();
    [DllImport("user32.dll", ExactSpelling=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool SetForegroundWindow(IntPtr window);
    [DllImport("user32.dll", ExactSpelling=true)] static extern IntPtr SetFocus(IntPtr window);
    [DllImport("user32.dll", ExactSpelling=true)] static extern IntPtr GetFocus();
    [DllImport("user32.dll",CharSet=CharSet.Unicode,ExactSpelling=true,SetLastError=true)] static extern Desktop OpenDesktopW(string name,uint flags,bool inherit,uint access);
    [StructLayout(LayoutKind.Sequential)] struct Point { public int X,Y; }
    [StructLayout(LayoutKind.Sequential)] struct Message { public IntPtr Window;public uint Id;public UIntPtr WParam;public IntPtr LParam;public uint Time;public Point Position;public uint Private; }
    [DllImport("user32.dll",ExactSpelling=true,SetLastError=true)] static extern uint MsgWaitForMultipleObjectsEx(uint count,IntPtr[] handles,uint milliseconds,uint mask,uint flags);
    [DllImport("user32.dll",ExactSpelling=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool PeekMessageW(out Message message,IntPtr window,uint first,uint last,uint flags);
    [DllImport("user32.dll",ExactSpelling=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool TranslateMessage(ref Message message);
    [DllImport("user32.dll",ExactSpelling=true)] static extern IntPtr DispatchMessageW(ref Message message);
    public sealed class ForegroundResult {
        public string Stage;public bool SetForegroundReturned,ActualForeground,ActualFocus;public int CaptureHr,Width,Height;
    }
    static DefaultAnchor anchor;
    public static bool GrantSelfOnActiveInputDesktop() {
        bool granted=false;Exception failure=null;
        var t=new Thread(()=>{try { using(var d=OpenInputDesktop(0,false,0x1ff)) {
            if(d==null || d.IsInvalid || !SetThreadDesktop(d)) throw new Win32Exception(Marshal.GetLastWin32Error());
            granted=AllowSetForegroundWindow(GetCurrentProcessId());
        }}catch(Win32Exception e){failure=e;}});
        t.Start();if(!t.Join(TimeSpan.FromSeconds(10)))throw new TimeoutException("Owned grant worker timeout");
        if(failure!=null)throw new InvalidOperationException("Owned grant failed",failure);
        return granted;
    }
    public static void StartDefaultAnchor() { anchor=new DefaultAnchor(); }
    public static ForegroundResult ActivateDefaultAnchor(string stage) { return anchor.Activate(stage); }
    public static void StopDefaultAnchor() { if(anchor!=null){anchor.Dispose();anchor=null;} }
    sealed class DefaultAnchor : IDisposable {
        readonly AutoResetEvent request=new AutoResetEvent(false),stop=new AutoResetEvent(false);
        readonly ManualResetEventSlim ready=new ManualResetEventSlim(false),finished=new ManualResetEventSlim(false);
        readonly Thread thread;Exception failure;ForegroundResult result;string stage;
        public DefaultAnchor() { thread=new Thread(Run){IsBackground=true};thread.Start();if(!ready.Wait(TimeSpan.FromSeconds(10)))throw new TimeoutException("Default owned window startup timeout");if(failure!=null)throw new InvalidOperationException("Default window startup failed",failure); }
        void Run() {
            try { using(var desktop=OpenDesktopW("Default",0,false,0x1ff)) {
                if(desktop==null || desktop.IsInvalid || !SetThreadDesktop(desktop)) throw new Win32Exception(Marshal.GetLastWin32Error());
                int initialized=RoInitialize(1);if(initialized<0)Marshal.ThrowExceptionForHR(initialized);
                try { using(var window=CreateWindowExW(0x00040000,"STATIC","Owned default foreground control",0x90000000,Math.Max(0,GetSystemMetrics(0)-12),Math.Max(0,GetSystemMetrics(1)-12),12,12,IntPtr.Zero,IntPtr.Zero,IntPtr.Zero,IntPtr.Zero)) {
                    if(window==null || window.IsInvalid)throw new Win32Exception(Marshal.GetLastWin32Error());
                    ready.Set();var waits=new[]{request.SafeWaitHandle.DangerousGetHandle(),stop.SafeWaitHandle.DangerousGetHandle()};
                    for(;;){uint wait=MsgWaitForMultipleObjectsEx(2,waits,0xffffffff,0x04ff,4);
                        if(wait==1)break;
                        if(wait==0){IntPtr h=window.DangerousGetHandle();bool set=SetForegroundWindow(h);SetFocus(h);int factory;int captured=Admit(h,out factory);
                            result=new ForegroundResult{Stage=stage,SetForegroundReturned=set,ActualForeground=GetForegroundWindow()==h,ActualFocus=GetFocus()==h,CaptureHr=captured,Width=lastWidth,Height=lastHeight};finished.Set();}
                        else if(wait==2){Message m;while(PeekMessageW(out m,IntPtr.Zero,0,0,1)){TranslateMessage(ref m);DispatchMessageW(ref m);}}
                        else throw new Win32Exception(Marshal.GetLastWin32Error());
                    }
                }}finally{RoUninitialize();}
            }}catch(Win32Exception e){failure=e;}catch(COMException e){failure=e;}finally{ready.Set();finished.Set();}
        }
        public ForegroundResult Activate(string value){stage=value;finished.Reset();request.Set();if(!finished.Wait(TimeSpan.FromSeconds(10)))throw new TimeoutException("Default activation timeout");if(failure!=null)throw new InvalidOperationException("Default owned activation failed",failure);return result;}
        public void Dispose(){stop.Set();if(!thread.Join(TimeSpan.FromSeconds(10)))throw new TimeoutException("Default control retirement timeout");ready.Dispose();finished.Dispose();request.Dispose();stop.Dispose();if(failure!=null)throw new InvalidOperationException("Default control failed",failure);}
    }
    public sealed class Result {
        public string Desktop; public uint ExtendedStyle; public int FactoryHr; public int ItemHr;
        public int FlushHr; public bool Visible; public bool? SameDesktopForegroundUnchanged;
        public string Apartment; public int ItemWidth, ItemHeight, ItemSizeHr;
    }
    static int Admit(IntPtr hwnd,out int factoryHr) {
        lastWidth=0;lastHeight=0;lastSizeHr=unchecked((int)0x8000000a);
        const string cls="Windows.Graphics.Capture.GraphicsCaptureItem";
        HString name; int hr=WindowsCreateString(cls,(uint)cls.Length,out name);
        factoryHr=hr;
        using(name) {
            if(hr<0) return hr;
            Guid interopId=new Guid("3628e81b-3cac-4c60-b7f4-23ce0e0c3356");
            ComOwner factory; factoryHr=RoGetActivationFactory(name,ref interopId,out factory);
            using(factory) {
                if(factoryHr<0) return factoryHr;
                IntPtr table=Marshal.ReadIntPtr(factory.DangerousGetHandle());
                var create=(CreateForWindow)Marshal.GetDelegateForFunctionPointer(Marshal.ReadIntPtr(table,3*IntPtr.Size),typeof(CreateForWindow));
                Guid itemId=new Guid("79c3f95b-31f7-4ec2-a464-632ef5d30760");
                ComOwner item; hr=create(factory.DangerousGetHandle(),hwnd,ref itemId,out item);
                using(item) {
                    if(hr>=0 && item!=null && !item.IsInvalid) {
                        IntPtr itemTable=Marshal.ReadIntPtr(item.DangerousGetHandle());
                        var getSize=(GetItemSize)Marshal.GetDelegateForFunctionPointer(Marshal.ReadIntPtr(itemTable,7*IntPtr.Size),typeof(GetItemSize));
                        ItemSize size;lastSizeHr=getSize(item.DangerousGetHandle(),out size);lastWidth=size.Width;lastHeight=size.Height;
                    }
                    return hr;
                }
            }
        }
    }
    public static Result[] Run() {
        var results=new List<Result>();
        IntPtr foreground=GetForegroundWindow();
        OriginalForegroundBefore=foreground.ToInt64();
        string name="RedSalamander.Probe."+Guid.NewGuid().ToString("N");
        using(var desktop=CreateDesktopW(name,IntPtr.Zero,IntPtr.Zero,0,0x1ff,IntPtr.Zero)) using(var input=OpenInputDesktop(0,false,0x1ff)) {
            if(input==null || input.IsInvalid) throw new Win32Exception(Marshal.GetLastWin32Error(),"Retain input desktop for owned API control");
            if(desktop==null || desktop.IsInvalid) throw new Win32Exception(Marshal.GetLastWin32Error(),"Create inactive probe desktop");
            foreach(bool alternate in new[]{false,true}) foreach(uint apartment in new uint[]{0,1}) {
                Exception error=null;
                var thread=new Thread(()=>{
                    try {
                        if(!SetThreadDesktop(alternate?desktop:input)) throw new Win32Exception(Marshal.GetLastWin32Error(),"Attach explicit probe desktop");
                        int initialized=RoInitialize(apartment);
                        if(initialized<0) Marshal.ThrowExceptionForHR(initialized);
                        try {
                            foreach(uint style in new uint[]{0x08000080,0x08040080,0x08040000}) {
                                using(var window=CreateWindowExW(style,"STATIC","RedSalamander capture admission control",0x90000000,
                                    Math.Max(0,GetSystemMetrics(0)-12),Math.Max(0,GetSystemMetrics(1)-12),12,12,IntPtr.Zero,IntPtr.Zero,IntPtr.Zero,IntPtr.Zero)) {
                                    if(window==null || window.IsInvalid) throw new Win32Exception(Marshal.GetLastWin32Error(),"Create nonactivating owned probe window");
                                    int factoryHr;
                                    int flush=DwmFlush();
                                    int itemHr=Admit(window.DangerousGetHandle(),out factoryHr);
                                    results.Add(new Result{Desktop=alternate?"inactive-owned":"active-input",ExtendedStyle=style,FactoryHr=factoryHr,ItemHr=itemHr,FlushHr=flush,Visible=IsWindowVisible(window.DangerousGetHandle()),SameDesktopForegroundUnchanged=alternate?(bool?)null:GetForegroundWindow()==foreground,Apartment=apartment==0?"STA":"MTA",ItemWidth=lastWidth,ItemHeight=lastHeight,ItemSizeHr=lastSizeHr});
                                }
                            }
                            if(!alternate) {
                                int factoryHr;int hr=Admit(IntPtr.Zero,out factoryHr);
                                results.Add(new Result{Desktop="original-null-window-control",FactoryHr=factoryHr,ItemHr=hr,SameDesktopForegroundUnchanged=alternate?(bool?)null:GetForegroundWindow()==foreground,Apartment=apartment==0?"STA":"MTA",ItemWidth=lastWidth,ItemHeight=lastHeight,ItemSizeHr=lastSizeHr});
                            }
                        } finally { RoUninitialize(); }
                    } catch(Win32Exception e) { error=e; }
                      catch(COMException e) { error=e; }
                });
                thread.SetApartmentState(apartment==0?ApartmentState.STA:ApartmentState.MTA);
                thread.Start();
                if(!thread.Join(TimeSpan.FromSeconds(30))) throw new TimeoutException("Owned admission probe did not retire");
                if(error!=null) throw new InvalidOperationException("Independent admission probe failed",error);
            }
        }
        OriginalForegroundAfter=GetForegroundWindow().ToInt64();
        return results.ToArray();
    }
}
'@
$initialForegroundProcess=[WgcDesktopAdmissionProbe]::CurrentForegroundProcess()
$initialForegroundTitle=[WgcDesktopAdmissionProbe]::CurrentForegroundTitle()
$initialForegroundIdentity=if($initialForegroundProcess){Get-CimInstance Win32_Process -Filter ('ProcessId='+$initialForegroundProcess)|Select-Object Name,ExecutablePath,CommandLine,SessionId}else{$null}
$result=[WgcDesktopAdmissionProbe]::Run()
$ownedResult=@()
$foregroundBefore=$null;$foregroundAfter=$null;$grant=$null
if($env:GITHUB_ACTIONS -ceq 'true' -and $env:RUNNER_ENVIRONMENT -ceq 'github-hosted'){
    # Reuse the canonical entry lease definition, changing only its script-relative
    # module locator because this diagnostic has no native application/build role.
    $parseErrors=$null;$tokens=$null
    $runnerAst=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'Tools/Run-AllTests.ps1'),[ref]$tokens,[ref]$parseErrors)
    if($parseErrors.Count){throw 'Canonical runner parse failed'}
    $definition=@($runnerAst.FindAll({param($node)$node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq 'Invoke-RSInteractiveDesktopOperation'},$false))
    if($definition.Count -ne 1){throw 'Expected one canonical interactive lease'}
    $definitionText=$definition[0].Extent.Text.Replace("(Join-Path `$PSScriptRoot 'Modules/Testing/HostedTestDesktop.psm1')","(Join-Path `$repo 'Tools/Modules/Testing/HostedTestDesktop.psm1')")
    . ([scriptblock]::Create($definitionText))
    $priorDesktopRequest=$env:REDSALAMANDER_HOSTED_TEST_DESKTOP
    try{
        [WgcDesktopAdmissionProbe]::StartDefaultAnchor()
        $foregroundBefore=[WgcDesktopAdmissionProbe]::ActivateDefaultAnchor('before-bootstrap')
        $env:REDSALAMANDER_HOSTED_TEST_DESKTOP='owned'
        $ownedResult=@(Invoke-RSInteractiveDesktopOperation -RequiresInteractiveDesktop $true -Operation {[WgcDesktopAdmissionProbe]::Run();$script:wgcSelfGrant=[WgcDesktopAdmissionProbe]::GrantSelfOnActiveInputDesktop()})
        $grant=$script:wgcSelfGrant
        $foregroundAfter=[WgcDesktopAdmissionProbe]::ActivateDefaultAnchor('after-bootstrap-specific-self-grant')
    }finally{[WgcDesktopAdmissionProbe]::StopDefaultAnchor();[Environment]::SetEnvironmentVariable('REDSALAMANDER_HOSTED_TEST_DESKTOP',$priorDesktopRequest,'Process')}
}
$path=Join-Path $taskTemp 'wgc-admission.json'
if(Test-Path -LiteralPath $path){throw 'Preserve existing probe evidence'}
$report=[ordered]@{Schema='private.wgc-admission.v1';Scope='Metadata admission only: no image or injected input; active owned desktop uses canonical hosted entry lease; no qualification receipt';Head=(& git -C $repo rev-parse HEAD);Architecture=$env:PROCESSOR_ARCHITECTURE;InitialForegroundProcess=$initialForegroundProcess;InitialForegroundTitle=$initialForegroundTitle;InitialForegroundIdentity=$initialForegroundIdentity;OriginalForegroundBefore=[WgcDesktopAdmissionProbe]::OriginalForegroundBefore;OriginalForegroundAfter=[WgcDesktopAdmissionProbe]::OriginalForegroundAfter;Results=$result;OwnedActiveResults=$ownedResult;DefaultForegroundBefore=$foregroundBefore;SpecificSelfGrant=$grant;DefaultForegroundAfter=$foregroundAfter}
[IO.File]::WriteAllText($path,($report|ConvertTo-Json -Depth 6),[Text.UTF8Encoding]::new($false))
$report|ConvertTo-Json -Depth 6 -Compress
Write-Host ('Probe evidence: '+$path)
