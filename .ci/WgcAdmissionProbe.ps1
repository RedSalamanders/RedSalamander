[CmdletBinding()]
param([string]$TestRoot='D:\RedSalamander.Perf',[switch]$DismissHostedAccountSurface,
      [switch]$CloseOnly,[switch]$SnapshotOnly,[string]$RepoRoot)
$ErrorActionPreference='Stop'
$repo=if($RepoRoot){[IO.Path]::GetFullPath($RepoRoot)}else{Split-Path -Parent $PSScriptRoot}
if($CloseOnly -and -not $DismissHostedAccountSurface){throw 'CloseOnly requires the explicit hosted account-surface authorization.'}
if($SnapshotOnly -and ($CloseOnly -or $DismissHostedAccountSurface)){throw 'SnapshotOnly cannot request a close.'}
if(($CloseOnly -or $SnapshotOnly) -and ($env:GITHUB_ACTIONS -cne 'true' -or $env:RUNNER_ENVIRONMENT -cne 'github-hosted' -or
    $env:GITHUB_REPOSITORY -cne 'RedSalamanders/RedSalamander' -or $env:GITHUB_EVENT_NAME -cne 'workflow_dispatch' -or
    [Runtime.InteropServices.RuntimeInformation]::OSArchitecture.ToString() -cne 'Arm64' -or
    [Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture.ToString() -cne 'Arm64')){
    throw 'This diagnostic requires this repository manual hosted native ARM64 execution.'
}
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
    [DllImport("user32.dll", CharSet=CharSet.Unicode, ExactSpelling=true)] static extern int GetClassNameW(IntPtr window,StringBuilder title,int capacity);
    [DllImport("user32.dll", ExactSpelling=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool IsWindow(IntPtr window);
    [DllImport("user32.dll", ExactSpelling=true, SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool PostMessageW(IntPtr window,uint message,UIntPtr wParam,IntPtr lParam);
    [DllImport("kernel32.dll", ExactSpelling=true, SetLastError=true)] static extern Microsoft.Win32.SafeHandles.SafeProcessHandle OpenProcess(uint access,bool inherit,uint process);
    [DllImport("kernel32.dll", CharSet=CharSet.Unicode, ExactSpelling=true, SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool QueryFullProcessImageNameW(Microsoft.Win32.SafeHandles.SafeProcessHandle process,uint flags,StringBuilder image,ref uint capacity);
    [DllImport("kernel32.dll", CharSet=CharSet.Unicode, ExactSpelling=true)] static extern int GetApplicationUserModelId(Microsoft.Win32.SafeHandles.SafeProcessHandle process,ref uint capacity,StringBuilder id);
    public static uint CurrentForegroundProcess() { uint process;GetWindowThreadProcessId(GetForegroundWindow(),out process);return process; }
    public static long CurrentForegroundWindow() { return GetForegroundWindow().ToInt64(); }
    public static uint CurrentForegroundThread() { uint process;return GetWindowThreadProcessId(GetForegroundWindow(),out process); }
    public static string CurrentForegroundClass() {var cls=new StringBuilder(256);GetClassNameW(GetForegroundWindow(),cls,cls.Capacity);return cls.ToString();}
    public static string CurrentForegroundTitle() {var title=new StringBuilder(1024);GetWindowTextW(GetForegroundWindow(),title,title.Capacity);return title.ToString();}
    public static string CurrentForegroundApplicationId() {
        using(var process=OpenProcess(0x1000,false,CurrentForegroundProcess())) {
            if(process==null || process.IsInvalid)return "unavailable:open-process";
            uint capacity=2048;var id=new StringBuilder((int)capacity);int status=GetApplicationUserModelId(process,ref capacity,id);
            return status==0?id.ToString():"unavailable:"+status.ToString();
        }
    }
    public static bool CloseVerifiedHostedAccountSurface(long expectedWindow,uint expectedThread,uint expectedProcess) {
        // This method remains unreachable without the explicit manual workflow
        // switch. It posts one ordinary close request, never kills a process,
        // sends authentication input, or changes a global foreground policy.
        if(Environment.GetEnvironmentVariable("GITHUB_ACTIONS")!="true" ||
           Environment.GetEnvironmentVariable("RUNNER_ENVIRONMENT")!="github-hosted" ||
           Environment.GetEnvironmentVariable("GITHUB_REPOSITORY")!="RedSalamanders/RedSalamander" ||
           Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE")!="ARM64")
            throw new InvalidOperationException("Account-surface dismissal is confined to this project's ephemeral hosted ARM64 runner.");
        IntPtr window=GetForegroundWindow();uint processId;uint threadId=GetWindowThreadProcessId(window,out processId);
        if(window==IntPtr.Zero || window.ToInt64()!=expectedWindow || threadId==0 || threadId!=expectedThread || processId!=expectedProcess)
            throw new InvalidOperationException("Account-surface identity changed.");
        using(var process=OpenProcess(0x1000,false,processId)) {
            if(process==null || process.IsInvalid)throw new Win32Exception(Marshal.GetLastWin32Error(),"Account-surface process identity unavailable");
            uint capacity=4096;var image=new StringBuilder((int)capacity);
            if(!QueryFullProcessImageNameW(process,0,image,ref capacity))throw new Win32Exception(Marshal.GetLastWin32Error());
            string expected=System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System),"WWAHost.exe");
            var title=new StringBuilder(1024);var cls=new StringBuilder(256);
            GetWindowTextW(window,title,title.Capacity);GetClassNameW(window,cls,cls.Capacity);
            uint appCapacity=2048;var application=new StringBuilder((int)appCapacity);
            int appStatus=GetApplicationUserModelId(process,ref appCapacity,application);
            if(!String.Equals(image.ToString(),expected,StringComparison.OrdinalIgnoreCase) ||
               title.ToString()!="Microsoft account" || cls.ToString()!="Windows.UI.Core.CoreWindow" ||
               appStatus!=0 || application.ToString()!="Microsoft.Windows.CloudExperienceHost_cw5n1h2txyewy!App")
                throw new InvalidOperationException("The foreground surface is not the verified stock account window.");
            uint checkedProcess;uint checkedThread=GetWindowThreadProcessId(window,out checkedProcess);
            if(GetForegroundWindow()!=window || checkedProcess!=processId || checkedThread!=threadId || CurrentForegroundTitle()!="Microsoft account")
                throw new InvalidOperationException("The approved account surface drifted before its close request.");
            if(!PostMessageW(window,0x0010,UIntPtr.Zero,IntPtr.Zero))throw new Win32Exception(Marshal.GetLastWin32Error(),"Account window refused its ordinary close request");
            var deadline=System.Diagnostics.Stopwatch.StartNew();
            while(IsWindow(window) && (IsWindowVisible(window) || GetForegroundWindow()==window) && deadline.ElapsedMilliseconds<5000)Thread.Sleep(25);
            if(IsWindow(window) && (IsWindowVisible(window) || GetForegroundWindow()==window))throw new InvalidOperationException("The stock account window did not close; no process termination or input fallback is permitted.");
            return true;
        }
    }
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
function Get-PrivateForegroundSnapshot {
    $window=[WgcDesktopAdmissionProbe]::CurrentForegroundWindow()
    $process=[WgcDesktopAdmissionProbe]::CurrentForegroundProcess()
    $title=[WgcDesktopAdmissionProbe]::CurrentForegroundTitle()
    $class=[WgcDesktopAdmissionProbe]::CurrentForegroundClass()
    $application=[WgcDesktopAdmissionProbe]::CurrentForegroundApplicationId()
    $image=$null;$session=$null;$queryError=$null
    if($process){
        try{$info=Get-CimInstance Win32_Process -Filter ('ProcessId='+$process) -ErrorAction Stop
            if($null -ne $info){$image=$info.ExecutablePath;$session=$info.SessionId}else{$queryError='process-not-found'}}
        catch{$queryError=$_.Exception.Message}
    }
    # A changed foreground during this read is surfaced, never silently paired.
    $afterWindow=[WgcDesktopAdmissionProbe]::CurrentForegroundWindow();$afterProcess=[WgcDesktopAdmissionProbe]::CurrentForegroundProcess()
    return [ordered]@{ObservedUtc=[DateTime]::UtcNow.ToString('o');Window=$window;Process=$process;Title=$title;Class=$class;ApplicationId=$application;
        Image=$image;Session=$session;QueryError=$queryError;Stable=($window -eq $afterWindow -and $process -eq $afterProcess)}
}
if($SnapshotOnly){
    $snapshotPath=Join-Path $taskTemp 'foreground-snapshot.json'
    if(Test-Path -LiteralPath $snapshotPath){throw 'Preserve existing foreground snapshot.'}
    [IO.File]::WriteAllText($snapshotPath,((Get-PrivateForegroundSnapshot)|ConvertTo-Json -Depth 5),[Text.UTF8Encoding]::new($false))
    Write-Host "Foreground snapshot: $snapshotPath"
    return
}
$initialForegroundWindow=[WgcDesktopAdmissionProbe]::CurrentForegroundWindow()
$initialForegroundThread=[WgcDesktopAdmissionProbe]::CurrentForegroundThread()
$initialForegroundProcess=[WgcDesktopAdmissionProbe]::CurrentForegroundProcess()
$initialForegroundTitle=[WgcDesktopAdmissionProbe]::CurrentForegroundTitle()
$initialForegroundApplicationId=[WgcDesktopAdmissionProbe]::CurrentForegroundApplicationId()
$initialForegroundIdentity=if($initialForegroundProcess){Get-CimInstance Win32_Process -Filter ('ProcessId='+$initialForegroundProcess)|Select-Object Name,ExecutablePath,CommandLine,SessionId}else{$null}
$accountSurfaceDismissed=$false
if($DismissHostedAccountSurface){
    Import-Module (Join-Path $repo 'Tools/Modules/Testing/HostedTestDesktop.psm1')
    $context=(Get-RSHostedDesktopDiagnostics).DesktopContext
    if($context.WindowStation -cne 'WinSta0' -or $context.InputDesktop -cne 'Default' -or $context.ThreadDesktop -cne 'Default'){
        throw 'Refuse account-window intervention on an unknown or secure desktop.'
    }
    $intentPath=Join-Path $taskTemp 'account-dismissal-intent.json'
    $resultPath=Join-Path $taskTemp 'account-dismissal-result.json'
    if((Test-Path -LiteralPath $intentPath) -or (Test-Path -LiteralPath $resultPath)){throw 'Preserve account-window intervention evidence.'}
    $intent=[ordered]@{Scope='Explicit manual hosted ARM64 account-window close only';RequestedUtc=[DateTime]::UtcNow.ToString('o');Head=(& git -C $repo rev-parse HEAD);Window=$initialForegroundWindow;Thread=$initialForegroundThread;Process=$initialForegroundProcess;Title=$initialForegroundTitle;ApplicationId=$initialForegroundApplicationId;Identity=$initialForegroundIdentity;DesktopContext=$context}
    [IO.File]::WriteAllText($intentPath,($intent|ConvertTo-Json -Depth 6),[Text.UTF8Encoding]::new($false))
    $closeError=$null
    $noTarget=$false
    try{if($CloseOnly -and $initialForegroundWindow -eq 0 -and $initialForegroundThread -eq 0 -and $initialForegroundProcess -eq 0 -and [WgcDesktopAdmissionProbe]::CurrentForegroundWindow() -eq 0){$noTarget=$true}else{$accountSurfaceDismissed=[WgcDesktopAdmissionProbe]::CloseVerifiedHostedAccountSurface($initialForegroundWindow,$initialForegroundThread,$initialForegroundProcess)}}catch{$closeError=$_.Exception.Message;throw}finally{
        $postClose=$null;$snapshotError=$null
        try{$postClose=Get-PrivateForegroundSnapshot}catch{$snapshotError=$_.Exception.Message}
        $outcome=[ordered]@{CompletedUtc=[DateTime]::UtcNow.ToString('o');IntentSha256=(Get-FileHash -LiteralPath $intentPath -Algorithm SHA256).Hash;Dismissed=$accountSurfaceDismissed;NoTarget=$noTarget;NoTargetScope='No foreground window observed, not a desktop-wide stock-window absence proof';Error=$closeError;PostCloseForeground=$postClose;SnapshotError=$snapshotError}
        [IO.File]::WriteAllText($resultPath,($outcome|ConvertTo-Json),[Text.UTF8Encoding]::new($false))
        if($CloseOnly){Write-Host "Close intent: $intentPath";Write-Host "Close result: $resultPath"}
    }
}
if($CloseOnly){return}
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
$finalForeground=Get-PrivateForegroundSnapshot
$report=[ordered]@{Schema='private.wgc-admission.v1';Scope='Private desktop control: no application build, image, authentication input, or qualification receipt';Head=(& git -C $repo rev-parse HEAD);Architecture=$env:PROCESSOR_ARCHITECTURE;InitialForegroundProcess=$initialForegroundProcess;InitialForegroundTitle=$initialForegroundTitle;InitialForegroundApplicationId=$initialForegroundApplicationId;InitialForegroundIdentity=$initialForegroundIdentity;AccountSurfaceDismissed=$accountSurfaceDismissed;OriginalForegroundBefore=[WgcDesktopAdmissionProbe]::OriginalForegroundBefore;OriginalForegroundAfter=[WgcDesktopAdmissionProbe]::OriginalForegroundAfter;Results=$result;OwnedActiveResults=$ownedResult;DefaultForegroundBefore=$foregroundBefore;SpecificSelfGrant=$grant;DefaultForegroundAfter=$foregroundAfter;FinalForeground=$finalForeground}
[IO.File]::WriteAllText($path,($report|ConvertTo-Json -Depth 6),[Text.UTF8Encoding]::new($false))
$report|ConvertTo-Json -Depth 6 -Compress
Write-Host ('Probe evidence: '+$path)
