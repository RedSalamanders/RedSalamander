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
    [DllImport("user32.dll", CharSet=CharSet.Unicode, ExactSpelling=true, SetLastError=true)] static extern Window CreateWindowExW(uint ex,string cls,string title,uint style,int x,int y,int w,int h,IntPtr parent,IntPtr menu,IntPtr module,IntPtr param);
    [DllImport("user32.dll", ExactSpelling=true)] static extern int GetSystemMetrics(int index);
    [DllImport("user32.dll", ExactSpelling=true)] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll", ExactSpelling=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("dwmapi.dll", ExactSpelling=true)] static extern int DwmFlush();
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
        using(var desktop=CreateDesktopW(name,IntPtr.Zero,IntPtr.Zero,0,0x1ff,IntPtr.Zero)) {
            if(desktop==null || desktop.IsInvalid) throw new Win32Exception(Marshal.GetLastWin32Error(),"Create inactive probe desktop");
            foreach(bool alternate in new[]{false,true}) foreach(uint apartment in new uint[]{0,1}) {
                Exception error=null;
                var thread=new Thread(()=>{
                    try {
                        if(alternate && !SetThreadDesktop(desktop)) throw new Win32Exception(Marshal.GetLastWin32Error(),"Attach inactive probe desktop");
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
                                    results.Add(new Result{Desktop=alternate?"inactive-owned":"original",ExtendedStyle=style,FactoryHr=factoryHr,ItemHr=itemHr,FlushHr=flush,Visible=IsWindowVisible(window.DangerousGetHandle()),SameDesktopForegroundUnchanged=alternate?(bool?)null:GetForegroundWindow()==foreground,Apartment=apartment==0?"STA":"MTA",ItemWidth=lastWidth,ItemHeight=lastHeight,ItemSizeHr=lastSizeHr});
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
$result=[WgcDesktopAdmissionProbe]::Run()
$path=Join-Path $taskTemp 'wgc-admission.json'
if(Test-Path -LiteralPath $path){throw 'Preserve existing probe evidence'}
$report=[ordered]@{Schema='private.wgc-admission.v1';Scope='Metadata admission only: no image, input or desktop switch; no qualification receipt';Head=(& git -C $repo rev-parse HEAD);Architecture=$env:PROCESSOR_ARCHITECTURE;OriginalForegroundBefore=[WgcDesktopAdmissionProbe]::OriginalForegroundBefore;OriginalForegroundAfter=[WgcDesktopAdmissionProbe]::OriginalForegroundAfter;Results=$result}
[IO.File]::WriteAllText($path,($report|ConvertTo-Json -Depth 6),[Text.UTF8Encoding]::new($false))
$report|ConvertTo-Json -Depth 6 -Compress
Write-Host ('Probe evidence: '+$path)
