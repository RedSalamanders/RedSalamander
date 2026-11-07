using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Microsoft.Win32.SafeHandles;
namespace PrivateWslStartup {
 public sealed class LocalArguments : SafeHandleZeroOrMinusOneIsInvalid {
  public LocalArguments():base(true){}
  [DllImport("kernel32.dll",ExactSpelling=true)] static extern IntPtr LocalFree(IntPtr handle);
  protected override bool ReleaseHandle(){return LocalFree(handle)==IntPtr.Zero;}
 }
 public sealed class CapturedWindow : IDisposable {
  [DllImport("user32.dll",ExactSpelling=true)] static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll",ExactSpelling=true)] static extern uint GetWindowThreadProcessId(IntPtr window,out uint process);
  [DllImport("user32.dll",ExactSpelling=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool IsWindow(IntPtr window);
  [DllImport("user32.dll",ExactSpelling=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool IsWindowVisible(IntPtr window);
  [DllImport("user32.dll",CharSet=CharSet.Unicode,ExactSpelling=true)] static extern int GetWindowTextW(IntPtr window,StringBuilder text,int count);
  [DllImport("user32.dll",CharSet=CharSet.Unicode,ExactSpelling=true)] static extern int GetClassNameW(IntPtr window,StringBuilder text,int count);
  [DllImport("user32.dll",ExactSpelling=true,SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool PostMessageW(IntPtr window,uint message,UIntPtr w,IntPtr l);
  [DllImport("kernel32.dll",ExactSpelling=true)] static extern uint GetCurrentProcessId();
  [DllImport("kernel32.dll",ExactSpelling=true,SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool ProcessIdToSessionId(uint process,out uint session);
  [DllImport("kernel32.dll",ExactSpelling=true,SetLastError=true)] static extern SafeProcessHandle OpenProcess(uint access,[MarshalAs(UnmanagedType.Bool)] bool inherit,uint process);
  [DllImport("kernel32.dll",ExactSpelling=true)] static extern uint GetProcessId(SafeProcessHandle process);
  [DllImport("kernel32.dll",CharSet=CharSet.Unicode,ExactSpelling=true,SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool QueryFullProcessImageNameW(SafeProcessHandle process,uint flags,StringBuilder path,ref uint size);
  [DllImport("kernel32.dll",ExactSpelling=true,SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool GetProcessTimes(SafeProcessHandle process,out long created,out long exit,out long kernel,out long user);
  [DllImport("shell32.dll",CharSet=CharSet.Unicode,ExactSpelling=true,SetLastError=true)] static extern LocalArguments CommandLineToArgvW(string command,out int count);
  readonly IntPtr window; readonly SafeProcessHandle process; readonly long creationFileTime; bool closeRequested;
  public long Window {get{return window.ToInt64();}} public uint ProcessId {get;private set;} public uint ThreadId {get;private set;} public uint SessionId {get;private set;}
  public DateTime CreationUtc {get{return DateTime.FromFileTimeUtc(creationFileTime);}}
  CapturedWindow(IntPtr h,uint pid,uint tid,uint session,SafeProcessHandle p,long created){window=h;ProcessId=pid;ThreadId=tid;SessionId=session;process=p;creationFileTime=created;}
  static string Title(IntPtr h){var s=new StringBuilder(1024);GetWindowTextW(h,s,s.Capacity);return s.ToString();}
  static string Class(IntPtr h){var s=new StringBuilder(256);GetClassNameW(h,s,s.Capacity);return s.ToString();}
  static string Image(SafeProcessHandle p){uint n=4096;var s=new StringBuilder((int)n);if(!QueryFullProcessImageNameW(p,0,s,ref n))throw new Win32Exception(Marshal.GetLastWin32Error(),"Cannot query retained WSL process image.");return s.ToString();}
  static string ExpectedImage(){return System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System),"wsl.exe");}
  static bool ExactWindow(IntPtr h){return Title(h)==@"C:\Windows\system32\wsl.exe" && Class(h)=="ConsoleWindowClass";}
  public static CapturedWindow Capture(){
   IntPtr h=GetForegroundWindow();uint pid;uint tid=GetWindowThreadProcessId(h,out pid),session,own;
   if(h==IntPtr.Zero || pid==0 || tid==0 || !IsWindow(h) || !IsWindowVisible(h) || !ExactWindow(h))throw new InvalidOperationException("Foreground is not the exact approved WSL update console identity.");
   if(!ProcessIdToSessionId(pid,out session) || !ProcessIdToSessionId(GetCurrentProcessId(),out own) || session!=own)throw new InvalidOperationException("WSL and runner sessions differ or are unavailable.");
   var p=OpenProcess(0x1000,false,pid);if(p==null || p.IsInvalid){int error=Marshal.GetLastWin32Error();if(p!=null)p.Dispose();throw new Win32Exception(error,"Cannot retain WSL process.");}
   bool transfer=false;try{long created,ex,k,u;if(GetProcessId(p)!=pid || !GetProcessTimes(p,out created,out ex,out k,out u))throw new Win32Exception(Marshal.GetLastWin32Error(),"Cannot bind WSL process creation.");
    if(!String.Equals(Image(p),ExpectedImage(),StringComparison.OrdinalIgnoreCase) || GetForegroundWindow()!=h || GetWindowThreadProcessId(h,out pid)!=tid || pid!=GetProcessId(p))throw new InvalidOperationException("WSL process or foreground changed while captured.");
    var result=new CapturedWindow(h,pid,tid,session,p,created);transfer=true;return result;
   }finally{if(!transfer)p.Dispose();}
  }
  public static bool ExactArguments(string command,out int count){
   count=0;if(String.IsNullOrWhiteSpace(command))return false;
   using(var args=CommandLineToArgvW(command,out count)){if(args==null || args.IsInvalid || count!=4)return false;
    string exe=Marshal.PtrToStringUni(Marshal.ReadIntPtr(args.DangerousGetHandle(),0));
    bool executable=String.Equals(exe,ExpectedImage(),StringComparison.OrdinalIgnoreCase) || String.Equals(exe,"wsl.exe",StringComparison.OrdinalIgnoreCase);
    return executable && Marshal.PtrToStringUni(Marshal.ReadIntPtr(args.DangerousGetHandle(),IntPtr.Size))=="--update" &&
     Marshal.PtrToStringUni(Marshal.ReadIntPtr(args.DangerousGetHandle(),IntPtr.Size*2))=="--confirm" &&
     Marshal.PtrToStringUni(Marshal.ReadIntPtr(args.DangerousGetHandle(),IntPtr.Size*3))=="--prompt-before-exit";
   }
  }
  public void CloseOnce(){
   if(closeRequested)throw new InvalidOperationException("Only one WSL close request is permitted.");
   uint pid,session,own;uint tid=GetWindowThreadProcessId(window,out pid);long created,ex,k,u;
   if(GetForegroundWindow()!=window || !IsWindow(window) || !IsWindowVisible(window) || pid!=ProcessId || tid!=ThreadId || GetProcessId(process)!=ProcessId ||
      !ProcessIdToSessionId(pid,out session) || !ProcessIdToSessionId(GetCurrentProcessId(),out own) || session!=SessionId || session!=own ||
      !GetProcessTimes(process,out created,out ex,out k,out u) || created!=creationFileTime || !ExactWindow(window) || !String.Equals(Image(process),ExpectedImage(),StringComparison.OrdinalIgnoreCase))
    throw new InvalidOperationException("WSL foreground/process identity changed before close.");
   closeRequested=true;
   if(!PostMessageW(window,0x0010,UIntPtr.Zero,IntPtr.Zero))throw new Win32Exception(Marshal.GetLastWin32Error(),"WSL update console refused WM_CLOSE.");
   var timer=System.Diagnostics.Stopwatch.StartNew();while(timer.ElapsedMilliseconds<5000){
    if(!IsWindow(window))return;tid=GetWindowThreadProcessId(window,out pid);
    if(pid!=ProcessId || tid!=ThreadId)throw new InvalidOperationException("WSL window identity changed during retirement.");
    if(GetForegroundWindow()!=window && !IsWindowVisible(window))return;Thread.Sleep(25);
   }
   throw new TimeoutException("WSL update console did not retire within five seconds.");
  }
  public void Dispose(){process.Dispose();}
 }
}
