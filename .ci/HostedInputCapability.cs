using System;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
namespace PrivateHosted {
 public sealed class InputDesktop : SafeHandleZeroOrMinusOneIsInvalid {
  public InputDesktop():base(true) {}
  [DllImport("user32.dll",ExactSpelling=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool CloseDesktop(IntPtr h);
  protected override bool ReleaseHandle(){return CloseDesktop(handle);}
 }
 public sealed class WtsMemory : SafeHandleZeroOrMinusOneIsInvalid {
  public WtsMemory():base(true) {}
  [DllImport("wtsapi32.dll",ExactSpelling=true)] static extern void WTSFreeMemory(IntPtr p);
  protected override bool ReleaseHandle(){WTSFreeMemory(handle);return true;}
 }
 [StructLayout(LayoutKind.Sequential)] public struct Flags {public int Inherit,Reserved;public uint Value;}
 public sealed class Result {
  public bool SessionAvailable;public uint SessionId;public int SessionError;public uint ActiveConsoleSessionId;
  public bool ConnectStateAvailable;public int ConnectState;public int ConnectStateError;public uint ConnectStateBytes;
  public bool StationFlagsAvailable;public uint StationFlags;public bool StationVisible;public int StationFlagsError;
  public bool ThreadIoAvailable;public bool ThreadReceivesInput;public int ThreadIoError;
  public bool InputDesktopOpened;public int InputOpenError;public bool InputIoAvailable;public bool InputReceivesInput;public int InputIoError;
 }
 public static class InputCapability {
  [DllImport("kernel32.dll",ExactSpelling=true)] static extern uint GetCurrentProcessId();
  [DllImport("kernel32.dll",ExactSpelling=true)] static extern uint GetCurrentThreadId();
  [DllImport("kernel32.dll",ExactSpelling=true,SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool ProcessIdToSessionId(uint p,out uint s);
  [DllImport("kernel32.dll",ExactSpelling=true)] static extern uint WTSGetActiveConsoleSessionId();
  [DllImport("wtsapi32.dll",ExactSpelling=true,SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool WTSQuerySessionInformationW(IntPtr server,uint session,int kind,out WtsMemory mem,out uint size);
  [DllImport("user32.dll",ExactSpelling=true,SetLastError=true)] static extern IntPtr GetProcessWindowStation();
  [DllImport("user32.dll",ExactSpelling=true,SetLastError=true)] static extern IntPtr GetThreadDesktop(uint t);
  [DllImport("user32.dll",ExactSpelling=true,SetLastError=true)] static extern InputDesktop OpenInputDesktop(uint f,[MarshalAs(UnmanagedType.Bool)] bool inherit,uint access);
  [DllImport("user32.dll",EntryPoint="GetUserObjectInformationW",ExactSpelling=true,SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool GetFlags(IntPtr h,int index,ref Flags f,uint len,out uint needed);
  [DllImport("user32.dll",EntryPoint="GetUserObjectInformationW",ExactSpelling=true,SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool GetIo(IntPtr h,int index,out int v,uint len,out uint needed);
  public static Result Read() {
   var r=new Result();uint session;r.SessionAvailable=ProcessIdToSessionId(GetCurrentProcessId(),out session);r.SessionId=session;if(!r.SessionAvailable)r.SessionError=Marshal.GetLastWin32Error();
   r.ActiveConsoleSessionId=WTSGetActiveConsoleSessionId();
   WtsMemory mem;uint size;bool ok=WTSQuerySessionInformationW(IntPtr.Zero,0xffffffff,8,out mem,out size);int err=ok?0:Marshal.GetLastWin32Error();
   using(mem){r.ConnectStateBytes=size;if(ok && mem!=null && !mem.IsInvalid && size>=4){r.ConnectState=Marshal.ReadInt32(mem.DangerousGetHandle());r.ConnectStateAvailable=true;}else r.ConnectStateError=err;}
   IntPtr station=GetProcessWindowStation();var flags=new Flags();uint needed;
   r.StationFlagsAvailable=station!=IntPtr.Zero && GetFlags(station,1,ref flags,(uint)Marshal.SizeOf<Flags>(),out needed);
   if(r.StationFlagsAvailable){r.StationFlags=flags.Value;r.StationVisible=(flags.Value & 1)!=0;}else r.StationFlagsError=Marshal.GetLastWin32Error();
   IntPtr desk=GetThreadDesktop(GetCurrentThreadId());int value=0;
   r.ThreadIoAvailable=desk!=IntPtr.Zero && GetIo(desk,6,out value,4,out needed);
   if(r.ThreadIoAvailable)r.ThreadReceivesInput=value!=0;else r.ThreadIoError=Marshal.GetLastWin32Error();
   using(var input=OpenInputDesktop(0,false,1)){
    r.InputDesktopOpened=input!=null && !input.IsInvalid;
    if(!r.InputDesktopOpened)r.InputOpenError=Marshal.GetLastWin32Error();
    else {r.InputIoAvailable=GetIo(input.DangerousGetHandle(),6,out value,4,out needed);if(r.InputIoAvailable)r.InputReceivesInput=value!=0;else r.InputIoError=Marshal.GetLastWin32Error();}
   }
   return r;
  }
 }
}
