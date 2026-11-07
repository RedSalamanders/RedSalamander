using System;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
namespace PrivateHosted {
 public sealed class SafeNativeBuffer : SafeHandleZeroOrMinusOneIsInvalid {
  public SafeNativeBuffer(int bytes):base(true){SetHandle(Marshal.AllocHGlobal(bytes));}
  protected override bool ReleaseHandle(){Marshal.FreeHGlobal(handle);return true;}
 }
 public sealed class TokenValue { public bool Available; public uint Value; public int Error; }
 public sealed class TokenFacts {
  public bool Opened;public int OpenError;
  public TokenValue IntegrityRid=new TokenValue(),UIAccess=new TokenValue(),Elevation=new TokenValue(),ElevationType=new TokenValue(),AppContainer=new TokenValue(),SessionId=new TokenValue();
 }
 public sealed class SecurityFacts {
  public uint CurrentProcessId,ForegroundProcessId,ForegroundThreadId,ForegroundAfterProcessId,ForegroundAfterThreadId;
  public long ForegroundWindow,ForegroundAfterWindow;public bool ForegroundStable;public int ForegroundProcessOpenError;
  public TokenFacts Current=new TokenFacts(),Foreground=new TokenFacts();
  public bool ThreadTokenAvailable;public int ThreadTokenError;
  public bool JobMembershipAvailable,InJob;public int JobMembershipError;public bool JobUiRestrictionsAvailable;public uint JobUiRestrictions;public int JobUiRestrictionsError;
 }
 public static class InputSecurity {
  [DllImport("kernel32.dll",ExactSpelling=true)] static extern IntPtr GetCurrentProcess();
  [DllImport("kernel32.dll",ExactSpelling=true)] static extern IntPtr GetCurrentThread();
  [DllImport("kernel32.dll",ExactSpelling=true)] static extern uint GetCurrentProcessId();
  [DllImport("kernel32.dll",ExactSpelling=true,SetLastError=true)] static extern SafeProcessHandle OpenProcess(uint access,[MarshalAs(UnmanagedType.Bool)] bool inherit,uint pid);
  [DllImport("user32.dll",ExactSpelling=true)] static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll",ExactSpelling=true)] static extern uint GetWindowThreadProcessId(IntPtr h,out uint pid);
  [DllImport("advapi32.dll",ExactSpelling=true,SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool OpenProcessToken(IntPtr p,uint access,out SafeAccessTokenHandle token);
  [DllImport("advapi32.dll",ExactSpelling=true,SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool OpenThreadToken(IntPtr t,uint access,[MarshalAs(UnmanagedType.Bool)] bool openAsSelf,out SafeAccessTokenHandle token);
  [DllImport("advapi32.dll",ExactSpelling=true,SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool GetTokenInformation(SafeAccessTokenHandle token,int kind,IntPtr buffer,uint size,out uint needed);
  [DllImport("advapi32.dll",ExactSpelling=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool IsValidSid(IntPtr sid);
  [DllImport("advapi32.dll",ExactSpelling=true)] static extern IntPtr GetSidSubAuthorityCount(IntPtr sid);
  [DllImport("advapi32.dll",ExactSpelling=true)] static extern IntPtr GetSidSubAuthority(IntPtr sid,uint index);
  [DllImport("kernel32.dll",ExactSpelling=true,SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool IsProcessInJob(IntPtr process,IntPtr job,[MarshalAs(UnmanagedType.Bool)] out bool result);
  [DllImport("kernel32.dll",ExactSpelling=true,SetLastError=true)] [return:MarshalAs(UnmanagedType.Bool)] static extern bool QueryInformationJobObject(IntPtr job,int info,out uint data,uint length,out uint returned);
  static void ReadScalar(SafeAccessTokenHandle token,int kind,TokenValue value) {
   using(var b=new SafeNativeBuffer(4)){uint needed;bool ok=GetTokenInformation(token,kind,b.DangerousGetHandle(),4,out needed);int error=ok?0:Marshal.GetLastWin32Error();if(ok && needed==4){value.Available=true;value.Value=unchecked((uint)Marshal.ReadInt32(b.DangerousGetHandle()));}else value.Error=ok?13:error;}
  }
  static void ReadIntegrity(SafeAccessTokenHandle token,TokenValue value) {
   uint needed;GetTokenInformation(token,25,IntPtr.Zero,0,out needed);int error=Marshal.GetLastWin32Error();
   if(needed==0 || needed>65536){value.Error=error;return;}
   using(var b=new SafeNativeBuffer((int)needed)){
    uint returned;if(!GetTokenInformation(token,25,b.DangerousGetHandle(),needed,out returned)){value.Error=Marshal.GetLastWin32Error();return;}
    IntPtr sid=Marshal.ReadIntPtr(b.DangerousGetHandle());if(sid==IntPtr.Zero || !IsValidSid(sid)){value.Error=13;return;}
    byte count=Marshal.ReadByte(GetSidSubAuthorityCount(sid));if(count==0){value.Error=13;return;}
    value.Value=unchecked((uint)Marshal.ReadInt32(GetSidSubAuthority(sid,(uint)(count-1))));value.Available=true;
   }
  }
  static TokenFacts ReadToken(IntPtr process) {
   var facts=new TokenFacts();SafeAccessTokenHandle token;bool ok=OpenProcessToken(process,8,out token);int error=ok?0:Marshal.GetLastWin32Error();
   using(token){facts.Opened=ok && token!=null && !token.IsInvalid;if(!facts.Opened){facts.OpenError=error;return facts;}
    ReadIntegrity(token,facts.IntegrityRid);ReadScalar(token,26,facts.UIAccess);ReadScalar(token,20,facts.Elevation);ReadScalar(token,18,facts.ElevationType);ReadScalar(token,29,facts.AppContainer);ReadScalar(token,12,facts.SessionId);}
   return facts;
  }
  public static SecurityFacts Read() {
   var r=new SecurityFacts();r.CurrentProcessId=GetCurrentProcessId();r.Current=ReadToken(GetCurrentProcess());
   SafeAccessTokenHandle thread;bool threadOk=OpenThreadToken(GetCurrentThread(),8,true,out thread);int threadError=threadOk?0:Marshal.GetLastWin32Error();using(thread){r.ThreadTokenAvailable=threadOk && thread!=null && !thread.IsInvalid;r.ThreadTokenError=threadError;}
   bool inJob;r.JobMembershipAvailable=IsProcessInJob(GetCurrentProcess(),IntPtr.Zero,out inJob);r.InJob=inJob;if(!r.JobMembershipAvailable)r.JobMembershipError=Marshal.GetLastWin32Error();
   if(r.JobMembershipAvailable && inJob){uint flags,returned;r.JobUiRestrictionsAvailable=QueryInformationJobObject(IntPtr.Zero,4,out flags,4,out returned);r.JobUiRestrictions=flags;if(!r.JobUiRestrictionsAvailable)r.JobUiRestrictionsError=Marshal.GetLastWin32Error();else if(returned!=4){r.JobUiRestrictionsAvailable=false;r.JobUiRestrictionsError=13;}}
   IntPtr foreground=GetForegroundWindow();uint pid=0;uint tid=foreground==IntPtr.Zero?0:GetWindowThreadProcessId(foreground,out pid);r.ForegroundWindow=foreground.ToInt64();r.ForegroundProcessId=pid;r.ForegroundThreadId=tid;
   if(pid!=0){using(var process=OpenProcess(0x1000,false,pid)){if(process==null || process.IsInvalid)r.ForegroundProcessOpenError=Marshal.GetLastWin32Error();else r.Foreground=ReadToken(process.DangerousGetHandle());}}
   IntPtr after=GetForegroundWindow();uint afterPid=0;uint afterTid=after==IntPtr.Zero?0:GetWindowThreadProcessId(after,out afterPid);r.ForegroundAfterWindow=after.ToInt64();r.ForegroundAfterProcessId=afterPid;r.ForegroundAfterThreadId=afterTid;r.ForegroundStable=foreground==after && pid==afterPid && tid==afterTid;
   return r;
  }
 }
}
