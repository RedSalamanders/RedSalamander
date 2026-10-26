using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;
public sealed class SortProcessSample {
 public string Name {get;set;} = "";
 public int Id {get;set;}
 public long StartTicks {get;set;}
 public long CpuTicks {get;set;}
}
public static class SortProcessScanner {
 static readonly HashSet<string> Names=new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
  "MSBuild","cl","link","lib","rc","mt","clang","clang++","clang-cl","lld","lld-link","ninja","cmake","vcpkg","zig","gcc","g++","cc1","cc1plus","rustc","cargo","csc","vbc","dotnet","make","gmake","clang-scan-deps","pwsh","powershell","cmd","wsl","python","python3","node","DxUi.ControlTests","DxUi.EmbeddedTests","DxUi.FoundationTests","DxUi.PerformanceBench","RedPrism"
 };
 public static SortProcessSample[] Scan(int self) {
  var found=new List<SortProcessSample>();
  foreach(var p in Process.GetProcesses()) {
   using(p) {try {
    if(p.Id==self)continue;
    string name=p.ProcessName;
    if(!Names.Contains(name))continue;
    long ticks=0; try{ticks=p.StartTime.ToUniversalTime().Ticks;}catch(Win32Exception){}catch(InvalidOperationException){}
    found.Add(new SortProcessSample{Name=name,Id=p.Id,StartTicks=ticks,CpuTicks=p.TotalProcessorTime.Ticks});
   }catch(Win32Exception){}catch(InvalidOperationException){}}
  }
  return found.ToArray();
 }
 [DllImport("kernel32.dll",SetLastError=true)] static extern SafeProcessHandle OpenProcess(uint access,bool inherit,int pid);
 [DllImport("kernel32.dll",CharSet=CharSet.Unicode,SetLastError=true)] static extern bool QueryFullProcessImageName(SafeProcessHandle h,uint flags,StringBuilder b,ref int n);
 public static string Image(int pid) {
  using(var h=OpenProcess(0x1000,false,pid)) {
   if(h.IsInvalid)throw new Win32Exception();
   var b=new StringBuilder(32768);int n=b.Capacity;
   if(!QueryFullProcessImageName(h,0,b,ref n))throw new Win32Exception();
   return b.ToString();
  }
 }
}
