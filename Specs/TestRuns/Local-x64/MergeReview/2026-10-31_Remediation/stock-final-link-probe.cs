using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
public static class RsGitLinkProbe {
  [StructLayout(LayoutKind.Sequential)] public struct NativeFileTime { public uint Low; public uint High; }
  [StructLayout(LayoutKind.Sequential)] public struct BY_HANDLE_FILE_INFORMATION { public uint Attributes; public NativeFileTime Creation; public NativeFileTime Access; public NativeFileTime Write; public uint VolumeSerial; public uint SizeHigh; public uint SizeLow; public uint NumberOfLinks; public uint FileIndexHigh; public uint FileIndexLow; }
  [DllImport("kernel32.dll", SetLastError=true)] public static extern bool GetFileInformationByHandle(SafeFileHandle h, out BY_HANDLE_FILE_INFORMATION info);
}
