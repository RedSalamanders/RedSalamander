if (-not ('RedSalamander.Tooling.KillOnCloseJob' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Microsoft.Win32.SafeHandles;

namespace RedSalamander.Tooling
{
    public sealed class KillOnCloseJob : IDisposable
    {
        private const int JobObjectExtendedLimitInformation = 9;
        private const uint JobObjectLimitKillOnJobClose = 0x00002000;

        [StructLayout(LayoutKind.Sequential)]
        private struct BasicLimitInformation
        {
            public long PerProcessUserTimeLimit;
            public long PerJobUserTimeLimit;
            public uint LimitFlags;
            public UIntPtr MinimumWorkingSetSize;
            public UIntPtr MaximumWorkingSetSize;
            public uint ActiveProcessLimit;
            public UIntPtr Affinity;
            public uint PriorityClass;
            public uint SchedulingClass;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct IoCounters
        {
            public ulong ReadOperationCount;
            public ulong WriteOperationCount;
            public ulong OtherOperationCount;
            public ulong ReadTransferCount;
            public ulong WriteTransferCount;
            public ulong OtherTransferCount;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct ExtendedLimitInformation
        {
            public BasicLimitInformation BasicLimitInformation;
            public IoCounters IoInfo;
            public UIntPtr ProcessMemoryLimit;
            public UIntPtr JobMemoryLimit;
            public UIntPtr PeakProcessMemoryUsed;
            public UIntPtr PeakJobMemoryUsed;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr CreateJobObject(IntPtr jobAttributes, string name);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetInformationJobObject(
            SafeFileHandle job,
            int informationClass,
            IntPtr information,
            uint informationLength);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool IsProcessInJob(IntPtr process, IntPtr job, out bool result);

        private readonly SafeFileHandle job;

        internal SafeFileHandle Handle
        {
            get { return job; }
        }

        public KillOnCloseJob()
        {
            IntPtr rawJob = CreateJobObject(IntPtr.Zero, null);
            if (rawJob == IntPtr.Zero || rawJob == new IntPtr(-1))
            {
                throw new Win32Exception(Marshal.GetLastWin32Error(), "CreateJobObject failed.");
            }

            job = new SafeFileHandle(rawJob, true);
            try
            {
                ExtendedLimitInformation limits = new ExtendedLimitInformation();
                limits.BasicLimitInformation.LimitFlags = JobObjectLimitKillOnJobClose;
                int size = Marshal.SizeOf(typeof(ExtendedLimitInformation));
                IntPtr buffer = Marshal.AllocHGlobal(size);
                try
                {
                    Marshal.StructureToPtr(limits, buffer, false);
                    if (!SetInformationJobObject(job, JobObjectExtendedLimitInformation, buffer, (uint)size))
                    {
                        throw new Win32Exception(Marshal.GetLastWin32Error(), "SetInformationJobObject failed.");
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(buffer);
                }
            }
            catch
            {
                job.Dispose();
                throw;
            }
        }

        public static bool IsCurrentProcessInJob()
        {
            using (Process process = Process.GetCurrentProcess())
            {
                bool result;
                if (!IsProcessInJob(process.Handle, IntPtr.Zero, out result))
                {
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "IsProcessInJob failed.");
                }

                return result;
            }
        }

        internal bool ContainsProcess(SafeFileHandle process)
        {
            bool result;
            if (!IsProcessInJob(process.DangerousGetHandle(), job.DangerousGetHandle(), out result))
            {
                throw new Win32Exception(Marshal.GetLastWin32Error(), "IsProcessInJob failed for the containment job.");
            }
            return result;
        }

        public void Dispose()
        {
            job.Dispose();
            GC.SuppressFinalize(this);
        }
    }

    public sealed class ContainedProcess : IDisposable
    {
        private const uint CreateNoWindow = 0x08000000;
        private const uint CreateUnicodeEnvironment = 0x00000400;
        private const uint ExtendedStartupInfoPresent = 0x00080000;
        private const int StartfUseShowWindow = 0x00000001;
        private const int StartfUseStdHandles = 0x00000100;
        private const short SwHide = 0;
        private const uint HandleFlagInherit = 0x00000001;
        private const uint DuplicateSameAccess = 0x00000002;
        private const int StdInputHandle = -10;
        private const int StdOutputHandle = -11;
        private const int StdErrorHandle = -12;
        private const uint GenericRead = 0x80000000;
        private const uint GenericWrite = 0x40000000;
        private const uint FileShareRead = 0x00000001;
        private const uint FileShareWrite = 0x00000002;
        private const uint OpenExisting = 3;
        private const uint Infinite = 0xFFFFFFFF;
        private const uint WaitObject0 = 0;
        private const uint WaitTimeout = 258;
        private const uint WaitFailed = 0xFFFFFFFF;
        private static readonly IntPtr ProcThreadAttributeHandleList = new IntPtr(0x00020002);
        private static readonly IntPtr ProcThreadAttributeJobList = new IntPtr(0x0002000D);

        [StructLayout(LayoutKind.Sequential)]
        private struct SecurityAttributes
        {
            public int Length;
            public IntPtr SecurityDescriptor;
            [MarshalAs(UnmanagedType.Bool)]
            public bool InheritHandle;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct StartupInfo
        {
            public int Size;
            public string Reserved;
            public string Desktop;
            public string Title;
            public int X;
            public int Y;
            public int XSize;
            public int YSize;
            public int XCountChars;
            public int YCountChars;
            public int FillAttribute;
            public int Flags;
            public short ShowWindow;
            public short ReservedByteCount;
            public IntPtr ReservedBytes;
            public IntPtr StandardInput;
            public IntPtr StandardOutput;
            public IntPtr StandardError;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct StartupInfoEx
        {
            public StartupInfo StartupInfo;
            public IntPtr AttributeList;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct ProcessInformation
        {
            public IntPtr Process;
            public IntPtr Thread;
            public int ProcessId;
            public int ThreadId;
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool InitializeProcThreadAttributeList(
            IntPtr attributeList,
            int attributeCount,
            int flags,
            ref IntPtr size);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool UpdateProcThreadAttribute(
            IntPtr attributeList,
            uint flags,
            IntPtr attribute,
            IntPtr value,
            IntPtr size,
            IntPtr previousValue,
            IntPtr returnSize);

        [DllImport("kernel32.dll")]
        private static extern void DeleteProcThreadAttributeList(IntPtr attributeList);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CreateProcessW(
            string applicationName,
            StringBuilder commandLine,
            IntPtr processAttributes,
            IntPtr threadAttributes,
            [MarshalAs(UnmanagedType.Bool)] bool inheritHandles,
            uint creationFlags,
            IntPtr environment,
            string currentDirectory,
            ref StartupInfoEx startupInfo,
            out ProcessInformation processInformation);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CreatePipe(
            out SafeFileHandle readPipe,
            out SafeFileHandle writePipe,
            ref SecurityAttributes pipeAttributes,
            int size);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetHandleInformation(
            SafeFileHandle handle,
            uint mask,
            uint flags);

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetCurrentProcess();

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool DuplicateHandle(
            IntPtr sourceProcess,
            IntPtr sourceHandle,
            IntPtr targetProcess,
            out SafeFileHandle targetHandle,
            uint desiredAccess,
            [MarshalAs(UnmanagedType.Bool)] bool inheritHandle,
            uint options);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GetStdHandle(int standardHandle);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafeFileHandle CreateFileW(
            string fileName,
            uint desiredAccess,
            uint shareMode,
            ref SecurityAttributes securityAttributes,
            uint creationDisposition,
            uint flagsAndAttributes,
            IntPtr templateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern uint WaitForSingleObject(SafeFileHandle handle, uint milliseconds);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetExitCodeProcess(SafeFileHandle process, out uint exitCode);

        private KillOnCloseJob job;
        private SafeFileHandle process;
        private StreamWriter standardInput;
        private StreamReader standardOutput;
        private StreamReader standardError;
        private readonly int processId;
        private int disposed;

        private ContainedProcess(
            KillOnCloseJob job,
            SafeFileHandle process,
            int processId,
            StreamWriter standardInput,
            StreamReader standardOutput,
            StreamReader standardError)
        {
            this.job = job;
            this.process = process;
            this.processId = processId;
            this.standardInput = standardInput;
            this.standardOutput = standardOutput;
            this.standardError = standardError;
        }

        public int Id { get { return processId; } }
        public StreamWriter StandardInput { get { return standardInput; } }
        public StreamReader StandardOutput { get { return standardOutput; } }
        public StreamReader StandardError { get { return standardError; } }
        public bool IsInContainmentJob { get { return job.ContainsProcess(process); } }
        public object ArtifactOperationDelegation { get; set; }

        public bool HasExited
        {
            get
            {
                uint result = WaitForSingleObject(process, 0);
                if (result == WaitObject0)
                {
                    return true;
                }
                if (result == WaitTimeout)
                {
                    return false;
                }
                throw new Win32Exception(Marshal.GetLastWin32Error(), "WaitForSingleObject failed.");
            }
        }

        public int ExitCode
        {
            get
            {
                if (!HasExited)
                {
                    throw new InvalidOperationException("The contained process has not exited.");
                }

                uint exitCode;
                if (!GetExitCodeProcess(process, out exitCode))
                {
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "GetExitCodeProcess failed.");
                }
                return unchecked((int)exitCode);
            }
        }

        public static ContainedProcess Start(ProcessStartInfo startInfo, string commandLine)
        {
            if (startInfo == null)
            {
                throw new ArgumentNullException("startInfo");
            }
            if (startInfo.UseShellExecute)
            {
                throw new InvalidOperationException("Contained processes require UseShellExecute=false.");
            }
            if (String.IsNullOrWhiteSpace(startInfo.FileName))
            {
                throw new ArgumentException("A process executable is required.", "startInfo");
            }
            if (!String.IsNullOrWhiteSpace(startInfo.UserName) ||
                !String.IsNullOrWhiteSpace(startInfo.Verb) ||
                startInfo.ErrorDialog)
            {
                throw new InvalidOperationException("Contained processes do not support credentials, shell verbs, or error-dialog mode.");
            }
            if (String.IsNullOrWhiteSpace(commandLine))
            {
                throw new ArgumentException("A complete command line including argv[0] is required.", "commandLine");
            }

            KillOnCloseJob job = null;
            SafeFileHandle process = null;
            SafeFileHandle thread = null;
            SafeFileHandle childInput = null;
            SafeFileHandle childOutput = null;
            SafeFileHandle childError = null;
            SafeFileHandle parentInput = null;
            SafeFileHandle parentOutput = null;
            SafeFileHandle parentError = null;
            StreamWriter inputWriter = null;
            StreamReader outputReader = null;
            StreamReader errorReader = null;
            IntPtr attributeList = IntPtr.Zero;
            IntPtr attributeListSize = IntPtr.Zero;
            IntPtr jobValue = IntPtr.Zero;
            IntPtr handleValues = IntPtr.Zero;
            IntPtr environment = IntPtr.Zero;

            try
            {
                job = new KillOnCloseJob();
                bool useStandardHandles = startInfo.RedirectStandardInput ||
                    startInfo.RedirectStandardOutput || startInfo.RedirectStandardError;
                List<SafeFileHandle> childHandles = new List<SafeFileHandle>();

                if (useStandardHandles)
                {
                    if (startInfo.RedirectStandardInput)
                    {
                        CreateInputPipe(out childInput, out parentInput);
                    }
                    else
                    {
                        childInput = DuplicateStandardHandle(StdInputHandle, GenericRead);
                    }

                    if (startInfo.RedirectStandardOutput)
                    {
                        CreateOutputPipe(out parentOutput, out childOutput);
                    }
                    else
                    {
                        childOutput = DuplicateStandardHandle(StdOutputHandle, GenericWrite);
                    }

                    if (startInfo.RedirectStandardError)
                    {
                        CreateOutputPipe(out parentError, out childError);
                    }
                    else
                    {
                        childError = DuplicateStandardHandle(StdErrorHandle, GenericWrite);
                    }

                    childHandles.Add(childInput);
                    childHandles.Add(childOutput);
                    childHandles.Add(childError);
                }

                int attributeCount = useStandardHandles ? 2 : 1;
                InitializeProcThreadAttributeList(IntPtr.Zero, attributeCount, 0, ref attributeListSize);
                attributeList = Marshal.AllocHGlobal(attributeListSize);
                if (!InitializeProcThreadAttributeList(attributeList, attributeCount, 0, ref attributeListSize))
                {
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "InitializeProcThreadAttributeList failed.");
                }

                jobValue = Marshal.AllocHGlobal(IntPtr.Size);
                Marshal.WriteIntPtr(jobValue, job.Handle.DangerousGetHandle());
                UpdateAttribute(attributeList, ProcThreadAttributeJobList, jobValue, new IntPtr(IntPtr.Size), "job list");

                if (useStandardHandles)
                {
                    handleValues = Marshal.AllocHGlobal(IntPtr.Size * childHandles.Count);
                    for (int index = 0; index < childHandles.Count; ++index)
                    {
                        Marshal.WriteIntPtr(handleValues, index * IntPtr.Size, childHandles[index].DangerousGetHandle());
                    }
                    UpdateAttribute(
                        attributeList,
                        ProcThreadAttributeHandleList,
                        handleValues,
                        new IntPtr(IntPtr.Size * childHandles.Count),
                        "handle list");
                }

                StartupInfoEx startupInfo = new StartupInfoEx();
                startupInfo.StartupInfo.Size = Marshal.SizeOf(typeof(StartupInfoEx));
                startupInfo.AttributeList = attributeList;
                if (useStandardHandles)
                {
                    startupInfo.StartupInfo.Flags |= StartfUseStdHandles;
                    startupInfo.StartupInfo.StandardInput = childInput.DangerousGetHandle();
                    startupInfo.StartupInfo.StandardOutput = childOutput.DangerousGetHandle();
                    startupInfo.StartupInfo.StandardError = childError.DangerousGetHandle();
                }
                if (startInfo.WindowStyle == ProcessWindowStyle.Hidden)
                {
                    startupInfo.StartupInfo.Flags |= StartfUseShowWindow;
                    startupInfo.StartupInfo.ShowWindow = SwHide;
                }

                environment = BuildEnvironmentBlock(startInfo);
                uint creationFlags = ExtendedStartupInfoPresent | CreateUnicodeEnvironment;
                if (startInfo.CreateNoWindow)
                {
                    creationFlags |= CreateNoWindow;
                }

                ProcessInformation processInformation;
                StringBuilder mutableCommandLine = new StringBuilder(commandLine);
                string workingDirectory = String.IsNullOrWhiteSpace(startInfo.WorkingDirectory)
                    ? Environment.CurrentDirectory
                    : startInfo.WorkingDirectory;
                string applicationName = startInfo.FileName.IndexOf('\\') >= 0 || startInfo.FileName.IndexOf('/') >= 0
                    ? startInfo.FileName
                    : null;
                if (!CreateProcessW(
                        applicationName,
                        mutableCommandLine,
                        IntPtr.Zero,
                        IntPtr.Zero,
                        useStandardHandles,
                        creationFlags,
                        environment,
                        workingDirectory,
                        ref startupInfo,
                        out processInformation))
                {
                    throw new Win32Exception(Marshal.GetLastWin32Error(), "CreateProcessW failed.");
                }

                process = new SafeFileHandle(processInformation.Process, true);
                thread = new SafeFileHandle(processInformation.Thread, true);
                DisposeHandle(ref childInput);
                DisposeHandle(ref childOutput);
                DisposeHandle(ref childError);

                if (parentInput != null)
                {
                    FileStream inputStream = new FileStream(parentInput, FileAccess.Write, 4096, false);
                    parentInput = null;
                    inputWriter = new StreamWriter(inputStream, new UTF8Encoding(false));
                    inputWriter.AutoFlush = true;
                }
                if (parentOutput != null)
                {
                    FileStream outputStream = new FileStream(parentOutput, FileAccess.Read, 4096, false);
                    parentOutput = null;
                    Encoding outputEncoding = startInfo.StandardOutputEncoding ?? Console.OutputEncoding;
                    outputReader = new StreamReader(outputStream, outputEncoding, true, 4096, false);
                }
                if (parentError != null)
                {
                    FileStream errorStream = new FileStream(parentError, FileAccess.Read, 4096, false);
                    parentError = null;
                    Encoding errorEncoding = startInfo.StandardErrorEncoding ?? Console.OutputEncoding;
                    errorReader = new StreamReader(errorStream, errorEncoding, true, 4096, false);
                }

                ContainedProcess result = new ContainedProcess(
                    job,
                    process,
                    processInformation.ProcessId,
                    inputWriter,
                    outputReader,
                    errorReader);
                job = null;
                process = null;
                inputWriter = null;
                outputReader = null;
                errorReader = null;
                return result;
            }
            finally
            {
                DisposeHandle(ref thread);
                DisposeHandle(ref childInput);
                DisposeHandle(ref childOutput);
                DisposeHandle(ref childError);
                DisposeHandle(ref parentInput);
                DisposeHandle(ref parentOutput);
                DisposeHandle(ref parentError);
                if (inputWriter != null) inputWriter.Dispose();
                if (outputReader != null) outputReader.Dispose();
                if (errorReader != null) errorReader.Dispose();
                DisposeHandle(ref process);
                if (job != null) job.Dispose();
                if (attributeList != IntPtr.Zero) DeleteProcThreadAttributeList(attributeList);
                if (attributeList != IntPtr.Zero) Marshal.FreeHGlobal(attributeList);
                if (jobValue != IntPtr.Zero) Marshal.FreeHGlobal(jobValue);
                if (handleValues != IntPtr.Zero) Marshal.FreeHGlobal(handleValues);
                if (environment != IntPtr.Zero) Marshal.FreeHGlobal(environment);
            }
        }

        public bool WaitForExit(int milliseconds)
        {
            uint timeout = milliseconds < 0 ? Infinite : unchecked((uint)milliseconds);
            uint result = WaitForSingleObject(process, timeout);
            if (result == WaitObject0) return true;
            if (result == WaitTimeout) return false;
            throw new Win32Exception(Marshal.GetLastWin32Error(), "WaitForSingleObject failed.");
        }

        public void WaitForExit()
        {
            if (!WaitForExit(-1))
            {
                throw new InvalidOperationException("Infinite process wait returned without process exit.");
            }
        }

        public void Refresh()
        {
            // Native process state is queried directly from the retained process handle.
        }

        public void Dispose()
        {
            if (Interlocked.Exchange(ref disposed, 1) != 0)
            {
                return;
            }

            if (job != null)
            {
                job.Dispose();
                job = null;
            }
            if (process != null && !process.IsInvalid && !process.IsClosed)
            {
                WaitForSingleObject(process, 5000);
            }
            if (standardInput != null) standardInput.Dispose();
            if (standardOutput != null) standardOutput.Dispose();
            if (standardError != null) standardError.Dispose();
            standardInput = null;
            standardOutput = null;
            standardError = null;
            DisposeHandle(ref process);
            GC.SuppressFinalize(this);
        }

        private static void UpdateAttribute(
            IntPtr attributeList,
            IntPtr attribute,
            IntPtr value,
            IntPtr size,
            string description)
        {
            if (!UpdateProcThreadAttribute(
                    attributeList,
                    0,
                    attribute,
                    value,
                    size,
                    IntPtr.Zero,
                    IntPtr.Zero))
            {
                throw new Win32Exception(
                    Marshal.GetLastWin32Error(),
                    "UpdateProcThreadAttribute failed for " + description + ".");
            }
        }

        private static void CreateOutputPipe(out SafeFileHandle parentRead, out SafeFileHandle childWrite)
        {
            SecurityAttributes attributes = NewInheritableSecurityAttributes();
            if (!CreatePipe(out parentRead, out childWrite, ref attributes, 0))
            {
                throw new Win32Exception(Marshal.GetLastWin32Error(), "CreatePipe failed.");
            }
            if (!SetHandleInformation(parentRead, HandleFlagInherit, 0))
            {
                parentRead.Dispose();
                childWrite.Dispose();
                throw new Win32Exception(Marshal.GetLastWin32Error(), "SetHandleInformation failed.");
            }
        }

        private static void CreateInputPipe(out SafeFileHandle childRead, out SafeFileHandle parentWrite)
        {
            SecurityAttributes attributes = NewInheritableSecurityAttributes();
            if (!CreatePipe(out childRead, out parentWrite, ref attributes, 0))
            {
                throw new Win32Exception(Marshal.GetLastWin32Error(), "CreatePipe failed.");
            }
            if (!SetHandleInformation(parentWrite, HandleFlagInherit, 0))
            {
                childRead.Dispose();
                parentWrite.Dispose();
                throw new Win32Exception(Marshal.GetLastWin32Error(), "SetHandleInformation failed.");
            }
        }

        private static SafeFileHandle DuplicateStandardHandle(int standardHandle, uint fallbackAccess)
        {
            IntPtr source = GetStdHandle(standardHandle);
            SafeFileHandle duplicate;
            IntPtr currentProcess = GetCurrentProcess();
            if (source != IntPtr.Zero && source != new IntPtr(-1) &&
                DuplicateHandle(
                    currentProcess,
                    source,
                    currentProcess,
                    out duplicate,
                    0,
                    true,
                    DuplicateSameAccess))
            {
                return duplicate;
            }

            SecurityAttributes attributes = NewInheritableSecurityAttributes();
            SafeFileHandle nullHandle = CreateFileW(
                "NUL",
                fallbackAccess,
                FileShareRead | FileShareWrite,
                ref attributes,
                OpenExisting,
                0,
                IntPtr.Zero);
            if (nullHandle.IsInvalid)
            {
                int error = Marshal.GetLastWin32Error();
                nullHandle.Dispose();
                throw new Win32Exception(error, "Failed to duplicate a standard handle or open NUL.");
            }
            return nullHandle;
        }

        private static SecurityAttributes NewInheritableSecurityAttributes()
        {
            SecurityAttributes attributes = new SecurityAttributes();
            attributes.Length = Marshal.SizeOf(typeof(SecurityAttributes));
            attributes.InheritHandle = true;
            return attributes;
        }

        private static IntPtr BuildEnvironmentBlock(ProcessStartInfo startInfo)
        {
            List<string> names = new List<string>();
            foreach (string name in startInfo.EnvironmentVariables.Keys)
            {
                names.Add(name);
            }
            names.Sort(StringComparer.OrdinalIgnoreCase);

            StringBuilder block = new StringBuilder();
            foreach (string name in names)
            {
                string value = startInfo.EnvironmentVariables[name] ?? String.Empty;
                if (name.IndexOf('\0') >= 0 || value.IndexOf('\0') >= 0)
                {
                    throw new InvalidOperationException("Process environment names and values cannot contain embedded NUL characters.");
                }
                block.Append(name);
                block.Append('=');
                block.Append(value);
                block.Append('\0');
            }
            block.Append('\0');
            if (names.Count == 0)
            {
                block.Append('\0');
            }
            return Marshal.StringToHGlobalUni(block.ToString());
        }

        private static void DisposeHandle(ref SafeFileHandle handle)
        {
            if (handle != null)
            {
                handle.Dispose();
                handle = null;
            }
        }
    }
}
'@
}

function Get-RSCanonicalEnvironmentName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if ([string]::Equals($Name, 'PATH', [System.StringComparison]::OrdinalIgnoreCase)) {
        return 'Path'
    }

    return $Name
}

function Add-RSPathSegments {
    param(
        [System.Collections.Generic.List[string]]$PathSegments,

        [System.Collections.Generic.HashSet[string]]$SeenPathSegments,

        [AllowNull()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return
    }

    foreach ($segment in ($Value -split ';')) {
        $trimmedSegment = $segment.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmedSegment)) {
            continue
        }

        if ($SeenPathSegments.Add($trimmedSegment)) {
            [void]$PathSegments.Add($trimmedSegment)
        }
    }
}

function Get-RSNormalizedEnvironmentMap {
    param(
        [AllowNull()]
        [System.Collections.IDictionary]$ProcessEnvironment = $null,

        [AllowNull()]
        [string]$ProcessPathValue = $null
    )

    $environment = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $pathSegments = [System.Collections.Generic.List[string]]::new()
    $seenPathSegments = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    if ($null -eq $ProcessEnvironment) {
        $ProcessEnvironment = [System.Environment]::GetEnvironmentVariables('Process')
    }

    if ($null -eq $ProcessPathValue) {
        $ProcessPathValue = [System.Environment]::GetEnvironmentVariable('Path', 'Process')
    }

    Add-RSPathSegments -PathSegments $pathSegments -SeenPathSegments $seenPathSegments -Value $ProcessPathValue

    foreach ($entry in $ProcessEnvironment.GetEnumerator()) {
        $name = [string]$entry.Key
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        $canonicalName = Get-RSCanonicalEnvironmentName -Name $name
        $value = [string]$entry.Value
        if ([string]::Equals($canonicalName, 'Path', [System.StringComparison]::OrdinalIgnoreCase)) {
            Add-RSPathSegments -PathSegments $pathSegments -SeenPathSegments $seenPathSegments -Value $value
        }
        else {
            $environment[$canonicalName] = $value
        }
    }

    if ($pathSegments.Count -gt 0) {
        $environment['Path'] = ($pathSegments -join ';')
    }

    return $environment
}

function ConvertTo-RSQuotedProcessArgument {
    param(
        [AllowNull()]
        [string]$Argument
    )

    if ($null -eq $Argument -or $Argument.Length -eq 0) {
        return '""'
    }

    if ($Argument -notmatch '[\s"]') {
        return $Argument
    }

    $builder = [System.Text.StringBuilder]::new()
    [void]$builder.Append('"')
    $backslashCount = 0
    foreach ($character in $Argument.ToCharArray()) {
        if ($character -eq '\') {
            $backslashCount++
            continue
        }

        if ($character -eq '"') {
            if ($backslashCount -gt 0) {
                [void]$builder.Append(('\' * ($backslashCount * 2)))
                $backslashCount = 0
            }

            [void]$builder.Append('\"')
            continue
        }

        if ($backslashCount -gt 0) {
            [void]$builder.Append(('\' * $backslashCount))
            $backslashCount = 0
        }

        [void]$builder.Append($character)
    }

    if ($backslashCount -gt 0) {
        [void]$builder.Append(('\' * ($backslashCount * 2)))
    }

    [void]$builder.Append('"')
    return $builder.ToString()
}

function Set-RSProcessArguments {
    param(
        [Parameter(Mandatory = $true)]
        [System.Diagnostics.ProcessStartInfo]$ProcessStartInfo,

        [string[]]$Arguments = @()
    )

    $argumentListProperty = $ProcessStartInfo.PSObject.Properties['ArgumentList']
    if ($argumentListProperty -and $null -ne $ProcessStartInfo.ArgumentList) {
        foreach ($argument in $Arguments) {
            [void]$ProcessStartInfo.ArgumentList.Add($argument)
        }
        return
    }

    $ProcessStartInfo.Arguments = (($Arguments | ForEach-Object { ConvertTo-RSQuotedProcessArgument $_ }) -join ' ')
}

function Get-RSProcessEnvironmentDictionary {
    param(
        [Parameter(Mandatory = $true)]
        [System.Diagnostics.ProcessStartInfo]$ProcessStartInfo
    )

    $environmentProperty = $ProcessStartInfo.PSObject.Properties['Environment']
    if ($environmentProperty -and $null -ne $ProcessStartInfo.Environment) {
        return ,$ProcessStartInfo.Environment
    }

    $environmentVariablesProperty = $ProcessStartInfo.PSObject.Properties['EnvironmentVariables']
    if ($environmentVariablesProperty -and $null -ne $ProcessStartInfo.EnvironmentVariables) {
        return ,$ProcessStartInfo.EnvironmentVariables
    }

    return ,$ProcessStartInfo.EnvironmentVariables
}

function Test-RSProcessEnvironmentContains {
    param(
        [Parameter(Mandatory = $true)]
        [object]$EnvironmentDictionary,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if ($null -ne $EnvironmentDictionary.PSObject.Methods['ContainsKey']) {
        return [bool]$EnvironmentDictionary.ContainsKey($Name)
    }

    return [bool]$EnvironmentDictionary.Contains($Name)
}

function Set-RSProcessEnvironmentValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$EnvironmentDictionary,

        [Parameter(Mandatory = $true)]
        [string]$Name,

        [AllowNull()]
        [string]$Value
    )

    $canonicalName = Get-RSCanonicalEnvironmentName -Name $Name

    if ($null -ne $Value) {
        $EnvironmentDictionary[$canonicalName] = $Value
        return
    }

    if (Test-RSProcessEnvironmentContains -EnvironmentDictionary $EnvironmentDictionary -Name $canonicalName) {
        [void]$EnvironmentDictionary.Remove($canonicalName)
    }

    if (-not [string]::Equals($canonicalName, $Name, [System.StringComparison]::Ordinal) -and
        (Test-RSProcessEnvironmentContains -EnvironmentDictionary $EnvironmentDictionary -Name $Name)) {
        [void]$EnvironmentDictionary.Remove($Name)
    }
}

function New-RSProcessStartInfo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [string[]]$Arguments = @(),

        [string]$WorkingDirectory = (Get-Location).Path,

        [hashtable]$AdditionalEnvironment = @{}
    )

    $psi = [System.Diagnostics.ProcessStartInfo]::new()
    $psi.FileName = $FilePath
    Set-RSProcessArguments -ProcessStartInfo $psi -Arguments $Arguments

    $psi.WorkingDirectory = $WorkingDirectory
    $psi.UseShellExecute = $false
    $environmentDictionary = Get-RSProcessEnvironmentDictionary -ProcessStartInfo $psi
    $environmentDictionary.Clear()
    foreach ($entry in (Get-RSNormalizedEnvironmentMap).GetEnumerator()) {
        Set-RSProcessEnvironmentValue -EnvironmentDictionary $environmentDictionary -Name $entry.Key -Value $entry.Value
    }

    foreach ($entry in $AdditionalEnvironment.GetEnumerator()) {
        $name = [string]$entry.Key
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        if ($null -eq $entry.Value) {
            Set-RSProcessEnvironmentValue -EnvironmentDictionary $environmentDictionary -Name $name -Value $null
            continue
        }

        Set-RSProcessEnvironmentValue -EnvironmentDictionary $environmentDictionary -Name $name -Value ([string]$entry.Value)
    }

    return $psi
}

function Get-RSProcessCommandLine {
    param(
        [Parameter(Mandatory = $true)]
        [System.Diagnostics.ProcessStartInfo]$ProcessStartInfo
    )

    $commandLine = [System.Text.StringBuilder]::new()
    [void]$commandLine.Append((ConvertTo-RSQuotedProcessArgument -Argument $ProcessStartInfo.FileName))

    $argumentListProperty = $ProcessStartInfo.PSObject.Properties['ArgumentList']
    if ($argumentListProperty -and $null -ne $ProcessStartInfo.ArgumentList -and $ProcessStartInfo.ArgumentList.Count -gt 0) {
        foreach ($argument in $ProcessStartInfo.ArgumentList) {
            [void]$commandLine.Append(' ')
            [void]$commandLine.Append((ConvertTo-RSQuotedProcessArgument -Argument ([string]$argument)))
        }
    }
    elseif (-not [string]::IsNullOrWhiteSpace($ProcessStartInfo.Arguments)) {
        [void]$commandLine.Append(' ')
        [void]$commandLine.Append($ProcessStartInfo.Arguments)
    }

    return $commandLine.ToString()
}

function Start-RSContainedProcess {
    param(
        [Parameter(Mandatory = $true)]
        [System.Diagnostics.ProcessStartInfo]$ProcessStartInfo,

        [switch]$DelegateArtifactOperation
    )

    $delegation = $null
    $process = $null
    $environmentDictionary = $null
    $savedEnvironment = [System.Collections.Generic.List[object]]::new()
    try {
        try {
            if ($DelegateArtifactOperation -and
                $null -ne (Get-Command New-RSArtifactOperationChildDelegation -ErrorAction SilentlyContinue)) {
                $delegation = New-RSArtifactOperationChildDelegation
                if ($null -ne $delegation) {
                    $environmentDictionary = Get-RSProcessEnvironmentDictionary -ProcessStartInfo $ProcessStartInfo
                    foreach ($entry in $delegation.Environment.GetEnumerator()) {
                        $name = Get-RSCanonicalEnvironmentName -Name ([string]$entry.Key)
                        $hadValue = Test-RSProcessEnvironmentContains `
                            -EnvironmentDictionary $environmentDictionary `
                            -Name $name
                        $previousValue = if ($hadValue) { [string]$environmentDictionary[$name] } else { $null }
                        $savedEnvironment.Add([pscustomobject]@{
                                Name = $name
                                HadValue = $hadValue
                                Value = $previousValue
                            })
                        Set-RSProcessEnvironmentValue `
                            -EnvironmentDictionary $environmentDictionary `
                            -Name $name `
                            -Value ([string]$entry.Value)
                    }
                }
            }

            $commandLine = Get-RSProcessCommandLine -ProcessStartInfo $ProcessStartInfo
            $process = [RedSalamander.Tooling.ContainedProcess]::Start($ProcessStartInfo, $commandLine)

            if ($null -ne $delegation) {
                if (-not $delegation.ReadyEvent.Set()) {
                    throw 'Failed to signal the contained artifact-operation child delegation.'
                }
                $process.ArtifactOperationDelegation = $delegation
            }
        }
        finally {
            if ($null -ne $environmentDictionary) {
                foreach ($saved in $savedEnvironment) {
                    $restoreValue = if ($saved.HadValue) { [string]$saved.Value } else { $null }
                    Set-RSProcessEnvironmentValue `
                        -EnvironmentDictionary $environmentDictionary `
                        -Name ([string]$saved.Name) `
                        -Value $restoreValue
                }
            }
        }
    }
    catch {
        if ($null -ne $process) {
            $process.Dispose()
        }
        if ($null -ne $delegation) {
            Complete-RSArtifactOperationChildDelegation -Delegation $delegation
        }
        throw
    }

    return $process
}

function Close-RSContainedProcess {
    param(
        [AllowNull()]
        [RedSalamander.Tooling.ContainedProcess]$Process
    )

    if ($null -eq $Process) {
        return
    }

    $delegation = $Process.ArtifactOperationDelegation
    $Process.ArtifactOperationDelegation = $null

    try {
        $Process.Dispose()
    }
    finally {
        if ($null -ne $delegation -and
            $null -ne (Get-Command Complete-RSArtifactOperationChildDelegation -ErrorAction SilentlyContinue)) {
            Complete-RSArtifactOperationChildDelegation -Delegation $delegation
        }
    }
}

function Invoke-RSProcess {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [string[]]$Arguments = @(),

        [string]$WorkingDirectory = (Get-Location).Path,

        [hashtable]$AdditionalEnvironment = @{}
    )

    $psi = New-RSProcessStartInfo `
        -FilePath $FilePath `
        -Arguments $Arguments `
        -WorkingDirectory $WorkingDirectory `
        -AdditionalEnvironment $AdditionalEnvironment

    $psi.CreateNoWindow = $false

    $process = $null

    try {
        $process = Start-RSContainedProcess -ProcessStartInfo $psi -DelegateArtifactOperation
        $process.WaitForExit()
        $global:LASTEXITCODE = $process.ExitCode
        return $process.ExitCode
    }
    finally {
        Close-RSContainedProcess -Process $process
    }
}
