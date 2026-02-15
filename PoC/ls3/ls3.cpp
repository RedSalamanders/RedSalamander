// ntfastwalk_rel.cpp � ultra-fast lister using NtQueryDirectoryFile + relative NtCreateFile
// Build: cl /O2 /std:c++20 ntfastwalk_rel.cpp
#include <string>
#include <vector>
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <winternl.h> // NTSTATUS, IO_STATUS_BLOCK, UNICODE_STRING, OBJECT_ATTRIBUTES

// --- Minimal NT Helpers/defs (avoid extra headers) ---------------------------
#ifndef NT_SUCCESS
#define NT_SUCCESS(Status) (((NTSTATUS)(Status)) >= 0)
#endif
#ifndef STATUS_NO_MORE_FILES
#define STATUS_NO_MORE_FILES ((NTSTATUS)0x80000006L)
#endif
#ifndef FILE_OPEN
#define FILE_OPEN 0x00000001
#endif
#ifndef FILE_DIRECTORY_FILE
#define FILE_DIRECTORY_FILE 0x00000001
#endif
#ifndef FILE_SYNCHRONOUS_IO_NONALERT
#define FILE_SYNCHRONOUS_IO_NONALERT 0x00000020
#endif
#ifndef FILE_OPEN_FOR_BACKUP_INTENT
#define FILE_OPEN_FOR_BACKUP_INTENT 0x00004000
#endif
#ifndef OBJ_CASE_INSENSITIVE
#define OBJ_CASE_INSENSITIVE 0x00000040L
#endif

using NtQueryDirectoryFile_t = NTSTATUS(NTAPI*)(HANDLE FileHandle,
                                                HANDLE Event,
                                                PIO_APC_ROUTINE ApcRoutine,
                                                PVOID ApcContext,
                                                PIO_STATUS_BLOCK IoStatusBlock,
                                                PVOID FileInformation,
                                                ULONG Length,
                                                FILE_INFORMATION_CLASS FileInformationClass,
                                                BOOLEAN ReturnSingleEntry,
                                                PUNICODE_STRING FileName,
                                                BOOLEAN RestartScan);

using NtCreateFile_t = NTSTATUS(NTAPI*)(PHANDLE FileHandle,
                                        ACCESS_MASK DesiredAccess,
                                        POBJECT_ATTRIBUTES ObjectAttributes,
                                        PIO_STATUS_BLOCK IoStatusBlock,
                                        PLARGE_INTEGER AllocationSize,
                                        ULONG FileAttributes,
                                        ULONG ShareAccess,
                                        ULONG CreateDisposition,
                                        ULONG CreateOptions,
                                        PVOID EaBuffer,
                                        ULONG EaLength);

typedef struct _FILE_BOTH_DIR_INFORMATION
{
    ULONG NextEntryOffset;
    ULONG FileIndex;
    LARGE_INTEGER CreationTime;
    LARGE_INTEGER LastAccessTime;
    LARGE_INTEGER LastWriteTime;
    LARGE_INTEGER ChangeTime;
    LARGE_INTEGER EndOfFile;
    LARGE_INTEGER AllocationSize;
    ULONG FileAttributes;
    ULONG FileNameLength;
    ULONG EaSize;
    CCHAR ShortNameLength;
    WCHAR ShortName[12];
    WCHAR FileName[1]; // variable length
} FILE_BOTH_DIR_INFORMATION, *PFILE_BOTH_DIR_INFORMATION;

typedef struct _FILE_NAMES_INFORMATION
{
    ULONG NextEntryOffset;
    ULONG FileIndex;
    ULONG FileNameLength;
    WCHAR FileName[1];
} FILE_NAMES_INFORMATION, *PFILE_NAMES_INFORMATION;

typedef enum _FILE_INFORMATION_CLASS_ALL // from ntifs.h
{
    FileDirectoryInformationFronNTifs = 1,
    FileFullDirectoryInformation,            // 2
    FileBothDirectoryInformation,            // 3
    FileBasicInformation,                    // 4
    FileStandardInformation,                 // 5
    FileInternalInformation,                 // 6
    FileEaInformation,                       // 7
    FileAccessInformation,                   // 8
    FileNameInformation,                     // 9
    FileRenameInformation,                   // 10
    FileLinkInformation,                     // 11
    FileNamesInformation,                    // 12
    FileDispositionInformation,              // 13
    FilePositionInformation,                 // 14
    FileFullEaInformation,                   // 15
    FileModeInformation,                     // 16
    FileAlignmentInformation,                // 17
    FileAllInformation,                      // 18
    FileAllocationInformation,               // 19
    FileEndOfFileInformation,                // 20
    FileAlternateNameInformation,            // 21
    FileStreamInformation,                   // 22
    FilePipeInformation,                     // 23
    FilePipeLocalInformation,                // 24
    FilePipeRemoteInformation,               // 25
    FileMailslotQueryInformation,            // 26
    FileMailslotSetInformation,              // 27
    FileCompressionInformation,              // 28
    FileObjectIdInformation,                 // 29
    FileCompletionInformation,               // 30
    FileMoveClusterInformation,              // 31
    FileQuotaInformation,                    // 32
    FileReparsePointInformation,             // 33
    FileNetworkOpenInformation,              // 34
    FileAttributeTagInformation,             // 35
    FileTrackingInformation,                 // 36
    FileIdBothDirectoryInformation,          // 37
    FileIdFullDirectoryInformation,          // 38
    FileValidDataLengthInformation,          // 39
    FileShortNameInformation,                // 40
    FileIoCompletionNotificationInformation, // 41
    FileIoStatusBlockRangeInformation,       // 42
    FileIoPriorityHintInformation,           // 43
    FileSfioReserveInformation,              // 44
    FileSfioVolumeInformation,               // 45
    FileHardLinkInformation,                 // 46
    FileProcessIdsUsingFileInformation,      // 47
    FileNormalizedNameInformation,           // 48
    FileNetworkPhysicalNameInformation,      // 49
    FileIdGlobalTxDirectoryInformation,      // 50
    FileIsRemoteDeviceInformation,           // 51
    FileUnusedInformation,                   // 52
    FileNumaNodeInformation,                 // 53
    FileStandardLinkInformation,             // 54
    FileRemoteProtocolInformation,           // 55

    //
    //  These are special versions of these operations (defined earlier)
    //  which can be used by kernel mode drivers only to bypass security
    //  access checks for Rename and HardLink operations.  These operations
    //  are only recognized by the IOManager, a file system should never
    //  receive these.
    //
    FileRenameInformationBypassAccessCheck, // 56
    FileLinkInformationBypassAccessCheck,   // 57
    FileVolumeNameInformation,              // 58
    FileIdInformation,                      // 59
    FileIdExtdDirectoryInformation,         // 60
    FileReplaceCompletionInformation,       // 61
    FileHardLinkFullIdInformation,          // 62
    FileIdExtdBothDirectoryInformation,     // 63
    FileMaximumInformation
} FILE_INFORMATION_CLASS_ALL,
    *PFILE_INFORMATION_CLASS_ALL;

// --- tiny utils --------------------------------------------------------------
static inline bool IsDotOrDotDot(const wchar_t* name)
{
    return (name[0] == L'.' && (name[1] == L'\0' || (name[1] == L'.' && name[2] == L'\0')));
}

static inline void EnsureConsoleUtf8()
{
    SetConsoleOutputCP(CP_UTF8);
}

static void WriteLineUtf8(const std::wstring& wline)
{
    int need = WideCharToMultiByte(CP_UTF8, 0, wline.c_str(), (int)wline.size(), nullptr, 0, nullptr, nullptr);
    if (need <= 0)
        return;
    std::string utf8;
    utf8.resize((size_t)need);
    WideCharToMultiByte(CP_UTF8, 0, wline.c_str(), (int)wline.size(), (LPSTR)(utf8.data()), need, nullptr, nullptr);
    DWORD written = 0;
    HANDLE h      = GetStdHandle(STD_OUTPUT_HANDLE);
    WriteFile(h, utf8.data(), (DWORD)utf8.size(), &written, nullptr);
    static const char nl = '\n';
    WriteFile(h, &nl, 1, &written, nullptr);
}

static std::wstring JoinPath(const std::wstring& base, const std::wstring& child)
{
    if (base.empty())
        return child;
    wchar_t last   = base.back();
    bool needSlash = (last != L'\\' && last != L'/');
    return needSlash ? (base + L'\\' + child) : (base + child);
}

// For long-path-safe root open (only used for the initial directory)
static std::wstring ToExtendedPath(const std::wstring& p)
{
    if (p.rfind(L"\\\\?\\", 0) == 0)
        return p;
    if (p.rfind(L"\\\\", 0) == 0)
        return L"\\\\?\\UNC\\" + p.substr(2);
    return L"\\\\?\\" + p;
}

static HANDLE OpenRootDirHandle(const std::wstring& path)
{
    std::wstring xp = ToExtendedPath(path);
    return CreateFileW(xp.c_str(),
                       FILE_LIST_DIRECTORY | SYNCHRONIZE,
                       FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                       nullptr,
                       OPEN_EXISTING,
                       FILE_FLAG_BACKUP_SEMANTICS, // open directory
                       nullptr);
}

// Open a CHILD directory by NAME **relative** to parent handle (no full path re-parse)
static HANDLE OpenSubdirRelative(NtCreateFile_t NtCreateFile, HANDLE parent, const std::wstring& childName)
{
    UNICODE_STRING us;
    us.Buffer        = const_cast<wchar_t*>(childName.c_str());
    us.Length        = (USHORT)(childName.size() * sizeof(wchar_t));
    us.MaximumLength = us.Length;

    OBJECT_ATTRIBUTES oa;
    InitializeObjectAttributes(&oa, &us, OBJ_CASE_INSENSITIVE, parent, nullptr);

    IO_STATUS_BLOCK iosb{};
    HANDLE h    = nullptr;
    NTSTATUS st = NtCreateFile(&h,
                               FILE_LIST_DIRECTORY | SYNCHRONIZE,
                               &oa,
                               &iosb,
                               nullptr,
                               FILE_ATTRIBUTE_NORMAL,
                               FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                               FILE_OPEN,
                               FILE_DIRECTORY_FILE | FILE_SYNCHRONOUS_IO_NONALERT | FILE_OPEN_FOR_BACKUP_INTENT,
                               nullptr,
                               0);
    return NT_SUCCESS(st) ? h : INVALID_HANDLE_VALUE;
}

// Frame keeps one open handle (depth-bounded), plus its child names to visit
struct Frame
{
    HANDLE hDir = INVALID_HANDLE_VALUE;
    std::wstring displayPath;          // for printing
    std::vector<std::wstring> subdirs; // names (relative to this frame)
    size_t next = 0;
};

// Enumerate a directory: print all entries with full display path,
// and collect immediate subdirectory **names** (skip reparse points).
static void EnumerateDir(NtQueryDirectoryFile_t NtQueryDirectoryFile,
                         HANDLE hDir,
                         const std::wstring& displayPath,
                         void* buffer,
                         ULONG bufSize,
                         std::vector<std::wstring>& outSubdirs)
{
    IO_STATUS_BLOCK iosb{};
    BOOLEAN restart = TRUE;

    for (;;)
    {
        NTSTATUS st = NtQueryDirectoryFile(hDir,
                                           nullptr,
                                           nullptr,
                                           nullptr,
                                           &iosb,
                                           buffer,
                                           bufSize,
                                           (FILE_INFORMATION_CLASS)FileBothDirectoryInformation, // portable across NTFS/ReFS
                                           FALSE,
                                           nullptr,
                                           restart);
        restart     = FALSE;

        if (st == STATUS_NO_MORE_FILES)
            break;
        if (! NT_SUCCESS(st) || iosb.Information == 0)
            break;

        BYTE* base   = static_cast<BYTE*>(buffer);
        ULONG offset = 0;

        for (;;)
        {
            auto* info             = reinterpret_cast<PFILE_BOTH_DIR_INFORMATION>(base + offset);
            const WCHAR* fname     = info->FileName;
            const ULONG fnameBytes = info->FileNameLength;
            std::wstring name(fname, fname + (fnameBytes / 2));

            if (! IsDotOrDotDot(name.c_str()))
            {
                // Print full path
                std::wstring full = JoinPath(displayPath, name);
                WriteLineUtf8(full);

                // Collect subdir names (avoid reparse points)
                if ((info->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) && ((info->FileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) == 0))
                {
                    outSubdirs.emplace_back(std::move(name));
                }
            }

            if (info->NextEntryOffset == 0)
                break;
            offset += info->NextEntryOffset;
        }
    }
}

int wmain(int argc, wchar_t** argv)
{
    SetErrorMode(SEM_FAILCRITICALERRORS | SEM_NOOPENFILEERRORBOX);
    EnsureConsoleUtf8();

    std::wstring start = (argc >= 2) ? argv[1] : L".";

    // Load NT functions
    HMODULE ntdll = GetModuleHandleW(L"ntdll.dll");
    if (! ntdll)
        ntdll = LoadLibrary(L"ntdll.dll");
    if (! ntdll)
        return 1;

    auto NtQueryDirectoryFile = reinterpret_cast<NtQueryDirectoryFile_t>(GetProcAddress(ntdll, "NtQueryDirectoryFile"));
    auto NtCreateFile         = reinterpret_cast<NtCreateFile_t>(GetProcAddress(ntdll, "NtCreateFile"));
    if (! NtQueryDirectoryFile || ! NtCreateFile)
        return 2;

    // Open root directory with Win32 (long-path-safe), then switch to relative opens
    HANDLE hRoot = OpenRootDirHandle(start);
    if (hRoot == INVALID_HANDLE_VALUE)
        return 3;

    // Big shared buffer for all queries
    const ULONG BUF_SIZE = 512 * 1024; // try 1*1024*1024 for giant dirs
    void* buffer         = VirtualAlloc(nullptr, BUF_SIZE, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE);
    if (! buffer)
    {
        CloseHandle(hRoot);
        return 4;
    }

    // Print the starting directory (like 'dir /s' headers)
    WriteLineUtf8(start);

    // Create first frame and enumerate it
    Frame root;
    root.hDir        = hRoot;
    root.displayPath = start;
    EnumerateDir(NtQueryDirectoryFile, root.hDir, root.displayPath, buffer, BUF_SIZE, root.subdirs);

    // Depth-first traversal, depth-bounded handle usage
    std::vector<Frame> stack;
    stack.reserve(256);
    stack.push_back(std::move(root));

    while (! stack.empty())
    {
        Frame& f = stack.back();

        if (f.next < f.subdirs.size())
        {
            // Open next child **relative** to current handle
            const std::wstring& childName = f.subdirs[f.next++];
            HANDLE hChild                 = OpenSubdirRelative(NtCreateFile, f.hDir, childName);
            if (hChild != INVALID_HANDLE_VALUE)
            {
                Frame child;
                child.hDir        = hChild;
                child.displayPath = JoinPath(f.displayPath, childName);
                EnumerateDir(NtQueryDirectoryFile, child.hDir, child.displayPath, buffer, BUF_SIZE, child.subdirs);
                stack.push_back(std::move(child));
            }
            // if open fails (ACL, transient), we just skip
        }
        else
        {
            // Done with this directory
            CloseHandle(f.hDir);
            stack.pop_back();
        }
    }

    VirtualFree(buffer, 0, MEM_RELEASE);
    return 0;
}
