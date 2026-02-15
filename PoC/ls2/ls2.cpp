// ntfastwalk.cpp — super-fast directory lister using NtQueryDirectoryFile
// Build: cl /O2 /std:c++20 ntfastwalk.cpp
#include <cstdio>
#include <string>
#include <vector>
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <winternl.h> // NTSTATUS, IO_STATUS_BLOCK, FILE_INFORMATION_CLASS, etc.

#ifndef NT_SUCCESS
#define NT_SUCCESS(Status) (((NTSTATUS)(Status)) >= 0)
#endif
// Minimal status we care about (avoid ntstatus.h include headaches)
#ifndef STATUS_NO_MORE_FILES
#define STATUS_NO_MORE_FILES ((NTSTATUS)0x80000006L)
#endif

// NtQueryDirectoryFile typedef (we’ll fetch from ntdll at runtime)
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

// We’ll use FILE_BOTH_DIR_INFORMATION for broad compatibility (NTFS & ReFS).
// (If you want 128-bit IDs, switch to FileIdExtdDirectoryInformation
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

static inline bool IsDotOrDotDot(const wchar_t* name)
{
    return (name[0] == L'.' && (name[1] == L'\0' || (name[1] == L'.' && name[2] == L'\0')));
}

// Convert to extended-length path so we bypass legacy MAX_PATH and normalization overhead.
static std::wstring ToExtendedPath(const std::wstring& p)
{
    if (p.rfind(L"\\\\?\\", 0) == 0)
        return p; // already extended
    if (p.rfind(L"\\\\", 0) == 0)
    {
        // \\server\share\... -> \\?\UNC\server\share\...
        return L"\\\\?\\UNC\\" + p.substr(2);
    }
    return L"\\\\?\\" + p;
}

// Append child to base with exactly one backslash.
static std::wstring JoinPath(const std::wstring& base, const std::wstring& child)
{
    if (base.empty())
        return child;
    wchar_t last   = base.back();
    bool needSlash = (last != L'\\' && last != L'/');
    return needSlash ? (base + L'\\' + child) : (base + child);
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

static HANDLE OpenDirHandle(const std::wstring& path)
{
    return CreateFileW(path.c_str(),
                       FILE_LIST_DIRECTORY, // enumerate
                       FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                       nullptr,
                       OPEN_EXISTING,
                       FILE_FLAG_BACKUP_SEMANTICS | FILE_ATTRIBUTE_NORMAL,
                       nullptr);
}

static void WalkTree_NtQuery(const std::wstring& startInput)
{
    // Resolve starting root and open handle
    std::wstring start    = startInput.empty() ? L"." : startInput;
    std::wstring rootPath = ToExtendedPath(start);

    SetErrorMode(SEM_FAILCRITICALERRORS | SEM_NOOPENFILEERRORBOX);
    EnsureConsoleUtf8();

    // Load NtQueryDirectoryFile
    HMODULE ntdll = GetModuleHandleW(L"ntdll.dll");
    if (! ntdll)
        ntdll = LoadLibrary(L"ntdll.dll");
    if (! ntdll)
        return;
    auto NtQueryDirectoryFile = reinterpret_cast<NtQueryDirectoryFile_t>(GetProcAddress(ntdll, "NtQueryDirectoryFile"));
    if (! NtQueryDirectoryFile)
        return;

    // Print the starting directory, then DFS using explicit stack of paths.
    WriteLineUtf8(start);

    std::vector<std::wstring> stack;
    stack.reserve(4096);
    stack.push_back(rootPath);

    // Big buffer to reduce syscalls (512 KiB is a good sweet spot; 1 MiB also fine)
    const ULONG BUF_SIZE = 512 * 1024;
    void* buffer         = VirtualAlloc(nullptr, BUF_SIZE, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE);
    if (! buffer)
        return;

    IO_STATUS_BLOCK iosb{};

    while (! stack.empty())
    {
        std::wstring dirPath = std::move(stack.back());
        stack.pop_back();

        HANDLE hDir = OpenDirHandle(dirPath);
        if (hDir == INVALID_HANDLE_VALUE)
        {
            continue; // access denied, disappeared, etc.
        }

        BOOLEAN restart = TRUE;

        for (;;)
        {
            NTSTATUS st = NtQueryDirectoryFile(hDir,
                                               nullptr,
                                               nullptr,
                                               nullptr,
                                               &iosb,
                                               buffer,
                                               BUF_SIZE,
                                               (FILE_INFORMATION_CLASS)FileBothDirectoryInformation, // broad compatibility (NTFS/ReFS)
                                               FALSE,                                                // ReturnSingleEntry = FALSE => batch
                                               nullptr,                                              // no pattern => all
                                               restart                                               // restart on first call
            );
            restart     = FALSE;

            if (st == STATUS_NO_MORE_FILES)
                break;
            if (! NT_SUCCESS(st) || iosb.Information == 0)
                break;

            // Walk the MULTI-SZ-like linked list via NextEntryOffset
            BYTE* base   = static_cast<BYTE*>(buffer);
            ULONG offset = 0;

            for (;;)
            {
                auto* info = reinterpret_cast<PFILE_BOTH_DIR_INFORMATION>(base + offset);

                // Extract name
                const WCHAR* fname     = info->FileName;
                const ULONG fnameBytes = info->FileNameLength; // bytes, UTF-16LE
                std::wstring name(fname, fname + (fnameBytes / 2));

                if (! IsDotOrDotDot(name.c_str()))
                {
                    std::wstring full = JoinPath(dirPath, name);
                    WriteLineUtf8(full);

                    // Recurse into real directories (skip reparse points: symlinks/junctions)
                    if ((info->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) && ((info->FileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) == 0))
                    {
                        stack.push_back(full);
                    }
                }

                if (info->NextEntryOffset == 0)
                    break;
                offset += info->NextEntryOffset;
            }
        }

        CloseHandle(hDir);
    }

    VirtualFree(buffer, 0, MEM_RELEASE);
}

int wmain(int argc, wchar_t** argv)
{
    std::wstring start = (argc >= 2) ? argv[1] : L".";
    WalkTree_NtQuery(start);
    return 0;
}
