#include "FileSystem.Internal.h"

#include <algorithm>
#include <chrono>
#include <limits>
#include <memory>
#include <string_view>
#include <vector>

#include <lm.h>       // NetShareEnum, NetApiBufferFree
#include <winternl.h> // NTSTATUS, IO_STATUS_BLOCK

using namespace FileSystemInternal;

namespace
{
constexpr size_t kDefaultBufferSize = 512 * 1024;

static_assert(sizeof(FileInfo) == sizeof(FILE_FULL_DIR_INFO), "FileInfo must match FILE_FULL_DIR_INFO layout.");
static_assert(offsetof(FileInfo, FileName) == offsetof(FILE_FULL_DIR_INFO, FileName), "FileInfo must match FILE_FULL_DIR_INFO layout.");
static_assert(alignof(FileInfo) == alignof(FILE_FULL_DIR_INFO), "FileInfo must match FILE_FULL_DIR_INFO alignment.");

constexpr size_t kFileInfoEntryAlignment = alignof(FileInfo);
static_assert((kFileInfoEntryAlignment & (kFileInfoEntryAlignment - 1u)) == 0u, "FileInfo alignment must be power-of-two.");

#ifndef NT_SUCCESS
#define NT_SUCCESS(Status) (((NTSTATUS)(Status)) >= 0)
#endif

#ifndef STATUS_NO_MORE_FILES
#define STATUS_NO_MORE_FILES ((NTSTATUS)0x80000006L)
#endif

enum class NtFileInformationClass : int
{
    FileDirectoryInformation     = 1,
    FileFullDirectoryInformation = 2,
};

using NtQueryDirectoryFile_t = NTSTATUS(NTAPI*)(HANDLE FileHandle,
                                                HANDLE Event,
                                                PIO_APC_ROUTINE ApcRoutine,
                                                PVOID ApcContext,
                                                PIO_STATUS_BLOCK IoStatusBlock,
                                                PVOID FileInformation,
                                                ULONG Length,
                                                NtFileInformationClass FileInformationClass,
                                                BOOLEAN ReturnSingleEntry,
                                                PUNICODE_STRING FileName,
                                                BOOLEAN RestartScan);

using RtlNtStatusToDosError_t = ULONG(NTAPI*)(NTSTATUS Status);

NtQueryDirectoryFile_t GetNtQueryDirectoryFile() noexcept
{
    static const NtQueryDirectoryFile_t fn = []() noexcept -> NtQueryDirectoryFile_t
    {
        HMODULE ntdll = ::GetModuleHandleW(L"ntdll.dll");
        if (! ntdll)
        {
            return nullptr;
        }

#pragma warning(push)
#pragma warning(disable : 4191) // C4191: 'reinterpret_cast': unsafe conversion from 'FARPROC'
        return reinterpret_cast<NtQueryDirectoryFile_t>(::GetProcAddress(ntdll, "NtQueryDirectoryFile"));
#pragma warning(pop)
    }();

    return fn;
}

RtlNtStatusToDosError_t GetRtlNtStatusToDosError() noexcept
{
    static const RtlNtStatusToDosError_t fn = []() noexcept -> RtlNtStatusToDosError_t
    {
        HMODULE ntdll = ::GetModuleHandleW(L"ntdll.dll");
        if (! ntdll)
        {
            return nullptr;
        }

#pragma warning(push)
#pragma warning(disable : 4191) // C4191: 'reinterpret_cast': unsafe conversion from 'FARPROC'
        return reinterpret_cast<RtlNtStatusToDosError_t>(::GetProcAddress(ntdll, "RtlNtStatusToDosError"));
#pragma warning(pop)
    }();

    return fn;
}

template <typename T> constexpr T AlignUp(T value, size_t alignment) noexcept
{
    const size_t mask = alignment - 1;
    return static_cast<T>((static_cast<size_t>(value) + mask) & ~mask);
}

bool IsExtendedUncPath(std::wstring_view path) noexcept
{
    return path.rfind(L"\\\\?\\UNC\\", 0) == 0;
}

bool IsExtendedWslPath(std::wstring_view path) noexcept
{
    // WSL provider uses UNC-style paths, e.g. \\wsl.localhost\\Ubuntu or \\wsl$\\Ubuntu
    return path.rfind(L"\\\\?\\UNC\\wsl.localhost\\", 0) == 0 || path.rfind(L"\\\\?\\UNC\\wsl$\\", 0) == 0;
}

bool IsExtendedDriveLetterPath(std::wstring_view path) noexcept
{
    if (path.size() < 7)
    {
        return false;
    }

    if (path.rfind(L"\\\\?\\", 0) != 0)
    {
        return false;
    }

    const wchar_t drive = path[4];
    if (! ((drive >= L'A' && drive <= L'Z') || (drive >= L'a' && drive <= L'z')))
    {
        return false;
    }

    return path[5] == L':' && (path[6] == L'\\' || path[6] == L'/');
}

bool ShouldUseHandleEnumeration(std::wstring_view extendedPath) noexcept
{
    if (extendedPath.empty())
    {
        return false;
    }

    if (IsExtendedUncPath(extendedPath) || IsExtendedWslPath(extendedPath))
    {
        return false;
    }

    if (IsExtendedDriveLetterPath(extendedPath))
    {
        wchar_t root[]       = {extendedPath[4], L':', L'\\', L'\0'};
        const UINT driveType = ::GetDriveTypeW(root);
        if (driveType == DRIVE_REMOTE || driveType == DRIVE_UNKNOWN || driveType == DRIVE_NO_ROOT_DIR)
        {
            return false;
        }
    }

    return true;
}

__int64 FileTimeToInt64(const FILETIME& fileTime) noexcept
{
    ULARGE_INTEGER value{};
    value.LowPart  = fileTime.dwLowDateTime;
    value.HighPart = fileTime.dwHighDateTime;
    return static_cast<__int64>(value.QuadPart);
}
} // namespace

FilesInformation::FilesInformation()
{
    _buffer.resize(kDefaultBufferSize, std::byte{0});
}

FilesInformation::~FilesInformation()
{
    ResetDirectoryState(true);
}

HRESULT STDMETHODCALLTYPE FilesInformation::QueryInterface(REFIID riid, void** ppvObject) noexcept
{
    if (ppvObject == nullptr)
    {
        return E_POINTER;
    }

    if (riid == __uuidof(IUnknown) || riid == __uuidof(IFilesInformation))
    {
        *ppvObject = static_cast<IFilesInformation*>(this);
        AddRef();
        return S_OK;
    }

    *ppvObject = nullptr;
    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE FilesInformation::AddRef() noexcept
{
    return static_cast<ULONG>(_refCount.fetch_add(1, std::memory_order_relaxed) + 1);
}

ULONG STDMETHODCALLTYPE FilesInformation::Release() noexcept
{
    const ULONG current = static_cast<ULONG>(_refCount.fetch_sub(1, std::memory_order_acq_rel) - 1);
    if (current == 0)
    {
        delete this;
    }
    return current;
}

HRESULT STDMETHODCALLTYPE FilesInformation::GetBuffer(FileInfo** ppFileInfo) noexcept
{
    if (ppFileInfo == nullptr)
    {
        return E_POINTER;
    }

    if (_count == 0 || _usedBytes == 0)
    {
        *ppFileInfo = nullptr;
        return S_OK;
    }

    *ppFileInfo = reinterpret_cast<FileInfo*>(_buffer.data());
    return S_OK;
}

HRESULT FilesInformation::BeginWrite(FileInfo** ppFileInfo) noexcept
{
    if (ppFileInfo == nullptr)
    {
        return E_POINTER;
    }

    _count      = 0;
    _usedBytes  = 0;
    *ppFileInfo = reinterpret_cast<FileInfo*>(_buffer.data());
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FilesInformation::GetAllocatedSize(unsigned long* pSize) noexcept
{
    if (pSize == nullptr)
    {
        return E_POINTER;
    }

    if (_buffer.size() > std::numeric_limits<unsigned long>::max())
    {
        *pSize = 0;
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    *pSize = static_cast<unsigned long>(_buffer.size());
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FilesInformation::GetBufferSize(unsigned long* pSize) noexcept
{
    if (pSize == nullptr)
    {
        return E_POINTER;
    }

    *pSize = _usedBytes;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FilesInformation::GetCount(unsigned long* pCount) noexcept
{
    if (pCount == nullptr)
    {
        return E_POINTER;
    }

    *pCount = _count;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FilesInformation::Get(unsigned long index, FileInfo** ppEntry) noexcept
{
    if (ppEntry == nullptr)
    {
        return E_POINTER;
    }

    if (index >= _count || _usedBytes == 0)
    {
        return HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES);
    }

    return LocateEntry(index, ppEntry);
}

void FilesInformation::UpdateUsage(unsigned long bytesUsed, unsigned long count) noexcept
{
    _usedBytes = bytesUsed;
    _count     = count;
}

void FilesInformation::MaybeTrimBuffer() noexcept
{
    // Enumeration scratch state is not needed after ReadDirectoryInfo() completes; avoid holding extra memory/handles.
    _directoryHandle.reset();
    std::vector<std::byte>().swap(_enumerationBuffer);
    _enumerationBufferOffset     = 0;
    _enumerationBufferBytesValid = 0;
    _enumerationRestartScan      = true;
    _useHandleEnumeration        = false;

    const size_t allocated = _buffer.size();
    const size_t used      = static_cast<size_t>(_usedBytes);

    if (allocated == 0)
    {
        return;
    }

    if (used == 0)
    {
        // Empty directory: freeing the default 512KB is a big win for the global cache.
        Debug::Perf::Scope perf(L"FileSystem.DirectoryOps.TrimBuffer");
        perf.SetDetail(_requestedPath);
        perf.SetValue0(allocated);
        perf.SetValue1(allocated);
        _buffer.clear();
        _buffer.shrink_to_fit();
        return;
    }

    if (used > allocated)
    {
        return;
    }

    const size_t saved = allocated - used;
    if (saved == 0)
    {
        return;
    }

    // Heuristic: trimming reallocates+copies `used` bytes; only do it when the space win is meaningful.
    // - Require at least 25% waste.
    // - Require either "saved >= used" (win >= copy) or "saved >= 8 MiB" (large win).
    constexpr size_t kMinSavedBytes   = 128u * 1024u;
    constexpr size_t kLargeSavedBytes = 8u * 1024u * 1024u;

    if (saved < kMinSavedBytes)
    {
        return;
    }

    const bool hasMeaningfulWaste = (saved * 4u) >= allocated; // >= 25%
    const bool savedBeatsCopy     = saved >= used;
    const bool veryLargeSavings   = saved >= kLargeSavedBytes;

    if (! hasMeaningfulWaste || (! savedBeatsCopy && ! veryLargeSavings))
    {
        return;
    }

    Debug::Perf::Scope perf(L"FileSystem.DirectoryOps.TrimBuffer");
    perf.SetDetail(_requestedPath);
    perf.SetValue0(allocated);
    perf.SetValue1(saved);
    _buffer.resize(used);
    _buffer.shrink_to_fit();
}

void FilesInformation::ResetDirectoryState(bool clearPath) noexcept
{
    _findHandle.reset();
    _directoryHandle.reset();
    _hasPendingEntry             = false;
    _pendingEntry                = WIN32_FIND_DATAW{};
    _enumerationInitialized      = false;
    _enumerationComplete         = false;
    _useHandleEnumeration        = false;
    _enumerationRestartScan      = true;
    _enumerationBufferOffset     = 0;
    _enumerationBufferBytesValid = 0;
    if (clearPath)
    {
        _requestedPath.clear();
    }
}

void FilesInformation::ResizeBuffer(size_t newSize) noexcept
{
    _buffer.resize(newSize, std::byte{0});
}

bool FilesInformation::PathEquals(const std::wstring& other) const noexcept
{
    if (_requestedPath.empty())
    {
        return false;
    }
    return _wcsicmp(_requestedPath.c_str(), other.c_str()) == 0;
}

size_t FilesInformation::ComputeEntrySize(const FileInfo* entry) const noexcept
{
    if (entry == nullptr)
    {
        return 0;
    }

    const size_t baseSize = offsetof(FileInfo, FileName);
    const size_t nameSize = static_cast<size_t>(entry->FileNameSize);
    const size_t total    = AlignUp(baseSize + nameSize + sizeof(wchar_t), kFileInfoEntryAlignment);
    return total;
}

HRESULT FilesInformation::LocateEntry(unsigned long index, FileInfo** ppEntry) const noexcept
{
    const std::byte* base      = _buffer.data();
    size_t offset              = 0;
    unsigned long currentIndex = 0;

    while (offset < _usedBytes && offset + sizeof(FileInfo) <= _buffer.size())
    {
        auto* entry = reinterpret_cast<FileInfo const*>(base + offset);

        if (currentIndex == index)
        {
            *ppEntry = const_cast<FileInfo*>(entry);
            return S_OK;
        }

        const size_t advance = (entry->NextEntryOffset != 0) ? static_cast<size_t>(entry->NextEntryOffset) : ComputeEntrySize(entry);
        if (advance == 0)
        {
            break;
        }

        offset += advance;
        ++currentIndex;
    }

    return HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES);
}

HRESULT STDMETHODCALLTYPE FileSystem::ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept
{
    // Reads ALL files and folders from the specified directory in a single call.
    // Uses progressive buffer growth (512KB → 2MB → 8MB → 32MB → ... → 512MB) to handle large directories.
    // The enumeration resumes when the buffer grows (no restart), avoiding O(N) re-enumeration passes.
    // If the directory exceeds 512MB, we may grow further (up to a hard cap) as a fallback.
    // Returns ERROR_INSUFFICIENT_BUFFER if the directory exceeds maximum capacity.

    if (ppFilesInformation == nullptr)
    {
        return E_POINTER;
    }

    *ppFilesInformation = nullptr;

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    auto infoImpl = std::unique_ptr<FilesInformation>(new (std::nothrow) FilesInformation());
    if (! infoImpl)
    {
        return E_OUTOFMEMORY;
    }

    std::wstring requestedPath(path);
    requestedPath = MakeAbsolutePath(requestedPath);
    if (requestedPath.empty())
    {
        requestedPath.assign(path);
    }

    unsigned long bytesWritten = 0;
    unsigned long entryCount   = 0;
    const HRESULT hr           = PopulateFilesInformation(*infoImpl, requestedPath, bytesWritten, entryCount);
    if (FAILED(hr))
    {
        return hr;
    }

    infoImpl->MaybeTrimBuffer();

    *ppFilesInformation = infoImpl.release();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::GetAttributes(const wchar_t* path, unsigned long* fileAttributes) noexcept
{
    if (fileAttributes == nullptr)
    {
        return E_POINTER;
    }

    *fileAttributes = 0;

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    // Allow navigation to UNC server roots (e.g. "\\\\server\\") which are treated as pseudo-directories listing shares.
    std::wstring serverName;
    if (TryGetUncServerRoot(path, serverName))
    {
        *fileAttributes = FILE_ATTRIBUTE_DIRECTORY;
        return S_OK;
    }

    const DWORD attrs = GetFileAttributesW(path);
    if (attrs == INVALID_FILE_ATTRIBUTES)
    {
        const DWORD lastError = GetLastError();
        return lastError != 0 ? HRESULT_FROM_WIN32(lastError) : E_FAIL;
    }

    *fileAttributes = attrs;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::CreateDirectory(const wchar_t* path) noexcept
{
    if (path == nullptr)
    {
        return E_POINTER;
    }

    if (path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (::CreateDirectoryW(path, nullptr) == 0)
    {
        const DWORD lastError = GetLastError();
        return lastError != 0 ? HRESULT_FROM_WIN32(lastError) : E_FAIL;
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::GetDirectorySize(
    const wchar_t* path, FileSystemFlags flags, IFileSystemDirectorySizeCallback* callback, void* cookie, FileSystemDirectorySizeResult* result) noexcept
{
    if (path == nullptr || result == nullptr)
    {
        return E_POINTER;
    }

    if (path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    result->totalBytes     = 0;
    result->fileCount      = 0;
    result->directoryCount = 0;
    result->status         = S_OK;

    // Verify path is a directory.
    const DWORD attrs = ::GetFileAttributesW(path);
    if (attrs == INVALID_FILE_ATTRIBUTES)
    {
        const DWORD lastError = ::GetLastError();
        result->status        = HRESULT_FROM_WIN32(lastError);
        return result->status;
    }

    if ((attrs & FILE_ATTRIBUTE_DIRECTORY) == 0)
    {
        WIN32_FILE_ATTRIBUTE_DATA fileData{};
        if (::GetFileAttributesExW(path, GetFileExInfoStandard, &fileData) == 0)
        {
            const DWORD lastError = ::GetLastError();
            result->status        = HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_GEN_FAILURE);
            return result->status;
        }

        result->totalBytes = (static_cast<unsigned __int64>(fileData.nFileSizeHigh) << 32) | fileData.nFileSizeLow;
        result->fileCount  = 1;

        if (callback != nullptr)
        {
            callback->DirectorySizeProgress(1, result->totalBytes, result->fileCount, result->directoryCount, path, cookie);
            BOOL cancel = FALSE;
            callback->DirectorySizeShouldCancel(&cancel, cookie);
            if (cancel)
            {
                result->status = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                return result->status;
            }

            callback->DirectorySizeProgress(1, result->totalBytes, result->fileCount, result->directoryCount, nullptr, cookie);
        }

        result->status = S_OK;
        return S_OK;
    }

    if ((attrs & FILE_ATTRIBUTE_REPARSE_POINT) != 0)
    {
        // Root reparse points are treated as leaf links for sizing to match safe copy/delete policy.
        result->status = S_OK;
        return S_OK;
    }

    const bool recursive                             = (flags & FILESYSTEM_FLAG_RECURSIVE) != 0;
    constexpr unsigned long kProgressIntervalEntries = 100;
    constexpr ULONGLONG kProgressIntervalMs          = 200;

    unsigned __int64 scannedEntries = 0;
    ULONGLONG lastProgressTime      = ::GetTickCount64();

#ifdef _DEBUG
    const unsigned int delayMs = _directorySizeDelayMs;
#else
    const unsigned int delayMs = 0u;
#endif

    auto maybeReportProgress = [&](const wchar_t* currentPath) -> bool
    {
        if (callback == nullptr)
        {
            return true;
        }

        const bool entryThreshold = (scannedEntries % kProgressIntervalEntries) == 0;
        const ULONGLONG now       = ::GetTickCount64();
        const bool timeThreshold  = (now - lastProgressTime) >= kProgressIntervalMs;

        if (entryThreshold || timeThreshold)
        {
            lastProgressTime = now;
            callback->DirectorySizeProgress(scannedEntries, result->totalBytes, result->fileCount, result->directoryCount, currentPath, cookie);

            BOOL cancel = FALSE;
            callback->DirectorySizeShouldCancel(&cancel, cookie);
            if (cancel)
            {
                result->status = HRESULT_FROM_WIN32(ERROR_CANCELLED);
                return false;
            }
        }
        return true;
    };

    struct DirectoryFrame final
    {
        DirectoryFrame()                                 = default;
        DirectoryFrame(const DirectoryFrame&)            = delete;
        DirectoryFrame(DirectoryFrame&&)                 = default;
        DirectoryFrame& operator=(const DirectoryFrame&) = delete;
        DirectoryFrame& operator=(DirectoryFrame&&)      = default;

        std::wstring directoryPath;
        wil::unique_hfind findHandle;
        WIN32_FIND_DATAW data{};
        bool hasData = false;
    };

    std::vector<DirectoryFrame> stack;

    auto pushDirectory = [&](std::wstring directoryPath) noexcept -> void
    {
        std::wstring searchPath = directoryPath;
        if (! searchPath.empty() && searchPath.back() != L'\\' && searchPath.back() != L'/')
        {
            searchPath += L'\\';
        }
        searchPath += L'*';

        WIN32_FIND_DATAW findData{};
        wil::unique_hfind findHandle(
            ::FindFirstFileExW(searchPath.c_str(), FindExInfoBasic, &findData, FindExSearchNameMatch, nullptr, FIND_FIRST_EX_LARGE_FETCH));
        if (! findHandle)
        {
            const DWORD lastError = ::GetLastError();
            if (lastError != ERROR_FILE_NOT_FOUND && lastError != ERROR_ACCESS_DENIED)
            {
                if (SUCCEEDED(result->status))
                {
                    result->status = HRESULT_FROM_WIN32(lastError);
                }
            }
            return;
        }

        DirectoryFrame frame{};
        frame.directoryPath = std::move(directoryPath);
        frame.findHandle    = std::move(findHandle);
        frame.data          = findData;
        frame.hasData       = true;

        stack.emplace_back(std::move(frame));
    };

    const auto advanceFrame = [&](DirectoryFrame& frame) noexcept
    {
        WIN32_FIND_DATAW next{};
        if (::FindNextFileW(frame.findHandle.get(), &next) != 0)
        {
            frame.data    = next;
            frame.hasData = true;
            return;
        }

        const DWORD lastError = ::GetLastError();
        if (lastError != ERROR_NO_MORE_FILES)
        {
            if (SUCCEEDED(result->status))
            {
                result->status = HRESULT_FROM_WIN32(lastError);
            }
        }

        frame.hasData = false;
    };

    pushDirectory(std::wstring(path));

    while (! stack.empty())
    {
        DirectoryFrame& frame = stack.back();
        if (! frame.hasData)
        {
            stack.pop_back();
            continue;
        }

        // Copy the current entry so we can safely mutate the frame (advance/push) without accidentally reusing stale data.
        const WIN32_FIND_DATAW currentData = frame.data;

        // Skip . and ..
        if (currentData.cFileName[0] == L'.')
        {
            if (currentData.cFileName[1] == L'\0')
            {
                advanceFrame(frame);
                continue;
            }
            if (currentData.cFileName[1] == L'.' && currentData.cFileName[2] == L'\0')
            {
                advanceFrame(frame);
                continue;
            }
        }

        ++scannedEntries;
        if (delayMs > 0u)
        {
            ::Sleep(delayMs);
        }

        const bool isDirectory    = (currentData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
        const bool isReparsePoint = (currentData.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;

        if (isDirectory)
        {
            ++result->directoryCount;
        }
        else
        {
            ++result->fileCount;
            const unsigned __int64 fileSize = (static_cast<unsigned __int64>(currentData.nFileSizeHigh) << 32) | currentData.nFileSizeLow;
            result->totalBytes += fileSize;
        }

        if (! maybeReportProgress(frame.directoryPath.c_str()))
        {
            return result->status;
        }

        if (recursive && isDirectory && ! isReparsePoint)
        {
            std::wstring childPath = frame.directoryPath;
            if (! childPath.empty() && childPath.back() != L'\\' && childPath.back() != L'/')
            {
                childPath += L'\\';
            }
            childPath += currentData.cFileName;

            // Advance the parent directory BEFORE descending; this keeps the parent frame consistent if the stack reallocates.
            advanceFrame(frame);

            pushDirectory(std::move(childPath));
            continue;
        }

        advanceFrame(frame);
    }

    // Final progress report.
    if (callback != nullptr)
    {
        callback->DirectorySizeProgress(scannedEntries, result->totalBytes, result->fileCount, result->directoryCount, nullptr, cookie);
    }

    return result->status;
}

void FileSystem::EnsureRequestedPath(FilesInformation& info, const std::wstring& path) noexcept
{
    if (! info.PathEquals(path))
    {
        info.ResetDirectoryState(true);
        info._requestedPath = path;
    }
}

HRESULT FileSystem::PopulateServerShares(FilesInformation& info, std::wstring_view serverName, unsigned long& bytesWritten, unsigned long& entryCount) noexcept
{
    bytesWritten = 0;
    entryCount   = 0;

    FileInfo* buffer = nullptr;
    HRESULT hr       = info.BeginWrite(&buffer);
    if (FAILED(hr))
    {
        return hr;
    }

    unsigned long capacityBytes = 0;
    hr                          = info.GetAllocatedSize(&capacityBytes);
    if (FAILED(hr))
    {
        return hr;
    }

    if (serverName.empty())
    {
        return E_INVALIDARG;
    }

    std::wstring serverPath;
    serverPath.reserve(serverName.size() + 2);
    serverPath.append(L"\\\\");
    serverPath.append(serverName);

    std::vector<std::wstring> shares;
    DWORD resumeHandle = 0;

    for (;;)
    {
        LPBYTE shareBufferRaw = nullptr;
        DWORD entriesRead     = 0;
        DWORD totalEntries    = 0;
        const NET_API_STATUS status =
            ::NetShareEnum(const_cast<wchar_t*>(serverPath.c_str()), 1, &shareBufferRaw, MAX_PREFERRED_LENGTH, &entriesRead, &totalEntries, &resumeHandle);

        wil::unique_any<LPBYTE, decltype(&::NetApiBufferFree), ::NetApiBufferFree> shareBuffer(shareBufferRaw);

        if (status != NERR_Success && status != ERROR_MORE_DATA)
        {
            return HRESULT_FROM_WIN32(status);
        }

        const auto* shareInfo = reinterpret_cast<const SHARE_INFO_1*>(shareBuffer.get());
        for (DWORD index = 0; index < entriesRead; ++index)
        {
            const auto& entry = shareInfo[index];
            if (! entry.shi1_netname || entry.shi1_netname[0] == L'\0')
            {
                continue;
            }

            const DWORD shareType = entry.shi1_type & STYPE_MASK;
            if (shareType != STYPE_DISKTREE)
            {
                continue;
            }

            shares.emplace_back(entry.shi1_netname);
        }

        if (status == NERR_Success)
        {
            break;
        }
    }

    if (shares.empty())
    {
        return S_OK;
    }

    std::sort(shares.begin(), shares.end(), [](const std::wstring& a, const std::wstring& b) { return OrdinalString::LessNoCase(a, b); });

    size_t requiredTotal = 0;
    for (const auto& share : shares)
    {
        const size_t nameChars = share.size();
        if (nameChars > (std::numeric_limits<unsigned long>::max)() / sizeof(wchar_t))
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        const size_t nameBytes  = nameChars * sizeof(wchar_t);
        const size_t entryBytes = AlignUp(offsetof(FileInfo, FileName) + nameBytes + sizeof(wchar_t), kFileInfoEntryAlignment);
        if (requiredTotal > static_cast<size_t>(capacityBytes) - entryBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_MORE_DATA);
        }

        requiredTotal += entryBytes;
    }

    std::byte* destination  = reinterpret_cast<std::byte*>(buffer);
    unsigned long remaining = capacityBytes;
    FileInfo* lastEntry     = nullptr;
    size_t lastEntrySize    = 0;

    for (const auto& share : shares)
    {
        const size_t nameChars = share.size();
        const size_t nameBytes = nameChars * sizeof(wchar_t);
        const size_t entrySize = AlignUp(offsetof(FileInfo, FileName) + nameBytes + sizeof(wchar_t), kFileInfoEntryAlignment);
        if (entrySize > remaining)
        {
            return HRESULT_FROM_WIN32(ERROR_MORE_DATA);
        }

        auto* entry = reinterpret_cast<FileInfo*>(destination);
        std::memset(entry, 0, entrySize);

        entry->FileAttributes  = FILE_ATTRIBUTE_DIRECTORY;
        entry->FileNameSize    = static_cast<unsigned long>(nameBytes);
        entry->EaSize          = 0;
        entry->FileIndex       = 0;
        entry->NextEntryOffset = 0;

        if (nameBytes > 0)
        {
            std::memcpy(entry->FileName, share.data(), nameBytes);
        }
        entry->FileName[nameChars] = L'\0';

        if (lastEntry != nullptr)
        {
            lastEntry->NextEntryOffset = static_cast<unsigned long>(lastEntrySize);
        }

        lastEntry     = entry;
        lastEntrySize = entrySize;

        destination += entrySize;
        remaining -= static_cast<unsigned long>(entrySize);
        bytesWritten += static_cast<unsigned long>(entrySize);
        ++entryCount;
    }

    if (lastEntry != nullptr)
    {
        lastEntry->NextEntryOffset = 0;
    }

    return S_OK;
}

HRESULT FileSystem::PopulateFilesInformation(FilesInformation& info, const std::wstring& path, unsigned long& bytesWritten, unsigned long& entryCount) noexcept
{
    Debug::Perf::Scope perf(L"FileSystem.DirectoryOps.Enumerate");
    perf.SetDetail(path);

    EnsureRequestedPath(info, path);

    std::wstring serverName;
    const bool isUncServerRoot = TryGetUncServerRoot(path, serverName);

    // Progressive buffer growth (single enumeration pass): 512KB → 2MB → 8MB → 32MB → ... → soft cap.
    // Fallback: if the directory exceeds the soft cap, we may grow further (up to a hard cap) to avoid hard failure in extreme cases.
    // The enumeration resumes when the buffer grows (no restart), so we avoid O(N) re-enumeration passes.
    constexpr size_t kGrowthFactor            = 4u;
    constexpr size_t kMiB                     = 1024u * 1024u;
    const size_t maxBufferBytesLimit          = static_cast<size_t>((std::numeric_limits<unsigned long>::max)());
    unsigned long enumerationSoftMaxBufferMiB = kDefaultEnumerationSoftMaxBufferMiB;
    unsigned long enumerationHardMaxBufferMiB = kDefaultEnumerationHardMaxBufferMiB;
    {
        std::lock_guard lock(_stateMutex);
        enumerationSoftMaxBufferMiB = _enumerationSoftMaxBufferMiB;
        enumerationHardMaxBufferMiB = _enumerationHardMaxBufferMiB;
    }
    const size_t softMaxBufferSize =
        std::clamp(static_cast<size_t>(enumerationSoftMaxBufferMiB) * kMiB, static_cast<size_t>(kDefaultBufferSize), maxBufferBytesLimit);
    const size_t hardMaxBufferSize = std::clamp(static_cast<size_t>(enumerationHardMaxBufferMiB) * kMiB, softMaxBufferSize, maxBufferBytesLimit);
    const bool canFallback         = hardMaxBufferSize > softMaxBufferSize;

    uint64_t growCount       = 0;
    size_t currentBufferSize = info._buffer.size();
    size_t peakBufferSize    = currentBufferSize;
    size_t maxBufferSize     = softMaxBufferSize;
    bool usedFallback        = false;
    if (currentBufferSize < kDefaultBufferSize)
    {
        currentBufferSize = kDefaultBufferSize;
        info.ResizeBuffer(currentBufferSize);
        peakBufferSize = currentBufferSize;
    }

    auto finalizePerf = [&](HRESULT hr) noexcept
    {
        perf.SetValue0(static_cast<uint64_t>(peakBufferSize));
        perf.SetValue1(growCount);

        if (FAILED(hr))
        {
            perf.SetHr(hr);
            return;
        }

        perf.SetHr(usedFallback ? S_FALSE : S_OK);
    };

    if (isUncServerRoot)
    {
        for (;;)
        {
            HRESULT hr = PopulateServerShares(info, serverName, bytesWritten, entryCount);

            // Share enumeration doesn't use ERROR_NO_MORE_FILES; empty server just yields an empty result set.

            constexpr HRESULT errorMoreData      = HRESULT_FROM_WIN32(ERROR_MORE_DATA);
            constexpr HRESULT insufficientBuffer = HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);

            if (hr == errorMoreData || hr == insufficientBuffer)
            {
                if (currentBufferSize >= maxBufferSize)
                {
                    if (! usedFallback && canFallback)
                    {
                        maxBufferSize = hardMaxBufferSize;
                        usedFallback  = true;
                    }
                    else
                    {
                        const HRESULT finalHr = HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
                        finalizePerf(finalHr);
                        return finalHr;
                    }
                }

                ++growCount;
                const size_t grown = currentBufferSize * kGrowthFactor;
                currentBufferSize  = (grown < currentBufferSize || grown > maxBufferSize) ? maxBufferSize : grown;
                peakBufferSize     = std::max(peakBufferSize, currentBufferSize);
                info.ResizeBuffer(currentBufferSize);
                continue;
            }

            if (SUCCEEDED(hr))
            {
                info.UpdateUsage(bytesWritten, entryCount);
                finalizePerf(S_OK);
                return S_OK;
            }

            finalizePerf(hr);
            return hr;
        }
    }

    bytesWritten         = 0;
    entryCount           = 0;
    size_t lastEntrySize = 0;

    FileInfo* buffer = nullptr;
    HRESULT hr       = info.BeginWrite(&buffer);
    if (FAILED(hr))
    {
        finalizePerf(hr);
        return hr;
    }

    for (;;)
    {
        hr = EnsureEnumeration(info, path);
        if (hr == HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES))
        {
            info.UpdateUsage(bytesWritten, entryCount);
            finalizePerf(S_OK);
            return S_OK;
        }
        if (FAILED(hr))
        {
            finalizePerf(hr);
            return hr;
        }

        hr = PopulateBuffer(info, bytesWritten, entryCount, lastEntrySize);

        // Completed directory listing.
        if (SUCCEEDED(hr))
        {
            info.UpdateUsage(bytesWritten, entryCount);
            finalizePerf(S_OK);
            return S_OK;
        }

        constexpr HRESULT errorMoreData      = HRESULT_FROM_WIN32(ERROR_MORE_DATA);
        constexpr HRESULT insufficientBuffer = HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);

        if (hr == errorMoreData || hr == insufficientBuffer)
        {
            if (currentBufferSize >= maxBufferSize)
            {
                if (! usedFallback && canFallback)
                {
                    maxBufferSize = hardMaxBufferSize;
                    usedFallback  = true;
                }
                else
                {
                    const HRESULT finalHr = HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
                    finalizePerf(finalHr);
                    return finalHr;
                }
            }

            ++growCount;
            const size_t grown = currentBufferSize * kGrowthFactor;
            currentBufferSize  = (grown < currentBufferSize || grown > maxBufferSize) ? maxBufferSize : grown;
            peakBufferSize     = std::max(peakBufferSize, currentBufferSize);
            info.ResizeBuffer(currentBufferSize);
            continue;
        }

        finalizePerf(hr);
        return hr;
    }
}

HRESULT FileSystem::EnsureEnumeration(FilesInformation& info, const std::wstring& path) noexcept
{
    if (info._enumerationComplete)
    {
        return HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES);
    }

    if (! info._enumerationInitialized)
    {
        return StartEnumeration(info, path);
    }

    return S_OK;
}

HRESULT FileSystem::StartEnumeration(FilesInformation& info, const std::wstring& path) noexcept
{
    const std::wstring extendedPath = ToExtendedPath(path);
    if (ShouldUseHandleEnumeration(extendedPath))
    {
        const HRESULT hr = StartEnumerationHandle(info, extendedPath);
        if (SUCCEEDED(hr))
        {
            return hr;
        }

        // Graceful fallback for network paths and edge cases.
        info.ResetDirectoryState(false);
    }

    return StartEnumerationWin32(info, path);
}

HRESULT FileSystem::StartEnumerationWin32(FilesInformation& info, const std::wstring& path) noexcept
{
    std::wstring extendedPath = ToExtendedPath(path);
    if (! extendedPath.empty())
    {
        const wchar_t last = extendedPath.back();
        if (last != L'\\' && last != L'/')
        {
            extendedPath.push_back(L'\\');
        }
    }
    extendedPath.append(L"*");

    WIN32_FIND_DATAW findData{};
    wil::unique_hfind findHandle(FindFirstFileExW(extendedPath.c_str(), FindExInfoBasic, &findData, FindExSearchNameMatch, nullptr, FIND_FIRST_EX_LARGE_FETCH));
    if (! findHandle)
    {
        const DWORD err = GetLastError();
        return HRESULT_FROM_WIN32(err);
    }

    info._findHandle             = std::move(findHandle);
    info._pendingEntry           = findData;
    info._hasPendingEntry        = true;
    info._enumerationInitialized = true;
    info._enumerationComplete    = false;
    return S_OK;
}

HRESULT FileSystem::StartEnumerationHandle(FilesInformation& info, const std::wstring& path) noexcept
{
    std::wstring extendedPath = ToExtendedPath(path);
    wil::unique_handle directory(CreateFileW(extendedPath.c_str(),
                                             FILE_LIST_DIRECTORY | SYNCHRONIZE,
                                             FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                             nullptr,
                                             OPEN_EXISTING,
                                             FILE_FLAG_BACKUP_SEMANTICS,
                                             nullptr));
    if (! directory)
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    // Ensure scratch buffer exists. The API doesn't report "bytes returned"; use the linked entry offsets for traversal.
    constexpr size_t kEnumerationBufferBytes = 512u * 1024u;
    if (info._enumerationBuffer.size() != kEnumerationBufferBytes)
    {
        info._enumerationBuffer.assign(kEnumerationBufferBytes, std::byte{0});
    }

    info._directoryHandle             = std::move(directory);
    info._useHandleEnumeration        = true;
    info._enumerationRestartScan      = true;
    info._enumerationBufferOffset     = 0;
    info._enumerationBufferBytesValid = 0;
    info._enumerationInitialized      = true;
    info._enumerationComplete         = false;
    return S_OK;
}

HRESULT FileSystem::FetchNextEntryWin32(FilesInformation& info, WIN32_FIND_DATAW& data) noexcept
{
    while (true)
    {
        if (info._hasPendingEntry)
        {
            data                  = info._pendingEntry;
            info._hasPendingEntry = false;
        }
        else
        {
            if (! info._findHandle)
            {
                info._enumerationComplete = true;
                return HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES);
            }

            if (! FindNextFileW(info._findHandle.get(), &data))
            {
                const DWORD err = GetLastError();
                if (err == ERROR_NO_MORE_FILES)
                {
                    info._findHandle.reset();
                    info._enumerationInitialized = false;
                    info._enumerationComplete    = true;
                    return HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES);
                }

                info.ResetDirectoryState(false);
                return HRESULT_FROM_WIN32(err);
            }
        }

        if (! IsDotOrDotDot(std::wstring_view(data.cFileName)))
        {
            return S_OK;
        }
    }
}

HRESULT FileSystem::PopulateBuffer(FilesInformation& info, unsigned long& bytesWritten, unsigned long& entryCount, size_t& lastEntrySize) noexcept
{
    if (info._useHandleEnumeration)
    {
        return PopulateBufferHandle(info, bytesWritten, entryCount, lastEntrySize);
    }

    return PopulateBufferWin32(info, bytesWritten, entryCount, lastEntrySize);
}

HRESULT FileSystem::PopulateBufferWin32(FilesInformation& info, unsigned long& bytesWritten, unsigned long& entryCount, size_t& lastEntrySize) noexcept
{
    if (info._buffer.size() > std::numeric_limits<unsigned long>::max())
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    const unsigned long capacityBytes = static_cast<unsigned long>(info._buffer.size());

    if (bytesWritten > capacityBytes)
    {
        return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
    }

    if ((entryCount == 0) != (bytesWritten == 0))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    if (entryCount > 0 && (lastEntrySize == 0 || bytesWritten < lastEntrySize))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    std::byte* base = info._buffer.data();

    for (;;)
    {
        const unsigned long remaining = capacityBytes - bytesWritten;

        WIN32_FIND_DATAW findData{};
        HRESULT hr = FetchNextEntryWin32(info, findData);
        if (hr == HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES))
        {
            return S_OK;
        }
        if (FAILED(hr))
        {
            return hr;
        }

        const size_t nameLength   = wcslen(findData.cFileName);
        const size_t nameBytes    = nameLength * sizeof(wchar_t);
        const size_t requiredSize = AlignUp(offsetof(FileInfo, FileName) + nameBytes + sizeof(wchar_t), kFileInfoEntryAlignment);

        if (requiredSize > remaining)
        {
            info._pendingEntry    = findData;
            info._hasPendingEntry = true;

            if (entryCount == 0)
            {
                return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
            }

            return HRESULT_FROM_WIN32(ERROR_MORE_DATA);
        }

        if (entryCount > 0)
        {
            auto* prev            = reinterpret_cast<FileInfo*>(base + (static_cast<size_t>(bytesWritten) - lastEntrySize));
            prev->NextEntryOffset = static_cast<unsigned long>(lastEntrySize);
        }

        auto* entry = reinterpret_cast<FileInfo*>(base + bytesWritten);

        entry->FileNameSize = static_cast<unsigned long>(nameBytes);
        if (nameBytes > 0)
        {
            std::memcpy(entry->FileName, findData.cFileName, nameBytes);
        }
        entry->FileName[nameLength] = L'\0';

        entry->FileAttributes = findData.dwFileAttributes;
        entry->CreationTime   = FileTimeToInt64(findData.ftCreationTime);
        entry->LastAccessTime = FileTimeToInt64(findData.ftLastAccessTime);
        entry->LastWriteTime  = FileTimeToInt64(findData.ftLastWriteTime);
        entry->ChangeTime     = entry->LastWriteTime;

        const unsigned long long fileSize =
            (static_cast<unsigned long long>(findData.nFileSizeHigh) << 32) | static_cast<unsigned long long>(findData.nFileSizeLow);
        entry->EndOfFile       = static_cast<__int64>(fileSize);
        entry->AllocationSize  = entry->EndOfFile;
        entry->EaSize          = 0;
        entry->FileIndex       = 0;
        entry->NextEntryOffset = 0;

        bytesWritten += static_cast<unsigned long>(requiredSize);
        ++entryCount;
        lastEntrySize = requiredSize;
    }
}

HRESULT FileSystem::PopulateBufferHandle(FilesInformation& info, unsigned long& bytesWritten, unsigned long& entryCount, size_t& lastEntrySize) noexcept
{
    if (info._buffer.size() > std::numeric_limits<unsigned long>::max())
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    const unsigned long capacityBytes = static_cast<unsigned long>(info._buffer.size());

    if (bytesWritten > capacityBytes)
    {
        return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
    }

    if ((entryCount == 0) != (bytesWritten == 0))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    if (entryCount > 0 && (lastEntrySize == 0 || bytesWritten < lastEntrySize))
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    if (! info._directoryHandle || info._enumerationBuffer.empty())
    {
        return S_OK;
    }

    if (info._enumerationBuffer.size() > static_cast<size_t>(std::numeric_limits<DWORD>::max()))
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    std::byte* base = info._buffer.data();

    for (;;)
    {
        const unsigned long remaining = capacityBytes - bytesWritten;

        // Refresh batch buffer if consumed.
        if (info._enumerationBufferOffset >= info._enumerationBufferBytesValid)
        {
            if (const auto NtQueryDirectoryFile = GetNtQueryDirectoryFile(); NtQueryDirectoryFile != nullptr)
            {
                IO_STATUS_BLOCK iosb{};
                const BOOLEAN restart = info._enumerationRestartScan ? TRUE : FALSE;
                const NTSTATUS status = NtQueryDirectoryFile(info._directoryHandle.get(),
                                                             nullptr,
                                                             nullptr,
                                                             nullptr,
                                                             &iosb,
                                                             info._enumerationBuffer.data(),
                                                             static_cast<ULONG>(info._enumerationBuffer.size()),
                                                             NtFileInformationClass::FileFullDirectoryInformation,
                                                             FALSE,
                                                             nullptr,
                                                             restart);
                if (status == STATUS_NO_MORE_FILES)
                {
                    info._enumerationComplete    = true;
                    info._enumerationInitialized = false;
                    info._useHandleEnumeration   = false;
                    info._directoryHandle.reset();
                    return S_OK;
                }

                if (! NT_SUCCESS(status))
                {
                    if (const auto RtlNtStatusToDosError = GetRtlNtStatusToDosError(); RtlNtStatusToDosError != nullptr)
                    {
                        const DWORD error = RtlNtStatusToDosError(status);
                        return HRESULT_FROM_WIN32(error != 0 ? error : ERROR_GEN_FAILURE);
                    }

                    return HRESULT_FROM_WIN32(ERROR_GEN_FAILURE);
                }

                const size_t bytesValid = static_cast<size_t>(iosb.Information);
                if (bytesValid == 0 || bytesValid > info._enumerationBuffer.size())
                {
                    return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
                }

                info._enumerationRestartScan      = false;
                info._enumerationBufferOffset     = 0;
                info._enumerationBufferBytesValid = bytesValid;
            }
            else
            {
                const FILE_INFO_BY_HANDLE_CLASS cls = info._enumerationRestartScan ? FileFullDirectoryRestartInfo : FileFullDirectoryInfo;
                if (! ::GetFileInformationByHandleEx(
                        info._directoryHandle.get(), cls, info._enumerationBuffer.data(), static_cast<DWORD>(info._enumerationBuffer.size())))
                {
                    const DWORD err = ::GetLastError();
                    if (err == ERROR_NO_MORE_FILES)
                    {
                        info._enumerationComplete    = true;
                        info._enumerationInitialized = false;
                        info._useHandleEnumeration   = false;
                        info._directoryHandle.reset();
                        return S_OK;
                    }

                    return HRESULT_FROM_WIN32(err);
                }

                info._enumerationRestartScan      = false;
                info._enumerationBufferOffset     = 0;
                info._enumerationBufferBytesValid = info._enumerationBuffer.size();
            }
        }

        const size_t sourceOffset = info._enumerationBufferOffset;
        if ((sourceOffset % kFileInfoEntryAlignment) != 0u)
        {
            return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
        }
        if (info._enumerationBufferBytesValid <= sourceOffset || info._enumerationBufferBytesValid - sourceOffset < offsetof(FILE_FULL_DIR_INFO, FileName))
        {
            return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
        }

        const auto* source = reinterpret_cast<const FILE_FULL_DIR_INFO*>(info._enumerationBuffer.data() + sourceOffset);

        size_t nextOffset = info._enumerationBufferBytesValid;
        if (source->NextEntryOffset != 0)
        {
            if ((static_cast<size_t>(source->NextEntryOffset) % kFileInfoEntryAlignment) != 0u)
            {
                return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
            }

            if (static_cast<size_t>(source->NextEntryOffset) > info._enumerationBufferBytesValid - sourceOffset)
            {
                return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
            }

            nextOffset = sourceOffset + static_cast<size_t>(source->NextEntryOffset);
        }

        const unsigned long nameBytes = source->FileNameLength;
        if ((nameBytes % static_cast<unsigned long>(sizeof(wchar_t))) != 0u)
        {
            info._enumerationBufferOffset = nextOffset;
            continue;
        }

        const size_t nameChars  = static_cast<size_t>(nameBytes / sizeof(wchar_t));
        const size_t nameOffset = offsetof(FILE_FULL_DIR_INFO, FileName);
        if (nameOffset + static_cast<size_t>(nameBytes) > info._enumerationBufferBytesValid - sourceOffset)
        {
            return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
        }

        const std::wstring_view name(source->FileName, nameChars);
        if (IsDotOrDotDot(name))
        {
            info._enumerationBufferOffset = nextOffset;
            continue;
        }

        const size_t requiredSize = AlignUp(offsetof(FileInfo, FileName) + static_cast<size_t>(nameBytes) + sizeof(wchar_t), kFileInfoEntryAlignment);
        if (requiredSize > remaining)
        {
            if (entryCount == 0)
            {
                return HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER);
            }

            return HRESULT_FROM_WIN32(ERROR_MORE_DATA);
        }

        if (entryCount > 0)
        {
            auto* prev            = reinterpret_cast<FileInfo*>(base + (static_cast<size_t>(bytesWritten) - lastEntrySize));
            prev->NextEntryOffset = static_cast<unsigned long>(lastEntrySize);
        }

        auto* entry = reinterpret_cast<FileInfo*>(base + bytesWritten);
        std::memcpy(entry, source, offsetof(FileInfo, FileName));
        entry->NextEntryOffset = 0;

        entry->FileNameSize = nameBytes;
        if (nameBytes > 0)
        {
            std::memcpy(entry->FileName, source->FileName, static_cast<size_t>(nameBytes));
        }
        entry->FileName[nameChars] = L'\0';

        bytesWritten += static_cast<unsigned long>(requiredSize);
        ++entryCount;
        lastEntrySize                 = requiredSize;
        info._enumerationBufferOffset = nextOffset;
    }
}
