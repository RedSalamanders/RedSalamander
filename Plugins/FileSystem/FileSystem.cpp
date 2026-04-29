#include "FileSystem.Internal.h"

#include <cwctype>
#include <limits>

#include <yyjson.h>

#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Netapi32.lib")

namespace
{
[[nodiscard]] std::string Utf8FromUtf16(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    const int len = WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
    if (len <= 0)
    {
        return {};
    }

    std::string result;
    result.resize(static_cast<size_t>(len));
    const int written = WideCharToMultiByte(CP_UTF8, 0, text.data(), static_cast<int>(text.size()), result.data(), len, nullptr, nullptr);
    if (written != len)
    {
        return {};
    }

    return result;
}

[[nodiscard]] std::string FormatFileTimeLocal(const FILETIME& fileTime) noexcept
{
    if (fileTime.dwLowDateTime == 0u && fileTime.dwHighDateTime == 0u)
    {
        return {};
    }

    FILETIME localFileTime{};
    if (FileTimeToLocalFileTime(&fileTime, &localFileTime) == 0)
    {
        return {};
    }

    SYSTEMTIME localSystemTime{};
    if (FileTimeToSystemTime(&localFileTime, &localSystemTime) == 0)
    {
        return {};
    }

    return Utf8FromUtf16(std::format(L"{:04}-{:02}-{:02} {:02}:{:02}:{:02}",
                                     localSystemTime.wYear,
                                     localSystemTime.wMonth,
                                     localSystemTime.wDay,
                                     localSystemTime.wHour,
                                     localSystemTime.wMinute,
                                     localSystemTime.wSecond));
}

[[nodiscard]] std::string FormatFileAttributeFlags(DWORD attributes) noexcept
{
    std::wstring text;
    const auto appendFlag = [&](const DWORD flag, const std::wstring_view label) noexcept
    {
        if ((attributes & flag) == 0u)
        {
            return;
        }

        if (! text.empty())
        {
            text.append(L", ");
        }
        text.append(label);
    };

    appendFlag(FILE_ATTRIBUTE_ARCHIVE, L"Archive");
    appendFlag(FILE_ATTRIBUTE_COMPRESSED, L"Compressed");
    appendFlag(FILE_ATTRIBUTE_DIRECTORY, L"Directory");
    appendFlag(FILE_ATTRIBUTE_ENCRYPTED, L"Encrypted");
    appendFlag(FILE_ATTRIBUTE_HIDDEN, L"Hidden");
    appendFlag(FILE_ATTRIBUTE_NOT_CONTENT_INDEXED, L"Not content indexed");
    appendFlag(FILE_ATTRIBUTE_OFFLINE, L"Offline");
    appendFlag(FILE_ATTRIBUTE_READONLY, L"Read-only");
    appendFlag(FILE_ATTRIBUTE_REPARSE_POINT, L"Reparse point");
    appendFlag(FILE_ATTRIBUTE_SPARSE_FILE, L"Sparse");
    appendFlag(FILE_ATTRIBUTE_SYSTEM, L"System");
    appendFlag(FILE_ATTRIBUTE_TEMPORARY, L"Temporary");

    if (text.empty())
    {
        text = L"None";
    }

    return Utf8FromUtf16(text);
}

[[nodiscard]] FileSystemReparsePointPolicy ParseReparsePointPolicy(std::string_view policy) noexcept
{
    if (policy == "copyReparse")
    {
        return FileSystemReparsePointPolicy::CopyReparse;
    }
    if (policy == "followTargets")
    {
        return FileSystemReparsePointPolicy::FollowTargets;
    }
    if (policy == "skip")
    {
        return FileSystemReparsePointPolicy::Skip;
    }

    return FileSystemReparsePointPolicy::CopyReparse;
}

[[nodiscard]] const char* ReparsePointPolicyToString(FileSystemReparsePointPolicy policy) noexcept
{
    switch (policy)
    {
        case FileSystemReparsePointPolicy::CopyReparse: return "copyReparse";
        case FileSystemReparsePointPolicy::FollowTargets: return "followTargets";
        case FileSystemReparsePointPolicy::Skip: return "skip";
    }

    return "copyReparse";
}

[[nodiscard]] FileSystemSearchBackendPreference ParseSearchBackendPreference(std::string_view preference) noexcept
{
    if (preference == "service")
    {
        return FileSystemSearchBackendPreference::Service;
    }
    if (preference == "local-index")
    {
        return FileSystemSearchBackendPreference::LocalIndex;
    }
    if (preference == "scan")
    {
        return FileSystemSearchBackendPreference::Scan;
    }

    return FileSystemSearchBackendPreference::Auto;
}

[[nodiscard]] const char* SearchBackendPreferenceToString(FileSystemSearchBackendPreference preference) noexcept
{
    switch (preference)
    {
        case FileSystemSearchBackendPreference::Auto: return "auto";
        case FileSystemSearchBackendPreference::Service: return "service";
        case FileSystemSearchBackendPreference::LocalIndex: return "local-index";
        case FileSystemSearchBackendPreference::Scan: return "scan";
    }

    return "auto";
}

[[nodiscard]] FileSystemConcurrencyMode ParseConcurrencyMode(std::string_view mode) noexcept
{
    if (mode == "manual")
    {
        return FileSystemConcurrencyMode::Manual;
    }

    return FileSystemConcurrencyMode::Auto;
}

[[nodiscard]] const char* ConcurrencyModeToString(FileSystemConcurrencyMode mode) noexcept
{
    switch (mode)
    {
        case FileSystemConcurrencyMode::Auto: return "auto";
        case FileSystemConcurrencyMode::Manual: return "manual";
    }

    return "auto";
}

[[nodiscard]] std::wstring_view StripExtendedPrefix(std::wstring_view path) noexcept
{
    if (path.rfind(L"\\\\?\\UNC\\", 0) == 0)
    {
        return path.substr(6);
    }
    if (path.rfind(L"\\\\?\\", 0) == 0)
    {
        return path.substr(4);
    }
    return path;
}

[[nodiscard]] bool IsUncPath(std::wstring_view path) noexcept
{
    const std::wstring_view normalized = StripExtendedPrefix(path);
    return normalized.size() >= 2 && normalized[0] == L'\\' && normalized[1] == L'\\';
}

[[nodiscard]] std::wstring ExtractDriveRoot(std::wstring_view path) noexcept
{
    const std::wstring_view normalized = StripExtendedPrefix(path);
    if (normalized.size() >= 3 && normalized[1] == L':' && (normalized[2] == L'\\' || normalized[2] == L'/'))
    {
        return std::format(L"{}:\\", static_cast<wchar_t>(towupper(normalized[0])));
    }

    if (! IsUncPath(path))
    {
        return {};
    }

    size_t serverEnd = normalized.find_first_of(L"\\/", 2);
    if (serverEnd == std::wstring_view::npos)
    {
        return {};
    }

    size_t shareEnd = normalized.find_first_of(L"\\/", serverEnd + 1);
    if (shareEnd == std::wstring_view::npos)
    {
        return std::wstring(normalized) + L"\\";
    }

    return std::wstring(normalized.substr(0, shareEnd + 1));
}

[[nodiscard]] unsigned long ClampPreferredBridgeBufferBytes(uint32_t preferredBytes) noexcept
{
    constexpr uint64_t kMinBytes = 512ull * 1024ull;
    constexpr uint64_t kMaxBytes = 16ull * 1024ull * 1024ull;
    const uint64_t clampedBytes  = (std::clamp)(static_cast<uint64_t>(preferredBytes), kMinBytes, kMaxBytes);
    return static_cast<unsigned long>((std::min)(clampedBytes, static_cast<uint64_t>(std::numeric_limits<unsigned long>::max())));
}

void FillTransferHintsLocal(FileSystemTransferHints& hints, bool highLatency) noexcept
{
    hints.latencyClass              = highLatency ? FILESYSTEM_TRANSFER_LATENCY_LAN : FILESYSTEM_TRANSFER_LATENCY_LOCAL;
    hints.flags                     = FILESYSTEM_TRANSFER_HINT_PREFERS_SEQUENTIAL_IO;
    hints.preferredBufferBytes      = highLatency ? (8u * 1024u * 1024u) : (2u * 1024u * 1024u);
    hints.preferredProgressPeriodMs = 200u;
    if (highLatency)
    {
        hints.flags |= FILESYSTEM_TRANSFER_HINT_PREFERS_LARGE_BUFFERS;
    }
}

void FillStorageCharacteristicsLocal(FileSystemStorageCharacteristics& characteristics, bool highLatency) noexcept
{
    characteristics.storageKind                  = highLatency ? FILESYSTEM_STORAGE_NETWORK_SHARE : FILESYSTEM_STORAGE_UNKNOWN;
    characteristics.flags                        = FILESYSTEM_STORAGE_FLAG_PREFERS_SEQUENTIAL_IO;
    characteristics.queueDepthHint               = highLatency ? 8u : 4u;
    characteristics.preferredCopyMoveConcurrency = highLatency ? 8u : 4u;
    characteristics.preferredDeleteConcurrency   = 8u;
    if (highLatency)
    {
        characteristics.flags |= FILESYSTEM_STORAGE_FLAG_HIGH_LATENCY | FILESYSTEM_STORAGE_FLAG_SUPPORTS_DEEP_QUEUE;
    }
}

[[nodiscard]] std::string BuildConfigurationJson(FileSystemConcurrencyMode concurrencyMode,
                                                 unsigned int copyMoveMaxConcurrency,
                                                 unsigned int deleteMaxConcurrency,
                                                 unsigned int deleteRecycleBinMaxConcurrency,
                                                 unsigned int recycleBinBatchSize,
                                                 unsigned long enumerationSoftMaxBufferMiB,
                                                 unsigned long enumerationHardMaxBufferMiB,
                                                 FileSystemReparsePointPolicy reparsePointPolicy,
                                                 FileSystemSearchBackendPreference searchBackendPreference,
                                                 unsigned int searchMaxDirectoryWalkers) noexcept
{
    return std::format("{{\"concurrencyMode\":\"{}\",\"copyMoveMaxConcurrency\":{},\"deleteMaxConcurrency\":{},\"deleteRecycleBinMaxConcurrency\":{},"
                       "\"recycleBinBatchSize\":{},\"enumerationSoftMaxBufferMiB\":{},\"enumerationHardMaxBufferMiB\":{},\"reparsePointPolicy\":\"{}\","
                       "\"searchBackendPreference\":\"{}\",\"searchMaxDirectoryWalkers\":{}}}",
                       ConcurrencyModeToString(concurrencyMode),
                       copyMoveMaxConcurrency,
                       deleteMaxConcurrency,
                       deleteRecycleBinMaxConcurrency,
                       recycleBinBatchSize,
                       enumerationSoftMaxBufferMiB,
                       enumerationHardMaxBufferMiB,
                       ReparsePointPolicyToString(reparsePointPolicy),
                       SearchBackendPreferenceToString(searchBackendPreference),
                       searchMaxDirectoryWalkers);
}

class Win32FileReader final : public IFileReader
{
public:
    Win32FileReader(wil::unique_handle file, uint64_t sizeBytes) noexcept : _file(std::move(file)), _sizeBytes(sizeBytes)
    {
    }

    Win32FileReader(const Win32FileReader&)            = delete;
    Win32FileReader(Win32FileReader&&)                 = delete;
    Win32FileReader& operator=(const Win32FileReader&) = delete;
    Win32FileReader& operator=(Win32FileReader&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileReader))
        {
            *ppvObject = static_cast<IFileReader*>(this);
            AddRef();
            return S_OK;
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    HRESULT STDMETHODCALLTYPE GetSize(uint64_t* sizeBytes) noexcept override
    {
        if (sizeBytes == nullptr)
        {
            return E_POINTER;
        }

        *sizeBytes = _sizeBytes;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Seek(__int64 offset, unsigned long origin, uint64_t* newPosition) noexcept override
    {
        if (newPosition == nullptr)
        {
            return E_POINTER;
        }

        *newPosition = 0;

        if (! _file)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        }

        if (origin != FILE_BEGIN && origin != FILE_CURRENT && origin != FILE_END)
        {
            return E_INVALIDARG;
        }

        LARGE_INTEGER distance{};
        distance.QuadPart = offset;

        LARGE_INTEGER moved{};
        if (SetFilePointerEx(_file.get(), distance, &moved, origin) == 0)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        if (moved.QuadPart < 0)
        {
            return HRESULT_FROM_WIN32(ERROR_NEGATIVE_SEEK);
        }

        *newPosition = static_cast<uint64_t>(moved.QuadPart);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Read(void* buffer, unsigned long bytesToRead, unsigned long* bytesRead) noexcept override
    {
        if (bytesRead == nullptr)
        {
            return E_POINTER;
        }

        *bytesRead = 0;

        if (bytesToRead == 0)
        {
            return S_OK;
        }

        if (buffer == nullptr)
        {
            return E_POINTER;
        }

        if (! _file)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        }

        DWORD read = 0;
        if (ReadFile(_file.get(), buffer, bytesToRead, &read, nullptr) == 0)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        *bytesRead = static_cast<unsigned long>(read);
        return S_OK;
    }

private:
    ~Win32FileReader() = default;

    std::atomic_ulong _refCount{1};
    wil::unique_handle _file;
    uint64_t _sizeBytes = 0;
};

class Win32FileWriter final : public IFileWriter
{
public:
    Win32FileWriter(wil::unique_handle file, std::wstring path) noexcept : _file(std::move(file)), _path(std::move(path))
    {
    }

    Win32FileWriter(const Win32FileWriter&)            = delete;
    Win32FileWriter(Win32FileWriter&&)                 = delete;
    Win32FileWriter& operator=(const Win32FileWriter&) = delete;
    Win32FileWriter& operator=(Win32FileWriter&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileWriter))
        {
            *ppvObject = static_cast<IFileWriter*>(this);
            AddRef();
            return S_OK;
        }

        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG current = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (current == 0)
        {
            delete this;
        }
        return current;
    }

    HRESULT STDMETHODCALLTYPE GetPosition(uint64_t* positionBytes) noexcept override
    {
        if (positionBytes == nullptr)
        {
            return E_POINTER;
        }

        *positionBytes = 0;

        if (! _file)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        }

        LARGE_INTEGER distance{};
        LARGE_INTEGER moved{};
        if (SetFilePointerEx(_file.get(), distance, &moved, FILE_CURRENT) == 0)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        if (moved.QuadPart < 0)
        {
            return HRESULT_FROM_WIN32(ERROR_NEGATIVE_SEEK);
        }

        *positionBytes = static_cast<uint64_t>(moved.QuadPart);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Write(const void* buffer, unsigned long bytesToWrite, unsigned long* bytesWritten) noexcept override
    {
        if (bytesWritten == nullptr)
        {
            return E_POINTER;
        }

        *bytesWritten = 0;

        if (bytesToWrite == 0)
        {
            return S_OK;
        }

        if (buffer == nullptr)
        {
            return E_POINTER;
        }

        if (! _file)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        }

        DWORD wrote = 0;
        if (WriteFile(_file.get(), buffer, bytesToWrite, &wrote, nullptr) == 0)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        *bytesWritten = static_cast<unsigned long>(wrote);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Commit() noexcept override
    {
        if (! _file)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        }

        if (FlushFileBuffers(_file.get()) == 0)
        {
            return HRESULT_FROM_WIN32(GetLastError());
        }

        _committed = true;
        return S_OK;
    }

private:
    ~Win32FileWriter()
    {
        if (_committed)
        {
            return;
        }

        _file.reset();
        if (! _path.empty())
        {
            static_cast<void>(DeleteFileW(_path.c_str()));
        }
    }

    std::atomic_ulong _refCount{1};
    wil::unique_handle _file;
    std::wstring _path;
    bool _committed = false;
};
} // namespace

HRESULT STDMETHODCALLTYPE FileSystem::QueryInterface(REFIID riid, void** ppvObject) noexcept
{
    if (ppvObject == nullptr)
    {
        return E_POINTER;
    }

    if (riid == __uuidof(IUnknown) || riid == __uuidof(IFileSystem))
    {
        *ppvObject = static_cast<IFileSystem*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemSearch))
    {
        *ppvObject = static_cast<IFileSystemSearch*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemIO))
    {
        *ppvObject = static_cast<IFileSystemIO*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemDirectoryOperations))
    {
        *ppvObject = static_cast<IFileSystemDirectoryOperations*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemDirectoryWatch))
    {
        *ppvObject = static_cast<IFileSystemDirectoryWatch*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IInformations))
    {
        *ppvObject = static_cast<IInformations*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(INavigationMenu))
    {
        *ppvObject = static_cast<INavigationMenu*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IDriveInfo))
    {
        *ppvObject = static_cast<IDriveInfo*>(this);
        AddRef();
        return S_OK;
    }

    *ppvObject = nullptr;
    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE FileSystem::AddRef() noexcept
{
    return static_cast<ULONG>(_refCount.fetch_add(1, std::memory_order_relaxed) + 1);
}

ULONG STDMETHODCALLTYPE FileSystem::Release() noexcept
{
    const ULONG current = static_cast<ULONG>(_refCount.fetch_sub(1, std::memory_order_acq_rel) - 1);
    if (current == 0)
    {
        delete this;
    }
    return current;
}

HRESULT STDMETHODCALLTYPE FileSystem::CreateFileReader(const wchar_t* path, IFileReader** reader) noexcept
{
    if (reader == nullptr)
    {
        return E_POINTER;
    }

    *reader = nullptr;

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    wil::unique_handle file(
        CreateFileW(path, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
    if (! file)
    {
        const DWORD lastError = GetLastError();
        return HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_FILE_NOT_FOUND);
    }

    LARGE_INTEGER fileSize{};
    if (GetFileSizeEx(file.get(), &fileSize) == 0)
    {
        const DWORD lastError = GetLastError();
        return HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_GEN_FAILURE);
    }

    if (fileSize.QuadPart < 0)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }

    auto* impl = new (std::nothrow) Win32FileReader(std::move(file), static_cast<uint64_t>(fileSize.QuadPart));
    if (! impl)
    {
        return E_OUTOFMEMORY;
    }

    *reader = impl;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::CreateFileWriter(const wchar_t* path, FileSystemFlags flags, IFileWriter** writer) noexcept
{
    if (writer == nullptr)
    {
        return E_POINTER;
    }

    *writer = nullptr;

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const bool allowOverwrite       = (flags & FILESYSTEM_FLAG_ALLOW_OVERWRITE) != 0;
    const bool allowReplaceReadOnly = (flags & FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY) != 0;
    const DWORD creationDisposition = allowOverwrite ? CREATE_ALWAYS : CREATE_NEW;

    auto tryCreate = [&](DWORD* outLastError) -> wil::unique_handle
    {
        wil::unique_handle file(CreateFileW(path, GENERIC_WRITE, FILE_SHARE_READ, nullptr, creationDisposition, FILE_ATTRIBUTE_NORMAL, nullptr));
        if (file)
        {
            if (outLastError)
            {
                *outLastError = ERROR_SUCCESS;
            }
            return file;
        }

        const DWORD lastError = GetLastError();
        if (outLastError)
        {
            *outLastError = lastError != 0 ? lastError : ERROR_GEN_FAILURE;
        }
        return {};
    };

    DWORD lastError         = ERROR_SUCCESS;
    wil::unique_handle file = tryCreate(&lastError);
    if (! file && allowReplaceReadOnly && (lastError == ERROR_ACCESS_DENIED || lastError == ERROR_SHARING_VIOLATION))
    {
        const DWORD attributes = GetFileAttributesW(path);
        if (attributes != INVALID_FILE_ATTRIBUTES && (attributes & FILE_ATTRIBUTE_READONLY) != 0)
        {
            static_cast<void>(SetFileAttributesW(path, attributes & ~FILE_ATTRIBUTE_READONLY));
            file = tryCreate(&lastError);
        }
    }

    if (! file)
    {
        return HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_GEN_FAILURE);
    }

    auto* impl = new (std::nothrow) Win32FileWriter(std::move(file), std::wstring(path));
    if (! impl)
    {
        return E_OUTOFMEMORY;
    }

    *writer = impl;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::GetFileBasicInformation(const wchar_t* path, FileSystemBasicInformation* info) noexcept
{
    if (info == nullptr)
    {
        return E_POINTER;
    }

    if (info->sizeBytes != sizeof(FileSystemBasicInformation))
    {
        return E_INVALIDARG;
    }

    info->creationTime   = 0;
    info->lastAccessTime = 0;
    info->lastWriteTime  = 0;
    info->attributes     = 0;

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    wil::unique_handle file(CreateFileW(
        path, FILE_READ_ATTRIBUTES, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, nullptr));
    if (! file)
    {
        const DWORD lastError = GetLastError();
        return HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_FILE_NOT_FOUND);
    }

    FILE_BASIC_INFO basic{};
    if (! GetFileInformationByHandleEx(file.get(), FileBasicInfo, &basic, sizeof(basic)))
    {
        const DWORD lastError = GetLastError();
        return HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_GEN_FAILURE);
    }

    info->creationTime   = basic.CreationTime.QuadPart;
    info->lastAccessTime = basic.LastAccessTime.QuadPart;
    info->lastWriteTime  = basic.LastWriteTime.QuadPart;
    info->attributes     = basic.FileAttributes;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::SetFileBasicInformation(const wchar_t* path, const FileSystemBasicInformation* info) noexcept
{
    if (info == nullptr)
    {
        return E_POINTER;
    }

    if (info->sizeBytes != sizeof(FileSystemBasicInformation))
    {
        return E_INVALIDARG;
    }

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    wil::unique_handle file(CreateFileW(
        path, FILE_WRITE_ATTRIBUTES, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, nullptr));
    if (! file)
    {
        const DWORD lastError = GetLastError();
        return HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_FILE_NOT_FOUND);
    }

    FILE_BASIC_INFO basic{};
    basic.CreationTime.QuadPart   = info->creationTime;
    basic.LastAccessTime.QuadPart = info->lastAccessTime;
    basic.LastWriteTime.QuadPart  = info->lastWriteTime;
    basic.FileAttributes          = info->attributes;

    if (! SetFileInformationByHandle(file.get(), FileBasicInfo, &basic, sizeof(basic)))
    {
        const DWORD lastError = GetLastError();
        return HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_GEN_FAILURE);
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::GetCapabilities(const char** jsonUtf8) noexcept
{
    if (jsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);

    if (_capabilitiesJson.empty())
    {
        UpdateCapabilitiesJson(); // requires _stateMutex
    }

    *jsonUtf8 = _capabilitiesJson.empty() ? "{}" : _capabilitiesJson.c_str();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::GetTransferHints(const wchar_t* path,
                                                       [[maybe_unused]] FileSystemOperation operationType,
                                                       [[maybe_unused]] FileSystemTransferEndpoint endpoint,
                                                       FileSystemTransferHints* hints) noexcept
{
    if (path == nullptr || path[0] == L'\0' || hints == nullptr)
    {
        return E_INVALIDARG;
    }
    if (hints->sizeBytes < sizeof(FileSystemTransferHints))
    {
        return E_INVALIDARG;
    }

    hints->latencyClass              = FILESYSTEM_TRANSFER_LATENCY_UNKNOWN;
    hints->flags                     = FILESYSTEM_TRANSFER_HINT_NONE;
    hints->preferredBufferBytes      = 0;
    hints->preferredProgressPeriodMs = 0;

    const std::wstring driveRoot = ExtractDriveRoot(path);
    const UINT driveType         = ! driveRoot.empty() ? GetDriveTypeW(driveRoot.c_str()) : DRIVE_UNKNOWN;
    FillTransferHintsLocal(*hints, driveType == DRIVE_REMOTE || IsUncPath(path));
    hints->preferredBufferBytes = ClampPreferredBridgeBufferBytes(hints->preferredBufferBytes);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::GetStorageCharacteristics(const wchar_t* path, FileSystemStorageCharacteristics* characteristics) noexcept
{
    if (path == nullptr || path[0] == L'\0' || characteristics == nullptr)
    {
        return E_INVALIDARG;
    }
    if (characteristics->sizeBytes < sizeof(FileSystemStorageCharacteristics))
    {
        return E_INVALIDARG;
    }

    characteristics->storageKind                  = FILESYSTEM_STORAGE_UNKNOWN;
    characteristics->flags                        = FILESYSTEM_STORAGE_FLAG_NONE;
    characteristics->queueDepthHint               = 0;
    characteristics->preferredCopyMoveConcurrency = 0;
    characteristics->preferredDeleteConcurrency   = 0;

    const std::wstring driveRoot = ExtractDriveRoot(path);
    const UINT driveType         = ! driveRoot.empty() ? GetDriveTypeW(driveRoot.c_str()) : DRIVE_UNKNOWN;
    FillStorageCharacteristicsLocal(*characteristics, driveType == DRIVE_REMOTE || IsUncPath(path));
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::GetItemProperties(const wchar_t* path, const char** jsonUtf8) noexcept
{
    if (jsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *jsonUtf8 = nullptr;

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    WIN32_FILE_ATTRIBUTE_DATA data{};
    if (GetFileAttributesExW(path, GetFileExInfoStandard, &data) == 0)
    {
        const DWORD lastError = GetLastError();
        return HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_GEN_FAILURE);
    }

    const bool isDirectory   = (data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
    const uint64_t sizeBytes = isDirectory ? 0ull : (static_cast<uint64_t>(data.nFileSizeHigh) << 32u) | static_cast<uint64_t>(data.nFileSizeLow);

    yyjson_mut_doc* doc = yyjson_mut_doc_new(nullptr);
    if (! doc)
    {
        return E_OUTOFMEMORY;
    }
    auto freeDoc = wil::scope_exit([&] { yyjson_mut_doc_free(doc); });

    yyjson_mut_val* root = yyjson_mut_obj(doc);
    yyjson_mut_doc_set_root(doc, root);

    yyjson_mut_obj_add_int(doc, root, "version", 1);
    yyjson_mut_obj_add_str(doc, root, "title", "properties");

    yyjson_mut_val* sections = yyjson_mut_arr(doc);
    yyjson_mut_obj_add_val(doc, root, "sections", sections);

    auto addSection = [&](const char* title) noexcept
    {
        yyjson_mut_val* section = yyjson_mut_obj(doc);
        yyjson_mut_arr_add_val(sections, section);
        yyjson_mut_obj_add_str(doc, section, "title", title);

        yyjson_mut_val* sectionFields = yyjson_mut_arr(doc);
        yyjson_mut_obj_add_val(doc, section, "fields", sectionFields);
        return sectionFields;
    };

    const std::wstring fullPath(path);
    const std::filesystem::path fullPathFs(fullPath);
    const std::wstring name       = fullPathFs.filename().wstring();
    const std::wstring extension  = ! isDirectory ? fullPathFs.extension().wstring() : std::wstring{};
    const std::wstring parentPath = fullPathFs.parent_path().wstring();
    const std::wstring rootPath   = fullPathFs.root_path().wstring();

    auto addField = [&](yyjson_mut_val* sectionFields, const char* key, const std::string& value) noexcept
    {
        if (! sectionFields || value.empty())
        {
            return;
        }

        yyjson_mut_val* field = yyjson_mut_obj(doc);
        yyjson_mut_obj_add_strcpy(doc, field, "key", key);
        yyjson_mut_obj_add_strncpy(doc, field, "value", value.data(), value.size());
        yyjson_mut_arr_add_val(sectionFields, field);
    };

    yyjson_mut_val* generalFields = addSection("General");
    addField(generalFields, "Name", Utf8FromUtf16(name.empty() ? std::wstring_view(fullPath) : std::wstring_view(name)));
    addField(generalFields, "Path", Utf8FromUtf16(fullPath));
    addField(generalFields, "Parent", Utf8FromUtf16(parentPath));
    addField(generalFields, "Root", Utf8FromUtf16(rootPath));
    addField(generalFields, "Type", isDirectory ? std::string("Directory") : std::string("File"));
    if (! extension.empty())
    {
        addField(generalFields, "Extension", Utf8FromUtf16(extension));
    }
    if (! isDirectory)
    {
        addField(generalFields, "Size", std::format("{} bytes", sizeBytes));
    }

    yyjson_mut_val* timestampFields = addSection("Timestamps");
    addField(timestampFields, "Created", FormatFileTimeLocal(data.ftCreationTime));
    addField(timestampFields, "Modified", FormatFileTimeLocal(data.ftLastWriteTime));
    addField(timestampFields, "Accessed", FormatFileTimeLocal(data.ftLastAccessTime));

    yyjson_mut_val* attributeFields = addSection("Attributes");
    addField(attributeFields, "Raw", std::format("0x{:08X}", data.dwFileAttributes));
    addField(attributeFields, "Flags", FormatFileAttributeFlags(data.dwFileAttributes));

    const char* written = yyjson_mut_write(doc, YYJSON_WRITE_NOFLAG, nullptr);
    if (! written)
    {
        return E_OUTOFMEMORY;
    }
    auto freeWritten = wil::scope_exit([&] { free(const_cast<char*>(written)); });

    {
        std::scoped_lock lock(_propertiesMutex);
        _lastPropertiesJson.assign(written);
        *jsonUtf8 = _lastPropertiesJson.c_str();
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::GetMetaData(const PluginMetaData** metaData) noexcept
{
    if (metaData == nullptr)
    {
        return E_POINTER;
    }

    *metaData = &_metaData;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::GetConfigurationSchema(const char** schemaJsonUtf8) noexcept
{
    if (schemaJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *schemaJsonUtf8 = StaticConfigurationSchema();
    return S_OK;
}

const char* GetFileSystemStaticConfigurationSchema() noexcept
{
    return FileSystem::StaticConfigurationSchema();
}

const char* FileSystem::StaticConfigurationSchema() noexcept
{
    return kSchemaJson;
}

HRESULT STDMETHODCALLTYPE FileSystem::SetConfiguration(const char* configurationJsonUtf8) noexcept
{
    FileSystemConcurrencyMode concurrencyMode                 = kDefaultConcurrencyMode;
    unsigned int copyMoveMaxConcurrency                       = kDefaultCopyMoveMaxConcurrency;
    unsigned int deleteMaxConcurrency                         = kDefaultDeleteMaxConcurrency;
    unsigned int deleteRecycleBinMaxConcurrency               = kDefaultDeleteRecycleBinMaxConcurrency;
    unsigned int recycleBinBatchSize                          = kDefaultRecycleBinBatchSize;
    unsigned long enumerationSoftMaxBufferMiB                 = kDefaultEnumerationSoftMaxBufferMiB;
    unsigned long enumerationHardMaxBufferMiB                 = kDefaultEnumerationHardMaxBufferMiB;
    FileSystemReparsePointPolicy reparsePointPolicy           = kDefaultReparsePointPolicy;
    FileSystemSearchBackendPreference searchBackendPreference = kDefaultSearchBackendPreference;
    unsigned int searchMaxDirectoryWalkers                    = kDefaultSearchMaxDirectoryWalkers;
#ifdef _DEBUG
    unsigned int directorySizeDelayMs = 0u;
#endif

    constexpr unsigned long kMiB     = 1024u * 1024u;
    const unsigned long maxBufferMiB = (std::numeric_limits<unsigned long>::max)() / kMiB;

    if (configurationJsonUtf8 != nullptr && configurationJsonUtf8[0] != '\0')
    {
        const std::string_view configText(configurationJsonUtf8);

        yyjson_doc* doc = yyjson_read(configText.data(), configText.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
        if (doc)
        {
            auto freeDoc = wil::scope_exit([&] { yyjson_doc_free(doc); });

            yyjson_val* root = yyjson_doc_get_root(doc);
            if (root && yyjson_is_obj(root))
            {
                yyjson_val* concurrencyModeVal = yyjson_obj_get(root, "concurrencyMode");
                if (concurrencyModeVal && yyjson_is_str(concurrencyModeVal))
                {
                    const char* valueText = yyjson_get_str(concurrencyModeVal);
                    if (valueText && valueText[0] != '\0')
                    {
                        concurrencyMode = ParseConcurrencyMode(valueText);
                    }
                }

                yyjson_val* copyMoveVal = yyjson_obj_get(root, "copyMoveMaxConcurrency");
                if (copyMoveVal && yyjson_is_int(copyMoveVal))
                {
                    const int64_t value = yyjson_get_int(copyMoveVal);
                    if (value >= 1)
                    {
                        copyMoveMaxConcurrency = static_cast<unsigned int>(std::min<int64_t>(value, static_cast<int64_t>(kMaxCopyMoveMaxConcurrency)));
                    }
                }

                yyjson_val* deleteVal = yyjson_obj_get(root, "deleteMaxConcurrency");
                if (deleteVal && yyjson_is_int(deleteVal))
                {
                    const int64_t value = yyjson_get_int(deleteVal);
                    if (value >= 1)
                    {
                        deleteMaxConcurrency = static_cast<unsigned int>(std::min<int64_t>(value, static_cast<int64_t>(kMaxDeleteMaxConcurrency)));
                    }
                }

                yyjson_val* deleteRecycleVal = yyjson_obj_get(root, "deleteRecycleBinMaxConcurrency");
                if (deleteRecycleVal && yyjson_is_int(deleteRecycleVal))
                {
                    const int64_t value = yyjson_get_int(deleteRecycleVal);
                    if (value >= 1)
                    {
                        deleteRecycleBinMaxConcurrency =
                            static_cast<unsigned int>(std::min<int64_t>(value, static_cast<int64_t>(kMaxDeleteRecycleBinMaxConcurrency)));
                    }
                }

                yyjson_val* recycleBinBatchSizeVal = yyjson_obj_get(root, "recycleBinBatchSize");
                if (recycleBinBatchSizeVal && yyjson_is_int(recycleBinBatchSizeVal))
                {
                    const int64_t value = yyjson_get_int(recycleBinBatchSizeVal);
                    if (value >= 1)
                    {
                        recycleBinBatchSize = static_cast<unsigned int>(std::min<int64_t>(value, static_cast<int64_t>(kMaxRecycleBinBatchSize)));
                    }
                }

                yyjson_val* softMaxVal = yyjson_obj_get(root, "enumerationSoftMaxBufferMiB");
                if (softMaxVal && yyjson_is_int(softMaxVal))
                {
                    const int64_t value = yyjson_get_int(softMaxVal);
                    if (value >= 1)
                    {
                        enumerationSoftMaxBufferMiB = static_cast<unsigned long>(std::min<int64_t>(value, static_cast<int64_t>(maxBufferMiB)));
                    }
                }

                yyjson_val* hardMaxVal = yyjson_obj_get(root, "enumerationHardMaxBufferMiB");
                if (hardMaxVal && yyjson_is_int(hardMaxVal))
                {
                    const int64_t value = yyjson_get_int(hardMaxVal);
                    if (value >= 1)
                    {
                        enumerationHardMaxBufferMiB = static_cast<unsigned long>(std::min<int64_t>(value, static_cast<int64_t>(maxBufferMiB)));
                    }
                }

                yyjson_val* reparsePolicyVal = yyjson_obj_get(root, "reparsePointPolicy");
                if (reparsePolicyVal && yyjson_is_str(reparsePolicyVal))
                {
                    const char* valueText = yyjson_get_str(reparsePolicyVal);
                    if (valueText && valueText[0] != '\0')
                    {
                        reparsePointPolicy = ParseReparsePointPolicy(valueText);
                    }
                }

                yyjson_val* searchBackendVal = yyjson_obj_get(root, "searchBackendPreference");
                if (searchBackendVal && yyjson_is_str(searchBackendVal))
                {
                    const char* valueText = yyjson_get_str(searchBackendVal);
                    if (valueText && valueText[0] != '\0')
                    {
                        searchBackendPreference = ParseSearchBackendPreference(valueText);
                    }
                }

                yyjson_val* searchWalkersVal = yyjson_obj_get(root, "searchMaxDirectoryWalkers");
                if (searchWalkersVal && yyjson_is_int(searchWalkersVal))
                {
                    const int64_t value = yyjson_get_int(searchWalkersVal);
                    if (value >= 1)
                    {
                        searchMaxDirectoryWalkers = static_cast<unsigned int>(std::min<int64_t>(value, static_cast<int64_t>(kMaxSearchMaxDirectoryWalkers)));
                    }
                }

#ifdef _DEBUG
                yyjson_val* delayVal = yyjson_obj_get(root, "directorySizeDelayMs");
                if (delayVal && yyjson_is_int(delayVal))
                {
                    const int64_t value = yyjson_get_int(delayVal);
                    if (value >= 0)
                    {
                        directorySizeDelayMs = static_cast<unsigned int>(std::min<int64_t>(value, 50));
                    }
                }
#endif
            }
        }
    }

    copyMoveMaxConcurrency         = std::clamp(copyMoveMaxConcurrency, 1u, kMaxCopyMoveMaxConcurrency);
    deleteMaxConcurrency           = std::clamp(deleteMaxConcurrency, 1u, kMaxDeleteMaxConcurrency);
    deleteRecycleBinMaxConcurrency = std::clamp(deleteRecycleBinMaxConcurrency, 1u, kMaxDeleteRecycleBinMaxConcurrency);
    recycleBinBatchSize            = std::clamp(recycleBinBatchSize, 1u, kMaxRecycleBinBatchSize);
    searchMaxDirectoryWalkers      = std::clamp(searchMaxDirectoryWalkers, 1u, kMaxSearchMaxDirectoryWalkers);

    enumerationSoftMaxBufferMiB = std::clamp(enumerationSoftMaxBufferMiB, 1ul, maxBufferMiB);
    enumerationHardMaxBufferMiB = std::clamp(enumerationHardMaxBufferMiB, enumerationSoftMaxBufferMiB, maxBufferMiB);

    std::string newConfigurationJson = BuildConfigurationJson(concurrencyMode,
                                                              copyMoveMaxConcurrency,
                                                              deleteMaxConcurrency,
                                                              deleteRecycleBinMaxConcurrency,
                                                              recycleBinBatchSize,
                                                              enumerationSoftMaxBufferMiB,
                                                              enumerationHardMaxBufferMiB,
                                                              reparsePointPolicy,
                                                              searchBackendPreference,
                                                              searchMaxDirectoryWalkers);

    std::lock_guard lock(_stateMutex);

    _concurrencyMode                = concurrencyMode;
    _copyMoveMaxConcurrency         = copyMoveMaxConcurrency;
    _deleteMaxConcurrency           = deleteMaxConcurrency;
    _deleteRecycleBinMaxConcurrency = deleteRecycleBinMaxConcurrency;
    _recycleBinBatchSize            = recycleBinBatchSize;
    _enumerationSoftMaxBufferMiB    = enumerationSoftMaxBufferMiB;
    _enumerationHardMaxBufferMiB    = enumerationHardMaxBufferMiB;
    _reparsePointPolicy             = reparsePointPolicy;
    _searchBackendPreference        = searchBackendPreference;
    _searchMaxDirectoryWalkers      = searchMaxDirectoryWalkers;
#ifdef _DEBUG
    _directorySizeDelayMs = directorySizeDelayMs;
#endif

    _configurationJson = std::move(newConfigurationJson);
    UpdateCapabilitiesJson();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::GetConfiguration(const char** configurationJsonUtf8) noexcept
{
    if (configurationJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);

    *configurationJsonUtf8 = _configurationJson.empty() ? "{}" : _configurationJson.c_str();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::SomethingToSave(BOOL* pSomethingToSave) noexcept
{
    if (pSomethingToSave == nullptr)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);
    const bool isDefault = _concurrencyMode == kDefaultConcurrencyMode && _copyMoveMaxConcurrency == kDefaultCopyMoveMaxConcurrency &&
                           _deleteMaxConcurrency == kDefaultDeleteMaxConcurrency && _deleteRecycleBinMaxConcurrency == kDefaultDeleteRecycleBinMaxConcurrency &&
                           _recycleBinBatchSize == kDefaultRecycleBinBatchSize && _enumerationSoftMaxBufferMiB == kDefaultEnumerationSoftMaxBufferMiB &&
                           _enumerationHardMaxBufferMiB == kDefaultEnumerationHardMaxBufferMiB && _reparsePointPolicy == kDefaultReparsePointPolicy &&
                           _searchBackendPreference == kDefaultSearchBackendPreference && _searchMaxDirectoryWalkers == kDefaultSearchMaxDirectoryWalkers;
    *pSomethingToSave    = isDefault ? FALSE : TRUE;
    return S_OK;
}

void FileSystem::UpdateCapabilitiesJson() noexcept
{
    // NOTE: Caller must hold _stateMutex.
    _capabilitiesJson = std::format(
        R"json({{"version":1,"operations":{{"copy":true,"move":true,"delete":true,"rename":true,"properties":true,"read":true,"write":true}},"search":{{"version":1,"name":true,"content":true,"indexed":true,"serviceBacked":true,"supportsRegex":true,"supportsSnippets":true,"preferredBackend":"service"}},"concurrency":{{"copyMoveMax":{},"deleteMax":{},"deleteRecycleBinMax":{}}},"crossFileSystem":{{"export":{{"copy":["*"],"move":["*"]}},"import":{{"copy":["*"],"move":["*"]}}}}}})json",
        std::clamp(_copyMoveMaxConcurrency, 1u, kMaxCopyMoveMaxConcurrency),
        std::clamp(_deleteMaxConcurrency, 1u, kMaxDeleteMaxConcurrency),
        std::clamp(_deleteRecycleBinMaxConcurrency, 1u, kMaxDeleteRecycleBinMaxConcurrency));
}
