#include "FileSystem.Internal.h"

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

    HRESULT STDMETHODCALLTYPE GetSize(unsigned __int64* sizeBytes) noexcept override
    {
        if (sizeBytes == nullptr)
        {
            return E_POINTER;
        }

        *sizeBytes = _sizeBytes;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Seek(__int64 offset, unsigned long origin, unsigned __int64* newPosition) noexcept override
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

        *newPosition = static_cast<unsigned __int64>(moved.QuadPart);
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

    HRESULT STDMETHODCALLTYPE GetPosition(unsigned __int64* positionBytes) noexcept override
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

        *positionBytes = static_cast<unsigned __int64>(moved.QuadPart);
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
    const ULONG current = static_cast<ULONG>(_refCount.fetch_sub(1, std::memory_order_relaxed) - 1);
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

    *info = {};

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

    const bool isDirectory = (data.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
    const unsigned __int64 sizeBytes =
        isDirectory ? 0ull : (static_cast<unsigned __int64>(data.nFileSizeHigh) << 32u) | static_cast<unsigned __int64>(data.nFileSizeLow);

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

    yyjson_mut_val* general = yyjson_mut_obj(doc);
    yyjson_mut_arr_add_val(sections, general);
    yyjson_mut_obj_add_str(doc, general, "title", "general");

    yyjson_mut_val* fields = yyjson_mut_arr(doc);
    yyjson_mut_obj_add_val(doc, general, "fields", fields);

    const std::wstring fullPath(path);
    const std::wstring name = std::filesystem::path(fullPath).filename().wstring();

    auto addField = [&](const char* key, const std::string& value)
    {
        yyjson_mut_val* field = yyjson_mut_obj(doc);
        yyjson_mut_obj_add_strcpy(doc, field, "key", key);
        yyjson_mut_obj_add_strncpy(doc, field, "value", value.data(), value.size());
        yyjson_mut_arr_add_val(fields, field);
    };

    addField("name", Utf8FromUtf16(name.empty() ? std::wstring_view(fullPath) : std::wstring_view(name)));
    addField("path", Utf8FromUtf16(fullPath));
    addField("type", isDirectory ? std::string("directory") : std::string("file"));
    if (! isDirectory)
    {
        addField("sizeBytes", std::format("{}", sizeBytes));
    }

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

    *schemaJsonUtf8 = kSchemaJson;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem::SetConfiguration(const char* configurationJsonUtf8) noexcept
{
    unsigned int copyMoveMaxConcurrency             = kDefaultCopyMoveMaxConcurrency;
    unsigned int deleteMaxConcurrency               = kDefaultDeleteMaxConcurrency;
    unsigned int deleteRecycleBinMaxConcurrency     = kDefaultDeleteRecycleBinMaxConcurrency;
    unsigned long enumerationSoftMaxBufferMiB       = kDefaultEnumerationSoftMaxBufferMiB;
    unsigned long enumerationHardMaxBufferMiB       = kDefaultEnumerationHardMaxBufferMiB;
    FileSystemReparsePointPolicy reparsePointPolicy = kDefaultReparsePointPolicy;
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

    enumerationSoftMaxBufferMiB = std::clamp(enumerationSoftMaxBufferMiB, 1ul, maxBufferMiB);
    enumerationHardMaxBufferMiB = std::clamp(enumerationHardMaxBufferMiB, enumerationSoftMaxBufferMiB, maxBufferMiB);

    std::string newConfigurationJson;
    newConfigurationJson = std::format("{{\"copyMoveMaxConcurrency\":{},\"deleteMaxConcurrency\":{},\"deleteRecycleBinMaxConcurrency\":{},"
                                       "\"enumerationSoftMaxBufferMiB\":{},\"enumerationHardMaxBufferMiB\":{},\"reparsePointPolicy\":\"{}\"}}",
                                       copyMoveMaxConcurrency,
                                       deleteMaxConcurrency,
                                       deleteRecycleBinMaxConcurrency,
                                       enumerationSoftMaxBufferMiB,
                                       enumerationHardMaxBufferMiB,
                                       ReparsePointPolicyToString(reparsePointPolicy));

    std::lock_guard lock(_stateMutex);

    _copyMoveMaxConcurrency         = copyMoveMaxConcurrency;
    _deleteMaxConcurrency           = deleteMaxConcurrency;
    _deleteRecycleBinMaxConcurrency = deleteRecycleBinMaxConcurrency;
    _enumerationSoftMaxBufferMiB    = enumerationSoftMaxBufferMiB;
    _enumerationHardMaxBufferMiB    = enumerationHardMaxBufferMiB;
    _reparsePointPolicy             = reparsePointPolicy;
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
    const bool isDefault = _copyMoveMaxConcurrency == kDefaultCopyMoveMaxConcurrency && _deleteMaxConcurrency == kDefaultDeleteMaxConcurrency &&
                           _deleteRecycleBinMaxConcurrency == kDefaultDeleteRecycleBinMaxConcurrency &&
                           _enumerationSoftMaxBufferMiB == kDefaultEnumerationSoftMaxBufferMiB &&
                           _enumerationHardMaxBufferMiB == kDefaultEnumerationHardMaxBufferMiB && _reparsePointPolicy == kDefaultReparsePointPolicy;
    *pSomethingToSave = isDefault ? FALSE : TRUE;
    return S_OK;
}

void FileSystem::UpdateCapabilitiesJson() noexcept
{
    // NOTE: Caller must hold _stateMutex.
    _capabilitiesJson = std::format(
        R"json({{"version":1,"operations":{{"copy":true,"move":true,"delete":true,"rename":true,"properties":true,"read":true,"write":true}},"concurrency":{{"copyMoveMax":{},"deleteMax":{},"deleteRecycleBinMax":{}}},"crossFileSystem":{{"export":{{"copy":["*"],"move":["*"]}},"import":{{"copy":["*"],"move":["*"]}}}}}})json",
        std::clamp(_copyMoveMaxConcurrency, 1u, kMaxCopyMoveMaxConcurrency),
        std::clamp(_deleteMaxConcurrency, 1u, kMaxDeleteMaxConcurrency),
        std::clamp(_deleteRecycleBinMaxConcurrency, 1u, kMaxDeleteRecycleBinMaxConcurrency));
}
