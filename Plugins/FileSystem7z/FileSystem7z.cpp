#include <algorithm>
#include <condition_variable>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <cwctype>
#include <filesystem>
#include <format>
#include <limits>
#include <mutex>
#include <new>
#include <string>
#include <string_view>
#include <system_error>
#include <thread>
#include <utility>
#include <vector>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <oleauto.h>

#ifndef INITGUID
#define INITGUID
#define REDSALAMANDER_UNDEF_INITGUID
#endif
#include <7zip/CPP/7zip/Archive/IArchive.h>
#include <7zip/CPP/7zip/IPassword.h>
#ifdef REDSALAMANDER_UNDEF_INITGUID
#undef INITGUID
#undef REDSALAMANDER_UNDEF_INITGUID
#endif

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#pragma warning(push)
// (C6297) Arithmetic overflow. Results might not be an expected value.
// (C28182) Dereferencing NULL pointer.
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

#include "FileSystem7z.h"

#include "Helpers.h"

#pragma comment(lib, "Ole32.lib")
#pragma comment(lib, "OleAut32.lib")

namespace
{
std::wstring Utf16FromMultiByte(std::string_view text, UINT codePage, DWORD flags) noexcept
{
    if (text.empty())
    {
        return {};
    }

    if (text.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return {};
    }

    const int required = MultiByteToWideChar(codePage, flags, text.data(), static_cast<int>(text.size()), nullptr, 0);
    if (required <= 0)
    {
        return {};
    }

    std::wstring result(static_cast<size_t>(required), L'\0');
    const int written = MultiByteToWideChar(codePage, flags, text.data(), static_cast<int>(text.size()), result.data(), required);
    if (written != required)
    {
        return {};
    }

    return result;
}

std::optional<std::wstring> TryGetJsonString(yyjson_val* obj, const char* key) noexcept
{
    if (! obj || ! key)
    {
        return std::nullopt;
    }

    yyjson_val* val = yyjson_obj_get(obj, key);
    if (! val || ! yyjson_is_str(val))
    {
        return std::nullopt;
    }

    const char* s = yyjson_get_str(val);
    if (! s)
    {
        return std::nullopt;
    }

    const size_t len        = yyjson_get_len(val);
    const std::wstring wide = Utf16FromMultiByte(std::string_view(s, len), CP_UTF8, MB_ERR_INVALID_CHARS);
    if (wide.empty() && len != 0u)
    {
        return std::nullopt;
    }

    return wide;
}

HRESULT
CreateSevenZipItemFileReader(std::wstring archivePath, std::wstring password, uint32_t itemIndex, uint64_t sizeBytes, IFileReader** outReader) noexcept;
} // namespace

// FilesInformation7z

HRESULT STDMETHODCALLTYPE FilesInformation7z::QueryInterface(REFIID riid, void** ppvObject) noexcept
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

ULONG STDMETHODCALLTYPE FilesInformation7z::AddRef() noexcept
{
    return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
}

ULONG STDMETHODCALLTYPE FilesInformation7z::Release() noexcept
{
    const ULONG result = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
    if (result == 0)
    {
        delete this;
    }
    return result;
}

HRESULT STDMETHODCALLTYPE FilesInformation7z::GetBuffer(FileInfo** ppFileInfo) noexcept
{
    if (ppFileInfo == nullptr)
    {
        return E_POINTER;
    }

    if (_usedBytes == 0 || _buffer.empty())
    {
        *ppFileInfo = nullptr;
        return S_OK;
    }

    *ppFileInfo = reinterpret_cast<FileInfo*>(_buffer.data());
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FilesInformation7z::GetBufferSize(unsigned long* pSize) noexcept
{
    if (pSize == nullptr)
    {
        return E_POINTER;
    }

    *pSize = _usedBytes;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FilesInformation7z::GetAllocatedSize(unsigned long* pSize) noexcept
{
    if (pSize == nullptr)
    {
        return E_POINTER;
    }

    if (_buffer.size() > static_cast<size_t>(std::numeric_limits<unsigned long>::max()))
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    *pSize = static_cast<unsigned long>(_buffer.size());
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FilesInformation7z::GetCount(unsigned long* pCount) noexcept
{
    if (pCount == nullptr)
    {
        return E_POINTER;
    }

    *pCount = _count;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FilesInformation7z::Get(unsigned long index, FileInfo** ppEntry) noexcept
{
    if (ppEntry == nullptr)
    {
        return E_POINTER;
    }

    *ppEntry = nullptr;

    if (index >= _count)
    {
        return HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES);
    }

    return LocateEntry(index, ppEntry);
}

size_t FilesInformation7z::AlignUp(size_t value, size_t alignment) noexcept
{
    const size_t mask = alignment - 1u;
    return (value + mask) & ~mask;
}

size_t FilesInformation7z::ComputeEntrySizeBytes(std::wstring_view name) noexcept
{
    const size_t baseSize = offsetof(FileInfo, FileName);
    const size_t nameSize = name.size() * sizeof(wchar_t);
    return AlignUp(baseSize + nameSize + sizeof(wchar_t), sizeof(unsigned long));
}

HRESULT FilesInformation7z::BuildFromEntries(std::vector<Entry> entries) noexcept
{
    _buffer.clear();
    _count     = 0;
    _usedBytes = 0;

    if (entries.empty())
    {
        return S_OK;
    }

    std::sort(entries.begin(),
              entries.end(),
              [](const Entry& a, const Entry& b)
              {
                  const int cmp = OrdinalString::Compare(a.name, b.name, true);
                  if (cmp != 0)
                  {
                      return cmp < 0;
                  }

                  const bool aDir = (a.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                  const bool bDir = (b.attributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                  if (aDir != bDir)
                  {
                      return aDir;
                  }
                  return a.sizeBytes < b.sizeBytes;
              });

    size_t totalBytes = 0;
    for (const auto& entry : entries)
    {
        totalBytes += ComputeEntrySizeBytes(entry.name);
        if (totalBytes > static_cast<size_t>(std::numeric_limits<unsigned long>::max()))
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }
    }

    _buffer.resize(totalBytes, std::byte{0});

    std::byte* base     = _buffer.data();
    size_t offset       = 0;
    FileInfo* previous  = nullptr;
    size_t previousSize = 0;

    for (const auto& source : entries)
    {
        const size_t entrySize = ComputeEntrySizeBytes(source.name);
        if (offset + entrySize > _buffer.size())
        {
            return E_FAIL;
        }

        auto* entry = reinterpret_cast<FileInfo*>(base + offset);
        std::memset(entry, 0, entrySize);

        const size_t nameBytes = source.name.size() * sizeof(wchar_t);
        if (nameBytes > static_cast<size_t>(std::numeric_limits<unsigned long>::max()))
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        entry->FileAttributes = static_cast<unsigned long>(source.attributes);
        entry->EndOfFile      = static_cast<__int64>(source.sizeBytes);
        entry->AllocationSize = static_cast<__int64>(source.sizeBytes);

        entry->CreationTime   = source.lastWriteTime;
        entry->LastAccessTime = source.lastWriteTime;
        entry->LastWriteTime  = source.lastWriteTime;
        entry->ChangeTime     = source.lastWriteTime;

        entry->FileNameSize = static_cast<unsigned long>(nameBytes);
        if (! source.name.empty())
        {
            std::memcpy(entry->FileName, source.name.data(), nameBytes);
        }
        entry->FileName[source.name.size()] = L'\0';

        if (previous)
        {
            previous->NextEntryOffset = static_cast<unsigned long>(previousSize);
        }

        previous     = entry;
        previousSize = entrySize;

        offset += entrySize;
        ++_count;
    }

    _usedBytes = static_cast<unsigned long>(_buffer.size());
    return S_OK;
}

HRESULT FilesInformation7z::LocateEntry(unsigned long index, FileInfo** ppEntry) const noexcept
{
    const std::byte* base      = _buffer.data();
    size_t offset              = 0;
    unsigned long currentIndex = 0;

    while (offset < _usedBytes && offset + sizeof(FileInfo) <= _buffer.size())
    {
        auto* entry = reinterpret_cast<const FileInfo*>(base + offset);
        if (currentIndex == index)
        {
            *ppEntry = const_cast<FileInfo*>(entry);
            return S_OK;
        }

        const size_t advance = (entry->NextEntryOffset != 0)
                                   ? static_cast<size_t>(entry->NextEntryOffset)
                                   : ComputeEntrySizeBytes(std::wstring_view(entry->FileName, static_cast<size_t>(entry->FileNameSize) / sizeof(wchar_t)));
        if (advance == 0)
        {
            break;
        }

        offset += advance;
        ++currentIndex;
    }

    return HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES);
}

// FileSystem7z

FileSystem7z::FileSystem7z()
{
    _metaData.id          = kPluginId;
    _metaData.shortId     = kPluginShortId;
    _metaData.name        = kPluginName;
    _metaData.description = kPluginDescription;
    _metaData.author      = kPluginAuthor;
    _metaData.version     = kPluginVersion;

    _configurationJson = "{}";

    _driveFileSystem = L"7z";
}

HRESULT STDMETHODCALLTYPE FileSystem7z::QueryInterface(REFIID riid, void** ppvObject) noexcept
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

    if (riid == __uuidof(IFileSystemInitialize))
    {
        *ppvObject = static_cast<IFileSystemInitialize*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IFileSystemDirectoryOperations))
    {
        *ppvObject = static_cast<IFileSystemDirectoryOperations*>(this);
        AddRef();
        return S_OK;
    }

    *ppvObject = nullptr;
    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE FileSystem7z::AddRef() noexcept
{
    return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
}

ULONG STDMETHODCALLTYPE FileSystem7z::Release() noexcept
{
    const ULONG result = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
    if (result == 0)
    {
        delete this;
    }
    return result;
}

HRESULT STDMETHODCALLTYPE FileSystem7z::GetMetaData(const PluginMetaData** metaData) noexcept
{
    if (metaData == nullptr)
    {
        return E_POINTER;
    }

    *metaData = &_metaData;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem7z::GetConfigurationSchema(const char** schemaJsonUtf8) noexcept
{
    if (schemaJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *schemaJsonUtf8 = kSchemaJson;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem7z::SetConfiguration(const char* configurationJsonUtf8) noexcept
{
    std::lock_guard lock(_stateMutex);

    _defaultPassword.clear();

    if (configurationJsonUtf8 == nullptr || configurationJsonUtf8[0] == '\0')
    {
        _configurationJson = "{}";
        return S_OK;
    }

    _configurationJson = configurationJsonUtf8;

    yyjson_read_err err{};
    yyjson_doc* doc = yyjson_read_opts(_configurationJson.data(), _configurationJson.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, &err);
    if (! doc)
    {
        return S_OK;
    }

    auto freeDoc = wil::scope_exit([&] { yyjson_doc_free(doc); });

    yyjson_val* root = yyjson_doc_get_root(doc);
    if (! root || ! yyjson_is_obj(root))
    {
        return S_OK;
    }

    const auto password = TryGetJsonString(root, "defaultPassword");
    if (password.has_value())
    {
        _defaultPassword = password.value();
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem7z::GetConfiguration(const char** configurationJsonUtf8) noexcept
{
    if (configurationJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);
    *configurationJsonUtf8 = _configurationJson.c_str();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem7z::SomethingToSave(BOOL* pSomethingToSave) noexcept
{
    if (pSomethingToSave == nullptr)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);
    const bool hasNonDefault = ! _configurationJson.empty() && _configurationJson != "{}";
    *pSomethingToSave        = hasNonDefault ? TRUE : FALSE;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem7z::GetMenuItems(const NavigationMenuItem** items, unsigned int* count) noexcept
{
    if (items == nullptr || count == nullptr)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);
    const std::wstring archivePath = _archivePath;

    _menuEntries.clear();
    _menuEntryView.clear();

    const std::wstring pluginName = _metaData.name ? _metaData.name : L"7-Zip";

    MenuEntry header;
    header.flags = NAV_MENU_ITEM_FLAG_HEADER;
    header.label = pluginName;
    _menuEntries.push_back(std::move(header));

    if (! archivePath.empty())
    {
        MenuEntry mountHeader;
        mountHeader.flags = NAV_MENU_ITEM_FLAG_HEADER;
        mountHeader.label = archivePath;
        _menuEntries.push_back(std::move(mountHeader));
    }

    MenuEntry separator;
    separator.flags = NAV_MENU_ITEM_FLAG_SEPARATOR;
    _menuEntries.push_back(std::move(separator));

    MenuEntry root;
    root.label    = L"/";
    root.path     = L"/";
    root.iconPath = archivePath;
    _menuEntries.push_back(std::move(root));

    _menuEntryView.reserve(_menuEntries.size());
    for (const auto& e : _menuEntries)
    {
        NavigationMenuItem item{};
        item.flags     = e.flags;
        item.label     = e.label.empty() ? nullptr : e.label.c_str();
        item.path      = e.path.empty() ? nullptr : e.path.c_str();
        item.iconPath  = e.iconPath.empty() ? nullptr : e.iconPath.c_str();
        item.commandId = e.commandId;
        _menuEntryView.push_back(item);
    }

    *items = _menuEntryView.empty() ? nullptr : _menuEntryView.data();
    *count = static_cast<unsigned int>(_menuEntryView.size());
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem7z::ExecuteMenuCommand([[maybe_unused]] unsigned int commandId) noexcept
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE FileSystem7z::SetCallback(INavigationMenuCallback* callback, void* cookie) noexcept
{
    std::lock_guard lock(_stateMutex);
    _navigationMenuCallback       = callback;
    _navigationMenuCallbackCookie = callback != nullptr ? cookie : nullptr;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem7z::GetDriveInfo(const wchar_t* /*path*/, DriveInfo* info) noexcept
{
    if (info == nullptr)
    {
        return E_POINTER;
    }

    std::lock_guard lock(_stateMutex);

    UpdateDriveInfoStringsLocked();

    _driveInfo = {};

    if (! _driveDisplayName.empty())
    {
        _driveInfo.flags       = static_cast<DriveInfoFlags>(_driveInfo.flags | DRIVE_INFO_FLAG_HAS_DISPLAY_NAME);
        _driveInfo.displayName = _driveDisplayName.c_str();
    }

    if (! _driveFileSystem.empty())
    {
        _driveInfo.flags      = static_cast<DriveInfoFlags>(_driveInfo.flags | DRIVE_INFO_FLAG_HAS_FILE_SYSTEM);
        _driveInfo.fileSystem = _driveFileSystem.c_str();
    }

    _driveVolumeLabel.clear();
    if (! _archivePath.empty())
    {
        const std::wstring_view archivePath = _archivePath;
        const size_t lastSlash              = archivePath.find_last_of(L"\\/");
        if (lastSlash == std::wstring_view::npos)
        {
            _driveVolumeLabel.assign(archivePath);
        }
        else if ((lastSlash + 1u) < archivePath.size())
        {
            _driveVolumeLabel.assign(archivePath.substr(lastSlash + 1u));
        }
    }

    if (! _driveVolumeLabel.empty())
    {
        _driveInfo.flags       = static_cast<DriveInfoFlags>(_driveInfo.flags | DRIVE_INFO_FLAG_HAS_VOLUME_LABEL);
        _driveInfo.volumeLabel = _driveVolumeLabel.c_str();
    }

    if (! _archivePath.empty())
    {
        WIN32_FILE_ATTRIBUTE_DATA attributes{};
        if (GetFileAttributesExW(_archivePath.c_str(), GetFileExInfoStandard, &attributes) != 0)
        {
            ULARGE_INTEGER size{};
            size.LowPart  = attributes.nFileSizeLow;
            size.HighPart = attributes.nFileSizeHigh;

            _driveInfo.flags      = static_cast<DriveInfoFlags>(_driveInfo.flags | DRIVE_INFO_FLAG_HAS_TOTAL_BYTES);
            _driveInfo.totalBytes = size.QuadPart;
        }
    }

    *info = _driveInfo;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem7z::GetDriveMenuItems(const wchar_t* /*path*/, const NavigationMenuItem** items, unsigned int* count) noexcept
{
    if (items == nullptr || count == nullptr)
    {
        return E_POINTER;
    }

    *items = nullptr;
    *count = 0;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem7z::ExecuteDriveMenuCommand(unsigned int /*commandId*/, const wchar_t* /*path*/) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystem7z::Initialize(const wchar_t* rootPath, const char* optionsJsonUtf8) noexcept
{
    if (rootPath == nullptr || rootPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    std::wstring normalizedArchivePath;
    {
        std::wstring_view text = Trim(std::wstring_view(rootPath));

        if (text.size() >= 2u && text.front() == L'"' && text.back() == L'"')
        {
            text.remove_prefix(1u);
            text.remove_suffix(1u);
            text = Trim(text);
        }

        if (text.size() >= 3u && CompareStringOrdinal(text.data(), 3, L"7z:", 3, TRUE) == CSTR_EQUAL)
        {
            text.remove_prefix(3u);
            text = Trim(text);
        }
        else if (text.size() >= 5u && CompareStringOrdinal(text.data(), 5, L"file:", 5, TRUE) == CSTR_EQUAL)
        {
            text.remove_prefix(5u);
            text = Trim(text);
        }

        const size_t bar = text.find(L'|');
        if (bar != std::wstring_view::npos)
        {
            text = Trim(text.substr(0, bar));
        }

        if (text.empty())
        {
            return E_INVALIDARG;
        }

        normalizedArchivePath.assign(text);
        std::replace(normalizedArchivePath.begin(), normalizedArchivePath.end(), L'/', L'\\');

        while (normalizedArchivePath.size() > 3u && (normalizedArchivePath.back() == L'\\' || normalizedArchivePath.back() == L'/'))
        {
            normalizedArchivePath.pop_back();
        }

        const bool isExtended = normalizedArchivePath.rfind(L"\\\\?\\", 0) == 0 || normalizedArchivePath.rfind(L"\\\\.\\", 0) == 0;
        if (! isExtended)
        {
            const DWORD required = GetFullPathNameW(normalizedArchivePath.c_str(), 0, nullptr, nullptr);
            if (required > 0)
            {
                std::wstring absolute;
                absolute.resize(static_cast<size_t>(required));

                const DWORD written = GetFullPathNameW(normalizedArchivePath.c_str(), required, absolute.data(), nullptr);
                if (written > 0 && written < required)
                {
                    absolute.resize(static_cast<size_t>(written));
                    normalizedArchivePath = std::move(absolute);
                }
            }
        }
    }

    std::lock_guard lock(_stateMutex);

    _archivePath = std::move(normalizedArchivePath);
    _password.clear();

    if (optionsJsonUtf8 && optionsJsonUtf8[0] != '\0')
    {
        const std::string_view utf8(optionsJsonUtf8);
        if (! utf8.empty())
        {
            std::string mutableJson;
            mutableJson.assign(utf8.data(), utf8.size());

            yyjson_read_err err{};
            yyjson_doc* doc = yyjson_read_opts(mutableJson.data(), mutableJson.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, &err);
            if (doc)
            {
                auto freeDoc     = wil::scope_exit([&] { yyjson_doc_free(doc); });
                yyjson_val* root = yyjson_doc_get_root(doc);
                if (root && yyjson_is_obj(root))
                {
                    const auto password = TryGetJsonString(root, "password");
                    if (password.has_value())
                    {
                        _password = password.value();
                    }
                }
            }
        }
    }

    if (_password.empty())
    {
        _password = _defaultPassword;
    }

    ClearIndexLocked();
    UpdateDriveInfoStringsLocked();
    return S_OK;
}

void FileSystem7z::ClearIndexLocked() noexcept
{
    _indexReady  = false;
    _indexStatus = S_OK;
    _indexedArchivePath.clear();
    _indexedPassword.clear();
    _entries.clear();
    _children.clear();
}

HRESULT FileSystem7z::EnsureIndex() noexcept
{
    std::lock_guard lock(_stateMutex);

    if (_archivePath.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
    }

    const bool matches = _indexReady && EqualsNoCase(_indexedArchivePath, _archivePath) && _indexedPassword == _password;
    if (matches)
    {
        return _indexStatus;
    }

    ClearIndexLocked();

    _indexStatus = BuildIndexLocked();
    _indexReady  = true;
    if (SUCCEEDED(_indexStatus))
    {
        _indexedArchivePath = _archivePath;
        _indexedPassword    = _password;
    }

    return _indexStatus;
}

HRESULT STDMETHODCALLTYPE FileSystem7z::ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept
{
    if (ppFilesInformation == nullptr)
    {
        return E_POINTER;
    }

    *ppFilesInformation = nullptr;

    if (path == nullptr || path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const HRESULT idxHr = EnsureIndex();
    if (FAILED(idxHr))
    {
        return idxHr;
    }

    std::vector<FilesInformation7z::Entry> entries;

    {
        std::lock_guard lock(_stateMutex);
        const std::wstring key = NormalizeInternalPath(path);
        const HRESULT hr       = GetEntriesForDirectory(key, entries);
        if (FAILED(hr))
        {
            return hr;
        }
    }

    auto infoImpl = std::unique_ptr<FilesInformation7z>(new (std::nothrow) FilesInformation7z());
    if (! infoImpl)
    {
        return E_OUTOFMEMORY;
    }

    const HRESULT buildHr = infoImpl->BuildFromEntries(std::move(entries));
    if (FAILED(buildHr))
    {
        return buildHr;
    }

    *ppFilesInformation = infoImpl.release();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem7z::GetAttributes(const wchar_t* path, unsigned long* fileAttributes) noexcept
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

    const HRESULT idxHr = EnsureIndex();
    if (FAILED(idxHr))
    {
        return idxHr;
    }

    std::lock_guard lock(_stateMutex);
    const std::wstring key = NormalizeInternalPath(path);

    if (key.empty())
    {
        *fileAttributes = FILE_ATTRIBUTE_DIRECTORY;
        return S_OK;
    }

    const auto it = _entries.find(key);
    if (it == _entries.end())
    {
        return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
    }

    *fileAttributes = it->second.isDirectory ? FILE_ATTRIBUTE_DIRECTORY : FILE_ATTRIBUTE_ARCHIVE;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem7z::CreateFileReader(const wchar_t* path, IFileReader** reader) noexcept
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

    const HRESULT idxHr = EnsureIndex();
    if (FAILED(idxHr))
    {
        return idxHr;
    }

    std::wstring archivePath;
    std::wstring password;
    uint32_t itemIndex = 0;
    uint64_t sizeBytes = 0;

    {
        std::lock_guard lock(_stateMutex);

        const std::wstring key = NormalizeInternalPath(path);
        if (key.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_DIRECTORY);
        }

        const auto it = _entries.find(key);
        if (it == _entries.end())
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }

        if (it->second.isDirectory)
        {
            return HRESULT_FROM_WIN32(ERROR_DIRECTORY);
        }

        if (! it->second.itemIndex.has_value())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        itemIndex   = it->second.itemIndex.value();
        sizeBytes   = static_cast<uint64_t>(it->second.sizeBytes);
        archivePath = _archivePath;
        password    = _password;
    }

    if (archivePath.empty())
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
    }

    return CreateSevenZipItemFileReader(std::move(archivePath), std::move(password), itemIndex, sizeBytes, reader);
}

HRESULT STDMETHODCALLTYPE FileSystem7z::CreateFileWriter([[maybe_unused]] const wchar_t* path,
                                                         [[maybe_unused]] FileSystemFlags flags,
                                                         IFileWriter** writer) noexcept
{
    if (writer == nullptr)
    {
        return E_POINTER;
    }

    *writer = nullptr;
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystem7z::GetFileBasicInformation([[maybe_unused]] const wchar_t* path, FileSystemBasicInformation* info) noexcept
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

    const HRESULT idxHr = EnsureIndex();
    if (FAILED(idxHr))
    {
        return idxHr;
    }

    std::wstring key;
    ArchiveEntry entry{};
    {
        std::lock_guard lock(_stateMutex);

        key = NormalizeInternalPath(path);
        if (key.empty())
        {
            // Root directory.
            return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        }

        const auto it = _entries.find(key);
        if (it == _entries.end())
        {
            return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
        }

        entry = it->second;
    }

    // Only file items provide meaningful basic info for cross-FS metadata propagation.
    if (entry.isDirectory)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    if (entry.lastWriteTime == 0)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }

    info->attributes     = FILE_ATTRIBUTE_NORMAL;
    info->lastWriteTime  = entry.lastWriteTime;
    info->creationTime   = entry.lastWriteTime;
    info->lastAccessTime = entry.lastWriteTime;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE FileSystem7z::SetFileBasicInformation([[maybe_unused]] const wchar_t* path,
                                                                [[maybe_unused]] const FileSystemBasicInformation* info) noexcept
{
    if (info == nullptr)
    {
        return E_POINTER;
    }

    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystem7z::GetItemProperties([[maybe_unused]] const wchar_t* path, const char** jsonUtf8) noexcept
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

    const HRESULT idxHr = EnsureIndex();
    if (FAILED(idxHr))
    {
        return idxHr;
    }

    std::wstring archivePath;
    std::wstring pluginPath;
    std::wstring name;
    bool isDirectory = false;
    std::optional<uint32_t> itemIndex;
    uint64_t sizeBytes    = 0;
    int64_t lastWriteTime = 0;

    {
        std::lock_guard lock(_stateMutex);

        archivePath = _archivePath;

        const std::wstring key = NormalizeInternalPath(path);
        if (key.empty())
        {
            pluginPath  = L"/";
            name        = L"/";
            isDirectory = true;
        }
        else
        {
            const auto it = _entries.find(key);
            if (it == _entries.end())
            {
                return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
            }

            pluginPath    = std::wstring(L"/") + key;
            name          = LeafName(key);
            isDirectory   = it->second.isDirectory;
            itemIndex     = it->second.itemIndex;
            sizeBytes     = it->second.sizeBytes;
            lastWriteTime = it->second.lastWriteTime;
        }
    }

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

    auto addField = [&](const char* key, const std::string& value)
    {
        yyjson_mut_val* field = yyjson_mut_obj(doc);
        yyjson_mut_obj_add_strcpy(doc, field, "key", key);
        yyjson_mut_obj_add_strncpy(doc, field, "value", value.data(), value.size());
        yyjson_mut_arr_add_val(fields, field);
    };

    addField("name", Utf8FromUtf16(name));
    addField("path", Utf8FromUtf16(pluginPath));
    addField("type", isDirectory ? std::string("directory") : std::string("file"));
    if (! isDirectory)
    {
        addField("sizeBytes", std::format("{}", sizeBytes));
    }
    if (lastWriteTime != 0)
    {
        addField("lastWriteTime", std::format("{}", lastWriteTime));
    }
    if (itemIndex.has_value())
    {
        addField("archiveItemIndex", std::format("{}", itemIndex.value()));
    }
    if (! archivePath.empty())
    {
        addField("archivePath", Utf8FromUtf16(archivePath));
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

HRESULT STDMETHODCALLTYPE FileSystem7z::CreateDirectory(const wchar_t* path) noexcept
{
    if (path == nullptr)
    {
        return E_POINTER;
    }

    if (path[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystem7z::GetDirectorySize(
    const wchar_t* path, FileSystemFlags flags, IFileSystemDirectorySizeCallback* callback, void* cookie, FileSystemDirectorySizeResult* result) noexcept
{
    if (path == nullptr || result == nullptr)
    {
        return E_POINTER;
    }

    result->totalBytes     = 0;
    result->fileCount      = 0;
    result->directoryCount = 0;
    result->status         = S_OK;

    HRESULT hr = EnsureIndex();
    if (FAILED(hr))
    {
        result->status = hr;
        return hr;
    }

    const std::wstring normalizedPath                = NormalizeInternalPath(path);
    const std::wstring searchPrefix                  = normalizedPath.empty() ? L"" : (normalizedPath + L"/");
    const bool recursive                             = (flags & FILESYSTEM_FLAG_RECURSIVE) != 0;
    constexpr unsigned long kProgressIntervalEntries = 100;
    constexpr ULONGLONG kProgressIntervalMs          = 200;

    uint64_t scannedEntries    = 0;
    ULONGLONG lastProgressTime = ::GetTickCount64();

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

    bool rootIsFile       = false;
    uint64_t rootFileSize = 0;
    {
        std::scoped_lock lock(_stateMutex);

        // Verify root path exists and classify directory/file root.
        if (! normalizedPath.empty())
        {
            auto it = _entries.find(normalizedPath);
            if (it == _entries.end())
            {
                result->status = HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
                return result->status;
            }
            if (! it->second.isDirectory)
            {
                rootIsFile   = true;
                rootFileSize = it->second.sizeBytes;
            }
        }

        if (rootIsFile)
        {
            // File root: nothing else to enumerate in archive index.
            result->totalBytes = rootFileSize;
            result->fileCount  = 1;
            scannedEntries     = 1;
        }
        else
        {
            for (const auto& [key, entry] : _entries)
            {
                // Skip root itself.
                if (key == normalizedPath)
                {
                    continue;
                }

                // Check if this entry is under the target path.
                bool isChild = false;
                if (normalizedPath.empty())
                {
                    isChild = true; // Root: all entries are descendants.
                }
                else if (key.size() > searchPrefix.size() && key.compare(0, searchPrefix.size(), searchPrefix) == 0)
                {
                    isChild = true;
                }

                if (! isChild)
                {
                    continue;
                }

                // For non-recursive, only count immediate children.
                if (! recursive)
                {
                    const std::wstring_view remainder(key.data() + searchPrefix.size(), key.size() - searchPrefix.size());
                    if (remainder.find(L'/') != std::wstring_view::npos)
                    {
                        continue; // Not an immediate child.
                    }
                }

                ++scannedEntries;

                if (entry.isDirectory)
                {
                    ++result->directoryCount;
                }
                else
                {
                    ++result->fileCount;
                    result->totalBytes += entry.sizeBytes;
                }

                if (! maybeReportProgress(path))
                {
                    return result->status;
                }
            }
        }
    }

    if (rootIsFile)
    {
        if (! maybeReportProgress(path))
        {
            return result->status;
        }
    }

    // Final progress report.
    if (callback != nullptr)
    {
        callback->DirectorySizeProgress(scannedEntries, result->totalBytes, result->fileCount, result->directoryCount, nullptr, cookie);
    }

    return result->status;
}

HRESULT STDMETHODCALLTYPE FileSystem7z::CopyItem(const wchar_t* /*sourcePath*/,
                                                 const wchar_t* /*destinationPath*/,
                                                 FileSystemFlags /*flags*/,
                                                 const FileSystemOptions* /*options*/,
                                                 IFileSystemCallback* /*callback*/,
                                                 void* /*cookie*/) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystem7z::MoveItem(const wchar_t* /*sourcePath*/,
                                                 const wchar_t* /*destinationPath*/,
                                                 FileSystemFlags /*flags*/,
                                                 const FileSystemOptions* /*options*/,
                                                 IFileSystemCallback* /*callback*/,
                                                 void* /*cookie*/) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystem7z::DeleteItem(
    const wchar_t* /*path*/, FileSystemFlags /*flags*/, const FileSystemOptions* /*options*/, IFileSystemCallback* /*callback*/, void* /*cookie*/) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystem7z::RenameItem(const wchar_t* /*sourcePath*/,
                                                   const wchar_t* /*destinationPath*/,
                                                   FileSystemFlags /*flags*/,
                                                   const FileSystemOptions* /*options*/,
                                                   IFileSystemCallback* /*callback*/,
                                                   void* /*cookie*/) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystem7z::CopyItems(const wchar_t* const* /*sourcePaths*/,
                                                  unsigned long /*count*/,
                                                  const wchar_t* /*destinationFolder*/,
                                                  FileSystemFlags /*flags*/,
                                                  const FileSystemOptions* /*options*/,
                                                  IFileSystemCallback* /*callback*/,
                                                  void* /*cookie*/) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystem7z::MoveItems(const wchar_t* const* /*sourcePaths*/,
                                                  unsigned long /*count*/,
                                                  const wchar_t* /*destinationFolder*/,
                                                  FileSystemFlags /*flags*/,
                                                  const FileSystemOptions* /*options*/,
                                                  IFileSystemCallback* /*callback*/,
                                                  void* /*cookie*/) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystem7z::DeleteItems(const wchar_t* const* /*paths*/,
                                                    unsigned long /*count*/,
                                                    FileSystemFlags /*flags*/,
                                                    const FileSystemOptions* /*options*/,
                                                    IFileSystemCallback* /*callback*/,
                                                    void* /*cookie*/) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystem7z::RenameItems(const FileSystemRenamePair* /*items*/,
                                                    unsigned long /*count*/,
                                                    FileSystemFlags /*flags*/,
                                                    const FileSystemOptions* /*options*/,
                                                    IFileSystemCallback* /*callback*/,
                                                    void* /*cookie*/) noexcept
{
    return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
}

HRESULT STDMETHODCALLTYPE FileSystem7z::GetCapabilities(const char** jsonUtf8) noexcept
{
    if (jsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *jsonUtf8 = kCapabilitiesJson;
    return S_OK;
}

std::wstring_view FileSystem7z::Trim(std::wstring_view text) noexcept
{
    while (! text.empty() && std::iswspace(static_cast<wint_t>(text.front())) != 0)
    {
        text.remove_prefix(1);
    }
    while (! text.empty() && std::iswspace(static_cast<wint_t>(text.back())) != 0)
    {
        text.remove_suffix(1);
    }
    return text;
}

bool FileSystem7z::EqualsNoCase(std::wstring_view a, std::wstring_view b) noexcept
{
    if (a.size() != b.size())
    {
        return false;
    }

    if (a.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return false;
    }

    const int len = static_cast<int>(a.size());
    return CompareStringOrdinal(a.data(), len, b.data(), len, TRUE) == CSTR_EQUAL;
}

std::string FileSystem7z::Utf8FromUtf16(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    if (text.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return {};
    }

    const int required = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
    if (required <= 0)
    {
        return {};
    }

    std::string result(static_cast<size_t>(required), '\0');
    const int written =
        WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), result.data(), required, nullptr, nullptr);
    if (written != required)
    {
        return {};
    }

    return result;
}

std::wstring FileSystem7z::Utf16FromUtf8OrAcp(std::string_view text) noexcept
{
    std::wstring utf8 = Utf16FromMultiByte(text, CP_UTF8, MB_ERR_INVALID_CHARS);
    if (! utf8.empty())
    {
        return utf8;
    }

    return Utf16FromMultiByte(text, CP_ACP, 0);
}

std::wstring FileSystem7z::NormalizeInternalPath(std::wstring_view path) noexcept
{
    std::wstring normalized(path);
    std::replace(normalized.begin(), normalized.end(), L'\\', L'/');

    if (normalized == L"/")
    {
        return {};
    }

    while (! normalized.empty() && normalized.front() == L'/')
    {
        normalized.erase(normalized.begin());
    }

    while (! normalized.empty() && normalized.back() == L'/')
    {
        normalized.pop_back();
    }

    return normalized;
}

std::wstring FileSystem7z::NormalizeArchiveEntryKey(std::wstring_view path) noexcept
{
    std::wstring key(path);
    key = std::wstring(Trim(key));
    std::replace(key.begin(), key.end(), L'\\', L'/');

    while (! key.empty() && (key.front() == L'/' || key.front() == L'.'))
    {
        if (key.front() == L'/')
        {
            key.erase(key.begin());
            continue;
        }
        if (key.front() == L'.')
        {
            if (key.size() >= 2 && key[1] == L'/')
            {
                key.erase(0, 2);
                continue;
            }
        }
        break;
    }

    while (! key.empty() && key.back() == L'/')
    {
        key.pop_back();
    }

    return key;
}

std::wstring FileSystem7z::ParentKey(std::wstring_view key) noexcept
{
    const size_t pos = key.find_last_of(L'/');
    if (pos == std::wstring_view::npos)
    {
        return {};
    }
    return std::wstring(key.substr(0, pos));
}

std::wstring FileSystem7z::LeafName(std::wstring_view key) noexcept
{
    const size_t pos = key.find_last_of(L'/');
    if (pos == std::wstring_view::npos)
    {
        return std::wstring(key);
    }
    return std::wstring(key.substr(pos + 1));
}

bool FileSystem7z::TryParseModifiedLocalTime(std::wstring_view text, int64_t& outFileTimeUtc) noexcept
{
    outFileTimeUtc = 0;

    text = Trim(text);
    if (text.empty())
    {
        return false;
    }

    int year   = 0;
    int month  = 0;
    int day    = 0;
    int hour   = 0;
    int minute = 0;
    int second = 0;

    if (swscanf_s(std::wstring(text).c_str(), L"%d-%d-%d %d:%d:%d", &year, &month, &day, &hour, &minute, &second) != 6)
    {
        return false;
    }

    SYSTEMTIME local{};
    local.wYear   = static_cast<WORD>(year);
    local.wMonth  = static_cast<WORD>(month);
    local.wDay    = static_cast<WORD>(day);
    local.wHour   = static_cast<WORD>(hour);
    local.wMinute = static_cast<WORD>(minute);
    local.wSecond = static_cast<WORD>(second);

    SYSTEMTIME utc{};
    if (TzSpecificLocalTimeToSystemTime(nullptr, &local, &utc) == 0)
    {
        utc = local;
    }

    FILETIME ft{};
    if (SystemTimeToFileTime(&utc, &ft) == 0)
    {
        return false;
    }

    ULARGE_INTEGER uli{};
    uli.LowPart    = ft.dwLowDateTime;
    uli.HighPart   = ft.dwHighDateTime;
    outFileTimeUtc = static_cast<int64_t>(uli.QuadPart);
    return true;
}

void FileSystem7z::UpdateDriveInfoStringsLocked() noexcept
{
    _driveFileSystem  = L"7z";
    _driveDisplayName = _archivePath.empty() ? std::wstring(L"7z") : _archivePath;
}

HRESULT FileSystem7z::GetEntriesForDirectory(std::wstring_view dirKey, std::vector<FilesInformation7z::Entry>& out) const noexcept
{
    out.clear();

    if (! dirKey.empty())
    {
        const auto it = _entries.find(std::wstring(dirKey));
        if (it == _entries.end() || ! it->second.isDirectory)
        {
            return HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND);
        }
    }

    const auto it = _children.find(std::wstring(dirKey));
    if (it == _children.end())
    {
        return S_OK;
    }

    std::vector<std::wstring> children = it->second;
    std::sort(children.begin(), children.end());
    children.erase(std::unique(children.begin(), children.end()), children.end());

    out.reserve(children.size());

    for (const auto& childKey : children)
    {
        const auto entryIt = _entries.find(childKey);
        if (entryIt == _entries.end())
        {
            continue;
        }

        FilesInformation7z::Entry e{};
        e.name          = LeafName(childKey);
        e.attributes    = entryIt->second.isDirectory ? FILE_ATTRIBUTE_DIRECTORY : FILE_ATTRIBUTE_ARCHIVE;
        e.sizeBytes     = entryIt->second.sizeBytes;
        e.lastWriteTime = entryIt->second.lastWriteTime;
        out.emplace_back(std::move(e));
    }

    return S_OK;
}

namespace
{
std::wstring GetModuleFileNameString(HMODULE module) noexcept
{
    std::wstring buffer;
    buffer.resize(MAX_PATH);

    for (;;)
    {
        const DWORD length = GetModuleFileNameW(module, buffer.data(), static_cast<DWORD>(buffer.size()));
        if (length == 0)
        {
            return {};
        }

        if (length < buffer.size())
        {
            buffer.resize(length);
            return buffer;
        }

        if (buffer.size() >= 32768u)
        {
            return {};
        }

        buffer.resize(buffer.size() * 2u);
    }
}

std::filesystem::path GetModuleDirectory(HMODULE module) noexcept
{
    const std::wstring pathText = GetModuleFileNameString(module);
    if (pathText.empty())
    {
        return {};
    }
    return std::filesystem::path(pathText).parent_path();
}

std::filesystem::path GetThisModuleDirectory() noexcept
{
    HMODULE module = nullptr;
    const BOOL ok  = GetModuleHandleExW(
        GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT, reinterpret_cast<LPCWSTR>(&GetThisModuleDirectory), &module);
    if (ok == 0 || module == nullptr)
    {
        return {};
    }
    return GetModuleDirectory(module);
}

struct SevenZipExports
{
    wil::unique_hmodule module;
    Func_CreateObject createObject               = nullptr;
    Func_GetNumberOfFormats getNumberOfFormats   = nullptr;
    Func_GetHandlerProperty2 getHandlerProperty2 = nullptr;
};

HRESULT LoadSevenZipExports(SevenZipExports& out) noexcept
{
    if (out.module && out.createObject && out.getNumberOfFormats && out.getHandlerProperty2)
    {
        return S_OK;
    }

    out.module.reset();
    out.createObject        = nullptr;
    out.getNumberOfFormats  = nullptr;
    out.getHandlerProperty2 = nullptr;

    const std::filesystem::path moduleDir = GetThisModuleDirectory();
    if (moduleDir.empty())
    {
        Debug::Error(L"Failed to determine module directory for locating 7zip.dll.");
        return HRESULT_FROM_WIN32(ERROR_NOT_FOUND);
    }

    const std::filesystem::path pluginDir(moduleDir / L"7zip.dll");

    wil::unique_hmodule module(LoadLibrary(pluginDir.c_str()));
    if (! module)
    {
        auto lastError = Debug::ErrorWithLastError(L"7zip.dll not found in plugin directory");
        return HRESULT_FROM_WIN32(lastError);
    }

#pragma warning(push)
#pragma warning(disable : 4191) // unsafe conversion from FARPROC
    const auto createObject        = reinterpret_cast<Func_CreateObject>(GetProcAddress(module.get(), "CreateObject"));
    const auto getNumberOfFormats  = reinterpret_cast<Func_GetNumberOfFormats>(GetProcAddress(module.get(), "GetNumberOfFormats"));
    const auto getHandlerProperty2 = reinterpret_cast<Func_GetHandlerProperty2>(GetProcAddress(module.get(), "GetHandlerProperty2"));
#pragma warning(pop)

    if (! createObject || ! getNumberOfFormats || ! getHandlerProperty2)
    {
        Debug::Error(L"7zip.dll is missing required exports.");
        return HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND);
    }

    out.module              = std::move(module);
    out.createObject        = createObject;
    out.getNumberOfFormats  = getNumberOfFormats;
    out.getHandlerProperty2 = getHandlerProperty2;
    return S_OK;
}

class SevenZipLibrary final
{
public:
    HRESULT EnsureLoaded() noexcept
    {
        std::lock_guard lock(_mutex);
        if (_loaded)
        {
            return _loadStatus;
        }

        _loadStatus = LoadSevenZipExports(_exports);
        _loaded     = true;
        return _loadStatus;
    }

    const SevenZipExports& Exports() const noexcept
    {
        return _exports;
    }

private:
    std::mutex _mutex;
    bool _loaded        = false;
    HRESULT _loadStatus = E_FAIL;
    SevenZipExports _exports;
};

SevenZipLibrary& GetSevenZipLibrary() noexcept
{
    static SevenZipLibrary library;
    return library;
}

class SevenZipFileInStream final : public IInStream, public IStreamGetSize
{
public:
    static HRESULT Create(std::wstring_view path, wil::com_ptr<IInStream>& out) noexcept
    {
        out.reset();

        if (path.empty())
        {
            return E_INVALIDARG;
        }

        wil::unique_handle file(CreateFileW(
            path.data(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
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

        auto* impl = new (std::nothrow) SevenZipFileInStream(std::move(file), static_cast<uint64_t>(fileSize.QuadPart));
        if (! impl)
        {
            return E_OUTOFMEMORY;
        }

        out.attach(impl);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** outObject) noexcept override
    {
        if (outObject == nullptr)
        {
            return E_POINTER;
        }

        *outObject = nullptr;

        if (iid == IID_IUnknown)
        {
            *outObject = static_cast<IInStream*>(this);
        }
        else if (iid == IID_ISequentialInStream)
        {
            *outObject = static_cast<ISequentialInStream*>(this);
        }
        else if (iid == IID_IInStream)
        {
            *outObject = static_cast<IInStream*>(this);
        }
        else if (iid == IID_IStreamGetSize)
        {
            *outObject = static_cast<IStreamGetSize*>(this);
        }
        else
        {
            return E_NOINTERFACE;
        }

        AddRef();
        return S_OK;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG result = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (result == 0)
        {
            delete this;
        }
        return result;
    }

    HRESULT STDMETHODCALLTYPE Read(void* data, UInt32 size, UInt32* processedSize) noexcept override
    {
        if (processedSize == nullptr)
        {
            return E_POINTER;
        }

        *processedSize = 0;
        if (size == 0)
        {
            return S_OK;
        }

        DWORD bytesRead = 0;
        const BOOL ok   = ReadFile(_file.get(), data, size, &bytesRead, nullptr);
        *processedSize  = static_cast<UInt32>(bytesRead);
        if (ok == 0)
        {
            const DWORD lastError = GetLastError();
            return HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_READ_FAULT);
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Seek(Int64 offset, UInt32 seekOrigin, UInt64* newPosition) noexcept override
    {
        DWORD method = FILE_BEGIN;

        switch (seekOrigin)
        {
            case 0: method = FILE_BEGIN; break;
            case 1: method = FILE_CURRENT; break;
            case 2: method = FILE_END; break;
            default: return STG_E_INVALIDFUNCTION;
        }

        LARGE_INTEGER distance{};
        distance.QuadPart = offset;

        LARGE_INTEGER pos{};
        if (SetFilePointerEx(_file.get(), distance, &pos, method) == 0)
        {
            const DWORD lastError = GetLastError();
            return HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_SEEK);
        }

        if (newPosition)
        {
            *newPosition = static_cast<UInt64>(pos.QuadPart);
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetSize(UInt64* size) noexcept override
    {
        if (size == nullptr)
        {
            return E_POINTER;
        }

        *size = static_cast<UInt64>(_sizeBytes);
        return S_OK;
    }

private:
    SevenZipFileInStream(wil::unique_handle file, uint64_t sizeBytes) noexcept : _file(std::move(file)), _sizeBytes(sizeBytes)
    {
    }

    ~SevenZipFileInStream() = default;

    std::atomic_ulong _refCount{1};
    wil::unique_handle _file;
    uint64_t _sizeBytes = 0;
};

class SevenZipOpenCallback final : public IArchiveOpenCallback, public IArchiveOpenVolumeCallback, public ICryptoGetTextPassword, public ICryptoGetTextPassword2
{
public:
    SevenZipOpenCallback(std::wstring archivePath, std::wstring password) : _archivePath(std::move(archivePath)), _password(std::move(password))
    {
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid, void** outObject) noexcept override
    {
        if (outObject == nullptr)
        {
            return E_POINTER;
        }

        *outObject = nullptr;

        if (iid == IID_IUnknown)
        {
            *outObject = static_cast<IArchiveOpenCallback*>(this);
        }
        else if (iid == IID_IArchiveOpenCallback)
        {
            *outObject = static_cast<IArchiveOpenCallback*>(this);
        }
        else if (iid == IID_IArchiveOpenVolumeCallback)
        {
            *outObject = static_cast<IArchiveOpenVolumeCallback*>(this);
        }
        else if (iid == IID_ICryptoGetTextPassword)
        {
            *outObject = static_cast<ICryptoGetTextPassword*>(this);
        }
        else if (iid == IID_ICryptoGetTextPassword2)
        {
            *outObject = static_cast<ICryptoGetTextPassword2*>(this);
        }
        else
        {
            return E_NOINTERFACE;
        }

        AddRef();
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetProperty(PROPID propID, PROPVARIANT* value) noexcept override
    {
        if (value == nullptr)
        {
            return E_POINTER;
        }

        static_cast<void>(PropVariantClear(value));

        if (propID == kpidName)
        {
            value->vt      = VT_BSTR;
            value->bstrVal = SysAllocString(_archivePath.c_str());
            if (! value->bstrVal)
            {
                value->vt = VT_EMPTY;
                return E_OUTOFMEMORY;
            }
            return S_OK;
        }

        value->vt = VT_EMPTY;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GetStream(const wchar_t* name, IInStream** inStream) noexcept override
    {
        if (inStream == nullptr)
        {
            return E_POINTER;
        }

        *inStream = nullptr;

        if (name == nullptr || name[0] == L'\0')
        {
            return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
        }

        std::wstring volumePath;
        std::filesystem::path path(name);
        if (path.is_relative())
        {
            path = std::filesystem::path(_archivePath).parent_path() / path;
        }
        volumePath = path.wstring();

        wil::com_ptr<IInStream> stream;
        const HRESULT hr = SevenZipFileInStream::Create(volumePath, stream);
        if (FAILED(hr))
        {
            return hr;
        }

        // COM ownership transfer to the caller via the out-param.
        *inStream = stream.detach();
        return S_OK;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG result = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (result == 0)
        {
            delete this;
        }
        return result;
    }

    HRESULT STDMETHODCALLTYPE SetTotal(const UInt64* /*files*/, const UInt64* /*bytes*/) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE SetCompleted(const UInt64* /*files*/, const UInt64* /*bytes*/) noexcept override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE CryptoGetTextPassword(BSTR* password) noexcept override
    {
        if (password == nullptr)
        {
            return E_POINTER;
        }

        *password = nullptr;
        if (_password.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_PASSWORD);
        }

        const UINT length = static_cast<UINT>(std::min<size_t>(_password.size(), std::numeric_limits<UINT>::max()));
        BSTR allocated    = SysAllocStringLen(_password.data(), length);
        if (! allocated)
        {
            return E_OUTOFMEMORY;
        }

        *password = allocated;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE CryptoGetTextPassword2(Int32* passwordIsDefined, BSTR* password) noexcept override
    {
        if (passwordIsDefined == nullptr || password == nullptr)
        {
            return E_POINTER;
        }

        *password = nullptr;

        if (_password.empty())
        {
            *passwordIsDefined = 0;
            return S_OK;
        }

        *passwordIsDefined = 1;

        const UINT length = static_cast<UINT>(std::min<size_t>(_password.size(), std::numeric_limits<UINT>::max()));
        BSTR allocated    = SysAllocStringLen(_password.data(), length);
        if (! allocated)
        {
            return E_OUTOFMEMORY;
        }

        *password = allocated;
        return S_OK;
    }

private:
    ~SevenZipOpenCallback() = default;

    std::atomic_ulong _refCount{1};
    std::wstring _archivePath;
    std::wstring _password;
};

std::wstring PropVariantToWideString(const PROPVARIANT& value) noexcept
{
    if (value.vt == VT_BSTR && value.bstrVal)
    {
        const UINT length = SysStringLen(value.bstrVal);
        return std::wstring(value.bstrVal, value.bstrVal + length);
    }
    if (value.vt == VT_LPWSTR && value.pwszVal)
    {
        return std::wstring(value.pwszVal);
    }
    return {};
}

std::optional<GUID> PropVariantToGuidBinaryBstr(const PROPVARIANT& value) noexcept
{
    if (value.vt != VT_BSTR || ! value.bstrVal)
    {
        return std::nullopt;
    }

    const UINT bytes = SysStringByteLen(value.bstrVal);
    if (bytes != sizeof(GUID))
    {
        return std::nullopt;
    }

    GUID guid{};
    std::memcpy(&guid, value.bstrVal, sizeof(guid));
    return guid;
}

bool ExtensionListContains(std::wstring_view list, std::wstring_view extensionNoDotLower) noexcept
{
    if (list.empty() || extensionNoDotLower.empty())
    {
        return false;
    }

    size_t pos = 0;
    while (pos < list.size())
    {
        while (pos < list.size() && std::iswspace(static_cast<wint_t>(list[pos])) != 0)
        {
            ++pos;
        }

        const size_t start = pos;
        while (pos < list.size() && std::iswspace(static_cast<wint_t>(list[pos])) == 0)
        {
            ++pos;
        }

        if (start == pos)
        {
            break;
        }

        std::wstring_view token(list.data() + start, pos - start);
        while (! token.empty() && token.front() == L'.')
        {
            token.remove_prefix(1);
        }

        if (token.empty() || token.size() != extensionNoDotLower.size())
        {
            continue;
        }

        const int len = static_cast<int>(token.size());
        if (CompareStringOrdinal(token.data(), len, extensionNoDotLower.data(), len, TRUE) == CSTR_EQUAL)
        {
            return true;
        }
    }

    return false;
}

std::wstring GetArchiveExtensionNoDotLower(std::wstring_view archivePath) noexcept
{
    std::wstring ext = std::filesystem::path(std::wstring(archivePath)).extension().wstring();
    if (ext.empty())
    {
        return {};
    }

    if (ext.front() == L'.')
    {
        ext.erase(ext.begin());
    }

    std::transform(ext.begin(), ext.end(), ext.begin(), [](wchar_t ch) { return static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch))); });
    return ext;
}

std::optional<GUID> TryGetFormatClassIdForExtension(const SevenZipExports& api, std::wstring_view extensionNoDotLower) noexcept
{
    if (! api.getNumberOfFormats || ! api.getHandlerProperty2)
    {
        return std::nullopt;
    }

    UInt32 numFormats     = 0;
    const HRESULT countHr = api.getNumberOfFormats(&numFormats);
    if (FAILED(countHr))
    {
        return std::nullopt;
    }

    for (UInt32 i = 0; i < numFormats; ++i)
    {
        PROPVARIANT extVar{};
        PropVariantInit(&extVar);
        extVar.vt  = VT_EMPTY;
        HRESULT hr = api.getHandlerProperty2(i, NArchive::NHandlerPropID::kExtension, &extVar);
        if (FAILED(hr))
        {
            continue;
        }
        auto clearExt              = wil::scope_exit([&] { PropVariantClear(&extVar); });
        const std::wstring extList = PropVariantToWideString(extVar);

        PROPVARIANT addExtVar{};
        PropVariantInit(&addExtVar);
        addExtVar.vt         = VT_EMPTY;
        hr                   = api.getHandlerProperty2(i, NArchive::NHandlerPropID::kAddExtension, &addExtVar);
        const bool hasAddExt = SUCCEEDED(hr);
        auto clearAddExt     = wil::scope_exit(
            [&]
            {
                if (hasAddExt)
                {
                    PropVariantClear(&addExtVar);
                }
            });
        const std::wstring addExtList = hasAddExt ? PropVariantToWideString(addExtVar) : std::wstring();

        if (! ExtensionListContains(extList, extensionNoDotLower) && ! ExtensionListContains(addExtList, extensionNoDotLower))
        {
            continue;
        }

        PROPVARIANT clsidVar{};
        PropVariantInit(&clsidVar);
        clsidVar.vt = VT_EMPTY;
        hr          = api.getHandlerProperty2(i, NArchive::NHandlerPropID::kClassID, &clsidVar);
        if (FAILED(hr))
        {
            continue;
        }
        auto clearClsid = wil::scope_exit([&] { PropVariantClear(&clsidVar); });
        const auto guid = PropVariantToGuidBinaryBstr(clsidVar);
        if (guid.has_value())
        {
            return guid;
        }
    }

    return std::nullopt;
}

HRESULT CreateAndOpenArchive(const SevenZipExports& api,
                             const GUID& classId,
                             std::wstring_view archivePath,
                             std::wstring_view password,
                             wil::com_ptr<IInArchive>& outArchive,
                             wil::com_ptr<IInStream>& outStream,
                             wil::com_ptr<IArchiveOpenCallback>& outOpenCallback) noexcept
{
    outArchive.reset();
    outStream.reset();
    outOpenCallback.reset();

    if (! api.createObject)
    {
        return HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND);
    }

    wil::com_ptr<IInArchive> archive;
    HRESULT hr = api.createObject(&classId, &IID_IInArchive, archive.put_void());
    if (FAILED(hr) || ! archive)
    {
        return FAILED(hr) ? hr : E_NOINTERFACE;
    }

    wil::com_ptr<IInStream> stream;
    hr = SevenZipFileInStream::Create(archivePath, stream);
    if (FAILED(hr))
    {
        return hr;
    }

    auto* callbackImpl = new (std::nothrow) SevenZipOpenCallback(std::wstring(archivePath), std::wstring(password));
    if (! callbackImpl)
    {
        return E_OUTOFMEMORY;
    }

    wil::com_ptr<IArchiveOpenCallback> callback;
    callback.attach(callbackImpl);

    hr = archive->Open(stream.get(), nullptr, callback.get());
    if (FAILED(hr))
    {
        return hr;
    }

    outArchive      = std::move(archive);
    outStream       = std::move(stream);
    outOpenCallback = std::move(callback);
    return S_OK;
}

HRESULT OpenArchiveAuto(const SevenZipExports& api,
                        std::wstring_view archivePath,
                        std::wstring_view password,
                        wil::com_ptr<IInArchive>& outArchive,
                        wil::com_ptr<IInStream>& outStream,
                        wil::com_ptr<IArchiveOpenCallback>& outOpenCallback) noexcept
{
    outArchive.reset();
    outStream.reset();
    outOpenCallback.reset();

    const std::wstring extNoDotLower = GetArchiveExtensionNoDotLower(archivePath);
    if (! extNoDotLower.empty())
    {
        const auto clsid = TryGetFormatClassIdForExtension(api, extNoDotLower);
        if (clsid.has_value())
        {
            const HRESULT hr = CreateAndOpenArchive(api, clsid.value(), archivePath, password, outArchive, outStream, outOpenCallback);
            if (SUCCEEDED(hr))
            {
                return S_OK;
            }
        }
    }

    if (! api.getNumberOfFormats || ! api.getHandlerProperty2)
    {
        return HRESULT_FROM_WIN32(ERROR_PROC_NOT_FOUND);
    }

    UInt32 numFormats = 0;
    HRESULT hr        = api.getNumberOfFormats(&numFormats);
    if (FAILED(hr))
    {
        return hr;
    }

    HRESULT lastError = HRESULT_FROM_WIN32(ERROR_INVALID_DATA);

    for (UInt32 i = 0; i < numFormats; ++i)
    {
        PROPVARIANT clsidVar{};
        PropVariantInit(&clsidVar);
        clsidVar.vt = VT_EMPTY;

        hr = api.getHandlerProperty2(i, NArchive::NHandlerPropID::kClassID, &clsidVar);
        if (FAILED(hr))
        {
            continue;
        }

        auto clearClsid    = wil::scope_exit([&] { PropVariantClear(&clsidVar); });
        const auto classId = PropVariantToGuidBinaryBstr(clsidVar);
        if (! classId.has_value())
        {
            continue;
        }

        lastError = CreateAndOpenArchive(api, classId.value(), archivePath, password, outArchive, outStream, outOpenCallback);
        if (SUCCEEDED(lastError))
        {
            return S_OK;
        }
    }

    return lastError;
}

class SevenZipItemFileReader final : public IFileReader
{
public:
    static HRESULT Create(std::wstring archivePath, std::wstring password, uint32_t itemIndex, uint64_t sizeBytes, IFileReader** outReader) noexcept
    {
        if (outReader == nullptr)
        {
            return E_POINTER;
        }

        *outReader = nullptr;

        auto* impl = new (std::nothrow) SevenZipItemFileReader(std::move(archivePath), std::move(password), itemIndex, sizeBytes);
        if (! impl)
        {
            return E_OUTOFMEMORY;
        }

        auto cleanup = wil::scope_exit([&] { impl->Release(); });

        const HRESULT initHr = impl->Initialize();
        if (FAILED(initHr))
        {
            return initHr;
        }

        cleanup.release();
        *outReader = impl;
        return S_OK;
    }

    SevenZipItemFileReader(const SevenZipItemFileReader&)            = delete;
    SevenZipItemFileReader(SevenZipItemFileReader&&)                 = delete;
    SevenZipItemFileReader& operator=(const SevenZipItemFileReader&) = delete;
    SevenZipItemFileReader& operator=(SevenZipItemFileReader&&)      = delete;

    ~SevenZipItemFileReader() noexcept
    {
        if (_extractThread.joinable())
        {
            {
                std::lock_guard lock(_extractMutex);
                _extractStopRequested = true;
            }

            _extractCv.notify_all();
            _extractThread.join();
        }

        if (_archive)
        {
            static_cast<void>(_archive->Close());
        }
    }

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

        *sizeBytes = _fileSizeBytes;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Seek(__int64 offset, unsigned long origin, uint64_t* newPosition) noexcept override
    {
        if (newPosition == nullptr)
        {
            return E_POINTER;
        }

        *newPosition = 0;

        if (origin != FILE_BEGIN && origin != FILE_CURRENT && origin != FILE_END)
        {
            return E_INVALIDARG;
        }

        __int64 base = 0;
        if (origin == FILE_CURRENT)
        {
            if (_positionBytes > static_cast<uint64_t>((std::numeric_limits<__int64>::max)()))
            {
                return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
            }
            base = static_cast<__int64>(_positionBytes);
        }
        else if (origin == FILE_END)
        {
            if (_fileSizeBytes > static_cast<uint64_t>((std::numeric_limits<__int64>::max)()))
            {
                return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
            }
            base = static_cast<__int64>(_fileSizeBytes);
        }

        const __int64 next = base + offset;
        if (next < 0)
        {
            return HRESULT_FROM_WIN32(ERROR_NEGATIVE_SEEK);
        }

        _positionBytes = static_cast<uint64_t>(next);
        *newPosition   = _positionBytes;
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

        if (_positionBytes >= _fileSizeBytes)
        {
            if (! _terminalStatusReported && FAILED(_terminalReadStatus))
            {
                _terminalStatusReported = true;
                return _terminalReadStatus;
            }
            return S_OK;
        }

        if (_useInMemorySpool)
        {
            const uint64_t remaining = _fileSizeBytes - _positionBytes;
            const unsigned long take = (remaining > static_cast<uint64_t>(bytesToRead)) ? bytesToRead : static_cast<unsigned long>(remaining);

            const uint64_t end    = _positionBytes + static_cast<uint64_t>(take);
            const HRESULT spoolHr = EnsureSpooledUntil(end);
            if (FAILED(spoolHr))
            {
                return spoolHr;
            }

            const uint64_t available    = (_spooledBytes > _positionBytes) ? (_spooledBytes - _positionBytes) : 0;
            const unsigned long canTake = (available > static_cast<uint64_t>(take)) ? take : static_cast<unsigned long>(available);
            if (canTake == 0)
            {
                return S_OK;
            }

            if (_positionBytes > static_cast<uint64_t>(std::numeric_limits<size_t>::max()))
            {
                return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
            }

            const size_t offset = static_cast<size_t>(_positionBytes);
            memcpy(buffer, _spool.data() + offset, canTake);

            _positionBytes += static_cast<uint64_t>(canTake);
            *bytesRead = canTake;
            return S_OK;
        }

        if (_itemStream)
        {
            return ReadStreamingItemStream(buffer, bytesToRead, bytesRead);
        }

        return ReadStreamingExtractPipe(buffer, bytesToRead, bytesRead);
    }

private:
    SevenZipItemFileReader(std::wstring archivePath, std::wstring password, uint32_t itemIndex, uint64_t sizeBytes) noexcept
        : _archivePath(std::move(archivePath)),
          _password(std::move(password)),
          _itemIndex(itemIndex),
          _fileSizeBytes(sizeBytes)
    {
    }

    class SpoolOutStream final : public ISequentialOutStream
    {
    public:
        explicit SpoolOutStream(SevenZipItemFileReader* owner) noexcept : _owner(owner)
        {
        }

        SpoolOutStream(const SpoolOutStream&)            = delete;
        SpoolOutStream(SpoolOutStream&&)                 = delete;
        SpoolOutStream& operator=(const SpoolOutStream&) = delete;
        SpoolOutStream& operator=(SpoolOutStream&&)      = delete;

        HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
        {
            if (ppvObject == nullptr)
            {
                return E_POINTER;
            }

            if (riid == IID_IUnknown || riid == IID_ISequentialOutStream)
            {
                *ppvObject = static_cast<ISequentialOutStream*>(this);
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

        HRESULT STDMETHODCALLTYPE Write(const void* data, UInt32 size, UInt32* processedSize) noexcept override
        {
            if (processedSize == nullptr)
            {
                return E_POINTER;
            }

            *processedSize = 0;

            if (size == 0)
            {
                return S_OK;
            }

            if (data == nullptr)
            {
                return E_POINTER;
            }

            if (_owner == nullptr)
            {
                return E_POINTER;
            }

            return _owner->WriteExtractBytes(data, size, processedSize);
        }

    private:
        std::atomic_ulong _refCount{1};
        SevenZipItemFileReader* _owner = nullptr; // non-owning
    };

    class ExtractCallback final : public IArchiveExtractCallback, public ICryptoGetTextPassword, public ICryptoGetTextPassword2
    {
    public:
        ExtractCallback(SevenZipItemFileReader* owner, UInt32 itemIndex, const std::wstring* password) noexcept
            : _owner(owner),
              _itemIndex(itemIndex),
              _password(password)
        {
        }

        ExtractCallback(const ExtractCallback&)            = delete;
        ExtractCallback(ExtractCallback&&)                 = delete;
        ExtractCallback& operator=(const ExtractCallback&) = delete;
        ExtractCallback& operator=(ExtractCallback&&)      = delete;

        HRESULT Result() const noexcept
        {
            return OperationResultToHr(_operationResult);
        }

        HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
        {
            if (ppvObject == nullptr)
            {
                return E_POINTER;
            }

            if (riid == IID_IUnknown || riid == IID_IProgress)
            {
                *ppvObject = static_cast<IProgress*>(this);
                AddRef();
                return S_OK;
            }

            if (riid == IID_IArchiveExtractCallback)
            {
                *ppvObject = static_cast<IArchiveExtractCallback*>(this);
                AddRef();
                return S_OK;
            }

            if (riid == IID_ICryptoGetTextPassword)
            {
                *ppvObject = static_cast<ICryptoGetTextPassword*>(this);
                AddRef();
                return S_OK;
            }

            if (riid == IID_ICryptoGetTextPassword2)
            {
                *ppvObject = static_cast<ICryptoGetTextPassword2*>(this);
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

        HRESULT STDMETHODCALLTYPE SetTotal(UInt64 /*total*/) noexcept override
        {
            return S_OK;
        }

        HRESULT STDMETHODCALLTYPE SetCompleted(const UInt64* /*completeValue*/) noexcept override
        {
            return S_OK;
        }

        HRESULT STDMETHODCALLTYPE GetStream(UInt32 index, ISequentialOutStream** outStream, Int32 askExtractMode) noexcept override
        {
            if (outStream == nullptr)
            {
                return E_POINTER;
            }

            *outStream = nullptr;

            if (askExtractMode != NArchive::NExtract::NAskMode::kExtract)
            {
                return S_OK;
            }

            if (index != _itemIndex)
            {
                return S_OK;
            }

            if (_owner == nullptr)
            {
                return E_POINTER;
            }

            if (! _stream)
            {
                auto* impl = new (std::nothrow) SpoolOutStream(_owner);
                if (! impl)
                {
                    return E_OUTOFMEMORY;
                }
                _stream.attach(impl);
            }

            *outStream = _stream.get();
            _stream->AddRef();
            return S_OK;
        }

        HRESULT STDMETHODCALLTYPE PrepareOperation(Int32 /*askExtractMode*/) noexcept override
        {
            return S_OK;
        }

        HRESULT STDMETHODCALLTYPE SetOperationResult(Int32 opRes) noexcept override
        {
            _operationResult = opRes;
            return S_OK;
        }

        HRESULT STDMETHODCALLTYPE CryptoGetTextPassword(BSTR* password) noexcept override
        {
            if (password == nullptr)
            {
                return E_POINTER;
            }

            *password = nullptr;

            if (_password == nullptr || _password->empty())
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_PASSWORD);
            }

            const UINT length = static_cast<UINT>(std::min<size_t>(_password->size(), std::numeric_limits<UINT>::max()));
            BSTR allocated    = SysAllocStringLen(_password->data(), length);
            if (! allocated)
            {
                return E_OUTOFMEMORY;
            }

            *password = allocated;
            return S_OK;
        }

        HRESULT STDMETHODCALLTYPE CryptoGetTextPassword2(Int32* passwordIsDefined, BSTR* password) noexcept override
        {
            if (passwordIsDefined == nullptr || password == nullptr)
            {
                return E_POINTER;
            }

            *password = nullptr;

            if (_password == nullptr || _password->empty())
            {
                *passwordIsDefined = 0;
                return S_OK;
            }

            *passwordIsDefined = 1;

            const UINT length = static_cast<UINT>(std::min<size_t>(_password->size(), std::numeric_limits<UINT>::max()));
            BSTR allocated    = SysAllocStringLen(_password->data(), length);
            if (! allocated)
            {
                return E_OUTOFMEMORY;
            }

            *password = allocated;
            return S_OK;
        }

    private:
        static HRESULT OperationResultToHr(Int32 opRes) noexcept
        {
            switch (opRes)
            {
                case NArchive::NExtract::NOperationResult::kOK: return S_OK;
                case NArchive::NExtract::NOperationResult::kUnsupportedMethod: return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
                case NArchive::NExtract::NOperationResult::kCRCError: return HRESULT_FROM_WIN32(ERROR_CRC);
                case NArchive::NExtract::NOperationResult::kWrongPassword: return HRESULT_FROM_WIN32(ERROR_INVALID_PASSWORD);
                case NArchive::NExtract::NOperationResult::kUnavailable: return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
                case NArchive::NExtract::NOperationResult::kUnexpectedEnd: return HRESULT_FROM_WIN32(ERROR_HANDLE_EOF);
                case NArchive::NExtract::NOperationResult::kDataError:
                case NArchive::NExtract::NOperationResult::kDataAfterEnd:
                case NArchive::NExtract::NOperationResult::kIsNotArc:
                case NArchive::NExtract::NOperationResult::kHeadersError: return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
                default: return E_FAIL;
            }
        }

        std::atomic_ulong _refCount{1};
        SevenZipItemFileReader* _owner = nullptr; // non-owning
        UInt32 _itemIndex              = 0;
        const std::wstring* _password  = nullptr; // non-owning
        wil::com_ptr<ISequentialOutStream> _stream;
        Int32 _operationResult = NArchive::NExtract::NOperationResult::kOK;
    };

    HRESULT EnsureItemStreamPosition() noexcept
    {
        if (! _itemStream || ! _archiveGetStream)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }

        if (_itemStreamPositionBytes == _positionBytes)
        {
            return S_OK;
        }

        if (_itemStreamPositionBytes > _positionBytes)
        {
            _terminalReadStatus     = S_OK;
            _terminalStatusReported = false;

            _itemStream.reset();
            const HRESULT streamHr = _archiveGetStream->GetStream(static_cast<UInt32>(_itemIndex), _itemStream.put());
            if (FAILED(streamHr) || ! _itemStream)
            {
                return FAILED(streamHr) ? streamHr : E_NOINTERFACE;
            }
            _itemStreamPositionBytes = 0;
        }

        if (_itemStreamPositionBytes < _positionBytes)
        {
            constexpr size_t kChunkSize = 256u * 1024u;
            if (_scratch.size() < kChunkSize)
            {
                _scratch.resize(kChunkSize);
            }

            uint64_t skipRemaining = _positionBytes - _itemStreamPositionBytes;
            while (skipRemaining != 0)
            {
                const UInt32 request =
                    (skipRemaining > static_cast<uint64_t>(kChunkSize)) ? static_cast<UInt32>(kChunkSize) : static_cast<UInt32>(skipRemaining);

                UInt32 processed = 0;
                const HRESULT hr = _itemStream->Read(_scratch.data(), request, &processed);
                if (FAILED(hr))
                {
                    return hr;
                }

                if (processed == 0)
                {
                    return HRESULT_FROM_WIN32(ERROR_HANDLE_EOF);
                }

                _itemStreamPositionBytes += processed;
                skipRemaining -= processed;
            }
        }

        return S_OK;
    }

    HRESULT ReadStreamingItemStream(void* buffer, unsigned long bytesToRead, unsigned long* bytesRead) noexcept
    {
        if (bytesRead == nullptr)
        {
            return E_POINTER;
        }

        *bytesRead = 0;

        const HRESULT alignHr = EnsureItemStreamPosition();
        if (FAILED(alignHr))
        {
            return alignHr;
        }

        const uint64_t remaining = _fileSizeBytes - _positionBytes;
        const unsigned long take = (remaining > static_cast<uint64_t>(bytesToRead)) ? bytesToRead : static_cast<unsigned long>(remaining);

        UInt32 processed = 0;
        HRESULT hr       = _itemStream->Read(buffer, static_cast<UInt32>(take), &processed);
        if (hr == S_FALSE)
        {
            hr = S_OK;
        }

        if (FAILED(hr))
        {
            if (processed == 0)
            {
                return hr;
            }

            _terminalReadStatus     = hr;
            _terminalStatusReported = false;
            hr                      = S_OK;
        }

        if (processed == 0 && _positionBytes < _fileSizeBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_HANDLE_EOF);
        }

        _positionBytes += static_cast<uint64_t>(processed);
        _itemStreamPositionBytes += static_cast<uint64_t>(processed);
        *bytesRead = processed;
        return hr;
    }

    void ConsumePipeLocked(size_t bytes, uint8_t* outBuffer) noexcept
    {
        if (bytes == 0 || _pipe.empty())
        {
            return;
        }

        size_t remaining = bytes;
        size_t outOffset = 0;
        while (remaining != 0)
        {
            const size_t contiguous = std::min(remaining, _pipe.size() - _pipeReadIndex);
            if (outBuffer)
            {
                memcpy(outBuffer + outOffset, _pipe.data() + _pipeReadIndex, contiguous);
                outOffset += contiguous;
            }

            _pipeReadIndex = (_pipeReadIndex + contiguous) % _pipe.size();
            _pipeSizeBytes -= contiguous;
            _pipeStartOffsetBytes += static_cast<uint64_t>(contiguous);
            remaining -= contiguous;
        }
    }

    HRESULT RestartExtractStreaming(uint64_t positionBytes) noexcept
    {
        if (_extractThread.joinable())
        {
            {
                std::lock_guard lock(_extractMutex);
                _extractStopRequested = true;
            }

            _extractCv.notify_all();
            _extractThread.join();
        }

        if (_archive)
        {
            static_cast<void>(_archive->Close());
        }

        _archive.reset();
        _archiveStream.reset();
        _openCallback.reset();
        _archiveGetStream.reset();
        _itemStream.reset();
        _itemStreamPositionBytes = 0;

        SevenZipLibrary& library = GetSevenZipLibrary();
        const HRESULT loadHr     = library.EnsureLoaded();
        if (FAILED(loadHr))
        {
            return loadHr;
        }

        const SevenZipExports& api = library.Exports();

        wil::com_ptr<IInArchive> archive;
        wil::com_ptr<IInStream> stream;
        wil::com_ptr<IArchiveOpenCallback> openCallback;
        const HRESULT openHr = OpenArchiveAuto(api, _archivePath, _password, archive, stream, openCallback);
        if (FAILED(openHr))
        {
            return openHr;
        }

        {
            std::lock_guard lock(_extractMutex);
            _archive       = std::move(archive);
            _archiveStream = std::move(stream);
            _openCallback  = std::move(openCallback);

            _pipeReadIndex        = 0;
            _pipeWriteIndex       = 0;
            _pipeSizeBytes        = 0;
            _pipeStartOffsetBytes = 0;

            _extractStarted       = false;
            _extractFinished      = false;
            _extractStopRequested = false;
            _extractWantedBytes   = 0;
            _extractStatus        = S_OK;

            _terminalReadStatus     = S_OK;
            _terminalStatusReported = false;
            _positionBytes          = positionBytes;
        }

        _extractCv.notify_all();
        return S_OK;
    }

    HRESULT EnsureExtractPipeAlignedLocked(std::unique_lock<std::mutex>& lock) noexcept
    {
        while (_pipeStartOffsetBytes > _positionBytes)
        {
            const uint64_t target = _positionBytes;
            lock.unlock();
            const HRESULT restartHr = RestartExtractStreaming(target);
            lock.lock();
            if (FAILED(restartHr))
            {
                return restartHr;
            }

            if (! _archive)
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
            }

            const HRESULT startHr = StartExtractThreadIfNeededLocked();
            if (FAILED(startHr))
            {
                return startHr;
            }

            _extractCv.notify_all();
        }

        while (_pipeStartOffsetBytes < _positionBytes)
        {
            const uint64_t needSkip = _positionBytes - _pipeStartOffsetBytes;

            while (_pipeSizeBytes == 0 && ! _extractFinished && ! _extractStopRequested)
            {
                _extractCv.wait(lock);
            }

            if (_extractStopRequested)
            {
                return E_ABORT;
            }

            if (_pipeSizeBytes == 0)
            {
                return FAILED(_extractStatus) ? _extractStatus : HRESULT_FROM_WIN32(ERROR_HANDLE_EOF);
            }

            const size_t skipNow = static_cast<size_t>(std::min<uint64_t>(needSkip, static_cast<uint64_t>(_pipeSizeBytes)));
            ConsumePipeLocked(skipNow, nullptr);
            _extractCv.notify_all();
        }

        return S_OK;
    }

    HRESULT ReadStreamingExtractPipe(void* buffer, unsigned long bytesToRead, unsigned long* bytesRead) noexcept
    {
        if (bytesRead == nullptr)
        {
            return E_POINTER;
        }

        *bytesRead = 0;

        if (_pipe.empty())
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
        }

        std::unique_lock lock(_extractMutex);

        if (! _archive)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        }

        const HRESULT startHr = StartExtractThreadIfNeededLocked();
        if (FAILED(startHr))
        {
            return startHr;
        }

        _extractCv.notify_all();

        const HRESULT alignHr = EnsureExtractPipeAlignedLocked(lock);
        if (FAILED(alignHr))
        {
            return alignHr;
        }

        const uint64_t remaining      = _fileSizeBytes - _positionBytes;
        const unsigned long requested = (remaining > static_cast<uint64_t>(bytesToRead)) ? bytesToRead : static_cast<unsigned long>(remaining);

        while (_pipeSizeBytes == 0 && ! _extractFinished && ! _extractStopRequested)
        {
            _extractCv.wait(lock);
        }

        if (_extractStopRequested)
        {
            return E_ABORT;
        }

        if (_pipeSizeBytes == 0)
        {
            return FAILED(_extractStatus) ? _extractStatus : HRESULT_FROM_WIN32(ERROR_HANDLE_EOF);
        }

        const size_t take = std::min<size_t>(static_cast<size_t>(requested), _pipeSizeBytes);
        ConsumePipeLocked(take, static_cast<uint8_t*>(buffer));

        _positionBytes += static_cast<uint64_t>(take);
        *bytesRead = static_cast<unsigned long>(take);

        lock.unlock();
        _extractCv.notify_all();
        return S_OK;
    }

    HRESULT WriteExtractBytes(const void* data, UInt32 size, UInt32* processedSize) noexcept
    {
        if (processedSize == nullptr)
        {
            return E_POINTER;
        }

        *processedSize = 0;

        if (size == 0)
        {
            return S_OK;
        }

        if (data == nullptr)
        {
            return E_POINTER;
        }

        std::unique_lock lock(_extractMutex);

        if (! _useInMemorySpool)
        {
            if (_pipe.empty())
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_STATE);
            }

            const auto* src  = static_cast<const uint8_t*>(data);
            size_t remaining = size;
            size_t offset    = 0;

            while (remaining != 0)
            {
                while (_pipeSizeBytes >= _pipe.size() && ! _extractStopRequested && ! _extractFinished)
                {
                    _extractCv.wait(lock);
                }

                if (_extractStopRequested)
                {
                    return E_ABORT;
                }

                if (_extractFinished)
                {
                    break;
                }

                const size_t freeBytes = _pipe.size() - _pipeSizeBytes;
                const size_t writeNow  = std::min(remaining, freeBytes);
                if (writeNow == 0)
                {
                    continue;
                }

                const size_t first = std::min(writeNow, _pipe.size() - _pipeWriteIndex);
                memcpy(_pipe.data() + _pipeWriteIndex, src + offset, first);
                _pipeWriteIndex = (_pipeWriteIndex + first) % _pipe.size();
                _pipeSizeBytes += first;
                offset += first;
                remaining -= first;

                const size_t second = writeNow - first;
                if (second != 0)
                {
                    memcpy(_pipe.data() + _pipeWriteIndex, src + offset, second);
                    _pipeWriteIndex = (_pipeWriteIndex + second) % _pipe.size();
                    _pipeSizeBytes += second;
                    offset += second;
                    remaining -= second;
                }

                lock.unlock();
                _extractCv.notify_all();
                lock.lock();
            }

            *processedSize = static_cast<UInt32>(offset);
            lock.unlock();
            _extractCv.notify_all();
            return S_OK;
        }

        constexpr uint64_t kExtractPrefetchBytes = 256u * 1024u;

        while (! _extractStopRequested && ! _extractFinished)
        {
            uint64_t limit = std::numeric_limits<uint64_t>::max();
            if (_extractWantedBytes <= (std::numeric_limits<uint64_t>::max() - kExtractPrefetchBytes))
            {
                limit = _extractWantedBytes + kExtractPrefetchBytes;
            }

            if (_spooledBytes < limit)
            {
                break;
            }

            _extractCv.wait(lock);
        }

        if (_extractStopRequested)
        {
            return E_ABORT;
        }

        const size_t currentSize = _spool.size();
        if (currentSize > (std::numeric_limits<size_t>::max() - static_cast<size_t>(size)))
        {
            return E_OUTOFMEMORY;
        }

        const size_t targetSize = currentSize + static_cast<size_t>(size);
        _spool.resize(targetSize);

        memcpy(_spool.data() + currentSize, data, size);
        *processedSize = size;

        _spooledBytes = static_cast<uint64_t>(_spool.size());

        lock.unlock();
        _extractCv.notify_all();
        return S_OK;
    }

    HRESULT StartExtractThreadIfNeededLocked() noexcept
    {
        if (_extractStarted)
        {
            return S_OK;
        }

        if (! _archive)
        {
            return E_FAIL;
        }

        _extractStarted       = true;
        _extractFinished      = false;
        _extractStatus        = S_OK;
        _extractStopRequested = false;

        _extractThread = std::thread([this] { ExtractThreadMain(); });

        return S_OK;
    }

    void ExtractThreadMain() noexcept
    {
        HRESULT status = S_OK;

        auto* callbackImpl = new (std::nothrow) ExtractCallback(this, static_cast<UInt32>(_itemIndex), &_password);
        if (! callbackImpl)
        {
            status = E_OUTOFMEMORY;
        }
        else
        {
            wil::com_ptr<IArchiveExtractCallback> callback;
            callback.attach(callbackImpl);

            const UInt32 index      = static_cast<UInt32>(_itemIndex);
            const HRESULT extractHr = _archive->Extract(&index, 1, 0, callback.get());
            if (FAILED(extractHr))
            {
                status = extractHr;
            }
            else
            {
                status = callbackImpl->Result();
            }
        }

        {
            std::lock_guard lock(_extractMutex);
            _extractStatus   = status;
            _extractFinished = true;
            if (FAILED(status))
            {
                _terminalReadStatus     = status;
                _terminalStatusReported = false;
            }
        }

        _extractCv.notify_all();

        if (_archive)
        {
            static_cast<void>(_archive->Close());
        }
    }

    HRESULT EnsureExtractUntil(uint64_t endExclusive) noexcept
    {
        std::unique_lock lock(_extractMutex);

        if (! _archive)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_HANDLE);
        }

        if (endExclusive > _extractWantedBytes)
        {
            _extractWantedBytes = endExclusive;
        }

        const HRESULT startHr = StartExtractThreadIfNeededLocked();
        if (FAILED(startHr))
        {
            return startHr;
        }

        _extractCv.notify_all();

        while (_spooledBytes < endExclusive && ! _extractFinished)
        {
            _extractCv.wait(lock);
        }

        if (_spooledBytes >= endExclusive)
        {
            return S_OK;
        }

        return FAILED(_extractStatus) ? _extractStatus : S_OK;
    }

    HRESULT Initialize() noexcept
    {
        SevenZipLibrary& library = GetSevenZipLibrary();
        const HRESULT loadHr     = library.EnsureLoaded();
        if (FAILED(loadHr))
        {
            return loadHr;
        }

        const SevenZipExports& api = library.Exports();

        HRESULT openHr = OpenArchiveAuto(api, _archivePath, _password, _archive, _archiveStream, _openCallback);
        if (FAILED(openHr))
        {
            Debug::Error(L"FileSystem7Z: Failed to open archive: {} (0x{:08X})", _archivePath.c_str(), openHr);
            return openHr;
        }

        _terminalReadStatus      = S_OK;
        _terminalStatusReported  = false;
        _positionBytes           = 0;
        _itemStreamPositionBytes = 0;

        constexpr uint64_t kMaxInMemorySpoolBytes = 32u * 1024u * 1024u;
        _useInMemorySpool                         = (_fileSizeBytes == 0) || (_fileSizeBytes <= kMaxInMemorySpoolBytes);

        _archiveGetStream.reset();
        _itemStream.reset();

        _spool.clear();
        _spooledBytes = 0;

        constexpr uint64_t kMaxInitialReserveBytes = 4u * 1024u * 1024u;
        if (_useInMemorySpool && _fileSizeBytes != 0 && _fileSizeBytes <= static_cast<uint64_t>(std::numeric_limits<size_t>::max()))
        {
            const uint64_t reserveBytes = std::min<uint64_t>(_fileSizeBytes, kMaxInitialReserveBytes);
            _spool.reserve(static_cast<size_t>(reserveBytes));
        }

        const HRESULT qiHr = _archive->QueryInterface(IID_IInArchiveGetStream, _archiveGetStream.put_void());
        if (SUCCEEDED(qiHr) && _archiveGetStream)
        {
            const HRESULT streamHr = _archiveGetStream->GetStream(static_cast<UInt32>(_itemIndex), _itemStream.put());
            if (FAILED(streamHr))
            {
                _itemStream.reset();
            }
        }

        _pipeReadIndex        = 0;
        _pipeWriteIndex       = 0;
        _pipeSizeBytes        = 0;
        _pipeStartOffsetBytes = 0;

        if (! _useInMemorySpool && ! _itemStream)
        {
            constexpr size_t kPipeCapacityBytes = 4u * 1024u * 1024u;
            if (_pipe.size() != kPipeCapacityBytes)
            {
                _pipe.clear();
                _pipe.resize(kPipeCapacityBytes);
            }
        }

        return S_OK;
    }

    HRESULT EnsureSpooledUntil(uint64_t endExclusive) noexcept
    {
        if (_fileSizeBytes != 0 && endExclusive > _fileSizeBytes)
        {
            endExclusive = _fileSizeBytes;
        }

        if (! _itemStream)
        {
            return EnsureExtractUntil(endExclusive);
        }

        if (endExclusive <= _spooledBytes)
        {
            return S_OK;
        }

        constexpr size_t kChunkSize = 256u * 1024u;
        if (_scratch.size() < kChunkSize)
        {
            _scratch.resize(kChunkSize);
        }

        while (_spooledBytes < endExclusive)
        {
            const uint64_t remaining = endExclusive - _spooledBytes;
            const UInt32 request     = remaining > static_cast<uint64_t>(kChunkSize) ? static_cast<UInt32>(kChunkSize) : static_cast<UInt32>(remaining);

            UInt32 processed = 0;
            const HRESULT hr = _itemStream->Read(_scratch.data(), request, &processed);
            if (FAILED(hr))
            {
                return hr;
            }

            if (processed == 0)
            {
                return HRESULT_FROM_WIN32(ERROR_HANDLE_EOF);
            }

            const size_t currentSize = _spool.size();
            if (currentSize > (std::numeric_limits<size_t>::max() - static_cast<size_t>(processed)))
            {
                return E_OUTOFMEMORY;
            }

            const size_t targetSize = currentSize + static_cast<size_t>(processed);
            _spool.resize(targetSize);

            memcpy(_spool.data() + currentSize, _scratch.data(), processed);
            _spooledBytes += static_cast<uint64_t>(processed);
        }

        return S_OK;
    }

    std::atomic_ulong _refCount{1};

    std::wstring _archivePath;
    std::wstring _password;
    uint32_t _itemIndex     = 0;
    uint64_t _fileSizeBytes = 0;

    wil::com_ptr<IInArchive> _archive;
    wil::com_ptr<IInStream> _archiveStream;
    wil::com_ptr<IArchiveOpenCallback> _openCallback;
    wil::com_ptr<IInArchiveGetStream> _archiveGetStream;
    wil::com_ptr<ISequentialInStream> _itemStream;

    uint64_t _itemStreamPositionBytes = 0;

    bool _useInMemorySpool       = true;
    HRESULT _terminalReadStatus  = S_OK;
    bool _terminalStatusReported = false;

    std::vector<uint8_t> _spool;
    uint64_t _spooledBytes = 0;

    std::vector<uint8_t> _pipe;
    size_t _pipeReadIndex          = 0;
    size_t _pipeWriteIndex         = 0;
    size_t _pipeSizeBytes          = 0;
    uint64_t _pipeStartOffsetBytes = 0;

    uint64_t _positionBytes = 0;
    std::vector<uint8_t> _scratch;

    std::mutex _extractMutex;
    std::condition_variable _extractCv;
    std::thread _extractThread;
    uint64_t _extractWantedBytes = 0;
    HRESULT _extractStatus       = S_OK;
    bool _extractStarted         = false;
    bool _extractFinished        = false;
    bool _extractStopRequested   = false;
};

HRESULT
CreateSevenZipItemFileReader(std::wstring archivePath, std::wstring password, uint32_t itemIndex, uint64_t sizeBytes, IFileReader** outReader) noexcept
{
    return SevenZipItemFileReader::Create(std::move(archivePath), std::move(password), itemIndex, sizeBytes, outReader);
}

std::wstring ArchiveStringProperty(IInArchive* archive, UInt32 index, PROPID propId) noexcept
{
    if (! archive)
    {
        return {};
    }

    PROPVARIANT var{};
    PropVariantInit(&var);
    var.vt = VT_EMPTY;

    const HRESULT hr = archive->GetProperty(index, propId, &var);
    if (FAILED(hr))
    {
        return {};
    }

    auto clear = wil::scope_exit([&] { PropVariantClear(&var); });
    return PropVariantToWideString(var);
}

bool ArchiveBoolProperty(IInArchive* archive, UInt32 index, PROPID propId, bool& outValue) noexcept
{
    outValue = false;

    if (! archive)
    {
        return false;
    }

    PROPVARIANT var{};
    PropVariantInit(&var);
    var.vt = VT_EMPTY;

    const HRESULT hr = archive->GetProperty(index, propId, &var);
    if (FAILED(hr))
    {
        return false;
    }

    auto clear = wil::scope_exit([&] { PropVariantClear(&var); });

    if (var.vt == VT_BOOL)
    {
        outValue = (var.boolVal != VARIANT_FALSE);
        return true;
    }

    if (var.vt == VT_UI4)
    {
        outValue = (var.ulVal != 0);
        return true;
    }

    if (var.vt == VT_I4)
    {
        outValue = (var.lVal != 0);
        return true;
    }

    return false;
}

bool ArchiveUInt64Property(IInArchive* archive, UInt32 index, PROPID propId, uint64_t& outValue) noexcept
{
    outValue = 0;

    if (! archive)
    {
        return false;
    }

    PROPVARIANT var{};
    PropVariantInit(&var);
    var.vt = VT_EMPTY;

    const HRESULT hr = archive->GetProperty(index, propId, &var);
    if (FAILED(hr))
    {
        return false;
    }

    auto clear = wil::scope_exit([&] { PropVariantClear(&var); });

    if (var.vt == VT_UI8)
    {
        outValue = static_cast<uint64_t>(var.uhVal.QuadPart);
        return true;
    }

    if (var.vt == VT_UI4)
    {
        outValue = static_cast<uint64_t>(var.ulVal);
        return true;
    }

    if (var.vt == VT_I8 && var.hVal.QuadPart >= 0)
    {
        outValue = static_cast<uint64_t>(var.hVal.QuadPart);
        return true;
    }

    if (var.vt == VT_I4 && var.lVal >= 0)
    {
        outValue = static_cast<uint64_t>(var.lVal);
        return true;
    }

    return false;
}

bool ArchiveFileTimePropertyUtc(IInArchive* archive, UInt32 index, PROPID propId, int64_t& outFileTimeUtc) noexcept
{
    outFileTimeUtc = 0;

    if (! archive)
    {
        return false;
    }

    PROPVARIANT var{};
    PropVariantInit(&var);
    var.vt = VT_EMPTY;

    const HRESULT hr = archive->GetProperty(index, propId, &var);
    if (FAILED(hr))
    {
        return false;
    }

    auto clear = wil::scope_exit([&] { PropVariantClear(&var); });

    if (var.vt != VT_FILETIME)
    {
        return false;
    }

    ULARGE_INTEGER uli{};
    uli.LowPart    = var.filetime.dwLowDateTime;
    uli.HighPart   = var.filetime.dwHighDateTime;
    outFileTimeUtc = static_cast<int64_t>(uli.QuadPart);
    return true;
}
} // namespace

HRESULT FileSystem7z::BuildIndexLocked() noexcept
{
    const DWORD attrs = GetFileAttributesW(_archivePath.c_str());
    if (attrs == INVALID_FILE_ATTRIBUTES)
    {
        const DWORD lastError = GetLastError();
        return HRESULT_FROM_WIN32(lastError != 0 ? lastError : ERROR_FILE_NOT_FOUND);
    }

    if ((attrs & FILE_ATTRIBUTE_DIRECTORY) != 0)
    {
        return HRESULT_FROM_WIN32(ERROR_DIRECTORY);
    }

    SevenZipLibrary& library = GetSevenZipLibrary();
    const HRESULT loadHr     = library.EnsureLoaded();
    if (FAILED(loadHr))
    {
        return loadHr;
    }

    const SevenZipExports& api = library.Exports();

    wil::com_ptr<IInArchive> archive;
    wil::com_ptr<IInStream> stream;
    wil::com_ptr<IArchiveOpenCallback> openCallback;

    const HRESULT openHr = OpenArchiveAuto(api, _archivePath, _password, archive, stream, openCallback);
    if (FAILED(openHr))
    {
        return openHr;
    }

    auto closeArchive = wil::scope_exit([&] { static_cast<void>(archive->Close()); });

    UInt32 numItems = 0;
    HRESULT hr      = archive->GetNumberOfItems(&numItems);
    if (FAILED(hr))
    {
        return hr;
    }

    struct Raw
    {
        std::wstring key;
        bool isDirectory      = false;
        uint64_t sizeBytes    = 0;
        int64_t lastWriteTime = 0;
        std::optional<uint32_t> itemIndex;
    };

    std::vector<Raw> raws;
    raws.reserve(static_cast<size_t>(numItems));

    for (UInt32 i = 0; i < numItems; ++i)
    {
        std::wstring pathText = ArchiveStringProperty(archive.get(), i, kpidPath);
        if (pathText.empty())
        {
            pathText = ArchiveStringProperty(archive.get(), i, kpidName);
        }

        if (pathText.empty())
        {
            continue;
        }

        Raw raw{};
        raw.key = NormalizeArchiveEntryKey(pathText);
        if (raw.key.empty())
        {
            continue;
        }
        raw.itemIndex = static_cast<uint32_t>(i);

        bool isDir        = false;
        const bool hasDir = ArchiveBoolProperty(archive.get(), i, kpidIsDir, isDir);
        if (! hasDir)
        {
            if (! pathText.empty() && (pathText.back() == L'/' || pathText.back() == L'\\'))
            {
                isDir = true;
            }
        }
        raw.isDirectory = isDir;

        if (! isDir)
        {
            static_cast<void>(ArchiveUInt64Property(archive.get(), i, kpidSize, raw.sizeBytes));
        }

        static_cast<void>(ArchiveFileTimePropertyUtc(archive.get(), i, kpidMTime, raw.lastWriteTime));

        raws.emplace_back(std::move(raw));
    }

    _entries.clear();
    _children.clear();

    _entries.emplace(std::wstring(), ArchiveEntry{true, 0, 0, std::nullopt});

    const auto ensureDir = [&](const std::wstring& key)
    {
        if (key.empty())
        {
            return;
        }

        if (_entries.contains(key))
        {
            return;
        }

        _entries.emplace(key, ArchiveEntry{true, 0, 0, std::nullopt});

        const std::wstring parent = ParentKey(key);
        _children[parent].push_back(key);
    };

    for (const auto& raw : raws)
    {
        if (raw.key.empty())
        {
            continue;
        }

        std::wstring parent = ParentKey(raw.key);
        if (! parent.empty())
        {
            size_t start = 0;
            while (start < raw.key.size())
            {
                const size_t slash = raw.key.find(L'/', start);
                if (slash == std::wstring::npos)
                {
                    break;
                }

                const std::wstring dirKey = raw.key.substr(0, slash);
                ensureDir(dirKey);
                start = slash + 1;
            }
        }

        if (raw.isDirectory)
        {
            ensureDir(raw.key);
        }

        ArchiveEntry entry{};
        entry.isDirectory   = raw.isDirectory;
        entry.sizeBytes     = raw.isDirectory ? 0 : raw.sizeBytes;
        entry.lastWriteTime = raw.lastWriteTime;
        entry.itemIndex     = raw.itemIndex;

        _entries[raw.key] = entry;
        _children[parent].push_back(raw.key);
    }

    for (auto& [_, list] : _children)
    {
        std::sort(list.begin(), list.end());
        list.erase(std::unique(list.begin(), list.end()), list.end());
    }

    return S_OK;
}
