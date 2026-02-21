#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <atomic>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <mutex>
#include <new>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182) // WIL headers: deleted copy/move and unreferenced inline Helpers
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "PlugInterfaces/DriveInfo.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Informations.h"

class FilesInformation7z final : public IFilesInformation
{
public:
    FilesInformation7z()  = default;
    ~FilesInformation7z() = default;

    FilesInformation7z(const FilesInformation7z&)            = delete;
    FilesInformation7z(FilesInformation7z&&)                 = delete;
    FilesInformation7z& operator=(const FilesInformation7z&) = delete;
    FilesInformation7z& operator=(FilesInformation7z&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override;
    ULONG STDMETHODCALLTYPE AddRef() noexcept override;
    ULONG STDMETHODCALLTYPE Release() noexcept override;

    HRESULT STDMETHODCALLTYPE GetBuffer(FileInfo** ppFileInfo) noexcept override;
    HRESULT STDMETHODCALLTYPE GetBufferSize(unsigned long* pSize) noexcept override;
    HRESULT STDMETHODCALLTYPE GetAllocatedSize(unsigned long* pSize) noexcept override;
    HRESULT STDMETHODCALLTYPE GetCount(unsigned long* pCount) noexcept override;
    HRESULT STDMETHODCALLTYPE Get(unsigned long index, FileInfo** ppEntry) noexcept override;

    struct Entry
    {
        std::wstring name;
        DWORD attributes      = 0;
        uint64_t sizeBytes    = 0;
        int64_t lastWriteTime = 0;
    };

    HRESULT BuildFromEntries(std::vector<Entry> entries) noexcept;

private:
    static size_t ComputeEntrySizeBytes(std::wstring_view name) noexcept;
    static size_t AlignUp(size_t value, size_t alignment) noexcept;

    HRESULT LocateEntry(unsigned long index, FileInfo** ppEntry) const noexcept;

    std::atomic_ulong _refCount{1};
    std::vector<std::byte> _buffer;
    unsigned long _count     = 0;
    unsigned long _usedBytes = 0;
};

class FileSystem7z final : public IFileSystem,
                           public IFileSystemIO,
                           public IFileSystemDirectoryOperations,
                           public IInformations,
                           public INavigationMenu,
                           public IDriveInfo,
                           public IFileSystemInitialize
{
public:
    FileSystem7z();

    FileSystem7z(const FileSystem7z&)            = delete;
    FileSystem7z(FileSystem7z&&)                 = delete;
    FileSystem7z& operator=(const FileSystem7z&) = delete;
    FileSystem7z& operator=(FileSystem7z&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override;
    ULONG STDMETHODCALLTYPE AddRef() noexcept override;
    ULONG STDMETHODCALLTYPE Release() noexcept override;

    HRESULT STDMETHODCALLTYPE GetMetaData(const PluginMetaData** metaData) noexcept override;
    HRESULT STDMETHODCALLTYPE GetConfigurationSchema(const char** schemaJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE SetConfiguration(const char* configurationJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE GetConfiguration(const char** configurationJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE SomethingToSave(BOOL* pSomethingToSave) noexcept override;

    HRESULT STDMETHODCALLTYPE GetMenuItems(const NavigationMenuItem** items, unsigned int* count) noexcept override;
    HRESULT STDMETHODCALLTYPE ExecuteMenuCommand(unsigned int commandId) noexcept override;
    HRESULT STDMETHODCALLTYPE SetCallback(INavigationMenuCallback* callback, void* cookie) noexcept override;

    HRESULT STDMETHODCALLTYPE GetDriveInfo(const wchar_t* path, DriveInfo* info) noexcept override;
    HRESULT STDMETHODCALLTYPE GetDriveMenuItems(const wchar_t* path, const NavigationMenuItem** items, unsigned int* count) noexcept override;
    HRESULT STDMETHODCALLTYPE ExecuteDriveMenuCommand(unsigned int commandId, const wchar_t* path) noexcept override;

    HRESULT STDMETHODCALLTYPE Initialize(const wchar_t* rootPath, const char* optionsJsonUtf8 = nullptr) noexcept override;

    HRESULT STDMETHODCALLTYPE ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept override;
    HRESULT STDMETHODCALLTYPE CopyItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options = nullptr,
                                       IFileSystemCallback* callback    = nullptr,
                                       void* cookie                     = nullptr) noexcept override;
    HRESULT STDMETHODCALLTYPE MoveItem(const wchar_t* sourcePath,
                                       const wchar_t* destinationPath,
                                       FileSystemFlags flags,
                                       const FileSystemOptions* options = nullptr,
                                       IFileSystemCallback* callback    = nullptr,
                                       void* cookie                     = nullptr) noexcept override;
    HRESULT STDMETHODCALLTYPE DeleteItem(const wchar_t* path,
                                         FileSystemFlags flags,
                                         const FileSystemOptions* options = nullptr,
                                         IFileSystemCallback* callback    = nullptr,
                                         void* cookie                     = nullptr) noexcept override;
    HRESULT STDMETHODCALLTYPE RenameItem(const wchar_t* sourcePath,
                                         const wchar_t* destinationPath,
                                         FileSystemFlags flags,
                                         const FileSystemOptions* options = nullptr,
                                         IFileSystemCallback* callback    = nullptr,
                                         void* cookie                     = nullptr) noexcept override;
    HRESULT STDMETHODCALLTYPE CopyItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options = nullptr,
                                        IFileSystemCallback* callback    = nullptr,
                                        void* cookie                     = nullptr) noexcept override;
    HRESULT STDMETHODCALLTYPE MoveItems(const wchar_t* const* sourcePaths,
                                        unsigned long count,
                                        const wchar_t* destinationFolder,
                                        FileSystemFlags flags,
                                        const FileSystemOptions* options = nullptr,
                                        IFileSystemCallback* callback    = nullptr,
                                        void* cookie                     = nullptr) noexcept override;
    HRESULT STDMETHODCALLTYPE DeleteItems(const wchar_t* const* paths,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options = nullptr,
                                          IFileSystemCallback* callback    = nullptr,
                                          void* cookie                     = nullptr) noexcept override;
    HRESULT STDMETHODCALLTYPE RenameItems(const FileSystemRenamePair* items,
                                          unsigned long count,
                                          FileSystemFlags flags,
                                          const FileSystemOptions* options = nullptr,
                                          IFileSystemCallback* callback    = nullptr,
                                          void* cookie                     = nullptr) noexcept override;

    HRESULT STDMETHODCALLTYPE GetCapabilities(const char** jsonUtf8) noexcept override;

    HRESULT STDMETHODCALLTYPE GetAttributes(const wchar_t* path, unsigned long* fileAttributes) noexcept override;
    HRESULT STDMETHODCALLTYPE CreateFileReader(const wchar_t* path, IFileReader** reader) noexcept override;
    HRESULT STDMETHODCALLTYPE CreateFileWriter(const wchar_t* path, FileSystemFlags flags, IFileWriter** writer) noexcept override;
    HRESULT STDMETHODCALLTYPE GetFileBasicInformation(const wchar_t* path, FileSystemBasicInformation* info) noexcept override;
    HRESULT STDMETHODCALLTYPE SetFileBasicInformation(const wchar_t* path, const FileSystemBasicInformation* info) noexcept override;
    HRESULT STDMETHODCALLTYPE GetItemProperties(const wchar_t* path, const char** jsonUtf8) noexcept override;

    HRESULT STDMETHODCALLTYPE CreateDirectory(const wchar_t* path) noexcept override;
    HRESULT STDMETHODCALLTYPE GetDirectorySize(const wchar_t* path,
                                               FileSystemFlags flags,
                                               IFileSystemDirectorySizeCallback* callback,
                                               void* cookie,
                                               FileSystemDirectorySizeResult* result) noexcept override;

private:
    ~FileSystem7z() = default;

    static constexpr wchar_t kPluginId[]          = L"builtin/file-system-7z";
    static constexpr wchar_t kPluginShortId[]     = L"7z";
    static constexpr wchar_t kPluginName[]        = L"7-Zip";
    static constexpr wchar_t kPluginDescription[] = L"Browse archive files as a virtual file system.";
    static constexpr wchar_t kPluginAuthor[]      = L"RedSalamander";
    static constexpr wchar_t kPluginVersion[]     = L"0.1";

    static constexpr char kCapabilitiesJson[] = R"json(
{
  "version": 1,
  "operations": {
    "copy": false,
    "move": false,
    "delete": false,
    "rename": false,
    "properties": true,
    "read": true,
    "write": false
  },
  "concurrency": {
    "copyMoveMax": 1,
    "deleteMax": 1,
    "deleteRecycleBinMax": 1
  },
  "crossFileSystem": {
    "export": { "copy": ["*"], "move": [] },
    "import": { "copy": [], "move": [] }
  }
}
)json";

    static constexpr char kSchemaJson[] = R"json(
{
  "version": 1,
  "title": "7-Zip",
  "fields": [
    {
      "key": "defaultPassword",
      "label": "Default password",
      "type": "text",
      "default": "",
      "description": "Optional password used when listing encrypted archives (stored in settings as plain text)."
    }
  ]
}
)json";

    struct ArchiveEntry
    {
        bool isDirectory      = false;
        uint64_t sizeBytes    = 0;
        int64_t lastWriteTime = 0;
        std::optional<uint32_t> itemIndex;
    };

    HRESULT EnsureIndex() noexcept;
    void ClearIndexLocked() noexcept;

    HRESULT BuildIndexLocked() noexcept;

    static std::wstring NormalizeInternalPath(std::wstring_view path) noexcept;
    static std::wstring NormalizeArchiveEntryKey(std::wstring_view path) noexcept;
    static std::wstring ParentKey(std::wstring_view key) noexcept;
    static std::wstring LeafName(std::wstring_view key) noexcept;
    static bool TryParseModifiedLocalTime(std::wstring_view text, int64_t& outFileTimeUtc) noexcept;

    static std::wstring_view Trim(std::wstring_view text) noexcept;
    static std::wstring Utf16FromUtf8OrAcp(std::string_view text) noexcept;
    static std::string Utf8FromUtf16(std::wstring_view text) noexcept;
    static bool EqualsNoCase(std::wstring_view a, std::wstring_view b) noexcept;

    void UpdateDriveInfoStringsLocked() noexcept;
    HRESULT GetEntriesForDirectory(std::wstring_view dirKey, std::vector<FilesInformation7z::Entry>& out) const noexcept;

    std::atomic_ulong _refCount{1};

    PluginMetaData _metaData{};
    std::string _configurationJson;
    std::wstring _defaultPassword;

    std::mutex _stateMutex;
    std::mutex _propertiesMutex;
    std::string _lastPropertiesJson;
    std::wstring _archivePath;
    std::wstring _password;

    bool _indexReady     = false;
    HRESULT _indexStatus = S_OK;
    std::wstring _indexedArchivePath;
    std::wstring _indexedPassword;

    // Key format: forward-slash-separated, no leading slash. Root is "".
    std::unordered_map<std::wstring, ArchiveEntry> _entries;
    std::unordered_map<std::wstring, std::vector<std::wstring>> _children;

    // DriveInfo string storage.
    std::wstring _driveDisplayName;
    std::wstring _driveVolumeLabel;
    std::wstring _driveFileSystem;
    DriveInfo _driveInfo{};

    struct MenuEntry
    {
        NavigationMenuItemFlags flags = NAV_MENU_ITEM_FLAG_NONE;
        std::wstring label;
        std::wstring path;
        std::wstring iconPath;
        unsigned int commandId = 0;
    };

    std::vector<MenuEntry> _menuEntries;
    std::vector<NavigationMenuItem> _menuEntryView;
    INavigationMenuCallback* _navigationMenuCallback = nullptr;
    void* _navigationMenuCallbackCookie              = nullptr;
};
