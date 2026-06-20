#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <atomic>
#include <condition_variable>
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
#pragma warning(disable : 4625 4626 5026 5027 4514 28182) // WIL headers: deleted copy/move / unused inline Helpers / Deferencing NULL Pointer
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "PlugInterfaces/DriveInfo.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Host.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/NavigationMenu.h"

namespace Aws
{
namespace S3Crt
{
class S3CrtClient;
} // namespace S3Crt
} // namespace Aws

class FileSystemS3;

namespace FileSystemS3Internal
{
struct ResolvedAwsContext;
[[nodiscard]] std::shared_ptr<Aws::S3Crt::S3CrtClient> GetS3Client(FileSystemS3& fs, const ResolvedAwsContext& ctx) noexcept;
} // namespace FileSystemS3Internal

enum class FileSystemS3Mode
{
    S3,
    S3Table,
};

namespace FileSystemS3Internal
{
struct ResolvedAwsContext;
}

class FilesInformationS3 final : public IFilesInformation
{
public:
    struct Entry
    {
        std::wstring name;
        unsigned long fileIndex  = 0;
        unsigned long attributes = 0;
        uint64_t sizeBytes       = 0;
        __int64 creationTime     = 0;
        __int64 lastAccessTime   = 0;
        __int64 lastWriteTime    = 0;
        __int64 changeTime       = 0;
    };

    FilesInformationS3()  = default;
    ~FilesInformationS3() = default;

    FilesInformationS3(const FilesInformationS3&)            = delete;
    FilesInformationS3(FilesInformationS3&&)                 = delete;
    FilesInformationS3& operator=(const FilesInformationS3&) = delete;
    FilesInformationS3& operator=(FilesInformationS3&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override;
    ULONG STDMETHODCALLTYPE AddRef() noexcept override;
    ULONG STDMETHODCALLTYPE Release() noexcept override;

    HRESULT STDMETHODCALLTYPE GetBuffer(FileInfo** ppFileInfo) noexcept override;
    HRESULT STDMETHODCALLTYPE GetBufferSize(unsigned long* pSize) noexcept override;
    HRESULT STDMETHODCALLTYPE GetAllocatedSize(unsigned long* pSize) noexcept override;
    HRESULT STDMETHODCALLTYPE GetCount(unsigned long* pCount) noexcept override;
    HRESULT STDMETHODCALLTYPE Get(unsigned long index, FileInfo** ppEntry) noexcept override;

    HRESULT BuildFromEntries(std::vector<Entry> entries) noexcept;

private:
    static size_t AlignUp(size_t value, size_t alignment) noexcept;
    static size_t ComputeEntrySizeBytes(std::wstring_view name) noexcept;
    HRESULT LocateEntry(unsigned long index, FileInfo** ppEntry) const noexcept;

    std::atomic_ulong _refCount{1};
    std::vector<std::byte> _buffer;
    unsigned long _count     = 0;
    unsigned long _usedBytes = 0;
};

class FileSystemS3 final : public IFileSystem,
                           public IFileSystemIO,
                           public IFileSystemDirectoryOperations,
                           public IFileSystemDirectoryWatch,
                           public IInformations,
                           public INavigationMenu,
                           public IDriveInfo
{
public:
    explicit FileSystemS3(FileSystemS3Mode mode, IHost* host);

    FileSystemS3(const FileSystemS3&)            = delete;
    FileSystemS3(FileSystemS3&&)                 = delete;
    FileSystemS3& operator=(const FileSystemS3&) = delete;
    FileSystemS3& operator=(FileSystemS3&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override;
    ULONG STDMETHODCALLTYPE AddRef() noexcept override;
    ULONG STDMETHODCALLTYPE Release() noexcept override;

    // IInformations
    HRESULT STDMETHODCALLTYPE GetMetaData(const PluginMetaData** metaData) noexcept override;
    HRESULT STDMETHODCALLTYPE GetConfigurationSchema(const char** schemaJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE SetConfiguration(const char* configurationJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE GetConfiguration(const char** configurationJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE SomethingToSave(BOOL* pSomethingToSave) noexcept override;
    [[nodiscard]] static const char* StaticConfigurationSchema(FileSystemS3Mode mode) noexcept;

    // INavigationMenu
    HRESULT STDMETHODCALLTYPE GetMenuItems(const NavigationMenuItem** items, unsigned int* count) noexcept override;
    HRESULT STDMETHODCALLTYPE ExecuteMenuCommand(unsigned int commandId) noexcept override;
    HRESULT STDMETHODCALLTYPE SetCallback(INavigationMenuCallback* callback, void* cookie) noexcept override;

    // IDriveInfo
    HRESULT STDMETHODCALLTYPE GetDriveInfo(const wchar_t* path, DriveInfo* info) noexcept override;
    HRESULT STDMETHODCALLTYPE GetDriveMenuItems(const wchar_t* path, const NavigationMenuItem** items, unsigned int* count) noexcept override;
    HRESULT STDMETHODCALLTYPE ExecuteDriveMenuCommand(unsigned int commandId, const wchar_t* path) noexcept override;

    // IFileSystem
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
    HRESULT STDMETHODCALLTYPE GetTransferHints(const wchar_t* path,
                                               FileSystemOperation operationType,
                                               FileSystemTransferEndpoint endpoint,
                                               FileSystemTransferHints* hints) noexcept override;
    HRESULT STDMETHODCALLTYPE GetStorageCharacteristics(const wchar_t* path, FileSystemStorageCharacteristics* characteristics) noexcept override;

    // IFileSystemIO
    HRESULT STDMETHODCALLTYPE GetAttributes(const wchar_t* path, unsigned long* fileAttributes) noexcept override;
    HRESULT STDMETHODCALLTYPE CreateFileReader(const wchar_t* path, IFileReader** reader) noexcept override;
    HRESULT STDMETHODCALLTYPE CreateFileWriter(const wchar_t* path, FileSystemFlags flags, IFileWriter** writer) noexcept override;
    HRESULT STDMETHODCALLTYPE GetFileBasicInformation(const wchar_t* path, FileSystemBasicInformation* info) noexcept override;
    HRESULT STDMETHODCALLTYPE SetFileBasicInformation(const wchar_t* path, const FileSystemBasicInformation* info) noexcept override;
    HRESULT STDMETHODCALLTYPE GetItemProperties(const wchar_t* path, const char** jsonUtf8) noexcept override;

    // IFileSystemDirectoryOperations
    HRESULT STDMETHODCALLTYPE CreateDirectory(const wchar_t* path) noexcept override;
    HRESULT STDMETHODCALLTYPE GetDirectorySize(const wchar_t* path,
                                               FileSystemFlags flags,
                                               IFileSystemDirectorySizeCallback* callback,
                                               void* cookie,
                                               FileSystemDirectorySizeResult* result) noexcept override;

    HRESULT STDMETHODCALLTYPE WatchDirectory(const wchar_t* path, IFileSystemDirectoryWatchCallback* callback, void* cookie) noexcept override;
    HRESULT STDMETHODCALLTYPE UnwatchDirectory(const wchar_t* path) noexcept override;

    struct Settings
    {
        std::wstring defaultRegion = L"us-east-1";
        std::wstring defaultEndpointOverride;
        bool useHttps                 = true;
        bool verifyTls                = true;
        bool useVirtualAddressing     = true;
        unsigned long maxKeys         = 1000;
        unsigned long maxTableResults = 1000;
        uint32_t connectTimeoutMs     = 10'000;
        uint32_t requestTimeoutMs     = 30'000;
    };

private:
    ~FileSystemS3();

    struct SyntheticWatchRegistration
    {
        SyntheticWatchRegistration()                                             = default;
        SyntheticWatchRegistration(const SyntheticWatchRegistration&)            = delete;
        SyntheticWatchRegistration& operator=(const SyntheticWatchRegistration&) = delete;
        SyntheticWatchRegistration(SyntheticWatchRegistration&&)                 = delete;
        SyntheticWatchRegistration& operator=(SyntheticWatchRegistration&&)      = delete;

        std::wstring watchedPath;
        std::wstring watchedPathKey;
        IFileSystemDirectoryWatchCallback* callback = nullptr;
        void* cookie                                = nullptr;
        std::atomic<bool> active{true};
        std::atomic<uint32_t> inFlight{0};
        std::atomic<DWORD> callbackThreadId{0};
        std::mutex drainMutex;
        std::condition_variable drainCv;
    };

    struct SyntheticWatchChange
    {
        FileSystemDirectoryChangeAction action = FILESYSTEM_DIR_CHANGE_MODIFIED;
        std::wstring relativePath;
    };

    static std::wstring MakeWatchPathKey(std::wstring_view path) noexcept;
    static std::wstring RelativeWatchPath(std::wstring_view watchedPath, std::wstring_view fullPath) noexcept;
    void EmitSyntheticWatchNotification(std::wstring_view watchedPath, const std::vector<SyntheticWatchChange>& changes, bool overflow = false) noexcept;

public:
    void NotifySyntheticPathCreated(std::wstring_view fullPath) noexcept;
    void NotifySyntheticPathDeleted(std::wstring_view fullPath) noexcept;
    void NotifySyntheticFolderChanged(std::wstring_view folderPath) noexcept;

private:
    struct MenuEntry
    {
        NavigationMenuItemFlags flags = NAV_MENU_ITEM_FLAG_NONE;
        std::wstring label;
        std::wstring path;
        std::wstring iconPath;
        unsigned int commandId = 0;
    };

    static constexpr wchar_t kPluginIdS3[]      = L"builtin/file-system-s3";
    static constexpr wchar_t kPluginShortIdS3[] = L"s3";

    static constexpr wchar_t kPluginIdS3Table[]      = L"builtin/file-system-s3table";
    static constexpr wchar_t kPluginShortIdS3Table[] = L"s3table";

    static constexpr wchar_t kPluginAuthor[]  = L"RedSalamander";
    static constexpr wchar_t kPluginVersion[] = VERSINFO_PLUGIN_VERSION;

    static constexpr char kCapabilitiesJsonS3[] = R"json(
{
  "version": 1,
  "operations": {
    "copy": true,
    "move": true,
    "delete": true,
    "rename": true,
    "properties": true,
    "read": true,
    "write": true
  },
  "concurrency": {
    "copyMoveMax": 1,
    "deleteMax": 8,
    "deleteRecycleBinMax": 1
  },
  "crossFileSystem": {
    "export": { "copy": ["*"], "move": ["*"] },
    "import": { "copy": ["*"], "move": ["*"] }
  },
  "pathIdentity": {
    "version": 1,
    "pathTextStableIdentity": true,
    "componentComparison": "ordinalCaseSensitive",
    "normalization": "none",
    "preferredSeparator": "/",
    "acceptedSeparators": ["/"],
    "casePreserving": true,
    "caseOnlyRename": "supported"
  }
}
)json";

    static constexpr char kCapabilitiesJsonS3Table[] = R"json(
{
  "version": 1,
  "operations": {
    "copy": false,
    "move": false,
    "delete": false,
    "rename": false,
    "properties": true,
    "read": true,
    "write": true
  },
  "concurrency": {
    "copyMoveMax": 1,
    "deleteMax": 1,
    "deleteRecycleBinMax": 1
  },
  "crossFileSystem": {
    "export": { "copy": ["*"], "move": [] },
    "import": { "copy": ["*"], "move": ["*"] }
  },
  "pathIdentity": {
    "version": 1,
    "pathTextStableIdentity": true,
    "componentComparison": "ordinalCaseSensitive",
    "normalization": "none",
    "preferredSeparator": "/",
    "acceptedSeparators": ["/"],
    "casePreserving": true,
    "caseOnlyRename": "notApplicable"
  }
}
)json";

    static constexpr char kSchemaJsonS3[] = R"json(
{
  "version": 1,
  "title": "S3",
  "fields": [
    {
      "key": "defaultRegion",
      "label": "Default region",
      "type": "text",
      "default": "us-east-1",
      "description": "AWS region used when no Connection Manager profile is selected."
    },
    {
      "key": "defaultEndpointOverride",
      "label": "Default endpoint override",
      "type": "text",
      "default": "",
      "description": "Optional endpoint override (for S3-compatible storage). Examples: https://s3.us-east-1.amazonaws.com, http://localhost:9000"
    },
    {
      "key": "useHttps",
      "label": "Use HTTPS",
      "type": "bool",
      "default": true
    },
    {
      "key": "verifyTls",
      "label": "Verify TLS certificate",
      "type": "bool",
      "default": true
    },
    {
      "key": "connectTimeoutMs",
      "label": "Connect timeout (ms)",
      "type": "value",
      "default": 10000,
      "description": "TCP connect timeout used for S3 requests (ms).",
      "min": 1,
      "max": 600000
    },
    {
      "key": "requestTimeoutMs",
      "label": "Network stall timeout (ms)",
      "type": "value",
      "default": 30000,
      "description": "No-progress / socket-read timeout for S3 requests (ms). Does not cap total transfer time if data keeps flowing.",
      "min": 1,
      "max": 600000
    },
    {
      "key": "useVirtualAddressing",
      "label": "Use virtual-hosted style addressing",
      "type": "bool",
      "default": true,
      "description": "When off, path-style addressing is used (often required for some S3-compatible endpoints)."
    },
    {
      "key": "maxKeys",
      "label": "Max keys per request",
      "type": "value",
      "default": 1000,
      "description": "S3 folder listings are paginated; this controls the number of keys requested per page (1..1000).",
      "min": 1,
      "max": 1000
    }
  ]
}
)json";

    static constexpr char kSchemaJsonS3Table[] = R"json(
{
  "version": 1,
  "title": "S3 Table",
  "fields": [
    {
      "key": "defaultRegion",
      "label": "Default region",
      "type": "text",
      "default": "us-east-1"
    },
    {
      "key": "defaultEndpointOverride",
      "label": "Default endpoint override",
      "type": "text",
      "default": ""
    },
    {
      "key": "useHttps",
      "label": "Use HTTPS",
      "type": "bool",
      "default": true
    },
    {
      "key": "verifyTls",
      "label": "Verify TLS certificate",
      "type": "bool",
      "default": true
    },
    {
      "key": "connectTimeoutMs",
      "label": "Connect timeout (ms)",
      "type": "value",
      "default": 10000,
      "description": "TCP connect timeout used for S3 requests (ms).",
      "min": 1,
      "max": 600000
    },
    {
      "key": "requestTimeoutMs",
      "label": "Network stall timeout (ms)",
      "type": "value",
      "default": 30000,
      "description": "No-progress / socket-read timeout for S3 requests (ms). Does not cap total transfer time if data keeps flowing.",
      "min": 1,
      "max": 600000
    },
    {
      "key": "maxTableResults",
      "label": "Max results per request",
      "type": "value",
      "default": 1000,
      "description": "S3 Table listings are paginated; this controls the number of results requested per page (1..1000).",
      "min": 1,
      "max": 1000
    }
  ]
}
)json";

    FileSystemS3Mode _mode = FileSystemS3Mode::S3;
    PluginMetaData _metaData{};

    std::atomic_ulong _refCount{1};

    wil::com_ptr<IHostConnections> _hostConnections;

    std::mutex _stateMutex;
    Settings _settings{};
    std::string _configurationJson = "{}";

    std::mutex _propertiesMutex;
    std::string _lastPropertiesJson;

    // Navigation menu state
    INavigationMenuCallback* _navigationMenuCallback = nullptr;
    void* _navigationMenuCallbackCookie              = nullptr;
    std::vector<MenuEntry> _menuEntries;
    std::vector<NavigationMenuItem> _menuEntryView;

    std::mutex _watchMutex;
    std::vector<std::shared_ptr<SyntheticWatchRegistration>> _syntheticWatches;

    // Drive info
    std::wstring _driveDisplayName;
    std::wstring _driveFileSystem;

    // S3 cache (best-effort)
    std::unordered_map<std::wstring, std::string> _s3BucketRegionByName;
    std::unordered_map<std::string, std::shared_ptr<Aws::S3Crt::S3CrtClient>> _s3ClientsByCtxKey;

    // S3 Tables cache (best-effort)
    std::unordered_map<std::wstring, std::string> _s3TableBucketArnByName;

    friend std::optional<std::string> LookupS3BucketRegion(FileSystemS3& fs, std::wstring_view bucketName) noexcept;
    friend void SetS3BucketRegion(FileSystemS3& fs, std::wstring_view bucketName, std::string region) noexcept;

    friend HRESULT ListS3TableBuckets(FileSystemS3& fs,
                                      const FileSystemS3Internal::ResolvedAwsContext& ctx,
                                      std::vector<FilesInformationS3::Entry>& out) noexcept;
    friend std::optional<std::string> LookupS3TableBucketArn(FileSystemS3& fs, std::wstring_view bucketName) noexcept;
    friend std::shared_ptr<Aws::S3Crt::S3CrtClient> FileSystemS3Internal::GetS3Client(FileSystemS3& fs,
                                                                                      const FileSystemS3Internal::ResolvedAwsContext& ctx) noexcept;
};

[[nodiscard]] const char* GetFileSystemS3StaticConfigurationSchema(FileSystemS3Mode mode) noexcept;
