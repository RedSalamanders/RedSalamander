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
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182)
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "PlugInterfaces/DriveInfo.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Host.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/NavigationMenu.h"

enum class FileSystemMicrosoftDriveMode
{
    OneDrivePersonal,
    OneDriveBusiness,
    SharePoint,
};

class FilesInformationMicrosoftDrive final : public IFilesInformation
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

    FilesInformationMicrosoftDrive()  = default;
    ~FilesInformationMicrosoftDrive() = default;

    FilesInformationMicrosoftDrive(const FilesInformationMicrosoftDrive&)            = delete;
    FilesInformationMicrosoftDrive(FilesInformationMicrosoftDrive&&)                 = delete;
    FilesInformationMicrosoftDrive& operator=(const FilesInformationMicrosoftDrive&) = delete;
    FilesInformationMicrosoftDrive& operator=(FilesInformationMicrosoftDrive&&)      = delete;

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

class FileSystemMicrosoftDrive final : public IFileSystem,
                                       public IFileSystemIO,
                                       public IFileSystemDirectoryOperations,
                                       public IInformations,
                                       public INavigationMenu,
                                       public IDriveInfo
{
public:
    static constexpr wchar_t kDefaultClientId[] = L"90cdea53-7c21-48b0-959e-b4024209027b";

    explicit FileSystemMicrosoftDrive(FileSystemMicrosoftDriveMode mode, IHost* host);

    FileSystemMicrosoftDrive(const FileSystemMicrosoftDrive&)            = delete;
    FileSystemMicrosoftDrive(FileSystemMicrosoftDrive&&)                 = delete;
    FileSystemMicrosoftDrive& operator=(const FileSystemMicrosoftDrive&) = delete;
    FileSystemMicrosoftDrive& operator=(FileSystemMicrosoftDrive&&)      = delete;

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

    struct Settings
    {
        std::wstring clientId     = kDefaultClientId;
        uint32_t connectTimeoutMs = 10'000;
        uint32_t requestTimeoutMs = 60'000;
        uint32_t pageSize         = 200;
        uint32_t uploadChunkMiB   = 8;
    };

    // Internal helpers used by the plugin implementation TUs.
    [[nodiscard]] Settings SnapshotSettings() const noexcept;
    [[nodiscard]] IHostConnections* GetHostConnections() const noexcept;
    HRESULT AcquireAccessTokenForConnection(std::wstring_view connectionName,
                                            std::wstring_view userName,
                                            std::wstring_view authority,
                                            std::wstring_view scopeText,
                                            bool persistRefreshToken,
                                            std::string& accessTokenOut) noexcept;
    void ClearCachedAccessToken(std::wstring_view connectionName) noexcept;
    void UpdateDriveInfoSnapshot(std::wstring_view displayName, std::wstring_view volumeLabel) noexcept;
    HRESULT SetPropertiesJson(std::string jsonUtf8) noexcept;
    [[nodiscard]] FileSystemMicrosoftDriveMode Mode() const noexcept;
    [[nodiscard]] bool TryGetCachedDrive(std::wstring_view connectionName,
                                         std::wstring& siteIdOut,
                                         std::wstring& driveIdOut,
                                         std::wstring& displayNameOut,
                                         std::wstring& volumeLabelOut,
                                         std::wstring& webUrlOut) const noexcept;
    void StoreCachedDrive(std::wstring_view connectionName,
                          std::wstring_view siteId,
                          std::wstring_view driveId,
                          std::wstring_view displayName,
                          std::wstring_view volumeLabel,
                          std::wstring_view webUrl) noexcept;

private:
    ~FileSystemMicrosoftDrive();
    [[nodiscard]] IHostAlerts* GetHostAlerts() const noexcept;
    void ShowMissingClientIdAlert() const noexcept;

    struct CachedToken
    {
        std::string accessToken;
        uint64_t expiresAtTickMs = 0;
    };

    struct CachedDrive
    {
        std::wstring siteId;
        std::wstring driveId;
        std::wstring displayName;
        std::wstring volumeLabel;
        std::wstring webUrl;
    };

    static constexpr wchar_t kPluginIdOneDrivePersonal[]      = L"builtin/file-system-onedrive-personal";
    static constexpr wchar_t kPluginShortIdOneDrivePersonal[] = L"onedrivep";
    static constexpr wchar_t kPluginNameOneDrivePersonal[]    = L"OneDrive Personal";
    static constexpr wchar_t kPluginDescOneDrivePersonal[]    = L"Microsoft OneDrive Personal storage over Microsoft Graph.";

    static constexpr wchar_t kPluginIdOneDriveBusiness[]      = L"builtin/file-system-onedrive-business";
    static constexpr wchar_t kPluginShortIdOneDriveBusiness[] = L"onedriveb";
    static constexpr wchar_t kPluginNameOneDriveBusiness[]    = L"OneDrive Business";
    static constexpr wchar_t kPluginDescOneDriveBusiness[]    = L"Microsoft OneDrive for Business storage over Microsoft Graph.";

    static constexpr wchar_t kPluginIdSharePoint[]      = L"builtin/file-system-sharepoint";
    static constexpr wchar_t kPluginShortIdSharePoint[] = L"sharepoint";
    static constexpr wchar_t kPluginNameSharePoint[]    = L"SharePoint";
    static constexpr wchar_t kPluginDescSharePoint[]    = L"Microsoft SharePoint document libraries over Microsoft Graph.";

    static constexpr wchar_t kPluginAuthor[]  = L"RedSalamander";
    static constexpr wchar_t kPluginVersion[] = L"0.1";

    static constexpr char kSchemaJson[] = R"json(
{
  "version": 1,
  "title": "Microsoft Drive",
  "fields": [
    {
      "key": "clientId",
      "type": "text",
      "label": "Client ID",
      "description": "Microsoft Entra application (client) ID used for delegated Graph sign-in.",
      "default": "90cdea53-7c21-48b0-959e-b4024209027b"
    },
    {
      "key": "connectTimeoutMs",
      "type": "value",
      "label": "Connect timeout (ms)",
      "default": 10000,
      "min": 1000,
      "max": 600000
    },
    {
      "key": "requestTimeoutMs",
      "type": "value",
      "label": "Request timeout (ms)",
      "default": 60000,
      "min": 1000,
      "max": 600000
    },
    {
      "key": "pageSize",
      "type": "value",
      "label": "Directory page size",
      "default": 200,
      "min": 1,
      "max": 999
    },
    {
      "key": "uploadChunkMiB",
      "type": "value",
      "label": "Upload chunk size (MiB)",
      "default": 8,
      "min": 1,
      "max": 32
    }
  ]
}
)json";

    static constexpr char kCapabilitiesJson[] = R"json(
{
  "version": 1,
  "operations": {
    "copy": false,
    "move": true,
    "delete": true,
    "rename": true,
    "properties": true,
    "read": true,
    "write": true
  },
  "concurrency": {
    "copyMoveMax": 1,
    "deleteMax": 4,
    "deleteRecycleBinMax": 1
  },
  "crossFileSystem": {
    "export": { "copy": ["*"], "move": [] },
    "import": { "copy": ["*"], "move": [] }
  }
}
)json";

    FileSystemMicrosoftDriveMode _mode;
    PluginMetaData _metaData{};
    std::atomic_ulong _refCount{1};
    mutable std::mutex _stateMutex;
    Settings _settings{};
    std::string _configurationJson = "{}";
    std::string _propertiesJson    = "{}";
    std::wstring _driveDisplayName;
    std::wstring _driveVolumeLabel;
    std::wstring _driveFileSystem;
    wil::com_ptr<IHostAlerts> _hostAlerts;
    wil::com_ptr<IHostConnections> _hostConnections;
    std::unordered_map<std::wstring, CachedToken> _tokenCacheByConnectionName;
    std::unordered_map<std::wstring, CachedDrive> _driveCacheByConnectionName;
    INavigationMenuCallback* _navigationMenuCallback = nullptr;
    void* _navigationMenuCookie                      = nullptr;
};
