#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <atomic>
#include <cstddef>
#include <cstdint>
#include <condition_variable>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182)
#include <wil/com.h>
#pragma warning(pop)

#include "PlugInterfaces/DriveInfo.h"
#include "PlugInterfaces/FileSystem.h"
#include "FileSystemRouteProviderBase.h"
#include "PlugInterfaces/Host.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/NavigationMenu.h"
#include "Helpers.h"
#include "PackedFileInfoBuffer.h"

class FilesInformationGoogleDrive final : public IFilesInformation
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

    FilesInformationGoogleDrive()  = default;
    ~FilesInformationGoogleDrive() = default;

    FilesInformationGoogleDrive(const FilesInformationGoogleDrive&)            = delete;
    FilesInformationGoogleDrive(FilesInformationGoogleDrive&&)                 = delete;
    FilesInformationGoogleDrive& operator=(const FilesInformationGoogleDrive&) = delete;
    FilesInformationGoogleDrive& operator=(FilesInformationGoogleDrive&&)      = delete;

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
    std::atomic_ulong _refCount{1};
    Common::Plugins::PackedFileInfoBuffer _packedBuffer;
};

class FileSystemGoogleDrive final : public IFileSystem,
                                    public FileSystemRouteCapabilitiesBase,
                                    public IFileSystemIO,
                                    public IFileSystemDirectoryOperations,
                                    public IFileSystemAtomicWriter,
                                    public IFileSystemIdentityDelete,
                                    public IInformations,
                                    public INavigationMenu,
                                    public IDriveInfo
{
public:
    explicit FileSystemGoogleDrive(IHost* host);

    FileSystemGoogleDrive(const FileSystemGoogleDrive&)            = delete;
    FileSystemGoogleDrive(FileSystemGoogleDrive&&)                 = delete;
    FileSystemGoogleDrive& operator=(const FileSystemGoogleDrive&) = delete;
    FileSystemGoogleDrive& operator=(FileSystemGoogleDrive&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override;
    ULONG STDMETHODCALLTYPE AddRef() noexcept override;
    ULONG STDMETHODCALLTYPE Release() noexcept override;

    HRESULT STDMETHODCALLTYPE GetMetaData(const PluginMetaData** metaData) noexcept override;
    HRESULT STDMETHODCALLTYPE GetConfigurationSchema(const char** schemaJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE SetConfiguration(const char* configurationJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE GetConfiguration(const char** configurationJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE SomethingToSave(BOOL* pSomethingToSave) noexcept override;
    [[nodiscard]] static const char* StaticConfigurationSchema() noexcept;

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

    HRESULT STDMETHODCALLTYPE GetPathCapabilities(const wchar_t* path,
                                                  FileSystemOperation operation,
                                                  const char** jsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE GetTransferHints(const wchar_t* path,
                                               FileSystemOperation operationType,
                                               FileSystemTransferEndpoint endpoint,
                                               FileSystemTransferHints* hints) noexcept override;
    HRESULT STDMETHODCALLTYPE GetStorageCharacteristics(const wchar_t* path, FileSystemStorageCharacteristics* characteristics) noexcept override;

    // IFileSystemIO (R0f-GDrive)
    HRESULT STDMETHODCALLTYPE GetAttributes(const wchar_t* path, unsigned long* fileAttributes) noexcept override;
    HRESULT STDMETHODCALLTYPE CreateFileReader(const wchar_t* path, IFileReader** reader) noexcept override;
    HRESULT STDMETHODCALLTYPE CreateFileWriter(const wchar_t* path, FileSystemFlags flags, IFileWriter** writer) noexcept override;
    HRESULT STDMETHODCALLTYPE GetFileBasicInformation(const wchar_t* path, FileSystemBasicInformation* info) noexcept override;
    HRESULT STDMETHODCALLTYPE SetFileBasicInformation(const wchar_t* path, const FileSystemBasicInformation* info) noexcept override;
    HRESULT STDMETHODCALLTYPE GetItemProperties(const wchar_t* path, const char** jsonUtf8) noexcept override;

    // IFileSystemDirectoryOperations (R0f-GDrive)
    HRESULT STDMETHODCALLTYPE CreateDirectory(const wchar_t* path) noexcept override;
    HRESULT STDMETHODCALLTYPE GetDirectorySize(const wchar_t* path,
                                               FileSystemFlags flags,
                                               IFileSystemDirectorySizeCallback* callback,
                                               void* cookie,
                                               FileSystemDirectorySizeResult* result) noexcept override;

    // IFileSystemAtomicWriter (R0f-GDrive): a resumable upload publishes the file only when its last
    // chunk is acknowledged, so the host bridge may publish new names through the provider writer.
    HRESULT STDMETHODCALLTYPE SupportsAtomicWriterCommit(const wchar_t* path, FileSystemFlags flags, BOOL* supported) noexcept override;

    // IFileSystemIdentityDelete (C10): a Permanent Delete pins the Drive file id it confirmed and
    // deletes only that id.
    HRESULT STDMETHODCALLTYPE ResolveDeleteIdentity(const wchar_t* path, const FileSystemOptions* options, FileSystemDeleteIdentity* identity) noexcept override;
    HRESULT STDMETHODCALLTYPE DeleteIfIdentity(const wchar_t* path,
                                               const FileSystemDeleteIdentity* identity,
                                               FileSystemFlags flags,
                                               const FileSystemOptions* options,
                                               IFileSystemCallback* callback,
                                               void* cookie) noexcept override;

    // Drive item helpers shared with the reader and writer objects of this module (not a host surface).
    struct ResolvedConnection;
    struct GoogleItem;
    enum class AuthorizedRequestRetry
    {
        ReadOnly,
        Mutation,
    };
    HRESULT ResolveConnection(const wchar_t* path, bool acquireSecrets, ResolvedConnection& outConnection);
    HRESULT GetAccessToken(const ResolvedConnection& connection, std::wstring& accessToken);
    HRESULT PerformAuthorizedJsonGet(const ResolvedConnection& connection, std::string_view url, std::string& body);
    HRESULT PerformAuthorizedRequest(const ResolvedConnection& connection,
                                     std::string_view method,
                                     std::string_view url,
                                     const std::vector<std::string>& extraHeaders,
                                     std::string_view body,
                                     size_t maxResponseBytes,
                                     long& statusCodeOut,
                                     std::string& bodyOut,
                                     std::string* locationOut,
                                     std::string* rangeOut,
                                     AuthorizedRequestRetry retryPolicy);
    HRESULT ListChildren(const ResolvedConnection& connection, std::wstring_view parentId, std::vector<GoogleItem>& items);
    HRESULT ResolveItemByPath(const ResolvedConnection& connection, std::wstring_view canonicalPath, GoogleItem& item);
    HRESULT GetItemById(const ResolvedConnection& connection, std::wstring_view id, GoogleItem& item);
    HRESULT ResolveChildByName(const ResolvedConnection& connection, std::wstring_view parentId, std::wstring_view exposedName, GoogleItem& child);
    HRESULT ResolveParentAndLeaf(const ResolvedConnection& connection, std::wstring_view canonicalPath, GoogleItem& parent, std::wstring& leafName);
    HRESULT CreateFolderItem(const ResolvedConnection& connection, std::wstring_view parentId, std::wstring_view name, GoogleItem& created);
    HRESULT UpdateItem(const ResolvedConnection& connection,
                       std::wstring_view id,
                       std::wstring_view newName,
                       std::wstring_view addParent,
                       std::wstring_view removeParent,
                       std::optional<bool> trashed,
                       GoogleItem& updated);
    HRESULT DeleteItemPermanently(const ResolvedConnection& connection, std::wstring_view id);
    HRESULT CopyFileItem(const ResolvedConnection& connection, std::wstring_view id, std::wstring_view parentId, std::wstring_view name, GoogleItem& copy);
    HRESULT DownloadRange(const ResolvedConnection& connection, std::wstring_view id, uint64_t offset, size_t bytes, std::string& data);
    HRESULT UploadResumable(const ResolvedConnection& connection,
                            std::wstring_view existingId,
                            std::wstring_view parentId,
                            std::wstring_view name,
                            HANDLE file,
                            uint64_t sizeBytes,
                            GoogleItem& result);

    struct Settings
    {
        std::wstring defaultClientId;
        uint32_t connectTimeoutMs = 10'000;
        uint32_t requestTimeoutMs = 30'000;
        unsigned long pageSize    = 200;
    };

#if defined(_DEBUG)
    static HRESULT RunDebugSelfTests(unsigned int* passed, unsigned int* failed) noexcept;
#endif

protected:
    HRESULT BuildFileSystemRouteDescriptor(const wchar_t* path,
                                           FileSystemOperation operation,
                                           FileSystemRouteDescriptor& descriptor) noexcept override;

private:
    ~FileSystemGoogleDrive();
    [[nodiscard]] IHostAlerts* GetHostAlerts() const noexcept;
    void ShowMissingClientIdAlert() const noexcept;

    struct DriveInfoPayload;
    struct MenuEntry
    {
        NavigationMenuItemFlags flags = NAV_MENU_ITEM_FLAG_NONE;
        std::wstring label;
        std::wstring path;
        std::wstring iconPath;
        unsigned int commandId = 0;
    };

    using NavigationMenuCallbackState    = RegistrationCallbackState<INavigationMenuCallback>;
    using NavigationMenuCallbackSnapshot = NavigationMenuCallbackState::Snapshot;

    HRESULT SetConfigurationImpl(const char* configurationJsonUtf8);
    HRESULT ReadDirectoryInfoImpl(const wchar_t* path, IFilesInformation** ppFilesInformation);
    HRESULT GetDriveInfoImpl(const wchar_t* path, DriveInfo* info);
    [[nodiscard]] bool TryCaptureNavigationMenuCallback(NavigationMenuCallbackSnapshot& snapshot) noexcept;
    HRESULT InvokeNavigationMenuCallback(const NavigationMenuCallbackSnapshot& snapshot, const wchar_t* path) noexcept;

    HRESULT FetchDriveInfoPayload(const ResolvedConnection& connection, DriveInfoPayload& payload);

    PluginMetaData _metaData{};
    std::atomic_ulong _refCount{1};

    wil::com_ptr<IHostAlerts> _hostAlerts;
    wil::com_ptr<IHostConnections> _hostConnections;

    std::mutex _stateMutex;
    Settings _settings{};
    std::string _configurationJsonStorage[2] = {"{}", "{}"}; // Double-buffer to keep old pointer valid
    size_t _configurationJsonIndex           = 0;
    std::string _capabilitiesJson;
    std::string _itemPropertiesJson; // GetItemProperties result; valid until the next call

    NavigationMenuCallbackState _navigationMenuCallbackState;
    std::vector<MenuEntry> _menuEntries;
    std::vector<NavigationMenuItem> _menuEntryView;

    std::wstring _driveDisplayName;
    std::wstring _driveVolumeLabel;
    std::wstring _driveFileSystem = L"Google Drive";

    struct AccessTokenCacheEntry
    {
        AccessTokenCacheEntry() = default;
        ~AccessTokenCacheEntry()
        {
            SecureWipe::SecureClear(token);
        }
        AccessTokenCacheEntry(const AccessTokenCacheEntry&)            = delete;
        AccessTokenCacheEntry& operator=(const AccessTokenCacheEntry&) = delete;
        AccessTokenCacheEntry(AccessTokenCacheEntry&& other) noexcept
            : token(std::move(other.token)),
              expiresAtTickMs(std::exchange(other.expiresAtTickMs, 0u))
        {
            SecureWipe::SecureClear(other.token);
        }
        AccessTokenCacheEntry& operator=(AccessTokenCacheEntry&& other) noexcept
        {
            if (this != &other)
            {
                SecureWipe::SecureClear(token);
                token           = std::move(other.token);
                expiresAtTickMs = std::exchange(other.expiresAtTickMs, 0u);
                SecureWipe::SecureClear(other.token);
            }
            return *this;
        }

        std::wstring token;
        uint64_t expiresAtTickMs = 0;
    };

    std::mutex _tokenMutex;
    std::condition_variable _tokenCv;
    std::unordered_map<std::wstring, AccessTokenCacheEntry> _accessTokensByConnectionKey;
    std::unordered_set<std::wstring> _tokenRefreshesInFlight;
    std::unordered_map<std::wstring, HRESULT> _lastTokenRefreshStatus;
    uint64_t _tokenCacheGeneration = 0u;
};

[[nodiscard]] const char* GetFileSystemGoogleDriveStaticConfigurationSchema() noexcept;

namespace FileSystemGoogleDriveInternal
{
[[nodiscard]] bool CanCreateInstance() noexcept;
void BeginShutdown() noexcept;
[[nodiscard]] bool CanUnloadNow() noexcept;
// R0f-GDrive provider-owned bound: one token refresh plus one request, each bounded by the connect
// timeout and the hard request timeout (libcurl's total-time limit); transport failures are not retried.
[[nodiscard]] unsigned long DriveProviderWatchdogTimeoutMs(uint32_t connectTimeoutMs, uint32_t requestTimeoutMs) noexcept;
#if defined(_DEBUG)
[[nodiscard]] HRESULT RunDebugCurlRuntimeProbe() noexcept;
#endif
} // namespace FileSystemGoogleDriveInternal

#if defined(ENABLE_TESTS)
namespace FileSystemGoogleDriveSelfTest
{
// R0f-GDrive fixture seam (test-enabled builds only): redirect the Drive API and the token endpoint
// to a loopback origin (`http://127.0.0.1:<port>`) and resolve `/@conn:google-drive-selftest/...`
// to a synthetic connection; `enable == false` restores the service.
void ConfigureFakeDrive(const wchar_t* origin, bool enable) noexcept;
void RunDriveStalledRequestCancelSelfTests(unsigned int& passed, unsigned int& failed) noexcept;
void RunDriveZeroTimestampReplaceSelfTests(unsigned int& passed, unsigned int& failed) noexcept;
} // namespace FileSystemGoogleDriveSelfTest
#endif
