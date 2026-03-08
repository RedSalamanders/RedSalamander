#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <atomic>
#include <cstddef>
#include <cstdint>
#include <mutex>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182)
#include <wil/com.h>
#pragma warning(pop)

#include "PlugInterfaces/DriveInfo.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Host.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/NavigationMenu.h"

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
    static size_t AlignUp(size_t value, size_t alignment) noexcept;
    static size_t ComputeEntrySizeBytes(std::wstring_view name) noexcept;
    HRESULT LocateEntry(unsigned long index, FileInfo** ppEntry) const noexcept;

    std::atomic_ulong _refCount{1};
    std::vector<std::byte> _buffer;
    unsigned long _count     = 0;
    unsigned long _usedBytes = 0;
};

class FileSystemGoogleDrive final : public IFileSystem, public IInformations, public INavigationMenu, public IDriveInfo
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

    struct Settings
    {
        std::wstring defaultClientId;
        uint32_t connectTimeoutMs = 10'000;
        uint32_t requestTimeoutMs = 30'000;
        unsigned long pageSize    = 200;
    };

private:
    ~FileSystemGoogleDrive();

    struct ResolvedConnection;
    struct GoogleItem;
    struct DriveInfoPayload;
    struct MenuEntry
    {
        NavigationMenuItemFlags flags = NAV_MENU_ITEM_FLAG_NONE;
        std::wstring label;
        std::wstring path;
        std::wstring iconPath;
        unsigned int commandId = 0;
    };

    HRESULT SetConfigurationImpl(const char* configurationJsonUtf8);
    HRESULT ReadDirectoryInfoImpl(const wchar_t* path, IFilesInformation** ppFilesInformation);
    HRESULT GetDriveInfoImpl(const wchar_t* path, DriveInfo* info);

    HRESULT ResolveConnection(const wchar_t* path, bool acquireSecrets, ResolvedConnection& outConnection);
    HRESULT GetAccessToken(const ResolvedConnection& connection, std::wstring& accessToken);
    HRESULT PerformAuthorizedJsonGet(const ResolvedConnection& connection, std::string_view url, std::string& body);
    HRESULT ListChildren(const ResolvedConnection& connection, std::wstring_view parentId, std::vector<GoogleItem>& items);
    HRESULT ResolveItemByPath(const ResolvedConnection& connection, std::wstring_view canonicalPath, GoogleItem& item);
    HRESULT FetchDriveInfoPayload(const ResolvedConnection& connection, DriveInfoPayload& payload);

    PluginMetaData _metaData{};
    std::atomic_ulong _refCount{1};

    wil::com_ptr<IHostConnections> _hostConnections;

    std::mutex _stateMutex;
    Settings _settings{};
    std::string _configurationJson = "{}";

    INavigationMenuCallback* _navigationMenuCallback = nullptr;
    void* _navigationMenuCallbackCookie              = nullptr;
    std::vector<MenuEntry> _menuEntries;
    std::vector<NavigationMenuItem> _menuEntryView;

    std::wstring _driveDisplayName;
    std::wstring _driveVolumeLabel;
    std::wstring _driveFileSystem = L"Google Drive";

    struct AccessTokenCacheEntry
    {
        std::wstring token;
        uint64_t expiresAtTickMs = 0;
    };

    std::mutex _tokenMutex;
    std::unordered_map<std::wstring, AccessTokenCacheEntry> _accessTokensByConnectionKey;
};
