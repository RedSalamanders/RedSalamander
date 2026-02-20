#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <atomic>
#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514) // WIL headers: deleted copy/move and unused inline Helpers
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "PlugInterfaces/DriveInfo.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/NavigationMenu.h"

struct IHost;
struct IHostConnections;

enum class FileSystemCurlProtocol
{
    Ftp,
    Sftp,
    Scp,
    Imap,
};

class FilesInformationCurl final : public IFilesInformation
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

    FilesInformationCurl()  = default;
    ~FilesInformationCurl() = default;

    FilesInformationCurl(const FilesInformationCurl&)            = delete;
    FilesInformationCurl(FilesInformationCurl&&)                 = delete;
    FilesInformationCurl& operator=(const FilesInformationCurl&) = delete;
    FilesInformationCurl& operator=(FilesInformationCurl&&)      = delete;

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

class FileSystemCurl final : public IFileSystem,
                             public IFileSystemIO,
                             public IFileSystemDirectoryOperations,
                             public IInformations,
                             public INavigationMenu,
                             public IDriveInfo
{
public:
    explicit FileSystemCurl(FileSystemCurlProtocol protocol, IHost* host);

    FileSystemCurl(const FileSystemCurl&)            = delete;
    FileSystemCurl(FileSystemCurl&&)                 = delete;
    FileSystemCurl& operator=(const FileSystemCurl&) = delete;
    FileSystemCurl& operator=(FileSystemCurl&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override;
    ULONG STDMETHODCALLTYPE AddRef() noexcept override;
    ULONG STDMETHODCALLTYPE Release() noexcept override;

    // IInformations
    HRESULT STDMETHODCALLTYPE GetMetaData(const PluginMetaData** metaData) noexcept override;
    HRESULT STDMETHODCALLTYPE GetConfigurationSchema(const char** schemaJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE SetConfiguration(const char* configurationJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE GetConfiguration(const char** configurationJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE SomethingToSave(BOOL* pSomethingToSave) noexcept override;

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

private:
    ~FileSystemCurl() = default;

    static constexpr wchar_t kPluginIdFtp[]  = L"builtin/file-system-ftp";
    static constexpr wchar_t kPluginIdSftp[] = L"builtin/file-system-sftp";
    static constexpr wchar_t kPluginIdScp[]  = L"builtin/file-system-scp";
    static constexpr wchar_t kPluginIdImap[] = L"builtin/file-system-imap";

    static constexpr wchar_t kPluginShortIdFtp[]  = L"ftp";
    static constexpr wchar_t kPluginShortIdSftp[] = L"sftp";
    static constexpr wchar_t kPluginShortIdScp[]  = L"scp";
    static constexpr wchar_t kPluginShortIdImap[] = L"imap";

    static constexpr wchar_t kPluginNameFtp[]  = L"FTP";
    static constexpr wchar_t kPluginNameSftp[] = L"SFTP";
    static constexpr wchar_t kPluginNameScp[]  = L"SCP";
    static constexpr wchar_t kPluginNameImap[] = L"IMAP";

    static constexpr wchar_t kPluginDescriptionFtp[]  = L"FTP virtual file system.";
    static constexpr wchar_t kPluginDescriptionSftp[] = L"SFTP virtual file system (SSH File Transfer Protocol).";
    static constexpr wchar_t kPluginDescriptionScp[]  = L"SCP virtual file system (secure copy over SSH).";
    static constexpr wchar_t kPluginDescriptionImap[] = L"IMAP virtual mail file system.";
    static constexpr wchar_t kPluginAuthor[]          = L"RedSalamander";
    static constexpr wchar_t kPluginVersion[]         = L"0.3";

    static constexpr char kCapabilitiesJsonFtp[] = R"json(
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
    "deleteMax": 1,
    "deleteRecycleBinMax": 1
  },
  "crossFileSystem": {
    "export": { "copy": ["*"], "move": ["*"] },
    "import": { "copy": ["*"], "move": ["*"] }
  }
}
)json";

    static constexpr char kCapabilitiesJsonSftp[] = R"json(
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
    "deleteMax": 1,
    "deleteRecycleBinMax": 1
  },
  "crossFileSystem": {
    "export": { "copy": ["*"], "move": ["*"] },
    "import": { "copy": ["*"], "move": ["*"] }
  }
}
)json";

    static constexpr char kCapabilitiesJsonScp[] = R"json(
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
    "deleteMax": 1,
    "deleteRecycleBinMax": 1
  },
  "crossFileSystem": {
    "export": { "copy": ["*"], "move": ["*"] },
    "import": { "copy": ["*"], "move": ["*"] }
  }
}
)json";

    static constexpr char kCapabilitiesJsonImap[] = R"json(
{
  "version": 1,
  "operations": {
    "copy": false,
    "move": false,
    "delete": true,
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
    "export": { "copy": ["*"], "move": ["*"] },
    "import": { "copy": [], "move": [] }
  }
}
)json";

    static constexpr char kSchemaJsonFtp[] = R"json(
{
  "version": 1,
  "title": "FTP",
   "fields": [
     {
       "key": "defaultHost",
       "label": "Default host (for ftp:/)",
       "type": "text",
       "default": "",
       "description": "Host name used when navigating to ftp:/ (example: example.com)."
     },
     {
       "key": "defaultPort",
       "label": "Default port (0 = default)",
      "type": "value",
      "default": 0,
      "min": 0,
      "max": 65535
    },
     {
       "key": "defaultUser",
       "label": "Default user",
       "type": "text",
       "default": "",
       "description": "User name used when not provided in the URI."
     },
     {
       "key": "defaultPassword",
       "label": "Default password",
       "type": "text",
       "default": "",
       "description": "Password used when not provided in the URI (stored in settings as plain text)."
     },
     {
       "key": "defaultBasePath",
       "label": "Default base path",
       "type": "text",
       "default": "/",
       "description": "Remote base folder for ftp:/ (example: /pub)."
     },
     {
       "key": "connectTimeoutMs",
       "label": "Connect timeout (ms, 0 = libcurl default)",
      "type": "value",
      "default": 10000,
      "min": 0,
      "max": 600000
    },
    {
      "key": "operationTimeoutMs",
      "label": "Operation timeout (ms, 0 = no timeout)",
      "type": "value",
      "default": 0,
      "min": 0,
      "max": 3600000
    },
     {
       "key": "ftpUseEpsv",
       "label": "FTP: Use EPSV",
       "type": "bool",
       "default": true,
       "description": "Enables EPSV for FTP (recommended; disable only for legacy servers)."
     }
   ]
 }
 )json";

    static constexpr char kSchemaJsonSftp[] = R"json(
{
  "version": 1,
  "title": "SFTP",
  "fields": [
    {
      "key": "defaultHost",
      "label": "Default host (for sftp:/)",
      "type": "text",
      "default": "",
      "description": "Host name used when navigating to sftp:/ (example: example.com)."
    },
    {
      "key": "defaultPort",
      "label": "Default port (0 = default)",
      "type": "value",
      "default": 0,
      "min": 0,
      "max": 65535
    },
    {
      "key": "defaultUser",
      "label": "Default user",
      "type": "text",
      "default": "",
      "description": "User name used when not provided in the URI."
    },
    {
      "key": "defaultPassword",
      "label": "Default password",
      "type": "text",
      "default": "",
      "description": "Password used when not provided in the URI (stored in settings as plain text)."
    },
    {
      "key": "defaultBasePath",
      "label": "Default base path",
      "type": "text",
      "default": "/",
      "description": "Remote base folder for sftp:/ (example: /home/user)."
    },
    {
      "key": "connectTimeoutMs",
      "label": "Connect timeout (ms, 0 = libcurl default)",
      "type": "value",
      "default": 10000,
      "min": 0,
      "max": 600000
    },
    {
      "key": "operationTimeoutMs",
      "label": "Operation timeout (ms, 0 = no timeout)",
      "type": "value",
      "default": 0,
      "min": 0,
      "max": 3600000
    },
    {
      "key": "sshPrivateKey",
      "label": "SSH private key file",
      "type": "text",
      "default": "",
      "description": "Optional path to private key file for SFTP authentication."
    },
    {
      "key": "sshPublicKey",
      "label": "SSH public key file",
      "type": "text",
      "default": "",
      "description": "Optional path to public key file for SFTP authentication."
    },
    {
      "key": "sshKeyPassphrase",
      "label": "SSH key passphrase",
      "type": "text",
      "default": "",
      "description": "Optional passphrase for the SSH private key (stored in settings as plain text)."
    },
    {
      "key": "sshKnownHosts",
      "label": "SSH known_hosts file",
      "type": "text",
      "default": "",
      "description": "Optional known_hosts path for host key verification (empty disables strict host key checking)."
    }
  ]
}
)json";

    static constexpr char kSchemaJsonScp[] = R"json(
{
  "version": 1,
  "title": "SCP",
  "fields": [
    {
      "key": "defaultHost",
      "label": "Default host (for scp:/)",
      "type": "text",
      "default": "",
      "description": "Host name used when navigating to scp:/ (example: example.com)."
    },
    {
      "key": "defaultPort",
      "label": "Default port (0 = default)",
      "type": "value",
      "default": 0,
      "min": 0,
      "max": 65535
    },
    {
      "key": "defaultUser",
      "label": "Default user",
      "type": "text",
      "default": "",
      "description": "User name used when not provided in the URI."
    },
    {
      "key": "defaultPassword",
      "label": "Default password",
      "type": "text",
      "default": "",
      "description": "Password used for SSH authentication when not provided elsewhere (stored in settings as plain text)."
    },
    {
      "key": "defaultBasePath",
      "label": "Default base path",
      "type": "text",
      "default": "/",
      "description": "Remote base folder for scp:/ (example: /home/user)."
    },
    {
      "key": "connectTimeoutMs",
      "label": "Connect timeout (ms, 0 = libcurl default)",
      "type": "value",
      "default": 10000,
      "min": 0,
      "max": 600000
    },
    {
      "key": "operationTimeoutMs",
      "label": "Operation timeout (ms, 0 = no timeout)",
      "type": "value",
      "default": 0,
      "min": 0,
      "max": 3600000
    },
    {
      "key": "sshPrivateKey",
      "label": "SSH private key file",
      "type": "text",
      "default": "",
      "description": "Optional path to private key file for SCP authentication."
    },
    {
      "key": "sshPublicKey",
      "label": "SSH public key file",
      "type": "text",
      "default": "",
      "description": "Optional path to public key file for SCP authentication."
    },
    {
      "key": "sshKeyPassphrase",
      "label": "SSH key passphrase",
      "type": "text",
      "default": "",
      "description": "Optional passphrase for the SSH private key (stored in settings as plain text)."
    },
    {
      "key": "sshKnownHosts",
      "label": "SSH known_hosts file",
      "type": "text",
      "default": "",
      "description": "Optional known_hosts path for host key verification (empty disables strict host key checking)."
    }
  ]
}
)json";

    static constexpr char kSchemaJsonImap[] = R"json(
{
  "version": 1,
  "title": "IMAP",
  "fields": [
    {
      "key": "defaultHost",
      "label": "Default host (for imap:/)",
      "type": "text",
      "default": "",
      "description": "Host name used when navigating to imap:/ (example: imap.example.com)."
    },
    {
      "key": "defaultPort",
      "label": "Default port (0 = default)",
      "type": "value",
      "default": 0,
      "min": 0,
      "max": 65535
    },
    {
      "key": "ignoreSslTrust",
      "label": "Ignore trust for SSL",
      "type": "bool",
      "default": false,
      "description": "When enabled, TLS certificate validation is skipped (allows self-signed certificates; not recommended)."
    },
    {
      "key": "defaultUser",
      "label": "Default user",
      "type": "text",
      "default": "",
      "description": "User name used when not provided in the URI."
    },
    {
      "key": "defaultPassword",
      "label": "Default password",
      "type": "text",
      "default": "",
      "description": "Password used when not provided in the URI (stored in settings as plain text)."
    },
    {
      "key": "defaultBasePath",
      "label": "Default base path",
      "type": "text",
      "default": "/",
      "description": "Mailbox prefix for imap:/ (example: / for all mailboxes, or /INBOX to start in INBOX)."
    },
    {
      "key": "connectTimeoutMs",
      "label": "Connect timeout (ms, 0 = libcurl default)",
      "type": "value",
      "default": 10000,
      "min": 0,
      "max": 600000
    },
    {
      "key": "operationTimeoutMs",
      "label": "Operation timeout (ms, 0 = no timeout)",
      "type": "value",
      "default": 0,
      "min": 0,
      "max": 3600000
    }
  ]
}
)json";

public:
    struct Settings
    {
        std::wstring defaultHost;
        unsigned int defaultPort = 0;
        std::wstring defaultUser;
        std::wstring defaultPassword;
        std::wstring defaultBasePath = L"/";

        unsigned long connectTimeoutMs   = 10000;
        unsigned long operationTimeoutMs = 0;

        bool ignoreSslTrust = false;
        bool ftpUseEpsv     = true;

        std::wstring sshPrivateKey;
        std::wstring sshPublicKey;
        std::wstring sshKeyPassphrase;
        std::wstring sshKnownHosts;
    };

private:
    struct MenuEntry
    {
        NavigationMenuItemFlags flags = NAV_MENU_ITEM_FLAG_NONE;
        std::wstring label;
        std::wstring path;
        std::wstring iconPath;
        unsigned int commandId = 0;
    };

    std::atomic_ulong _refCount{1};

    std::mutex _stateMutex;
    FileSystemCurlProtocol _protocol = FileSystemCurlProtocol::Sftp;
    PluginMetaData _metaData{};
    std::string _configurationJson;
    Settings _settings;
    wil::com_ptr<IHostConnections> _hostConnections;

    // NavigationMenu state.
    std::vector<MenuEntry> _menuEntries;
    std::vector<NavigationMenuItem> _menuEntryView;
    INavigationMenuCallback* _navigationMenuCallback = nullptr;
    void* _navigationMenuCallbackCookie              = nullptr;

    // DriveInfo state.
    std::wstring _driveDisplayName;
    std::wstring _driveFileSystem;
    DriveInfo _driveInfo{};
    std::vector<MenuEntry> _driveMenuEntries;
    std::vector<NavigationMenuItem> _driveMenuEntryView;

    std::mutex _propertiesMutex;
    std::string _lastPropertiesJson;
};
