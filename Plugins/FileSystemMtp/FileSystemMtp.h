#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <array>
#include <atomic>
#include <condition_variable>
#include <cstddef>
#include <cstdint>
#include <functional>
#include <memory>
#include <mutex>
#include <string>
#include <string_view>
#include <thread>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182) // WIL headers: deleted copy/move and unused inline helpers
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "FileSystemMtp.Internal.h"
#include "PlugInterfaces/DriveInfo.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Host.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/NavigationMenu.h"
#include "PackedFileInfoBuffer.h"

class FilesInformationMtp final : public IFilesInformation
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

    FilesInformationMtp()  = default;
    ~FilesInformationMtp() = default;

    FilesInformationMtp(const FilesInformationMtp&)            = delete;
    FilesInformationMtp(FilesInformationMtp&&)                 = delete;
    FilesInformationMtp& operator=(const FilesInformationMtp&) = delete;
    FilesInformationMtp& operator=(FilesInformationMtp&&)      = delete;

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

class MtpBackendReader;

class FileSystemMtp final : public IFileSystem,
                            public IFileSystemIO,
                            public IFileSystemDirectoryOperations,
                            public IFileSystemInitialize,
                            public IInformations,
                            public INavigationMenu,
                            public IDriveInfo
{
public:
    explicit FileSystemMtp(IHost* host) noexcept;
    FileSystemMtp(IHost* host, std::unique_ptr<FileSystemMtpInternal::IMtpBackend> backend) noexcept;

    FileSystemMtp(const FileSystemMtp&)            = delete;
    FileSystemMtp(FileSystemMtp&&)                 = delete;
    FileSystemMtp& operator=(const FileSystemMtp&) = delete;
    FileSystemMtp& operator=(FileSystemMtp&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override;
    ULONG STDMETHODCALLTYPE AddRef() noexcept override;
    ULONG STDMETHODCALLTYPE Release() noexcept override;

    HRESULT STDMETHODCALLTYPE GetMetaData(const PluginMetaData** metaData) noexcept override;
    HRESULT STDMETHODCALLTYPE GetConfigurationSchema(const char** schemaJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE SetConfiguration(const char* configurationJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE GetConfiguration(const char** configurationJsonUtf8) noexcept override;
    HRESULT STDMETHODCALLTYPE SomethingToSave(BOOL* pSomethingToSave) noexcept override;
    [[nodiscard]] static const char* StaticConfigurationSchema() noexcept;

    HRESULT STDMETHODCALLTYPE Initialize(const wchar_t* rootPath, const char* optionsJsonUtf8) noexcept override;

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
    HRESULT STDMETHODCALLTYPE GetTransferHints(const wchar_t* path,
                                               FileSystemOperation operationType,
                                               FileSystemTransferEndpoint endpoint,
                                               FileSystemTransferHints* hints) noexcept override;
    HRESULT STDMETHODCALLTYPE GetStorageCharacteristics(const wchar_t* path, FileSystemStorageCharacteristics* characteristics) noexcept override;

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

    HRESULT CommitFileWriter(std::wstring_view normalizedPath, FileSystemFlags flags, std::span<const std::byte> bytes) noexcept;

private:
    friend class MtpBackendReader;

    ~FileSystemMtp();

    struct Settings
    {
        bool readOnly                     = true;
        std::string byteVerifyOnOverwrite = "transmitHash";
        uint32_t commandTimeoutMs         = 120'000;
        uint32_t enumerationPageSize      = 32;
        bool failOverwriteJournalWrites   = false;
        std::wstring connectionHost;
        std::wstring connectionInitialPath = L"/";
        std::wstring connectionDevicePuid;
        std::wstring connectionFriendlyName;
    };

    struct JsonReturnBuffers
    {
        std::array<std::string, 2> slots{};
        uint8_t active = 0;
    };

    struct MenuEntry
    {
        std::wstring label;
        std::wstring path;
        std::wstring iconPath;
        NavigationMenuItemFlags flags = NAV_MENU_ITEM_FLAG_NONE;
        unsigned int commandId        = 0;
    };

    enum class CreatedObjectPuidPolicy : uint8_t
    {
        Unknown,
        Supported,
        Unsupported,
    };

    [[nodiscard]] const char* StoreJson(JsonReturnBuffers& buffers, std::string jsonUtf8) noexcept;
    [[nodiscard]] std::string BuildConfigurationJson() const;
    [[nodiscard]] std::string BuildCapabilitiesJson() const;
    [[nodiscard]] HRESULT NormalizeInputPath(const wchar_t* path, std::wstring& normalized) const noexcept;
    [[nodiscard]] bool MutationsAllowed() const noexcept;
    [[nodiscard]] HRESULT CheckMutationAllowed() const noexcept;
    [[nodiscard]] HRESULT CheckConnected() const noexcept;
    [[nodiscard]] bool CreatedObjectPuidUnsupported() const noexcept;
    void RecordCreatedObjectPuidProbe(bool supported) noexcept;
    void AbandonBackendSessionLocked(const std::shared_ptr<FileSystemMtpInternal::IMtpBackend>& abandonedBackend) noexcept;
    [[nodiscard]] HRESULT CreateBackendWorkerLocked() noexcept;
    [[nodiscard]] std::wstring OverwriteJournalIdentityForPath(std::wstring_view normalizedPath) const noexcept;
    HRESULT RunBackendCommand(std::function<HRESULT(FileSystemMtpInternal::IMtpBackend&)> command,
                              std::wstring recoveryDeviceIdentity = {},
                              FileSystemMtpInternal::MtpBackendCommandKind kind = FileSystemMtpInternal::MtpBackendCommandKind::ReadOnly,
                              uint64_t requiredBackendGeneration = 0u) noexcept;
    HRESULT CompleteSingleItem(FileSystemOperation operationType,
                               unsigned long itemIndex,
                               const wchar_t* sourcePath,
                               const wchar_t* destinationPath,
                               HRESULT status,
                               const FileSystemOptions* options,
                               IFileSystemCallback* callback,
                               void* cookie) noexcept;
    HRESULT ReportSingleItemStartAndCheckCancel(FileSystemOperation operationType,
                                                const wchar_t* sourcePath,
                                                const wchar_t* destinationPath,
                                                const FileSystemOptions* options,
                                                IFileSystemCallback* callback,
                                                void* cookie) noexcept;
    HRESULT CopyOrMoveItems(bool move,
                            const wchar_t* const* sourcePaths,
                            unsigned long count,
                            const wchar_t* destinationFolder,
                            FileSystemFlags flags,
                            const FileSystemOptions* options,
                            IFileSystemCallback* callback,
                            void* cookie) noexcept;
    HRESULT AccumulateDirectorySize(std::wstring_view path,
                                    bool recursive,
                                    IFileSystemDirectorySizeCallback* callback,
                                    void* cookie,
                                    FileSystemDirectorySizeResult& result,
                                    uint64_t& scannedEntries) noexcept;

    static constexpr wchar_t kPluginId[]      = L"builtin/file-system-mtp";
    static constexpr wchar_t kPluginShortId[] = L"mtp";
    static constexpr wchar_t kPluginAuthor[]  = L"RedSalamander";
    static constexpr wchar_t kPluginVersion[] = VERSINFO_PLUGIN_VERSION;

    PluginMetaData _metaData{};
    std::atomic_ulong _refCount{1};
    mutable std::mutex _stateMutex;
    mutable std::mutex _jsonMutex;
    Settings _settings{};
    bool _initialized      = false;
    std::wstring _rootPath = L"/";
    std::string _configurationSource;
    JsonReturnBuffers _configurationJson;
    JsonReturnBuffers _capabilitiesJson;
    JsonReturnBuffers _itemPropertiesJson;
    std::shared_ptr<FileSystemMtpInternal::IMtpBackend> _backend;
    std::shared_ptr<FileSystemMtpInternal::MtpBackendCommandQueue> _backendWorker;
    HRESULT _backendWorkerCreationHr = E_PENDING;
    uint64_t _backendGeneration      = 1u;
    wil::com_ptr<IHostConnections> _hostConnections;
    bool _disconnected                               = false;
    CreatedObjectPuidPolicy _createdObjectPuidPolicy = CreatedObjectPuidPolicy::Unknown;

    std::vector<MenuEntry> _menuEntries;
    std::vector<NavigationMenuItem> _menuEntryView;
    std::vector<MenuEntry> _driveMenuEntries;
    std::vector<NavigationMenuItem> _driveMenuEntryView;
    INavigationMenuCallback* _navigationMenuCallback = nullptr;
    void* _navigationMenuCookie                      = nullptr;
    std::condition_variable _navigationMenuCallbackCv;
    uint32_t _navigationMenuActiveCallbacks = 0;

    std::wstring _driveDisplayName;
    std::wstring _driveVolumeLabel;
    std::wstring _driveFileSystem;
    DriveInfo _driveInfo{};
};

[[nodiscard]] const char* GetFileSystemMtpStaticConfigurationSchema() noexcept;
void ShutdownFileSystemMtpModule() noexcept;
[[nodiscard]] bool CanUnloadFileSystemMtpModule() noexcept;
[[nodiscard]] bool RetainFileSystemMtpModuleUntilProcessExit() noexcept;
