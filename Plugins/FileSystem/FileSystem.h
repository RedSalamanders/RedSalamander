#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <algorithm>
#include <atomic>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <cwchar>
#include <format>
#include <memory>
#include <mutex>
#include <new>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4514 28182) // WIL headers: deleted copy/move and unreferenced inline Helpers
#include <wil/com.h>
#include <wil/resource.h>
#pragma warning(pop)

#include "PlugInterfaces/DriveInfo.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/NavigationMenu.h"

enum class FileSystemReparsePointPolicy : uint8_t
{
    CopyReparse,
    FollowTargets,
    Skip,
};

class FilesInformation final : public IFilesInformation
{
public:
    FilesInformation();
    ~FilesInformation();

    // Explicitly delete copy/move operations (COM objects are not copyable/movable)
    FilesInformation(const FilesInformation&)            = delete;
    FilesInformation(FilesInformation&&)                 = delete;
    FilesInformation& operator=(const FilesInformation&) = delete;
    FilesInformation& operator=(FilesInformation&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override;
    ULONG STDMETHODCALLTYPE AddRef() noexcept override;
    ULONG STDMETHODCALLTYPE Release() noexcept override;

    HRESULT STDMETHODCALLTYPE GetBuffer(FileInfo** ppFileInfo) noexcept override;
    HRESULT STDMETHODCALLTYPE GetBufferSize(unsigned long* pSize) noexcept override;
    HRESULT STDMETHODCALLTYPE GetAllocatedSize(unsigned long* pSize) noexcept override;
    HRESULT STDMETHODCALLTYPE GetCount(unsigned long* pCount) noexcept override;
    HRESULT STDMETHODCALLTYPE Get(unsigned long index, FileInfo** ppEntry) noexcept override;

private:
    // Internal writer API (not part of the public COM interface)
    // Used by FileSystem to build the buffer during ReadDirectoryInfo().
    HRESULT BeginWrite(FileInfo** ppFileInfo) noexcept;
    void UpdateUsage(unsigned long bytesUsed, unsigned long count) noexcept;
    void MaybeTrimBuffer() noexcept;

    friend class FileSystem;

    void ResetDirectoryState(bool clearPath) noexcept;
    void ResizeBuffer(size_t newSize) noexcept;
    bool PathEquals(const std::wstring& other) const noexcept;
    size_t ComputeEntrySize(const FileInfo* entry) const noexcept;
    HRESULT LocateEntry(unsigned long index, FileInfo** ppEntry) const noexcept;

    std::atomic_ulong _refCount{1};
    std::vector<std::byte> _buffer;
    unsigned long _count     = 0;
    unsigned long _usedBytes = 0;

    std::wstring _requestedPath;
    bool _enumerationInitialized = false;
    bool _enumerationComplete    = false;
    bool _hasPendingEntry        = false;
    WIN32_FIND_DATAW _pendingEntry{};
    wil::unique_hfind _findHandle;

    // Handle-based enumeration state (GetFileInformationByHandleEx) for local disks.
    bool _useHandleEnumeration          = false;
    bool _enumerationRestartScan        = true;
    size_t _enumerationBufferOffset     = 0;
    size_t _enumerationBufferBytesValid = 0;
    wil::unique_handle _directoryHandle;
    std::vector<std::byte> _enumerationBuffer;
};

class FileSystem final : public IFileSystem,
                         public IFileSystemIO,
                         public IFileSystemDirectoryOperations,
                         public IFileSystemDirectoryWatch,
                         public IInformations,
                         public INavigationMenu,
                         public IDriveInfo
{
public:
    FileSystem();

    // Explicitly delete copy/move operations (COM objects are not copyable/movable)
    FileSystem(const FileSystem&)            = delete;
    FileSystem(FileSystem&&)                 = delete;
    FileSystem& operator=(const FileSystem&) = delete;
    FileSystem& operator=(FileSystem&&)      = delete;

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
    HRESULT STDMETHODCALLTYPE CreateDirectory(const wchar_t* path) noexcept override;
    HRESULT STDMETHODCALLTYPE GetDirectorySize(const wchar_t* path,
                                               FileSystemFlags flags,
                                               IFileSystemDirectorySizeCallback* callback,
                                               void* cookie,
                                               FileSystemDirectorySizeResult* result) noexcept override;
    HRESULT STDMETHODCALLTYPE WatchDirectory(const wchar_t* path, IFileSystemDirectoryWatchCallback* callback, void* cookie) noexcept override;
    HRESULT STDMETHODCALLTYPE UnwatchDirectory(const wchar_t* path) noexcept override;
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

    HRESULT STDMETHODCALLTYPE GetAttributes(const wchar_t* path, unsigned long* fileAttributes) noexcept override;

    HRESULT STDMETHODCALLTYPE CreateFileReader(const wchar_t* path, IFileReader** reader) noexcept override;

    HRESULT STDMETHODCALLTYPE CreateFileWriter(const wchar_t* path, FileSystemFlags flags, IFileWriter** writer) noexcept override;

    HRESULT STDMETHODCALLTYPE GetFileBasicInformation(const wchar_t* path, FileSystemBasicInformation* info) noexcept override;

    HRESULT STDMETHODCALLTYPE SetFileBasicInformation(const wchar_t* path, const FileSystemBasicInformation* info) noexcept override;

    HRESULT STDMETHODCALLTYPE GetItemProperties(const wchar_t* path, const char** jsonUtf8) noexcept override;

    HRESULT STDMETHODCALLTYPE GetCapabilities(const char** jsonUtf8) noexcept override;

private:
    ~FileSystem();

    static constexpr wchar_t kPluginId[]          = L"builtin/file-system";
    static constexpr wchar_t kPluginShortId[]     = L"file";
    static constexpr wchar_t kPluginName[]        = L"File System";
    static constexpr wchar_t kPluginDescription[] = L"Local Windows file system implementation.";
    static constexpr wchar_t kPluginAuthor[]      = L"RedSalamander";
    static constexpr wchar_t kPluginVersion[]     = L"1.0";

    static constexpr char kSchemaJson[] = R"json(
{
  "version": 1,
  "title": "File System",
  "fields": [
    {
      "key": "copyMoveMaxConcurrency",
      "type": "value",
      "label": "Copy/Move max concurrency",
      "description": "Maximum number of worker threads used by Copy/Move batch operations (top-level items). 1 disables internal parallelism.",
      "default": 4,
      "min": 1,
      "max": 8
    },
    {
      "key": "deleteMaxConcurrency",
      "type": "value",
      "label": "Delete max concurrency",
      "description": "Maximum number of worker threads used by Delete batch operations (top-level items).",
      "default": 8,
      "min": 1,
      "max": 64
    },
    {
      "key": "deleteRecycleBinMaxConcurrency",
      "type": "value",
      "label": "Delete (Recycle Bin) max concurrency",
      "description": "Maximum concurrency for Recycle Bin deletes. Keep this small to reduce shell overhead.",
      "default": 2,
      "min": 1,
      "max": 16
    },
    {
      "key": "enumerationSoftMaxBufferMiB",
      "type": "value",
      "label": "Enumeration soft cap (MiB)",
      "description": "Soft cap for the directory enumeration buffer (in MiB). When exceeded, the plugin may grow up to the hard cap as a fallback.",
      "default": 512,
      "min": 1,
      "max": 4095
    },
    {
      "key": "enumerationHardMaxBufferMiB",
      "type": "value",
      "label": "Enumeration hard cap (MiB)",
      "description": "Hard cap for the directory enumeration buffer (in MiB). Must be >= soft cap; exceeding returns ERROR_INSUFFICIENT_BUFFER.",
      "default": 2048,
      "min": 1,
      "max": 4095
    },
    {
      "key": "reparsePointPolicy",
      "type": "option",
      "label": "Reparse points (symlinks/junctions)",
      "description": "When recursively copying/moving directories, controls whether reparse points are recreated as links, followed, or skipped. Default is safest.",
      "default": "copyReparse",
      "options": [
        { "value": "copyReparse", "label": "Create link/reparse at destination" },
        { "value": "followTargets", "label": "Follow targets (can loop / escape tree)" },
        { "value": "skip", "label": "Skip reparse points" }
      ]
    }
  ]
}
)json";

    static constexpr unsigned int kDefaultCopyMoveMaxConcurrency             = 4u;
    static constexpr unsigned int kDefaultDeleteMaxConcurrency               = 8u;
    static constexpr unsigned int kDefaultDeleteRecycleBinMaxConcurrency     = 2u;
    static constexpr unsigned long kDefaultEnumerationSoftMaxBufferMiB       = 512ul;
    static constexpr unsigned long kDefaultEnumerationHardMaxBufferMiB       = 2048ul;
    static constexpr FileSystemReparsePointPolicy kDefaultReparsePointPolicy = FileSystemReparsePointPolicy::CopyReparse;

    static constexpr unsigned int kMaxCopyMoveMaxConcurrency         = 8u;
    static constexpr unsigned int kMaxDeleteMaxConcurrency           = 64u;
    static constexpr unsigned int kMaxDeleteRecycleBinMaxConcurrency = 16u;

    PluginMetaData _metaData{};

    // Guards configuration/capabilities and other shared state that is mutated during normal operation
    // (Preferences updates, menu/drive info string storage, etc.). File operations capture a snapshot
    // of the relevant settings at the start of each call to avoid holding this lock during IO.
    std::mutex _stateMutex;

    std::string _configurationJson;
    std::string _capabilitiesJson;
    mutable std::mutex _propertiesMutex;
    std::string _lastPropertiesJson;

    unsigned int _copyMoveMaxConcurrency             = kDefaultCopyMoveMaxConcurrency;
    unsigned int _deleteMaxConcurrency               = kDefaultDeleteMaxConcurrency;
    unsigned int _deleteRecycleBinMaxConcurrency     = kDefaultDeleteRecycleBinMaxConcurrency;
    unsigned long _enumerationSoftMaxBufferMiB       = kDefaultEnumerationSoftMaxBufferMiB;
    unsigned long _enumerationHardMaxBufferMiB       = kDefaultEnumerationHardMaxBufferMiB;
    FileSystemReparsePointPolicy _reparsePointPolicy = kDefaultReparsePointPolicy;
#ifdef _DEBUG
    unsigned int _directorySizeDelayMs = 0u;
#endif

    struct MenuEntry
    {
        std::wstring label;
        std::wstring path;
        std::wstring iconPath;
        NavigationMenuItemFlags flags = NAV_MENU_ITEM_FLAG_NONE;
        unsigned int commandId        = 0;
    };

    std::vector<MenuEntry> _menuEntries;
    std::vector<NavigationMenuItem> _menuEntryView;
    INavigationMenuCallback* _navigationMenuCallback = nullptr;
    void* _navigationMenuCallbackCookie              = nullptr;
    std::vector<MenuEntry> _driveMenuEntries;
    std::vector<NavigationMenuItem> _driveMenuEntryView;

    std::wstring _driveDisplayName;
    std::wstring _driveVolumeLabel;
    std::wstring _driveFileSystem;
    DriveInfo _driveInfo{};

    void EnsureRequestedPath(FilesInformation& info, const std::wstring& path) noexcept;

    HRESULT PopulateFilesInformation(FilesInformation& info, const std::wstring& path, unsigned long& bytesWritten, unsigned long& entryCount) noexcept;
    HRESULT PopulateServerShares(FilesInformation& info, std::wstring_view serverName, unsigned long& bytesWritten, unsigned long& entryCount) noexcept;
    HRESULT EnsureEnumeration(FilesInformation& info, const std::wstring& path) noexcept;
    HRESULT StartEnumeration(FilesInformation& info, const std::wstring& path) noexcept;
    HRESULT PopulateBuffer(FilesInformation& info, unsigned long& bytesWritten, unsigned long& entryCount, size_t& lastEntrySize) noexcept;

    HRESULT StartEnumerationWin32(FilesInformation& info, const std::wstring& path) noexcept;
    HRESULT StartEnumerationHandle(FilesInformation& info, const std::wstring& path) noexcept;
    HRESULT FetchNextEntryWin32(FilesInformation& info, WIN32_FIND_DATAW& data) noexcept;
    HRESULT PopulateBufferWin32(FilesInformation& info, unsigned long& bytesWritten, unsigned long& entryCount, size_t& lastEntrySize) noexcept;
    HRESULT PopulateBufferHandle(FilesInformation& info, unsigned long& bytesWritten, unsigned long& entryCount, size_t& lastEntrySize) noexcept;

    class DirectoryWatch;

    std::mutex _watchMutex;
    std::unordered_map<std::wstring, std::unique_ptr<DirectoryWatch>> _directoryWatches;

    std::atomic_ulong _refCount{1};

    void UpdateCapabilitiesJson() noexcept;
};
