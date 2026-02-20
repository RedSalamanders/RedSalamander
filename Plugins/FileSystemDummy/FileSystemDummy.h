#pragma once

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <atomic>
#include <condition_variable>
#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <memory>
#include <mutex>
#include <random>
#include <string>
#include <string_view>
#include <vector>

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027)
#include <wil/com.h>
#pragma warning(pop)

#include "PlugInterfaces/DriveInfo.h"
#include "PlugInterfaces/FileSystem.h"
#include "PlugInterfaces/Informations.h"
#include "PlugInterfaces/NavigationMenu.h"

class DummyFilesInformation final : public IFilesInformation
{
public:
    DummyFilesInformation(std::vector<std::byte> buffer, unsigned long count, unsigned long usedBytes) noexcept;
    ~DummyFilesInformation() = default;

    // Explicitly delete copy/move operations (COM objects are not copyable/movable).
    DummyFilesInformation(const DummyFilesInformation&)            = delete;
    DummyFilesInformation(DummyFilesInformation&&)                 = delete;
    DummyFilesInformation& operator=(const DummyFilesInformation&) = delete;
    DummyFilesInformation& operator=(DummyFilesInformation&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override;
    ULONG STDMETHODCALLTYPE AddRef() noexcept override;
    ULONG STDMETHODCALLTYPE Release() noexcept override;

    HRESULT STDMETHODCALLTYPE GetBuffer(FileInfo** ppFileInfo) noexcept override;
    HRESULT STDMETHODCALLTYPE GetBufferSize(unsigned long* pSize) noexcept override;
    HRESULT STDMETHODCALLTYPE GetAllocatedSize(unsigned long* pSize) noexcept override;
    HRESULT STDMETHODCALLTYPE GetCount(unsigned long* pCount) noexcept override;
    HRESULT STDMETHODCALLTYPE Get(unsigned long index, FileInfo** ppEntry) noexcept override;

private:
    HRESULT LocateEntry(unsigned long index, FileInfo** ppEntry) const noexcept;
    size_t ComputeEntrySize(const FileInfo* entry) const noexcept;

    std::atomic_ulong _refCount{1};
    std::vector<std::byte> _buffer;
    unsigned long _count     = 0;
    unsigned long _usedBytes = 0;
};

class FileSystemDummy final : public IFileSystem,
                              public IFileSystemIO,
                              public IFileSystemDirectoryOperations,
                              public IFileSystemDirectoryWatch,
                              public IInformations,
                              public INavigationMenu,
                              public IDriveInfo
{
public:
    FileSystemDummy();

    // Explicitly delete copy/move operations (COM objects are not copyable/movable).
    FileSystemDummy(const FileSystemDummy&)            = delete;
    FileSystemDummy(FileSystemDummy&&)                 = delete;
    FileSystemDummy& operator=(const FileSystemDummy&) = delete;
    FileSystemDummy& operator=(FileSystemDummy&&)      = delete;

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
    HRESULT STDMETHODCALLTYPE GetAttributes(const wchar_t* path, unsigned long* fileAttributes) noexcept override;
    HRESULT STDMETHODCALLTYPE CreateFileReader(const wchar_t* path, IFileReader** reader) noexcept override;
    HRESULT STDMETHODCALLTYPE CreateFileWriter(const wchar_t* path, FileSystemFlags flags, IFileWriter** writer) noexcept override;
    HRESULT STDMETHODCALLTYPE GetFileBasicInformation(const wchar_t* path, FileSystemBasicInformation* info) noexcept override;
    HRESULT STDMETHODCALLTYPE SetFileBasicInformation(const wchar_t* path, const FileSystemBasicInformation* info) noexcept override;
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

    HRESULT STDMETHODCALLTYPE GetItemProperties(const wchar_t* path, const char** jsonUtf8) noexcept override;

    HRESULT STDMETHODCALLTYPE GetCapabilities(const char** jsonUtf8) noexcept override;

    // Internal helper used by IFileWriter implementations.
    HRESULT
    CommitFileWriter(const std::filesystem::path& normalizedPath, FileSystemFlags flags, const std::shared_ptr<std::vector<std::byte>>& buffer) noexcept;

private:
    struct DummyNode
    {
        std::wstring name;
        bool isDirectory                = false;
        DWORD attributes                = 0;
        uint64_t sizeBytes              = 0;
        __int64 creationTime            = 0;
        __int64 lastAccessTime          = 0;
        __int64 lastWriteTime           = 0;
        __int64 changeTime              = 0;
        std::uint64_t generationSeed    = 0;
        unsigned long plannedChildCount = 0;
        bool childrenGenerated          = false;
        std::shared_ptr<std::vector<std::byte>> materializedContent;
        DummyNode* parent = nullptr;
        std::vector<std::unique_ptr<DummyNode>> children;
    };

#pragma warning(push)
// C4625 (copy ctor deleted), C4626 (copy assign deleted)
#pragma warning(disable : 4625 4626)
    struct DummyRoot
    {
        std::wstring rootPath;
        std::unique_ptr<DummyNode> node;
    };
#pragma warning(pop)

    ~FileSystemDummy();

    static constexpr wchar_t kPluginId[]          = L"builtin/file-system-dummy";
    static constexpr wchar_t kPluginShortId[]     = L"fk";
    static constexpr wchar_t kPluginName[]        = L"File System Dummy";
    static constexpr wchar_t kPluginDescription[] = L"Deterministic in-memory dummy file system for testing.";
    static constexpr wchar_t kPluginAuthor[]      = L"RedSalamander";
    static constexpr wchar_t kPluginVersion[]     = L"1.0";

    static constexpr char kSchemaJson[] = R"json({
	    "version":1,
	    "title":"File System Dummy",
	    "fields":[
	        {
	            "key":"maxChildrenPerDirectory",
	            "type":"value",
	            "label":"Max children per directory",
	            "description":"Upper bound for how many children are generated in each directory.",
	            "default":42,
	            "min":0,
	            "max":20000
	        },
	        {
	            "key":"maxDepth",
	            "type":"value",
	            "label":"Max depth",
	            "description":"Maximum generated directory depth (0 = unlimited).",
	            "default":10,
	            "min":0,
	            "max":1024
	        },
	        {
	            "key":"seed",
	            "type":"value",
	            "label":"Random seed (0 = random)",
	            "description":"Seed used by the deterministic generator; 0 picks a random seed.",
            "default":42,
            "min":0,
            "max":4294967295
        },
        {
            "key":"latencyMs",
            "type":"value",
            "label":"Latency (ms)",
            "description":"Artificial latency per file access and per directory entry enumerated (0 = none).",
            "default":0,
            "min":0,
            "max":1000
        },
        {
            "key":"virtualSpeedLimit",
            "type":"text",
            "label":"Virtual speed limit",
            "description":"Maximum copy/move throughput for the dummy file system (0 = unlimited). Examples: 3KB, 4MB.",
            "default":"0"
        }
    ]
})json";

    static constexpr char kCapabilitiesJson[] = R"json(
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
    "copyMoveMax": 4,
    "deleteMax": 8,
    "deleteRecycleBinMax": 2
  },
  "crossFileSystem": {
    "export": { "copy": ["*"], "move": ["*"] },
    "import": { "copy": ["*"], "move": ["*"] }
  }
}
)json";

    PluginMetaData _metaData{};
    inline static std::string _configurationJson;
    inline static unsigned long _maxChildrenPerDirectory = 42;
    inline static unsigned long _maxDepth                = 10;
    inline static unsigned int _seed                     = 42;
    inline static unsigned long _latencyMilliseconds     = 0;
    inline static std::wstring _virtualSpeedLimitText    = L"0";
    inline static std::atomic<uint64_t> _virtualSpeedLimitBytesPerSecond{0};

    mutable std::mutex _propertiesMutex;
    std::string _lastPropertiesJson;

    struct MenuEntry
    {
        std::wstring label;
        std::wstring path;
        std::wstring iconPath;
        NavigationMenuItemFlags flags = NAV_MENU_ITEM_FLAG_NONE;
        unsigned int commandId        = 0;
    };

    std::mutex _stateMutex;

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

    // Clears _roots iteratively to avoid stack overflow on deeply-nested trees.
    // Must be called while _mutex is held.
    static void ClearRootsIteratively() noexcept;

    std::filesystem::path NormalizePath(std::wstring_view path) const;
    DummyRoot* FindRoot(std::wstring_view rootPath) noexcept;
    DummyRoot* GetOrCreateRoot(std::wstring_view rootPath);
    HRESULT ResolvePath(const std::filesystem::path& path, DummyNode** outNode, bool createMissing, bool requireDirectory) noexcept;
    DummyNode* FindChild(DummyNode* parent, std::wstring_view name) const noexcept;
    std::unique_ptr<DummyNode> ExtractChild(DummyNode* parent, DummyNode* child) noexcept;
    DummyNode* AddChild(DummyNode* parent, std::unique_ptr<DummyNode> child);
    std::unique_ptr<DummyNode> CreateNode(std::wstring_view name, bool isDirectory, std::uint64_t generationSeed);
    void EnsureChildrenGenerated(DummyNode& node);
    void GenerateChildren(DummyNode& node);
    bool IsNameValid(std::wstring_view name) const noexcept;
    std::wstring MakeUniqueName(DummyNode* parent, std::wstring_view baseName) const;
    std::wstring MakeRandomName(std::mt19937& rng, bool isDirectory);
    std::wstring MakeRandomBaseName(std::mt19937& rng);
    void TouchNode(DummyNode& node) noexcept;
    void TouchParent(DummyNode* parent) noexcept;
    uint64_t ComputeNodeBytes(const DummyNode& node) const noexcept;
    bool IsAncestor(const DummyNode& node, const DummyNode& possibleDescendant) const noexcept;
    std::unique_ptr<DummyNode> CloneNode(const DummyNode& source);
    HRESULT CreateDirectoryClone(
        const DummyNode& sourceDirectory, DummyNode& destinationParent, std::wstring_view destinationName, FileSystemFlags flags, DummyNode** outDirectory);
    HRESULT CopyNode(DummyNode& source, DummyNode& destinationParent, std::wstring_view destinationName, FileSystemFlags flags, uint64_t* outBytes);
    HRESULT MoveNode(DummyNode& source, DummyNode& destinationParent, std::wstring_view destinationName, FileSystemFlags flags, uint64_t* outBytes);
    HRESULT DeleteNode(DummyNode& target, FileSystemFlags flags);
    void SimulateLatency(unsigned long itemCount) const noexcept;

    struct DirectoryWatchRegistration
    {
        DirectoryWatchRegistration() = default;

        DirectoryWatchRegistration(const DirectoryWatchRegistration&)            = delete;
        DirectoryWatchRegistration& operator=(const DirectoryWatchRegistration&) = delete;
        DirectoryWatchRegistration(DirectoryWatchRegistration&&)                 = delete;
        DirectoryWatchRegistration& operator=(DirectoryWatchRegistration&&)      = delete;

        const FileSystemDummy* owner = nullptr;
        std::wstring watchedPath;
        IFileSystemDirectoryWatchCallback* callback = nullptr;
        void* cookie                                = nullptr;
        std::atomic_uint32_t inFlight{0};
        std::atomic_bool active{true};
    };

    void NotifyDirectoryWatchers(std::wstring_view watchedPath, std::wstring_view relativePath, FileSystemDirectoryChangeAction action) noexcept;
    void NotifyDirectoryWatchers(std::wstring_view watchedPath, std::wstring_view oldRelativePath, std::wstring_view newRelativePath) noexcept;

    std::atomic_ulong _refCount{1};
    inline static std::mutex _mutex;
    inline static std::uint64_t _effectiveSeed      = 0;
    inline static std::uint64_t _generationBaseTime = 0;
    inline static std::vector<std::unique_ptr<DummyRoot>> _roots;

    inline static std::mutex _watchMutex;
    inline static std::condition_variable _watchCv;
    inline static std::vector<std::shared_ptr<DirectoryWatchRegistration>> _directoryWatches;
};
