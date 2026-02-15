#pragma once

#include <limits.h>
#include <unknwn.h>
#include <wchar.h>
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#pragma warning(push)
#pragma warning(disable : 4820) // padding in data structure
struct FileInfo
{
    unsigned long NextEntryOffset;
    unsigned long FileIndex;
    __int64 CreationTime;
    __int64 LastAccessTime;
    __int64 LastWriteTime;
    __int64 ChangeTime;
    __int64 EndOfFile;
    __int64 AllocationSize;
    unsigned long FileAttributes;
    // Length of the file name in bytes (not characters).
    // Callers MUST use this length and MUST NOT assume FileName is null-terminated.
    unsigned long FileNameSize;
    unsigned long EaSize;
    wchar_t FileName[1];
};

enum FileSystemOperation : uint32_t
{
    FILESYSTEM_COPY   = 1,
    FILESYSTEM_MOVE   = 2,
    FILESYSTEM_DELETE = 3,
    FILESYSTEM_RENAME = 4,
};

enum FileSystemFlags : uint32_t
{
    FILESYSTEM_FLAG_NONE                   = 0,
    FILESYSTEM_FLAG_ALLOW_OVERWRITE        = 0x1,
    FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY = 0x2,
    FILESYSTEM_FLAG_RECURSIVE              = 0x4,
    FILESYSTEM_FLAG_USE_RECYCLE_BIN        = 0x8,
    FILESYSTEM_FLAG_CONTINUE_ON_ERROR      = 0x10,
};

enum class FileSystemIssueAction : uint8_t
{
    None = 0,
    Overwrite,
    ReplaceReadOnly,
    PermanentDelete,
    Retry,
    Skip,
    Cancel,
};

struct FileSystemOptions
{
    // 0 = unlimited (use all available bandwidth).
    // Callbacks receive an in/out FileSystemOptions* so the host can tweak it on progress updates (e.g. changing the limit mid-flight).
    // Plugins MAY also write back an effective applied limit (e.g. internal clamping or combining with a plugin-specific cap).
    unsigned __int64 bandwidthLimitBytesPerSecond;
};

struct FileSystemRenamePair
{
    // Pointers reference NUL-terminated UTF-16 strings stored in a caller-owned arena.
    // Arrays of FileSystemRenamePair are allocated from the same arena as their strings.
    const wchar_t* sourcePath;
    const wchar_t* newName; // Leaf name only (no path separators).
};

// All pointer fields in FileSystemRenamePair, FileSystemSearchQuery, FileSystemSearchMatch, and callback string
// parameters must be arena-backed UTF-16 strings.
// Arrays passed to CopyItems/MoveItems/DeleteItems and arrays of FileSystemRenamePair must be allocated from the same
// arena as their strings.
// Arena strings are NUL-terminated.
struct FileSystemArena
{
    unsigned char* buffer;
    unsigned long capacityBytes;
    unsigned long usedBytes;
};

enum FileSystemSearchFlags : uint32_t
{
    FILESYSTEM_SEARCH_NONE                = 0,
    FILESYSTEM_SEARCH_RECURSIVE           = 0x1,
    FILESYSTEM_SEARCH_INCLUDE_FILES       = 0x2,
    FILESYSTEM_SEARCH_INCLUDE_DIRECTORIES = 0x4,
    FILESYSTEM_SEARCH_MATCH_CASE          = 0x8,
    FILESYSTEM_SEARCH_FOLLOW_SYMLINKS     = 0x10,
    FILESYSTEM_SEARCH_USE_REGEX           = 0x20,
};

struct FileSystemSearchQuery
{
    const wchar_t* rootPath;
    const wchar_t* pattern; // nullptr/empty = L"*"
    FileSystemSearchFlags flags;
    unsigned long maxResults; // 0 = unlimited
};

struct FileSystemSearchMatch
{
    const wchar_t* fullPath;
    unsigned long fullPathSize;
    unsigned long fileAttributes;
    __int64 creationTime;
    __int64 lastAccessTime;
    __int64 lastWriteTime;
    __int64 changeTime;
    __int64 endOfFile;
    __int64 allocationSize;
};

struct FileSystemSearchProgress
{
    unsigned __int64 scannedEntries;
    unsigned __int64 matchedEntries;
    const wchar_t* currentPath;
};
#pragma warning(pop)

// Keeps a complete listing of directory entries in memory as a contiguous buffer of FileInfo structs.
interface __declspec(uuid("0d9ef549-4e54-4086-8a5c-f9d3e6120211")) __declspec(novtable) IFilesInformation : public IUnknown
{
    // Returns the head of a contiguous buffer containing FileInfo entries linked by NextEntryOffset.
    // The buffer is owned by the IFilesInformation instance; the caller MUST NOT free it.
    // If there are no entries, *ppFileInfo is set to nullptr and S_OK is returned.
    virtual HRESULT STDMETHODCALLTYPE GetBuffer(FileInfo * *ppFileInfo) noexcept = 0;
    // Returns how many bytes in the buffer are committed/used by the current result set.
    virtual HRESULT STDMETHODCALLTYPE GetBufferSize(unsigned long* pSize) noexcept = 0;
    // Returns the allocated capacity of the backing buffer in bytes.
    // This may be larger than the committed/used bytes for the current result set.
    virtual HRESULT STDMETHODCALLTYPE GetAllocatedSize(unsigned long* pSize) noexcept = 0;
    // Helper methods.
    // Normally you would enumerate the buffer yourself, but these methods are provided for convenience.
    virtual HRESULT STDMETHODCALLTYPE GetCount(unsigned long* pCount) noexcept              = 0;
    virtual HRESULT STDMETHODCALLTYPE Get(unsigned long index, FileInfo** ppEntry) noexcept = 0;
};

// Host callback for file operation progress.
// Notes:
// - This is NOT a COM interface (no IUnknown inheritance); lifetime is managed by the host.
// - The cookie is provided by the host at call time and must be passed back verbatim by the plugin.
// - Plugins MUST NOT invoke these callbacks concurrently for a single operation (the host is not required to be thread-safe).
// - Callbacks may be invoked on background threads.
// - Callbacks may block (e.g. host-driven Pause); plugins SHOULD avoid holding locks that could deadlock if callbacks block,
//   and SHOULD reach progress checkpoints frequently enough for pause/cancel responsiveness.
interface __declspec(novtable) IFileSystemCallback
{
    // options may be nullptr; implementations must check before reading/writing to it.
    // If options is non-null, it is an in/out object:
    // - the host may update fields (e.g. speed limit changes)
    // - plugins may write back an effective applied value (e.g. clamping / combining with internal caps)
    // Plugins SHOULD read options after the callback returns.
    // Notes:
    // - totalItems/totalBytes MAY be 0 if the plugin does not know totals; hosts MAY provide totals via pre-calculation.
    // - completedBytes SHOULD be monotonic when reported (best-effort); it MAY be 0 for operations where bytes are not meaningful.
    // - currentItem*Bytes refer to the in-flight item (typically a file); they MAY be 0 for directory operations or when unknown.
    // - progressStreamId identifies a concurrent progress stream (e.g. a worker). When a plugin executes items in parallel,
    //   each active worker MUST report a distinct progressStreamId. The ID MUST remain stable across progress callbacks for that worker,
    //   even as it advances to new items.
    virtual HRESULT STDMETHODCALLTYPE FileSystemProgress(FileSystemOperation operationType,
                                                         unsigned long totalItems,
                                                         unsigned long completedItems,
                                                         unsigned __int64 totalBytes,
                                                         unsigned __int64 completedBytes,
                                                         const wchar_t* currentSourcePath,
                                                         const wchar_t* currentDestinationPath,
                                                         unsigned __int64 currentItemTotalBytes,
                                                         unsigned __int64 currentItemCompletedBytes,
                                                         FileSystemOptions* options,
                                                         unsigned __int64 progressStreamId,
                                                         void* cookie) noexcept = 0;
    // options may be nullptr; implementations must check before reading/writing to it.
    // Notes:
    // - itemIndex is the logical index of the completed item within the original request array (0..count-1).
    // - Plugins MAY complete items out-of-order when executing in parallel; hosts MUST NOT assume ascending completion order.
    virtual HRESULT STDMETHODCALLTYPE FileSystemItemCompleted(FileSystemOperation operationType,
                                                              unsigned long itemIndex,
                                                              const wchar_t* sourcePath,
                                                              const wchar_t* destinationPath,
                                                              HRESULT status,
                                                              FileSystemOptions* options,
                                                              void* cookie) noexcept                = 0;
    virtual HRESULT STDMETHODCALLTYPE FileSystemShouldCancel(BOOL * pCancel, void* cookie) noexcept = 0;

    // Invoked by plugins when an operation hits a conflict/issue that requires a user decision (retry/skip/etc.).
    // Notes:
    // - sourcePath/destinationPath are best-effort; they may be nullptr for some operations (e.g. delete destination).
    // - action must be non-null. Implementations should set it even when returning failure/cancellation.
    // - This callback may block (host-driven inline conflict UI).
    virtual HRESULT STDMETHODCALLTYPE FileSystemIssue(FileSystemOperation operationType,
                                                      const wchar_t* sourcePath,
                                                      const wchar_t* destinationPath,
                                                      HRESULT status,
                                                      FileSystemIssueAction* action,
                                                      FileSystemOptions* options,
                                                      void* cookie) noexcept = 0;
};

interface __declspec(uuid("12519afa-30e7-4e3a-9db2-7990c4be9a21")) __declspec(novtable) IFileSystem : public IUnknown
{
    // Lists the contents of a directory into an IFilesInformation object.
    // On success, ppFilesInformation receives a valid instance of IFilesInformation.
    virtual HRESULT STDMETHODCALLTYPE ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept = 0;

    virtual HRESULT STDMETHODCALLTYPE CopyItem(const wchar_t* sourcePath,
                                               const wchar_t* destinationPath,
                                               FileSystemFlags flags,
                                               const FileSystemOptions* options = nullptr,
                                               IFileSystemCallback* callback    = nullptr,
                                               void* cookie                     = nullptr) noexcept = 0;

    virtual HRESULT STDMETHODCALLTYPE MoveItem(const wchar_t* sourcePath,
                                               const wchar_t* destinationPath,
                                               FileSystemFlags flags,
                                               const FileSystemOptions* options = nullptr,
                                               IFileSystemCallback* callback    = nullptr,
                                               void* cookie                     = nullptr) noexcept = 0;

    virtual HRESULT STDMETHODCALLTYPE DeleteItem(const wchar_t* path,
                                                 FileSystemFlags flags,
                                                 const FileSystemOptions* options = nullptr,
                                                 IFileSystemCallback* callback    = nullptr,
                                                 void* cookie                     = nullptr) noexcept = 0;

    virtual HRESULT STDMETHODCALLTYPE RenameItem(const wchar_t* sourcePath,
                                                 const wchar_t* destinationPath,
                                                 FileSystemFlags flags,
                                                 const FileSystemOptions* options = nullptr,
                                                 IFileSystemCallback* callback    = nullptr,
                                                 void* cookie                     = nullptr) noexcept = 0;

    virtual HRESULT STDMETHODCALLTYPE CopyItems(const wchar_t* const* sourcePaths,
                                                unsigned long count,
                                                const wchar_t* destinationFolder,
                                                FileSystemFlags flags,
                                                const FileSystemOptions* options = nullptr,
                                                IFileSystemCallback* callback    = nullptr,
                                                void* cookie                     = nullptr) noexcept = 0;

    virtual HRESULT STDMETHODCALLTYPE MoveItems(const wchar_t* const* sourcePaths,
                                                unsigned long count,
                                                const wchar_t* destinationFolder,
                                                FileSystemFlags flags,
                                                const FileSystemOptions* options = nullptr,
                                                IFileSystemCallback* callback    = nullptr,
                                                void* cookie                     = nullptr) noexcept = 0;

    virtual HRESULT STDMETHODCALLTYPE DeleteItems(const wchar_t* const* paths,
                                                  unsigned long count,
                                                  FileSystemFlags flags,
                                                  const FileSystemOptions* options = nullptr,
                                                  IFileSystemCallback* callback    = nullptr,
                                                  void* cookie                     = nullptr) noexcept = 0;

    virtual HRESULT STDMETHODCALLTYPE RenameItems(const FileSystemRenamePair* items,
                                                  unsigned long count,
                                                  FileSystemFlags flags,
                                                  const FileSystemOptions* options = nullptr,
                                                  IFileSystemCallback* callback    = nullptr,
                                                  void* cookie                     = nullptr) noexcept = 0;

    // Optional: returns filesystem capabilities as a UTF-8 JSON document.
    // Notes:
    // - Returned pointers are owned by the plugin and remain valid until the next call to GetCapabilities or object release.
    // - JSON strings are UTF-8, NUL-terminated.
    // - Host-recognized optional shape:
    //   {
    //     "version": 1,
    //     "operations": { ... },
    //     "concurrency": {
    //       "copyMoveMax": 4,
    //       "deleteMax": 8,
    //       "deleteRecycleBinMax": 2
    //     }
    //   }
    //   If "concurrency" is absent, host per-item concurrency falls back to 1.
    // - Implementations SHOULD return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) when unsupported.
    virtual HRESULT STDMETHODCALLTYPE GetCapabilities(const char** jsonUtf8) noexcept = 0;
};

// Minimal Win32-like file reader for filesystem plugins.
// Notes:
// - The reader is read-only.
// - Implementations MUST be safe for large files (64-bit offsets/sizes).
interface __declspec(uuid("b1d0c2b8-0e37-4d6f-8c2c-2cc4f0d1c6b8")) __declspec(novtable) IFileReader : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE GetSize(unsigned __int64* sizeBytes) noexcept                                      = 0;
    virtual HRESULT STDMETHODCALLTYPE Seek(__int64 offset, unsigned long origin, unsigned __int64* newPosition) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE Read(void* buffer, unsigned long bytesToRead, unsigned long* bytesRead) noexcept   = 0;
};

// Minimal Win32-like file writer for filesystem plugins.
// Notes:
// - Implementations MUST be safe for large files (64-bit offsets/sizes).
// - Implementations MUST tolerate being released without Commit() (treat as abort / best-effort cleanup).
interface __declspec(uuid("b6f0a9e1-8c8b-4b72-9f3e-2f2b4b8b9c41")) __declspec(novtable) IFileWriter : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE GetPosition(unsigned __int64* positionBytes) noexcept                                       = 0;
    virtual HRESULT STDMETHODCALLTYPE Write(const void* buffer, unsigned long bytesToWrite, unsigned long* bytesWritten) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE Commit() noexcept                                                                           = 0;
};

struct FileSystemBasicInformation
{
    __int64 creationTime     = 0; // FILETIME ticks (100ns intervals since 1601-01-01 UTC)
    __int64 lastAccessTime   = 0; // FILETIME ticks
    __int64 lastWriteTime    = 0; // FILETIME ticks
    unsigned long attributes = 0; // FILE_ATTRIBUTE_* flags
};

// Optional I/O interface for filesystem plugins.
// Notes:
// - Implementations MUST interpret `path` as a filesystem-internal path (not necessarily a Win32 path).
// - Implementations SHOULD return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND/ERROR_PATH_NOT_FOUND) when the item does not exist.
// - On success, fileAttributes is set to FILE_ATTRIBUTE_* flags (e.g. FILE_ATTRIBUTE_DIRECTORY).
interface __declspec(uuid("2c7c32b3-8a0f-4e25-8d3a-6a5f1d0a1e2c")) __declspec(novtable) IFileSystemIO : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE GetAttributes(const wchar_t* path, unsigned long* fileAttributes) noexcept                    = 0;
    virtual HRESULT STDMETHODCALLTYPE CreateFileReader(const wchar_t* path, IFileReader** reader) noexcept                          = 0;
    virtual HRESULT STDMETHODCALLTYPE CreateFileWriter(const wchar_t* path, FileSystemFlags flags, IFileWriter** writer) noexcept   = 0;
    virtual HRESULT STDMETHODCALLTYPE GetFileBasicInformation(const wchar_t* path, FileSystemBasicInformation* info) noexcept       = 0;
    virtual HRESULT STDMETHODCALLTYPE SetFileBasicInformation(const wchar_t* path, const FileSystemBasicInformation* info) noexcept = 0;

    // Optional: returns item properties as a UTF-8 JSON document.
    // Notes:
    // - Returned pointers are owned by the plugin and remain valid until the next call to GetItemProperties or object release.
    // - JSON strings are UTF-8, NUL-terminated.
    // - Implementations SHOULD return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) when unsupported.
    virtual HRESULT STDMETHODCALLTYPE GetItemProperties(const wchar_t* path, const char** jsonUtf8) noexcept = 0;
};

// Result structure for directory size computation.
struct FileSystemDirectorySizeResult
{
    unsigned __int64 totalBytes;     // Total size in bytes (sum of file sizes).
    unsigned __int64 fileCount;      // Number of files counted.
    unsigned __int64 directoryCount; // Number of directories counted (excluding root).
    HRESULT status;                  // S_OK, HRESULT_FROM_WIN32(ERROR_CANCELLED), or first error.
};

// Host callback for directory size computation progress.
// Notes:
// - This is NOT a COM interface (no IUnknown inheritance); lifetime is managed by the host.
// - The cookie is provided by the host at call time and must be passed back verbatim by the plugin.
// - Callbacks may block (e.g. host-driven Pause/Skip); plugins SHOULD avoid holding locks that could deadlock if callbacks block,
//   and SHOULD reach progress checkpoints frequently enough for responsiveness.
interface __declspec(novtable) IFileSystemDirectorySizeCallback
{
    virtual HRESULT STDMETHODCALLTYPE DirectorySizeProgress(unsigned __int64 scannedEntries,
                                                            unsigned __int64 totalBytes,
                                                            unsigned __int64 fileCount,
                                                            unsigned __int64 directoryCount,
                                                            const wchar_t* currentPath,
                                                            void* cookie) noexcept = 0;

    virtual HRESULT STDMETHODCALLTYPE DirectorySizeShouldCancel(BOOL * pCancel, void* cookie) noexcept = 0;
};

// Optional directory operations interface.
// Notes:
// - The host obtains this interface via QueryInterface on the active IFileSystem instance.
// - Implementations should return HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) when the target already exists.
interface __declspec(uuid("4a8f7cf2-f81c-4278-b182-7183e6bed6f3")) __declspec(novtable) IFileSystemDirectoryOperations : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE CreateDirectory(const wchar_t* path) noexcept = 0;

    // Compute the total size of a directory.
    // - path: Root item to start from.
    //   - If path is a directory: compute directory size (recursive or immediate children based on flags).
    //   - If path is a file: return file-root sizing (totalBytes=file size, fileCount=1, directoryCount=0, status=S_OK).
    // - flags: Use FILESYSTEM_FLAG_RECURSIVE for recursive computation; otherwise only immediate children.
    // - callback: Optional progress callback (may be nullptr for synchronous completion).
    // - cookie: Opaque value passed to callback.
    // - result: [out] Output result structure.
    // Returns: S_OK on success, HRESULT_FROM_WIN32(ERROR_CANCELLED) if cancelled via callback.
    virtual HRESULT STDMETHODCALLTYPE GetDirectorySize(const wchar_t* path,
                                                       FileSystemFlags flags,
                                                       IFileSystemDirectorySizeCallback* callback,
                                                       void* cookie,
                                                       FileSystemDirectorySizeResult* result) noexcept = 0;
};

// Directory watch actions (best-effort; plugins may coalesce or drop events).
enum FileSystemDirectoryChangeAction : uint32_t
{
    FILESYSTEM_DIR_CHANGE_UNKNOWN          = 0,
    FILESYSTEM_DIR_CHANGE_ADDED            = 1,
    FILESYSTEM_DIR_CHANGE_REMOVED          = 2,
    FILESYSTEM_DIR_CHANGE_MODIFIED         = 3,
    FILESYSTEM_DIR_CHANGE_RENAMED_OLD_NAME = 4,
    FILESYSTEM_DIR_CHANGE_RENAMED_NEW_NAME = 5,
};

struct FileSystemDirectoryChange
{
    FileSystemDirectoryChangeAction action;
    // Relative path to the watched folder; NOT required to be NUL-terminated.
    const wchar_t* relativePath;
    unsigned long relativePathSize; // bytes (not characters)
};

struct FileSystemDirectoryChangeNotification
{
    // Path originally passed to WatchDirectory; NUL-terminated UTF-16.
    const wchar_t* watchedPath;
    unsigned long watchedPathSize; // bytes (not characters)

    const FileSystemDirectoryChange* changes;
    unsigned long changeCount;
    // TRUE if changes were dropped/coalesced (OS overflow, internal caps, parse failure, queue pressure, etc.).
    // If overflow is TRUE, incremental events are not trustworthy and the host SHOULD perform a full resync of the watched folder.
    BOOL overflow;
};

// Host callback for directory watch notifications.
// Notes:
// - This is NOT a COM interface (no IUnknown inheritance); lifetime is managed by the host.
// - The cookie is provided by the host at WatchDirectory time and must be passed back verbatim by the plugin.
// - Plugins MUST NOT invoke these callbacks concurrently for a single watch registration (the host is not required to be thread-safe).
// - Callbacks may be invoked on background threads.
interface __declspec(novtable) IFileSystemDirectoryWatchCallback
{
    virtual HRESULT STDMETHODCALLTYPE FileSystemDirectoryChanged(const FileSystemDirectoryChangeNotification* notification, void* cookie) noexcept = 0;
};

// Optional directory watch interface for plugins that can report change notifications.
// Notes:
// - The host obtains this interface via QueryInterface on the active IFileSystem instance.
// - UnwatchDirectory MUST guarantee no callbacks for that path after it returns.
interface __declspec(uuid("d00f72a2-faf2-47c4-abbe-85dab1e67132")) __declspec(novtable) IFileSystemDirectoryWatch : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE WatchDirectory(const wchar_t* path, IFileSystemDirectoryWatchCallback* callback, void* cookie) noexcept = 0;

    virtual HRESULT STDMETHODCALLTYPE UnwatchDirectory(const wchar_t* path) noexcept = 0;
};

// Optional per-instance initialization interface.
// Implementations can use this to accept a "root" context (e.g. archive path, remote endpoint)
// and an optional JSON/JSON5 options payload (e.g. password, initial directory).
interface __declspec(uuid("a4bdbb56-4f3f-4c1b-9b28-2f4c4a08d7af")) __declspec(novtable) IFileSystemInitialize : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE Initialize(const wchar_t* rootPath, const char* optionsJsonUtf8) noexcept = 0;
};

// Host callback for file system search results.
// Notes:
// - This is NOT a COM interface (no IUnknown inheritance); lifetime is managed by the host.
// - The cookie is provided by the host at call time and must be passed back verbatim by the plugin.
interface __declspec(novtable) IFileSystemSearchCallback
{
    virtual HRESULT STDMETHODCALLTYPE FileSystemSearchMatch(const FileSystemSearchMatch* match, void* cookie) noexcept          = 0;
    virtual HRESULT STDMETHODCALLTYPE FileSystemSearchProgress(const FileSystemSearchProgress* progress, void* cookie) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE FileSystemSearchShouldCancel(BOOL * pCancel, void* cookie) noexcept                       = 0;
};

interface __declspec(uuid("00417f3e-f0f5-4add-8dea-4407d5169ef6")) __declspec(novtable) IFileSystemSearch : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE Search(const FileSystemSearchQuery* query, IFileSystemSearchCallback* callback, void* cookie) noexcept = 0;
};

inline HRESULT InitializeFileSystemArena(FileSystemArena* arena, unsigned long capacityBytes) noexcept
{
    if (! arena)
    {
        return E_POINTER;
    }

    if (arena->buffer != nullptr)
    {
        return E_INVALIDARG;
    }

    if (capacityBytes == 0)
    {
        arena->buffer        = nullptr;
        arena->capacityBytes = 0;
        arena->usedBytes     = 0;
        return S_OK;
    }

    arena->buffer = static_cast<unsigned char*>(::HeapAlloc(::GetProcessHeap(), 0, capacityBytes));
    if (! arena->buffer)
    {
        return E_OUTOFMEMORY;
    }

    arena->capacityBytes = capacityBytes;
    arena->usedBytes     = 0;
    return S_OK;
}

inline void DestroyFileSystemArena(FileSystemArena* arena) noexcept
{
    if (! arena)
    {
        return;
    }

    if (arena->buffer)
    {
        ::HeapFree(::GetProcessHeap(), 0, arena->buffer);
        arena->buffer = nullptr;
    }

    arena->capacityBytes = 0;
    arena->usedBytes     = 0;
}

inline void* AllocateFromFileSystemArena(FileSystemArena* arena, unsigned long sizeBytes, unsigned long alignment) noexcept
{
    if (! arena || ! arena->buffer || sizeBytes == 0)
    {
        return nullptr;
    }

    if (alignment == 0 || (alignment & (alignment - 1u)) != 0u)
    {
        return nullptr;
    }

    const unsigned long mask     = alignment - 1u;
    const unsigned long aligned  = (arena->usedBytes + mask) & ~mask;
    const unsigned long capacity = arena->capacityBytes;

    if (aligned > capacity || sizeBytes > capacity - aligned)
    {
        return nullptr;
    }

    void* result     = arena->buffer + aligned;
    arena->usedBytes = aligned + sizeBytes;
    return result;
}

// Builds an arena containing a const wchar_t* path array for every entry in filesInformation.
inline HRESULT BuildFileSystemPathListArenaFromFilesInformation(
    const wchar_t* sourceRoot, IFilesInformation* filesInformation, FileSystemArena* arena, const wchar_t*** outPaths, unsigned long* outCount) noexcept
{
    if (! sourceRoot || ! filesInformation || ! arena || ! outPaths || ! outCount)
    {
        return E_POINTER;
    }

    if (arena->buffer != nullptr)
    {
        return E_INVALIDARG;
    }

    unsigned long entryCount = 0;
    HRESULT hr               = filesInformation->GetCount(&entryCount);
    if (FAILED(hr))
    {
        return hr;
    }

    if (entryCount == 0)
    {
        *outPaths = nullptr;
        *outCount = 0;
        return S_OK;
    }

    FileInfo* buffer = nullptr;
    hr               = filesInformation->GetBuffer(&buffer);
    if (FAILED(hr))
    {
        return hr;
    }

    if (! buffer)
    {
        return E_POINTER;
    }

    unsigned long bufferSize = 0;
    hr                       = filesInformation->GetBufferSize(&bufferSize);
    if (FAILED(hr))
    {
        return hr;
    }

    if (bufferSize == 0)
    {
        return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
    }

    const unsigned long wcharSizeBytes = static_cast<unsigned long>(sizeof(wchar_t));

    const size_t sourceRootSize = ::wcslen(sourceRoot);
    if (sourceRootSize > ULONG_MAX)
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    const unsigned long sourceRootChars = static_cast<unsigned long>(sourceRootSize);
    bool sourceNeedsSeparator           = false;
    if (sourceRootChars > 0)
    {
        const wchar_t lastChar = sourceRoot[sourceRootChars - 1];
        if (lastChar != L'\\' && lastChar != L'/')
        {
            sourceNeedsSeparator = true;
        }
    }

    unsigned __int64 totalBytes = static_cast<unsigned __int64>(entryCount) * sizeof(const wchar_t*);
    if (totalBytes > ULONG_MAX)
    {
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    unsigned char* bufferBytes = reinterpret_cast<unsigned char*>(buffer);
    FileInfo* entry            = buffer;
    unsigned long offset       = 0;

    for (unsigned long index = 0; index < entryCount; ++index)
    {
        if (! entry)
        {
            return E_POINTER;
        }

        if ((entry->FileNameSize % wcharSizeBytes) != 0u)
        {
            return E_INVALIDARG;
        }

        const unsigned long nameChars = entry->FileNameSize / wcharSizeBytes;
        unsigned __int64 sourceChars  = static_cast<unsigned __int64>(sourceRootChars);
        if (sourceNeedsSeparator)
        {
            sourceChars += 1u;
        }

        sourceChars += static_cast<unsigned __int64>(nameChars);
        if (sourceChars > ULONG_MAX)
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        unsigned __int64 sourceBytes = (sourceChars + 1u) * wcharSizeBytes;
        totalBytes += sourceBytes;
        if (totalBytes > ULONG_MAX)
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        if (entry->NextEntryOffset == 0)
        {
            if (index + 1u < entryCount)
            {
                return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
            }

            break;
        }

        if (entry->NextEntryOffset > bufferSize - offset)
        {
            return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
        }

        offset += entry->NextEntryOffset;
        if (offset >= bufferSize)
        {
            return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
        }

        entry = reinterpret_cast<FileInfo*>(bufferBytes + offset);
    }

    hr = InitializeFileSystemArena(arena, static_cast<unsigned long>(totalBytes));
    if (FAILED(hr))
    {
        return hr;
    }

    const unsigned __int64 pathsBytes64 = static_cast<unsigned __int64>(entryCount) * sizeof(const wchar_t*);
    if (pathsBytes64 > ULONG_MAX)
    {
        DestroyFileSystemArena(arena);
        return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
    }

    const unsigned long pathsBytes = static_cast<unsigned long>(pathsBytes64);
    const wchar_t** paths = static_cast<const wchar_t**>(AllocateFromFileSystemArena(arena, pathsBytes, static_cast<unsigned long>(alignof(const wchar_t*))));
    if (! paths)
    {
        DestroyFileSystemArena(arena);
        return E_OUTOFMEMORY;
    }

    entry  = buffer;
    offset = 0;

    for (unsigned long index = 0; index < entryCount; ++index)
    {
        if (! entry)
        {
            DestroyFileSystemArena(arena);
            return E_POINTER;
        }

        const unsigned long nameChars = entry->FileNameSize / wcharSizeBytes;
        unsigned long sourceChars     = sourceRootChars;
        if (sourceNeedsSeparator)
        {
            sourceChars += 1u;
        }

        sourceChars += nameChars;
        const unsigned __int64 sourceBytes64 = (static_cast<unsigned __int64>(sourceChars) + 1u) * wcharSizeBytes;
        if (sourceBytes64 > ULONG_MAX)
        {
            DestroyFileSystemArena(arena);
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        const unsigned long sourceBytes = static_cast<unsigned long>(sourceBytes64);
        wchar_t* sourcePath             = static_cast<wchar_t*>(AllocateFromFileSystemArena(arena, sourceBytes, static_cast<unsigned long>(alignof(wchar_t))));
        if (! sourcePath)
        {
            DestroyFileSystemArena(arena);
            return E_OUTOFMEMORY;
        }

        unsigned long pathOffset = 0;
        if (sourceRootChars > 0)
        {
            const SIZE_T rootBytes = static_cast<SIZE_T>(sourceRootChars) * wcharSizeBytes;
            ::CopyMemory(sourcePath, sourceRoot, rootBytes);
            pathOffset = sourceRootChars;
        }

        if (sourceNeedsSeparator)
        {
            sourcePath[pathOffset] = L'\\';
            pathOffset += 1u;
        }

        if (nameChars > 0)
        {
            const SIZE_T nameBytes = static_cast<SIZE_T>(entry->FileNameSize);
            ::CopyMemory(sourcePath + pathOffset, entry->FileName, nameBytes);
        }

        sourcePath[pathOffset + nameChars] = L'\0';
        paths[index]                       = sourcePath;

        if (entry->NextEntryOffset == 0)
        {
            if (index + 1u < entryCount)
            {
                DestroyFileSystemArena(arena);
                return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
            }

            break;
        }

        if (entry->NextEntryOffset > bufferSize - offset)
        {
            DestroyFileSystemArena(arena);
            return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
        }

        offset += entry->NextEntryOffset;
        if (offset >= bufferSize)
        {
            DestroyFileSystemArena(arena);
            return HRESULT_FROM_WIN32(ERROR_BAD_LENGTH);
        }

        entry = reinterpret_cast<FileInfo*>(bufferBytes + offset);
    }

    *outPaths = paths;
    *outCount = entryCount;
    return S_OK;
}

class FileSystemArenaOwner final
{
public:
    FileSystemArenaOwner() noexcept = default;

    ~FileSystemArenaOwner() noexcept
    {
        DestroyFileSystemArena(&_arena);
    }

    FileSystemArenaOwner(const FileSystemArenaOwner&)            = delete;
    FileSystemArenaOwner& operator=(const FileSystemArenaOwner&) = delete;

    FileSystemArenaOwner(FileSystemArenaOwner&& other) noexcept
    {
        _arena                     = other._arena;
        other._arena.buffer        = nullptr;
        other._arena.capacityBytes = 0;
        other._arena.usedBytes     = 0;
    }

    FileSystemArenaOwner& operator=(FileSystemArenaOwner&& other) noexcept
    {
        if (this == &other)
        {
            return *this;
        }

        DestroyFileSystemArena(&_arena);
        _arena                     = other._arena;
        other._arena.buffer        = nullptr;
        other._arena.capacityBytes = 0;
        other._arena.usedBytes     = 0;
        return *this;
    }

    FileSystemArena* Get() noexcept
    {
        return &_arena;
    }

    const FileSystemArena* Get() const noexcept
    {
        return &_arena;
    }

    void Reset() noexcept
    {
        DestroyFileSystemArena(&_arena);
    }

    HRESULT Initialize(unsigned long capacityBytes) noexcept
    {
        DestroyFileSystemArena(&_arena);
        return InitializeFileSystemArena(&_arena, capacityBytes);
    }

    HRESULT BuildPathListFromFilesInformation(const wchar_t* sourceRoot,
                                              IFilesInformation* filesInformation,
                                              const wchar_t*** outPaths,
                                              unsigned long* outCount) noexcept
    {
        DestroyFileSystemArena(&_arena);
        return BuildFileSystemPathListArenaFromFilesInformation(sourceRoot, filesInformation, &_arena, outPaths, outCount);
    }

private:
    FileSystemArena _arena{};
};
