#pragma once

#include <cstdint>
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

// ABI versioning note:
// For any struct that includes a `sizeBytes` field, the creator MUST set `sizeBytes = sizeof(StructName)` before
// passing it across the host<->plugin boundary. Consumers MUST validate `sizeBytes` before reading other fields.
// In this repo's ABI-breaking sweep, mismatched `sizeBytes` is treated as a contract violation: fail the call with `E_INVALIDARG`.
// For [out] structs, the caller MUST initialize `sizeBytes` before calling into the callee, and the callee MUST NOT
// write beyond `sizeBytes`.
struct FileSystemOptions
{
    uint32_t sizeBytes; // sizeof(FileSystemOptions)

    // 0 = unlimited (use all available bandwidth).
    // Callbacks receive an in/out FileSystemOptions* so the host can tweak it on progress updates (e.g. changing the limit mid-flight).
    // Plugins MAY also write back an effective applied limit (e.g. internal clamping or combining with a plugin-specific cap).
    uint64_t bandwidthLimitBytesPerSecond;
};

enum FileSystemTransferEndpoint : uint32_t
{
    FILESYSTEM_TRANSFER_SOURCE_READ       = 1,
    FILESYSTEM_TRANSFER_DESTINATION_WRITE = 2,
};

enum FileSystemTransferLatencyClass : uint32_t
{
    FILESYSTEM_TRANSFER_LATENCY_UNKNOWN = 0,
    FILESYSTEM_TRANSFER_LATENCY_LOCAL   = 1,
    FILESYSTEM_TRANSFER_LATENCY_LAN     = 2,
    FILESYSTEM_TRANSFER_LATENCY_WAN     = 3,
    FILESYSTEM_TRANSFER_LATENCY_CLOUD   = 4,
};

enum FileSystemTransferHintFlags : uint32_t
{
    FILESYSTEM_TRANSFER_HINT_NONE                  = 0,
    FILESYSTEM_TRANSFER_HINT_PREFERS_LARGE_BUFFERS = 0x1,
    FILESYSTEM_TRANSFER_HINT_PREFERS_SEQUENTIAL_IO = 0x2,
    FILESYSTEM_TRANSFER_HINT_HIGH_METADATA_COST    = 0x4,
};

struct FileSystemTransferHints
{
    uint32_t sizeBytes; // sizeof(FileSystemTransferHints)

    uint32_t latencyClass;              // FileSystemTransferLatencyClass
    uint32_t flags;                     // FileSystemTransferHintFlags
    uint32_t preferredBufferBytes;      // e.g. 2 MiB, 4 MiB, 8 MiB
    uint32_t preferredProgressPeriodMs; // e.g. 100..250ms
};

enum FileSystemStorageKind : uint32_t
{
    FILESYSTEM_STORAGE_UNKNOWN       = 0,
    FILESYSTEM_STORAGE_HDD           = 1,
    FILESYSTEM_STORAGE_SSD           = 2,
    FILESYSTEM_STORAGE_NVME          = 3,
    FILESYSTEM_STORAGE_NETWORK_SHARE = 4,
    FILESYSTEM_STORAGE_CLOUD         = 5,
    FILESYSTEM_STORAGE_VIRTUAL       = 6,
    FILESYSTEM_STORAGE_MEMORY        = 7,
};

enum FileSystemStorageFlags : uint32_t
{
    FILESYSTEM_STORAGE_FLAG_NONE                  = 0,
    FILESYSTEM_STORAGE_FLAG_ROTATIONAL            = 0x1,
    FILESYSTEM_STORAGE_FLAG_HIGH_LATENCY          = 0x2,
    FILESYSTEM_STORAGE_FLAG_PREFERS_SEQUENTIAL_IO = 0x4,
    FILESYSTEM_STORAGE_FLAG_SUPPORTS_DEEP_QUEUE   = 0x8,
};

struct FileSystemStorageCharacteristics
{
    uint32_t sizeBytes; // sizeof(FileSystemStorageCharacteristics)

    uint32_t storageKind; // FileSystemStorageKind
    uint32_t flags;       // FileSystemStorageFlags
    uint32_t queueDepthHint;
    uint32_t preferredCopyMoveConcurrency;
    uint32_t preferredDeleteConcurrency;
};

struct FileSystemRenamePair
{
    uint32_t sizeBytes; // sizeof(FileSystemRenamePair)

    // Pointers reference NUL-terminated UTF-16 strings stored in a caller-owned arena.
    // Arrays of FileSystemRenamePair are allocated from the same arena as their strings.
    const wchar_t* sourcePath;
    const wchar_t* newName; // Leaf name only (no path separators).
};

// Arrays passed to CopyItems/MoveItems/DeleteItems and arrays of FileSystemRenamePair must be allocated from the same
// arena as their referenced UTF-16 strings.
// Arena strings are NUL-terminated.
// Search query strings are caller-owned and only need to remain valid for the duration of Search().
// Search match/progress strings are plugin-owned and only need to remain valid until the callback returns.
// Search payload strings are not required to come from FileSystemArena.
struct FileSystemArena
{
    unsigned char* buffer;
    unsigned long capacityBytes;
    unsigned long usedBytes;
};

// Per-call search options and backend hints.
enum FileSystemSearchFlags : uint32_t
{
    FILESYSTEM_SEARCH_NONE                = 0,
    FILESYSTEM_SEARCH_RECURSIVE           = 0x1,
    FILESYSTEM_SEARCH_INCLUDE_FILES       = 0x2,
    FILESYSTEM_SEARCH_INCLUDE_DIRECTORIES = 0x4,
    FILESYSTEM_SEARCH_FOLLOW_SYMLINKS     = 0x8,
    FILESYSTEM_SEARCH_MATCH_CASE_NAME     = 0x10,
    FILESYSTEM_SEARCH_MATCH_CASE_CONTENT  = 0x20,
    FILESYSTEM_SEARCH_WANT_SNIPPETS       = 0x40,
    FILESYSTEM_SEARCH_PREFER_INDEX        = 0x80,
    FILESYSTEM_SEARCH_FORCE_SCAN          = 0x100,
};

enum FileSystemSearchNameMode : uint32_t
{
    FILESYSTEM_SEARCH_NAME_DISABLED = 0,
    FILESYSTEM_SEARCH_NAME_WILDCARD = 1,
    FILESYSTEM_SEARCH_NAME_LITERAL  = 2,
    FILESYSTEM_SEARCH_NAME_REGEX    = 3,
};

enum FileSystemSearchContentMode : uint32_t
{
    FILESYSTEM_SEARCH_CONTENT_DISABLED     = 0,
    FILESYSTEM_SEARCH_CONTENT_TEXT_LITERAL = 1,
    FILESYSTEM_SEARCH_CONTENT_TEXT_REGEX   = 2,
};

enum FileSystemSearchBackend : uint32_t
{
    FILESYSTEM_SEARCH_BACKEND_UNKNOWN = 0,
    FILESYSTEM_SEARCH_BACKEND_SCAN    = 1,
    FILESYSTEM_SEARCH_BACKEND_INDEX   = 2,
    FILESYSTEM_SEARCH_BACKEND_SERVICE = 3,
};

enum FileSystemSearchMatchSource : uint32_t
{
    FILESYSTEM_SEARCH_MATCH_SOURCE_NONE    = 0,
    FILESYSTEM_SEARCH_MATCH_SOURCE_NAME    = 0x1,
    FILESYSTEM_SEARCH_MATCH_SOURCE_CONTENT = 0x2,
};

enum FileSystemSearchWarningFlags : uint32_t
{
    FILESYSTEM_SEARCH_WARNING_NONE                  = 0,
    FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_INDEX     = 0x1,
    FILESYSTEM_SEARCH_WARNING_DEGRADED_NO_CONTENT   = 0x2,
    FILESYSTEM_SEARCH_WARNING_ACCESS_DENIED_SKIPPED = 0x4,
    FILESYSTEM_SEARCH_WARNING_OVERFLOW              = 0x8,
    FILESYSTEM_SEARCH_WARNING_SERVICE_UNAVAILABLE   = 0x10,
    FILESYSTEM_SEARCH_WARNING_REGEX_REJECTED        = 0x20,
};

enum FileSystemSearchPhase : uint32_t
{
    FILESYSTEM_SEARCH_PHASE_INITIALIZING = 0,
    FILESYSTEM_SEARCH_PHASE_ENUMERATING  = 1,
    FILESYSTEM_SEARCH_PHASE_INDEX_LOOKUP = 2,
    FILESYSTEM_SEARCH_PHASE_CONTENT_SCAN = 3,
    FILESYSTEM_SEARCH_PHASE_COMPLETED    = 4,
};

// Search query payload passed to IFileSystemSearch::Search.
// sizeBytes must equal sizeof(FileSystemSearchQuery).
struct FileSystemSearchQuery
{
    uint32_t sizeBytes; // Must equal sizeof(FileSystemSearchQuery)

    const wchar_t* rootPath;                 // Required, NUL-terminated plugin path used as the search root.
    const wchar_t* namePattern;              // Required unless nameMode == FILESYSTEM_SEARCH_NAME_DISABLED.
    const wchar_t* contentPattern;           // Required unless contentMode == FILESYSTEM_SEARCH_CONTENT_DISABLED.
    FileSystemSearchFlags flags;             // Include/recurse/matching options and backend hints.
    FileSystemSearchNameMode nameMode;       // Wildcard, literal substring, regex, or disabled.
    FileSystemSearchContentMode contentMode; // Text literal, text regex, or disabled.
    uint64_t maxResults;                     // 0 = unlimited.
    uint64_t maxContentBytesPerFile;         // 0 = plugin default (64 MiB in the current built-in file plugin).
    uint32_t maxSnippetCharacters;           // 0 = plugin default (160 UTF-16 code units in the current built-in file plugin).
    uint32_t reserved;                       // Must be 0 for normal v1 callers; built-in host extensions use FILESYSTEM_SEARCH_HOST_EXTENSIONS_V1.
};

// Match payload delivered through IFileSystemSearchCallback::FileSystemSearchMatch.
// sizeBytes must equal sizeof(FileSystemSearchMatch).
struct FileSystemSearchMatch
{
    uint32_t sizeBytes; // Must equal sizeof(FileSystemSearchMatch)

    const wchar_t* fullPath;         // Canonical full path for the current plugin instance.
    unsigned long fullPathSize;      // Size in bytes, excluding the trailing NUL.
    const wchar_t* relativePath;     // Root-relative path, useful for result grouping and display.
    unsigned long relativePathSize;  // Size in bytes, excluding the trailing NUL.
    const wchar_t* displayName;      // Leaf display name.
    unsigned long displayNameSize;   // Size in bytes, excluding the trailing NUL.
    const wchar_t* previewText;      // Optional UTF-16 snippet around the first content hit.
    unsigned long previewTextSize;   // Size in bytes, excluding the trailing NUL.
    unsigned long fileAttributes;    // FILE_ATTRIBUTE_* bits.
    __int64 creationTime;            // FILETIME ticks.
    __int64 lastAccessTime;          // FILETIME ticks.
    __int64 lastWriteTime;           // FILETIME ticks.
    __int64 changeTime;              // FILETIME ticks.
    __int64 endOfFile;               // Logical file size in bytes, or 0 for directories/unknown.
    __int64 allocationSize;          // Allocated size in bytes, or 0 when unknown.
    uint32_t matchedBy;              // FileSystemSearchMatchSource bits.
    uint64_t contentMatchByteOffset; // Best-effort byte/character offset of the first content hit, or 0 when unavailable.
    uint32_t contentMatchByteLength; // Best-effort byte/character length of the first content hit, or 0 when unavailable.
    uint32_t reserved;               // Must be 0 for v1 implementations.
};

// Progress payload delivered through IFileSystemSearchCallback::FileSystemSearchProgress.
// sizeBytes must equal sizeof(FileSystemSearchProgress).
struct FileSystemSearchProgress
{
    uint32_t sizeBytes; // Must equal sizeof(FileSystemSearchProgress)

    FileSystemSearchPhase phase;     // Current logical stage of execution.
    FileSystemSearchBackend backend; // Backend currently producing results/progress.
    uint32_t warningFlags;           // FileSystemSearchWarningFlags bits.
    HRESULT statusHint;              // S_OK, S_FALSE, HRESULT_FROM_WIN32(ERROR_CANCELLED), etc.
    uint64_t scannedDirectories;     // Number of directories enumerated so far.
    uint64_t scannedFiles;           // Number of files examined so far.
    uint64_t candidateFiles;         // Number of files selected for content scanning.
    uint64_t matchedEntries;         // Number of matches emitted so far.
    const wchar_t* currentPath;      // Optional current path; may be nullptr for final completion updates.
    unsigned long currentPathSize;   // Size in bytes, excluding the trailing NUL.
};

// Internal host-extension payloads for the built-in local file-system search path.
// These are opt-in and do not change the FileSystemSearchProgress ABI.
inline constexpr uint32_t FILESYSTEM_SEARCH_HOST_EXTENSIONS_V1 = 0x52534631u; // "RSF1"

struct FileSystemSearchServiceStatus
{
    uint32_t sizeBytes; // Must equal sizeof(FileSystemSearchServiceStatus)

    uint32_t storeState;
    uint32_t syncPhase;
    uint32_t queryExecutionMode;
    uint32_t fallbackReason;
    uint64_t completedRoots;
    uint64_t totalRoots;
    const wchar_t* activeRoot;    // Optional UTF-16 path; may be nullptr.
    unsigned long activeRootSize; // Size in bytes, excluding the trailing NUL.
};

using FileSystemSearchServiceStatusCallbackFn = HRESULT(STDMETHODCALLTYPE*)(const FileSystemSearchServiceStatus* status, void* cookie) noexcept;

struct FileSystemSearchHostExtensions
{
    uint32_t sizeBytes; // Must equal sizeof(FileSystemSearchHostExtensions)
    uint32_t version;   // FILESYSTEM_SEARCH_HOST_EXTENSIONS_V1
    void* callbackCookie;
    FileSystemSearchServiceStatusCallbackFn serviceStatusCallback;
    void* serviceStatusCookie;
};
#pragma warning(pop)

// Keeps a complete listing of directory entries in memory as a contiguous buffer of FileInfo structs.
interface __declspec(uuid("0d9ef549-4e54-4086-8a5c-f9d3e6120211")) __declspec(novtable) IFilesInformation : public IUnknown
{
    // Returns the head of a contiguous buffer containing FileInfo entries linked by NextEntryOffset.
    // The buffer is owned by the IFilesInformation instance; the caller MUST NOT free it.
    // The returned pointer remains valid until the IFilesInformation instance is released.
    // Implementations MUST NOT mutate/reallocate the buffer after returning it (no async mutation after ReadDirectoryInfo returns).
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
// - This is a per-call callback passed to Copy*/Move*/Delete*/Rename* operations.
// - Implementations MUST NOT invoke these callbacks after the operation returns.
// - Plugins MUST NOT invoke these callbacks concurrently for a single operation (the host is not required to be thread-safe).
// - Callbacks may be invoked on background threads.
// - Callbacks may block (e.g. host-driven Pause); plugins SHOULD avoid holding locks that could deadlock if callbacks block,
//   and SHOULD reach progress checkpoints frequently enough for pause/cancel responsiveness.
// Host obligations:
// - The host MUST keep the callback object and its backing state alive until the operation call returns.
// - To tear down early, the host MUST signal cancellation via FileSystemShouldCancel and wait for the operation
//   to return before destroying callback-referenced state.
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
                                                         uint64_t totalBytes,
                                                         uint64_t completedBytes,
                                                         const wchar_t* currentSourcePath,
                                                         const wchar_t* currentDestinationPath,
                                                         uint64_t currentItemTotalBytes,
                                                         uint64_t currentItemCompletedBytes,
                                                         FileSystemOptions* options,
                                                         uint64_t progressStreamId,
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

    // Transfer hints for cross-filesystem bridge buffering and progress cadence.
    // - path is interpreted in the plugin's own path space.
    // - operationType is the top-level file operation (copy/move).
    // - endpoint identifies whether the host is reading from the source side or writing to the destination side.
    // - Caller must initialize hints->sizeBytes before calling.
    virtual HRESULT STDMETHODCALLTYPE GetTransferHints(
        const wchar_t* path, FileSystemOperation operationType, FileSystemTransferEndpoint endpoint, FileSystemTransferHints* hints) noexcept = 0;

    // Storage classification used by host-side auto-concurrency and diagnostics.
    // - path is interpreted in the plugin's own path space.
    // - Caller must initialize characteristics->sizeBytes before calling.
    virtual HRESULT STDMETHODCALLTYPE GetStorageCharacteristics(const wchar_t* path, FileSystemStorageCharacteristics* characteristics) noexcept = 0;
};

// Minimal Win32-like file reader for filesystem plugins.
// Notes:
// - The reader is read-only.
// - Implementations MUST be safe for large files (64-bit offsets/sizes).
interface __declspec(uuid("b1d0c2b8-0e37-4d6f-8c2c-2cc4f0d1c6b8")) __declspec(novtable) IFileReader : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE GetSize(uint64_t* sizeBytes) noexcept                                            = 0;
    virtual HRESULT STDMETHODCALLTYPE Seek(__int64 offset, unsigned long origin, uint64_t* newPosition) noexcept       = 0;
    virtual HRESULT STDMETHODCALLTYPE Read(void* buffer, unsigned long bytesToRead, unsigned long* bytesRead) noexcept = 0;
};

// Minimal Win32-like file writer for filesystem plugins.
// Notes:
// - Implementations MUST be safe for large files (64-bit offsets/sizes).
// - Implementations MUST tolerate being released without Commit() (treat as abort / best-effort cleanup).
interface __declspec(uuid("b6f0a9e1-8c8b-4b72-9f3e-2f2b4b8b9c41")) __declspec(novtable) IFileWriter : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE GetPosition(uint64_t* positionBytes) noexcept                                               = 0;
    virtual HRESULT STDMETHODCALLTYPE Write(const void* buffer, unsigned long bytesToWrite, unsigned long* bytesWritten) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE Commit() noexcept                                                                           = 0;
};

struct FileSystemBasicInformation
{
    uint32_t sizeBytes = 0; // sizeof(FileSystemBasicInformation)

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

// Optional item stream operations interface for filesystem plugins.
// Notes:
// - The host obtains this interface via QueryInterface on the active IFileSystem instance.
// - Implementations MUST interpret `path` as a filesystem-internal path.
// - `streamName` is the logical stream name surfaced by GetItemProperties JSON (for example "Zone.Identifier"),
//   not the full Win32 ":name:$DATA" stream spec.
// - Implementations SHOULD return HRESULT_FROM_WIN32(ERROR_NOT_FOUND) when the named stream does not exist.
// - Implementations SHOULD return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED) for backends that can list but not delete streams.
interface __declspec(uuid("9435eb43-828f-43d3-a9a9-8d9c7f7ebe36")) __declspec(novtable) IFileSystemItemStreams : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE DeleteItemStream(const wchar_t* path, const wchar_t* streamName) noexcept = 0;
};

// Result structure for directory size computation.
struct FileSystemDirectorySizeResult
{
    uint32_t sizeBytes; // sizeof(FileSystemDirectorySizeResult)

    uint64_t totalBytes;     // Total size in bytes (sum of file sizes).
    uint64_t fileCount;      // Number of files counted.
    uint64_t directoryCount; // Number of directories counted (excluding root).
    HRESULT status;          // S_OK, HRESULT_FROM_WIN32(ERROR_CANCELLED), or first error.
};

// Host callback for directory size computation progress.
// Notes:
// - This is NOT a COM interface (no IUnknown inheritance); lifetime is managed by the host.
// - The cookie is provided by the host at call time and must be passed back verbatim by the plugin.
// - Callbacks may block (e.g. host-driven Pause/Skip); plugins SHOULD avoid holding locks that could deadlock if callbacks block,
//   and SHOULD reach progress checkpoints frequently enough for responsiveness.
// - This is a per-call callback; the same host obligations apply as for IFileSystemCallback above.
// - Plugins MUST NOT invoke these callbacks concurrently for a single GetDirectorySize call.
interface __declspec(novtable) IFileSystemDirectorySizeCallback
{
    // Notes:
    // - This is a per-call callback passed to IFileSystemDirectoryOperations::GetDirectorySize.
    // - Implementations MUST NOT invoke these callbacks after GetDirectorySize returns.
    virtual HRESULT STDMETHODCALLTYPE DirectorySizeProgress(
        uint64_t scannedEntries, uint64_t totalBytes, uint64_t fileCount, uint64_t directoryCount, const wchar_t* currentPath, void* cookie) noexcept = 0;

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
    uint32_t sizeBytes; // sizeof(FileSystemDirectoryChangeNotification)

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
// Host obligations:
// - The host MUST NOT destroy callback-referenced state until UnwatchDirectory returns.
// - The callback implementation MUST be safe to invoke from any thread at any time before UnwatchDirectory returns.
// - The callback implementation MUST tolerate invocation racing with a teardown request (e.g. check an atomic
//   _stopping flag at entry; see FolderWatcher::OnPluginDirectoryChanged for the reference pattern).
// Deadlock avoidance:
// - Watch callbacks MUST NOT perform synchronous calls that depend on the thread calling UnwatchDirectory.
//   Use PostMessage / TrySubmitThreadpoolCallback, never SendMessage, from a watch callback.
interface __declspec(novtable) IFileSystemDirectoryWatchCallback
{
    virtual HRESULT STDMETHODCALLTYPE FileSystemDirectoryChanged(const FileSystemDirectoryChangeNotification* notification, void* cookie) noexcept = 0;
};

// Optional directory watch interface for plugins that can report change notifications.
// Notes:
// - The host obtains this interface via QueryInterface on the active IFileSystem instance.
// - UnwatchDirectory MUST synchronously drain: after UnwatchDirectory returns, no thread may still invoke callbacks
//   for that path, and any in-flight callback invocation for that registration must have completed.
// - The plugin's drain wait MUST NOT hold a lock that callback delivery also acquires.
//   (FileSystem unlocks _mutex before WaitForThreadpool*Callbacks; FileSystemDummy uses a CV wait that releases the lock.)
// - The _stopping / active flag MUST be set before initiating the drain wait.
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
// - This is a per-call callback passed to IFileSystemSearch::Search.
// - Implementations MUST NOT invoke these callbacks after Search returns.
// - Plugins MUST NOT invoke these callbacks concurrently for a single Search call (the host is not required to be thread-safe).
// - Callbacks may be invoked on background threads.
// - Returning E_ABORT or HRESULT_FROM_WIN32(ERROR_CANCELLED) from FileSystemSearchMatch/FileSystemSearchProgress
//   requests cancellation; Search must then return HRESULT_FROM_WIN32(ERROR_CANCELLED).
// Host obligations:
// - The host MUST keep the callback object and its backing state alive until Search returns.
// - To tear down early, the host MUST signal cancellation via FileSystemSearchShouldCancel and wait for Search
//   to return before destroying callback-referenced state.
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

    uint64_t totalBytes = static_cast<uint64_t>(entryCount) * sizeof(const wchar_t*);
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
        uint64_t sourceChars          = static_cast<uint64_t>(sourceRootChars);
        if (sourceNeedsSeparator)
        {
            sourceChars += 1u;
        }

        sourceChars += static_cast<uint64_t>(nameChars);
        if (sourceChars > ULONG_MAX)
        {
            return HRESULT_FROM_WIN32(ERROR_ARITHMETIC_OVERFLOW);
        }

        uint64_t sourceBytes = (sourceChars + 1u) * wcharSizeBytes;
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

    const uint64_t pathsBytes64 = static_cast<uint64_t>(entryCount) * sizeof(const wchar_t*);
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
        const uint64_t sourceBytes64 = (static_cast<uint64_t>(sourceChars) + 1u) * wcharSizeBytes;
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
