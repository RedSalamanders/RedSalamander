#pragma once

#include <cstddef>
#include <cstdint>
#include <limits.h>
#include <unknwn.h>
#include <wchar.h>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

// ABI versioning note:
// For any struct that includes a `sizeBytes` field, the creator MUST set `sizeBytes = sizeof(StructName)` before
// passing it across the host<->plugin boundary. Consumers MUST validate `sizeBytes` before reading other fields.
// Unless a struct explicitly defines accepted versioned prefixes, mismatched `sizeBytes` is treated as a contract violation:
// fail the call with `E_INVALIDARG`.
// For [out] structs, the caller MUST initialize `sizeBytes` before calling into the callee, and the callee MUST NOT
// write beyond `sizeBytes`.
// 
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
    FILESYSTEM_RENAME           = 4,
    FILESYSTEM_CREATE_DIRECTORY = 5,
};

enum FileSystemFlags : uint32_t
{
    FILESYSTEM_FLAG_NONE                   = 0,
    FILESYSTEM_FLAG_ALLOW_OVERWRITE        = 0x1,
    FILESYSTEM_FLAG_ALLOW_REPLACE_READONLY = 0x2,
    FILESYSTEM_FLAG_RECURSIVE              = 0x4,
    FILESYSTEM_FLAG_USE_RECYCLE_BIN        = 0x8,
    FILESYSTEM_FLAG_CONTINUE_ON_ERROR      = 0x10,
    // Valid only with ALLOW_OVERWRITE and an exact expectedDestination whose no-follow
    // snapshot kind is LINK. This is the typed Replace Link receipt; ordinary Overwrite
    // must never remove a link object.
    FILESYSTEM_FLAG_ALLOW_REPLACE_LINK     = 0x20,
};

enum FileSystemLinkPolicy : uint32_t
{
    FILESYSTEM_LINK_PRESERVE = 1,
    FILESYSTEM_LINK_SKIP     = 2,
};

enum FileSystemMoveMode : uint32_t
{
    // The provider may use any Move implementation allowed by its advertised profile.
    FILESYSTEM_MOVE_DEFAULT = 0,

    // The provider MUST perform only its qualified native Move route. It must fail before any
    // generic/unqualified fallback when that route cannot complete the item. A provider-native
    // conditional server-side strategy may still include deletion as part of its qualified route.
    FILESYSTEM_MOVE_NATIVE_ONLY = 1,
};

enum FileSystemDiscoveryMode : uint32_t
{
    // Keep the single traversal ahead of execution within the host/provider bounded queue.
    FILESYSTEM_DISCOVERY_AHEAD = 1,

    // Discover only enough work to sustain the currently executable transfer slots.
    FILESYSTEM_DISCOVERY_JUST_IN_TIME = 2,
};

struct FileSystemDiscoveryProgress
{
    uint32_t sizeBytes; // sizeof(FileSystemDiscoveryProgress)
    uint64_t discoveredBytes;
    uint64_t discoveredFiles;
    uint64_t discoveredDirectories;
    uint32_t queuedItems;
    BOOL traversalClosed;
};

// Synchronous per-operation control. This is a raw callback interface, not COM. The host owns it and
// keeps it alive until the provider call and every object opened with the corresponding options are released.
interface __declspec(novtable) IFileSystemOperationControl
{
    virtual HRESULT STDMETHODCALLTYPE FileSystemShouldAbort(BOOL* abort, void* cookie) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE FileSystemGetDiscoveryMode(FileSystemDiscoveryMode* mode, void* cookie) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE FileSystemReportDiscoveryProgress(const FileSystemDiscoveryProgress* progress,
                                                                         void* cookie) noexcept = 0;
};

enum class FileSystemIssueAction : uint8_t
{
    None = 0,
    Overwrite,
    ReplaceReadOnly,
    ReplaceLink,
    PermanentDelete,
    Retry,
    KeepBoth,
    Skip,
    Cancel,
};

enum class FileSystemOwnedStageDisposition : uint32_t
{
    NotApplicable = 0,
    NotCreated,
    Removed,
    Published,
    Retained,
    Unknown,
    // The final leaf is known to be the provider-created object and is known to
    // survive incomplete cleanup with only a prefix of the requested content.
    RetainedIncomplete,
};

// Per-item mutation truth reported with FileSystemItemCompleted. Providers initialize every field
// they advertise through sizeBytes even on failure. A null result is permitted only when the
// provider cannot expose mutation truth; the host then treats destructive recovery/escalation as
// indeterminate and fails closed. The v1 prefix ends before ownedStageDisposition so older providers
// remain valid. Stage disposition is independent of final-destination mutation truth: for example,
// a failed pre-publication copy can report known non-commit plus an Unknown retained stage.
struct FileSystemItemMutationResult
{
    uint32_t sizeBytes; // sizeof(FileSystemItemMutationResult)
    BOOL outcomeKnown;
    BOOL mutationCommitted;
    BOOL originalStillPresent;
    FileSystemOwnedStageDisposition ownedStageDisposition = FileSystemOwnedStageDisposition::NotApplicable;
};

inline constexpr uint32_t FILESYSTEM_ITEM_MUTATION_RESULT_V1_SIZE =
    static_cast<uint32_t>(offsetof(FileSystemItemMutationResult, ownedStageDisposition));

[[nodiscard]] inline bool FileSystemItemMutationResultHasSupportedSize(
    const FileSystemItemMutationResult& result) noexcept
{
    return result.sizeBytes == FILESYSTEM_ITEM_MUTATION_RESULT_V1_SIZE ||
           result.sizeBytes >= sizeof(FileSystemItemMutationResult);
}

[[nodiscard]] inline bool FileSystemItemMutationResultHasOwnedStageDisposition(
    const FileSystemItemMutationResult& result) noexcept
{
    return result.sizeBytes >= sizeof(FileSystemItemMutationResult);
}

[[nodiscard]] inline bool FileSystemOwnedStageDispositionIsValid(FileSystemOwnedStageDisposition disposition) noexcept
{
    return disposition >= FileSystemOwnedStageDisposition::NotApplicable &&
           disposition <= FileSystemOwnedStageDisposition::RetainedIncomplete;
}

struct FileSystemOptions
{
    uint32_t sizeBytes; // sizeof(FileSystemOptions)

    // 0 = unlimited (use all available bandwidth).
    // Callbacks receive an in/out FileSystemOptions* so the host can tweak it on progress updates (e.g. changing the limit mid-flight).
    // Plugins MAY also write back an effective applied limit (e.g. internal clamping or combining with a plugin-specific cap).
    uint64_t bandwidthLimitBytesPerSecond;

    // 0 = plugin default. Non-zero clamps copy/move fan-out for the whole call, including recursive
    // work a plugin schedules internally under one top-level host item.
    uint32_t copyMoveMaxConcurrency;

    // Per-call link handling. The removed Follow value (0) is reserved and invalid at the ABI boundary.
    FileSystemLinkPolicy linkPolicy = FILESYSTEM_LINK_PRESERVE;

    // Move strategy constraint. Non-Move calls require FILESYSTEM_MOVE_DEFAULT.
    FileSystemMoveMode moveMode = FILESYSTEM_MOVE_DEFAULT;

    // Absolute GetTickCount64() deadline. 0 means no deadline.
    uint64_t deadlineTickCount64 = 0;

    // Optional cooperative abort source. Providers must not retain either pointer beyond the operation
    // call or the lifetime of an object opened by that operation.
    IFileSystemOperationControl* operationControl = nullptr;
    void* operationControlCookie                  = nullptr;
};

[[nodiscard]] inline bool FileSystemOptionsHaveValidHeader(const FileSystemOptions* options) noexcept
{
    return options == nullptr ||
           (options->sizeBytes == sizeof(FileSystemOptions) &&
            (options->linkPolicy == FILESYSTEM_LINK_PRESERVE || options->linkPolicy == FILESYSTEM_LINK_SKIP));
}

inline constexpr uint64_t FILESYSTEM_OWNED_STAGE_CLEANUP_TIMEOUT_MS = 2'000u;

// Builds the short-lived control used only to remove an exact object already owned by the
// operation. Primary cancellation must not suppress compensation; the fresh deadline remains
// authoritative and this control must never be used to resume primary work.
[[nodiscard]] inline FileSystemOptions MakeOwnedStageCleanupOptions(const FileSystemOptions* primaryOptions) noexcept
{
    FileSystemOptions cleanupOptions{};
    if (primaryOptions != nullptr && FileSystemOptionsHaveValidHeader(primaryOptions))
    {
        cleanupOptions = *primaryOptions;
    }
    cleanupOptions.sizeBytes              = sizeof(cleanupOptions);
    cleanupOptions.operationControl       = nullptr;
    cleanupOptions.operationControlCookie = nullptr;
    const uint64_t now                     = GetTickCount64();
    cleanupOptions.deadlineTickCount64 =
        now > UINT64_MAX - FILESYSTEM_OWNED_STAGE_CLEANUP_TIMEOUT_MS
            ? UINT64_MAX
            : now + FILESYSTEM_OWNED_STAGE_CLEANUP_TIMEOUT_MS;
    return cleanupOptions;
}

// Canonical cooperative operation checkpoint shared by the host and every provider. A provider
// may call this before or between bounded units of work; it never owns or extends the lifetime of
// operationControl. ERROR_TIMEOUT is distinct from user cancellation so the host can report a
// provider deadline/quiet-point failure as indeterminate when mutation truth is not known.
[[nodiscard]] inline HRESULT FileSystemCheckOperationControl(const FileSystemOptions* options) noexcept
{
    if (! FileSystemOptionsHaveValidHeader(options))
    {
        return E_INVALIDARG;
    }
    if (options == nullptr)
    {
        return S_OK;
    }
    if (options->deadlineTickCount64 != 0u && GetTickCount64() >= options->deadlineTickCount64)
    {
        return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
    }
    if (options->operationControl == nullptr)
    {
        return S_OK;
    }

    BOOL abort = FALSE;
    const HRESULT abortHr = options->operationControl->FileSystemShouldAbort(&abort, options->operationControlCookie);
    if (FAILED(abortHr))
    {
        return abortHr;
    }
    return abort != FALSE ? HRESULT_FROM_WIN32(ERROR_CANCELLED) : S_OK;
}

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

enum FileSystemRouteAvailability : uint32_t
{
    FILESYSTEM_ROUTE_UNSUPPORTED = 0,
    FILESYSTEM_ROUTE_AVAILABLE   = 1,
};

enum FileSystemCancellationRoute : uint32_t
{
    FILESYSTEM_CANCELLATION_UNCONTAINED       = 0,
    FILESYSTEM_CANCELLATION_BOUNDED           = 1,
    FILESYSTEM_CANCELLATION_PROVIDER_WATCHDOG = 2,
};

enum FileSystemNamespaceKind : uint32_t
{
    FILESYSTEM_NAMESPACE_REAL_CONTAINER          = 1,
    FILESYSTEM_NAMESPACE_PROVIDER_VIRTUAL_FOLDER = 2,
    FILESYSTEM_NAMESPACE_FIXED_OBJECT_SET         = 3,
};

enum FileSystemRouteComponentComparison : uint32_t
{
    FILESYSTEM_ROUTE_COMPONENT_ORDINAL_IGNORE_CASE = 1,
    FILESYSTEM_ROUTE_COMPONENT_ORDINAL_CASE_SENSITIVE = 2,
};

enum FileSystemRouteNormalization : uint32_t
{
    FILESYSTEM_ROUTE_NORMALIZATION_NONE = 1,
};

enum FileSystemRouteCaseOnlyRename : uint32_t
{
    FILESYSTEM_ROUTE_CASE_ONLY_SUPPORTED      = 1,
    FILESYSTEM_ROUTE_CASE_ONLY_NO_OP          = 2,
    FILESYSTEM_ROUTE_CASE_ONLY_UNSUPPORTED    = 3,
    FILESYSTEM_ROUTE_CASE_ONLY_NOT_APPLICABLE = 4,
};

enum FileSystemRouteProofFlags : uint32_t
{
    FILESYSTEM_ROUTE_PROOF_NONE            = 0,
    FILESYSTEM_ROUTE_PROOF_HOST_READBACK   = 0x1,
    FILESYSTEM_ROUTE_PROOF_PROVIDER_BLAKE3 = 0x2,
    FILESYSTEM_ROUTE_PROOF_PROVIDER_REREAD = 0x4,
    FILESYSTEM_ROUTE_PROOF_WRITER_DIGEST   = 0x8, // R3-2: the route's writer proves published content (IFileWriterContentProof)
};

enum FileSystemTransferPeerRole : uint32_t
{
    FILESYSTEM_TRANSFER_PEER_EXPORT = 1,
    FILESYSTEM_TRANSFER_PEER_IMPORT = 2,
};

enum FileSystemChildNameStatus : uint32_t
{
    FILESYSTEM_CHILD_NAME_UNSUPPORTED = 0,
    FILESYSTEM_CHILD_NAME_VALID       = 1,
    FILESYSTEM_CHILD_NAME_INVALID     = 2,
};

// Typed, path-and-operation-scoped route facts. The caller supplies a FileSystemArena
// whose buffer is aligned for wchar_t. Strings point into that arena and remain valid
// only until the arena is reset or destroyed.
// Facts describe feasibility; exact object/container interfaces and mutation receipts
// remain the authority for destructive work.
struct FileSystemRouteFacts
{
    uint32_t sizeBytes; // sizeof(FileSystemRouteFacts)
    uint32_t requiredArenaBytes;

    FileSystemRouteAvailability availability;
    FileSystemCancellationRoute cancellationRoute;
    FileSystemNamespaceKind namespaceKind;
    FileSystemRouteComponentComparison componentComparison;
    FileSystemRouteNormalization normalization;
    FileSystemRouteCaseOnlyRename caseOnlyRename;
    uint32_t proofFlags;
    uint32_t providerWatchdogTimeoutMs;
    uint32_t copyMoveMaxConcurrency;
    uint32_t deleteMaxConcurrency;
    uint32_t deleteRecycleBinMaxConcurrency;
    uint64_t maxComponentUtf16;

    BOOL copyOperation;
    BOOL moveOperation;
    BOOL nativeMoveOperation;
    BOOL deleteOperation;
    BOOL renameOperation;
    BOOL createDirectoryOperation;
    BOOL propertiesOperation;
    BOOL readOperation;
    BOOL writeOperation;
    BOOL recycleOperation;
    BOOL boundDelete;
    BOOL conditionalDelete;
    BOOL exclusiveStage;
    BOOL conditionalPublish;
    BOOL committedSize;
    BOOL preserveFileLink;
    BOOL preserveDirectoryLink;
    BOOL retargetInTree;
    BOOL exactLinkRemoval;
    BOOL cancellationAbort;
    BOOL cancellationDeadline;
    BOOL pathTextStableIdentity;
    BOOL casePreserving;

    wchar_t preferredSeparator;
    const wchar_t* acceptedSeparators;
    const wchar_t* providerId;
    const wchar_t* pathProfileId;
    const wchar_t* rootId;
};

struct FileSystemChildNameValidation
{
    uint32_t sizeBytes; // sizeof(FileSystemChildNameValidation)
    FileSystemChildNameStatus status;
    HRESULT failureStatus;
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
    FILESYSTEM_SEARCH_WARNING_SERVICE_ROOT_REJECTED = 0x40,
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
    uint32_t reserved;                       // Must be 0 for normal callers; built-in host extensions use FILESYSTEM_SEARCH_HOST_EXTENSIONS_MARKER.
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
inline constexpr uint32_t FILESYSTEM_SEARCH_HOST_EXTENSIONS_MARKER = 0x52534631u; // "RSF1"

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
    uint32_t sizeBytes; // Caller-provided byte size; unknown tail fields are ignored.
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
interface IFileSystemBoundObject;

interface __declspec(novtable) IFileSystemCallback
{
    // options may be nullptr; implementations must check before reading/writing to it.
    // If options is non-null, it is an in/out object:
    // - the host may update fields (e.g. speed limit changes)
    // - plugins may write back an effective applied value (e.g. clamping / combining with internal caps)
    // Plugins SHOULD read options after the callback returns.
    // Notes:
    // - totalItems/totalBytes MAY be 0 while the single traversal is still open; hosts make totals authoritative only after discovery closes.
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
                                                              const FileSystemItemMutationResult* mutationResult,
                                                              FileSystemOptions* options,
                                                              void* cookie) noexcept                = 0;
    virtual HRESULT STDMETHODCALLTYPE FileSystemShouldCancel(BOOL * pCancel, void* cookie) noexcept = 0;

    // Invoked by plugins when an operation hits a conflict/issue that requires a user decision (retry/skip/etc.).
    // Notes:
    // - sourcePath/destinationPath are best-effort; they may be nullptr for some operations (e.g. delete destination).
    // - action and expectedDestination must be non-null. Implementations initialize both even when
    //   returning failure/cancellation.
    // - For Overwrite, ReplaceReadOnly, or ReplaceLink, the host returns the exact no-follow
    //   destination authority captured before the decision surface was exposed. The provider MUST
    //   consume that authority through PublishAs/RenameIfUnchanged; it MUST NOT reopen or delete the
    //   pathname as replacement authority. Non-destructive actions return a null authority.
    // - This callback may block (host-driven inline conflict UI).
    virtual HRESULT STDMETHODCALLTYPE FileSystemIssue(FileSystemOperation operationType,
                                                      const wchar_t* sourcePath,
                                                       const wchar_t* destinationPath,
                                                       HRESULT status,
                                                       FileSystemIssueAction* action,
                                                       IFileSystemBoundObject** expectedDestination,
                                                       FileSystemOptions* options,
                                                       void* cookie) noexcept = 0;
};

// Mandatory path-scoped capability interface for ABI v2. The returned UTF-8 JSON is owned by the
// provider and remains valid until the next call on this instance or instance release. The v2
// operations object contains both "move" (the provider Move entry point is admissible for this
// path/profile) and "nativeMove" (that entry point is a qualified native strategy). The host must
// never infer nativeMove from move.
interface __declspec(uuid("9be5ee85-f247-4a95-953b-5f18ed76719d")) __declspec(novtable) IFileSystemPathCapabilities2 : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE GetPathCapabilities(const wchar_t* path,
                                                           FileSystemOperation operation,
                                                           const char** jsonUtf8) noexcept = 0;
};

// The sole executable route-capability authority. Capability JSON remains available
// through IFileSystemPathCapabilities2 for diagnostics/extensions only.
interface __declspec(uuid("1e924d87-2e62-4ab4-9f37-c565d465f25e")) __declspec(novtable) IFileSystemRouteCapabilities : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE GetRouteFacts(const wchar_t* path,
                                                     FileSystemOperation operation,
                                                     FileSystemArena* arena,
                                                     FileSystemRouteFacts* facts) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE IsTransferPeerAllowed(const wchar_t* path,
                                                            FileSystemOperation operation,
                                                            FileSystemTransferPeerRole role,
                                                            const wchar_t* peerPluginId,
                                                            BOOL* allowed) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE ValidateChildName(const wchar_t* parentPath,
                                                        const wchar_t* childName,
                                                        FileSystemOperation operation,
                                                        FileSystemChildNameValidation* validation) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE GetChildNameCollisionKey(const wchar_t* parentPath,
                                                               const wchar_t* childName,
                                                               FileSystemOperation operation,
                                                               FileSystemArena* arena,
                                                               const wchar_t** key,
                                                               unsigned long* requiredArenaBytes) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE JoinPath(const wchar_t* parentPath,
                                               const wchar_t* childName,
                                               FileSystemOperation operation,
                                               FileSystemArena* arena,
                                               const wchar_t** joinedPath,
                                               unsigned long* requiredArenaBytes) noexcept = 0;
};

interface __declspec(uuid("12519afa-30e7-4e3a-9db2-7990c4be9a21")) __declspec(novtable) IFileSystem : public IFileSystemPathCapabilities2
{
    // Lists the contents of a directory into an IFilesInformation object.
    // On success, ppFilesInformation receives a valid instance of IFilesInformation.
    virtual HRESULT STDMETHODCALLTYPE ReadDirectoryInfo(const wchar_t* path, IFilesInformation** ppFilesInformation) noexcept = 0;

    // Directory-merge contract for Copy/Move (normative; see Specs/FileSystem/FileSystem_FileOperations.md,
    // "Conflict Handling / Defaults"):
    //  - A source directory whose destination already exists as a directory MUST be merged by
    //    recursing into the existing destination directory. Directory-vs-directory existence is
    //    NOT an overwrite conflict and MUST NOT fail the item with ERROR_ALREADY_EXISTS.
    //  - ERROR_ALREADY_EXISTS is reserved for file-vs-existing-file and file/directory type
    //    mismatches. Recursive implementations SHOULD raise such collisions per child through
    //    IFileSystemCallback::FileSystemIssue (most-specific failing path) and continue per the
    //    returned action, finishing with HRESULT_FROM_WIN32(ERROR_PARTIAL_COPY) when children
    //    were skipped — never abort a whole directory transfer on the first collision.
    //  - Move MUST apply the same merge rule, falling back to per-child move or copy+delete when
    //    the platform cannot rename a directory onto an existing directory. Skipped children keep
    //    their source ("source preserved") and the item ends as ERROR_PARTIAL_COPY.
    //  - Overwrite / ReplaceReadOnly granted from a child conflict are ONE-SHOT: they apply only to
    //    the child that was answered and MUST NOT leak to sibling or descendant children. Only the
    //    explicit "Apply to all similar" toggle broadens a grant across many children.
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

// Optional cancellation/progress contract for potentially expensive directory enumeration.
// The callback is per-call and synchronous: implementations MUST NOT retain it or invoke it
// after ReadDirectoryInfoCancellable returns. A failed progress callback or TRUE cancellation
// request must stop work promptly and return HRESULT_FROM_WIN32(ERROR_CANCELLED).
interface __declspec(novtable) IFileSystemDirectoryEnumerationCallback
{
    virtual HRESULT STDMETHODCALLTYPE DirectoryEnumerationProgress(uint64_t scannedEntries, uint64_t totalEntries, void* cookie) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE DirectoryEnumerationShouldCancel(BOOL * pCancel, void* cookie) noexcept                              = 0;
};

interface __declspec(uuid("32f1b16a-fb45-4e59-98ca-61e16445c527")) __declspec(novtable) IFileSystemCancellableDirectoryEnumeration : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE ReadDirectoryInfoCancellable(const wchar_t* path,
                                                                   IFileSystemDirectoryEnumerationCallback* callback,
                                                                   void* cookie,
                                                                   IFilesInformation** ppFilesInformation) noexcept = 0;
};

// Minimal Win32-like file reader for filesystem plugins.
// Notes:
// - The reader is read-only.
// - Implementations MUST be safe for large files (64-bit offsets/sizes).
// - Read returns S_OK with *bytesRead == 0 only at end-of-file or when bytesToRead == 0.
// - On success, *bytesRead MUST NOT exceed bytesToRead. Hosts reject a larger value as invalid provider data.
interface __declspec(uuid("b1d0c2b8-0e37-4d6f-8c2c-2cc4f0d1c6b8")) __declspec(novtable) IFileReader : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE GetSize(uint64_t* sizeBytes) noexcept                                            = 0;
    virtual HRESULT STDMETHODCALLTYPE Seek(__int64 offset, unsigned long origin, uint64_t* newPosition) noexcept       = 0;
    virtual HRESULT STDMETHODCALLTYPE Read(void* buffer, unsigned long bytesToRead, unsigned long* bytesRead) noexcept = 0;
};

// Optional reader contract (R0f-Curl-OR1) for readers created through IFileSystemIO::CreateFileReader,
// which carries no FileSystemOptions: the host hands the task's operation control right after
// creation and before the first Read. The reader keeps the pointer for its lifetime (the caller
// guarantees the options outlive the reader) and polls the control while a Read is in flight, so a
// cancelled or deadline-expired call ends the transfer and the Read with ERROR_CANCELLED within the
// control's poll period instead of only between reads. Readers opened from a bound object receive
// their options through IFileSystemBoundObject::OpenReader and do not need this interface.
interface __declspec(uuid("b0d9beb3-76eb-4427-bf21-16d0ae1938c8")) __declspec(novtable) IFileReaderOperationControl : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE SetOperationControl(const FileSystemOptions* options) noexcept = 0;
};

// Minimal Win32-like file writer for filesystem plugins.
// Notes:
// - Implementations MUST be safe for large files (64-bit offsets/sizes).
// - Implementations MUST tolerate being released without Commit() (treat as abort / best-effort cleanup).
// - Overwrite-capable implementations SHOULD preserve any pre-existing destination until Commit() succeeds.
// - On success, *bytesWritten MUST NOT exceed bytesToWrite. Hosts reject a larger value as invalid provider data.
// - GetPosition MUST report the writer-local persisted position without publishing the destination.
interface __declspec(uuid("b6f0a9e1-8c8b-4b72-9f3e-2f2b4b8b9c41")) __declspec(novtable) IFileWriter : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE GetPosition(uint64_t* positionBytes) noexcept                                               = 0;
    virtual HRESULT STDMETHODCALLTYPE Write(const void* buffer, unsigned long bytesToWrite, unsigned long* bytesWritten) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE Commit() noexcept                                                                           = 0;
};

// Optional writer contract used by streaming destinations that must know the final object size
// before accepting the first byte (for example, chunked cloud upload sessions).
// The host calls SetExpectedSize immediately after CreateFileWriter and before Write.
// Implementations that expose this interface MUST reject a different byte count at Commit.
interface __declspec(uuid("24719c4b-0b51-4d63-93aa-4c697eb8a732")) __declspec(novtable) IFileWriterExpectedSize : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE SetExpectedSize(uint64_t sizeBytes) noexcept = 0;
};

// Optional writer contract for destinations that can prove the exact object size accepted by a
// successful Commit without performing a destination metadata query. The host calls
// GetCommittedSize only after Commit returns success. Implementations MUST return
// HRESULT_FROM_WIN32(ERROR_INVALID_STATE) before a successful Commit, MUST return the exact byte
// count published by that Commit on success, and MUST answer from writer-local state without
// network or filesystem I/O. A host may carry this proof through a successful staged MoveItem
// promotion because moving the committed object must preserve its bytes.
interface __declspec(uuid("f64f3fed-7262-435c-b509-c9751b6e0280")) __declspec(novtable) IFileWriterCommitSizeProof : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE GetCommittedSize(uint64_t* sizeBytes) noexcept = 0;
};

// Optional filesystem capability for destinations whose writer Commit atomically publishes the
// requested path. The host uses this explicit contract to avoid a temp-name rename/copy; it never
// infers the capability from provider IDs or temporary-path spelling.
interface __declspec(uuid("bcf04a7a-9c62-4aa8-9847-d756cf432669")) __declspec(novtable) IFileSystemAtomicWriter : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE SupportsAtomicWriterCommit(const wchar_t* path, FileSystemFlags flags, BOOL* supported) noexcept = 0;
};

// Optional (C10): a provider without object binding that can still name the object at a path with a
// stable identity (an object key with its version or ETag, a provider file id). While a Permanent
// Delete prepares, the host resolves the identity of every selected root through
// ResolveDeleteIdentity (read-only, no-follow); after the card answer it resolves again and refuses
// a root whose identity changed; at execution DeleteIfIdentity deletes the path only while the live
// identity still equals the pinned one and otherwise behaves as DeleteItem, including its per-item
// receipt through the callback. A live identity that differs returns
// HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH) with nothing deleted. An empty identity text means the
// provider cannot name that object and the host refuses the delete. For a directory the identity
// names the directory object; the recursive delete keeps the provider's own per-observation rules
// for descendants. A route with neither object binding nor this contract deletes whatever the name
// refers to when the mutation runs, and the host's card says so before the answer.
struct FileSystemDeleteIdentity
{
    uint32_t sizeBytes;    // sizeof(FileSystemDeleteIdentity)
    BOOL isDirectory;
    wchar_t identity[256]; // provider-defined, NUL-terminated; empty = cannot be named
};

interface IFileSystemCallback;

interface __declspec(uuid("7c2d9a4e-5b31-4f8e-9a6d-0e3f1b7c8d52")) __declspec(novtable) IFileSystemIdentityDelete : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE ResolveDeleteIdentity(const wchar_t* path,
                                                            const FileSystemOptions* options,
                                                            FileSystemDeleteIdentity* identity) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE DeleteIfIdentity(const wchar_t* path,
                                                       const FileSystemDeleteIdentity* identity,
                                                       FileSystemFlags flags,
                                                       const FileSystemOptions* options,
                                                       IFileSystemCallback* callback,
                                                       void* cookie) noexcept = 0;
};

struct FileSystemBasicInformation
{
    uint32_t sizeBytes = 0; // sizeof(FileSystemBasicInformation)

    __int64 creationTime     = 0; // FILETIME ticks (100ns intervals since 1601-01-01 UTC)
    __int64 lastAccessTime   = 0; // FILETIME ticks
    __int64 lastWriteTime    = 0; // FILETIME ticks
    unsigned long attributes = 0; // FILE_ATTRIBUTE_* flags
};

// Optional writer contract for atomic-final writers on providers that cannot bind destination
// objects (no IFileSystemObjectBinding). When the host grants Overwrite or Replace read-only on such
// a route, it calls SetExpectedReplacement once, after CreateFileWriter and before the first Write,
// with the occupant it showed the user (basic information read no-follow at the decision; zero
// fields are unknown). The writer MUST then replace only that occupant: at Commit it revalidates
// the current occupant with the strongest identity it has (provider item id, ETag, or version,
// otherwise the size and last-write time captured at creation) and fails with
// HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH), publishing nothing, when the name is absent, owned
// by a different object, or changed. It MAY return that status from SetExpectedReplacement already
// when the mismatch is provable early. A provider whose SupportsAtomicWriterCommit accepts
// FILESYSTEM_FLAG_ALLOW_OVERWRITE MUST expose this interface on every writer created with that
// flag; the host withholds the destructive decision otherwise.
interface __declspec(uuid("72e8650c-f69d-4595-8bee-70851de51599")) __declspec(novtable) IFileWriterExpectedReplacement : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE SetExpectedReplacement(const FileSystemBasicInformation* expected) noexcept = 0;
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

enum FileSystemBindFlags : uint32_t
{
    FILESYSTEM_BIND_NO_FOLLOW     = 0x01,
    FILESYSTEM_BIND_READ_CONTENT  = 0x02,
    FILESYSTEM_BIND_READ_METADATA = 0x04,
    FILESYSTEM_BIND_DELETE        = 0x08,
    FILESYSTEM_BIND_RENAME        = 0x10,
    FILESYSTEM_BIND_PUBLICATION   = 0x20,
};

enum FileSystemLinkKind : uint32_t
{
    FILESYSTEM_LINK_KIND_SYMBOLIC_FILE      = 1,
    FILESYSTEM_LINK_KIND_SYMBOLIC_DIRECTORY = 2,
    FILESYSTEM_LINK_KIND_JUNCTION           = 3,
};

enum FileSystemLinkTargetMapping : uint32_t
{
    FILESYSTEM_LINK_TARGET_OUTSIDE_SOURCE_ROOT           = 1,
    FILESYSTEM_LINK_TARGET_MAPPED_INSIDE_SOURCE_ROOT     = 2,
    FILESYSTEM_LINK_TARGET_INSIDE_SOURCE_ROOT_UNMAPPABLE = 3,
};

// Bounded literal link payload. Lengths exclude the terminating NUL. For ReadBoundLink,
// capacities are caller-owned buffer capacities including the NUL; ERROR_INSUFFICIENT_BUFFER
// reports the required length without partial text. Preserve is literal: the payload carries the
// stored target text and relative flag unchanged, targetMapping is
// FILESYSTEM_LINK_TARGET_OUTSIDE_SOURCE_ROOT, and sourceRelativeTarget is empty. The other mapping
// values and the transform's component mappings are reserved for an optional retarget transform
// that no built-in provider implements. For CreateExclusiveLink, only targetBuffer is consumed and
// it must contain targetLengthUtf16 characters followed by NUL.
struct FileSystemLinkInformation
{
    uint32_t sizeBytes; // sizeof(FileSystemLinkInformation)
    FileSystemLinkKind kind;
    BOOL targetIsRelative;
    FileSystemLinkTargetMapping targetMapping;
    uint32_t targetLengthUtf16;
    uint32_t targetCapacityUtf16;
    wchar_t* targetBuffer;
    uint32_t sourceRelativeTargetLengthUtf16;
    uint32_t sourceRelativeTargetCapacityUtf16;
    wchar_t* sourceRelativeTargetBuffer;
};

struct FileSystemLinkComponentMapping
{
    const wchar_t* sourceRelativeComponentPath;
    const wchar_t* destinationComponentName;
};

// Link read context. Paths are provider-internal spellings. Literal Preserve validates the roots
// and component mappings and never uses them to rewrite a target.
struct FileSystemLinkTransform
{
    uint32_t sizeBytes; // sizeof(FileSystemLinkTransform)
    const wchar_t* sourceLinkPath;
    const wchar_t* destinationLinkPath;
    const wchar_t* sourceRootPath;
    const wchar_t* destinationRootPath;
    const FileSystemLinkComponentMapping* componentMappings;
    uint32_t componentMappingCount;
};

enum FileSystemBoundObjectKind : uint32_t
{
    FILESYSTEM_BOUND_REGULAR_FILE = 1,
    FILESYSTEM_BOUND_DIRECTORY    = 2,
    FILESYSTEM_BOUND_LINK         = 3,
    FILESYSTEM_BOUND_OTHER        = 4,
};

struct FileSystemBoundObjectSnapshot
{
    uint32_t sizeBytes;
    uint32_t kind;               // FileSystemBoundObjectKind
    const void* objectId;        // Provider-owned; immutable for the bound-object lifetime.
    uint32_t objectIdBytes;
    const void* revisionId;      // Empty only when the path profile declares no revision token.
    uint32_t revisionIdBytes;
    uint64_t committedSizeBytes; // UINT64_MAX when not applicable or unknown.
};

enum FileSystemContentProofAlgorithm : uint32_t
{
    FILESYSTEM_CONTENT_PROOF_NONE       = 0,
    FILESYSTEM_CONTENT_PROOF_BLAKE3_256 = 1,
    // R3-2 writer-side proofs: the digest a provider holds for the object it published.
    FILESYSTEM_CONTENT_PROOF_SHA256_256   = 2,
    FILESYSTEM_CONTENT_PROOF_SHA1_160     = 3,
    FILESYSTEM_CONTENT_PROOF_MD5_128      = 4,
    FILESYSTEM_CONTENT_PROOF_QUICKXOR_160 = 5,
    FILESYSTEM_CONTENT_PROOF_CRC64NVME_64 = 6,
};

[[nodiscard]] inline constexpr uint32_t FileSystemContentProofDigestBytes(FileSystemContentProofAlgorithm algorithm) noexcept
{
    switch (algorithm)
    {
        case FILESYSTEM_CONTENT_PROOF_BLAKE3_256: return 32u;
        case FILESYSTEM_CONTENT_PROOF_SHA256_256: return 32u;
        case FILESYSTEM_CONTENT_PROOF_SHA1_160: return 20u;
        case FILESYSTEM_CONTENT_PROOF_MD5_128: return 16u;
        case FILESYSTEM_CONTENT_PROOF_QUICKXOR_160: return 20u;
        case FILESYSTEM_CONTENT_PROOF_CRC64NVME_64: return 8u;
        case FILESYSTEM_CONTENT_PROOF_NONE:
        default: return 0u;
    }
}

inline constexpr uint32_t FILESYSTEM_BLAKE3_DIGEST_BYTES = 32u;

struct FileSystemContentProof
{
    uint32_t sizeBytes; // sizeof(FileSystemContentProof)
    FileSystemContentProofAlgorithm algorithm;
    uint64_t contentSizeBytes;
    uint8_t digest[FILESYSTEM_BLAKE3_DIGEST_BYTES];
    uint32_t reserved[4];
};

struct FileSystemConditionalMutationResult
{
    uint32_t sizeBytes;
    BOOL mutationCommitted;
    BOOL originalStillPresent;
    BOOL outcomeKnown;
};

// Optional ABI-v2 binding surface. BindObject is always no-follow; providers that cannot guarantee
// that return ERROR_NOT_SUPPORTED. A successful CreateExclusiveWriter returns both non-null outputs.
interface __declspec(uuid("d8ae290a-b84c-42ec-982c-7c01dedc7603")) __declspec(novtable) IFileSystemObjectBinding : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE BindObject(const wchar_t* path,
                                                 FileSystemBindFlags flags,
                                                 IFileSystemBoundObject** bound) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE CreateExclusiveWriter(const wchar_t* stagePath,
                                                             const FileSystemOptions* options,
                                                             IFileWriter** writer,
                                                             IFileSystemBoundObject** ownedStage) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE CreateExclusiveDirectory(const wchar_t* stagePath,
                                                                const FileSystemOptions* options,
                                                                IFileSystemBoundObject** ownedStage) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE ReadBoundLink(IFileSystemBoundObject* boundLink,
                                                    const FileSystemLinkTransform* transform,
                                                    const FileSystemOptions* options,
                                                    FileSystemLinkInformation* information) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE CreateExclusiveLink(const wchar_t* stagePath,
                                                          const FileSystemLinkInformation* information,
                                                          const FileSystemOptions* options,
                                                          IFileSystemBoundObject** ownedStage) noexcept = 0;
};

// Optional immutable object/revision token used for publication and exact conditional mutation.
// Every conditional method must initialize outcomeKnown even when transport fails. A null
// expectedDestination means "only if absent"; replacement requires the exact bound destination.
interface __declspec(uuid("5ed3921d-a486-42cc-a5c3-33d97f75c5b5")) __declspec(novtable) IFileSystemBoundObject : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE GetSnapshot(FileSystemBoundObjectSnapshot* snapshot) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE IsSameObject(IFileSystemBoundObject* other, BOOL* same) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE OpenReader(const FileSystemOptions* options, IFileReader** reader) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE GetBasicInformation(FileSystemBasicInformation* info) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE SetBasicInformation(const FileSystemBasicInformation* info) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE PublishAs(const wchar_t* finalPath,
                                                IFileSystemBoundObject* expectedDestination,
                                                FileSystemFlags flags,
                                                const FileSystemOptions* options,
                                                FileSystemConditionalMutationResult* result,
                                                IFileSystemBoundObject** published) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE RenameIfUnchanged(const wchar_t* destinationPath,
                                                        IFileSystemBoundObject* expectedDestination,
                                                        FileSystemFlags flags,
                                                        const FileSystemOptions* options,
                                                        FileSystemConditionalMutationResult* result,
                                                        IFileSystemBoundObject** renamed) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE DeleteIfUnchanged(FileSystemFlags flags,
                                                        const FileSystemOptions* options,
                                                        FileSystemConditionalMutationResult* result) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE AbortOwnedObject(const FileSystemOptions* options,
                                                       FileSystemConditionalMutationResult* result) noexcept = 0;
};

// Optional exact-object content proof. The provider must read the already-bound object/revision;
// it must never reopen a pathname. A proof is evidence only for optional post-publication
// verification and can never authorize overwrite, skip, cleanup, or source deletion.
interface __declspec(uuid("ed24d2c9-3377-468a-a8b2-72862353f7d1")) __declspec(novtable) IFileSystemBoundContentProof : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE GetContentProof(const FileSystemOptions* options,
                                                      FileSystemContentProof* proof) noexcept = 0;
};

// R3-2: writer-side content proof for atomic-final routes without bound objects. Before the first
// Write the host asks which digests the writer can report for the object it will publish (a mask
// of `1u << FileSystemContentProofAlgorithm`) and hashes the bytes it streams with those
// algorithms; after a successful Commit it asks for the proof the provider holds for the published
// object (the service's own digest, never one the writer recomputed from the bytes it sent). A
// matching proof verifies the published content and is the route's content proof before Managed
// Move source cleanup; a missing or different proof keeps the source. The proof can never
// authorize overwrite, skip, or deletion by itself. Routes that implement this advertise
// FILESYSTEM_ROUTE_PROOF_WRITER_DIGEST.
interface __declspec(uuid("122eb65f-ef47-4d11-8c8d-d2fb144df1e3")) __declspec(novtable) IFileWriterContentProof : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE GetContentProofAlgorithms(uint32_t* algorithmMask) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE GetCommittedContentProof(FileSystemContentProof* proof) noexcept = 0;
};

// Exact-object metadata contract used by the host bridge. The interface is optional and is
// queried from an IFileSystemBoundObject; no method may reopen the object's pathname. Feature
// masks are deliberately orthogonal so a provider can preserve one class while reporting exact
// loss or change for another.
enum FileSystemMetadataFeature : uint32_t
{
    FILESYSTEM_METADATA_BASIC_TIMES_ATTRIBUTES = 0x0001,
    FILESYSTEM_METADATA_MOTW                   = 0x0002,
    FILESYSTEM_METADATA_ALTERNATE_STREAMS      = 0x0004,
    FILESYSTEM_METADATA_EXTENDED_ATTRIBUTES    = 0x0008,
    FILESYSTEM_METADATA_SECURITY               = 0x0010,
    FILESYSTEM_METADATA_SPARSE                 = 0x0020,
    FILESYSTEM_METADATA_COMPRESSION            = 0x0040,
    FILESYSTEM_METADATA_EFS                    = 0x0080,
    FILESYSTEM_METADATA_PLACEHOLDER            = 0x0100,
};

struct FileSystemMetadataSnapshot
{
    uint32_t sizeBytes; // sizeof(FileSystemMetadataSnapshot)
    uint32_t supportedFeatures;
    uint32_t presentFeatures;
    uint64_t logicalSizeBytes;
    uint64_t allocatedSizeBytes;
    unsigned long fileAttributes;
    unsigned long reparseTag;
};

enum FileSystemMetadataTransferPhase : uint32_t
{
    // Apply allocation/encryption properties before the first content byte is written.
    FILESYSTEM_METADATA_TRANSFER_PREPARE_CONTENT = 1,
    // Apply streams, extended attributes, timestamps/attributes and report security inheritance
    // after content is complete but before identity-owned publication.
    FILESYSTEM_METADATA_TRANSFER_FINALIZE = 2,
};

struct FileSystemMetadataTransferResult
{
    uint32_t sizeBytes; // sizeof(FileSystemMetadataTransferResult)
    uint32_t attemptedFeatures;
    uint32_t preservedFeatures;
    uint32_t changedFeatures;
    uint32_t lostFeatures;
    HRESULT firstFailure;
};

interface __declspec(uuid("145b7908-1876-4c76-9900-88961d43f2ab")) __declspec(novtable) IFileSystemBoundMetadata : public IUnknown
{
    virtual HRESULT STDMETHODCALLTYPE GetMetadataSnapshot(const FileSystemOptions* options,
                                                          FileSystemMetadataSnapshot* snapshot) noexcept = 0;
    virtual HRESULT STDMETHODCALLTYPE TransferMetadataTo(IFileSystemBoundObject* destination,
                                                         FileSystemMetadataTransferPhase phase,
                                                         const FileSystemOptions* options,
                                                         FileSystemMetadataTransferResult* result) noexcept = 0;
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
    uint64_t fileCount;      // Observed files, including unavailable sizes in a partial result.
    uint64_t directoryCount; // Number of directories counted (excluding root).
    HRESULT status;          // S_OK, ERROR_PARTIAL_COPY, cancellation, or the first global error (as HRESULT).
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
// - CreateDirectory returns HRESULT_FROM_WIN32(ERROR_ALREADY_EXISTS) when the target path is
//   already occupied. Callers that merge (cross-filesystem bridge) MUST distinguish the two
//   cases behind that code themselves: an existing DIRECTORY is a valid merge target (treat as
//   success), while an existing FILE is a genuine type-mismatch conflict — probe the path's
//   attributes before deciding.
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

namespace Common::Plugins
{
// Checked traversal for the variable-size FileInfo ABI record chain. This is
// the canonical reader for untrusted plugin buffers: no record field is read
// until its fixed header is in bounds, and every advance is aligned and
// contained within the advertised buffer.
class PackedFileInfoCursor final
{
public:
    PackedFileInfoCursor(const FileInfo* buffer, unsigned long bufferSize, unsigned long count) noexcept
        : _buffer(reinterpret_cast<const unsigned char*>(buffer)), _bufferSize(bufferSize), _count(count)
    {
    }

    [[nodiscard]] HRESULT Next(const FileInfo** result) noexcept
    {
        if (! result)
        {
            return E_POINTER;
        }
        *result = nullptr;
        if (_index >= _count)
        {
            return HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES);
        }
        if (! _buffer || _bufferSize == 0u ||
            (reinterpret_cast<uintptr_t>(_buffer) % alignof(FileInfo)) != 0u ||
            (_offset % alignof(FileInfo)) != 0u || _offset > _bufferSize)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        constexpr size_t headerBytes = offsetof(FileInfo, FileName);
        const size_t remaining = static_cast<size_t>(_bufferSize) - _offset;
        if (remaining < headerBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        const auto* entry = reinterpret_cast<const FileInfo*>(_buffer + _offset);
        if ((entry->FileNameSize % sizeof(wchar_t)) != 0u ||
            static_cast<size_t>(entry->FileNameSize) > remaining - headerBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        const size_t requiredBytes = headerBytes + static_cast<size_t>(entry->FileNameSize);
        const bool finalRecord = _index + 1u == _count;
        const size_t advance = static_cast<size_t>(entry->NextEntryOffset);
        if (finalRecord)
        {
            if (advance != 0u)
            {
                return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
            }
        }
        else if (advance == 0u || (advance % alignof(FileInfo)) != 0u ||
                 advance < requiredBytes || advance > remaining || remaining - advance < headerBytes)
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }

        *result = entry;
        ++_index;
        if (! finalRecord)
        {
            _offset += advance;
        }
        return S_OK;
    }

private:
    const unsigned char* _buffer = nullptr;
    size_t _bufferSize = 0u;
    size_t _offset = 0u;
    unsigned long _count = 0u;
    unsigned long _index = 0u;
};

[[nodiscard]] inline HRESULT LocatePackedFileInfoRecord(const FileInfo* buffer,
                                                        unsigned long bufferSize,
                                                        unsigned long count,
                                                        unsigned long index,
                                                        const FileInfo** result) noexcept
{
    if (! result)
    {
        return E_POINTER;
    }
    *result = nullptr;
    if (index >= count)
    {
        return HRESULT_FROM_WIN32(ERROR_NO_MORE_FILES);
    }
    PackedFileInfoCursor cursor(buffer, bufferSize, count);
    const FileInfo* located = nullptr;
    for (unsigned long current = 0u; current < count; ++current)
    {
        const FileInfo* entry = nullptr;
        const HRESULT hr = cursor.Next(&entry);
        if (FAILED(hr))
        {
            return hr;
        }
        if (current == index)
        {
            located = entry;
        }
    }
    *result = located;
    return S_OK;
}
} // namespace Common::Plugins

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

    Common::Plugins::PackedFileInfoCursor sizingCursor(buffer, bufferSize, entryCount);

    for (unsigned long index = 0; index < entryCount; ++index)
    {
        const FileInfo* entry = nullptr;
        hr = sizingCursor.Next(&entry);
        if (FAILED(hr))
        {
            return hr;
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

    Common::Plugins::PackedFileInfoCursor copyCursor(buffer, bufferSize, entryCount);

    for (unsigned long index = 0; index < entryCount; ++index)
    {
        const FileInfo* entry = nullptr;
        hr = copyCursor.Next(&entry);
        if (FAILED(hr))
        {
            DestroyFileSystemArena(arena);
            return hr;
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
