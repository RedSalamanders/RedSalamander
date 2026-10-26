#pragma once

#include "FileSystem.h"
#include "Helpers.h"

#include <array>
#include <cstddef>
#include <functional>
#include <vector>
// I35 R10.3 candidate (b): the Local writer reserves the destination's allocation for a known size
// before the first content byte (FileAllocationInfo only: the end of file still grows with the
// written bytes, so the written length stays the size proof). Plugin-private and used by the
// provider's own copy routes: the host bridge's IFileWriterExpectedSize carries a Commit size
// contract this reservation must not impose on sources whose size is unknown.
interface __declspec(uuid("fc287588-40a9-47d8-b3b8-ed0f0e046015")) __declspec(novtable) ILocalFileWriterAllocation : public IUnknown
{
    // S_OK: reserved. S_FALSE: the destination cannot reserve space; the bytes arrive as before.
    virtual HRESULT STDMETHODCALLTYPE ReserveAllocation(uint64_t sizeBytes) noexcept = 0;
};

namespace FileSystemInternal
{
struct PathInfo
{
    std::wstring display;
    std::wstring extended;
};

std::wstring MakeAbsolutePath(const std::wstring& path);
std::wstring ToExtendedPath(const std::wstring& path);

bool TryGetUncServerRoot(std::wstring_view path, std::wstring& serverName) noexcept;
// True when `path` lies on a network share: a UNC path, or a drive letter Windows reports as remote
// (a drive-type query, about 3 us).
[[nodiscard]] bool IsRemoteStoragePath(std::wstring_view path) noexcept;

[[nodiscard]] bool IsDotOrDotDot(std::wstring_view name) noexcept;

std::wstring AppendPath(const std::wstring& base, std::wstring_view leaf);
std::wstring AppendPath(const std::wstring& base, const wchar_t* leaf);

std::wstring_view TrimTrailingSeparators(std::wstring_view path) noexcept;
std::wstring_view GetPathLeaf(std::wstring_view path) noexcept;
std::wstring GetPathDirectory(std::wstring_view path);
[[nodiscard]] bool ContainsPathSeparator(std::wstring_view text) noexcept;

PathInfo MakePathInfo(const std::wstring& path);
PathInfo MakePathInfo(const wchar_t* path);

struct StagedPromotionOptions
{
    bool allowReplaceReadOnly        = false;
    bool preserveReplacementReadOnly = false;
    bool replacementIsReparsePoint   = false;
    bool stripTemporaryAttributes    = false;
    bool ignoreReplaceMergeErrors    = true;
};

[[nodiscard]] HRESULT PromoteStagedTempIntoFinalPath(const std::wstring& tempPath,
                                                     const std::wstring& finalPath,
                                                     const StagedPromotionOptions& options) noexcept;

[[nodiscard]] HRESULT ReadBoundLocalLink(HANDLE boundHandle, const FileSystemLinkTransform& transform, FileSystemLinkInformation& information) noexcept;
[[nodiscard]] HRESULT CreateExclusiveLocalLink(const wchar_t* stagePath,
                                               const FileSystemLinkInformation& information,
                                               wil::unique_handle& ownedHandle) noexcept;
[[nodiscard]] HRESULT DeleteBoundLocalDirectoryContents(const wchar_t* path, FileSystemFlags flags, const FileSystemOptions* options) noexcept;
// I35 R7.3 (D-7.3, pinned tree): one entry below a folder deleted by name, with the facts its directory
// listing reports. A link or junction is a leaf (never entered); a folder's size and last-write time are
// not compared (its parent's listing updates the time lazily, and what moves it is an entry of the tree).
struct LocalDeleteManifestEntry final
{
    std::wstring relativePath; // relative to the folder, '\\'-separated
    DWORD attributes   = 0u;
    uint64_t sizeBytes = 0u; // files and links
    uint64_t lastWrite = 0u; // FILETIME ticks
    uint64_t creation  = 0u;
    uint64_t fileIndex = 0u; // 0 when the listing reports none
};

// The whole no-follow subtree of a folder deleted by name, in canonical order: relative paths sorted
// ordinally over their upper-case form (ties by the paths themselves), which also puts every entry after
// the folder that holds it. The BLAKE3 digest folds every entry's path and facts; the corroboration text
// carries it with the counts, so the identity pins exactly the tree that was listed.
struct LocalDeleteManifest final
{
    std::vector<LocalDeleteManifestEntry> entries;
    uint64_t files       = 0u;
    uint64_t directories = 0u;
    uint64_t links       = 0u;
    uint64_t nameBytes   = 0u;
    uint64_t maxDepth    = 0u;
    std::array<std::byte, 32> digest{};
};

// A delete by name lists a folder whole before it is offered and again before it deletes anything, so the
// listing is bounded: entries, bytes of relative names, and Common::FileOperations::kTraversalMaxDepth
// levels. A folder over any bound is refused for a delete by name (ERROR_NOT_ENOUGH_QUOTA).
inline constexpr uint64_t kDeleteByNameManifestMaxEntries   = 200'000u;
inline constexpr uint64_t kDeleteByNameManifestMaxNameBytes = 32ull * 1024ull * 1024ull;

// How a folder's manifest is listed.
struct LocalDeleteManifestListing final
{
    // Polled by every listing thread (cancellation and deadline); may be null.
    const FileSystemOptions* options = nullptr;
    // Folders of one depth listed at the same time.
    unsigned int concurrency = 1u;
    // Called once per listed folder with what it holds (files and links, folders, bytes of files), one call
    // at a time, from whichever thread listed it.
    void (*onFolderListed)(void* cookie, uint64_t files, uint64_t directories, uint64_t bytes) noexcept = nullptr;
    void* cookie                                                                                        = nullptr;
};

// Lists the subtree of the folder open as `folder` (a no-follow handle with FILE_LIST_DIRECTORY, for
// `folderExtendedPath`), every folder of one depth at the same time, each through its own no-follow handle.
// Returns HRESULT_FROM_WIN32(ERROR_NOT_ENOUGH_QUOTA) over a bound (with a log record), ERROR_REVISION_MISMATCH
// when a listed sub-folder is gone or is no longer a folder when it is opened (the tree changed while it was
// listed), and ERROR_OPEN_FAILED for an open or a listing the location refuses with ERROR_NOT_SUPPORTED.
[[nodiscard]] HRESULT ListLocalDeleteManifest(HANDLE folder,
                                              const std::wstring& folderExtendedPath,
                                              const LocalDeleteManifestListing& listing,
                                              LocalDeleteManifest& manifest) noexcept;
// True while the facts of the no-follow handle opened for a listed entry still equal the listed ones: kind,
// compared attributes, creation time, the file index when both report one, and for a file or link its size
// and last-write time (a folder's last-write time moves as its listed children are deleted).
[[nodiscard]] bool LocalDeleteManifestEntryMatches(const LocalDeleteManifestEntry& listed, const BY_HANDLE_FILE_INFORMATION& live) noexcept;
// I35 R7.3 (D-7.3): the corroboration facts of the no-follow handle `file` (opened for `path`), as the
// FileSystemDeleteIdentity text of a delete by name: kind, compared attributes, size (a file), last-write
// and creation times, and the volume serial and 64-bit file index only when the location reports a
// non-zero index (corroboration, never proof: an inode number a server may reuse). A folder given its
// `manifest` also carries the manifest's digest and counts. Returns ERROR_NOT_SUPPORTED for an object that
// has an exact identity (FileIdInfo works): such an object is bound, never deleted by name.
[[nodiscard]] HRESULT ResolveLocalDeleteCorroboration(HANDLE file,
                                                      const wchar_t* path,
                                                      FileSystemDeleteIdentity& identity,
                                                      const LocalDeleteManifest* manifest = nullptr) noexcept;

// I35 NAS nested Copy: which object a folder the Local Copy walker created is. The walker keeps this and the
// folder's path, never a handle, while the folder's children copy, and finds the folder again with them to
// restore its basic information (FileSystem.cpp).
struct LocalDirectoryIdentity final
{
    enum class Proof : uint8_t
    {
        None,      // not recorded: the folder's basic information is not restored
        FileId,    // FileIdInfo: the volume serial number and the 128-bit file id
        FileIndex, // no FileIdInfo: the volume serial number and the nonzero 64-bit file index
    };

    Proof proof                 = Proof::None;
    uint64_t volumeSerialNumber = 0u;
    FILE_ID_128 fileId{};
    uint64_t fileIndex = 0u;
};

// Right after publication, while `published` still holds the folder at `path` open: its persistent FileIdInfo
// identity when it has one, otherwise the volume serial number and nonzero file index of that open handle.
[[nodiscard]] HRESULT CaptureLocalDirectoryIdentity(IFileSystemBoundObject* published, const std::wstring& path, LocalDirectoryIdentity& identity) noexcept;

// Opens `path` again for attribute access only (no DELETE, sharing everything, no-follow) and writes `info` through
// that handle only when it is a folder, not a reparse point, with `identity`. Changes nothing otherwise and returns
// the failure: HRESULT_FROM_WIN32(ERROR_FILE_INVALID) when the path names another object.
[[nodiscard]] HRESULT RestoreLocalDirectoryBasicInformation(const std::wstring& path,
                                                            const LocalDirectoryIdentity& identity,
                                                            const FileSystemBasicInformation& info) noexcept;

// Module anchor for AcquireModuleReferenceFromAddress — keeps the DLL loaded
// while background worker threads or threadpool callbacks are active.
extern const int kFileSystemModuleAnchor;

// Stops and joins the shared background copy/move worker threads.
// Intended to be invoked at a host "quiet point" when the last FileSystem instance is being destroyed.
void ShutdownSharedFileOpsJobScheduler() noexcept;

// Runs `run(index)` for every index of [0, count) on the shared File Operations scheduler, at most
// `concurrency` at once (all on the calling thread when concurrency is 1 or no worker is available), and
// returns once every index ran. `run` must not wait on that scheduler itself.
void RunOnSharedFileOpsScheduler(const FileSystemOptions* options, unsigned int concurrency, size_t count, const std::function<void(size_t)>& run) noexcept;

#if defined(_DEBUG)
void RunDebugPathNormalizationSelfTest(unsigned int& passed, unsigned int& failed) noexcept;
void RunDebugReparseCopyErrorMappingSelfTest(unsigned int& passed, unsigned int& failed) noexcept;
void RunDebugRecycleItemVerdictSelfTest(unsigned int& passed, unsigned int& failed) noexcept;
void RunDebugDirectorySizeErrorPolicySelfTest(unsigned int& passed, unsigned int& failed) noexcept;
void RunDebugSharedFileOpsSchedulerShutdownSelfTest(unsigned int& passed, unsigned int& failed) noexcept;
void RunDebugSearchServiceFallbackCandidateSelfTest(unsigned int& passed, unsigned int& failed) noexcept;
void RunDebugObjectBindingSelfTest(unsigned int& passed, unsigned int& failed) noexcept;
void RunDebugRecycleHeldCancellationSelfTest(unsigned int& passed, unsigned int& failed) noexcept;
#endif
} // namespace FileSystemInternal
