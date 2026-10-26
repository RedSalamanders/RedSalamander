#pragma once

#include "FileSystemPathIdentity.h"
#include "PlugInterfaces/FileSystem.h"

#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <limits>
#include <span>
#include <string>
#include <string_view>
#include <vector>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

namespace FileOperationArtifacts
{
enum class Classification : uint8_t
{
    Ordinary,
    Possible,
};

enum class ClassificationReason : uint8_t
{
    None,
    NameShapeOnly,
};

enum class ProbeState : uint8_t
{
    Indeterminate,
    Missing,
    Present,
};

struct Endpoint final
{
    std::wstring pluginId;
    std::wstring instanceId;
    std::wstring profileId;
    std::wstring rootId;
};

struct Identity final
{
    std::vector<std::byte> objectId;
    std::vector<std::byte> revisionId;
    std::wstring pathProfileId;
    FileSystemBoundObjectKind kind = FILESYSTEM_BOUND_OTHER;
};

[[nodiscard]] bool IdentityEquals(const Identity& left, const Identity& right) noexcept;

// I35 R7.1: why a Possible candidate has, or lacks, a current exact identity. A blocked request names
// each item with this reason and logs both results: the provider's no-follow open and the identity
// query that follows a successful open. The binding ABI returns one HRESULT, so the host derives the
// pair from the binding state: an unsupported identity means the open succeeded (or the provider has
// no binding at all) and the identity query is what failed; any other failure is the open's.
enum class IdentityProbeOutcome : uint8_t
{
    NotProbed,           // name shape only (no capabilities, or the provider path could not be probed)
    Bound,               // a current exact identity was captured
    IdentityUnsupported, // the location exposes no object identity (for example a NAS without FileIdInfo)
    OpenFailed,          // the no-follow open failed (sharing violation, access denied, delete pending, ...)
};

struct IdentityProbe final
{
    IdentityProbeOutcome outcome = IdentityProbeOutcome::NotProbed;
    HRESULT openStatus           = S_OK;
    HRESULT identityStatus       = S_FALSE; // S_FALSE: the identity query was not reached
};

// I35 R7.3/R7.6: how the Delete that selects a candidate confirms an item that has no exact identity
// (IdentityUnsupported) by itself. Such an item never enters a touch-guard receipt; this says who
// warns about it instead. Every other touch (Move, Rename, overwrite, publication target, external
// actions) and an item whose open failed keep the guard's refusal.
enum class DeleteHandoff : uint8_t
{
    None,          // not a Delete item, or its route asks nothing more: the guard refuses it
    PermanentCard, // the permanent-delete card (C1/C10 or by name) asks and carries the leftover warning
    RecycleWarn,   // a Recycle: the guard warns (OK/Cancel, no receipt) and a by-name escalation repeats it
};

struct Candidate final
{
    Endpoint endpoint;
    FileSystemPathIdentity pathIdentity;
    std::filesystem::path path;
    ProbeState probeState = ProbeState::Indeterminate;
    Identity currentIdentity;
    IdentityProbe identityProbe;
    DeleteHandoff deleteHandoff = DeleteHandoff::None;
};

struct ClassificationResult final
{
    Classification classification = Classification::Ordinary;
    ClassificationReason reason   = ClassificationReason::None;
    bool possibleNameShape        = false;
};

struct PresentationCapabilities final
{
    bool inspect      = false;
    bool reveal       = false;
    bool openLocation = false;
};

// Name-shape projection only. It never claims that a path is an app-owned artifact and carries no
// durable record or recovery authority.
struct Projection final
{
    Classification classification = Classification::Ordinary;
    ClassificationReason reason   = ClassificationReason::None;
    PresentationCapabilities capabilities;
    Endpoint endpoint;
    std::filesystem::path qualifiedLocation;
    // A backup (.rs_bak_, S3 .rs-bak-, a remote rollback sibling), a rename hop (.rs_ren_) or the
    // case-only rename's temporary name (.rs_case_tmp_) may hold the only complete copy of the user's
    // data; surfaces name it as a possible original, never as an incomplete file.
    bool mayHoldOriginal = false;
};

struct TouchGuardItem final
{
    Classification classification = Classification::Ordinary;
    Endpoint endpoint;
    FileSystemPathIdentity pathIdentity;
    std::filesystem::path path;
    Identity identity;
    IdentityProbe identityProbe;
    bool revalidatable = false; // this item supplied a usable current identity
    // I35 R7.3: a Recycle item without an exact identity: warned about, never in the receipt.
    bool handedOff = false;
};

struct TouchGuardRequest final
{
    std::vector<TouchGuardItem> items;
    bool containsPossible = false;
    // Every item that can enter the receipt (every item not handed off) supplied a usable identity.
    bool allItemsRevalidatable = true;
};

// Ephemeral, one-request receipt created only after the user accepts the warning. It is deliberately
// not serializable and carries no Apply-to-all scope.
struct TouchGuardReceipt final
{
    std::vector<TouchGuardItem> items;
};

struct ArtifactNameShapeProjectionDebugResult final
{
    uint64_t folderLookupUs         = 0u;
    uint64_t findLookupUs           = 0u;
    size_t folderOrdinaryProbeCount = 0u;
    size_t findOrdinaryProbeCount   = 0u;
    size_t folderPossibleProbeCount = 0u;
    size_t findPossibleProbeCount   = 0u;
};

// The classifier is a pure function of the candidate's leaf name. There is no durable record,
// nothing to load, and no recovery authority: a name-shape match only says "Possible".
[[nodiscard]] bool HasPossibleArtifactName(std::wstring_view leaf) noexcept;
// The leftover shapes that may hold the only complete copy of the user's data (.rs_bak_, .rs_ren_,
// .rs_case_tmp_, the S3 .rs-bak- key and the remote .redsalamander-rollback- sibling).
[[nodiscard]] bool HasPossibleOriginalArtifactName(std::wstring_view leaf) noexcept;
// I35 R7.3: a Delete item handed to its own route's consent (DeleteHandoff) instead of the receipt.
[[nodiscard]] bool IsDeleteHandoffCandidate(const Candidate& candidate) noexcept;
// True for the built-in Local file system provider, whose provider paths are Win32 paths that File
// Explorer can open. Only such endpoints offer Open location (I35 R7.2).
[[nodiscard]] bool IsWin32FileSystemEndpoint(const Endpoint& endpoint) noexcept;
// I35 R7.2: inspection-class shell action. Opens File Explorer on the item's parent folder with the
// item selected. It is prompt-free and grants no authority over the item. ENABLE_TESTS builds can
// record the request instead of launching Explorer, so self-tests never open a shell window.
[[nodiscard]] HRESULT OpenLocationInFileExplorer(const std::filesystem::path& win32Path) noexcept;
#ifdef ENABLE_TESTS
void DebugSetOpenLocationRecorderForTests(bool enabled) noexcept;
[[nodiscard]] std::vector<std::wstring> DebugTakeRecordedOpenLocationsForTests();
#endif
[[nodiscard]] ClassificationResult ClassifyCandidate(const Candidate& candidate) noexcept;
[[nodiscard]] Projection ProjectCandidate(const Candidate& candidate);
[[nodiscard]] HRESULT BuildTouchGuardRequest(std::span<const Candidate> candidates, TouchGuardRequest& out) noexcept;
[[nodiscard]] HRESULT AcceptTouchGuard(const TouchGuardRequest& request, TouchGuardReceipt& out) noexcept;
[[nodiscard]] HRESULT RevalidateTouchGuard(const TouchGuardReceipt& receipt, std::span<const Candidate> candidates) noexcept;

[[nodiscard]] HRESULT CaptureProviderObjectCandidate(
    IFileSystem* fileSystem, std::wstring_view providerPath, std::wstring_view pluginId, std::wstring_view instanceContext, Candidate& out) noexcept;
[[nodiscard]] HRESULT RevalidateProviderTouchGuard(IFileSystem* fileSystem,
                                                   std::wstring_view pluginId,
                                                   std::wstring_view instanceContext,
                                                   const TouchGuardReceipt& receipt,
                                                   std::span<const std::filesystem::path> providerPaths) noexcept;
[[nodiscard]] HRESULT ProjectProviderObject(
    IFileSystem* fileSystem, std::wstring_view providerPath, std::wstring_view pluginId, std::wstring_view instanceContext, Projection& out) noexcept;
[[nodiscard]] HRESULT ProjectProviderChildObject(IFileSystem* fileSystem,
                                                 std::wstring_view providerFolderPath,
                                                 std::wstring_view childLeaf,
                                                 std::wstring_view pluginId,
                                                 std::wstring_view instanceContext,
                                                 Projection& out) noexcept;

[[nodiscard]] HRESULT DebugMeasureArtifactNameShapeProjectionForTests(size_t ordinaryRowCount,
                                                                      size_t possibleRowCount,
                                                                      ArtifactNameShapeProjectionDebugResult& out) noexcept;
} // namespace FileOperationArtifacts
