#pragma once

#include "FileSystemPathIdentity.h"
#include "PlugInterfaces/FileSystem.h"

#include <Windows.h>

#include <cstddef>
#include <cstdint>
#include <filesystem>
#include <limits>
#include <span>
#include <string>
#include <string_view>
#include <vector>

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

struct Candidate final
{
    Endpoint endpoint;
    FileSystemPathIdentity pathIdentity;
    std::filesystem::path path;
    ProbeState probeState = ProbeState::Indeterminate;
    Identity currentIdentity;
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
};

struct TouchGuardItem final
{
    Classification classification = Classification::Ordinary;
    Endpoint endpoint;
    FileSystemPathIdentity pathIdentity;
    std::filesystem::path path;
    Identity identity;
};

struct TouchGuardRequest final
{
    std::vector<TouchGuardItem> items;
    bool containsPossible      = false;
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
