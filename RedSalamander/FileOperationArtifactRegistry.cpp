#include "Framework.h"

#include "FileOperationArtifactRegistry.h"

#include "Helpers.h"

#include <algorithm>
#include <chrono>
#include <format>

namespace FileOperationArtifacts
{
namespace
{
[[nodiscard]] bool IsAsciiHex(const wchar_t value) noexcept
{
    return (value >= L'0' && value <= L'9') || (value >= L'a' && value <= L'f') || (value >= L'A' && value <= L'F');
}

[[nodiscard]] bool IsAsciiHexRun(const std::wstring_view value) noexcept
{
    return ! value.empty() && std::ranges::all_of(value, IsAsciiHex);
}

[[nodiscard]] bool IsGuidWithBraces(const std::wstring_view value) noexcept
{
    if (value.size() != 38u || value.front() != L'{' || value.back() != L'}')
    {
        return false;
    }
    for (size_t index = 1u; index + 1u < value.size(); ++index)
    {
        const bool dash = index == 9u || index == 14u || index == 19u || index == 24u;
        if ((dash && value[index] != L'-') || (! dash && ! IsAsciiHex(value[index])))
        {
            return false;
        }
    }
    return true;
}

[[nodiscard]] bool HasTail(const std::wstring_view leaf, const std::wstring_view marker, bool (*matches)(std::wstring_view) noexcept) noexcept
{
    const size_t offset = leaf.rfind(marker);
    return offset != std::wstring_view::npos && matches(leaf.substr(offset + marker.size()));
}

[[nodiscard]] bool IsBridgeTail(const std::wstring_view value) noexcept
{
    return value.size() > 33u && IsAsciiHexRun(value.substr(0u, 32u)) && value[32u] == L'_' && IsAsciiHexRun(value.substr(33u));
}

[[nodiscard]] bool IsLegacyCopyTail(const std::wstring_view value) noexcept
{
    return value.size() == 34u && IsAsciiHexRun(value.substr(0u, 8u)) && value[8u] == L'_' && IsAsciiHexRun(value.substr(9u, 8u)) && value[17u] == L'_' &&
           IsAsciiHexRun(value.substr(18u, 16u));
}

[[nodiscard]] bool IsRenameTail(const std::wstring_view value) noexcept
{
    return value.size() == 32u && IsAsciiHexRun(value);
}

[[nodiscard]] bool IsBackupTail(const std::wstring_view value) noexcept
{
    return IsGuidWithBraces(value);
}

[[nodiscard]] bool IsWriterTail(std::wstring_view value) noexcept
{
    if (value.size() <= 4u || value.substr(value.size() - 4u) != L".tmp")
    {
        return false;
    }
    value.remove_suffix(4u);
    return IsGuidWithBraces(value) || (value.size() == 25u && IsAsciiHexRun(value.substr(0u, 8u)) && value[8u] == L'.' && IsAsciiHexRun(value.substr(9u, 16u)));
}

[[nodiscard]] bool EndpointEquals(const Endpoint& left, const Endpoint& right) noexcept
{
    return left.pluginId == right.pluginId && left.instanceId == right.instanceId && left.profileId == right.profileId && left.rootId == right.rootId;
}

[[nodiscard]] bool IdentityUsable(const Identity& identity) noexcept
{
    return ! identity.objectId.empty() && ! identity.pathProfileId.empty();
}
} // namespace

bool IdentityEquals(const Identity& left, const Identity& right) noexcept
{
    return left.objectId == right.objectId && left.revisionId == right.revisionId && left.pathProfileId == right.pathProfileId && left.kind == right.kind;
}

bool HasPossibleArtifactName(const std::wstring_view leaf) noexcept
{
    return HasTail(leaf, L".rs_tmp_", IsBridgeTail) || HasTail(leaf, L".rs_copy_tmp_", IsLegacyCopyTail) || HasTail(leaf, L".~rs-write-", IsWriterTail) ||
           HasTail(leaf, L".rs_bak_", IsBackupTail) || HasTail(leaf, L".rs_ren_", IsRenameTail);
}

ClassificationResult ClassifyCandidate(const Candidate& candidate) noexcept
{
    ClassificationResult result{};
    result.possibleNameShape = HasPossibleArtifactName(candidate.path.filename().native());
    if (result.possibleNameShape)
    {
        result.classification = Classification::Possible;
        result.reason         = ClassificationReason::NameShapeOnly;
    }
    return result;
}

Projection ProjectCandidate(const Candidate& candidate)
{
    const ClassificationResult classified = ClassifyCandidate(candidate);
    Projection projection{
        .classification    = classified.classification,
        .reason            = classified.reason,
        .endpoint          = candidate.endpoint,
        .qualifiedLocation = candidate.path,
    };
    if (classified.classification == Classification::Possible)
    {
        projection.capabilities.inspect      = true;
        projection.capabilities.reveal       = true;
        projection.capabilities.openLocation = true;
    }
    return projection;
}

HRESULT BuildTouchGuardRequest(const std::span<const Candidate> candidates, TouchGuardRequest& out) noexcept
{
    out = {};
    for (const Candidate& candidate : candidates)
    {
        const ClassificationResult classified = ClassifyCandidate(candidate);
        if (classified.classification == Classification::Ordinary)
        {
            continue;
        }
        out.containsPossible      = true;
        const bool revalidatable  = candidate.probeState == ProbeState::Present && IdentityUsable(candidate.currentIdentity);
        out.allItemsRevalidatable = out.allItemsRevalidatable && revalidatable;
        out.items.push_back(TouchGuardItem{
            .classification = classified.classification,
            .endpoint       = candidate.endpoint,
            .pathIdentity   = candidate.pathIdentity,
            .path           = candidate.path,
            .identity       = candidate.currentIdentity,
        });
    }
    return out.items.empty() ? S_FALSE : S_OK;
}

HRESULT AcceptTouchGuard(const TouchGuardRequest& request, TouchGuardReceipt& out) noexcept
{
    out = {};
    if (request.items.empty())
    {
        return E_INVALIDARG;
    }
    if (! request.allItemsRevalidatable)
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    out.items = request.items;
    return S_OK;
}

HRESULT RevalidateTouchGuard(const TouchGuardReceipt& receipt, const std::span<const Candidate> candidates) noexcept
{
    if (receipt.items.empty() || receipt.items.size() != candidates.size())
    {
        return E_INVALIDARG;
    }
    for (size_t index = 0u; index < candidates.size(); ++index)
    {
        const TouchGuardItem& accepted = receipt.items[index];
        const Candidate& current       = candidates[index];
        if (! EndpointEquals(accepted.endpoint, current.endpoint) || ! EquivalentPath(accepted.pathIdentity, accepted.path.native(), current.path.native()) ||
            current.probeState != ProbeState::Present || ! IdentityUsable(current.currentIdentity) ||
            ! IdentityEquals(accepted.identity, current.currentIdentity) || ClassifyCandidate(current).classification == Classification::Ordinary)
        {
            return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
        }
    }
    return S_OK;
}

HRESULT DebugMeasureArtifactNameShapeProjectionForTests(const size_t ordinaryRowCount,
                                                        const size_t possibleRowCount,
                                                        ArtifactNameShapeProjectionDebugResult& out) noexcept
{
    out = {};
    if (ordinaryRowCount == 0u || possibleRowCount == 0u)
    {
        return E_INVALIDARG;
    }

    const auto folderStartedAt = std::chrono::steady_clock::now();
    for (size_t item = 0u; item < ordinaryRowCount; ++item)
    {
        const std::wstring leaf = std::format(L"ordinary_{:05}.dat", item);
        if (HasPossibleArtifactName(leaf))
        {
            ++out.folderOrdinaryProbeCount;
        }
    }
    for (size_t item = 0u; item < possibleRowCount; ++item)
    {
        const std::wstring leaf = std::format(L"artifact_{:05}.rs_ren_{:032X}", item, item);
        if (HasPossibleArtifactName(leaf))
        {
            ++out.folderPossibleProbeCount;
        }
    }
    out.folderLookupUs = Debug::Perf::ElapsedUs(folderStartedAt);

    const auto findStartedAt = std::chrono::steady_clock::now();
    for (size_t item = 0u; item < ordinaryRowCount; ++item)
    {
        const std::wstring path = std::format(L"C:\\search\\ordinary_{:05}.dat", item);
        if (HasPossibleArtifactName(std::filesystem::path(path).filename().native()))
        {
            ++out.findOrdinaryProbeCount;
        }
    }
    for (size_t item = 0u; item < possibleRowCount; ++item)
    {
        const std::wstring path = std::format(L"C:\\search\\artifact_{:05}.rs_ren_{:032X}", item, item);
        if (HasPossibleArtifactName(std::filesystem::path(path).filename().native()))
        {
            ++out.findPossibleProbeCount;
        }
    }
    out.findLookupUs = Debug::Perf::ElapsedUs(findStartedAt);
    return S_OK;
}
} // namespace FileOperationArtifacts
