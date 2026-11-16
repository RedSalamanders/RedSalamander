#include "Framework.h"

#include "FileOperationArtifactRegistry.h"

#include "Helpers.h"
#include "PathUtils.h"

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

// I35 R7.6: `<32 hex>.tmp`, the tail of the MTP writer temp (`.<leaf>.rs-mtp-overwrite-`).
[[nodiscard]] bool IsMtpOverwriteLeaf(const std::wstring_view leaf) noexcept
{
    constexpr std::wstring_view kMarker = L".rs-mtp-overwrite-";
    constexpr size_t kTailSize          = 36u;
    if (leaf.size() < kMarker.size() + kTailSize)
    {
        return false;
    }
    // The fixed-width tail locates the marker without scanning a potentially long provider leaf.
    const size_t offset          = leaf.size() - kMarker.size() - kTailSize;
    const std::wstring_view tail = leaf.substr(leaf.size() - kTailSize);
    return Common::Paths::StartsWithAsciiNoCase(leaf.substr(offset), kMarker) && IsAsciiHexRun(tail.substr(0u, 32u)) &&
           Common::Paths::StartsWithAsciiNoCase(tail.substr(32u), L".tmp");
}

// I35 R7.6: an S3 hidden backup key, `.rs-bak-<32 hex>`, then `-<leaf>` unless the backed-up key was a folder.
[[nodiscard]] bool IsS3BackupLeaf(const std::wstring_view leaf) noexcept
{
    constexpr std::wstring_view kMarker = L".rs-bak-";
    if (! Common::Paths::StartsWithAsciiNoCase(leaf, kMarker) || leaf.size() < kMarker.size() + 32u || ! IsAsciiHexRun(leaf.substr(kMarker.size(), 32u)))
    {
        return false;
    }
    return leaf.size() == kMarker.size() + 32u || leaf[kMarker.size() + 32u] == L'-';
}

// I35 R7.6: `<marker><32 hex>` as the whole leaf: a Curl (FTP/SFTP/SCP) or Microsoft Drive sibling
// (`.redsalamander-<purpose>-`).
[[nodiscard]] bool IsWholeLeafHexName(const std::wstring_view leaf, const std::wstring_view marker) noexcept
{
    return leaf.size() == marker.size() + 32u && Common::Paths::StartsWithAsciiNoCase(leaf, marker) && IsAsciiHexRun(leaf.substr(marker.size()));
}

// I63: a staging artifact created on a Win32 path ends in a real extension, after a 32-hex token:
// `<readable prefix>.<32 hex>[_<stream>].rs-tmp` is incomplete content and
// `<readable prefix>.<32 hex>.rs-bak` may hold an original. The token is required: a name that only
// ends in the extension is an ordinary file. The extension test comes first and is cheap.
[[nodiscard]] bool EndsWithDotToken(const std::wstring_view value) noexcept
{
    return value.size() > 33u && value[value.size() - 33u] == L'.' && IsAsciiHexRun(value.substr(value.size() - 32u));
}

[[nodiscard]] bool HasTokenBeforeExtension(std::wstring_view leaf, const std::wstring_view extension, const bool allowStream) noexcept
{
    if (leaf.size() < extension.size() || ! Common::Paths::StartsWithAsciiNoCase(leaf.substr(leaf.size() - extension.size()), extension))
    {
        return false;
    }
    leaf.remove_suffix(extension.size());
    if (allowStream)
    {
        if (const size_t stream = leaf.rfind(L'_');
            stream != std::wstring_view::npos && IsAsciiHexRun(leaf.substr(stream + 1u)) && EndsWithDotToken(leaf.substr(0u, stream)))
        {
            return true;
        }
    }
    return EndsWithDotToken(leaf);
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

bool IsWin32FileSystemEndpoint(const Endpoint& endpoint) noexcept
{
    constexpr std::wstring_view kLocalFileSystemPluginId = L"builtin/file-system";
    return ! endpoint.pluginId.empty() && ::CompareStringOrdinal(endpoint.pluginId.data(),
                                                                 static_cast<int>(endpoint.pluginId.size()),
                                                                 kLocalFileSystemPluginId.data(),
                                                                 static_cast<int>(kLocalFileSystemPluginId.size()),
                                                                 TRUE) == CSTR_EQUAL;
}

// Each of these may hold the only complete copy of the user's original: a Local backup kept during an
// overwrite or the Local case-only rename's hop (`<leaf>.<32 hex>.rs-bak`), the S3 `.rs-bak-` key and
// the remote `.redsalamander-rollback-` sibling.
bool HasPossibleOriginalArtifactName(const std::wstring_view leaf) noexcept
{
    return HasTokenBeforeExtension(leaf, L".rs-bak", false) || IsS3BackupLeaf(leaf) || IsWholeLeafHexName(leaf, L".redsalamander-rollback-");
}

bool HasPossibleArtifactName(const std::wstring_view leaf) noexcept
{
    // Only the names the current generators write: the Win32 stages (`.rs-tmp`), the provider
    // staging names (I35 R7.6: the Curl upload and writer siblings, the MTP writer temp) and the
    // possible originals above. A name an earlier version wrote is an ordinary file.
    return HasTokenBeforeExtension(leaf, L".rs-tmp", true) || IsMtpOverwriteLeaf(leaf) || IsWholeLeafHexName(leaf, L".redsalamander-upload-") ||
           IsWholeLeafHexName(leaf, L".redsalamander-writer-") || HasPossibleOriginalArtifactName(leaf);
}

bool IsDeleteHandoffCandidate(const Candidate& candidate) noexcept
{
    // Only an item the location opened but cannot identify is handed off; a failed open is never.
    return candidate.deleteHandoff != DeleteHandoff::None && candidate.identityProbe.outcome == IdentityProbeOutcome::IdentityUnsupported;
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
        projection.mayHoldOriginal      = HasPossibleOriginalArtifactName(candidate.path.filename().native());
        projection.capabilities.inspect = true;
        projection.capabilities.reveal  = true;
        // File Explorer can open only a Win32 path; other providers keep Inspect and Reveal.
        projection.capabilities.openLocation = IsWin32FileSystemEndpoint(candidate.endpoint);
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
        // I35 R7.3: a Delete item the location cannot identify is never a receipt item. A permanent
        // delete's card asks for it (with the leftover warning); a Recycle is still warned about here.
        const bool handedOff = IsDeleteHandoffCandidate(candidate);
        if (handedOff && candidate.deleteHandoff == DeleteHandoff::PermanentCard)
        {
            continue;
        }
        out.containsPossible     = true;
        const bool revalidatable = ! handedOff && candidate.probeState == ProbeState::Present && IdentityUsable(candidate.currentIdentity);
        if (! handedOff)
        {
            out.allItemsRevalidatable = out.allItemsRevalidatable && revalidatable;
        }
        out.items.push_back(TouchGuardItem{
            .classification = classified.classification,
            .endpoint       = candidate.endpoint,
            .pathIdentity   = candidate.pathIdentity,
            .path           = candidate.path,
            .identity       = candidate.currentIdentity,
            .identityProbe  = candidate.identityProbe,
            .revalidatable  = revalidatable,
            .handedOff      = handedOff,
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
    // A receipt is minted only when every receipt item carries its own usable identity; the aggregate
    // flag alone is not trusted (a hand-built request could set it with an identity-less item). A
    // handed-off item (I35 R7.3) never enters the receipt and must be what that flag claims: an item
    // the location opened but could not identify.
    if (! request.allItemsRevalidatable || std::ranges::any_of(request.items, [](const TouchGuardItem& item) noexcept {
        return item.handedOff ? item.revalidatable || item.identityProbe.outcome != IdentityProbeOutcome::IdentityUnsupported : ! item.revalidatable;
    }))
    {
        return HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
    }
    for (const TouchGuardItem& item : request.items)
    {
        if (! item.handedOff)
        {
            out.items.push_back(item);
        }
    }
    return S_OK;
}

HRESULT RevalidateTouchGuard(const TouchGuardReceipt& receipt, const std::span<const Candidate> allCandidates) noexcept
{
    // I35 R7.3: a handed-off Delete item never entered the receipt; its own route corroborates it.
    // An accepted item that turned into one no longer matches the receipt and fails closed below.
    std::vector<const Candidate*> candidates;
    candidates.reserve(allCandidates.size());
    for (const Candidate& candidate : allCandidates)
    {
        if (! IsDeleteHandoffCandidate(candidate))
        {
            candidates.push_back(&candidate);
        }
    }
    if (receipt.items.empty())
    {
        return E_INVALIDARG;
    }
    if (receipt.items.size() != candidates.size())
    {
        return HRESULT_FROM_WIN32(ERROR_REVISION_MISMATCH);
    }
    for (size_t index = 0u; index < candidates.size(); ++index)
    {
        const TouchGuardItem& accepted = receipt.items[index];
        const Candidate& current       = *candidates[index];
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
        const std::wstring leaf =
            item % 2u == 0u ? std::format(L"artifact_{:05}.{:032x}_0.rs-tmp", item, item) : std::format(L"artifact_{:05}.{:032x}.rs-bak", item, item);
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
        const std::wstring path = item % 2u == 0u ? std::format(L"C:\\search\\artifact_{:05}.{:032x}_0.rs-tmp", item, item)
                                                  : std::format(L"C:\\search\\artifact_{:05}.{:032x}.rs-bak", item, item);
        if (HasPossibleArtifactName(std::filesystem::path(path).filename().native()))
        {
            ++out.findPossibleProbeCount;
        }
    }
    out.findLookupUs = Debug::Perf::ElapsedUs(findStartedAt);
    return S_OK;
}
} // namespace FileOperationArtifacts
