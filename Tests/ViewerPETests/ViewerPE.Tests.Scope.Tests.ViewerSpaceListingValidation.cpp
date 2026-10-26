#include "ViewerPE.Tests.Scope.Tests.ViewerSpaceListingValidation.h"

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <initializer_list>
#include <limits>
#include <span>
#include <stop_token>
#include <string_view>
#include <vector>

#include "ViewerSpace/ViewerSpace.ListingValidation.h"

namespace
{
struct Entry final
{
    std::wstring_view name;
    unsigned long fileIndex = 0u;
    bool isDirectory        = false;
};

[[nodiscard]] size_t EntryExtent(std::wstring_view name) noexcept
{
    size_t nameBytes = 0u;
    size_t withNull  = 0u;
    size_t extent    = 0u;
    if (name.size() > (std::numeric_limits<size_t>::max)() / sizeof(wchar_t))
    {
        return 0u;
    }
    nameBytes = name.size() * sizeof(wchar_t);
    if (! ViewerSpaceScan::TryAddSize(nameBytes, sizeof(wchar_t), withNull) || ! ViewerSpaceScan::TryAddSize(offsetof(FileInfo, FileName), withNull, extent) ||
        ! ViewerSpaceScan::TryAlignUp(extent, alignof(FileInfo), extent))
    {
        return 0u;
    }
    return extent;
}

[[nodiscard]] std::vector<std::byte> BuildBuffer(std::initializer_list<Entry> entries)
{
    size_t totalBytes = 0u;
    for (const Entry& source : entries)
    {
        const size_t extent = EntryExtent(source.name);
        if (extent == 0u || ! ViewerSpaceScan::TryAddSize(totalBytes, extent, totalBytes))
        {
            return {};
        }
    }

    std::vector<std::byte> buffer(totalBytes, std::byte{0});
    size_t offset = 0u;
    size_t index  = 0u;
    for (const Entry& source : entries)
    {
        const size_t extent   = EntryExtent(source.name);
        auto* entry           = reinterpret_cast<FileInfo*>(buffer.data() + offset);
        entry->FileIndex      = source.fileIndex;
        entry->FileAttributes = source.isDirectory ? FILE_ATTRIBUTE_DIRECTORY : FILE_ATTRIBUTE_NORMAL;
        entry->FileNameSize   = static_cast<unsigned long>(source.name.size() * sizeof(wchar_t));
        if (! source.name.empty())
        {
            std::memcpy(entry->FileName, source.name.data(), entry->FileNameSize);
        }
        entry->FileName[source.name.size()] = L'\0';
        entry->NextEntryOffset              = ++index < entries.size() ? static_cast<unsigned long>(extent) : 0u;
        offset += extent;
    }
    return buffer;
}
} // namespace

bool TestViewerSpaceListingValidation(std::wstring_view& failure) noexcept
{
    using ViewerSpaceScan::ValidationError;

    failure          = {};
    bool success     = true;
    const auto check = [&](bool condition, std::wstring_view label) noexcept
    {
        if (! condition)
        {
            success = false;
            if (failure.empty())
            {
                failure = label;
            }
        }
    };

    ViewerSpaceScan::ResourcePolicy policy = ViewerSpaceScan::kProductionResourcePolicy;
    policy.maxProviderBufferBytes          = 4096u;
    policy.maxProviderEntriesPerFolder     = 16u;
    policy.maxProviderNameBytes            = 128u;

    const std::vector<std::byte> empty;
    const auto emptyResult = ViewerSpaceScan::ValidateProviderListing(empty, 0u, 0u, policy, false, {});
    check(emptyResult.error == ValidationError::None && emptyResult.entryOffsets.empty(), L"empty provider listing is valid");

    const std::vector<std::byte> unsorted = BuildBuffer({{L"zeta", 1u}, {L"Alpha", 2u}, {L"beta", 3u}});
    const auto valid                      = ViewerSpaceScan::ValidateProviderListing(unsorted, unsorted.size(), 3u, policy, true, {});
    check(valid.error == ValidationError::None && ! valid.canceled && valid.entryOffsets.size() == 3u && valid.entryOffsets[0] == 0u &&
              valid.entryOffsets[1] > valid.entryOffsets[0] && valid.entryOffsets[2] > valid.entryOffsets[1],
          L"valid listing retains source-order offsets");

    const std::vector<std::byte> accountedBuffer = BuildBuffer({{L"folder", 10u, true}, {L"a.bin", 11u}, {L"long.bin", 12u}});
    const auto accounted                         = ViewerSpaceScan::ValidateProviderListing(accountedBuffer, accountedBuffer.size(), 3u, policy, false, {});
    check(accounted.error == ValidationError::None && accounted.directoryCount == 1u && accounted.directoryNameBytes == 14u && accounted.fileCount == 2u &&
              accounted.maximumFileNameBytes == 18u,
          L"successful decode reports exact directory/file and UTF-16 terminator byte bounds for staging admission");

    const ViewerSpaceScan::LogicalNameLess insensitive{true};
    const ViewerSpaceScan::LogicalNameLess sensitive{false};
    check(! insensitive(L"Name", L"name") && sensitive(L"Name", L"name") && ! sensitive(L"name", L"Name"),
          L"logical name comparison preserves case policy and ordinal tie semantics");

    const std::vector<std::byte> duplicateCase = BuildBuffer({{L"Name", 1u}, {L"middle", 3u}, {L"name", 2u}});
    const auto duplicateInsensitive            = ViewerSpaceScan::ValidateProviderListing(duplicateCase, duplicateCase.size(), 3u, policy, true, {});
    check(duplicateInsensitive.error == ValidationError::DuplicateName && duplicateInsensitive.entryOffsets.empty(),
          L"case-insensitive duplicate rejects the complete listing");
    const auto duplicateSensitive = ViewerSpaceScan::ValidateProviderListing(duplicateCase, duplicateCase.size(), 3u, policy, false, {});
    check(duplicateSensitive.error == ValidationError::None && duplicateSensitive.entryOffsets.size() == 3u, L"case-sensitive distinct names remain accepted");
    const std::vector<std::byte> numericNames = BuildBuffer({{L"file2", 1u}, {L"file02", 2u}, {L"file10", 3u}});
    for (const bool providerIgnoresCase : {false, true})
    {
        const auto numeric = ViewerSpaceScan::ValidateProviderListing(numericNames, numericNames.size(), 3u, policy, providerIgnoresCase, {});
        check(numeric.error == ValidationError::None && numeric.entryOffsets.size() == 3u,
              L"natural display sorting never merges numeric or leading-zero provider names");
    }
    const std::vector<std::byte> duplicateBoth = BuildBuffer({{L"same", 9u}, {L"same", 9u}});
    const auto duplicateBothResult             = ViewerSpaceScan::ValidateProviderListing(duplicateBoth, duplicateBoth.size(), 2u, policy, false, {});
    check(duplicateBothResult.error == ValidationError::DuplicateName, L"same-record duplicate name retains name-first rejection precedence");
    const std::vector<std::byte> idFirst = BuildBuffer({{L"first", 8u}, {L"second", 8u}, {L"second", 9u}});
    const auto idFirstResult             = ViewerSpaceScan::ValidateProviderListing(idFirst, idFirst.size(), 3u, policy, false, {});
    check(idFirstResult.error == ValidationError::DuplicateId, L"earlier duplicate id retains provider-order rejection precedence");
    const std::vector<std::byte> duplicateBeforeCountMismatch = BuildBuffer({{L"repeat", 1u}, {L"repeat", 2u}, {L"tail", 3u}});
    const auto duplicateBeforeCountMismatchResult =
        ViewerSpaceScan::ValidateProviderListing(duplicateBeforeCountMismatch, duplicateBeforeCountMismatch.size(), 4u, policy, false, {});
    check(duplicateBeforeCountMismatchResult.error == ValidationError::DuplicateName && duplicateBeforeCountMismatchResult.entryOffsets.empty(),
          L"an earlier duplicate retains precedence over a final advertised-count mismatch");
    std::vector<std::byte> duplicateBeforeMalformedTail = BuildBuffer({{L"repeat", 1u}, {L"repeat", 2u}, {L"tail", 3u}});
    const size_t malformedTailOffset                    = EntryExtent(L"repeat") * 2u;
    if (malformedTailOffset + offsetof(FileInfo, FileName) <= duplicateBeforeMalformedTail.size())
    {
        auto* tail         = reinterpret_cast<FileInfo*>(duplicateBeforeMalformedTail.data() + malformedTailOffset);
        tail->FileNameSize = 3u;
    }
    const auto duplicateBeforeMalformedTailResult =
        ViewerSpaceScan::ValidateProviderListing(duplicateBeforeMalformedTail, duplicateBeforeMalformedTail.size(), 3u, policy, false, {});
    check(duplicateBeforeMalformedTailResult.error == ValidationError::DuplicateName && duplicateBeforeMalformedTailResult.entryOffsets.empty(),
          L"an earlier duplicate retains precedence over a later malformed record");

    const std::vector<std::byte> duplicateIds = BuildBuffer({{L"first", 77u}, {L"second", 77u}});
    const auto duplicateIdResult              = ViewerSpaceScan::ValidateProviderListing(duplicateIds, duplicateIds.size(), 2u, policy, false, {});
    check(duplicateIdResult.error == ValidationError::DuplicateId && duplicateIdResult.entryOffsets.empty(),
          L"duplicate nonzero provider id rejects the complete listing");
    const std::vector<std::byte> zeroIds = BuildBuffer({{L"first", 0u}, {L"second", 0u}});
    const auto zeroIdResult              = ViewerSpaceScan::ValidateProviderListing(zeroIds, zeroIds.size(), 2u, policy, false, {});
    check(zeroIdResult.error == ValidationError::None && zeroIdResult.entryOffsets.size() == 2u, L"zero provider ids may repeat");

    const std::vector<std::byte> unsafeName = BuildBuffer({{L"bad/name", 1u}});
    const auto unsafeResult                 = ViewerSpaceScan::ValidateProviderListing(unsafeName, unsafeName.size(), 1u, policy, false, {});
    check(unsafeResult.error == ValidationError::UnsafeName && unsafeResult.entryOffsets.empty(), L"unsafe child name rejects before offsets are exposed");

    const std::vector<std::byte> countBuffer = BuildBuffer({{L"only", 1u}});
    const auto countResult                   = ViewerSpaceScan::ValidateProviderListing(countBuffer, countBuffer.size(), 2u, policy, false, {});
    check(countResult.error == ValidationError::CountMismatch && countResult.entryOffsets.empty() && countResult.directoryCount == 0u &&
              countResult.directoryNameBytes == 0u && countResult.fileCount == 0u && countResult.maximumFileNameBytes == 0u,
          L"advertised count mismatch exposes neither offsets nor staging-admission statistics");

    std::vector<std::byte> invalidOffset = BuildBuffer({{L"first", 1u}, {L"second", 2u}});
    if (invalidOffset.size() >= sizeof(FileInfo))
    {
        auto* first = reinterpret_cast<FileInfo*>(invalidOffset.data());
        first->NextEntryOffset += 1u;
    }
    const auto invalidOffsetResult = ViewerSpaceScan::ValidateProviderListing(invalidOffset, invalidOffset.size(), 2u, policy, false, {});
    check(invalidOffsetResult.error == ValidationError::InvalidNextOffset && invalidOffsetResult.entryOffsets.empty(),
          L"misaligned record offset rejects before offsets are exposed");

    std::stop_source canceledSource;
    canceledSource.request_stop();
    const auto canceled = ViewerSpaceScan::ValidateProviderListing(unsorted, unsorted.size(), 3u, policy, true, canceledSource.get_token());
    check(canceled.canceled && canceled.error == ValidationError::None && canceled.entryOffsets.empty(), L"pre-canceled validation exposes no offsets");

    return success;
}
