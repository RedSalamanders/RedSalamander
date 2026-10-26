#pragma once

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <functional>
#include <limits>
#include <numeric>
#include <span>
#include <stop_token>
#include <string_view>
#include <utility>
#include <vector>

#include "ViewerSpace.ScanPolicy.h"

namespace ViewerSpaceScan
{
struct LogicalNameLess final
{
    // Provider duplicate-name equivalence, deliberately ordinal. UI name sorting
    // uses Common/NaturalSort.h and must never redefine listing identity.
    bool caseInsensitive = false;

    [[nodiscard]] bool operator()(std::wstring_view left, std::wstring_view right) const noexcept
    {
        if (! caseInsensitive || left.empty() || right.empty())
        {
            return left < right;
        }

        const int comparison = CompareStringOrdinal(left.data(), static_cast<int>(left.size()), right.data(), static_cast<int>(right.size()), TRUE);
        if (comparison == CSTR_LESS_THAN)
        {
            return true;
        }
        if (comparison == CSTR_EQUAL || comparison == CSTR_GREATER_THAN)
        {
            return false;
        }
        return left < right;
    }
};

struct ValidatedListing final
{
    std::vector<uint32_t> entryOffsets;
    uint64_t directoryCount       = 0u;
    uint64_t directoryNameBytes   = 0u;
    uint64_t fileCount            = 0u;
    uint64_t maximumFileNameBytes = 0u;
    ValidationError error         = ValidationError::None;
    bool canceled                 = false;
};

struct ProviderIdOccurrence final
{
    unsigned long id       = 0u;
    uint32_t entryPosition = 0u;
};

namespace Details
{
template <typename T, typename Compare> [[nodiscard]] bool CancellableMergeSort(std::vector<T>& values, Compare compare, std::stop_token stopToken)
{
    if (stopToken.stop_requested())
    {
        return false;
    }

    if (values.size() < 2u)
    {
        return true;
    }

    std::vector<T> scratch(values.size());
    std::vector<T>* source                    = &values;
    std::vector<T>* destination               = &scratch;
    uint32_t operationsSinceCancellationCheck = 0u;

    auto wasCanceled = [&]() noexcept -> bool
    {
        operationsSinceCancellationCheck += 1u;
        if (operationsSinceCancellationCheck < 256u)
        {
            return false;
        }

        operationsSinceCancellationCheck = 0u;
        return stopToken.stop_requested();
    };

    for (size_t width = 1u; width < values.size();)
    {
        size_t begin = 0u;
        while (begin < values.size())
        {
            const size_t middle = begin + (std::min)(width, values.size() - begin);
            const size_t end    = middle + (std::min)(width, values.size() - middle);
            size_t left         = begin;
            size_t right        = middle;
            size_t output       = begin;

            while (left < middle && right < end)
            {
                if (wasCanceled())
                {
                    return false;
                }

                if (compare((*source)[right], (*source)[left]))
                {
                    (*destination)[output++] = (*source)[right++];
                }
                else
                {
                    (*destination)[output++] = (*source)[left++];
                }
            }

            while (left < middle)
            {
                if (wasCanceled())
                {
                    return false;
                }
                (*destination)[output++] = (*source)[left++];
            }

            while (right < end)
            {
                if (wasCanceled())
                {
                    return false;
                }
                (*destination)[output++] = (*source)[right++];
            }

            begin = end;
        }

        std::swap(source, destination);
        if (width > values.size() / 2u)
        {
            break;
        }
        width *= 2u;
    }

    if (source != &values)
    {
        values.swap(*source);
    }
    return ! stopToken.stop_requested();
}

[[nodiscard]] inline std::wstring_view NameAtOffset(std::span<const std::byte> buffer, uint32_t offset) noexcept
{
    const auto* entry = reinterpret_cast<const FileInfo*>(buffer.data() + static_cast<size_t>(offset));
    return std::wstring_view(entry->FileName, static_cast<size_t>(entry->FileNameSize / sizeof(wchar_t)));
}
} // namespace Details

[[nodiscard]] inline ValidatedListing ValidateProviderListing(std::span<const std::byte> buffer,
                                                              uint64_t allocatedBytes,
                                                              uint64_t advertisedCount,
                                                              const ResourcePolicy& policy,
                                                              bool caseInsensitiveNames,
                                                              std::stop_token stopToken)
{
    ValidatedListing result;
    const auto clearAcceptedListing = [&]() noexcept
    {
        result.entryOffsets.clear();
        result.directoryCount       = 0u;
        result.directoryNameBytes   = 0u;
        result.fileCount            = 0u;
        result.maximumFileNameBytes = 0u;
    };
    result.error = ValidateProviderBufferContract(reinterpret_cast<const FileInfo*>(buffer.data()), buffer.size(), allocatedBytes, advertisedCount, policy);
    if (result.error != ValidationError::None)
    {
        return result;
    }

    // Production provider buffers are limited to 64 MiB, so 32-bit offsets keep validation scratch compact.
    if (buffer.size() > static_cast<size_t>((std::numeric_limits<uint32_t>::max)()))
    {
        result.error = ValidationError::BufferLimit;
        return result;
    }
    if (stopToken.stop_requested())
    {
        result.canceled = true;
        return result;
    }

    const size_t headerBytes            = offsetof(FileInfo, FileName);
    const size_t maximumPossibleEntries = headerBytes == 0u ? 0u : (std::min)(static_cast<size_t>(advertisedCount), buffer.size() / headerBytes);
    result.entryOffsets.reserve(maximumPossibleEntries);
    std::vector<uint32_t> nameOrder;
    nameOrder.reserve(maximumPossibleEntries);
    std::vector<ProviderIdOccurrence> providerIds;
    providerIds.reserve((std::min)(maximumPossibleEntries, size_t{4096u}));

    size_t offset             = 0u;
    uint64_t validatedEntries = 0u;
    bool validatedTerminal    = false;
    while (offset < buffer.size())
    {
        if (stopToken.stop_requested())
        {
            clearAcceptedListing();
            result.canceled = true;
            return result;
        }

        ProviderEntryView entryView{};
        result.error = ValidateProviderEntry(buffer, offset, policy, entryView);
        if (result.error != ValidationError::None)
        {
            break;
        }

        validatedEntries = SaturatingAdd(validatedEntries, uint64_t{1u});
        if (validatedEntries > advertisedCount)
        {
            result.error = ValidationError::CountMismatch;
            break;
        }

        result.error = ValidateProviderChildName(entryView.name);
        if (result.error != ValidationError::None)
        {
            break;
        }

        const uint64_t terminatedNameBytes =
            SaturatingAdd(static_cast<uint64_t>(entryView.name.size()) * static_cast<uint64_t>(sizeof(wchar_t)), static_cast<uint64_t>(sizeof(wchar_t)));
        if ((entryView.entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0u)
        {
            result.directoryCount     = SaturatingAdd(result.directoryCount, uint64_t{1u});
            result.directoryNameBytes = SaturatingAdd(result.directoryNameBytes, terminatedNameBytes);
        }
        else
        {
            result.fileCount            = SaturatingAdd(result.fileCount, uint64_t{1u});
            result.maximumFileNameBytes = (std::max)(result.maximumFileNameBytes, terminatedNameBytes);
        }

        const uint32_t entryPosition = static_cast<uint32_t>(result.entryOffsets.size());
        result.entryOffsets.push_back(static_cast<uint32_t>(offset));
        nameOrder.push_back(entryPosition);

        const FileInfo& entry = *entryView.entry;
        if (entry.FileIndex != 0u)
        {
            providerIds.push_back(ProviderIdOccurrence{entry.FileIndex, entryPosition});
        }

        if (entryView.isLast)
        {
            result.error      = ValidateProviderTerminalExtent(buffer, offset, entryView);
            validatedTerminal = result.error == ValidationError::None;
            break;
        }
        offset += entryView.nextOffset;
    }

    if (result.error == ValidationError::None && ((! buffer.empty() && ! validatedTerminal) || validatedEntries != advertisedCount))
    {
        result.error = ValidationError::CountMismatch;
    }

    const LogicalNameLess less{caseInsensitiveNames};
    const auto nameAtIndex = [&](uint32_t index) noexcept -> std::wstring_view { return Details::NameAtOffset(buffer, result.entryOffsets[index]); };
    if (! Details::CancellableMergeSort(
            nameOrder, [&less, &nameAtIndex](uint32_t left, uint32_t right) noexcept { return less(nameAtIndex(left), nameAtIndex(right)); }, stopToken))
    {
        clearAcceptedListing();
        result.error    = ValidationError::None;
        result.canceled = true;
        return result;
    }

    size_t firstNameDuplicatePosition = result.entryOffsets.size();
    for (size_t i = 1u; i < nameOrder.size(); ++i)
    {
        if (stopToken.stop_requested())
        {
            clearAcceptedListing();
            result.error    = ValidationError::None;
            result.canceled = true;
            return result;
        }
        const std::wstring_view previous = nameAtIndex(nameOrder[i - 1u]);
        const std::wstring_view current  = nameAtIndex(nameOrder[i]);
        if (! less(previous, current) && ! less(current, previous))
        {
            firstNameDuplicatePosition = (std::min)(firstNameDuplicatePosition, static_cast<size_t>(nameOrder[i]));
        }
    }

    const auto idLess = [](const ProviderIdOccurrence& left, const ProviderIdOccurrence& right) noexcept
    { return left.id != right.id ? left.id < right.id : left.entryPosition < right.entryPosition; };
    if (! Details::CancellableMergeSort(providerIds, idLess, stopToken))
    {
        clearAcceptedListing();
        result.error    = ValidationError::None;
        result.canceled = true;
        return result;
    }
    size_t firstIdDuplicatePosition = result.entryOffsets.size();
    for (size_t i = 1u; i < providerIds.size(); ++i)
    {
        if (stopToken.stop_requested())
        {
            clearAcceptedListing();
            result.error    = ValidationError::None;
            result.canceled = true;
            return result;
        }
        if (providerIds[i - 1u].id == providerIds[i].id)
        {
            firstIdDuplicatePosition = (std::min)(firstIdDuplicatePosition, static_cast<size_t>(providerIds[i].entryPosition));
        }
    }

    if (firstNameDuplicatePosition < result.entryOffsets.size() && firstNameDuplicatePosition <= firstIdDuplicatePosition)
    {
        result.error = ValidationError::DuplicateName;
    }
    else if (firstIdDuplicatePosition < result.entryOffsets.size())
    {
        result.error = ValidationError::DuplicateId;
    }

    // Identity checks historically ran as each valid record was visited. If a later record is malformed (or the
    // advertised count is wrong), a duplicate in this validated prefix remains the earlier provider-order error.
    if (result.error != ValidationError::None)
    {
        clearAcceptedListing();
    }
    return result;
}
} // namespace ViewerSpaceScan
