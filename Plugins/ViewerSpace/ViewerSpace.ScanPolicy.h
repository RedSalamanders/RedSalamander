#pragma once

#include <cstddef>
#include <cstdint>
#include <limits>
#include <span>
#include <string_view>

#include "PlugInterfaces/FileSystem.h"

namespace ViewerSpaceScan
{
inline constexpr uint64_t kMebibyte = 1024ull * 1024ull;

struct ResourcePolicy final
{
    uint32_t maxRetainedDirectories      = 100'000u;
    uint32_t maxRetainedFileRecords      = 200'000u;
    uint32_t maxChildReferences          = 300'000u;
    uint32_t maxChildArenaSlots          = 1'200'000u;
    uint64_t maxRetainedNameBytes        = 64ull * kMebibyte;
    uint64_t maxProviderBufferBytes      = 64ull * kMebibyte;
    uint32_t maxProviderEntriesPerFolder = 500'000u;
    uint32_t maxProviderNameBytes        = 32u * 1024u;
    uint32_t maxTraversalDepth           = 512u;
    uint64_t maxTraversedDirectories     = 5'000'000u;
    uint32_t maxOutstandingDirectories   = 100'000u;
    uint64_t maxOutstandingPathBytes     = 64ull * kMebibyte;
    uint32_t maxPathChars                = 32'767u;
};

inline constexpr ResourcePolicy kProductionResourcePolicy{};

enum class ValidationError : uint8_t
{
    None = 0,
    NullBuffer,
    UsedExceedsAllocated,
    BufferLimit,
    AllocatedBufferLimit,
    EntryCountLimit,
    CountBufferMismatch,
    MisalignedEntry,
    TruncatedHeader,
    InvalidNameLength,
    NameLimit,
    TruncatedName,
    InvalidNextOffset,
    UnsafeName,
    DuplicateName,
    DuplicateId,
    AncestorCycle,
    CountMismatch,
    DepthLimit,
    TraversalLimit,
    OutstandingLimit,
    PathLimit,
    ItemIdLimit,
};

struct ProviderEntryView final
{
    const FileInfo* entry = nullptr;
    std::wstring_view name;
    size_t nextOffset = 0u;
    bool isLast       = false;
};

[[nodiscard]] constexpr uint64_t SaturatingAdd(uint64_t left, uint64_t right) noexcept
{
    return left > (std::numeric_limits<uint64_t>::max)() - right ? (std::numeric_limits<uint64_t>::max)() : left + right;
}

[[nodiscard]] constexpr uint32_t SaturatingAdd(uint32_t left, uint32_t right) noexcept
{
    return left > (std::numeric_limits<uint32_t>::max)() - right ? (std::numeric_limits<uint32_t>::max)() : left + right;
}

[[nodiscard]] constexpr bool TryAddSize(size_t left, size_t right, size_t& out) noexcept
{
    out = 0u;
    if (left > (std::numeric_limits<size_t>::max)() - right)
    {
        return false;
    }
    out = left + right;
    return true;
}

[[nodiscard]] constexpr bool TryAlignUp(size_t value, size_t alignment, size_t& out) noexcept
{
    out = 0u;
    if (alignment == 0u || (alignment & (alignment - 1u)) != 0u)
    {
        return false;
    }
    const size_t mask = alignment - 1u;
    if (value > (std::numeric_limits<size_t>::max)() - mask)
    {
        return false;
    }
    out = (value + mask) & ~mask;
    return true;
}

[[nodiscard]] constexpr ValidationError ValidateProviderBufferContract(const FileInfo* buffer,
                                                                        uint64_t usedBytes,
                                                                        uint64_t allocatedBytes,
                                                                        uint64_t advertisedCount,
                                                                        const ResourcePolicy& policy) noexcept
{
    if (usedBytes > allocatedBytes)
    {
        return ValidationError::UsedExceedsAllocated;
    }
    if (usedBytes > policy.maxProviderBufferBytes)
    {
        return ValidationError::BufferLimit;
    }
    if (allocatedBytes > policy.maxProviderBufferBytes)
    {
        return ValidationError::AllocatedBufferLimit;
    }
    if (advertisedCount > policy.maxProviderEntriesPerFolder)
    {
        return ValidationError::EntryCountLimit;
    }
    if ((buffer == nullptr) != (usedBytes == 0u) || (advertisedCount == 0u) != (usedBytes == 0u))
    {
        return ValidationError::CountBufferMismatch;
    }
    return ValidationError::None;
}

[[nodiscard]] inline ValidationError ValidateProviderEntry(std::span<const std::byte> buffer,
                                                            size_t offset,
                                                            const ResourcePolicy& policy,
                                                            ProviderEntryView& out) noexcept
{
    out = {};
    if (buffer.data() == nullptr)
    {
        return ValidationError::NullBuffer;
    }
    if (offset > buffer.size())
    {
        return ValidationError::TruncatedHeader;
    }

    const size_t remaining = buffer.size() - offset;
    const size_t headerBytes = offsetof(FileInfo, FileName);
    if (remaining < headerBytes)
    {
        return ValidationError::TruncatedHeader;
    }

    const std::byte* recordBytes = buffer.data() + offset;
    if ((reinterpret_cast<uintptr_t>(recordBytes) % alignof(FileInfo)) != 0u || (offset % alignof(FileInfo)) != 0u)
    {
        return ValidationError::MisalignedEntry;
    }

    const auto* entry = reinterpret_cast<const FileInfo*>(recordBytes);
    if ((entry->FileNameSize % sizeof(wchar_t)) != 0u)
    {
        return ValidationError::InvalidNameLength;
    }
    if (entry->FileNameSize > policy.maxProviderNameBytes)
    {
        return ValidationError::NameLimit;
    }

    size_t requiredBytes = 0u;
    if (! TryAddSize(headerBytes, static_cast<size_t>(entry->FileNameSize), requiredBytes) || requiredBytes > remaining)
    {
        return ValidationError::TruncatedName;
    }

    const size_t nextOffset = static_cast<size_t>(entry->NextEntryOffset);
    if (nextOffset != 0u)
    {
        size_t minimumAlignedBytes = 0u;
        if (! TryAlignUp(requiredBytes, alignof(FileInfo), minimumAlignedBytes) || nextOffset < minimumAlignedBytes || nextOffset > remaining ||
            (nextOffset % alignof(FileInfo)) != 0u)
        {
            return ValidationError::InvalidNextOffset;
        }
    }

    out.entry      = entry;
    out.name       = std::wstring_view(entry->FileName, static_cast<size_t>(entry->FileNameSize / sizeof(wchar_t)));
    out.nextOffset = nextOffset;
    out.isLast     = nextOffset == 0u;
    return ValidationError::None;
}

[[nodiscard]] inline ValidationError ValidateProviderChildName(std::wstring_view name) noexcept
{
    if (name.empty() || name == L"." || name == L"..")
    {
        return ValidationError::UnsafeName;
    }
    for (const wchar_t ch : name)
    {
        if (ch == L'\0' || ch == L'/' || ch == L'\\' || ch < 0x20)
        {
            return ValidationError::UnsafeName;
        }
    }
    return ValidationError::None;
}

[[nodiscard]] inline ValidationError ValidateProviderTerminalExtent(std::span<const std::byte> buffer,
                                                                     size_t offset,
                                                                     const ProviderEntryView& entry) noexcept
{
    if (entry.entry == nullptr || ! entry.isLast || offset > buffer.size())
    {
        return ValidationError::TruncatedHeader;
    }

    size_t requiredBytes = 0u;
    size_t nulTerminatedBytes = 0u;
    size_t maximumRecordBytes = 0u;
    if (! TryAddSize(offsetof(FileInfo, FileName), static_cast<size_t>(entry.entry->FileNameSize), requiredBytes) ||
        ! TryAddSize(requiredBytes, sizeof(wchar_t), nulTerminatedBytes) ||
        ! TryAlignUp(nulTerminatedBytes, alignof(FileInfo), maximumRecordBytes))
    {
        return ValidationError::InvalidNextOffset;
    }

    const size_t remaining = buffer.size() - offset;
    return remaining >= requiredBytes && remaining <= maximumRecordBytes ? ValidationError::None : ValidationError::CountMismatch;
}
} // namespace ViewerSpaceScan
