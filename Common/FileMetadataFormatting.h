#pragma once

#include <cstdint>
#include <format>
#include <string>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>

namespace Common::FileMetadata
{
// Metadata has already been normalized from the filesystem ABI into Windows FILETIME ticks and attribute bits.
struct NormalizedMetadata
{
    int64_t lastWriteTime100nsSince1601 = 0;
    uint32_t fileAttributes             = 0u;
};

enum class DisplayProfile
{
    CompactDetails,
};

struct DisplayFields
{
    std::wstring localTime;
    std::wstring attributes;
};

// CompactDetails uses local time at minute precision and the established RHSACETOP attribute ordering.
// Size formatting and omission remain with the caller because directories and unknown sizes have different
// localized presentation contracts in Folder View, Compare Directories, and the status bar.
[[nodiscard]] inline DisplayFields FormatDisplayFields(const NormalizedMetadata& metadata, DisplayProfile profile) noexcept
{
    if (profile != DisplayProfile::CompactDetails)
    {
        return {};
    }

    DisplayFields fields;
    if (metadata.lastWriteTime100nsSince1601 > 0)
    {
        ULARGE_INTEGER ticks{};
        ticks.QuadPart = static_cast<ULONGLONG>(metadata.lastWriteTime100nsSince1601);

        FILETIME utcFileTime{};
        utcFileTime.dwLowDateTime  = ticks.LowPart;
        utcFileTime.dwHighDateTime = ticks.HighPart;

        FILETIME localFileTime{};
        SYSTEMTIME localSystemTime{};
        if (FileTimeToLocalFileTime(&utcFileTime, &localFileTime) != FALSE && FileTimeToSystemTime(&localFileTime, &localSystemTime) != FALSE)
        {
            fields.localTime = std::format(L"{:04d}-{:02d}-{:02d} {:02d}:{:02d}",
                                           localSystemTime.wYear,
                                           localSystemTime.wMonth,
                                           localSystemTime.wDay,
                                           localSystemTime.wHour,
                                           localSystemTime.wMinute);
        }
    }

    fields.attributes.reserve(10u);
    const auto addAttribute = [&](uint32_t flag, wchar_t label) noexcept
    {
        if ((metadata.fileAttributes & flag) != 0u)
        {
            fields.attributes.push_back(label);
        }
    };

    addAttribute(FILE_ATTRIBUTE_READONLY, L'R');
    addAttribute(FILE_ATTRIBUTE_HIDDEN, L'H');
    addAttribute(FILE_ATTRIBUTE_SYSTEM, L'S');
    addAttribute(FILE_ATTRIBUTE_ARCHIVE, L'A');
    addAttribute(FILE_ATTRIBUTE_COMPRESSED, L'C');
    addAttribute(FILE_ATTRIBUTE_ENCRYPTED, L'E');
    addAttribute(FILE_ATTRIBUTE_TEMPORARY, L'T');
    addAttribute(FILE_ATTRIBUTE_OFFLINE, L'O');
    addAttribute(FILE_ATTRIBUTE_REPARSE_POINT, L'P');

    if (fields.attributes.empty())
    {
        fields.attributes = L"-";
    }
    return fields;
}
} // namespace Common::FileMetadata
