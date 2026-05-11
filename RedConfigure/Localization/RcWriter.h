#pragma once

#include "RcParser.h"

#include <span>
#include <string>
#include <string_view>
#include <vector>

namespace RedConfigure::Localization
{
struct MergedStringEntry
{
    std::wstring id;
    std::wstring sourceText;
    std::wstring targetText;
};

[[nodiscard]] std::vector<MergedStringEntry> MergeStringTables(std::span<const RcStringEntry> source, std::span<const RcStringEntry> target);

[[nodiscard]] std::wstring BuildSatelliteRcStringTable(std::wstring_view resourceHeader,
                                                       std::wstring_view cultureName,
                                                       std::span<const MergedStringEntry> entries);
} // namespace RedConfigure::Localization
