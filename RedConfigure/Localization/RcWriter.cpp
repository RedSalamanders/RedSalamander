#include "RcWriter.h"

#include <algorithm>
#include <unordered_map>

namespace
{
[[nodiscard]] std::wstring EscapeRcString(std::wstring_view text)
{
    std::wstring result;
    result.reserve(text.size());
    for (const wchar_t ch : text)
    {
        switch (ch)
        {
        case L'"':
            result.append(L"\"\"");
            break;
        case L'\r':
            result.append(L"\\r");
            break;
        case L'\n':
            result.append(L"\\n");
            break;
        case L'\t':
            result.append(L"\\t");
            break;
        case L'\\':
            result.append(L"\\\\");
            break;
        default:
            result.push_back(ch);
            break;
        }
    }
    return result;
}
} // namespace

namespace RedConfigure::Localization
{
std::vector<MergedStringEntry> MergeStringTables(std::span<const RcStringEntry> source, std::span<const RcStringEntry> target)
{
    std::unordered_map<std::wstring, std::wstring> targetById;
    targetById.reserve(target.size());
    for (const RcStringEntry& entry : target)
    {
        targetById.try_emplace(entry.id, entry.text);
    }

    std::vector<MergedStringEntry> merged;
    merged.reserve(source.size());
    for (const RcStringEntry& entry : source)
    {
        MergedStringEntry mergedEntry;
        mergedEntry.id         = entry.id;
        mergedEntry.sourceText = entry.text;
        if (const auto it = targetById.find(entry.id); it != targetById.end())
        {
            mergedEntry.targetText = it->second;
        }
        merged.push_back(std::move(mergedEntry));
    }

    std::sort(merged.begin(), merged.end(), [](const auto& lhs, const auto& rhs) noexcept { return lhs.id < rhs.id; });
    return merged;
}

std::wstring BuildSatelliteRcStringTable(std::wstring_view resourceHeader, std::wstring_view cultureName, std::span<const MergedStringEntry> entries)
{
    std::wstring output;
    output.append(L"#include \"");
    output.append(resourceHeader);
    output.append(L"\"\r\n\r\n");
    output.append(L"// Culture: ");
    output.append(cultureName);
    output.append(L"\r\n\r\n");
    output.append(L"STRINGTABLE\r\nBEGIN\r\n");

    for (const MergedStringEntry& entry : entries)
    {
        const std::wstring_view text = entry.targetText.empty() ? std::wstring_view(entry.sourceText) : std::wstring_view(entry.targetText);
        output.append(L"    ");
        output.append(entry.id);
        output.push_back(L' ');
        output.push_back(L'"');
        output.append(EscapeRcString(text));
        output.append(L"\"\r\n");
    }

    output.append(L"END\r\n");
    return output;
}
} // namespace RedConfigure::Localization
