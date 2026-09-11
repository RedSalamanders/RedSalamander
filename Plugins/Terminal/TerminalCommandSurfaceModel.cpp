#include "Framework.h"

#include "TerminalCommandSurfaceModel.h"

#include "StringConversion.h"

#include <algorithm>
#include <cwctype>
#include <limits>
#include <map>
#include <optional>
#include <ranges>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820)
#include <wil/resource.h>
#pragma warning(pop)

namespace TerminalCommandSurfaceModel
{
namespace
{
constexpr size_t kMaximumSuggestionCharacters = 32u * 1024u;

[[nodiscard]] std::optional<size_t> FindNoCase(std::wstring_view text, std::wstring_view query) noexcept
{
    if (query.empty())
    {
        return uint8_t{0};
    }
    if (text.size() > static_cast<size_t>((std::numeric_limits<int>::max)()) ||
        query.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return std::nullopt;
    }
    const int found = FindStringOrdinal(FIND_FROMSTART,
                                        text.data(),
                                        static_cast<int>(text.size()),
                                        query.data(),
                                        static_cast<int>(query.size()),
                                        TRUE);
    return found < 0 ? std::nullopt : std::optional<size_t>(static_cast<size_t>(found));
}

[[nodiscard]] bool StartsWithNoCase(std::wstring_view text, std::wstring_view query) noexcept
{
    return query.size() <= text.size() &&
        CompareStringOrdinal(text.data(), static_cast<int>(query.size()), query.data(), static_cast<int>(query.size()), TRUE) == CSTR_EQUAL;
}

[[nodiscard]] bool EqualsNoCase(std::wstring_view left, std::wstring_view right) noexcept
{
    return left.size() == right.size() &&
        CompareStringOrdinal(left.data(),
                             static_cast<int>(left.size()),
                             right.data(),
                             static_cast<int>(right.size()),
                             TRUE) == CSTR_EQUAL;
}

[[nodiscard]] bool HasTokenPrefixNoCase(std::wstring_view text, std::wstring_view query) noexcept
{
    if (query.empty())
    {
        return true;
    }
    for (size_t index = 0u; index + query.size() <= text.size(); ++index)
    {
        const bool tokenBoundary = index == 0u || ! std::iswalnum(static_cast<wint_t>(text[index - 1u]));
        if (tokenBoundary && StartsWithNoCase(text.substr(index), query))
        {
            return true;
        }
    }
    return false;
}

[[nodiscard]] bool IsFuzzySubsequenceNoCase(std::wstring_view text, std::wstring_view query) noexcept
{
    size_t queryIndex = 0u;
    for (const wchar_t character : text)
    {
        if (queryIndex < query.size() &&
            std::towupper(static_cast<wint_t>(character)) ==
                std::towupper(static_cast<wint_t>(query[queryIndex])))
        {
            ++queryIndex;
        }
    }
    return queryIndex == query.size();
}

[[nodiscard]] std::optional<uint8_t> SuggestionMatchRank(
    std::wstring_view text, std::wstring_view query) noexcept
{
    if (query.empty() || EqualsNoCase(text, query))
    {
        return uint8_t{0};
    }
    if (StartsWithNoCase(text, query))
    {
        return uint8_t{1};
    }
    if (HasTokenPrefixNoCase(text, query))
    {
        return uint8_t{2};
    }
    if (FindNoCase(text, query).has_value())
    {
        return uint8_t{3};
    }
    if (query.size() > 128u || std::ranges::any_of(query, [](wchar_t character) noexcept
        {
            return std::iswspace(static_cast<wint_t>(character)) != 0;
        }))
    {
        return std::nullopt;
    }
    return IsFuzzySubsequenceNoCase(text, query) ? std::optional<uint8_t>{uint8_t{4}} : std::nullopt;
}

struct OrdinalNoCaseLess final
{
    [[nodiscard]] bool operator()(const std::wstring& left, const std::wstring& right) const noexcept
    {
        return CompareStringOrdinal(left.data(),
                                    static_cast<int>(left.size()),
                                    right.data(),
                                    static_cast<int>(right.size()),
                                    TRUE) == CSTR_LESS_THAN;
    }
};

} // namespace

bool IsSafeSuggestion(std::wstring_view text) noexcept
{
    if (text.empty() || text.size() > kMaximumSuggestionCharacters)
    {
        return false;
    }
    return std::ranges::none_of(text, [](wchar_t character) noexcept
    {
        return character == L'\r' || character == L'\n' || character == L'\0' ||
            (character < L' ' && character != L'\t');
    });
}

FindResult Find(std::wstring_view snapshot, std::wstring_view query, size_t maximumRows, std::stop_token stopToken)
{
    FindResult result;
    if (query.empty() || maximumRows == 0u)
    {
        return result;
    }

    size_t lineNumber = 0u;
    size_t start = 0u;
    while (start <= snapshot.size())
    {
        if (stopToken.stop_requested())
        {
            result.cancelled = true;
            result.rows.clear();
            return result;
        }
        const size_t end = snapshot.find_first_of(L"\r\n", start);
        const size_t lineEnd = end == std::wstring_view::npos ? snapshot.size() : end;
        std::wstring_view line = snapshot.substr(start, lineEnd - start);
        if (const std::optional<size_t> offset = FindNoCase(line, query); offset.has_value())
        {
            ++result.totalMatches;
            if (result.rows.size() < maximumRows)
            {
                result.rows.push_back(FindMatch{.lineNumber = lineNumber, .matchOffset = offset.value(), .text = std::wstring(line)});
            }
        }
        ++result.scannedLines;
        ++lineNumber;
        if (end == std::wstring_view::npos)
        {
            break;
        }
        start = end + 1u;
        if (snapshot[end] == L'\r' && start < snapshot.size() && snapshot[start] == L'\n')
        {
            ++start;
        }
    }
    return result;
}

std::vector<std::wstring> LoadPowerShellHistory(std::wstring_view trustedPath,
                                               size_t maximumBytes,
                                               size_t maximumEntries,
                                               std::stop_token stopToken)
{
    std::vector<std::wstring> result;
    if (trustedPath.empty() || trustedPath.size() > 32767u || maximumBytes == 0u || maximumEntries == 0u ||
        stopToken.stop_requested())
    {
        return result;
    }
    const std::wstring path(trustedPath);
    wil::unique_hfile file(CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                      nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL | FILE_FLAG_SEQUENTIAL_SCAN, nullptr));
    if (! file)
    {
        return result;
    }
    LARGE_INTEGER size{};
    if (GetFileSizeEx(file.get(), &size) == FALSE || size.QuadPart <= 0)
    {
        return result;
    }
    const uint64_t fileSize = static_cast<uint64_t>(size.QuadPart);
    const size_t bytesToRead = static_cast<size_t>(std::min<uint64_t>(fileSize, maximumBytes));
    const uint64_t offset = fileSize - bytesToRead;
    if (offset != 0u)
    {
        LARGE_INTEGER seek{};
        seek.QuadPart = static_cast<LONGLONG>(offset);
        if (SetFilePointerEx(file.get(), seek, nullptr, FILE_BEGIN) == FALSE)
        {
            return result;
        }
    }
    std::string bytes(bytesToRead, '\0');
    DWORD read = 0u;
    if (ReadFile(file.get(), bytes.data(), static_cast<DWORD>(bytes.size()), &read, nullptr) == FALSE || read == 0u)
    {
        return result;
    }
    bytes.resize(read);
    const auto scrubBytes = wil::scope_exit([&bytes]() noexcept
    {
        SecureZeroMemory(bytes.data(), bytes.size());
    });
    if (stopToken.stop_requested())
    {
        return result;
    }
    size_t first = 0u;
    if (offset != 0u)
    {
        first = bytes.find('\n');
        if (first == std::string::npos)
        {
            return result;
        }
        ++first;
    }
    std::wstring decoded = Common::Strings::Utf16FromUtf8ReplacingInvalid(std::string_view(bytes).substr(first));
    const auto scrubDecoded = wil::scope_exit([&decoded]() noexcept
    {
        SecureZeroMemory(decoded.data(), decoded.size() * sizeof(wchar_t));
    });
    size_t start = 0u;
    while (start <= decoded.size())
    {
        if (stopToken.stop_requested())
        {
            return {};
        }
        const size_t end = decoded.find_first_of(L"\r\n", start);
        const size_t lineEnd = end == std::wstring::npos ? decoded.size() : end;
        std::wstring line(decoded.data() + start, lineEnd - start);
        if (IsSafeSuggestion(line))
        {
            result.push_back(std::move(line));
            if (result.size() > maximumEntries)
            {
                result.erase(result.begin(), result.begin() + static_cast<ptrdiff_t>(result.size() - maximumEntries));
            }
        }
        if (end == std::wstring::npos)
        {
            break;
        }
        start = end + 1u;
        if (decoded[end] == L'\r' && start < decoded.size() && decoded[start] == L'\n')
        {
            ++start;
        }
    }
    return result;
}

SuggestionResult FilterSuggestions(const std::vector<std::wstring>& persistedHistory,
                                   const std::vector<std::wstring>& currentSessionHistory,
                                   std::wstring_view query,
                                   size_t maximumRows,
                                   std::stop_token stopToken)
{
    SuggestionResult result;
    if (maximumRows == 0u)
    {
        return result;
    }
    struct Ranked final
    {
        Suggestion row;
        uint8_t matchRank = 0u;
        size_t recency = 0u;
        size_t frequency = 1u;
    };
    std::vector<Ranked> ranked;
    // The tree comparator is the exact same ordinal-ignore-case relation used
    // by search matching. This avoids combining locale-sensitive hashing with
    // a different Windows equality relation.
    std::map<std::wstring, size_t, OrdinalNoCaseLess> indices;
    const auto append = [&](const std::vector<std::wstring>& source, bool currentSession)
    {
        for (size_t reverseIndex = 0u; reverseIndex < source.size(); ++reverseIndex)
        {
            if (stopToken.stop_requested())
            {
                return false;
            }
            const std::wstring& value = source[source.size() - reverseIndex - 1u];
            result.sourceBytes += value.size() * sizeof(wchar_t);
            ++result.sourceEntries;
            if (! IsSafeSuggestion(value))
            {
                continue;
            }
            const std::optional<uint8_t> matchRank = SuggestionMatchRank(value, query);
            if (! matchRank.has_value())
            {
                continue;
            }
            const auto existing = indices.find(value);
            if (existing != indices.end())
            {
                ++ranked[existing->second].frequency;
                continue;
            }
            const size_t rankedIndex = ranked.size();
            ranked.push_back(Ranked{
                .row = Suggestion{.text = value, .currentSession = currentSession},
                .matchRank = matchRank.value(),
                .recency = reverseIndex,
            });
            indices.emplace(ranked.back().row.text, rankedIndex);
        }
        return true;
    };
    if (! append(currentSessionHistory, true) || ! append(persistedHistory, false))
    {
        result.cancelled = true;
        return result;
    }
    std::ranges::stable_sort(ranked, [](const Ranked& left, const Ranked& right) noexcept
    {
        if (left.matchRank != right.matchRank) return left.matchRank < right.matchRank;
        if (left.row.currentSession != right.row.currentSession) return left.row.currentSession;
        if (left.recency != right.recency) return left.recency < right.recency;
        if (left.frequency != right.frequency) return left.frequency > right.frequency;
        const int normalized = CompareStringOrdinal(left.row.text.data(),
                                                    static_cast<int>(left.row.text.size()),
                                                    right.row.text.data(),
                                                    static_cast<int>(right.row.text.size()),
                                                    TRUE);
        if (normalized != CSTR_EQUAL) return normalized == CSTR_LESS_THAN;
        return left.row.text < right.row.text;
    });
    result.rows.reserve(std::min(maximumRows, ranked.size()));
    for (Ranked& value : ranked | std::views::take(maximumRows))
    {
        result.rows.push_back(std::move(value.row));
    }
    return result;
}
} // namespace TerminalCommandSurfaceModel
