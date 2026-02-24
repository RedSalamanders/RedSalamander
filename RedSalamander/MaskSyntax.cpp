#include "MaskSyntax.h"

#include "Helpers.h"

#include <algorithm>
#include <cwctype>
#include <limits>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>

namespace
{
[[nodiscard]] bool EqualsNoCase(std::wstring_view a, std::wstring_view b) noexcept
{
    if (a.size() != b.size())
    {
        return false;
    }

    if (a.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        return false;
    }

    return OrdinalString::EqualsNoCase(a, b);
}

[[nodiscard]] wchar_t LowerInvariant(wchar_t ch) noexcept
{
    wchar_t buf[2] = {ch, L'\0'};
    ::CharLowerW(buf);
    return buf[0];
}

[[nodiscard]] bool WildcardMatchNoCase(std::wstring_view text, std::wstring_view pattern) noexcept
{
    // Glob match with '*' and '?', case-insensitive.
    size_t ti = 0;
    size_t pi = 0;

    size_t star  = std::wstring_view::npos;
    size_t match = 0;

    while (ti < text.size())
    {
        if (pi < pattern.size())
        {
            const wchar_t pch = pattern[pi];
            if (pch == L'?')
            {
                ++ti;
                ++pi;
                continue;
            }
            if (pch == L'*')
            {
                star  = pi++;
                match = ti;
                continue;
            }

            if (LowerInvariant(text[ti]) == LowerInvariant(pch))
            {
                ++ti;
                ++pi;
                continue;
            }
        }

        if (star != std::wstring_view::npos)
        {
            pi = star + 1;
            ++match;
            ti = match;
            continue;
        }

        return false;
    }

    while (pi < pattern.size() && pattern[pi] == L'*')
    {
        ++pi;
    }

    return pi == pattern.size();
}

[[nodiscard]] bool MatchesAnyMask(std::wstring_view text, const std::vector<std::wstring>& patterns) noexcept
{
    for (const auto& pat : patterns)
    {
        if (pat.empty())
        {
            continue;
        }
        if (WildcardMatchNoCase(text, pat))
        {
            return true;
        }
    }

    return false;
}

[[nodiscard]] std::vector<std::wstring> SplitMaskList(std::wstring_view text)
{
    std::vector<std::wstring> result;
    if (text.empty())
    {
        return result;
    }

    std::wstring token;
    token.reserve(text.size());

    for (size_t i = 0; i < text.size(); ++i)
    {
        const wchar_t ch = text[i];
        if (ch != L';')
        {
            token.push_back(ch);
            continue;
        }

        if ((i + 1) < text.size() && text[i + 1] == L';')
        {
            token.push_back(L';');
            ++i;
            continue;
        }

        std::wstring trimmed = StringUtils::TrimWhitespaceCopy(token);
        if (! trimmed.empty())
        {
            result.push_back(std::move(trimmed));
        }
        token.clear();
    }

    std::wstring trimmed = StringUtils::TrimWhitespaceCopy(token);
    if (! trimmed.empty())
    {
        result.push_back(std::move(trimmed));
    }

    return result;
}
} // namespace

MaskSyntax::WildcardMask MaskSyntax::ParseWildcardMask(std::wstring_view rawText)
{
    WildcardMask result;

    const std::wstring text = StringUtils::TrimWhitespaceCopy(rawText);
    if (text.empty())
    {
        return result;
    }

    const size_t pipe = text.find(L'|');
    std::wstring_view includeText(text);
    std::wstring_view excludeText;
    if (pipe != std::wstring::npos)
    {
        includeText = std::wstring_view(text).substr(0, pipe);
        excludeText = std::wstring_view(text).substr(pipe + 1);
    }

    result.includePatterns = SplitMaskList(includeText);
    result.excludePatterns = SplitMaskList(excludeText);
    return result;
}

bool MaskSyntax::MatchesWildcardMask(std::wstring_view text, const WildcardMask& mask) noexcept
{
    const bool includeMatch = mask.includePatterns.empty() ? true : MatchesAnyMask(text, mask.includePatterns);
    if (! includeMatch)
    {
        return false;
    }

    return ! MatchesAnyMask(text, mask.excludePatterns);
}

void MaskSyntax::NormalizeWildcardMaskHistory(std::vector<std::wstring>& history, size_t maxItems)
{
    std::vector<std::wstring> normalized;
    normalized.reserve(std::min(history.size(), maxItems));

    for (auto& entry : history)
    {
        std::wstring trimmed = StringUtils::TrimWhitespaceCopy(entry);
        if (trimmed.empty())
        {
            continue;
        }

        const std::wstring_view trimmedView(trimmed);
        const bool exists =
            std::find_if(normalized.begin(), normalized.end(), [&](const std::wstring& existing) noexcept { return EqualsNoCase(existing, trimmedView); }) !=
            normalized.end();
        if (exists)
        {
            continue;
        }

        normalized.push_back(std::move(trimmed));
        if (normalized.size() >= maxItems)
        {
            break;
        }
    }

    history = std::move(normalized);
}

void MaskSyntax::AddToWildcardMaskHistory(std::vector<std::wstring>& history, size_t maxItems, std::wstring_view entry)
{
    if (maxItems == 0)
    {
        return;
    }

    std::wstring trimmed = StringUtils::TrimWhitespaceCopy(entry);
    if (trimmed.empty())
    {
        return;
    }

    const std::wstring_view trimmedView(trimmed);
    auto it = std::find_if(history.begin(), history.end(), [&](const std::wstring& existing) noexcept { return EqualsNoCase(existing, trimmedView); });

    if (it != history.end())
    {
        if (it == history.begin())
        {
            return;
        }

        std::wstring moved = std::move(*it);
        history.erase(it);
        history.insert(history.begin(), std::move(moved));
        return;
    }

    history.insert(history.begin(), std::move(trimmed));
    if (history.size() > maxItems)
    {
        history.resize(maxItems);
    }
}
