#include "PlaceholderValidation.h"

#include <algorithm>
#include <cwctype>
#include <string>
#include <string_view>

namespace
{
[[nodiscard]] bool IsDigit(wchar_t ch) noexcept
{
    return ch >= L'0' && ch <= L'9';
}

[[nodiscard]] bool ContainsPrintfPlaceholder(std::wstring_view text) noexcept
{
    for (size_t index = 0u; index < text.size(); ++index)
    {
        if (text[index] != L'%')
        {
            continue;
        }
        if (index + 1u < text.size() && text[index + 1u] == L'%')
        {
            ++index;
            continue;
        }

        size_t pos = index + 1u;
        while (pos < text.size())
        {
            const wchar_t ch = text[pos];
            if (std::iswalpha(static_cast<wint_t>(ch)) != 0)
            {
                return ch == L's' || ch == L'S' || ch == L'd' || ch == L'i' || ch == L'u' || ch == L'x' || ch == L'X' || ch == L'f' ||
                       ch == L'c' || ch == L'C';
            }
            if (ch == L'{' || ch == L'}' || std::iswspace(static_cast<wint_t>(ch)) != 0)
            {
                break;
            }
            ++pos;
        }
    }
    return false;
}

[[nodiscard]] RedConfigure::Localization::PlaceholderStatus ExtractPlaceholders(std::wstring_view text, std::vector<std::wstring>& out)
{
    out.clear();
    for (size_t index = 0u; index < text.size(); ++index)
    {
        if (text[index] == L'{' && index + 1u < text.size() && text[index + 1u] == L'{')
        {
            ++index;
            continue;
        }
        if (text[index] != L'{')
        {
            continue;
        }

        const size_t start = index;
        const size_t close = text.find(L'}', index + 1u);
        if (close == std::wstring_view::npos)
        {
            continue;
        }

        const std::wstring_view body = text.substr(index + 1u, close - index - 1u);
        if (body.empty())
        {
            return RedConfigure::Localization::PlaceholderStatus::BarePlaceholder;
        }
        if (body.front() == L':')
        {
            return RedConfigure::Localization::PlaceholderStatus::UnindexedFormatSpec;
        }
        if (! IsDigit(body.front()))
        {
            index = close;
            continue;
        }

        size_t pos = 0u;
        while (pos < body.size() && IsDigit(body[pos]))
        {
            ++pos;
        }
        if (pos < body.size() && body[pos] != L':')
        {
            index = close;
            continue;
        }

        out.emplace_back(text.substr(start, close - start + 1u));
        index = close;
    }

    std::sort(out.begin(), out.end());
    return RedConfigure::Localization::PlaceholderStatus::Ok;
}
} // namespace

namespace RedConfigure::Localization
{
PlaceholderValidationResult ValidatePlaceholders(std::wstring_view source, std::wstring_view target)
{
    PlaceholderValidationResult result;
    if (ContainsPrintfPlaceholder(target))
    {
        result.status = PlaceholderStatus::PrintfPlaceholder;
        return result;
    }

    result.status = ExtractPlaceholders(source, result.sourcePlaceholders);
    if (result.status != PlaceholderStatus::Ok)
    {
        return result;
    }

    result.status = ExtractPlaceholders(target, result.targetPlaceholders);
    if (result.status != PlaceholderStatus::Ok)
    {
        return result;
    }

    if (result.sourcePlaceholders != result.targetPlaceholders)
    {
        result.status = PlaceholderStatus::PlaceholderMismatch;
        return result;
    }

    result.status = PlaceholderStatus::Ok;
    return result;
}
} // namespace RedConfigure::Localization
