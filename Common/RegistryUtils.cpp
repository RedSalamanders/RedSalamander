#include "RegistryUtils.h"

#include <algorithm>
#include <cwctype>
#include <string_view>

namespace
{
constexpr size_t kMaximumRegistryStringBytes = 1024u * 1024u;
constexpr unsigned int kMaximumReadAttempts  = 8u;

[[nodiscard]] bool IsAllowedType(DWORD type, const Common::Registry::StringReadOptions& options) noexcept
{
    return type == REG_SZ || (type == REG_EXPAND_SZ && options.allowExpandString);
}

void NormalizeString(std::wstring& value, bool trimWhitespace) noexcept
{
    while (! value.empty() && value.back() == L'\0')
    {
        value.pop_back();
    }

    if (! trimWhitespace)
    {
        return;
    }

    const auto first = std::find_if_not(value.begin(), value.end(), [](wchar_t ch) noexcept { return std::iswspace(ch) != 0; });
    const auto last  = std::find_if_not(value.rbegin(), value.rend(), [](wchar_t ch) noexcept { return std::iswspace(ch) != 0; }).base();
    if (first >= last)
    {
        value.clear();
        return;
    }

    value.assign(first, last);
}

[[nodiscard]] bool HasEmbeddedNull(std::wstring_view value) noexcept
{
    return value.find(L'\0') != std::wstring_view::npos;
}

[[nodiscard]] std::optional<std::wstring> ExpandEnvironmentValue(
    std::wstring_view rawValue, size_t maxBytes, unsigned int maxAttempts, bool trimWhitespace, bool allowEmpty) noexcept
{
    const size_t maxChars = maxBytes / sizeof(wchar_t);
    const std::wstring nullTerminated(rawValue);

    for (unsigned int attempt = 0u; attempt < maxAttempts; ++attempt)
    {
        const DWORD requiredChars = ExpandEnvironmentStringsW(nullTerminated.c_str(), nullptr, 0u);
        if (requiredChars == 0u || static_cast<size_t>(requiredChars) > maxChars + 1u)
        {
            return std::nullopt;
        }

        std::wstring expanded(requiredChars, L'\0');
        const DWORD writtenChars = ExpandEnvironmentStringsW(nullTerminated.c_str(), expanded.data(), requiredChars);
        if (writtenChars == 0u)
        {
            return std::nullopt;
        }
        if (writtenChars > requiredChars)
        {
            continue;
        }

        expanded.resize(writtenChars);
        NormalizeString(expanded, trimWhitespace);
        if (HasEmbeddedNull(expanded) || (! allowEmpty && expanded.empty()))
        {
            return std::nullopt;
        }
        return expanded;
    }

    return std::nullopt;
}
} // namespace

namespace Common::Registry
{
std::optional<std::wstring> ReadBoundedStringValue(HKEY key, const wchar_t* valueName, const StringReadOptions& options) noexcept
{
    if (key == nullptr || options.maxBytes == 0u || options.maxAttempts == 0u)
    {
        return std::nullopt;
    }

    const size_t maxBytes          = (std::min)(options.maxBytes, kMaximumRegistryStringBytes);
    const unsigned int maxAttempts = (std::min)(options.maxAttempts, kMaximumReadAttempts);

    for (unsigned int attempt = 0u; attempt < maxAttempts; ++attempt)
    {
        DWORD queriedType         = REG_NONE;
        DWORD queriedByteCount    = 0u;
        const LSTATUS queryStatus = RegQueryValueExW(key, valueName, nullptr, &queriedType, nullptr, &queriedByteCount);
        if (queryStatus != ERROR_SUCCESS || ! IsAllowedType(queriedType, options) || queriedByteCount > maxBytes || (queriedByteCount % sizeof(wchar_t)) != 0u)
        {
            return std::nullopt;
        }

        if (options.afterSizeReadObserver != nullptr)
        {
            options.afterSizeReadObserver(key, valueName, attempt, options.observerCookie);
        }

        const size_t queriedChars = static_cast<size_t>(queriedByteCount) / sizeof(wchar_t);
        std::wstring value(queriedChars + 1u, L'\0');
        DWORD readType           = REG_NONE;
        DWORD readByteCount      = static_cast<DWORD>(value.size() * sizeof(wchar_t));
        const LSTATUS readStatus = RegQueryValueExW(key, valueName, nullptr, &readType, reinterpret_cast<BYTE*>(value.data()), &readByteCount);
        if (readStatus == ERROR_MORE_DATA)
        {
            continue;
        }
        if (readStatus != ERROR_SUCCESS)
        {
            return std::nullopt;
        }
        if (readType != queriedType || readByteCount != queriedByteCount)
        {
            continue;
        }
        if (! IsAllowedType(readType, options) || readByteCount > maxBytes || (readByteCount % sizeof(wchar_t)) != 0u)
        {
            return std::nullopt;
        }

        value.resize(static_cast<size_t>(readByteCount) / sizeof(wchar_t));
        NormalizeString(value, options.trimWhitespace);
        if (HasEmbeddedNull(value) || (! options.allowEmpty && value.empty()))
        {
            return std::nullopt;
        }

        if (readType == REG_EXPAND_SZ && options.expandEnvironmentString)
        {
            return ExpandEnvironmentValue(value, maxBytes, maxAttempts, options.trimWhitespace, options.allowEmpty);
        }
        return value;
    }

    return std::nullopt;
}
} // namespace Common::Registry
