#pragma once

#include <cstddef>
#include <string>
#include <string_view>

namespace Common::Process
{
// Produces one always-quoted argument for the Windows CommandLineToArgvW/
// MSVC argv grammar. Process launch policy, argument validation, and the
// executable/token list remain the caller's responsibility.
[[nodiscard]] inline std::wstring QuoteWindowsCommandLineArgument(std::wstring_view value)
{
    std::wstring result(1u, L'"');
    size_t pendingBackslashes = 0u;
    for (const wchar_t character : value)
    {
        if (character == L'\\')
        {
            ++pendingBackslashes;
            continue;
        }
        if (character == L'"')
        {
            result.append(pendingBackslashes * 2u + 1u, L'\\');
            result.push_back(L'"');
            pendingBackslashes = 0u;
            continue;
        }
        result.append(pendingBackslashes, L'\\');
        pendingBackslashes = 0u;
        result.push_back(character);
    }
    result.append(pendingBackslashes * 2u, L'\\');
    result.push_back(L'"');
    return result;
}
} // namespace Common::Process
