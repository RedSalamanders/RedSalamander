#pragma once

#include <cstring>
#include <limits>
#include <string_view>

#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>

#pragma warning(push)
#pragma warning(disable : 4625 4626 5026 5027 4820 28182)
#include <wil/resource.h>
#pragma warning(pop)

namespace Common::Clipboard
{
enum class EmptyUnicodeTextPolicy
{
    Allow,
    Reject,
};

[[nodiscard]] inline bool TrySetUnicodeText(
    HWND ownerWindow, std::wstring_view text, EmptyUnicodeTextPolicy emptyPolicy = EmptyUnicodeTextPolicy::Allow) noexcept
{
    if (text.empty() && emptyPolicy == EmptyUnicodeTextPolicy::Reject)
    {
        return false;
    }
    if (text.size() >= (std::numeric_limits<size_t>::max)() / sizeof(wchar_t))
    {
        return false;
    }
    if (OpenClipboard(ownerWindow) == FALSE)
    {
        return false;
    }
    auto closeClipboard = wil::scope_exit([] { static_cast<void>(CloseClipboard()); });

    if (EmptyClipboard() == FALSE)
    {
        return false;
    }

    const SIZE_T bytes = (text.size() + 1u) * sizeof(wchar_t);
    wil::unique_hglobal storage(GlobalAlloc(GMEM_MOVEABLE, bytes));
    if (! storage)
    {
        return false;
    }

    auto* buffer = static_cast<wchar_t*>(GlobalLock(storage.get()));
    if (! buffer)
    {
        return false;
    }
    if (! text.empty())
    {
        std::memcpy(buffer, text.data(), text.size() * sizeof(wchar_t));
    }
    buffer[text.size()] = L'\0';
    static_cast<void>(GlobalUnlock(storage.get()));

    if (SetClipboardData(CF_UNICODETEXT, storage.get()) == nullptr)
    {
        return false;
    }

    storage.release();
    return true;
}
} // namespace Common::Clipboard
