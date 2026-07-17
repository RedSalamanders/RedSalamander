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

    const SIZE_T bytes = (text.size() + 1u) * sizeof(wchar_t);
    wil::unique_hglobal storage(GlobalAlloc(GMEM_MOVEABLE, bytes));
    if (! storage)
    {
        return false;
    }

    {
        auto* buffer = static_cast<wchar_t*>(GlobalLock(storage.get()));
        if (! buffer)
        {
            return false;
        }
        const auto unlockStorage = wil::scope_exit([handle = storage.get()]() noexcept { static_cast<void>(GlobalUnlock(handle)); });
        if (! text.empty())
        {
            std::memcpy(buffer, text.data(), text.size() * sizeof(wchar_t));
        }
        buffer[text.size()] = L'\0';
    }

    // Prepare the complete payload before changing the system clipboard. WIL retains ownership until
    // SetClipboardData succeeds and Windows takes it.
    if (OpenClipboard(ownerWindow) == FALSE)
    {
        return false;
    }
    const auto closeClipboard = wil::scope_exit([]() noexcept { static_cast<void>(CloseClipboard()); });

    if (EmptyClipboard() == FALSE || SetClipboardData(CF_UNICODETEXT, storage.get()) == nullptr)
    {
        return false;
    }
    static_cast<void>(storage.release());
    return true;
}
} // namespace Common::Clipboard
