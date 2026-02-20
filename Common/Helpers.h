#pragma once

#include <algorithm>
#include <array>
#include <atomic>
#include <chrono>
#include <cstdint>
#include <cstring>
#include <exception>
#include <filesystem>
#include <format>
#include <limits>
#include <locale>
#include <mutex>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>

#include <evntrace.h>

#pragma warning(push)
// WIL and TraceLogging: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted),
// C5027 (move assign deleted), C4820 (padding)
#pragma warning(disable : 4625 4626 5026 5027 4820)
#include <TraceLoggingProvider.h>
#include <wil/resource.h>
#pragma warning(pop)

#include <Helpers.h>

#ifdef min
#undef min
#endif
#ifdef max
#undef max
#endif

#pragma warning(push)
#pragma warning(disable : 4514) // unreferenced inline function has been removed

namespace Debug
{
// predefine a TraceLogging provider for use in other modules
template <typename... Args> inline DWORD ErrorWithLastError(std::wformat_string<Args...> format, Args&&... args) noexcept;
inline DWORD ErrorWithLastError(const std::wstring& message) noexcept;
}; // namespace Debug

namespace LocaleFormatting
{
inline void InvalidateFormatLocaleCache() noexcept;
inline const std::locale& GetFormatLocale() noexcept;
} // namespace LocaleFormatting

namespace OrdinalString
{
inline int Compare(std::wstring_view a, std::wstring_view b, bool ignoreCase) noexcept
{
    if (a.size() > static_cast<size_t>(std::numeric_limits<int>::max()) || b.size() > static_cast<size_t>(std::numeric_limits<int>::max()))
    {
        const int fallback = a.compare(b);
        return (fallback < 0) ? -1 : ((fallback > 0) ? 1 : 0);
    }

    const int aLen   = static_cast<int>(a.size());
    const int bLen   = static_cast<int>(b.size());
    const int result = CompareStringOrdinal(a.data(), aLen, b.data(), bLen, ignoreCase ? TRUE : FALSE);

    if (result == CSTR_LESS_THAN)
    {
        return -1;
    }
    if (result == CSTR_GREATER_THAN)
    {
        return 1;
    }
    if (result == CSTR_EQUAL)
    {
        return 0;
    }

    const int fallback = a.compare(b);
    return (fallback < 0) ? -1 : ((fallback > 0) ? 1 : 0);
}

inline bool EqualsNoCase(std::wstring_view a, std::wstring_view b) noexcept
{
    return Compare(a, b, true) == 0;
}

inline bool LessNoCase(std::wstring_view a, std::wstring_view b) noexcept
{
    const int cmp = Compare(a, b, true);
    if (cmp != 0)
    {
        return cmp < 0;
    }

    const int caseCmp = Compare(a, b, false);
    if (caseCmp != 0)
    {
        return caseCmp < 0;
    }

    return false;
}
} // namespace OrdinalString

// LoadString from resource ID
template <typename string_type, size_t stackBufferLength = 256>
int LoadStringResource(_In_opt_ HINSTANCE hInstance, _In_ UINT uID, string_type& result) WI_NOEXCEPT
{
    static_assert(stackBufferLength <= INT_MAX, "stackBufferLength must fit in int");
    const HINSTANCE instance = hInstance ? hInstance : GetModuleHandleW(nullptr);

    // LoadStringW supports returning a pointer directly to the resource string when cchBufferMax == 0.
    // This avoids guessing the required buffer size and supports embedded NULs (e.g. file dialog filters).
    PCWSTR ptr       = nullptr;
    const int length = ::LoadStringW(instance, uID, reinterpret_cast<LPWSTR>(&ptr), 0);
    if (length <= 0 || ! ptr)
    {
        result.clear();
        return 0;
    }

    result.assign(ptr, static_cast<size_t>(length));
    return length;
}

// Convenience overload returning std::wstring.
inline std::wstring LoadStringResource(_In_opt_ HINSTANCE hInstance, _In_ UINT uID) WI_NOEXCEPT
{
    std::wstring result;
    LoadStringResource(hInstance, uID, result);
    return result;
}

// Loads a resource string and formats it using std::format-style placeholders.
// Uses std::vformat since resource strings are runtime values (not compile-time format strings).
template <typename... Args> std::wstring FormatStringResource(_In_opt_ HINSTANCE hInstance, _In_ UINT uID, Args... args)
{
    std::wstring fmt;
    LoadStringResource(hInstance, uID, fmt);
    if (fmt.empty())
    {
        return {};
    }

    // std::make_wformat_args requires non-const lvalue references; take args by value so we can safely pass lvalues.
    return std::vformat(LocaleFormatting::GetFormatLocale(), std::wstring_view(fmt), std::make_wformat_args(args...));
}

namespace LocaleFormatting
{
inline std::atomic_uint32_t g_formatLocaleGeneration{1u};
inline std::mutex g_cachedFormatLocaleMutex;
inline std::locale g_cachedFormatLocale        = std::locale::classic();
inline uint32_t g_cachedFormatLocaleGeneration = 0u;

inline thread_local uint32_t g_threadFormatLocaleGeneration = 0u;
inline thread_local std::locale g_threadFormatLocale        = std::locale::classic();

inline void InvalidateFormatLocaleCache() noexcept
{
    g_formatLocaleGeneration.fetch_add(1u, std::memory_order_acq_rel);
}

inline const std::locale& GetFormatLocale() noexcept
{
    const uint32_t currentGeneration = g_formatLocaleGeneration.load(std::memory_order_acquire);
    if (g_threadFormatLocaleGeneration == currentGeneration)
    {
        return g_threadFormatLocale;
    }

    std::scoped_lock lock(g_cachedFormatLocaleMutex);
    if (g_cachedFormatLocaleGeneration != currentGeneration)
    {
        g_cachedFormatLocale           = std::locale("");
        g_cachedFormatLocaleGeneration = currentGeneration;
    }

    g_threadFormatLocale           = g_cachedFormatLocale;
    g_threadFormatLocaleGeneration = currentGeneration;
    return g_threadFormatLocale;
}
} // namespace LocaleFormatting

// Formats byte sizes as "B/KB/MB/GB/TB" with compact significant digits:
// - 1-digit integer part: 2 decimals (e.g. 4.60 MB)
// - 2-digit integer part: 1 decimal (e.g. 12.3 MB)
// - 3+ digit integer part: no decimals (e.g. 156 GB)
inline std::wstring FormatBytesCompact(uint64_t bytes)
{
    static constexpr std::array<std::wstring_view, 5> suffixes = {
        std::wstring_view(L"B"),
        std::wstring_view(L"KB"),
        std::wstring_view(L"MB"),
        std::wstring_view(L"GB"),
        std::wstring_view(L"TB"),
    };

    double value       = static_cast<double>(bytes);
    size_t suffixIndex = 0;
    while (value >= 1024.0 && (suffixIndex + 1) < suffixes.size())
    {
        value /= 1024.0;
        suffixIndex += 1;
    }

    if (suffixIndex == 0)
    {
        return std::format(LocaleFormatting::GetFormatLocale(), L"{:L} {}", bytes, suffixes[suffixIndex]);
    }

    int decimals = 0;
    if (value < 10.0)
    {
        decimals = (value >= 9.995) ? 1 : 2;
    }
    else if (value < 100.0)
    {
        decimals = (value >= 99.95) ? 0 : 1;
    }
    else
    {
        decimals = 0;
    }

    return std::format(LocaleFormatting::GetFormatLocale(), L"{:.{}Lf} {}", value, decimals, suffixes[suffixIndex]);
}

inline int MessageBoxThemedImpl(_In_opt_ HWND owner, PCWSTR text, PCWSTR caption, _In_ UINT type, bool centerOnOwner);

inline int MessageBoxResource(_In_opt_ HWND owner, _In_opt_ HINSTANCE hInstance, _In_ UINT textId, _In_ UINT captionId, _In_ UINT type)
{
    const std::wstring text    = LoadStringResource(hInstance, textId);
    const std::wstring caption = LoadStringResource(hInstance, captionId);
    return MessageBoxThemedImpl(owner, text.c_str(), caption.c_str(), type, false);
}

// Thread-local storage for MessageBox centering hook
namespace MessageBoxCenteringDetail
{
inline thread_local HWND g_centerOnWindow      = nullptr;
inline thread_local HHOOK g_hook               = nullptr;
inline thread_local WNDPROC g_msgBoxWndProc    = nullptr;
inline thread_local bool g_themeEnabled        = false;
inline thread_local bool g_themeUseDarkMode    = false;
inline thread_local COLORREF g_themeBackground = RGB(255, 255, 255);
inline thread_local COLORREF g_themeText       = RGB(0, 0, 0);
inline thread_local wil::unique_hbrush g_themeBrush;

inline std::atomic_bool g_defaultThemeEnabled{false};
inline std::atomic_bool g_defaultThemeUseDarkMode{false};
inline std::atomic_bool g_defaultThemeHighContrast{false};
inline std::atomic<DWORD> g_defaultThemeBackground{RGB(255, 255, 255)};
inline std::atomic<DWORD> g_defaultThemeText{RGB(0, 0, 0)};

inline void ApplyImmersiveDarkMode(HWND hwnd, bool enabled) noexcept
{
    if (! hwnd)
    {
        return;
    }

    using DwmSetWindowAttributeFunc          = HRESULT(WINAPI*)(HWND, DWORD, LPCVOID, DWORD);
    static DwmSetWindowAttributeFunc setAttr = []() noexcept -> DwmSetWindowAttributeFunc
    {
        HMODULE dwm = LoadLibrary(L"dwmapi.dll");
        if (! dwm)
        {
            Debug::ErrorWithLastError(L"Failed to load dwmapi.dll for ApplyImmersiveDarkMode.");
            return nullptr;
        }
#pragma warning(push)
#pragma warning(disable : 4191) // C4191: 'reinterpret_cast': unsafe conversion from 'FARPROC'
        return reinterpret_cast<DwmSetWindowAttributeFunc>(GetProcAddress(dwm, "DwmSetWindowAttribute"));
#pragma warning(pop)
    }();

    if (! setAttr)
    {
        return;
    }

    static constexpr DWORD kDwmwaUseImmersiveDarkMode19 = 19u;
    static constexpr DWORD kDwmwaUseImmersiveDarkMode20 = 20u;

    const BOOL darkMode = enabled ? TRUE : FALSE;
    setAttr(hwnd, kDwmwaUseImmersiveDarkMode20, &darkMode, sizeof(darkMode));
    setAttr(hwnd, kDwmwaUseImmersiveDarkMode19, &darkMode, sizeof(darkMode));
}

inline void ApplyWindowTheme(HWND hwnd, bool darkMode) noexcept
{
    if (! hwnd)
    {
        return;
    }

    using SetWindowThemeFunc           = HRESULT(WINAPI*)(HWND, LPCWSTR, LPCWSTR);
    static SetWindowThemeFunc setTheme = []() noexcept -> SetWindowThemeFunc
    {
        HMODULE uxTheme = LoadLibrary(L"uxtheme.dll");
        if (! uxTheme)
        {
            Debug::ErrorWithLastError(L"Failed to load uxtheme.dll for ApplyWindowTheme.");
            return nullptr;
        }
#pragma warning(push)
#pragma warning(disable : 4191) // C4191: 'reinterpret_cast': unsafe conversion from 'FARPROC'
        return reinterpret_cast<SetWindowThemeFunc>(GetProcAddress(uxTheme, "SetWindowTheme"));
#pragma warning(pop)
    }();

    if (! setTheme)
    {
        return;
    }

    setTheme(hwnd, darkMode ? L"DarkMode_Explorer" : L"Explorer", nullptr);
}

inline LRESULT OnThemedMessageBoxPaint(HWND hwnd) noexcept
{
    PAINTSTRUCT ps{};
    wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);
    if (hdc)
    {
        FillRect(hdc.get(), &ps.rcPaint, g_themeBrush.get());
    }
    return 0;
}

inline bool TryHandleThemedMessageBoxEraseBkgnd(HWND hwnd, HDC hdc) noexcept
{
    RECT client{};
    if (GetClientRect(hwnd, &client))
    {
        FillRect(hdc, &client, g_themeBrush.get());
        return true;
    }
    return false;
}

inline LRESULT OnThemedMessageBoxCtlColorText(HDC hdc) noexcept
{
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, g_themeText);
    return reinterpret_cast<LRESULT>(g_themeBrush.get());
}

inline LRESULT CALLBACK ThemedMessageBoxWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    if (g_themeEnabled && g_themeBrush)
    {
        switch (msg)
        {
            case WM_PAINT: return OnThemedMessageBoxPaint(hwnd);
            case WM_ERASEBKGND:
                if (TryHandleThemedMessageBoxEraseBkgnd(hwnd, reinterpret_cast<HDC>(wp)))
                {
                    return 1;
                }
                break;
            case WM_CTLCOLORDLG: return reinterpret_cast<LRESULT>(g_themeBrush.get());
            case WM_CTLCOLORSTATIC: return OnThemedMessageBoxCtlColorText(reinterpret_cast<HDC>(wp));
            case WM_CTLCOLORBTN: return OnThemedMessageBoxCtlColorText(reinterpret_cast<HDC>(wp));
        }
    }

    if (msg == WM_NCDESTROY && g_msgBoxWndProc)
    {
        WNDPROC original = g_msgBoxWndProc;
        g_msgBoxWndProc  = nullptr;
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
        SetWindowLongPtrW(hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(original));
        const LRESULT result = CallWindowProcW(original, hwnd, msg, wp, lp);
#pragma warning(pop)
        return result;
    }

    if (g_msgBoxWndProc)
    {
#pragma warning(push)
#pragma warning(disable : 5039) // C5039: passing potentially-throwing callback to extern "C" Win32 API under -EHc
        const LRESULT result = CallWindowProcW(g_msgBoxWndProc, hwnd, msg, wp, lp);
#pragma warning(pop)
        return result;
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

inline BOOL CALLBACK ApplyThemeToChildWindowsProc(HWND hwnd, LPARAM /*lp*/) noexcept
{
    ApplyWindowTheme(hwnd, g_themeUseDarkMode);
    SendMessageW(hwnd, WM_THEMECHANGED, 0, 0);
    return TRUE;
}

inline LRESULT CALLBACK CenteringHookProc(int nCode, WPARAM wParam, LPARAM lParam) noexcept
{
    if (nCode == HCBT_ACTIVATE && (g_centerOnWindow || g_themeEnabled))
    {
        HWND msgBox = reinterpret_cast<HWND>(wParam);

        if (g_themeEnabled && g_themeBrush)
        {
            ApplyImmersiveDarkMode(msgBox, g_themeUseDarkMode);
            ApplyWindowTheme(msgBox, g_themeUseDarkMode);
            EnumChildWindows(msgBox, ApplyThemeToChildWindowsProc, 0);
            SendMessageW(msgBox, WM_THEMECHANGED, 0, 0);

            if (! g_msgBoxWndProc)
            {
                g_msgBoxWndProc = reinterpret_cast<WNDPROC>(
                    SetWindowLongPtrW(msgBox, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(MessageBoxCenteringDetail::ThemedMessageBoxWndProc)));
            }
        }

        if (g_centerOnWindow)
        {
            HWND owner = g_centerOnWindow;

            RECT ownerRc{};
            RECT msgRc{};
            if (GetWindowRect(owner, &ownerRc) && GetWindowRect(msgBox, &msgRc))
            {
                const int ownerW = ownerRc.right - ownerRc.left;
                const int ownerH = ownerRc.bottom - ownerRc.top;
                const int msgW   = msgRc.right - msgRc.left;
                const int msgH   = msgRc.bottom - msgRc.top;

                const int x = ownerRc.left + (ownerW - msgW) / 2;
                const int y = ownerRc.top + (ownerH - msgH) / 2;

                SetWindowPos(msgBox, nullptr, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
            }
        }

        // Unhook after first activation
        if (g_hook)
        {
            UnhookWindowsHookEx(g_hook);
            g_hook = nullptr;
        }
        g_centerOnWindow = nullptr;
    }

    return CallNextHookEx(g_hook, nCode, wParam, lParam);
}
} // namespace MessageBoxCenteringDetail

struct MessageBoxTheme
{
    bool enabled        = false;
    bool useDarkMode    = false;
    bool highContrast   = false;
    COLORREF background = RGB(255, 255, 255);
    COLORREF text       = RGB(0, 0, 0);
};

inline void SetDefaultMessageBoxTheme(const MessageBoxTheme& theme) noexcept
{
    MessageBoxCenteringDetail::g_defaultThemeEnabled.store(theme.enabled, std::memory_order_relaxed);
    MessageBoxCenteringDetail::g_defaultThemeUseDarkMode.store(theme.useDarkMode, std::memory_order_relaxed);
    MessageBoxCenteringDetail::g_defaultThemeHighContrast.store(theme.highContrast, std::memory_order_relaxed);
    MessageBoxCenteringDetail::g_defaultThemeBackground.store(static_cast<DWORD>(theme.background), std::memory_order_relaxed);
    MessageBoxCenteringDetail::g_defaultThemeText.store(static_cast<DWORD>(theme.text), std::memory_order_relaxed);
}

inline void ClearDefaultMessageBoxTheme() noexcept
{
    MessageBoxTheme theme{};
    SetDefaultMessageBoxTheme(theme);
}

inline int MessageBoxThemedImpl(_In_opt_ HWND owner, PCWSTR text, PCWSTR caption, _In_ UINT type, bool centerOnOwner)
{
    const bool themeEnabled = MessageBoxCenteringDetail::g_defaultThemeEnabled.load(std::memory_order_relaxed) &&
                              ! MessageBoxCenteringDetail::g_defaultThemeHighContrast.load(std::memory_order_relaxed);

    if (centerOnOwner && owner && IsWindow(owner))
    {
        MessageBoxCenteringDetail::g_centerOnWindow = owner;
    }

    if (themeEnabled)
    {
        MessageBoxCenteringDetail::g_themeEnabled     = true;
        MessageBoxCenteringDetail::g_themeUseDarkMode = MessageBoxCenteringDetail::g_defaultThemeUseDarkMode.load(std::memory_order_relaxed);
        MessageBoxCenteringDetail::g_themeBackground =
            static_cast<COLORREF>(MessageBoxCenteringDetail::g_defaultThemeBackground.load(std::memory_order_relaxed));
        MessageBoxCenteringDetail::g_themeText = static_cast<COLORREF>(MessageBoxCenteringDetail::g_defaultThemeText.load(std::memory_order_relaxed));
        MessageBoxCenteringDetail::g_themeBrush.reset(CreateSolidBrush(MessageBoxCenteringDetail::g_themeBackground));
    }

    if ((MessageBoxCenteringDetail::g_centerOnWindow || themeEnabled) && ! MessageBoxCenteringDetail::g_hook)
    {
        MessageBoxCenteringDetail::g_hook = SetWindowsHookExW(WH_CBT, MessageBoxCenteringDetail::CenteringHookProc, nullptr, GetCurrentThreadId());
    }

    const int result = MessageBoxW(owner, text, caption, type);

    if (MessageBoxCenteringDetail::g_hook)
    {
        UnhookWindowsHookEx(MessageBoxCenteringDetail::g_hook);
        MessageBoxCenteringDetail::g_hook = nullptr;
    }

    MessageBoxCenteringDetail::g_centerOnWindow   = nullptr;
    MessageBoxCenteringDetail::g_msgBoxWndProc    = nullptr;
    MessageBoxCenteringDetail::g_themeEnabled     = false;
    MessageBoxCenteringDetail::g_themeUseDarkMode = false;
    MessageBoxCenteringDetail::g_themeBrush.reset();

    return result;
}

// MessageBox that is centered on the owner window
inline int MessageBoxCentered(_In_ HWND owner, _In_opt_ HINSTANCE hInstance, _In_ UINT textId, _In_ UINT captionId, _In_ UINT type)
{
    const std::wstring text    = LoadStringResource(hInstance, textId);
    const std::wstring caption = LoadStringResource(hInstance, captionId);
    return MessageBoxThemedImpl(owner, text.c_str(), caption.c_str(), type, true);
}

// MessageBox with caller-provided text, centered on the owner window.
inline int MessageBoxCenteredText(_In_opt_ HWND owner, const std::wstring& text, const std::wstring& caption, _In_ UINT type)
{
    return MessageBoxThemedImpl(owner, text.c_str(), caption.c_str(), type, true);
}

// class name for the Red Salamander Monitor window
constexpr auto g_redSalamanderMonitor          = L"Red Salamander Monitor";
constexpr auto g_redSalamanderMonitorClassName = L"RedSalamanderMonitor Window";

//////////////////////////////////////////////////////////////////////////////////
// DEBUG helpers
namespace Debug
{
struct TransportStats
{
    uint64_t etwWritten = 0;
    uint64_t etwFailed  = 0;
};

struct InfoParam // the real data/string will be after this structure
{
    enum Type : uint32_t
    {
        Text    = 0x0,
        Error   = 0x1,
        Warning = 0x2,
        Info    = 0x4,
        Debug   = 0x8,
        All     = 0x1F // Bitmask for all types enabled (bits 0-4)
    };
    FILETIME time; // More efficient: 8 bytes vs 16 bytes for SYSTEMTIME
    DWORD processID;
    DWORD threadID;
    Type type; // 0 - text, 1 - error, 2 - warning, 3 - info, 4 - debug

    // Helper method to get SYSTEMTIME when needed for display
    SYSTEMTIME GetLocalTime() const noexcept
    {
        SYSTEMTIME st{};
        FILETIME localFileTime{};
        if (FileTimeToLocalFileTime(&time, &localFileTime))
        {
            FileTimeToSystemTime(&localFileTime, &st);
        }
        return st;
    }

    // Helper method to get formatted time string
    std::wstring GetTimeString() const noexcept
    {
        SYSTEMTIME st = GetLocalTime();
        return std::format(L"{:02d}:{:02d}:{:02d}.{:03d}", st.wHour, st.wMinute, st.wSecond, st.wMilliseconds);
    }
};

// TraceLogging provider declaration
// Each module (EXE/DLL) must define its own provider instance using the same GUID
// to avoid cross-module provider handle issues. TraceLogging does NOT support
// sharing provider handles across DLL boundaries.
//
// Usage:
//   - In ONE .cpp file per module: #define REDSAL_DEFINE_TRACE_PROVIDER before including Helpers.h
//   - In all other files: Just include Helpers.h (uses TRACELOGGING_DECLARE_PROVIDER)

#if ! defined(REDSAL_DEFINE_TRACE_PROVIDER)
// Declaration only - provider will be defined elsewhere in this module
TRACELOGGING_DECLARE_PROVIDER(g_RedSalamanderProvider);
#else
// Definition - creates the actual provider storage in this module
TRACELOGGING_DEFINE_PROVIDER(g_RedSalamanderProvider,
                             "RedSalamanderMonitor",
                             // {440c70f6-6c6b-4ff7-9a3f-0b7db411b31a}
                             (0x440c70f6, 0x6c6b, 0x4ff7, 0x9a, 0x3f, 0x0b, 0x7d, 0xb4, 0x11, 0xb3, 0x1a));
#endif

namespace detail
{
inline std::atomic<uint64_t> g_etwWritten{0};
inline std::atomic<uint64_t> g_etwFailed{0};
inline std::once_flag g_traceLoggingRegisterOnce;
inline std::atomic<bool> g_etwRegistered{false};

struct IndentationState
{
    int level = 0;
    std::wstring prefix;
};

inline thread_local IndentationState g_indentation{};

inline void UpdateIndentationPrefix() noexcept
{
    constexpr int kMaxIndentLevel       = 64;
    constexpr int kIndentSpacesPerLevel = 2;
    const int boundedLevel              = std::clamp(g_indentation.level, 0, kMaxIndentLevel);

    if (boundedLevel <= 0)
    {
        g_indentation.prefix.clear();
        return;
    }

    const size_t spaceCount = static_cast<size_t>(kIndentSpacesPerLevel) * static_cast<size_t>(boundedLevel) - 1u;
    g_indentation.prefix.assign(spaceCount, L' ');
    g_indentation.prefix.append(L" - ");
}

inline void Indent() noexcept
{
    if (g_indentation.level < std::numeric_limits<int>::max())
    {
        ++g_indentation.level;
    }
    UpdateIndentationPrefix();
}

inline void Unindent() noexcept
{
    if (g_indentation.level > 0)
    {
        --g_indentation.level;
    }
    UpdateIndentationPrefix();
}

inline std::wstring_view GetIndentationPrefix() noexcept
{
    return g_indentation.prefix;
}

inline void PrependIndentation(std::wstring& message) noexcept
{
    const std::wstring_view prefix = GetIndentationPrefix();
    if (prefix.empty())
    {
        return;
    }

    const size_t newlineCount = static_cast<size_t>(std::count(message.begin(), message.end(), L'\n'));
    if (newlineCount == 0)
    {
        message.insert(0, prefix);
        return;
    }

    std::wstring indented;
    indented.reserve(message.size() + ((newlineCount + 1u) * prefix.size()));
    indented.append(prefix);
    for (size_t i = 0; i < message.size(); ++i)
    {
        const wchar_t ch = message[i];
        indented.push_back(ch);
        if (ch == L'\n' && (i + 1u) < message.size())
        {
            indented.append(prefix);
        }
    }

    message.swap(indented);
}

inline bool EnsureTraceLoggingRegistered() noexcept
{
    std::call_once(g_traceLoggingRegisterOnce,
                   []() noexcept
                   {
                       const HRESULT hr   = TraceLoggingRegister(g_RedSalamanderProvider);
                       const bool success = SUCCEEDED(hr);
                       g_etwRegistered.store(success, std::memory_order_release);

#ifdef _DEBUG
                       // Output detailed registration result for debugging
                       wchar_t msg[256]{};
                       const size_t msgMax = (sizeof(msg) / sizeof(msg[0])) - 1;
                       if (success)
                       {
                           const auto r           = std::format_to_n(msg, msgMax, L"ETW TraceLoggingRegister succeeded: 0x{:08X}\n", static_cast<unsigned>(hr));
                           using SizeType         = decltype(r.size);
                           const SizeType cap     = static_cast<SizeType>(msgMax);
                           const SizeType written = (r.size < 0) ? 0 : ((r.size > cap) ? cap : r.size);
                           msg[static_cast<size_t>(written)] = L'\0';
                       }
                       else
                       {
                           const wchar_t* const hrText = hr == static_cast<HRESULT>(0x80070005)   ? L"E_ACCESSDENIED"
                                                         : hr == E_INVALIDARG                     ? L"E_INVALIDARG"
                                                         : hr == static_cast<HRESULT>(0x800700B7) ? L"ERROR_ALREADY_EXISTS"
                                                                                                  : L"Unknown Error";

                           const auto r = std::format_to_n(msg, msgMax, L"ETW TraceLoggingRegister FAILED: 0x{:08X} ({})\n", static_cast<unsigned>(hr), hrText);
                           using SizeType                    = decltype(r.size);
                           const SizeType cap                = static_cast<SizeType>(msgMax);
                           const SizeType written            = (r.size < 0) ? 0 : ((r.size > cap) ? cap : r.size);
                           msg[static_cast<size_t>(written)] = L'\0';
                       }
                       OutputDebugStringW(msg);
#endif
                   });
    return g_etwRegistered.load(std::memory_order_acquire);
}

inline bool IsEtwRegistered() noexcept
{
    return g_etwRegistered.load(std::memory_order_acquire);
}

constexpr ULONGLONG kDebugKeyword = 0x0000000000000001ull;
constexpr ULONGLONG kPerfKeyword  = 0x0000000000000002ull;

inline bool IsEtwEnabled(ULONGLONG keyword) noexcept
{
    if (! EnsureTraceLoggingRegistered())
    {
        return false;
    }

    return TraceLoggingProviderEnabled(g_RedSalamanderProvider, TRACE_LEVEL_INFORMATION, keyword) != 0;
}

inline bool IsDebugEtwEnabled() noexcept
{
    return IsEtwEnabled(kDebugKeyword);
}

inline InfoParam BuildInfoParam(InfoParam::Type type) noexcept
{
    InfoParam dbg{};
    GetSystemTimeAsFileTime(&dbg.time);
    dbg.processID = GetCurrentProcessId();
    dbg.threadID  = GetCurrentThreadId();
    dbg.type      = type;
    return dbg;
}

inline bool EmitEtwEvent(const InfoParam& info, std::wstring_view message) noexcept
{
    if (! EnsureTraceLoggingRegistered())
    {
        g_etwFailed.fetch_add(1, std::memory_order_relaxed);
        return false;
    }

    if (TraceLoggingProviderEnabled(g_RedSalamanderProvider, TRACE_LEVEL_INFORMATION, kDebugKeyword) == 0)
    {
        return false;
    }

    ULARGE_INTEGER fileTime{};
    fileTime.LowPart  = info.time.dwLowDateTime;
    fileTime.HighPart = info.time.dwHighDateTime;
    USHORT length     = static_cast<USHORT>(std::min<size_t>(message.size(), std::numeric_limits<USHORT>::max()));

    // TraceLoggingWrite is a macro that doesn't return a value we can easily capture.
    // Once registration succeeds, write failures are extremely rare (only if provider disabled).
    // We count successful writes; failures would show as missing events in consumer.
    TraceLoggingWrite(g_RedSalamanderProvider,
                      "DebugMessage",
                      TraceLoggingLevel(TRACE_LEVEL_INFORMATION),
                      TraceLoggingKeyword(kDebugKeyword),
                      TraceLoggingUInt32(static_cast<UINT32>(info.type), "Type"),
                      TraceLoggingUInt32(info.processID, "ProcessId"),
                      TraceLoggingUInt32(info.threadID, "ThreadId"),
                      TraceLoggingUInt64(fileTime.QuadPart, "FileTime"),
                      TraceLoggingCountedWideString(message.data(), length, "Message"));

    g_etwWritten.fetch_add(1, std::memory_order_relaxed);
    return true;
}

inline void Publish(const InfoParam& dbg, std::wstring_view payload) noexcept
{
    EmitEtwEvent(dbg, payload);
}

inline void PublishString(std::wstring_view payload) noexcept
{
    const InfoParam dbg = BuildInfoParam(InfoParam::Type::Text);
    EmitEtwEvent(dbg, payload);
}
} // namespace detail

inline TransportStats GetTransportStats() noexcept
{
    TransportStats stats{};
    stats.etwWritten = detail::g_etwWritten.load(std::memory_order_relaxed);
    stats.etwFailed  = detail::g_etwFailed.load(std::memory_order_relaxed);
    return stats;
}

namespace Perf
{
inline bool IsEnabled() noexcept
{
    return detail::IsEtwEnabled(detail::kPerfKeyword);
}

inline void Emit(std::wstring_view name, std::wstring_view detail, uint64_t durationUs, uint64_t value0 = 0, uint64_t value1 = 0, HRESULT hr = S_OK) noexcept
{
    if (! IsEnabled())
    {
        return;
    }

    const USHORT nameLen   = static_cast<USHORT>(std::min<size_t>(name.size(), std::numeric_limits<USHORT>::max()));
    const USHORT detailLen = static_cast<USHORT>(std::min<size_t>(detail.size(), std::numeric_limits<USHORT>::max()));

    TraceLoggingWrite(g_RedSalamanderProvider,
                      "PerfScope",
                      TraceLoggingLevel(TRACE_LEVEL_INFORMATION),
                      TraceLoggingKeyword(detail::kPerfKeyword),
                      TraceLoggingCountedWideString(name.data() ? name.data() : L"", nameLen, "Name"),
                      TraceLoggingCountedWideString(detail.data() ? detail.data() : L"", detailLen, "Detail"),
                      TraceLoggingUInt64(durationUs, "DurationUs"),
                      TraceLoggingUInt64(value0, "Value0"),
                      TraceLoggingUInt64(value1, "Value1"),
                      TraceLoggingUInt32(static_cast<uint32_t>(hr), "Hr"));
}

class Scope final
{
public:
    explicit Scope(std::wstring_view name) noexcept
        : _enabled(IsEnabled()),
          _name(name),
          _start(_enabled ? std::chrono::steady_clock::now() : std::chrono::steady_clock::time_point{})
    {
    }

    Scope(const Scope&)            = delete;
    Scope& operator=(const Scope&) = delete;
    Scope(Scope&&)                 = delete;
    Scope& operator=(Scope&&)      = delete;

    ~Scope() noexcept
    {
        if (! _enabled)
        {
            return;
        }

        const auto elapsed        = std::chrono::steady_clock::now() - _start;
        const uint64_t durationUs = static_cast<uint64_t>(std::chrono::duration_cast<std::chrono::microseconds>(elapsed).count());

        const USHORT nameLen     = static_cast<USHORT>(std::min<size_t>(_name.size(), std::numeric_limits<USHORT>::max()));
        const USHORT detailLen   = static_cast<USHORT>(std::min<size_t>(_detail.size(), std::numeric_limits<USHORT>::max()));
        const wchar_t* namePtr   = _name.data() ? _name.data() : L"";
        const wchar_t* detailPtr = _detail.data() ? _detail.data() : L"";

        TraceLoggingWrite(g_RedSalamanderProvider,
                          "PerfScope",
                          TraceLoggingLevel(TRACE_LEVEL_INFORMATION),
                          TraceLoggingKeyword(detail::kPerfKeyword),
                          TraceLoggingCountedWideString(namePtr, nameLen, "Name"),
                          TraceLoggingCountedWideString(detailPtr, detailLen, "Detail"),
                          TraceLoggingUInt64(durationUs, "DurationUs"),
                          TraceLoggingUInt64(_value0, "Value0"),
                          TraceLoggingUInt64(_value1, "Value1"),
                          TraceLoggingUInt32(_hr, "Hr"));
    }

    void SetDetail(std::wstring_view detail) noexcept
    {
        _detail = detail;
    }

    void SetValue0(uint64_t value) noexcept
    {
        _value0 = value;
    }

    void SetValue1(uint64_t value) noexcept
    {
        _value1 = value;
    }

    void SetHr(HRESULT hr) noexcept
    {
        _hr = static_cast<uint32_t>(hr);
    }

private:
    bool _enabled = false;
    std::wstring_view _name;
    std::wstring_view _detail;
    std::chrono::steady_clock::time_point _start;
    uint64_t _value0 = 0;
    uint64_t _value1 = 0;
    uint32_t _hr     = static_cast<uint32_t>(S_OK);
};
} // namespace Perf

inline void Out(PCWSTR p) noexcept
{
    if (! p)
    {
        return;
    }

    if (! detail::IsDebugEtwEnabled())
    {
        return;
    }

    const std::wstring_view prefix = detail::GetIndentationPrefix();
    if (prefix.empty())
    {
        detail::PublishString(p);
        return;
    }

    std::wstring message{p};
    detail::PrependIndentation(message);
    detail::PublishString(message);
}

template <typename... Args> inline void Out(InfoParam::Type type, std::wformat_string<Args...> format, Args&&... args) noexcept
{
    if (! detail::IsDebugEtwEnabled())
    {
        return;
    }

    // Mandatory: noexcept boundary. Formatting can throw; keep debug output best-effort.
    try
    {
        std::wstring formattedString = std::vformat(format.get(), std::make_wformat_args(args...));
        const InfoParam dbg          = detail::BuildInfoParam(type);

        detail::PrependIndentation(formattedString);

#ifdef _DEBUG
        if (type & InfoParam::Type::Error)
        {
            OutputDebugStringW(formattedString.c_str());
        }
#endif

        detail::Publish(dbg, formattedString);
    }
    catch (const std::bad_alloc&)
    {
        // Out-of-memory is treated as fatal. Fail-fast so the crash pipeline can capture a dump.
        std::terminate();
    }
    catch (const std::format_error&)
    {
        // Fallback for format string / argument mismatches.
        Debug::Out(L"[Formatting Error in DbgOut]");
    }
    catch (const std::exception&)
    {
        // Fallback for unexpected failures inside debug output.
        Debug::Out(L"[Unexpected Error in DbgOut]");
    }
}

// returns the last error code
template <typename... Args> inline DWORD LastError(InfoParam::Type type, std::wformat_string<Args...> format, Args&&... args) noexcept
{
    const DWORD lastError = ::GetLastError();

    if (! detail::IsDebugEtwEnabled())
    {
        return lastError;
    }

    // Mandatory: noexcept boundary. Formatting can throw; keep debug output best-effort.
    try
    {
        std::wstring formattedString = std::vformat(format.get(), std::make_wformat_args(args...));
        if (lastError == 0)
        {
            formattedString.append(L" --> (NO ERROR)");
            Debug::Out(type, L"{}", formattedString);
            return 0;
        }

        wil::unique_hlocal_string message;
        const DWORD result = ::FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
                                             nullptr,
                                             lastError,
                                             MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
                                             reinterpret_cast<PWSTR>(message.addressof()),
                                             0,
                                             nullptr);

        if (result > 0 && message)
        {
            // Remove trailing newlines from system message
            std::wstring_view messageView{message.get()};
            while (! messageView.empty() && (messageView.back() == L'\r' || messageView.back() == L'\n'))
            {
                messageView.remove_suffix(1);
            }

            formattedString.append(std::format(L" --> ({}) {}", lastError, messageView));
        }
        else
        {
            formattedString.append(std::format(L" --> ({}) Unknown error", lastError));
        }

        Debug::Out(type, L"{}", formattedString);
        return lastError;
    }
    catch (const std::bad_alloc&)
    {
        // Out-of-memory is treated as fatal. Fail-fast so the crash pipeline can capture a dump.
        std::terminate();
    }
    catch (const std::format_error&)
    {
        // Best-effort: avoid throwing from this noexcept boundary.
        Debug::Out(type, L"[Formatting Error in Debug::OutLastError] LastError: {}", lastError);
        return lastError;
    }
    catch (const std::exception&)
    {
        // Best-effort: avoid throwing from this noexcept boundary.
        Debug::Out(type, L"[Unexpected Error in Debug::OutLastError] LastError: {}", lastError);
        return lastError;
    }
}

// Additional utility functions for common debug scenarios

template <typename... Args> inline void Info(std::wformat_string<Args...> format, Args&&... args) noexcept
{
    Debug::Out(InfoParam::Type::Info, format, std::forward<Args>(args)...);
}
inline void Info(const std::wstring& message) noexcept
{
    Debug::Out(InfoParam::Type::Info, L"{}", message);
}

template <typename... Args> inline void Warning(std::wformat_string<Args...> format, Args&&... args) noexcept
{
    Debug::Out(InfoParam::Type::Warning, format, std::forward<Args>(args)...);
}
inline void Warning(const std::wstring& message) noexcept
{
    Debug::Out(InfoParam::Type::Warning, L"{}", message);
}

template <typename... Args> inline void Error(std::wformat_string<Args...> format, Args&&... args) noexcept
{
    Debug::Out(InfoParam::Type::Error, format, std::forward<Args>(args)...);
}
inline void Error(const std::wstring& message) noexcept
{
    Debug::Out(InfoParam::Type::Error, L"{}", message);
}

template <typename... Args> inline DWORD ErrorWithLastError(std::wformat_string<Args...> format, Args&&... args) noexcept
{
    return Debug::LastError(InfoParam::Type::Error, format, std::forward<Args>(args)...);
}
inline DWORD ErrorWithLastError(const std::wstring& message) noexcept
{
    return Debug::LastError(InfoParam::Type::Error, L"{}", message);
}

} // namespace Debug

// ============================================================================
// Module Lifetime Helpers
// ============================================================================
/// Returns an owning module handle for the module that contains `address`.
/// This increments the module reference count so the module cannot be unloaded while the returned handle is alive.
/// Returns an empty handle on failure.
[[nodiscard]] inline wil::unique_hmodule AcquireModuleReferenceFromAddress(const void* address) noexcept
{
    if (! address)
    {
        return {};
    }

    HMODULE module = nullptr;
    const BOOL ok  = GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, reinterpret_cast<LPCWSTR>(address), &module);
    if (ok == 0 || module == nullptr)
    {
        return {};
    }

    wil::unique_hmodule owned;
    owned.reset(module);
    return owned;
}

// ============================================================================
// PostMessage Payload RAII Helpers
// ============================================================================
// These helpers provide safe ownership transfer for payloads sent via PostMessageW/SendMessageW.
// They eliminate raw new/delete by using std::unique_ptr for automatic cleanup.
//
// Usage pattern:
//   Sender:
//     auto payload = std::make_unique<MyPayload>();
//     // ... fill payload ...
//     if (!PostMessagePayload(hwnd, WM_MYMSG, 0, std::move(payload))) { /* handle error */ }
//
//   Receiver (WndProc):
//     auto payload = TakeMessagePayload<MyPayload>(lParam);
//     // NOTE: Receiver MUST use TakeMessagePayload<T> (not `std::unique_ptr<T>(reinterpret_cast<T*>(lParam))`)
//     // so the payload registry can unregister it and avoid double-delete during WM_NCDESTROY draining.
//     // ... use payload ...
//     // payload automatically deleted when scope exits
//
// Window teardown:
// - If an `HWND` is destroyed while messages are still queued, Windows may discard those messages without delivering them.
//   If those messages carry heap payload pointers, the payloads become unreachable (leak).
// - To prevent that, windows that receive payload messages should call `DrainPostedPayloadsForWindow(hwnd)` in `WM_NCDESTROY`
//   and call `InitPostedPayloadWindow(hwnd)` during create (`WM_NCCREATE`/`WM_CREATE`) to handle potential HWND reuse.

namespace detail
{
using MessagePayloadDeleter = void (*)(void*) noexcept;

struct PostedMessagePayloadEntry final
{
    HWND hwnd                 = nullptr;
    UINT msg                  = 0;
    MessagePayloadDeleter del = nullptr;
};

struct PostedMessagePayloadRegistry final
{
    PostedMessagePayloadRegistry()                                               = default;
    PostedMessagePayloadRegistry(const PostedMessagePayloadRegistry&)            = delete;
    PostedMessagePayloadRegistry& operator=(const PostedMessagePayloadRegistry&) = delete;
    PostedMessagePayloadRegistry(PostedMessagePayloadRegistry&&)                 = delete;
    PostedMessagePayloadRegistry& operator=(PostedMessagePayloadRegistry&&)      = delete;

    std::mutex mutex;
    std::unordered_map<void*, PostedMessagePayloadEntry> entriesByPtr;
    std::unordered_map<HWND, std::unordered_set<void*>> ptrsByHwnd;
    std::unordered_set<HWND> closedHwnds;
};

[[nodiscard]] inline PostedMessagePayloadRegistry& GetPostedMessagePayloadRegistry() noexcept
{
    static PostedMessagePayloadRegistry registry;
    return registry;
}

inline void InitPostedPayloadWindow(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    auto& registry = GetPostedMessagePayloadRegistry();
    std::lock_guard lock(registry.mutex);
    static_cast<void>(registry.closedHwnds.erase(hwnd));
}

[[nodiscard]] inline bool RegisterPostedMessagePayload(HWND hwnd, UINT msg, void* payload, MessagePayloadDeleter deleter) noexcept
{
    if (! hwnd || ! payload || deleter == nullptr)
    {
        return false;
    }

    bool shouldDelete = false;
    {
        auto& registry = GetPostedMessagePayloadRegistry();
        std::lock_guard lock(registry.mutex);

        if (registry.closedHwnds.contains(hwnd))
        {
            shouldDelete = true;
        }
        else
        {
            static_cast<void>(registry.entriesByPtr.emplace(payload, PostedMessagePayloadEntry{.hwnd = hwnd, .msg = msg, .del = deleter}));
            static_cast<void>(registry.ptrsByHwnd[hwnd].insert(payload));
        }
    }

    if (shouldDelete)
    {
        deleter(payload);
        return false;
    }

    return true;
}

inline void UnregisterPostedMessagePayload(void* payload) noexcept
{
    if (! payload)
    {
        return;
    }

    auto& registry = GetPostedMessagePayloadRegistry();
    std::lock_guard lock(registry.mutex);

    const auto it = registry.entriesByPtr.find(payload);
    if (it == registry.entriesByPtr.end())
    {
        return;
    }

    const HWND hwnd = it->second.hwnd;
    registry.entriesByPtr.erase(it);

    const auto hwIt = registry.ptrsByHwnd.find(hwnd);
    if (hwIt != registry.ptrsByHwnd.end())
    {
        hwIt->second.erase(payload);
        if (hwIt->second.empty())
        {
            registry.ptrsByHwnd.erase(hwIt);
        }
    }
}
} // namespace detail

// Call during window creation (WM_NCCREATE/WM_CREATE) for any window that can receive payload messages.
// This clears any previous drained state in case the HWND value is reused.
inline void InitPostedPayloadWindow(HWND hwnd) noexcept
{
    detail::InitPostedPayloadWindow(hwnd);
}

[[nodiscard]] inline size_t DrainPostedPayloadsForWindow(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return 0;
    }

    std::vector<std::pair<void*, detail::MessagePayloadDeleter>> toDelete;
    {
        auto& registry = detail::GetPostedMessagePayloadRegistry();
        std::lock_guard lock(registry.mutex);

        static_cast<void>(registry.closedHwnds.insert(hwnd));

        const auto hwIt = registry.ptrsByHwnd.find(hwnd);
        if (hwIt == registry.ptrsByHwnd.end())
        {
            return 0;
        }

        toDelete.reserve(hwIt->second.size());

        for (void* payload : hwIt->second)
        {
            const auto it = registry.entriesByPtr.find(payload);
            if (it == registry.entriesByPtr.end())
            {
                continue;
            }

            toDelete.emplace_back(payload, it->second.del);
            registry.entriesByPtr.erase(it);
        }

        registry.ptrsByHwnd.erase(hwIt);
    }

    for (const auto& [payload, deleter] : toDelete)
    {
        if (deleter)
        {
            deleter(payload);
        }
    }

    return toDelete.size();
}

/// Posts a message with a unique_ptr payload. If PostMessageW fails, the payload is automatically deleted.
/// Returns true on success, false on failure (payload is deleted, call GetLastError() for details).
template <typename T> [[nodiscard]] inline bool PostMessagePayload(HWND hwnd, UINT msg, WPARAM wParam, std::unique_ptr<T> payload) noexcept
{
    T* raw = payload.release();
    if (raw == nullptr)
    {
        return PostMessageW(hwnd, msg, wParam, 0) != 0;
    }

    const auto deleter = +[](void* ptr) noexcept { delete static_cast<T*>(ptr); };

    auto& registry = detail::GetPostedMessagePayloadRegistry();
    std::unique_lock lock(registry.mutex);

    if (! hwnd || registry.closedHwnds.contains(hwnd))
    {
        lock.unlock();
        deleter(raw);
        return false;
    }

    static_cast<void>(registry.entriesByPtr.emplace(raw, detail::PostedMessagePayloadEntry{.hwnd = hwnd, .msg = msg, .del = deleter}));
    static_cast<void>(registry.ptrsByHwnd[hwnd].insert(raw));

    if (! PostMessageW(hwnd, msg, wParam, reinterpret_cast<LPARAM>(raw)))
    {
        registry.entriesByPtr.erase(raw);

        const auto hwIt = registry.ptrsByHwnd.find(hwnd);
        if (hwIt != registry.ptrsByHwnd.end())
        {
            hwIt->second.erase(raw);
            if (hwIt->second.empty())
            {
                registry.ptrsByHwnd.erase(hwIt);
            }
        }

        lock.unlock();
        deleter(raw);
        return false;
    }

    return true;
}

/// Takes ownership of a message payload from LPARAM, wrapping it in a unique_ptr.
/// Use in WndProc message handlers to ensure automatic cleanup.
template <typename T> [[nodiscard]] inline std::unique_ptr<T> TakeMessagePayload(LPARAM lParam) noexcept
{
    detail::UnregisterPostedMessagePayload(reinterpret_cast<void*>(lParam));
    return std::unique_ptr<T>(reinterpret_cast<T*>(lParam));
}

// Macro Helpers for debug output
#ifdef _DEBUG
#define DBGOUT_INFO Debug::Info
#define DBGOUT_WARNING Debug::Warning
#define DBGOUT_ERROR Debug::Error
#define DBGOUT_ERROR_LASTERROR Debug::ErrorWithLastError
#else
#define DBGOUT_INFO __noop
#define DBGOUT_WARNING __noop
#define DBGOUT_ERROR __noop
#define DBGOUT_ERROR_LASTERROR __noop
#endif

// CallTracer: hierarchical indentation + performance measurement for Debug::Out messages on the current thread.
//
// Default behavior:
// - TRACER / TRACER_CTX: only logs the Exiting message (indentation still applies to all nested logs)
// - TRACER_INOUT / TRACER_INOUT_CTX: logs both Entering and Exiting messages
//
// Indentation is shared with Debug::Info/Warning/Error/Out on the same thread.

class CallTracer
{
public:
    enum class Mode : uint8_t
    {
        ExitOnly,
        EnterExit,
    };

    explicit CallTracer(const wchar_t* functionName, Mode mode = Mode::ExitOnly) noexcept : CallTracer(functionName, nullptr, mode)
    {
    }

    CallTracer(const wchar_t* functionName, const wchar_t* context, Mode mode = Mode::ExitOnly) noexcept
        : _enabled(Debug::detail::IsDebugEtwEnabled()),
          _functionName(functionName),
          _context(context),
          _mode(mode)
    {
        if (! _enabled)
        {
            return;
        }

        if (_mode == Mode::EnterExit)
        {
            if (_context)
            {
                Debug::Info(L"{} ({}) Entering", _functionName, _context);
            }
            else
            {
                Debug::Info(L"{} Entering", _functionName);
            }
        }

        Debug::detail::Indent();
        QueryPerformanceCounter(&_start);
    }

    ~CallTracer() noexcept
    {
        if (! _enabled)
        {
            return;
        }

        LARGE_INTEGER elapse{};
        QueryPerformanceCounter(&elapse);
        const double frequency = GetQpcFrequency();
        const double elapsedMs = (frequency > 0.0) ? (static_cast<double>(elapse.QuadPart - _start.QuadPart) * 1000.0 / frequency) : 0.0;

        Debug::detail::Unindent();

        if (_context)
        {
            Debug::Info(L"{} ({}) Exiting ({:.3f}ms)", _functionName, _context, elapsedMs);
        }
        else
        {
            Debug::Info(L"{} Exiting ({:.3f}ms)", _functionName, elapsedMs);
        }
    }

private:
    static double GetQpcFrequency() noexcept
    {
        static const double cached = []() noexcept
        {
            LARGE_INTEGER freq{};
            if (! QueryPerformanceFrequency(&freq) || freq.QuadPart <= 0)
            {
                return 0.0;
            }
            return static_cast<double>(freq.QuadPart);
        }();
        return cached;
    }

    bool _enabled                = false;
    const wchar_t* _functionName = nullptr;
    const wchar_t* _context      = nullptr;
    Mode _mode                   = Mode::ExitOnly;
    LARGE_INTEGER _start{};
};

// Helper macros for proper token concatenation
#define REDSAL_TRACER_CONCAT_IMPL(a, b) a##b
#define REDSAL_TRACER_CONCAT(a, b) REDSAL_TRACER_CONCAT_IMPL(a, b)

// Convert __FUNCTION__ (narrow string) to wide string at compile time
#define REDSAL_TRACER_WIDEN_IMPL(x) L##x
#define REDSAL_TRACER_WIDEN(x) REDSAL_TRACER_WIDEN_IMPL(x)

// Main tracing macros - use __FUNCTIONW__ (wide version) or convert __FUNCTION__
#if defined(__FUNCTIONW__)
#define TRACER [[maybe_unused]] CallTracer REDSAL_TRACER_CONCAT(_tracer_, __COUNTER__)(__FUNCTIONW__, CallTracer::Mode::ExitOnly)
#define TRACER_CTX(ctx) [[maybe_unused]] CallTracer REDSAL_TRACER_CONCAT(_tracer_, __COUNTER__)(__FUNCTIONW__, ctx, CallTracer::Mode::ExitOnly)
#define TRACER_INOUT [[maybe_unused]] CallTracer REDSAL_TRACER_CONCAT(_tracer_, __COUNTER__)(__FUNCTIONW__, CallTracer::Mode::EnterExit)
#define TRACER_INOUT_CTX(ctx) [[maybe_unused]] CallTracer REDSAL_TRACER_CONCAT(_tracer_, __COUNTER__)(__FUNCTIONW__, ctx, CallTracer::Mode::EnterExit)
#else
#define TRACER [[maybe_unused]] CallTracer REDSAL_TRACER_CONCAT(_tracer_, __COUNTER__)(REDSAL_TRACER_WIDEN(__FUNCTION__), CallTracer::Mode::ExitOnly)
#define TRACER_CTX(ctx)                                                                                                                                        \
    [[maybe_unused]] CallTracer REDSAL_TRACER_CONCAT(_tracer_, __COUNTER__)(REDSAL_TRACER_WIDEN(__FUNCTION__), ctx, CallTracer::Mode::ExitOnly)
#define TRACER_INOUT [[maybe_unused]] CallTracer REDSAL_TRACER_CONCAT(_tracer_, __COUNTER__)(REDSAL_TRACER_WIDEN(__FUNCTION__), CallTracer::Mode::EnterExit)
#define TRACER_INOUT_CTX(ctx)                                                                                                                                  \
    [[maybe_unused]] CallTracer REDSAL_TRACER_CONCAT(_tracer_, __COUNTER__)(REDSAL_TRACER_WIDEN(__FUNCTION__), ctx, CallTracer::Mode::EnterExit)
#endif

#define TRACER_CTW(ctx) TRACER_CTX(ctx)
#define TRACER_INOUT_CTW(ctx) TRACER_INOUT_CTX(ctx)

#pragma warning(pop)
