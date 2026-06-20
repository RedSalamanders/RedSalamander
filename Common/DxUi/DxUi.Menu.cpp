#include "DxUi.Internal.h"

#include <algorithm>
#include <chrono>
#include <cmath>
#include <d2d1effects.h>
#include <exception>
#include <format>
#include <limits>
#include <mutex>
#include <new>
#include <shellscalingapi.h>
#include <system_error>
#include <utility>
#include <windowsx.h>
#include <wincodec.h>

#include "Helpers.h"
#include "WindowMessages.h"

#pragma comment(lib, "shcore.lib")

#ifndef CLSID_WICImagingFactory2
#define CLSID_WICImagingFactory2 CLSID_WICImagingFactory
#endif

#if defined(ENABLE_TESTS)
#define DXUI_MENU_TRACE(...) Debug::Info(__VA_ARGS__)
#else
#define DXUI_MENU_TRACE(...) static_cast<void>(0)
#endif

#if defined(ENABLE_DXUI_MENU_DIAGNOSTICS) || (defined(ENABLE_TESTS) && ! defined(NDEBUG))
#define DXUI_MENU_PERSISTENT_DIAGNOSTICS 1
#else
#define DXUI_MENU_PERSISTENT_DIAGNOSTICS 0
#endif

// Persistent menu file tracing is compiled out of retail builds. To reactivate
// it for a support build, define ENABLE_DXUI_MENU_DIAGNOSTICS at compile time
// and set REDSALAMANDER_DXUI_MENU_TRACE=1 at runtime. Debug/test builds also
// keep it available while NDEBUG is not defined.
#if DXUI_MENU_PERSISTENT_DIAGNOSTICS
#define DXUI_MENU_DIAGNOSTICS_TRACE(...) WriteMenuDiagnosticsTrace(__VA_ARGS__)
#else
#define DXUI_MENU_DIAGNOSTICS_TRACE(...) static_cast<void>(0)
#endif

namespace RedSalamander::DxUi
{
namespace
{
// ---------------------------------------------------------------------------
// Menu layout constants (WinUI spec §3.2)
// ---------------------------------------------------------------------------

constexpr float kMenuCornerRadiusDip       = kPopupRoundSmallCornerRadiusDip;
constexpr float kMenuBorderDip             = 1.0f;
constexpr float kMenuPaddingTopDip         = 4.0f;
constexpr float kMenuPaddingBottomDip      = 4.0f;
constexpr float kMenuMinWidthDip           = 128.0f;
constexpr float kMenuMaxWidthDip           = 456.0f;
constexpr float kMenuShadowLeftDip         = 10.0f;
constexpr float kMenuShadowTopDip          = 8.0f;
constexpr float kMenuShadowRightDip        = 10.0f;
constexpr float kMenuShadowBottomDip       = 14.0f;
constexpr float kIconAreaWidthDip          = 28.0f;
constexpr float kTextLeftPaddingDip        = 36.0f;
constexpr float kTextToAccelGapDip         = 16.0f;
constexpr float kAccelRightPaddingDip      = 16.0f;
constexpr float kChevronRightPaddingDip    = 12.0f;
constexpr float kChevronAreaWidthDip       = 24.0f;
constexpr float kItemHoverRadiusDip        = 4.0f;
constexpr float kItemHoverInsetDip         = 4.0f;
constexpr float kIconSlotLeftInsetDip      = kMenuBorderDip + kItemHoverInsetDip + 2.0f;
constexpr float kIconSlotWidthDip          = kIconAreaWidthDip - kItemHoverInsetDip - 2.0f;
constexpr float kSeparatorHeightDip        = 9.0f; // 1px line + 4 DIP margin above/below
constexpr float kSeparatorMarginDip        = 4.0f;
constexpr float kSliderItemHeightDip       = 56.0f;
constexpr float kSliderTrackInsetDip       = 16.0f;
constexpr float kSliderTrackTopDip         = 36.0f;
constexpr float kSliderTrackHeightDip      = 4.0f;
constexpr float kSliderStopRadiusDip       = 5.0f;
constexpr float kSliderActiveStopRadiusDip = 6.0f;
constexpr float kSliderMenuMinWidthDip    = 260.0f;
constexpr float kSubmenuVerticalOffsetDip = 4.0f;
constexpr float kCascadeHoverDelayMs      = 400;
constexpr float kTextMeasureWidthDip      = 1024.0f;
#if DXUI_MENU_PERSISTENT_DIAGNOSTICS
constexpr uint64_t kMenuLoopTraceRepeatFlushCount     = 1024u;
#endif

constexpr wchar_t kChevronRightGlyph[]   = L"\uE76C";
constexpr wchar_t kCheckMarkGlyph[]      = L"\uE73E";
constexpr wchar_t kRadioBulletGlyph[]    = L"\uF137"; // Filled circle
constexpr GUID kMenuGaussianBlurEffectId = {0x1feb6d69, 0x2fe6, 0x4ac9, {0x8c, 0x58, 0x1d, 0x7f, 0x93, 0xe7, 0xa6, 0xa5}};

struct MenuController;
struct MenuPopup;

[[nodiscard]] const wchar_t* TraceMenuMessageName(UINT message) noexcept
{
    switch (message)
    {
        case WM_MOUSEMOVE: return L"WM_MOUSEMOVE";
        case WM_LBUTTONDOWN: return L"WM_LBUTTONDOWN";
        case WM_LBUTTONUP: return L"WM_LBUTTONUP";
        case WM_RBUTTONDOWN: return L"WM_RBUTTONDOWN";
        case WM_RBUTTONUP: return L"WM_RBUTTONUP";
        case WM_MBUTTONDOWN: return L"WM_MBUTTONDOWN";
        case WM_MBUTTONUP: return L"WM_MBUTTONUP";
        case WM_XBUTTONDOWN: return L"WM_XBUTTONDOWN";
        case WM_XBUTTONUP: return L"WM_XBUTTONUP";
        case WM_MOUSEWHEEL: return L"WM_MOUSEWHEEL";
        case WM_MOUSEHWHEEL: return L"WM_MOUSEHWHEEL";
        case WM_NCMOUSEMOVE: return L"WM_NCMOUSEMOVE";
        case WM_NCLBUTTONDOWN: return L"WM_NCLBUTTONDOWN";
        case WM_NCLBUTTONUP: return L"WM_NCLBUTTONUP";
        case WM_NCRBUTTONDOWN: return L"WM_NCRBUTTONDOWN";
        case WM_NCRBUTTONUP: return L"WM_NCRBUTTONUP";
        case WM_MOUSELEAVE: return L"WM_MOUSELEAVE";
        case WM_NCMOUSELEAVE: return L"WM_NCMOUSELEAVE";
        case WM_NULL: return L"WM_NULL";
        case WM_PAINT: return L"WM_PAINT";
        case WM_ERASEBKGND: return L"WM_ERASEBKGND";
        case WM_SETCURSOR: return L"WM_SETCURSOR";
        case WM_TIMER: return L"WM_TIMER";
        case WM_CAPTURECHANGED: return L"WM_CAPTURECHANGED";
        case WM_CANCELMODE: return L"WM_CANCELMODE";
        case WM_ACTIVATEAPP: return L"WM_ACTIVATEAPP";
        case WM_ACTIVATE: return L"WM_ACTIVATE";
        case WM_NCACTIVATE: return L"WM_NCACTIVATE";
        case WM_SETFOCUS: return L"WM_SETFOCUS";
        case WM_KILLFOCUS: return L"WM_KILLFOCUS";
        case WM_DESTROY: return L"WM_DESTROY";
        case WM_KEYDOWN: return L"WM_KEYDOWN";
        case WM_KEYUP: return L"WM_KEYUP";
        case WM_SYSKEYDOWN: return L"WM_SYSKEYDOWN";
        case WM_SYSKEYUP: return L"WM_SYSKEYUP";
        default: return L"message";
    }
}

[[nodiscard]] int TracePopupIndex(const MenuController& controller, const MenuPopup* popup) noexcept;

[[nodiscard]] uintptr_t TraceHwndValue(HWND hwnd) noexcept
{
    return reinterpret_cast<uintptr_t>(hwnd);
}

#if DXUI_MENU_PERSISTENT_DIAGNOSTICS
[[nodiscard]] const wchar_t* TraceBool(bool value) noexcept
{
    return value ? L"true" : L"false";
}
#endif

[[nodiscard]] bool IsMenuClientMouseMessage(UINT message) noexcept
{
    return message >= WM_MOUSEFIRST && message <= WM_MOUSELAST;
}

#if DXUI_MENU_PERSISTENT_DIAGNOSTICS
[[nodiscard]] bool IsMenuNonClientMouseMessage(UINT message) noexcept
{
    return message >= WM_NCMOUSEMOVE && message <= WM_NCXBUTTONDBLCLK;
}

[[nodiscard]] bool IsMenuKeyMessage(UINT message) noexcept
{
    return message >= WM_KEYFIRST && message <= WM_KEYLAST;
}

[[nodiscard]] bool IsMenuPriorityInputMessage(UINT message) noexcept
{
    return IsMenuClientMouseMessage(message) || IsMenuNonClientMouseMessage(message) || IsMenuKeyMessage(message);
}
#endif

[[nodiscard]] bool PeekMenuPriorityMessage(MSG& msg) noexcept
{
    if (PeekMessageW(&msg, nullptr, WM_MOUSEFIRST, WM_MOUSELAST, PM_REMOVE) != FALSE)
    {
        return true;
    }
    if (PeekMessageW(&msg, nullptr, WM_NCMOUSEMOVE, WM_NCXBUTTONDBLCLK, PM_REMOVE) != FALSE)
    {
        return true;
    }
    if (PeekMessageW(&msg, nullptr, WM_KEYFIRST, WM_KEYLAST, PM_REMOVE) != FALSE)
    {
        return true;
    }
    return false;
}

#if DXUI_MENU_PERSISTENT_DIAGNOSTICS
struct MenuDiagnosticsTraceState
{
    MenuDiagnosticsTraceState()                                         = default;
    MenuDiagnosticsTraceState(const MenuDiagnosticsTraceState&)         = delete;
    MenuDiagnosticsTraceState& operator=(const MenuDiagnosticsTraceState&) = delete;
    MenuDiagnosticsTraceState(MenuDiagnosticsTraceState&&)              = delete;
    MenuDiagnosticsTraceState& operator=(MenuDiagnosticsTraceState&&)   = delete;

    std::once_flag initOnce;
    std::atomic<bool> enabled = false;
    wil::unique_hfile file;
    std::mutex mutex;
    std::wstring path;
    std::atomic<uint64_t> sequence = 0u;
};

[[nodiscard]] MenuDiagnosticsTraceState& GetMenuDiagnosticsTraceState() noexcept
{
    static MenuDiagnosticsTraceState state;
    return state;
}

[[nodiscard]] bool MenuDiagnosticsEnvEquals(std::wstring_view value, std::wstring_view expected) noexcept
{
    if (value.size() != expected.size() || value.size() > static_cast<size_t>((std::numeric_limits<int>::max)()))
    {
        return false;
    }
    return CompareStringOrdinal(value.data(), static_cast<int>(value.size()), expected.data(), static_cast<int>(expected.size()), TRUE) == CSTR_EQUAL;
}

[[nodiscard]] std::wstring ReadMenuDiagnosticsEnvironmentVariable(const wchar_t* name)
{
    const DWORD required = GetEnvironmentVariableW(name, nullptr, 0u);
    if (required == 0u)
    {
        return {};
    }

    std::wstring value(required, L'\0');
    const DWORD copied = GetEnvironmentVariableW(name, value.data(), required);
    if (copied == 0u || copied >= required)
    {
        return {};
    }

    value.resize(copied);
    return value;
}

[[nodiscard]] bool IsMenuDiagnosticsTraceRequested(std::wstring_view value) noexcept
{
    if (value.empty())
    {
        return false;
    }
    if (value == L"0" || MenuDiagnosticsEnvEquals(value, L"false") || MenuDiagnosticsEnvEquals(value, L"off") ||
        MenuDiagnosticsEnvEquals(value, L"no"))
    {
        return false;
    }
    return true;
}

[[nodiscard]] std::wstring BuildDefaultMenuDiagnosticsTracePath()
{
    wchar_t tempPath[MAX_PATH + 1]{};
    const DWORD tempPathLength = GetTempPathW(static_cast<DWORD>(_countof(tempPath)), tempPath);
    std::wstring root = (tempPathLength > 0u && tempPathLength < _countof(tempPath)) ? std::wstring(tempPath, tempPathLength) : std::wstring(L".\\");
    if (! root.empty() && root.back() != L'\\' && root.back() != L'/')
    {
        root.push_back(L'\\');
    }
    return std::format(L"{}RedSalamander-DxUiMenuTrace-{}.log", root, GetCurrentProcessId());
}

void InitializeMenuDiagnosticsTrace() noexcept
{
    auto& state = GetMenuDiagnosticsTraceState();
    try
    {
        const std::wstring enabledValue = ReadMenuDiagnosticsEnvironmentVariable(L"REDSALAMANDER_DXUI_MENU_TRACE");
        if (! IsMenuDiagnosticsTraceRequested(enabledValue))
        {
            return;
        }

        std::wstring path = ReadMenuDiagnosticsEnvironmentVariable(L"REDSALAMANDER_DXUI_MENU_TRACE_FILE");
        if (path.empty())
        {
            path = BuildDefaultMenuDiagnosticsTracePath();
        }

        wil::unique_hfile file(CreateFileW(path.c_str(),
                                           FILE_APPEND_DATA,
                                           FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                           nullptr,
                                           CREATE_ALWAYS,
                                           FILE_ATTRIBUTE_NORMAL,
                                           nullptr));
        if (! file)
        {
            OutputDebugStringW(L"DxUi::MenuTrace: failed to open diagnostics trace file.\r\n");
            return;
        }

        constexpr wchar_t kUtf16Bom = 0xFEFF;
        DWORD bytesWritten          = 0u;
        static_cast<void>(WriteFile(file.get(), &kUtf16Bom, sizeof(kUtf16Bom), &bytesWritten, nullptr));

        state.path = std::move(path);
        state.file = std::move(file);
        state.enabled.store(true, std::memory_order_release);
        Debug::Info(L"DxUi::MenuTrace: writing context menu diagnostics to '{}'", state.path);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Diagnostics are optional evidence gathering; a trace setup failure must not alter menu input behavior.
        OutputDebugStringW(L"DxUi::MenuTrace: diagnostics initialization failed; continuing without trace file.\r\n");
    }
}

#if DXUI_MENU_PERSISTENT_DIAGNOSTICS
[[nodiscard]] bool ShouldTraceMenuLoopMessageImmediately(UINT message, bool popupMessage) noexcept
{
    if (popupMessage || IsMenuPriorityInputMessage(message))
    {
        return true;
    }

    switch (message)
    {
        case WM_TIMER:
        case WM_CAPTURECHANGED:
        case WM_CANCELMODE:
        case WM_ACTIVATEAPP:
        case WM_ACTIVATE:
        case WM_NCACTIVATE:
        case WM_SETFOCUS:
        case WM_KILLFOCUS:
        case WM_DESTROY:
        case WndMsg::kDxUiContextMenuRootHoverChanged:
            return true;
        default:
            return false;
    }
}
#endif

[[nodiscard]] bool IsMenuDiagnosticsTraceEnabled() noexcept
{
    try
    {
        auto& state = GetMenuDiagnosticsTraceState();
        std::call_once(state.initOnce, InitializeMenuDiagnosticsTrace);
        return state.enabled.load(std::memory_order_acquire);
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // std::call_once can fail before the flag is initialized; tracing stays disabled in that case.
        OutputDebugStringW(L"DxUi::MenuTrace: diagnostics enable check failed; continuing without trace file.\r\n");
        return false;
    }
}

void DisableMenuDiagnosticsTraceAfterFailure() noexcept
{
    auto& state = GetMenuDiagnosticsTraceState();
    state.enabled.store(false, std::memory_order_release);
    OutputDebugStringW(L"DxUi::MenuTrace: diagnostics disabled after a write failure.\r\n");
}

void WriteMenuDiagnosticsTraceLine(std::wstring_view eventName, std::wstring_view details) noexcept
{
    if (! IsMenuDiagnosticsTraceEnabled())
    {
        return;
    }

    auto& state = GetMenuDiagnosticsTraceState();
    try
    {
        SYSTEMTIME now{};
        GetLocalTime(&now);
        POINT cursor{};
        static_cast<void>(GetCursorPos(&cursor)); // getcursorpos-allow: diagnostic-only
        const uint64_t sequence = state.sequence.fetch_add(1u, std::memory_order_relaxed) + 1u;

        std::wstring line = std::format(L"{:04}-{:02}-{:02}T{:02}:{:02}:{:02}.{:03} seq={} pid={} tid={} event={} "
                                        L"cursor=({}, {}) capture={:#x} focus={:#x} active={:#x} foreground={:#x}",
                                        static_cast<unsigned int>(now.wYear),
                                        static_cast<unsigned int>(now.wMonth),
                                        static_cast<unsigned int>(now.wDay),
                                        static_cast<unsigned int>(now.wHour),
                                        static_cast<unsigned int>(now.wMinute),
                                        static_cast<unsigned int>(now.wSecond),
                                        static_cast<unsigned int>(now.wMilliseconds),
                                        sequence,
                                        static_cast<unsigned long>(GetCurrentProcessId()),
                                        static_cast<unsigned long>(GetCurrentThreadId()),
                                        eventName,
                                        cursor.x,
                                        cursor.y,
                                        TraceHwndValue(GetCapture()),
                                        TraceHwndValue(GetFocus()),
                                        TraceHwndValue(GetActiveWindow()),
                                        TraceHwndValue(GetForegroundWindow()));
        if (! details.empty())
        {
            line.push_back(L' ');
            line.append(details);
        }
        line.append(L"\r\n");

        const DWORD byteCount = static_cast<DWORD>(line.size() * sizeof(wchar_t));
        DWORD bytesWritten    = 0u;
        std::lock_guard lock(state.mutex);
        if (! state.file || WriteFile(state.file.get(), line.data(), byteCount, &bytesWritten, nullptr) == FALSE)
        {
            DisableMenuDiagnosticsTraceAfterFailure();
        }
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Diagnostics must never throw through a Win32 input path.
        DisableMenuDiagnosticsTraceAfterFailure();
    }
}

template <typename... Args>
void WriteMenuDiagnosticsTrace(std::wstring_view eventName, std::wformat_string<Args...> format, Args&&... args) noexcept
{
    if (! IsMenuDiagnosticsTraceEnabled())
    {
        return;
    }

    try
    {
        WriteMenuDiagnosticsTraceLine(eventName, std::format(format, std::forward<Args>(args)...));
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::exception&)
    {
        // Formatting is diagnostic only; disable tracing instead of disturbing menu dispatch.
        DisableMenuDiagnosticsTraceAfterFailure();
    }
}
#else
[[nodiscard]] constexpr bool IsMenuDiagnosticsTraceEnabled() noexcept
{
    return false;
}

void WriteMenuDiagnosticsTraceLine(std::wstring_view /*eventName*/, std::wstring_view /*details*/) noexcept
{
}

template <typename... Args>
void WriteMenuDiagnosticsTrace(std::wstring_view /*eventName*/, std::wformat_string<Args...> /*format*/, Args&&... /*args*/) noexcept
{
}
#endif

// ---------------------------------------------------------------------------
// Menu item visual style
// ---------------------------------------------------------------------------

struct MenuItemVisualStyle
{
    D2D1_COLOR_F background;
    D2D1_COLOR_F text;
    D2D1_COLOR_F accelText;
    D2D1_COLOR_F iconColor;
    D2D1_COLOR_F separatorColor;
    D2D1_COLOR_F hoverFill;
    D2D1_COLOR_F checkedFill;
    D2D1_COLOR_F checkColor;
    D2D1_COLOR_F chevronColor;
    D2D1_COLOR_F headerText;
    D2D1_COLOR_F borderColor;
};

struct MenuResolvedItemPaintStyle
{
    bool showHighlightFill     = false;
    bool usesRainbowHighlight  = false;
    D2D1_COLOR_F fill          = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F compositeFill = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F text          = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F accelText     = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F iconColor     = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F checkColor    = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
    D2D1_COLOR_F chevronColor  = D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f);
};

struct MenuSurfaceMaterialStyle
{
    D2D1_COLOR_F baseFill;
    D2D1_COLOR_F glazeTop;
    D2D1_COLOR_F glazeMid;
    D2D1_COLOR_F bottomShade;
    D2D1_COLOR_F outerBorder;
    D2D1_COLOR_F innerBorder;
    D2D1_COLOR_F topRim;
    float shadowInsetDip        = 3.0f;
    float shadowYOffsetDip      = 3.0f;
    float shadowSpreadDip       = 7.0f;
    float shadowOuterOpacity    = 0.18f;
    float shadowInnerOpacity    = 0.10f;
    float backdropOpacity       = 0.0f;
    float backdropBlurDip       = 0.0f;
    float backdropDetailOpacity = 0.0f;
};

struct MenuPopupShadowMargins
{
    float leftDip   = kMenuShadowLeftDip;
    float topDip    = kMenuShadowTopDip;
    float rightDip  = kMenuShadowRightDip;
    float bottomDip = kMenuShadowBottomDip;
};

struct ParsedMenuLabel
{
    std::wstring displayText;
    wchar_t mnemonic = L'\0';
};

struct DecodedMenuItemText
{
    std::wstring_view labelText;
    std::wstring_view acceleratorText;
};

struct MenuBackdropSnapshot
{
    WindowHostBitmapCapture capture;
    mutable wil::com_ptr<ID2D1Bitmap1> cachedBitmap;
    mutable wil::com_ptr<ID2D1Device> cachedDevice;
};

[[nodiscard]] std::wstring_view TrimMenuTextWhitespace(std::wstring_view text) noexcept
{
    size_t first = 0u;
    size_t last  = text.size();

    while (first < last && std::iswspace(static_cast<wint_t>(text[first])))
    {
        ++first;
    }

    while (last > first && std::iswspace(static_cast<wint_t>(text[last - 1u])))
    {
        --last;
    }

    return text.substr(first, last - first);
}

[[nodiscard]] DecodedMenuItemText DecodeMenuItemText(const MenuFlyoutItem& item) noexcept
{
    const std::wstring_view explicitAccelerator = TrimMenuTextWhitespace(item.acceleratorText);
    const std::wstring_view itemText            = item.text;
    std::wstring_view labelText                 = itemText;
    std::wstring_view embeddedAccelerator;

    if (const size_t tabPos = itemText.find(L'\t'); tabPos != std::wstring_view::npos)
    {
        labelText = itemText.substr(0u, tabPos);
        if (explicitAccelerator.empty() && (tabPos + 1u) < itemText.size())
        {
            embeddedAccelerator = itemText.substr(tabPos + 1u);
        }
    }

    return DecodedMenuItemText{
        .labelText       = TrimMenuTextWhitespace(labelText),
        .acceleratorText = explicitAccelerator.empty() ? TrimMenuTextWhitespace(embeddedAccelerator) : explicitAccelerator,
    };
}

[[nodiscard]] wchar_t NormalizeMenuMnemonicChar(wchar_t ch) noexcept
{
    return static_cast<wchar_t>(reinterpret_cast<UINT_PTR>(CharUpperW(reinterpret_cast<LPWSTR>(static_cast<UINT_PTR>(ch)))));
}

[[nodiscard]] ParsedMenuLabel ParseMenuLabel(std::wstring_view text) noexcept
{
    ParsedMenuLabel parsed{};
    parsed.displayText.reserve(text.size());

    for (size_t index = 0u; index < text.size(); ++index)
    {
        const wchar_t ch = text[index];
        if (ch == L'&')
        {
            if ((index + 1u) < text.size())
            {
                const wchar_t next = text[index + 1u];
                if (next == L'&')
                {
                    parsed.displayText.push_back(L'&');
                }
                else
                {
                    if (parsed.mnemonic == L'\0')
                    {
                        parsed.mnemonic = next;
                    }
                    parsed.displayText.push_back(next);
                }

                ++index;
                continue;
            }
        }

        parsed.displayText.push_back(ch);
    }

    return parsed;
}

[[nodiscard]] wchar_t ResolveMenuLabelMnemonic(const ParsedMenuLabel& label) noexcept
{
    if (label.mnemonic != L'\0')
    {
        return label.mnemonic;
    }

    for (const wchar_t ch : label.displayText)
    {
        if (! std::iswspace(static_cast<wint_t>(ch)))
        {
            return ch;
        }
    }

    return L'\0';
}

[[nodiscard]] D2D1_COLOR_F WithAlpha(const D2D1_COLOR_F& color, float alpha) noexcept
{
    return D2D1::ColorF(color.r, color.g, color.b, (std::clamp)(alpha, 0.0f, 1.0f));
}

[[nodiscard]] MenuResolvedItemPaintStyle ResolveMenuItemPaintStyle(
    const ThemePalette& theme, const MenuItemVisualStyle& style, const MenuFlyoutItem& item, std::wstring_view rainbowSeed, bool hovered) noexcept
{
    MenuResolvedItemPaintStyle resolved{
        .showHighlightFill    = false,
        .usesRainbowHighlight = false,
        .fill                 = style.hoverFill,
        .compositeFill        = style.background,
        .text                 = style.text,
        .accelText            = style.accelText,
        .iconColor            = style.iconColor,
        .checkColor           = style.checkColor,
        .chevronColor         = style.chevronColor,
    };

    if (! hovered || ! item.enabled)
    {
        return resolved;
    }

    resolved.showHighlightFill = true;
    if (theme.rainbowMode && ! theme.highContrast && ! rainbowSeed.empty())
    {
        resolved.usesRainbowHighlight = true;
        resolved.fill                 = RainbowMenuSelectionTint(rainbowSeed, theme.dark);
    }

    resolved.compositeFill                = CompositeOverBackground(resolved.fill, style.background);
    const D2D1_COLOR_F contrastForeground = ChooseContrastingTextColor(resolved.compositeFill);
    resolved.text                         = contrastForeground;
    resolved.accelText                    = contrastForeground;
    resolved.iconColor                    = contrastForeground;
    resolved.checkColor                   = contrastForeground;
    resolved.chevronColor                 = contrastForeground;
    return resolved;
}

[[nodiscard]] bool CaptureMenuBackdropScreenRegion(const RECT& screenRect, WindowHostBitmapCapture& outCapture) noexcept
{
    outCapture = {};

    const LONG widthPx  = screenRect.right - screenRect.left;
    const LONG heightPx = screenRect.bottom - screenRect.top;
    if (widthPx <= 0 || heightPx <= 0)
    {
        return false;
    }

    BITMAPINFO bmi{};
    bmi.bmiHeader.biSize        = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth       = widthPx;
    bmi.bmiHeader.biHeight      = -heightPx;
    bmi.bmiHeader.biPlanes      = 1;
    bmi.bmiHeader.biBitCount    = 32;
    bmi.bmiHeader.biCompression = BI_RGB;

    void* bits = nullptr;
    wil::unique_hdc_window screenDc{GetDC(nullptr)};
    if (! screenDc)
    {
        Debug::Warning(L"DxUi::Menu: unable to acquire screen DC for popup backdrop capture");
        return false;
    }

    wil::unique_hdc memoryDc{CreateCompatibleDC(screenDc.get())};
    if (! memoryDc)
    {
        Debug::Warning(L"DxUi::Menu: unable to create memory DC for popup backdrop capture");
        return false;
    }

    wil::unique_hbitmap bitmap{CreateDIBSection(screenDc.get(), &bmi, DIB_RGB_COLORS, &bits, nullptr, 0)};
    if (! bitmap || ! bits)
    {
        Debug::Warning(L"DxUi::Menu: unable to create DIB section for popup backdrop capture");
        return false;
    }

    [[maybe_unused]] const auto oldBitmap = wil::SelectObject(memoryDc.get(), bitmap.get());
    if (BitBlt(memoryDc.get(), 0, 0, widthPx, heightPx, screenDc.get(), screenRect.left, screenRect.top, SRCCOPY | CAPTUREBLT) == FALSE)
    {
        Debug::Warning(L"DxUi::Menu: BitBlt failed for popup backdrop capture (lastError={})", GetLastError());
        return false;
    }

    outCapture.widthPx  = static_cast<UINT>(widthPx);
    outCapture.heightPx = static_cast<UINT>(heightPx);
    outCapture.bgraPixels.resize(static_cast<size_t>(outCapture.widthPx) * static_cast<size_t>(outCapture.heightPx) * 4u);

    const auto* const sourceBytes = static_cast<const uint8_t*>(bits);
    std::copy_n(sourceBytes, outCapture.bgraPixels.size(), outCapture.bgraPixels.data());
    for (size_t offset = 3u; offset < outCapture.bgraPixels.size(); offset += 4u)
    {
        outCapture.bgraPixels[offset] = 0xFFu;
    }

    return true;
}

[[nodiscard]] ID2D1Bitmap1* EnsureMenuBackdropBitmap(WindowHost& host, MenuBackdropSnapshot& snapshot) noexcept
{
    if (snapshot.capture.widthPx == 0u || snapshot.capture.heightPx == 0u || snapshot.capture.bgraPixels.empty())
    {
        return nullptr;
    }

    ID2D1DeviceContext* const d2dContext = host.GetDeviceContext();
    if (! d2dContext)
    {
        return nullptr;
    }

    wil::com_ptr<ID2D1Device> device;
    d2dContext->GetDevice(device.put());
    if (! device)
    {
        return nullptr;
    }

    if (snapshot.cachedBitmap && snapshot.cachedDevice && snapshot.cachedDevice.get() == device.get())
    {
        return snapshot.cachedBitmap.get();
    }

    const D2D1_BITMAP_PROPERTIES1 bitmapProperties = D2D1::BitmapProperties1(
        D2D1_BITMAP_OPTIONS_NONE, D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED), host.GetDpi(), host.GetDpi());

    wil::com_ptr<ID2D1Bitmap1> bitmap;
    const UINT32 pitch = static_cast<UINT32>(snapshot.capture.widthPx * 4u);
    const HRESULT hr   = d2dContext->CreateBitmap(
        D2D1::SizeU(snapshot.capture.widthPx, snapshot.capture.heightPx), snapshot.capture.bgraPixels.data(), pitch, &bitmapProperties, bitmap.put());
    if (FAILED(hr) || ! bitmap)
    {
        Debug::Warning(L"DxUi::Menu: failed to create popup backdrop bitmap: 0x{:08X}", hr);
        return nullptr;
    }

    snapshot.cachedDevice = std::move(device);
    snapshot.cachedBitmap = std::move(bitmap);
    return snapshot.cachedBitmap.get();
}

[[nodiscard]] MenuItemVisualStyle ResolveMenuVisualStyle(const ThemePalette& theme) noexcept
{
    const D2D1_COLOR_F highlightBase = ChooseContrastingTextColor(theme.overlayBackground);
    D2D1_COLOR_F hoverFill           = theme.headerHovered;
    D2D1_COLOR_F checkedFill         = theme.selectionInactiveFill;
    if (theme.rainbowMode && ! theme.highContrast)
    {
        switch (theme.overlayMaterial)
        {
            case OverlayMaterial::Acrylic:
                hoverFill   = D2D1::ColorF(theme.accentPressed.r, theme.accentPressed.g, theme.accentPressed.b, theme.dark ? 0.82f : 0.72f);
                checkedFill = D2D1::ColorF(theme.accentHover.r, theme.accentHover.g, theme.accentHover.b, theme.dark ? 0.58f : 0.46f);
                break;
            case OverlayMaterial::MicaAlt:
                hoverFill   = D2D1::ColorF(theme.accentPressed.r, theme.accentPressed.g, theme.accentPressed.b, theme.dark ? 0.74f : 0.66f);
                checkedFill = D2D1::ColorF(theme.accentHover.r, theme.accentHover.g, theme.accentHover.b, theme.dark ? 0.50f : 0.40f);
                break;
            case OverlayMaterial::Mica:
                hoverFill   = D2D1::ColorF(theme.accentHover.r, theme.accentHover.g, theme.accentHover.b, theme.dark ? 0.52f : 0.42f);
                checkedFill = D2D1::ColorF(theme.accentHover.r, theme.accentHover.g, theme.accentHover.b, theme.dark ? 0.36f : 0.30f);
                break;
            case OverlayMaterial::Solid:
            default:
                hoverFill   = D2D1::ColorF(theme.accentHover.r, theme.accentHover.g, theme.accentHover.b, theme.dark ? 0.36f : 0.28f);
                checkedFill = D2D1::ColorF(theme.accentHover.r, theme.accentHover.g, theme.accentHover.b, theme.dark ? 0.28f : 0.22f);
                break;
        }
    }
    else
    {
        switch (theme.overlayMaterial)
        {
            case OverlayMaterial::Mica:
                hoverFill = WithAlpha(BlendColor(theme.headerHovered, highlightBase, theme.dark ? 0.16f : 0.08f), theme.dark ? 0.30f : 0.24f);
                break;
            case OverlayMaterial::MicaAlt:
                hoverFill = WithAlpha(BlendColor(theme.headerHovered, theme.accent, theme.dark ? 0.18f : 0.10f), theme.dark ? 0.34f : 0.28f);
                break;
            case OverlayMaterial::Acrylic:
                hoverFill = WithAlpha(BlendColor(theme.headerHovered, theme.accent, theme.dark ? 0.26f : 0.14f), theme.dark ? 0.42f : 0.34f);
                break;
            case OverlayMaterial::Solid:
            default: hoverFill = BlendColor(theme.headerHovered, theme.overlayBackground, theme.dark ? 0.10f : 0.04f); break;
        }
    }

    return MenuItemVisualStyle{
        .background     = theme.overlayBackground,
        .text           = theme.text,
        .accelText      = theme.subduedText,
        .iconColor      = theme.text,
        .separatorColor = BlendColor(theme.overlayBackground, theme.overlayBorder, theme.dark ? 0.62f : 0.46f),
        .hoverFill      = hoverFill,
        .checkedFill    = checkedFill,
        .checkColor     = theme.rainbowMode && ! theme.highContrast ? theme.accentPressed : theme.accent,
        .chevronColor   = theme.subduedText,
        .headerText     = theme.subduedText,
        .borderColor    = theme.overlayBorder,
    };
}

[[nodiscard]] wil::com_ptr<ID2D1LinearGradientBrush> CreateVerticalGradientBrush(
    ID2D1DeviceContext* dc, const D2D1_RECT_F& rect, const D2D1_COLOR_F& topColor, const D2D1_COLOR_F& midColor, const D2D1_COLOR_F& bottomColor) noexcept
{
    if (! dc)
    {
        return {};
    }

    const D2D1_GRADIENT_STOP stops[3] = {
        D2D1::GradientStop(0.0f, topColor),
        D2D1::GradientStop(0.42f, midColor),
        D2D1::GradientStop(1.0f, bottomColor),
    };

    wil::com_ptr<ID2D1GradientStopCollection> stopCollection;
    if (FAILED(dc->CreateGradientStopCollection(stops, static_cast<UINT32>(std::size(stops)), stopCollection.put())))
    {
        return {};
    }

    wil::com_ptr<ID2D1LinearGradientBrush> brush;
    if (FAILED(dc->CreateLinearGradientBrush(
            D2D1::LinearGradientBrushProperties(D2D1::Point2F(rect.left, rect.top), D2D1::Point2F(rect.left, rect.bottom)), stopCollection.get(), brush.put())))
    {
        return {};
    }

    return brush;
}

[[nodiscard]] MenuSurfaceMaterialStyle ResolveMenuSurfaceMaterialStyle(const ThemePalette& theme, const MenuItemVisualStyle& itemStyle) noexcept
{
    const D2D1_COLOR_F highlightBase = ChooseContrastingTextColor(theme.overlayBackground);
    MenuSurfaceMaterialStyle material{
        .baseFill           = WithAlpha(itemStyle.background, 0.96f),
        .glazeTop           = WithAlpha(BlendColor(theme.overlayBackground, highlightBase, theme.dark ? 0.12f : 0.08f), theme.dark ? 0.10f : 0.08f),
        .glazeMid           = WithAlpha(BlendColor(theme.overlayBackground, theme.windowBackground, theme.dark ? 0.32f : 0.18f), theme.dark ? 0.05f : 0.035f),
        .bottomShade        = D2D1::ColorF(0.0f, 0.0f, 0.0f, theme.dark ? 0.10f : 0.06f),
        .outerBorder        = WithAlpha(itemStyle.borderColor, theme.dark ? 0.76f : 0.62f),
        .innerBorder        = WithAlpha(BlendColor(itemStyle.borderColor, highlightBase, theme.dark ? 0.20f : 0.12f), theme.dark ? 0.48f : 0.42f),
        .topRim             = WithAlpha(highlightBase, theme.dark ? 0.18f : 0.14f),
        .shadowInsetDip     = 3.0f,
        .shadowYOffsetDip   = 3.0f,
        .shadowSpreadDip    = 9.0f,
        .shadowOuterOpacity = theme.dark ? 0.20f : 0.12f,
        .shadowInnerOpacity = theme.dark ? 0.12f : 0.07f,
        .backdropOpacity    = 0.0f,
        .backdropBlurDip    = 0.0f,
    };

    switch (theme.overlayMaterial)
    {
        case OverlayMaterial::Mica:
            material.baseFill = WithAlpha(BlendColor(theme.overlayBackground, theme.surfaceBackground, theme.dark ? 0.18f : 0.10f), theme.dark ? 0.58f : 0.62f);
            material.glazeTop = WithAlpha(BlendColor(theme.overlayBackground, highlightBase, theme.dark ? 0.24f : 0.14f), theme.dark ? 0.14f : 0.10f);
            material.glazeMid = WithAlpha(BlendColor(material.baseFill, theme.windowBackground, theme.dark ? 0.46f : 0.24f), theme.dark ? 0.08f : 0.05f);
            material.bottomShade        = D2D1::ColorF(0.0f, 0.0f, 0.0f, theme.dark ? 0.12f : 0.07f);
            material.outerBorder        = WithAlpha(BlendColor(itemStyle.borderColor, highlightBase, theme.dark ? 0.10f : 0.04f), theme.dark ? 0.82f : 0.68f);
            material.innerBorder        = WithAlpha(BlendColor(itemStyle.borderColor, highlightBase, theme.dark ? 0.30f : 0.20f), theme.dark ? 0.60f : 0.52f);
            material.topRim             = WithAlpha(highlightBase, theme.dark ? 0.24f : 0.19f);
            material.shadowOuterOpacity = theme.dark ? 0.24f : 0.16f;
            material.shadowInnerOpacity = theme.dark ? 0.14f : 0.09f;
            material.backdropOpacity    = ResolveOverlayBackdropOpacity(theme);
            material.backdropBlurDip    = ResolveOverlayBackdropBlurDip(theme);
            break;
        case OverlayMaterial::MicaAlt:
            material.baseFill = WithAlpha(BlendColor(theme.overlayBackground, theme.headerBackground, theme.dark ? 0.24f : 0.14f), theme.dark ? 0.48f : 0.58f);
            material.glazeTop = WithAlpha(BlendColor(material.baseFill, theme.headerHovered, theme.dark ? 0.48f : 0.28f), theme.dark ? 0.16f : 0.12f);
            material.glazeMid = WithAlpha(BlendColor(theme.headerHovered, theme.accent, theme.dark ? 0.40f : 0.24f), theme.dark ? 0.12f : 0.08f);
            material.bottomShade        = D2D1::ColorF(0.0f, 0.0f, 0.0f, theme.dark ? 0.14f : 0.08f);
            material.outerBorder        = WithAlpha(BlendColor(itemStyle.borderColor, theme.accent, theme.dark ? 0.22f : 0.12f), theme.dark ? 0.88f : 0.74f);
            material.innerBorder        = WithAlpha(BlendColor(material.outerBorder, highlightBase, theme.dark ? 0.36f : 0.22f), theme.dark ? 0.66f : 0.56f);
            material.topRim             = WithAlpha(BlendColor(highlightBase, theme.accent, 0.20f), theme.dark ? 0.28f : 0.22f);
            material.shadowOuterOpacity = theme.dark ? 0.26f : 0.17f;
            material.shadowInnerOpacity = theme.dark ? 0.16f : 0.10f;
            material.backdropOpacity    = ResolveOverlayBackdropOpacity(theme);
            material.backdropBlurDip    = ResolveOverlayBackdropBlurDip(theme);
            break;
        case OverlayMaterial::Acrylic:
            material.baseFill    = WithAlpha(BlendColor(theme.overlayBackground, theme.headerHovered, theme.dark ? 0.14f : 0.10f), theme.dark ? 0.12f : 0.16f);
            material.glazeTop    = WithAlpha(BlendColor(material.baseFill, highlightBase, theme.dark ? 0.24f : 0.14f), theme.dark ? 0.08f : 0.06f);
            material.glazeMid    = WithAlpha(BlendColor(theme.accent, theme.headerHovered, theme.dark ? 0.18f : 0.10f), theme.dark ? 0.04f : 0.03f);
            material.bottomShade = D2D1::ColorF(0.0f, 0.0f, 0.0f, theme.dark ? 0.05f : 0.03f);
            material.outerBorder = WithAlpha(BlendColor(itemStyle.borderColor, theme.accent, theme.dark ? 0.24f : 0.14f), theme.dark ? 0.76f : 0.64f);
            material.innerBorder = WithAlpha(BlendColor(material.outerBorder, highlightBase, theme.dark ? 0.32f : 0.18f), theme.dark ? 0.58f : 0.48f);
            material.topRim      = WithAlpha(BlendColor(highlightBase, theme.accent, 0.18f), theme.dark ? 0.14f : 0.12f);
            material.shadowInsetDip     = 2.0f;
            material.shadowSpreadDip    = 10.0f;
            material.shadowOuterOpacity = theme.dark ? 0.30f : 0.20f;
            material.shadowInnerOpacity = theme.dark ? 0.20f : 0.13f;
            material.backdropOpacity    = ResolveOverlayBackdropOpacity(theme);
            material.backdropBlurDip    = ResolveOverlayBackdropBlurDip(theme);
            break;
        case OverlayMaterial::Solid:
        default: break;
    }

    return material;
}

void PaintMenuBackdropSurface(
    WindowHost& host, MenuBackdropSnapshot* snapshot, const D2D1_RECT_F& menuRect, float cornerRadiusDip, const MenuSurfaceMaterialStyle& material) noexcept
{
    if (! snapshot || material.backdropOpacity <= 0.0f || material.backdropBlurDip <= 0.0f)
    {
        return;
    }

    auto* const dc             = host.GetDeviceContext();
    ID2D1Bitmap1* const bitmap = EnsureMenuBackdropBitmap(host, *snapshot);
    if (! dc || ! bitmap)
    {
        return;
    }

    wil::com_ptr<ID2D1Factory> factory;
    dc->GetFactory(factory.put());

    wil::com_ptr<ID2D1RoundedRectangleGeometry> roundedGeometry;
    if (! factory || FAILED(factory->CreateRoundedRectangleGeometry(D2D1::RoundedRect(menuRect, cornerRadiusDip, cornerRadiusDip), roundedGeometry.put())) ||
        ! roundedGeometry)
    {
        return;
    }

    wil::com_ptr<ID2D1Effect> blurEffect;
    if (FAILED(dc->CreateEffect(kMenuGaussianBlurEffectId, blurEffect.put())) || ! blurEffect)
    {
        return;
    }

    blurEffect->SetInput(0u, bitmap);
    blurEffect->SetValue(D2D1_GAUSSIANBLUR_PROP_STANDARD_DEVIATION, material.backdropBlurDip);
    blurEffect->SetValue(D2D1_GAUSSIANBLUR_PROP_OPTIMIZATION, D2D1_GAUSSIANBLUR_OPTIMIZATION_BALANCED);
    blurEffect->SetValue(D2D1_GAUSSIANBLUR_PROP_BORDER_MODE, D2D1_BORDER_MODE_HARD);

    wil::com_ptr<ID2D1Layer> layer;
    if (FAILED(dc->CreateLayer(layer.put())) || ! layer)
    {
        return;
    }

    const D2D1_RECT_F sourceRect = D2D1::RectF(
        0.0f, 0.0f, host.PixelsToDip(static_cast<float>(snapshot->capture.widthPx)), host.PixelsToDip(static_cast<float>(snapshot->capture.heightPx)));

    const D2D1_LAYER_PARAMETERS1 layerParameters = D2D1::LayerParameters1(menuRect,
                                                                          roundedGeometry.get(),
                                                                          D2D1_ANTIALIAS_MODE_PER_PRIMITIVE,
                                                                          D2D1::Matrix3x2F::Identity(),
                                                                          (std::clamp)(material.backdropOpacity, 0.0f, 1.0f),
                                                                          nullptr,
                                                                          D2D1_LAYER_OPTIONS1_NONE);
    const D2D1_POINT_2F targetOffset             = D2D1::Point2F(menuRect.left, menuRect.top);
    dc->PushLayer(layerParameters, layer.get());
    dc->DrawImage(blurEffect.get(), &targetOffset, &sourceRect, D2D1_INTERPOLATION_MODE_LINEAR, D2D1_COMPOSITE_MODE_SOURCE_OVER);
    if (material.backdropDetailOpacity > 0.0f)
    {
        dc->DrawBitmap(bitmap, menuRect, (std::clamp)(material.backdropDetailOpacity, 0.0f, 1.0f), D2D1_INTERPOLATION_MODE_LINEAR);
    }
    dc->PopLayer();
}

void PaintMenuMaterialSurface(
    WindowHost& host, MenuBackdropSnapshot* snapshot, const D2D1_RECT_F& menuRect, float cornerRadiusDip, const MenuSurfaceMaterialStyle& material) noexcept
{
    auto* const dc = host.GetDeviceContext();
    if (! dc)
    {
        return;
    }

    const D2D1_RECT_F shadowTargetRect   = InflateRect(menuRect, -material.shadowInsetDip, -material.shadowInsetDip);
    const D2D1_ROUNDED_RECT menuRounded  = D2D1::RoundedRect(menuRect, cornerRadiusDip, cornerRadiusDip);
    const D2D1_RECT_F innerRect          = InflateRect(menuRect, -1.0f, -1.0f);
    const D2D1_ROUNDED_RECT innerRounded = D2D1::RoundedRect(innerRect, (std::max)(0.0f, cornerRadiusDip - 1.0f), (std::max)(0.0f, cornerRadiusDip - 1.0f));

    DrawDropShadow(host,
                   shadowTargetRect,
                   (std::max)(0.0f, cornerRadiusDip - material.shadowInsetDip),
                   material.shadowYOffsetDip,
                   material.shadowSpreadDip,
                   material.shadowOuterOpacity,
                   material.shadowInnerOpacity);

    PaintMenuBackdropSurface(host, snapshot, menuRect, cornerRadiusDip, material);
    DrawRoundedRect(host, menuRect, material.baseFill, D2D1::ColorF(0, 0, 0, 0), cornerRadiusDip);

    if (wil::com_ptr<ID2D1LinearGradientBrush> glazeBrush = CreateVerticalGradientBrush(
            dc, menuRect, material.glazeTop, material.glazeMid, D2D1::ColorF(material.glazeMid.r, material.glazeMid.g, material.glazeMid.b, 0.0f)))
    {
        dc->FillRoundedRectangle(menuRounded, glazeBrush.get());
    }

    if (wil::com_ptr<ID2D1LinearGradientBrush> bottomShadeBrush =
            CreateVerticalGradientBrush(dc,
                                        menuRect,
                                        D2D1::ColorF(material.bottomShade.r, material.bottomShade.g, material.bottomShade.b, 0.0f),
                                        D2D1::ColorF(material.bottomShade.r, material.bottomShade.g, material.bottomShade.b, material.bottomShade.a * 0.35f),
                                        material.bottomShade))
    {
        dc->FillRoundedRectangle(menuRounded, bottomShadeBrush.get());
    }

    if (auto* const topRimBrush = host.GetSolidBrush(material.topRim))
    {
        const float rimY = menuRect.top + 1.0f;
        dc->DrawLine(
            D2D1::Point2F(menuRect.left + cornerRadiusDip - 1.0f, rimY), D2D1::Point2F(menuRect.right - cornerRadiusDip + 1.0f, rimY), topRimBrush, 1.0f);
    }

    if (auto* const outerBorderBrush = host.GetSolidBrush(material.outerBorder))
    {
        dc->DrawRoundedRectangle(menuRounded, outerBorderBrush, 1.0f);
    }
    if (auto* const innerBorderBrush = host.GetSolidBrush(material.innerBorder))
    {
        dc->DrawRoundedRectangle(innerRounded, innerBorderBrush, 1.0f);
    }
}

[[nodiscard]] IWICImagingFactory* GetMenuIconWicFactory() noexcept
{
    static wil::com_ptr<IWICImagingFactory> factory;
    if (factory)
    {
        return factory.get();
    }

    HRESULT hr = CoCreateInstance(CLSID_WICImagingFactory2, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(factory.put()));
    if (FAILED(hr))
    {
        hr = CoCreateInstance(CLSID_WICImagingFactory, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(factory.put()));
    }
    if (FAILED(hr))
    {
        Debug::Warning(L"DxUi::Menu: Failed to create WIC factory for popup icons: 0x{:08X}", hr);
        factory.reset();
        return nullptr;
    }

    return factory.get();
}

[[nodiscard]] ID2D1Bitmap1* EnsureMenuBitmapIconBitmap(WindowHost& host, const MenuFlyoutItem::BitmapIcon& icon) noexcept
{
    if (! icon.sourceBitmap)
    {
        return nullptr;
    }

    ID2D1DeviceContext* const d2dContext = host.GetDeviceContext();
    if (! d2dContext)
    {
        return nullptr;
    }

    wil::com_ptr<ID2D1Device> device;
    d2dContext->GetDevice(device.put());
    if (! device)
    {
        return nullptr;
    }

    if (icon.cachedBitmap && icon.cachedDevice && icon.cachedDevice.get() == device.get())
    {
        return icon.cachedBitmap.get();
    }

    IWICImagingFactory* const wicFactory = GetMenuIconWicFactory();
    if (! wicFactory)
    {
        return nullptr;
    }

    wil::com_ptr<IWICBitmap> wicBitmap;
    HRESULT hr = wicFactory->CreateBitmapFromHBITMAP(icon.sourceBitmap.get(), nullptr, WICBitmapUsePremultipliedAlpha, wicBitmap.put());
    if (FAILED(hr))
    {
        Debug::Warning(L"DxUi::Menu: Failed to create WIC bitmap from menu icon: 0x{:08X}", hr);
        return nullptr;
    }

    D2D1_BITMAP_PROPERTIES1 bitmapProperties = D2D1::BitmapProperties1(
        D2D1_BITMAP_OPTIONS_NONE, D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED), host.GetDpi(), host.GetDpi());

    wil::com_ptr<ID2D1Bitmap1> bitmap;
    hr = d2dContext->CreateBitmapFromWicBitmap(wicBitmap.get(), &bitmapProperties, bitmap.put());
    if (FAILED(hr))
    {
        Debug::Warning(L"DxUi::Menu: Failed to create D2D bitmap for menu icon: 0x{:08X}", hr);
        return nullptr;
    }

    icon.cachedDevice = std::move(device);
    icon.cachedBitmap = std::move(bitmap);
    return icon.cachedBitmap.get();
}

void DrawMenuBitmapIcon(WindowHost& host, const MenuFlyoutItem::BitmapIcon& icon, const D2D1_RECT_F& iconRectDip, float opacity) noexcept
{
    ID2D1Bitmap1* const bitmap = EnsureMenuBitmapIconBitmap(host, icon);
    if (! bitmap)
    {
        return;
    }

    const float sourceWidthDip  = icon.widthPx > 0u ? host.PixelsToDip(static_cast<float>(icon.widthPx)) : 16.0f;
    const float sourceHeightDip = icon.heightPx > 0u ? host.PixelsToDip(static_cast<float>(icon.heightPx)) : 16.0f;
    const float maxWidthDip     = (std::max)(0.0f, (iconRectDip.right - iconRectDip.left) - 4.0f);
    const float maxHeightDip    = (std::max)(0.0f, (iconRectDip.bottom - iconRectDip.top) - 4.0f);
    if (sourceWidthDip <= 0.0f || sourceHeightDip <= 0.0f || maxWidthDip <= 0.0f || maxHeightDip <= 0.0f)
    {
        return;
    }

    const float scale          = (std::min)(1.0f, (std::min)(maxWidthDip / sourceWidthDip, maxHeightDip / sourceHeightDip));
    const float drawWidthDip   = sourceWidthDip * scale;
    const float drawHeightDip  = sourceHeightDip * scale;
    const D2D1_RECT_F drawRect = D2D1::RectF(iconRectDip.left + (((iconRectDip.right - iconRectDip.left) - drawWidthDip) * 0.5f),
                                             iconRectDip.top + (((iconRectDip.bottom - iconRectDip.top) - drawHeightDip) * 0.5f),
                                             iconRectDip.left + (((iconRectDip.right - iconRectDip.left) - drawWidthDip) * 0.5f) + drawWidthDip,
                                             iconRectDip.top + (((iconRectDip.bottom - iconRectDip.top) - drawHeightDip) * 0.5f) + drawHeightDip);

    host.GetDeviceContext()->DrawBitmap(bitmap, drawRect, opacity, D2D1_INTERPOLATION_MODE_LINEAR);
}

[[nodiscard]] bool IsInvokableMenuItem(const MenuFlyoutItem& item) noexcept
{
    if (! item.enabled)
    {
        return false;
    }

    switch (item.kind)
    {
        case MenuItemKind::Separator:
        case MenuItemKind::Header:
        case MenuItemKind::Info: return false;
        case MenuItemKind::Slider:
            for (const auto& stop : item.sliderStops)
            {
                if (stop.commandId != 0)
                {
                    return true;
                }
            }
            return false;
        case MenuItemKind::Standard:
        case MenuItemKind::Toggle:
        case MenuItemKind::Radio: return ! item.children.empty() || item.commandId != 0;
    }

    return false;
}

[[nodiscard]] uint32_t ClampSliderValue(const MenuFlyoutItem& item) noexcept
{
    if (item.sliderStops.empty())
    {
        return 0u;
    }

    const uint32_t maxStop = static_cast<uint32_t>(item.sliderStops.size() - 1u);
    return (std::min)(item.sliderValue, maxStop);
}

[[nodiscard]] int GetSliderCommandId(const MenuFlyoutItem& item, uint32_t stopIndex) noexcept
{
    if (stopIndex >= item.sliderStops.size())
    {
        return 0;
    }
    return item.sliderStops[stopIndex].commandId;
}

[[nodiscard]] D2D1_RECT_F GetSliderTrackRect(const D2D1_RECT_F& itemRect) noexcept
{
    return D2D1::RectF(itemRect.left + kSliderTrackInsetDip,
                       itemRect.top + kSliderTrackTopDip,
                       itemRect.right - kSliderTrackInsetDip,
                       itemRect.top + kSliderTrackTopDip + kSliderTrackHeightDip);
}

[[nodiscard]] uint32_t HitTestSliderStop(const MenuFlyoutItem& item, const D2D1_RECT_F& itemRect, D2D1_POINT_2F pointDip) noexcept
{
    const size_t stopCount = item.sliderStops.size();
    if (stopCount <= 1u)
    {
        return 0u;
    }

    const D2D1_RECT_F trackRect = GetSliderTrackRect(itemRect);
    const float trackWidth      = (std::max)(1.0f, trackRect.right - trackRect.left);
    const float normalized      = std::clamp((pointDip.x - trackRect.left) / trackWidth, 0.0f, 1.0f);
    const float scaled          = normalized * static_cast<float>(stopCount - 1u);
    return static_cast<uint32_t>(std::lround(scaled));
}

struct MenuItemLayoutRects
{
    D2D1_RECT_F itemRectDip        = D2D1::RectF();
    D2D1_RECT_F iconRectDip        = D2D1::RectF();
    D2D1_RECT_F textRectDip        = D2D1::RectF();
    D2D1_RECT_F acceleratorRectDip = D2D1::RectF();
    D2D1_RECT_F chevronRectDip     = D2D1::RectF();
};

[[nodiscard]] int RoundToIntSaturated(double value) noexcept
{
    if (value <= static_cast<double>((std::numeric_limits<int>::min)()))
    {
        return (std::numeric_limits<int>::min)();
    }
    if (value >= static_cast<double>((std::numeric_limits<int>::max)()))
    {
        return (std::numeric_limits<int>::max)();
    }
    return static_cast<int>(std::lround(value));
}

[[nodiscard]] LONG RoundToLongSaturated(double value) noexcept
{
    if (value <= static_cast<double>((std::numeric_limits<LONG>::min)()))
    {
        return (std::numeric_limits<LONG>::min)();
    }
    if (value >= static_cast<double>((std::numeric_limits<LONG>::max)()))
    {
        return (std::numeric_limits<LONG>::max)();
    }
    return static_cast<LONG>(std::lround(value));
}

[[nodiscard]] int DipExtentToPixels(float extentDip, UINT dpi) noexcept
{
    const double scaled = static_cast<double>(extentDip) * static_cast<double>(dpi) / 96.0;
    return (std::max)(1, RoundToIntSaturated(scaled));
}

// ---------------------------------------------------------------------------
// Menu popup window class
// ---------------------------------------------------------------------------

static constexpr wchar_t kMenuWindowClass[] = L"DxUi_ContextMenu";
static std::atomic<bool> s_classRegistered  = false;
#if defined(ENABLE_TESTS)
static constexpr UINT kMenuDebugCaptureBitmapMessage = WM_APP + 0x214;
static constexpr UINT kMenuDebugGetItemTextMessage   = WM_APP + 0x215;
static constexpr UINT kMenuDebugSetBackdropMessage   = WM_APP + 0x216;
static constexpr UINT kMenuDebugGetStateMessage      = WM_APP + 0x217;
static constexpr UINT kMenuDebugGetItemRectMessage   = WM_APP + 0x218;
static constexpr UINT kMenuDebugGetItemPaintMessage  = WM_APP + 0x219;
#endif

struct MenuController; // forward
[[nodiscard]] D2D1_RECT_F GetItemRect(
    const MenuFlyoutItem* items, size_t count, size_t targetIndex, float menuWidthDip, float itemHeightDip, float headerHeightDip) noexcept;
[[nodiscard]] bool ProcessMenuPopupMessage(MenuController& controller, HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);
void FinalizeAsyncMenuController(MenuController& controller) noexcept;
void DestroyMenuPopupWindow(MenuPopup& popup) noexcept;
[[nodiscard]] UINT ResolveMenuPopupMessageDpi(HWND hwnd, UINT msg, WPARAM wp) noexcept;
void RelayoutMenuPopupForDpi(MenuPopup& popup, UINT dpi, const RECT* suggestedWindowRect) noexcept;

struct MenuPopup
{
    enum class ScrollbarHotPart : uint8_t
    {
        None,
        Track,
        Thumb,
    };

    enum class SubmenuHoverTimerKind : uint8_t
    {
        None,
        PendingOpen,
        PendingClose,
    };

    MenuPopup()                            = default;
    MenuPopup(const MenuPopup&)            = delete;
    MenuPopup& operator=(const MenuPopup&) = delete;
    MenuPopup(MenuPopup&&)                 = delete;
    MenuPopup& operator=(MenuPopup&&)      = delete;

    HWND hwnd = nullptr;
    WindowHost host;
    std::vector<MenuFlyoutItem> ownedItems;
    const MenuFlyoutItem* items = nullptr;
    size_t itemCount            = 0;
    MenuController* controller  = nullptr;
    std::optional<size_t> hoveredIndex;
    std::optional<size_t> keyboardIndex;
    std::optional<size_t> openedFromItemIndex;
    UINT dpi              = USER_DEFAULT_SCREEN_DPI;
    float menuWidthDip    = 0.0f;
    float menuHeightDip   = 0.0f;
    float windowWidthDip  = 0.0f;
    float windowHeightDip = 0.0f;
    RECT surfaceRectPx{};
    RECT windowRectPx{};
    float contentHeightDip          = 0.0f;
    float acceleratorColumnWidthDip = 0.0f;
    MenuBackdropSnapshot backdropSnapshot;
    MenuPopupShadowMargins shadowMargins;
    bool usesSystemBackdrop           = false;
    bool usesAppBackdropBlur          = false;
    bool isSubmenu                    = false;
    bool mouseInsideMenu              = false;
    float scrollOffsetDip             = 0.0f;
    bool draggingScrollbarThumb       = false;
    float scrollbarDragOffsetDip      = 0.0f;
    ScrollbarHotPart scrollbarHotPart = ScrollbarHotPart::None;

    UINT_PTR hoverTimerId                = 0;
    SubmenuHoverTimerKind hoverTimerKind = SubmenuHoverTimerKind::None;
    size_t hoverTimerItemIndex           = SIZE_MAX;

    [[nodiscard]] float DipToPixel(float dip) const noexcept
    {
        return dip * static_cast<float>(dpi) / 96.0f;
    }

    [[nodiscard]] float PixelToDip(float px) const noexcept
    {
        return px * 96.0f / static_cast<float>(dpi);
    }

    [[nodiscard]] bool IsNavigableItem(size_t index) const noexcept
    {
        if (index >= itemCount)
            return false;
        return items[index].enabled && items[index].kind != MenuItemKind::Separator && items[index].kind != MenuItemKind::Header &&
               items[index].kind != MenuItemKind::Info;
    }

    [[nodiscard]] bool NeedsScrollbar() const noexcept
    {
        return contentHeightDip > menuHeightDip && menuHeightDip > 0.0f;
    }

    [[nodiscard]] D2D1_RECT_F GetSurfaceRect() const noexcept
    {
        return D2D1::RectF(shadowMargins.leftDip, shadowMargins.topDip, shadowMargins.leftDip + menuWidthDip, shadowMargins.topDip + menuHeightDip);
    }

    [[nodiscard]] float GetScrollExtent() const noexcept
    {
        return (std::max)(0.0f, contentHeightDip - menuHeightDip);
    }

    void ClampScrollOffset() noexcept
    {
        scrollOffsetDip = (std::clamp)(scrollOffsetDip, 0.0f, GetScrollExtent());
    }

    [[nodiscard]] D2D1_RECT_F GetViewportRect() const noexcept
    {
        const D2D1_RECT_F surfaceRect = GetSurfaceRect();
        return NeedsScrollbar() ? D2D1::RectF(surfaceRect.left, surfaceRect.top, surfaceRect.right - kScrollbarThicknessDip, surfaceRect.bottom) : surfaceRect;
    }

    [[nodiscard]] D2D1_RECT_F GetScrollbarTrackRect() const noexcept
    {
        const D2D1_RECT_F surfaceRect = GetSurfaceRect();
        return D2D1::RectF(surfaceRect.right - kScrollbarThicknessDip, surfaceRect.top, surfaceRect.right, surfaceRect.bottom);
    }

    [[nodiscard]] D2D1_RECT_F GetScrollbarThumbRect() const noexcept
    {
        return ComputeScrollbarThumbRect(
            GetScrollbarTrackRect(), ScrollbarOrientation::Vertical, menuHeightDip, contentHeightDip, scrollOffsetDip, GetScrollExtent());
    }

    [[nodiscard]] float GetContentWidthDip() const noexcept
    {
        return NeedsScrollbar() ? menuWidthDip - kScrollbarThicknessDip : menuWidthDip;
    }

    [[nodiscard]] RECT GetInteractiveScreenRect() const noexcept
    {
        RECT windowRect{};
        if (! hwnd || GetWindowRect(hwnd, &windowRect) == FALSE)
        {
            return RECT{};
        }

        const auto roundToLong = [](double value) noexcept
        {
            if (value <= static_cast<double>((std::numeric_limits<LONG>::min)()))
            {
                return (std::numeric_limits<LONG>::min)();
            }
            if (value >= static_cast<double>((std::numeric_limits<LONG>::max)()))
            {
                return (std::numeric_limits<LONG>::max)();
            }
            return static_cast<LONG>(std::lround(value));
        };

        return RECT{roundToLong(static_cast<double>(windowRect.left) + DipToPixel(shadowMargins.leftDip)),
                    roundToLong(static_cast<double>(windowRect.top) + DipToPixel(shadowMargins.topDip)),
                    roundToLong(static_cast<double>(windowRect.left) + DipToPixel(shadowMargins.leftDip + menuWidthDip)),
                    roundToLong(static_cast<double>(windowRect.top) + DipToPixel(shadowMargins.topDip + menuHeightDip))};
    }

    void EnsureItemVisible(size_t index) noexcept
    {
        const ThemePalette& theme   = host.GetTheme();
        const float itemHeightDip   = ResolveMenuItemHeightDip(theme);
        const float headerHeightDip = ResolveMenuHeaderHeightDip(theme);
        const D2D1_RECT_F itemRect  = GetItemRect(items, itemCount, index, GetContentWidthDip(), itemHeightDip, headerHeightDip);
        if (itemRect.bottom <= itemRect.top || ! NeedsScrollbar())
        {
            return;
        }

        if (itemRect.top < scrollOffsetDip)
        {
            scrollOffsetDip = itemRect.top;
        }
        else if (itemRect.bottom > (scrollOffsetDip + menuHeightDip))
        {
            scrollOffsetDip = itemRect.bottom - menuHeightDip;
        }
        ClampScrollOffset();
    }

    [[nodiscard]] std::optional<size_t> FindNextNavigableItem(size_t from, bool forward) const noexcept
    {
        if (itemCount == 0)
            return std::nullopt;
        for (size_t i = 0; i < itemCount; ++i)
        {
            size_t idx;
            if (forward)
                idx = (from + 1 + i) % itemCount;
            else
                idx = (from + itemCount - 1 - i) % itemCount;
            if (IsNavigableItem(idx))
                return idx;
        }
        return std::nullopt;
    }

    [[nodiscard]] std::optional<size_t> FindFirstNavigableItem() const noexcept
    {
        for (size_t i = 0; i < itemCount; ++i)
        {
            if (IsNavigableItem(i))
                return i;
        }
        return std::nullopt;
    }

    [[nodiscard]] std::optional<size_t> FindInitialKeyboardItem(bool focusFirstNavigableItem) const noexcept
    {
        if (focusFirstNavigableItem)
        {
            return FindFirstNavigableItem();
        }

        return std::nullopt;
    }

    [[nodiscard]] std::optional<size_t> FindLastNavigableItem() const noexcept
    {
        for (size_t i = itemCount; i > 0; --i)
        {
            if (IsNavigableItem(i - 1))
                return i - 1;
        }
        return std::nullopt;
    }

    [[nodiscard]] std::optional<size_t> FindMnemonicItem(wchar_t ch) const noexcept
    {
        const wchar_t upper = NormalizeMenuMnemonicChar(ch);
        for (size_t i = 0; i < itemCount; ++i)
        {
            if (! IsNavigableItem(i) || ! items[i].enabled)
                continue;

            const ParsedMenuLabel label = ParseMenuLabel(DecodeMenuItemText(items[i]).labelText);
            const wchar_t itemMnemonic  = ResolveMenuLabelMnemonic(label);
            if (itemMnemonic != L'\0' && NormalizeMenuMnemonicChar(itemMnemonic) == upper)
            {
                return i;
            }
        }
        return std::nullopt;
    }
};

void InvalidatePopup(MenuPopup& popup) noexcept;

// ---------------------------------------------------------------------------
// Menu controller — owns the cascade chain, modal loop, result
// ---------------------------------------------------------------------------

struct MenuController
{
    HWND ownerHwnd = nullptr;
    ThemePalette theme;
    MenuItemVisualStyle style;
    ContextMenuSessionCallbacks sessionCallbacks;
    ContextMenuClosedCallback asyncOnClosed;
    std::optional<int> result;
    bool running = true;
    bool asyncSession = false;
    bool asyncFinalizing = false;
    bool asyncInteractionActive = false;
    // True while DestroyMenuPopupWindow runs DestroyWindow on one of this
    // controller's popups. Distinguishes controller-initiated popup destruction
    // (submenu close, chain replacement, root switch) from external teardown so
    // the popup WndProc does not finalize the async session re-entrantly.
    bool destroyingPopupWindow = false;
    HWND previousCapture = nullptr;
    HWND previousFocus = nullptr;
    bool ignoreInitialLeftButtonUp = false;
    bool ignoreInitialRightButtonUp = false;
    bool leftButtonDownInPopup = false;
    bool rightButtonDownInPopup = false;
    POINT lastPointerScreenPoint{};
    bool hasLastPointerScreenPoint = false;
    POINT lastRootSwitchPointerScreenPoint{};
    bool hasLastRootSwitchPointerScreenPoint = false;
    DWORD lastKeyboardRootSwitchMessageTime = 0;
    bool hasLastKeyboardRootSwitchMessageTime = false;
#if defined(ENABLE_TESTS)
    uint64_t rootPointerSwitchCount         = 0;
    uint64_t rootSwitchImmediateRenderCount = 0;
#endif

    // Cascade stack: [0] = root menu, [1..N] = submenus
    std::vector<std::unique_ptr<MenuPopup>> popups;

    // Root items (owned copy)
    std::vector<MenuFlyoutItem> rootItems;

    void Dismiss() noexcept
    {
        DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.dismiss",
                                  L"owner={:#x} popups={} resultBefore={} runningBefore={}",
                                  TraceHwndValue(ownerHwnd),
                                  popups.size(),
                                  result.value_or(-1),
                                  TraceBool(running));
        running = false;
    }

    void InvokeItem(int commandId) noexcept
    {
        DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.invoke",
                                  L"owner={:#x} commandId={} popups={} runningBefore={}",
                                  TraceHwndValue(ownerHwnd),
                                  commandId,
                                  popups.size(),
                                  TraceBool(running));
        result  = commandId;
        running = false;
    }

    void CloseTopmostSubmenu() noexcept
    {
        if (popups.size() > 1)
        {
            auto& parent = popups[popups.size() - 2];
            parent->hoveredIndex.reset();
            auto& child = popups.back();
            DestroyMenuPopupWindow(*child);
            popups.pop_back();
            if (! popups.empty() && popups.back()->hwnd)
            {
                InvalidatePopup(*popups.back());
            }
        }
    }

    [[nodiscard]] bool IsPointInAnyPopup(POINT screenPt) const noexcept
    {
        for (const auto& popup : popups)
        {
            if (popup->hwnd)
            {
                const RECT rc = popup->GetInteractiveScreenRect();
                if (PtInRect(&rc, screenPt))
                    return true;
            }
        }
        return false;
    }

    [[nodiscard]] MenuPopup* FindPopupForHwnd(HWND hwnd) const noexcept
    {
        for (const auto& popup : popups)
        {
            if (popup->hwnd == hwnd)
                return popup.get();
        }
        return nullptr;
    }

    [[nodiscard]] MenuPopup* GetRootPopup() const noexcept
    {
        return popups.empty() ? nullptr : popups.front().get();
    }

    [[nodiscard]] MenuPopup* GetTopmostPopup() const noexcept
    {
        return popups.empty() ? nullptr : popups.back().get();
    }
};

void InvalidatePopup(MenuPopup& popup) noexcept
{
    if (! popup.hwnd || IsWindow(popup.hwnd) == FALSE)
    {
        return;
    }

    InvalidateRect(popup.hwnd, nullptr, FALSE);
}

[[nodiscard]] int TracePopupIndex(const MenuController& controller, const MenuPopup* popup) noexcept
{
    if (! popup)
    {
        return -1;
    }

    for (size_t index = 0; index < controller.popups.size(); ++index)
    {
        if (controller.popups[index].get() == popup)
        {
            return index <= static_cast<size_t>(std::numeric_limits<int>::max()) ? static_cast<int>(index) : -1;
        }
    }

    return -1;
}

[[nodiscard]] POINT UnpackMousePoint(LPARAM lParam) noexcept
{
    return POINT{static_cast<LONG>(static_cast<short>(LOWORD(static_cast<DWORD_PTR>(lParam)))),
                 static_cast<LONG>(static_cast<short>(HIWORD(static_cast<DWORD_PTR>(lParam))))};
}

[[nodiscard]] POINT ResolveMouseScreenPoint(const MSG& msg) noexcept
{
    POINT screenPt{};
    if (msg.message == WM_MOUSEWHEEL || msg.message == WM_MOUSEHWHEEL)
    {
        return UnpackMousePoint(msg.lParam);
    }

    screenPt = UnpackMousePoint(msg.lParam);
    if (msg.hwnd && IsWindow(msg.hwnd) != FALSE)
    {
        POINT clientPoint = screenPt;
        if (ClientToScreen(msg.hwnd, &clientPoint) != FALSE)
        {
            return clientPoint;
        }
    }

    return screenPt;
}

enum class MenuInputSource : uint8_t
{
    ModalMessage,
    PopupWndProc,
};

enum class MenuPointerKind : uint8_t
{
    Move,
    LeftDown,
    LeftUp,
    RightDown,
    RightUp,
    Wheel,
};

enum class MenuKeyboardKind : uint8_t
{
    Up,
    Down,
    Left,
    Right,
    Home,
    End,
    Enter,
    Space,
    Escape,
    Tab,
    F10,
    Alt,
    Mnemonic,
};

struct MenuPointerEvent
{
    MenuInputSource source = MenuInputSource::ModalMessage;
    MenuPointerKind kind  = MenuPointerKind::Move;
    HWND hwnd             = nullptr;
    POINT screenPoint{};
    DWORD messageTime = 0;
    std::optional<D2D1_POINT_2F> deliveredPopupPointDip;
    WPARAM wParam       = 0;
    LPARAM lParam       = 0;
    bool mayInvoke      = false;
    bool mayDismiss     = false;
    bool maySwitchRoot  = false;
    bool mayTakeKeyboardFocus = false;
};

struct MenuKeyboardEvent
{
    MenuInputSource source = MenuInputSource::ModalMessage;
    MenuKeyboardKind kind  = MenuKeyboardKind::Mnemonic;
    HWND hwnd             = nullptr;
    UINT virtualKey       = 0;
    wchar_t mnemonic      = L'\0';
    DWORD messageTime     = 0;
    LPARAM lParam         = 0;
};

enum class MenuInputDisposition : uint8_t
{
    Ignored,
    Consumed,
    HoverChanged,
    Dismissed,
    Invoked,
    SwitchedRoot,
};

[[nodiscard]] bool IsMessageTimeAfter(DWORD candidate, DWORD reference) noexcept
{
    return static_cast<LONG>(candidate - reference) > 0;
}

[[nodiscard]] DWORD CurrentMessageTime() noexcept
{
    return static_cast<DWORD>(GetMessageTime());
}

void RememberKeyboardRootSwitchMessageTime(MenuController& controller, DWORD messageTime) noexcept
{
    controller.lastKeyboardRootSwitchMessageTime    = messageTime;
    controller.hasLastKeyboardRootSwitchMessageTime = true;
    controller.hasLastRootSwitchPointerScreenPoint  = false;
    controller.lastRootSwitchPointerScreenPoint     = {};
}

void ClearKeyboardRootSwitchMessageTime(MenuController& controller) noexcept
{
    controller.hasLastKeyboardRootSwitchMessageTime = false;
    controller.lastKeyboardRootSwitchMessageTime    = 0;
}

[[nodiscard]] bool IsStaleRootSwitchPointerMoveAfterKeyboard(const MenuController& controller, const MenuPointerEvent& event) noexcept
{
    if ((event.source != MenuInputSource::ModalMessage && event.source != MenuInputSource::PopupWndProc) || ! controller.hasLastKeyboardRootSwitchMessageTime)
    {
        return false;
    }

    if (! controller.hasLastRootSwitchPointerScreenPoint)
    {
        return true;
    }

    return event.messageTime != 0 && ! IsMessageTimeAfter(event.messageTime, controller.lastKeyboardRootSwitchMessageTime);
}

void RememberRootSwitchPointerPoint(MenuController& controller, POINT screenPoint) noexcept
{
    controller.lastRootSwitchPointerScreenPoint    = screenPoint;
    controller.hasLastRootSwitchPointerScreenPoint = true;
}

[[nodiscard]] bool ShouldRememberRejectedRootSwitchPointer(const MenuPointerEvent& event) noexcept
{
    return event.source == MenuInputSource::ModalMessage;
}

#if DXUI_MENU_PERSISTENT_DIAGNOSTICS
[[nodiscard]] const wchar_t* TraceInputSourceName(MenuInputSource source) noexcept
{
    switch (source)
    {
        case MenuInputSource::ModalMessage: return L"ModalMessage";
        case MenuInputSource::PopupWndProc: return L"PopupWndProc";
    }
    return L"unknown";
}

[[nodiscard]] const wchar_t* TracePointerKindName(MenuPointerKind kind) noexcept
{
    switch (kind)
    {
        case MenuPointerKind::Move: return L"Move";
        case MenuPointerKind::LeftDown: return L"LeftDown";
        case MenuPointerKind::LeftUp: return L"LeftUp";
        case MenuPointerKind::RightDown: return L"RightDown";
        case MenuPointerKind::RightUp: return L"RightUp";
        case MenuPointerKind::Wheel: return L"Wheel";
    }
    return L"unknown";
}

[[nodiscard]] const wchar_t* TraceKeyboardKindName(MenuKeyboardKind kind) noexcept
{
    switch (kind)
    {
        case MenuKeyboardKind::Up: return L"Up";
        case MenuKeyboardKind::Down: return L"Down";
        case MenuKeyboardKind::Left: return L"Left";
        case MenuKeyboardKind::Right: return L"Right";
        case MenuKeyboardKind::Home: return L"Home";
        case MenuKeyboardKind::End: return L"End";
        case MenuKeyboardKind::Enter: return L"Enter";
        case MenuKeyboardKind::Space: return L"Space";
        case MenuKeyboardKind::Escape: return L"Escape";
        case MenuKeyboardKind::Tab: return L"Tab";
        case MenuKeyboardKind::F10: return L"F10";
        case MenuKeyboardKind::Alt: return L"Alt";
        case MenuKeyboardKind::Mnemonic: return L"Mnemonic";
    }
    return L"unknown";
}

[[nodiscard]] const wchar_t* TraceDispositionName(MenuInputDisposition disposition) noexcept
{
    switch (disposition)
    {
        case MenuInputDisposition::Ignored: return L"Ignored";
        case MenuInputDisposition::Consumed: return L"Consumed";
        case MenuInputDisposition::HoverChanged: return L"HoverChanged";
        case MenuInputDisposition::Dismissed: return L"Dismissed";
        case MenuInputDisposition::Invoked: return L"Invoked";
        case MenuInputDisposition::SwitchedRoot: return L"SwitchedRoot";
    }
    return L"unknown";
}

[[nodiscard]] uint64_t TraceOptionalSizeValue(std::optional<size_t> value) noexcept
{
    return value.has_value() ? static_cast<uint64_t>(value.value()) : UINT64_MAX;
}
#endif

// ---------------------------------------------------------------------------
// Window class registration
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Window class registration — WndProc delegates to WindowHost
// ---------------------------------------------------------------------------

#if defined(ENABLE_TESTS)
struct MenuDebugGetItemTextRequest
{
    size_t itemIndex      = 0u;
    std::wstring* outText = nullptr;
};

struct MenuDebugGetItemRectRequest
{
    size_t itemIndex        = 0u;
    D2D1_RECT_F* outRectDip = nullptr;
};

struct MenuDebugGetItemPaintRequest
{
    size_t itemIndex                                = 0u;
    ContextMenuPopupItemPaintDebugState* outState = nullptr;
};

bool TryGetMenuPopupItemText(const MenuPopup& popup, size_t itemIndex, std::wstring& outText)
{
    outText.clear();
    if (itemIndex >= popup.itemCount)
    {
        return false;
    }

    outText = ParseMenuLabel(DecodeMenuItemText(popup.items[itemIndex]).labelText).displayText;
    return true;
}
#endif

static LRESULT CALLBACK MenuWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    if (msg == WM_MOUSEACTIVATE)
    {
        return MA_NOACTIVATE;
    }

    auto* popup = reinterpret_cast<MenuPopup*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (popup)
    {
        MenuController* const controller = popup->controller;
#if defined(ENABLE_TESTS)
        if (msg == kMenuDebugCaptureBitmapMessage)
        {
            auto* const outCapture = reinterpret_cast<WindowHostBitmapCapture*>(lp);
            if (! outCapture)
            {
                return FALSE;
            }

            *outCapture = {};
            return popup->host.DebugCaptureBitmap(*outCapture) ? TRUE : FALSE;
        }
        if (msg == kMenuDebugGetItemTextMessage)
        {
            auto* const request = reinterpret_cast<MenuDebugGetItemTextRequest*>(lp);
            if (! request || ! request->outText)
            {
                return FALSE;
            }

            return TryGetMenuPopupItemText(*popup, request->itemIndex, *request->outText) ? TRUE : FALSE;
        }
        if (msg == kMenuDebugGetStateMessage)
        {
            auto* const outState = reinterpret_cast<ContextMenuPopupDebugState*>(lp);
            return outState && DebugGetContextMenuPopupState(hwnd, *outState) ? TRUE : FALSE;
        }
        if (msg == kMenuDebugGetItemRectMessage)
        {
            auto* const request = reinterpret_cast<MenuDebugGetItemRectRequest*>(lp);
            if (! request || ! request->outRectDip)
            {
                return FALSE;
            }

            return DebugGetContextMenuPopupItemRect(hwnd, request->itemIndex, *request->outRectDip) ? TRUE : FALSE;
        }
        if (msg == kMenuDebugGetItemPaintMessage)
        {
            auto* const request = reinterpret_cast<MenuDebugGetItemPaintRequest*>(lp);
            if (! request || ! request->outState)
            {
                return FALSE;
            }

            return DebugGetContextMenuPopupItemPaint(hwnd, request->itemIndex, *request->outState) ? TRUE : FALSE;
        }
        if (msg == kMenuDebugSetBackdropMessage)
        {
            const auto* const capture = reinterpret_cast<const WindowHostBitmapCapture*>(lp);
            if (! capture || capture->widthPx == 0u || capture->heightPx == 0u || capture->bgraPixels.empty())
            {
                return FALSE;
            }

            popup->backdropSnapshot.capture      = *capture;
            popup->backdropSnapshot.cachedBitmap = nullptr;
            popup->backdropSnapshot.cachedDevice = nullptr;
            popup->usesAppBackdropBlur           = true;
            popup->host.Invalidate();
            return TRUE;
        }
#endif
        if (msg == WM_NCDESTROY)
        {
            popup->host.Detach();
            if (popup->hwnd == hwnd)
            {
                popup->hwnd = nullptr;
            }
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
            // Finalize only for external teardown (e.g. the owner window being
            // destroyed while the async menu is open). Controller-initiated
            // destroys (submenu close, chain replacement, root switch) reach this
            // handler synchronously from DestroyMenuPopupWindow's DestroyWindow
            // call; finalizing here would delete the controller and its popups
            // while that caller is still using them.
            if (controller && controller->asyncSession && ! controller->asyncFinalizing && ! controller->destroyingPopupWindow)
            {
                controller->Dismiss();
                FinalizeAsyncMenuController(*controller);
            }
            return DefWindowProcW(hwnd, msg, wp, lp);
        }
        if (msg == WM_DPICHANGED || msg == WM_DPICHANGED_AFTERPARENT)
        {
            const UINT dpi = ResolveMenuPopupMessageDpi(hwnd, msg, wp);
            RelayoutMenuPopupForDpi(*popup, dpi, msg == WM_DPICHANGED ? reinterpret_cast<const RECT*>(lp) : nullptr);
            return 0;
        }
        if (msg == WM_NCHITTEST)
        {
            const POINT screenPoint{GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
            if (! PtInRect(&popup->surfaceRectPx, screenPoint))
            {
                return HTTRANSPARENT;
            }

            return HTCLIENT;
        }

        if (controller && controller->asyncSession && controller->asyncInteractionActive && ! controller->asyncFinalizing &&
            ! controller->destroyingPopupWindow)
        {
            const bool dismissForActivation =
                (msg == WM_ACTIVATEAPP && wp == FALSE) ||
                (msg == WM_ACTIVATE && LOWORD(static_cast<DWORD_PTR>(wp)) == WA_INACTIVE) ||
                (msg == WM_NCACTIVATE && wp == FALSE);
            const bool dismissForCancel = msg == WM_CANCELMODE;
            const bool dismissForCaptureLoss =
                msg == WM_CAPTURECHANGED &&
                (! controller->GetRootPopup() || reinterpret_cast<HWND>(lp) != controller->GetRootPopup()->hwnd);
            if (dismissForActivation || dismissForCancel || dismissForCaptureLoss)
            {
                controller->Dismiss();
                FinalizeAsyncMenuController(*controller);
                return 0;
            }
        }

        if (controller && ProcessMenuPopupMessage(*controller, hwnd, msg, wp, lp))
        {
            if (controller->asyncSession && ! controller->running)
            {
                FinalizeAsyncMenuController(*controller);
            }
            return 0;
        }

        bool handled   = false;
        LRESULT result = popup->host.HandleMessage(hwnd, msg, wp, lp, handled);
        if (handled)
            return result;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

void EnsureMenuWindowClass(HINSTANCE hInstance)
{
    if (s_classRegistered.exchange(true, std::memory_order_acq_rel))
        return;

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = MenuWndProc;
    wc.hInstance     = hInstance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.lpszClassName = kMenuWindowClass;
    RegisterClassExW(&wc);
}

// ---------------------------------------------------------------------------
// Compute menu content size
// ---------------------------------------------------------------------------

[[nodiscard]] D2D1_SIZE_F ComputeMenuSize(const MenuFlyoutItem* items, size_t count, WindowHost& host, float* outAcceleratorColumnWidthDip = nullptr) noexcept
{
    const ThemePalette& theme   = host.GetTheme();
    const float itemHeightDip   = ResolveMenuItemHeightDip(theme);
    const float headerHeightDip = ResolveMenuHeaderHeightDip(theme);
    float maxTextWidth          = 0.0f;
    float maxAccelWidth         = 0.0f;
    float totalHeight           = kMenuPaddingTopDip + kMenuPaddingBottomDip;
    bool hasSubmenu             = false;
    bool hasSlider              = false;

    for (size_t i = 0; i < count; ++i)
    {
        const auto& item = items[i];
        switch (item.kind)
        {
            case MenuItemKind::Separator: totalHeight += kSeparatorHeightDip; break;
            case MenuItemKind::Header:
                totalHeight += headerHeightDip;
                // Measure header text
                if (auto* fmt = host.GetTextFormat(FontRole::Small))
                {
                    const ParsedMenuLabel label = ParseMenuLabel(DecodeMenuItemText(item).labelText);
                    wil::com_ptr<IDWriteTextLayout> layout;
                    if (auto* factory = host.GetWriteFactory())
                    {
                        if (SUCCEEDED(factory->CreateTextLayout(
                                label.displayText.c_str(), static_cast<UINT32>(label.displayText.size()), fmt, kTextMeasureWidthDip, headerHeightDip, &layout)))
                        {
                            DWRITE_TEXT_METRICS metrics{};
                            if (SUCCEEDED(layout->GetMetrics(&metrics)))
                            {
                                maxTextWidth = (std::max)(maxTextWidth, metrics.widthIncludingTrailingWhitespace);
                            }
                        }
                    }
                }
                break;
            case MenuItemKind::Slider:
                totalHeight += kSliderItemHeightDip;
                hasSlider = true;
                if (auto* fmt = host.GetTextFormat(FontRole::Body))
                {
                    const DecodedMenuItemText decoded = DecodeMenuItemText(item);
                    const ParsedMenuLabel label       = ParseMenuLabel(decoded.labelText);
                    wil::com_ptr<IDWriteTextLayout> layout;
                    if (auto* factory = host.GetWriteFactory())
                    {
                        if (SUCCEEDED(factory->CreateTextLayout(
                                label.displayText.c_str(), static_cast<UINT32>(label.displayText.size()), fmt, kTextMeasureWidthDip, itemHeightDip, &layout)))
                        {
                            DWRITE_TEXT_METRICS metrics{};
                            if (SUCCEEDED(layout->GetMetrics(&metrics)))
                            {
                                maxTextWidth = (std::max)(maxTextWidth, metrics.widthIncludingTrailingWhitespace);
                            }
                        }

                        if (! decoded.acceleratorText.empty())
                        {
                            if (SUCCEEDED(factory->CreateTextLayout(decoded.acceleratorText.data(),
                                                                    static_cast<UINT32>(decoded.acceleratorText.size()),
                                                                    fmt,
                                                                    kTextMeasureWidthDip,
                                                                    itemHeightDip,
                                                                    &layout)))
                            {
                                DWRITE_TEXT_METRICS metrics{};
                                if (SUCCEEDED(layout->GetMetrics(&metrics)))
                                {
                                    maxAccelWidth = (std::max)(maxAccelWidth, metrics.widthIncludingTrailingWhitespace);
                                }
                            }
                        }
                    }
                }
                break;
            case MenuItemKind::Info:
            case MenuItemKind::Standard:
            case MenuItemKind::Toggle:
            case MenuItemKind::Radio:
                totalHeight += itemHeightDip;
                if (! item.children.empty())
                    hasSubmenu = true;
                // Measure item text
                if (auto* fmt = host.GetTextFormat(FontRole::Body))
                {
                    const DecodedMenuItemText decoded = DecodeMenuItemText(item);
                    const ParsedMenuLabel label       = ParseMenuLabel(decoded.labelText);
                    wil::com_ptr<IDWriteTextLayout> layout;
                    if (auto* factory = host.GetWriteFactory())
                    {
                        if (SUCCEEDED(factory->CreateTextLayout(
                                label.displayText.c_str(), static_cast<UINT32>(label.displayText.size()), fmt, kTextMeasureWidthDip, itemHeightDip, &layout)))
                        {
                            DWRITE_TEXT_METRICS metrics{};
                            if (SUCCEEDED(layout->GetMetrics(&metrics)))
                            {
                                maxTextWidth = (std::max)(maxTextWidth, metrics.widthIncludingTrailingWhitespace);
                            }
                        }
                        // Measure accelerator text
                        if (! decoded.acceleratorText.empty())
                        {
                            if (SUCCEEDED(factory->CreateTextLayout(decoded.acceleratorText.data(),
                                                                    static_cast<UINT32>(decoded.acceleratorText.size()),
                                                                    fmt,
                                                                    kTextMeasureWidthDip,
                                                                    itemHeightDip,
                                                                    &layout)))
                            {
                                DWRITE_TEXT_METRICS metrics{};
                                if (SUCCEEDED(layout->GetMetrics(&metrics)))
                                {
                                    maxAccelWidth = (std::max)(maxAccelWidth, metrics.widthIncludingTrailingWhitespace);
                                }
                            }
                        }
                    }
                }
                break;
        }
    }

    if (outAcceleratorColumnWidthDip)
    {
        *outAcceleratorColumnWidthDip = maxAccelWidth;
    }

    float width = kTextLeftPaddingDip + maxTextWidth + 16.0f; // gap after text
    if (maxAccelWidth > 0.0f)
        width += kTextToAccelGapDip + maxAccelWidth + kAccelRightPaddingDip;
    if (hasSubmenu)
        width += kChevronAreaWidthDip;
    width += kMenuBorderDip * 2.0f;
    if (hasSlider)
        width = (std::max)(width, kSliderMenuMinWidthDip);

    width = std::clamp(width, kMenuMinWidthDip, kMenuMaxWidthDip);

    return D2D1::SizeF(width, totalHeight);
}

// ---------------------------------------------------------------------------
// Get item rect for a given index (in menu-local DIP coordinates)
// ---------------------------------------------------------------------------

[[nodiscard]] D2D1_RECT_F GetItemRect(
    const MenuFlyoutItem* items, size_t count, size_t targetIndex, float menuWidthDip, float itemHeightDip, float headerHeightDip) noexcept
{
    float y = kMenuPaddingTopDip;
    for (size_t i = 0; i < count; ++i)
    {
        float h = itemHeightDip;
        if (items[i].kind == MenuItemKind::Separator)
            h = kSeparatorHeightDip;
        else if (items[i].kind == MenuItemKind::Header)
            h = headerHeightDip;
        else if (items[i].kind == MenuItemKind::Slider)
            h = kSliderItemHeightDip;
        if (i == targetIndex)
        {
            return D2D1::RectF(kMenuBorderDip, y, menuWidthDip - kMenuBorderDip, y + h);
        }
        y += h;
    }
    return D2D1::RectF(0, 0, 0, 0);
}

[[nodiscard]] D2D1_RECT_F GetVisibleItemRect(const MenuPopup& popup, size_t targetIndex) noexcept
{
    const ThemePalette& theme   = popup.host.GetTheme();
    const float itemHeightDip   = ResolveMenuItemHeightDip(theme);
    const float headerHeightDip = ResolveMenuHeaderHeightDip(theme);
    D2D1_RECT_F rect            = GetItemRect(popup.items, popup.itemCount, targetIndex, popup.GetContentWidthDip(), itemHeightDip, headerHeightDip);
    rect.left += popup.shadowMargins.leftDip;
    rect.right += popup.shadowMargins.leftDip;
    rect.top += popup.shadowMargins.topDip;
    rect.bottom += popup.shadowMargins.topDip;
    rect.top -= popup.scrollOffsetDip;
    rect.bottom -= popup.scrollOffsetDip;
    return rect;
}

[[nodiscard]] D2D1_RECT_F ClampHorizontalRect(float left, float right, float top, float bottom, float minLeft) noexcept
{
    const float resolvedRight = (std::max)(right, minLeft);
    const float resolvedLeft  = (std::min)(left, resolvedRight);
    return D2D1::RectF(resolvedLeft, top, resolvedRight, bottom);
}

[[nodiscard]] MenuItemLayoutRects GetMenuItemLayoutRects(const MenuPopup& popup, size_t targetIndex) noexcept
{
    MenuItemLayoutRects layout{};
    if (targetIndex >= popup.itemCount)
    {
        return layout;
    }

    const ThemePalette& theme   = popup.host.GetTheme();
    const float itemHeightDip   = ResolveMenuItemHeightDip(theme);
    const float headerHeightDip = ResolveMenuHeaderHeightDip(theme);
    const auto& item            = popup.items[targetIndex];
    float rowHeightDip          = itemHeightDip;
    if (item.kind == MenuItemKind::Header)
        rowHeightDip = headerHeightDip;
    else if (item.kind == MenuItemKind::Slider)
        rowHeightDip = kSliderItemHeightDip;
    const float contentWidthDip             = popup.GetContentWidthDip();
    const float surfaceLeftDip              = popup.GetSurfaceRect().left;
    const std::wstring_view acceleratorText = DecodeMenuItemText(item).acceleratorText;

    layout.itemRectDip = GetVisibleItemRect(popup, targetIndex);
    layout.iconRectDip = D2D1::RectF(surfaceLeftDip + kIconSlotLeftInsetDip,
                                     layout.itemRectDip.top,
                                     surfaceLeftDip + kIconSlotLeftInsetDip + kIconSlotWidthDip,
                                     layout.itemRectDip.top + rowHeightDip);

    const float reservedChevronWidthDip   = item.children.empty() || item.kind == MenuItemKind::Slider ? 0.0f : kChevronAreaWidthDip;
    const float acceleratorColumnWidthDip = acceleratorText.empty() ? 0.0f : popup.acceleratorColumnWidthDip;
    const float reservedAccelWidthDip     = acceleratorColumnWidthDip > 0.0f ? (acceleratorColumnWidthDip + kTextToAccelGapDip) : 0.0f;
    const float textRightDip              = surfaceLeftDip + contentWidthDip - kAccelRightPaddingDip - reservedChevronWidthDip - reservedAccelWidthDip;
    layout.textRectDip                    = ClampHorizontalRect(surfaceLeftDip + kTextLeftPaddingDip,
                                                                textRightDip,
                                                                layout.itemRectDip.top,
                                                                layout.itemRectDip.top + rowHeightDip,
                                                                surfaceLeftDip + kTextLeftPaddingDip);

    if (acceleratorColumnWidthDip > 0.0f)
    {
        const float accelRightDip = surfaceLeftDip + contentWidthDip - kAccelRightPaddingDip - reservedChevronWidthDip;
        const float accelLeftDip  = accelRightDip - acceleratorColumnWidthDip;
        layout.acceleratorRectDip =
            ClampHorizontalRect(accelLeftDip, accelRightDip, layout.itemRectDip.top, layout.itemRectDip.top + rowHeightDip, layout.textRectDip.right);
    }

    if (! item.children.empty())
    {
        layout.chevronRectDip = D2D1::RectF(surfaceLeftDip + contentWidthDip - kChevronRightPaddingDip - 16.0f,
                                            layout.itemRectDip.top,
                                            surfaceLeftDip + contentWidthDip - kChevronRightPaddingDip,
                                            layout.itemRectDip.top + rowHeightDip);
    }

    if (item.kind == MenuItemKind::Slider)
    {
        const float sliderTextTopDip    = layout.itemRectDip.top + 6.0f;
        const float sliderTextBottomDip = sliderTextTopDip + 22.0f;
        layout.iconRectDip              = D2D1::RectF();
        layout.textRectDip              = ClampHorizontalRect(
            surfaceLeftDip + kAccelRightPaddingDip, textRightDip, sliderTextTopDip, sliderTextBottomDip, surfaceLeftDip + kAccelRightPaddingDip);
        if (acceleratorColumnWidthDip > 0.0f)
        {
            const float accelRightDip = surfaceLeftDip + contentWidthDip - kAccelRightPaddingDip;
            const float accelLeftDip  = accelRightDip - acceleratorColumnWidthDip;
            layout.acceleratorRectDip = ClampHorizontalRect(accelLeftDip, accelRightDip, sliderTextTopDip, sliderTextBottomDip, layout.textRectDip.right);
        }
    }

    return layout;
}

// ---------------------------------------------------------------------------
// Hit-test: find item index under a DIP point (menu-local)
// ---------------------------------------------------------------------------

[[nodiscard]] std::optional<size_t> HitTestMenuItem(
    const MenuFlyoutItem* items, size_t count, float menuWidthDip, float itemHeightDip, float headerHeightDip, D2D1_POINT_2F pointDip) noexcept
{
    float y = kMenuPaddingTopDip;
    for (size_t i = 0; i < count; ++i)
    {
        float h = itemHeightDip;
        if (items[i].kind == MenuItemKind::Separator)
            h = kSeparatorHeightDip;
        else if (items[i].kind == MenuItemKind::Header)
            h = headerHeightDip;
        else if (items[i].kind == MenuItemKind::Slider)
            h = kSliderItemHeightDip;
        if (pointDip.y >= y && pointDip.y < y + h && pointDip.x >= 0.0f && pointDip.x < menuWidthDip)
        {
            if (items[i].kind != MenuItemKind::Separator && items[i].kind != MenuItemKind::Header && items[i].kind != MenuItemKind::Info)
            {
                return i;
            }
            return std::nullopt; // Over a non-interactive item
        }
        y += h;
    }
    return std::nullopt;
}

[[nodiscard]] std::optional<size_t> HitTestMenuItem(const MenuPopup& popup, D2D1_POINT_2F pointDip) noexcept
{
    if (! PointInRect(popup.GetViewportRect(), pointDip))
    {
        return std::nullopt;
    }

    const ThemePalette& theme   = popup.host.GetTheme();
    const float itemHeightDip   = ResolveMenuItemHeightDip(theme);
    const float headerHeightDip = ResolveMenuHeaderHeightDip(theme);
    return HitTestMenuItem(popup.items,
                           popup.itemCount,
                           popup.GetContentWidthDip(),
                           itemHeightDip,
                           headerHeightDip,
                           D2D1::Point2F(pointDip.x, pointDip.y + popup.scrollOffsetDip));
}

// ---------------------------------------------------------------------------
// Paint menu content (called from WM_PAINT via WindowHost)
// ---------------------------------------------------------------------------

class MenuContentControl final : public Control
{
public:
    MenuPopup* popup = nullptr;

    void Paint(WindowHost& host) const override
    {
        if (! popup || ! popup->controller)
            return;

        Debug::Perf::Scope menuPaintPerf(L"dxui.menu.popup.paint");

        auto* dc = host.GetDeviceContext();
        if (! dc)
            return;

        const auto& style             = popup->controller->style;
        const auto& theme             = popup->controller->theme;
        const float itemHeightDip     = ResolveMenuItemHeightDip(theme);
        const float headerHeightDip   = ResolveMenuHeaderHeightDip(theme);
        const float contentWidth      = popup->GetContentWidthDip();
        const D2D1_RECT_F surfaceRect = popup->GetSurfaceRect();
        const float surfaceLeft       = surfaceRect.left;
        const float surfaceTop        = surfaceRect.top;

        // Background
        const MenuSurfaceMaterialStyle material = ResolveMenuSurfaceMaterialStyle(theme, style);
        PaintMenuMaterialSurface(host, popup->usesAppBackdropBlur ? &popup->backdropSnapshot : nullptr, surfaceRect, kMenuCornerRadiusDip, material);

        const D2D1_RECT_F viewportRect = popup->GetViewportRect();
        dc->PushAxisAlignedClip(viewportRect, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);

        // Items
        float y = kMenuPaddingTopDip;
        for (size_t i = 0; i < popup->itemCount; ++i)
        {
            const auto& item = popup->items[i];
            float itemHeight = itemHeightDip;
            if (item.kind == MenuItemKind::Separator)
                itemHeight = kSeparatorHeightDip;
            else if (item.kind == MenuItemKind::Header)
                itemHeight = headerHeightDip;
            else if (item.kind == MenuItemKind::Slider)
                itemHeight = kSliderItemHeightDip;

            const D2D1_RECT_F itemRect        = D2D1::RectF(kMenuBorderDip, y, contentWidth - kMenuBorderDip, y + itemHeight);
            const D2D1_RECT_F visibleItemRect = D2D1::RectF(surfaceLeft + itemRect.left,
                                                            surfaceTop + itemRect.top - popup->scrollOffsetDip,
                                                            surfaceLeft + itemRect.right,
                                                            surfaceTop + itemRect.bottom - popup->scrollOffsetDip);

            if (item.kind == MenuItemKind::Separator)
            {
                // Separator line
                const float lineY     = visibleItemRect.top + kSeparatorMarginDip + 0.5f;
                const float lineLeft  = surfaceLeft + kItemHoverInsetDip;
                const float lineRight = surfaceLeft + contentWidth - kItemHoverInsetDip;
                auto* brush           = host.GetSolidBrush(style.separatorColor);
                if (brush)
                {
                    dc->DrawLine(D2D1::Point2F(lineLeft, lineY), D2D1::Point2F(lineRight, lineY), brush, 1.0f);
                }
                y += itemHeight;
                continue;
            }

            if (item.kind == MenuItemKind::Header)
            {
                // Header text (Caption font, subdued)
                const MenuItemLayoutRects layout = GetMenuItemLayoutRects(*popup, i);
                const ParsedMenuLabel label      = ParseMenuLabel(DecodeMenuItemText(item).labelText);
                const D2D1_RECT_F textRect =
                    D2D1::RectF(layout.textRectDip.left, layout.textRectDip.top, surfaceLeft + contentWidth - kAccelRightPaddingDip, layout.textRectDip.bottom);
                DrawCenteredText(
                    host, label.displayText, textRect, FontRole::Small, style.headerText, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
                y += itemHeight;
                continue;
            }

            if (item.kind == MenuItemKind::Slider)
            {
                const MenuItemLayoutRects layout           = GetMenuItemLayoutRects(*popup, i);
                const bool isHovered                       = (popup->hoveredIndex.has_value() && popup->hoveredIndex.value() == i) ||
                                                             (popup->keyboardIndex.has_value() && popup->keyboardIndex.value() == i);
                const bool isDisabled                      = ! item.enabled;
                const DecodedMenuItemText decoded          = DecodeMenuItemText(item);
                const ParsedMenuLabel label                = ParseMenuLabel(decoded.labelText);
                const MenuResolvedItemPaintStyle itemPaint = ResolveMenuItemPaintStyle(theme, style, item, label.displayText, isHovered);
                const D2D1_COLOR_F textColor = isDisabled ? D2D1::ColorF(itemPaint.text.r, itemPaint.text.g, itemPaint.text.b, 0.4f) : itemPaint.text;
                const D2D1_COLOR_F accelColor =
                    isDisabled ? D2D1::ColorF(itemPaint.accelText.r, itemPaint.accelText.g, itemPaint.accelText.b, 0.4f) : itemPaint.accelText;

                if (itemPaint.showHighlightFill)
                {
                    const D2D1_RECT_F hoverRect = D2D1::RectF(
                        visibleItemRect.left + kItemHoverInsetDip, visibleItemRect.top, visibleItemRect.right - kItemHoverInsetDip, visibleItemRect.bottom);
                    const D2D1_COLOR_F transparent = D2D1::ColorF(0, 0, 0, 0);
                    DrawRoundedRect(host, hoverRect, itemPaint.fill, transparent, kItemHoverRadiusDip);
                }

                DrawCenteredText(
                    host, label.displayText, layout.textRectDip, FontRole::Body, textColor, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_CENTER);

                if (! decoded.acceleratorText.empty())
                {
                    DrawCenteredText(host,
                                     decoded.acceleratorText,
                                     layout.acceleratorRectDip,
                                     FontRole::Body,
                                     accelColor,
                                     DWRITE_TEXT_ALIGNMENT_TRAILING,
                                     DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
                }

                const D2D1_RECT_F trackRect = GetSliderTrackRect(layout.itemRectDip);
                const D2D1_COLOR_F trackColor =
                    isDisabled ? D2D1::ColorF(style.separatorColor.r, style.separatorColor.g, style.separatorColor.b, 0.45f) : style.separatorColor;
                const D2D1_COLOR_F activeColor =
                    isDisabled ? D2D1::ColorF(itemPaint.checkColor.r, itemPaint.checkColor.g, itemPaint.checkColor.b, 0.45f) : itemPaint.checkColor;
                auto* trackBrush  = host.GetSolidBrush(trackColor);
                auto* activeBrush = host.GetSolidBrush(activeColor);
                if (trackBrush)
                {
                    DrawRoundedRect(host, trackRect, trackColor, D2D1::ColorF(0, 0, 0, 0), kSliderTrackHeightDip * 0.5f);
                }

                const size_t stopCount = item.sliderStops.size();
                if (activeBrush && stopCount > 0u)
                {
                    const uint32_t sliderValue = ClampSliderValue(item);
                    const float trackCenterY   = (trackRect.top + trackRect.bottom) * 0.5f;
                    const float activeRight =
                        stopCount <= 1u
                            ? trackRect.left
                            : trackRect.left + ((trackRect.right - trackRect.left) * static_cast<float>(sliderValue) / static_cast<float>(stopCount - 1u));
                    if (activeRight > trackRect.left)
                    {
                        const D2D1_RECT_F activeRect = D2D1::RectF(trackRect.left, trackRect.top, activeRight, trackRect.bottom);
                        DrawRoundedRect(host, activeRect, activeColor, D2D1::ColorF(0, 0, 0, 0), kSliderTrackHeightDip * 0.5f);
                    }

                    for (size_t stopIndex = 0u; stopIndex < stopCount; ++stopIndex)
                    {
                        const float t                  = stopCount <= 1u ? 0.0f : static_cast<float>(stopIndex) / static_cast<float>(stopCount - 1u);
                        const float x                  = trackRect.left + ((trackRect.right - trackRect.left) * t);
                        const bool activeStop          = stopIndex == sliderValue;
                        const float radius             = activeStop ? kSliderActiveStopRadiusDip : kSliderStopRadiusDip;
                        const D2D1_ELLIPSE stopEllipse = D2D1::Ellipse(D2D1::Point2F(x, trackCenterY), radius, radius);
                        auto* stopBrush                = host.GetSolidBrush(activeStop ? activeColor : style.background);
                        auto* borderBrush              = host.GetSolidBrush(activeColor);
                        if (stopBrush)
                        {
                            dc->FillEllipse(stopEllipse, stopBrush);
                        }
                        if (borderBrush)
                        {
                            dc->DrawEllipse(stopEllipse, borderBrush, 1.0f);
                        }
                    }
                }

                y += itemHeight;
                continue;
            }

            if (item.kind == MenuItemKind::Info)
            {
                const MenuItemLayoutRects layout  = GetMenuItemLayoutRects(*popup, i);
                const DecodedMenuItemText decoded = DecodeMenuItemText(item);
                const ParsedMenuLabel label       = ParseMenuLabel(decoded.labelText);
                const D2D1_COLOR_F infoLabelColor = decoded.acceleratorText.empty() ? style.text : style.accelText;
                DrawCenteredText(host,
                                 label.displayText,
                                 layout.textRectDip,
                                 FontRole::Body,
                                 infoLabelColor,
                                 DWRITE_TEXT_ALIGNMENT_LEADING,
                                 DWRITE_PARAGRAPH_ALIGNMENT_CENTER);

                if (! decoded.acceleratorText.empty())
                {
                    DrawCenteredText(host,
                                     decoded.acceleratorText,
                                     layout.acceleratorRectDip,
                                     FontRole::Body,
                                     style.text,
                                     DWRITE_TEXT_ALIGNMENT_TRAILING,
                                     DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
                }

                y += itemHeight;
                continue;
            }

            // Standard/Toggle/Radio item
            const MenuItemLayoutRects layout           = GetMenuItemLayoutRects(*popup, i);
            const bool isHovered                       = (popup->hoveredIndex.has_value() && popup->hoveredIndex.value() == i) ||
                                                         (popup->keyboardIndex.has_value() && popup->keyboardIndex.value() == i);
            const bool isDisabled                      = ! item.enabled;
            const DecodedMenuItemText decoded          = DecodeMenuItemText(item);
            const ParsedMenuLabel label                = ParseMenuLabel(decoded.labelText);
            const MenuResolvedItemPaintStyle itemPaint = ResolveMenuItemPaintStyle(theme, style, item, label.displayText, isHovered);
            const D2D1_COLOR_F textColor               = isDisabled ? D2D1::ColorF(itemPaint.text.r, itemPaint.text.g, itemPaint.text.b, 0.4f) : itemPaint.text;
            const D2D1_COLOR_F accelColor =
                isDisabled ? D2D1::ColorF(itemPaint.accelText.r, itemPaint.accelText.g, itemPaint.accelText.b, 0.4f) : itemPaint.accelText;

            // Hover backplate
            if (itemPaint.showHighlightFill)
            {
                const D2D1_RECT_F hoverRect = D2D1::RectF(
                    visibleItemRect.left + kItemHoverInsetDip, visibleItemRect.top, visibleItemRect.right - kItemHoverInsetDip, visibleItemRect.bottom);
                const D2D1_COLOR_F transparent = D2D1::ColorF(0, 0, 0, 0);
                DrawRoundedRect(host, hoverRect, itemPaint.fill, transparent, kItemHoverRadiusDip);
            }

            // Icon area (check/radio indicator or custom icon)
            if (item.checked)
            {
                if (item.kind == MenuItemKind::Toggle)
                {
                    DrawCenteredText(host,
                                     kCheckMarkGlyph,
                                     layout.iconRectDip,
                                     FontRole::Icon,
                                     isDisabled ? D2D1::ColorF(itemPaint.checkColor.r, itemPaint.checkColor.g, itemPaint.checkColor.b, 0.4f)
                                                : itemPaint.checkColor,
                                     DWRITE_TEXT_ALIGNMENT_CENTER,
                                     DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
                }
                else if (item.kind == MenuItemKind::Radio)
                {
                    DrawCenteredText(host,
                                     kRadioBulletGlyph,
                                     layout.iconRectDip,
                                     FontRole::Icon,
                                     isDisabled ? D2D1::ColorF(itemPaint.checkColor.r, itemPaint.checkColor.g, itemPaint.checkColor.b, 0.4f)
                                                : itemPaint.checkColor,
                                     DWRITE_TEXT_ALIGNMENT_CENTER,
                                     DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
                }
            }
            else if (! item.iconGlyph.empty())
            {
                DrawCenteredText(host,
                                 item.iconGlyph,
                                 layout.iconRectDip,
                                 FontRole::Icon,
                                 isDisabled ? D2D1::ColorF(itemPaint.iconColor.r, itemPaint.iconColor.g, itemPaint.iconColor.b, 0.4f) : itemPaint.iconColor,
                                 DWRITE_TEXT_ALIGNMENT_CENTER,
                                 DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
            }
            else if (item.iconBitmap)
            {
                DrawMenuBitmapIcon(host, *item.iconBitmap, layout.iconRectDip, isDisabled ? 0.4f : 1.0f);
            }

            // Item text
            DrawCenteredText(
                host, label.displayText, layout.textRectDip, FontRole::Body, textColor, DWRITE_TEXT_ALIGNMENT_LEADING, DWRITE_PARAGRAPH_ALIGNMENT_CENTER);

            // Accelerator text
            if (! decoded.acceleratorText.empty())
            {
                DrawCenteredText(host,
                                 decoded.acceleratorText,
                                 layout.acceleratorRectDip,
                                 FontRole::Body,
                                 accelColor,
                                 DWRITE_TEXT_ALIGNMENT_TRAILING,
                                 DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
            }

            // Submenu chevron
            if (! item.children.empty())
            {
                DrawCenteredText(host,
                                 kChevronRightGlyph,
                                 layout.chevronRectDip,
                                 FontRole::Icon,
                                 isDisabled ? D2D1::ColorF(itemPaint.chevronColor.r, itemPaint.chevronColor.g, itemPaint.chevronColor.b, 0.4f)
                                            : itemPaint.chevronColor,
                                 DWRITE_TEXT_ALIGNMENT_CENTER,
                                 DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
            }

            y += itemHeight;
        }

        dc->PopAxisAlignedClip();

        if (popup->NeedsScrollbar())
        {
            const auto visuals = ResolveScrollbarVisuals(host.GetTheme(),
                                                         popup->scrollbarHotPart == MenuPopup::ScrollbarHotPart::Track,
                                                         popup->scrollbarHotPart == MenuPopup::ScrollbarHotPart::Thumb,
                                                         popup->draggingScrollbarThumb);
            PaintScrollbar(host, popup->GetScrollbarTrackRect(), popup->GetScrollbarThumbRect(), visuals);
        }
    }
};

// ---------------------------------------------------------------------------
// Popup positioning with screen-edge flip
// ---------------------------------------------------------------------------

[[nodiscard]] RECT GetPopupItemScreenRect(const MenuPopup& popup, size_t itemIndex) noexcept
{
    RECT windowRect{};
    if (! popup.hwnd || GetWindowRect(popup.hwnd, &windowRect) == FALSE)
    {
        return RECT{};
    }

    const D2D1_RECT_F itemRect = GetVisibleItemRect(popup, itemIndex);
    const float scale          = static_cast<float>(popup.dpi) / 96.0f;
    return RECT{RoundToLongSaturated(static_cast<double>(windowRect.left) + (static_cast<double>(itemRect.left) * static_cast<double>(scale))),
                RoundToLongSaturated(static_cast<double>(windowRect.top) + (static_cast<double>(itemRect.top) * static_cast<double>(scale))),
                RoundToLongSaturated(static_cast<double>(windowRect.left) + (static_cast<double>(itemRect.right) * static_cast<double>(scale))),
                RoundToLongSaturated(static_cast<double>(windowRect.top) + (static_cast<double>(itemRect.bottom) * static_cast<double>(scale)))};
}

[[nodiscard]] RECT ComputePopupPosition(POINT screenPoint,
                                        float widthDip,
                                        float heightDip,
                                        UINT dpi,
                                        bool isSubmenu,
                                        ContextMenuRootHorizontalAlignment rootHorizontalAlignment,
                                        ContextMenuRootVerticalPlacement rootVerticalPlacement,
                                        const RECT* parentRect,
                                        const RECT* parentItemRect) noexcept
{
    const float scale         = static_cast<float>(dpi) / 96.0f;
    const int desiredWidthPx  = DipExtentToPixels(widthDip, dpi);
    const int desiredHeightPx = DipExtentToPixels(heightDip, dpi);

    HMONITOR monitor = MonitorFromPoint(screenPoint, MONITOR_DEFAULTTONEAREST);
    MONITORINFO mi{};
    mi.cbSize = sizeof(mi);
    GetMonitorInfoW(monitor, &mi);
    const RECT work        = mi.rcWork;
    const int workWidthPx  = (std::max)(1, static_cast<int>(work.right - work.left));
    const int workHeightPx = (std::max)(1, static_cast<int>(work.bottom - work.top));
    const int widthPx      = (std::min)(desiredWidthPx, workWidthPx);
    const int heightPx     = (std::min)(desiredHeightPx, workHeightPx);

    int x = screenPoint.x;
    int y = screenPoint.y;

    if (isSubmenu && parentItemRect)
    {
        x = parentItemRect->right;
        y = parentItemRect->top + static_cast<int>(kSubmenuVerticalOffsetDip * scale + 0.5f);
    }
    else if (! isSubmenu)
    {
        if (rootHorizontalAlignment == ContextMenuRootHorizontalAlignment::End)
        {
            x -= widthPx;
            if (x < work.left)
            {
                x = screenPoint.x;
            }
        }

        if (rootVerticalPlacement == ContextMenuRootVerticalPlacement::Above)
        {
            y -= heightPx;
            if (y < work.top && screenPoint.y + heightPx <= work.bottom)
            {
                y = screenPoint.y;
            }
            y = (std::max)(y, static_cast<int>(work.top));
        }
    }

    // Flip horizontally if would go off-screen right
    if (x + widthPx > work.right)
    {
        if (isSubmenu)
        {
            const int anchorLeft = parentItemRect ? parentItemRect->left : (parentRect ? parentRect->left : x);
            x                    = anchorLeft - widthPx;
        }
        else
            x = work.right - widthPx;
    }

    // Flip vertically if would go off-screen bottom
    if (y + heightPx > work.bottom)
    {
        y = work.bottom - heightPx;
    }

    // Clamp to work area
    const int maxX = (std::max)(static_cast<int>(work.left), static_cast<int>(work.right - widthPx));
    const int maxY = (std::max)(static_cast<int>(work.top), static_cast<int>(work.bottom - heightPx));
    x              = std::clamp(x, static_cast<int>(work.left), maxX);
    y              = std::clamp(y, static_cast<int>(work.top), maxY);

    return RECT{x, y, x + widthPx, y + heightPx};
}

[[nodiscard]] RECT ComputePopupWindowRect(const RECT& surfaceRectPx, const MenuPopupShadowMargins& shadowMargins, UINT dpi) noexcept
{
    const int shadowLeftPx   = DipExtentToPixels(shadowMargins.leftDip, dpi);
    const int shadowTopPx    = DipExtentToPixels(shadowMargins.topDip, dpi);
    const int shadowRightPx  = DipExtentToPixels(shadowMargins.rightDip, dpi);
    const int shadowBottomPx = DipExtentToPixels(shadowMargins.bottomDip, dpi);

    return RECT{surfaceRectPx.left - shadowLeftPx, surfaceRectPx.top - shadowTopPx, surfaceRectPx.right + shadowRightPx, surfaceRectPx.bottom + shadowBottomPx};
}

[[nodiscard]] wil::unique_hrgn CreateMenuPopupWindowRegion(const MenuPopupShadowMargins& shadowMargins, UINT dpi, int widthPx, int heightPx) noexcept
{
    if (widthPx <= 0 || heightPx <= 0)
    {
        return nullptr;
    }

    const int menuCornerRadiusPx = DipExtentToPixels(kMenuCornerRadiusDip, dpi);
    const int shadowLeftPx       = DipExtentToPixels(shadowMargins.leftDip, dpi);
    const int shadowTopPx        = DipExtentToPixels(shadowMargins.topDip, dpi);
    const int shadowRightPx      = DipExtentToPixels(shadowMargins.rightDip, dpi);
    const int shadowBottomPx     = DipExtentToPixels(shadowMargins.bottomDip, dpi);

    const int topLeftRadiusXPx     = (std::clamp)(menuCornerRadiusPx + shadowLeftPx, 1, widthPx);
    const int topLeftRadiusYPx     = (std::clamp)(menuCornerRadiusPx + shadowTopPx, 1, heightPx);
    const int topRightRadiusXPx    = (std::clamp)(menuCornerRadiusPx + shadowRightPx, 1, widthPx);
    const int topRightRadiusYPx    = (std::clamp)(menuCornerRadiusPx + shadowTopPx, 1, heightPx);
    const int bottomLeftRadiusXPx  = (std::clamp)(menuCornerRadiusPx + shadowLeftPx, 1, widthPx);
    const int bottomLeftRadiusYPx  = (std::clamp)(menuCornerRadiusPx + shadowBottomPx, 1, heightPx);
    const int bottomRightRadiusXPx = (std::clamp)(menuCornerRadiusPx + shadowRightPx, 1, widthPx);
    const int bottomRightRadiusYPx = (std::clamp)(menuCornerRadiusPx + shadowBottomPx, 1, heightPx);

    wil::unique_hrgn region(CreateRectRgn(0, 0, widthPx, heightPx));
    if (! region)
    {
        return nullptr;
    }

    const auto subtractCornerCutout =
        [&](int boxLeft, int boxTop, int boxRight, int boxBottom, int ellipseLeft, int ellipseTop, int ellipseRight, int ellipseBottom) noexcept
    {
        if (boxRight <= boxLeft || boxBottom <= boxTop)
        {
            return true;
        }

        wil::unique_hrgn cornerBox(CreateRectRgn(boxLeft, boxTop, boxRight, boxBottom));
        wil::unique_hrgn ellipse(CreateEllipticRgn(ellipseLeft, ellipseTop, ellipseRight, ellipseBottom));
        wil::unique_hrgn quarter(CreateRectRgn(0, 0, 0, 0));
        wil::unique_hrgn cutout(CreateRectRgn(0, 0, 0, 0));
        if (! cornerBox || ! ellipse || ! quarter || ! cutout)
        {
            return false;
        }

        if (CombineRgn(quarter.get(), cornerBox.get(), ellipse.get(), RGN_AND) == ERROR)
        {
            return false;
        }

        if (CombineRgn(cutout.get(), cornerBox.get(), quarter.get(), RGN_DIFF) == ERROR)
        {
            return false;
        }

        return CombineRgn(region.get(), region.get(), cutout.get(), RGN_DIFF) != ERROR;
    };

    if (! subtractCornerCutout(0, 0, topLeftRadiusXPx, topLeftRadiusYPx, 0, 0, topLeftRadiusXPx * 2, topLeftRadiusYPx * 2) ||
        ! subtractCornerCutout(
            widthPx - topRightRadiusXPx, 0, widthPx, topRightRadiusYPx, widthPx - (topRightRadiusXPx * 2), 0, widthPx, topRightRadiusYPx * 2) ||
        ! subtractCornerCutout(
            0, heightPx - bottomLeftRadiusYPx, bottomLeftRadiusXPx, heightPx, 0, heightPx - (bottomLeftRadiusYPx * 2), bottomLeftRadiusXPx * 2, heightPx) ||
        ! subtractCornerCutout(widthPx - bottomRightRadiusXPx,
                               heightPx - bottomRightRadiusYPx,
                               widthPx,
                               heightPx,
                               widthPx - (bottomRightRadiusXPx * 2),
                               heightPx - (bottomRightRadiusYPx * 2),
                               widthPx,
                               heightPx))
    {
        return nullptr;
    }

    return region;
}

void ApplyMenuPopupWindowRegion(HWND hwnd, const MenuPopupShadowMargins& shadowMargins, UINT dpi, int widthPx, int heightPx) noexcept
{
    if (! hwnd || widthPx <= 0 || heightPx <= 0)
    {
        return;
    }

    wil::unique_hrgn region = CreateMenuPopupWindowRegion(shadowMargins, dpi, widthPx, heightPx);
    if (! region)
    {
        return;
    }

    if (SetWindowRgn(hwnd, region.release(), FALSE) == 0)
    {
        Debug::Warning(L"DxUi::ContextMenu: SetWindowRgn failed for popup host");
    }
}

[[nodiscard]] UINT ResolveMenuPopupMessageDpi(HWND hwnd, UINT msg, WPARAM wp) noexcept
{
    UINT dpi = msg == WM_DPICHANGED ? static_cast<UINT>(HIWORD(wp)) : 0u;
    if (dpi == 0u && hwnd && IsWindow(hwnd) != FALSE)
    {
        dpi = GetDpiForWindow(hwnd);
    }
    if (dpi == 0u)
    {
        dpi = GetDpiForSystem();
    }
    return dpi == 0u ? USER_DEFAULT_SCREEN_DPI : dpi;
}

[[nodiscard]] RECT ComputePopupSurfaceRectFromTopLeft(POINT topLeft, float widthDip, float heightDip, UINT dpi) noexcept
{
    const int widthPx  = DipExtentToPixels(widthDip, dpi);
    const int heightPx = DipExtentToPixels(heightDip, dpi);

    const HMONITOR monitor = MonitorFromPoint(topLeft, MONITOR_DEFAULTTONEAREST);
    MONITORINFO mi{};
    mi.cbSize = sizeof(mi);
    if (monitor == nullptr || GetMonitorInfoW(monitor, &mi) == FALSE)
    {
        mi.rcWork = RECT{GetSystemMetrics(SM_XVIRTUALSCREEN),
                         GetSystemMetrics(SM_YVIRTUALSCREEN),
                         GetSystemMetrics(SM_XVIRTUALSCREEN) + (std::max)(1, GetSystemMetrics(SM_CXVIRTUALSCREEN)),
                         GetSystemMetrics(SM_YVIRTUALSCREEN) + (std::max)(1, GetSystemMetrics(SM_CYVIRTUALSCREEN))};
    }

    const RECT work = mi.rcWork;
    const int maxX  = (std::max)(static_cast<int>(work.left), static_cast<int>(work.right - widthPx));
    const int maxY  = (std::max)(static_cast<int>(work.top), static_cast<int>(work.bottom - heightPx));
    const int x     = std::clamp(static_cast<int>(topLeft.x), static_cast<int>(work.left), maxX);
    const int y     = std::clamp(static_cast<int>(topLeft.y), static_cast<int>(work.top), maxY);
    return RECT{x, y, x + widthPx, y + heightPx};
}

void RefreshMenuPopupBackdrop(MenuPopup& popup, const RECT& surfaceRectPx) noexcept
{
    popup.backdropSnapshot = {};
    popup.usesAppBackdropBlur = false;

    const MenuItemVisualStyle itemStyle     = ResolveMenuVisualStyle(popup.host.GetTheme());
    const MenuSurfaceMaterialStyle material = ResolveMenuSurfaceMaterialStyle(popup.host.GetTheme(), itemStyle);
    if (material.backdropOpacity <= 0.0f || material.backdropBlurDip <= 0.0f)
    {
        return;
    }

    const auto backdropStartedAt = std::chrono::steady_clock::now();
    popup.usesAppBackdropBlur    = CaptureMenuBackdropScreenRegion(surfaceRectPx, popup.backdropSnapshot.capture);
    Debug::Perf::Emit(L"dxui.menu.backdrop_capture_us",
                      popup.isSubmenu ? L"submenu_dpi" : L"root_dpi",
                      Debug::Perf::ElapsedUs(backdropStartedAt),
                      static_cast<uint64_t>(surfaceRectPx.right - surfaceRectPx.left),
                      static_cast<uint64_t>(surfaceRectPx.bottom - surfaceRectPx.top),
                      popup.usesAppBackdropBlur ? S_OK : E_FAIL);
}

void RelayoutMenuPopupForDpi(MenuPopup& popup, UINT dpi, const RECT* suggestedWindowRect) noexcept
{
    if (! popup.hwnd || IsWindow(popup.hwnd) == FALSE || dpi == 0u)
    {
        return;
    }

    const auto startedAt = std::chrono::steady_clock::now();
    popup.dpi           = dpi;

    bool hostDpiHandled = false;
    static_cast<void>(popup.host.HandleMessage(popup.hwnd, WM_DPICHANGED, MAKEWPARAM(static_cast<WORD>(dpi), static_cast<WORD>(dpi)), 0, hostDpiHandled));

    POINT surfaceTopLeft{popup.surfaceRectPx.left, popup.surfaceRectPx.top};
    if (suggestedWindowRect)
    {
        surfaceTopLeft.x = suggestedWindowRect->left + DipExtentToPixels(popup.shadowMargins.leftDip, dpi);
        surfaceTopLeft.y = suggestedWindowRect->top + DipExtentToPixels(popup.shadowMargins.topDip, dpi);
    }

    const D2D1_SIZE_F sizeDip = ComputeMenuSize(popup.items, popup.itemCount, popup.host, &popup.acceleratorColumnWidthDip);
    const float requestedVisibleHeightDip =
        (! popup.isSubmenu && popup.controller && popup.controller->sessionCallbacks.maxRootHeightDip > 0.0f)
            ? (std::min)(sizeDip.height, popup.controller->sessionCallbacks.maxRootHeightDip)
            : sizeDip.height;

    const RECT surfaceRectPx     = ComputePopupSurfaceRectFromTopLeft(surfaceTopLeft, sizeDip.width, requestedVisibleHeightDip, dpi);
    const RECT windowRect        = ComputePopupWindowRect(surfaceRectPx, popup.shadowMargins, dpi);
    const int windowWidthPx      = windowRect.right - windowRect.left;
    const int windowHeightPx     = windowRect.bottom - windowRect.top;
    popup.menuWidthDip           = popup.PixelToDip(static_cast<float>(surfaceRectPx.right - surfaceRectPx.left));
    popup.menuHeightDip          = popup.PixelToDip(static_cast<float>(surfaceRectPx.bottom - surfaceRectPx.top));
    popup.windowWidthDip         = popup.PixelToDip(static_cast<float>(windowWidthPx));
    popup.windowHeightDip        = popup.PixelToDip(static_cast<float>(windowHeightPx));
    popup.surfaceRectPx          = surfaceRectPx;
    popup.windowRectPx           = windowRect;
    popup.contentHeightDip       = sizeDip.height;
    popup.draggingScrollbarThumb = false;
    popup.scrollbarHotPart       = MenuPopup::ScrollbarHotPart::None;
    popup.ClampScrollOffset();

    SetWindowPos(popup.hwnd, HWND_TOP, windowRect.left, windowRect.top, windowWidthPx, windowHeightPx, SWP_NOACTIVATE | SWP_NOOWNERZORDER);
    ApplyMenuPopupWindowRegion(popup.hwnd, popup.shadowMargins, dpi, windowWidthPx, windowHeightPx);
    if (Control* const root = popup.host.GetRoot())
    {
        root->SetBounds(D2D1::RectF(0.0f, 0.0f, popup.windowWidthDip, popup.windowHeightDip));
    }
    RefreshMenuPopupBackdrop(popup, surfaceRectPx);
    popup.host.Invalidate();

    Debug::Perf::Emit(L"dxui.menu.dpi_relayout_us",
                      popup.isSubmenu ? L"submenu" : L"root",
                      Debug::Perf::ElapsedUs(startedAt),
                      static_cast<uint64_t>(popup.itemCount),
                      static_cast<uint64_t>(dpi),
                      S_OK);
}

// ---------------------------------------------------------------------------
// Create/show a menu popup window
// ---------------------------------------------------------------------------

static constexpr UINT_PTR kSubmenuHoverTimerId = 1;

bool CreateMenuPopupWindow(MenuController& controller,
                           const MenuFlyoutItem* items,
                           size_t itemCount,
                           POINT screenPoint,
                           bool isSubmenu,
                           const RECT* parentWindowRect,
                           const RECT* parentItemRect,
                           bool forceInitialRender      = true,
                           bool focusFirstNavigableItem = false)
{
    DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.popup-create-begin",
                              L"owner={:#x} point=({}, {}) items={} submenu={} forceInitialRender={} focusFirst={} maxRootHeightDip={:.1f}",
                              TraceHwndValue(controller.ownerHwnd),
                              screenPoint.x,
                              screenPoint.y,
                              itemCount,
                              TraceBool(isSubmenu),
                              TraceBool(forceInitialRender),
                              TraceBool(focusFirstNavigableItem),
                              controller.sessionCallbacks.maxRootHeightDip);

    auto popup = std::make_unique<MenuPopup>();
    if (items && itemCount > 0u)
    {
        popup->ownedItems.assign(items, items + itemCount);
    }
    popup->items      = popup->ownedItems.data();
    popup->itemCount  = popup->ownedItems.size();
    popup->controller = &controller;
    popup->isSubmenu  = isSubmenu;

    HINSTANCE hInstance = GetModuleHandleW(nullptr);
    EnsureMenuWindowClass(hInstance);

    // DPI from target monitor
    HMONITOR monitor = MonitorFromPoint(screenPoint, MONITOR_DEFAULTTONEAREST);
    UINT dpiX        = USER_DEFAULT_SCREEN_DPI;
    UINT dpiY        = USER_DEFAULT_SCREEN_DPI;
    if (SUCCEEDED(GetDpiForMonitor(monitor, MDT_EFFECTIVE_DPI, &dpiX, &dpiY)))
        popup->dpi = dpiX;
    else
        popup->dpi = GetDpiForSystem();

    // Create a temporary dummy window to initialize WindowHost for text measurement
    // We need the host before we can compute the menu size
    const DWORD exStyle = WS_EX_TOOLWINDOW | WS_EX_NOREDIRECTIONBITMAP | (isSubmenu ? WS_EX_NOACTIVATE : 0u);
    const DWORD style   = WS_POPUP | WS_CLIPCHILDREN | WS_CLIPSIBLINGS;

    HWND hwnd = CreateWindowExW(exStyle, kMenuWindowClass, L"", style, 0, 0, 1, 1, controller.ownerHwnd, nullptr, hInstance, nullptr);

    if (! hwnd)
    {
        DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.popup-create-failed",
                                  L"owner={:#x} stage=CreateWindowEx lastError={}",
                                  TraceHwndValue(controller.ownerHwnd),
                                  static_cast<unsigned long>(GetLastError()));
        return false;
    }

    popup->hwnd = hwnd;

    // Store the popup pointer on the HWND before Attach so the WndProc can find it
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(popup.get()));

    // Attach WindowHost and set theme
    if (! popup->host.Attach(hwnd, WindowHost::AttachOptions{.presentationMode = WindowHost::PresentationMode::CompositionSwapChain}))
    {
        DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.popup-create-failed",
                                  L"owner={:#x} hwnd={:#x} stage=WindowHost.Attach",
                                  TraceHwndValue(controller.ownerHwnd),
                                  TraceHwndValue(hwnd));
        // Controller-initiated destroy: keep the WM_NCDESTROY external-teardown
        // guard from finalizing the async session re-entrantly (see
        // DestroyMenuPopupWindow).
        const bool previousDestroying     = controller.destroyingPopupWindow;
        controller.destroyingPopupWindow  = true;
        DestroyWindow(hwnd);
        controller.destroyingPopupWindow  = previousDestroying;
        return false;
    }
    popup->host.SetTheme(controller.theme);
    // Popup menus stay fully app-rendered even on the transparent composition host.
    // Applying a DWM system backdrop here pulls in OS palette colors that do not match
    // custom RedSalamander themes, so the popup HWND intentionally stays backdrop-free.
    popup->usesSystemBackdrop = false;

    // Compute menu size
    const D2D1_SIZE_F sizeDip = ComputeMenuSize(items, itemCount, popup->host, &popup->acceleratorColumnWidthDip);
    const float requestedVisibleHeightDip =
        (! isSubmenu && controller.sessionCallbacks.maxRootHeightDip > 0.0f)
            ? (std::min)(sizeDip.height, controller.sessionCallbacks.maxRootHeightDip)
            : sizeDip.height;

    // Position the window
    const RECT surfaceRectPx     = ComputePopupPosition(screenPoint,
                                                        sizeDip.width,
                                                        requestedVisibleHeightDip,
                                                        popup->dpi,
                                                        isSubmenu,
                                                        controller.sessionCallbacks.rootHorizontalAlignment,
                                                        controller.sessionCallbacks.rootVerticalPlacement,
                                                        parentWindowRect,
                                                        parentItemRect);
    const RECT windowRect        = ComputePopupWindowRect(surfaceRectPx, popup->shadowMargins, popup->dpi);
    const float visibleWidthDip  = popup->PixelToDip(static_cast<float>(surfaceRectPx.right - surfaceRectPx.left));
    const float visibleHeightDip = popup->PixelToDip(static_cast<float>(surfaceRectPx.bottom - surfaceRectPx.top));
    popup->menuWidthDip          = visibleWidthDip;
    popup->menuHeightDip         = visibleHeightDip;
    popup->windowWidthDip        = popup->PixelToDip(static_cast<float>(windowRect.right - windowRect.left));
    popup->windowHeightDip       = popup->PixelToDip(static_cast<float>(windowRect.bottom - windowRect.top));
    popup->surfaceRectPx         = surfaceRectPx;
    popup->windowRectPx          = windowRect;
    popup->contentHeightDip      = sizeDip.height;
    popup->ClampScrollOffset();

    const MenuItemVisualStyle popupItemStyle     = ResolveMenuVisualStyle(controller.theme);
    const MenuSurfaceMaterialStyle popupMaterial = ResolveMenuSurfaceMaterialStyle(controller.theme, popupItemStyle);
    if (popupMaterial.backdropOpacity > 0.0f && popupMaterial.backdropBlurDip > 0.0f)
    {
        const auto backdropStartedAt = std::chrono::steady_clock::now();
        popup->usesAppBackdropBlur   = CaptureMenuBackdropScreenRegion(surfaceRectPx, popup->backdropSnapshot.capture);
        Debug::Perf::Emit(L"dxui.menu.backdrop_capture_us",
                          isSubmenu ? L"submenu" : L"root",
                          Debug::Perf::ElapsedUs(backdropStartedAt),
                          static_cast<uint64_t>(surfaceRectPx.right - surfaceRectPx.left),
                          static_cast<uint64_t>(surfaceRectPx.bottom - surfaceRectPx.top),
                          popup->usesAppBackdropBlur ? S_OK : E_FAIL);
    }

    // Create the content control
    auto content   = std::make_unique<MenuContentControl>();
    content->popup = popup.get();
    content->SetBounds(D2D1::RectF(0, 0, popup->windowWidthDip, popup->windowHeightDip));
    popup->host.SetRoot(std::move(content));
    popup->keyboardIndex = popup->FindInitialKeyboardItem(focusFirstNavigableItem);
    MenuPopup* const registeredPopup = popup.get();
    controller.popups.push_back(std::move(popup));

    // Resize and show
    const int widthPx  = windowRect.right - windowRect.left;
    const int heightPx = windowRect.bottom - windowRect.top;
    SetWindowPos(hwnd, HWND_TOP, windowRect.left, windowRect.top, widthPx, heightPx, SWP_NOACTIVATE | SWP_NOOWNERZORDER);
    ApplyMenuPopupWindowRegion(hwnd, registeredPopup->shadowMargins, registeredPopup->dpi, widthPx, heightPx);
    [[maybe_unused]] const bool initialFrameReady =
        forceInitialRender ? registeredPopup->host.RenderInitialFrameForShow() : registeredPopup->host.PrimeForShow();
    ShowWindow(hwnd, SW_SHOWNOACTIVATE);

    InvalidateRect(hwnd, nullptr, FALSE);

    DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.popup-create-end",
                               L"owner={:#x} hwnd={:#x} popupIndex={} point=({}, {}) windowRect=({}, {}, {}, {}) surfaceRect=({}, {}, {}, {}) "
                               L"dpi={} items={} keyboardHas={} keyboard={} visibleDip=({:.1f}, {:.1f}) contentHeightDip={:.1f} "
                               L"maxRootHeightDip={:.1f} initialFrameReady={}",
                               TraceHwndValue(controller.ownerHwnd),
                               TraceHwndValue(hwnd),
                              TracePopupIndex(controller, registeredPopup),
                              screenPoint.x,
                              screenPoint.y,
                              windowRect.left,
                              windowRect.top,
                              windowRect.right,
                              windowRect.bottom,
                              surfaceRectPx.left,
                              surfaceRectPx.top,
                              surfaceRectPx.right,
                              surfaceRectPx.bottom,
                              static_cast<unsigned int>(registeredPopup->dpi),
                              registeredPopup->itemCount,
                              TraceBool(registeredPopup->keyboardIndex.has_value()),
                              TraceOptionalSizeValue(registeredPopup->keyboardIndex),
                               registeredPopup->menuWidthDip,
                               registeredPopup->menuHeightDip,
                               registeredPopup->contentHeightDip,
                               controller.sessionCallbacks.maxRootHeightDip,
                               TraceBool(initialFrameReady));

    return true;
}

// ---------------------------------------------------------------------------
// Open a submenu for a given item
// ---------------------------------------------------------------------------

void OpenSubmenu(MenuController& controller, MenuPopup& parent, size_t itemIndex, bool focusFirstNavigableItem)
{
    if (itemIndex >= parent.itemCount)
        return;
    const auto& item = parent.items[itemIndex];
    if (item.children.empty())
        return;

    // Close any existing submenu deeper than this parent
    while (controller.popups.size() > 1 && controller.popups.back().get() != &parent)
    {
        auto& child = controller.popups.back();
        DestroyMenuPopupWindow(*child);
        controller.popups.pop_back();
    }

    // Compute position: right edge of the parent item using the spec-defined cascade offset
    parent.EnsureItemVisible(itemIndex);
    RECT parentWindowRect{};
    GetWindowRect(parent.hwnd, &parentWindowRect);
    const RECT parentItemRect = GetPopupItemScreenRect(parent, itemIndex);
    POINT subPt{parentItemRect.right, parentItemRect.top};

    const auto startedAt = std::chrono::steady_clock::now();
    if (CreateMenuPopupWindow(
            controller, item.children.data(), item.children.size(), subPt, true, &parentWindowRect, &parentItemRect, true, focusFirstNavigableItem))
    {
        if (! controller.popups.empty())
        {
            controller.popups.back()->openedFromItemIndex = itemIndex;
        }

        Debug::Perf::Emit(
            L"DxUI::PopupShow", L"submenu", Debug::Perf::ElapsedUs(startedAt), static_cast<uint64_t>(item.children.size()), static_cast<uint64_t>(itemIndex));
    }
}

void CancelSubmenuHoverTimer(MenuPopup& popup) noexcept
{
    if (! popup.hoverTimerId || ! popup.hwnd)
    {
        popup.hoverTimerId        = 0;
        popup.hoverTimerKind      = MenuPopup::SubmenuHoverTimerKind::None;
        popup.hoverTimerItemIndex = SIZE_MAX;
        return;
    }

    KillTimer(popup.hwnd, kSubmenuHoverTimerId);
    popup.hoverTimerId        = 0;
    popup.hoverTimerKind      = MenuPopup::SubmenuHoverTimerKind::None;
    popup.hoverTimerItemIndex = SIZE_MAX;
}

void DestroyMenuPopupWindow(MenuPopup& popup) noexcept
{
    CancelSubmenuHoverTimer(popup);

    const HWND hwnd = popup.hwnd;
    if (hwnd && IsWindow(hwnd) != FALSE)
    {
        // DestroyWindow delivers WM_NCDESTROY (and possibly WM_CAPTURECHANGED)
        // synchronously while GWLP_USERDATA still points at this popup. Mark the
        // destroy as controller-initiated so the popup WndProc's external-teardown
        // guard does not Dismiss + FinalizeAsyncMenuController re-entrantly, which
        // would delete this popup (and the controller) while callers such as
        // CloseSubmenuChainFrom and CloseTopmostSubmenu are still using them.
        MenuController* const controller = popup.controller;
        const bool previousDestroying    = controller ? controller->destroyingPopupWindow : false;
        if (controller)
        {
            controller->destroyingPopupWindow = true;
        }
        DestroyWindow(hwnd);
        if (controller)
        {
            controller->destroyingPopupWindow = previousDestroying;
        }
    }

    popup.host.Detach();
    popup.hwnd = nullptr;
}

[[nodiscard]] bool HasPendingSubmenuOpenTimer(const MenuPopup& popup) noexcept
{
    return popup.hoverTimerId != 0 && popup.hoverTimerKind == MenuPopup::SubmenuHoverTimerKind::PendingOpen;
}

[[nodiscard]] std::optional<size_t> FindPopupIndex(const MenuController& controller, const MenuPopup& popup) noexcept
{
    for (size_t i = 0; i < controller.popups.size(); ++i)
    {
        if (controller.popups[i].get() == &popup)
        {
            return i;
        }
    }

    return std::nullopt;
}

[[nodiscard]] bool PopupOwnsOpenChildFromItem(const MenuController& controller, size_t popupIndex, size_t itemIndex) noexcept
{
    const size_t childIndex = popupIndex + 1u;
    if (childIndex >= controller.popups.size())
    {
        return false;
    }

    const std::optional<size_t> openedFromItemIndex = controller.popups[childIndex]->openedFromItemIndex;
    return openedFromItemIndex.has_value() && openedFromItemIndex.value() == itemIndex;
}

[[nodiscard]] bool PopupHoverOwnsOpenChild(const MenuController& controller, size_t popupIndex, const MenuPopup& popup) noexcept
{
    return popup.hoveredIndex.has_value() && PopupOwnsOpenChildFromItem(controller, popupIndex, popup.hoveredIndex.value());
}

void ScheduleSubmenuCloseTimer(MenuPopup& popup) noexcept
{
    if (! popup.hwnd)
    {
        popup.hoverTimerId        = 0;
        popup.hoverTimerKind      = MenuPopup::SubmenuHoverTimerKind::None;
        popup.hoverTimerItemIndex = SIZE_MAX;
        return;
    }

    if (popup.hoverTimerId && popup.hoverTimerKind == MenuPopup::SubmenuHoverTimerKind::PendingClose)
    {
        return;
    }

    CancelSubmenuHoverTimer(popup);
    popup.hoverTimerId        = SetTimer(popup.hwnd, kSubmenuHoverTimerId, static_cast<UINT>(kCascadeHoverDelayMs), nullptr);
    popup.hoverTimerKind      = MenuPopup::SubmenuHoverTimerKind::PendingClose;
    popup.hoverTimerItemIndex = SIZE_MAX;
}

void ScheduleSubmenuOpenTimer(MenuPopup& popup, size_t itemIndex) noexcept
{
    if (! popup.hwnd || itemIndex >= popup.itemCount)
    {
        CancelSubmenuHoverTimer(popup);
        return;
    }

    CancelSubmenuHoverTimer(popup);
    popup.hoverTimerId        = SetTimer(popup.hwnd, kSubmenuHoverTimerId, static_cast<UINT>(kCascadeHoverDelayMs), nullptr);
    popup.hoverTimerKind      = MenuPopup::SubmenuHoverTimerKind::PendingOpen;
    popup.hoverTimerItemIndex = itemIndex;
}

void CloseSubmenuChainFrom(MenuController& controller, MenuPopup& parent) noexcept
{
    bool closedAny = false;
    while (controller.popups.size() > 1 && controller.popups.back().get() != &parent)
    {
        auto& child = controller.popups.back();
        DestroyMenuPopupWindow(*child);
        controller.popups.pop_back();
        closedAny = true;
    }

    if (closedAny && parent.hwnd)
    {
        InvalidatePopup(parent);
    }
}

// Fires a pending submenu hover timer on a popup: opens the deferred submenu or
// closes the open submenu chain. Shared by the async WndProc path and the modal
// loop path, which receive the WM_TIMER through different message routes.
void HandleSubmenuHoverTimer(MenuController& controller, MenuPopup& popup)
{
    if (popup.hoverTimerId == 0 || popup.hoverTimerKind == MenuPopup::SubmenuHoverTimerKind::None)
    {
        return;
    }

    if (popup.hoverTimerKind == MenuPopup::SubmenuHoverTimerKind::PendingOpen && popup.hoverTimerItemIndex < popup.itemCount)
    {
        KillTimer(popup.hwnd, kSubmenuHoverTimerId);
        popup.hoverTimerId         = 0;
        popup.hoverTimerKind       = MenuPopup::SubmenuHoverTimerKind::None;
        const size_t openItemIndex = popup.hoverTimerItemIndex;
        popup.hoverTimerItemIndex  = SIZE_MAX;
        OpenSubmenu(controller, popup, openItemIndex, false);
    }
    else if (popup.hoverTimerKind == MenuPopup::SubmenuHoverTimerKind::PendingClose)
    {
        KillTimer(popup.hwnd, kSubmenuHoverTimerId);
        popup.hoverTimerId        = 0;
        popup.hoverTimerKind      = MenuPopup::SubmenuHoverTimerKind::None;
        popup.hoverTimerItemIndex = SIZE_MAX;
        CloseSubmenuChainFrom(controller, popup);
    }
}

// ---------------------------------------------------------------------------
// Process mouse move in a popup
// ---------------------------------------------------------------------------

void HandleMenuMouseMoveAtPointDip(MenuController& controller,
                                   MenuPopup& popup,
                                   D2D1_POINT_2F pointDip,
                                   bool pointerTakesKeyboardFocus = true) noexcept
{
    if (popup.draggingScrollbarThumb)
    {
        const D2D1_RECT_F track = popup.GetScrollbarTrackRect();
        const D2D1_RECT_F thumb = popup.GetScrollbarThumbRect();
        const float thumbHeight = thumb.bottom - thumb.top;
        const float available   = (std::max)(0.0f, (track.bottom - track.top) - thumbHeight);
        if (available > 0.0f)
        {
            const float thumbTop  = (std::clamp)(pointDip.y - popup.scrollbarDragOffsetDip, track.top, track.bottom - thumbHeight);
            popup.scrollOffsetDip = ((thumbTop - track.top) / available) * popup.GetScrollExtent();
            popup.ClampScrollOffset();
            InvalidatePopup(popup);
        }
        return;
    }

    MenuPopup::ScrollbarHotPart nextHotPart = MenuPopup::ScrollbarHotPart::None;
    if (popup.NeedsScrollbar() && PointInRect(popup.GetScrollbarTrackRect(), pointDip))
    {
        nextHotPart = PointInRect(popup.GetScrollbarThumbRect(), pointDip) ? MenuPopup::ScrollbarHotPart::Thumb : MenuPopup::ScrollbarHotPart::Track;
    }
    if (popup.scrollbarHotPart != nextHotPart)
    {
        popup.scrollbarHotPart = nextHotPart;
        InvalidatePopup(popup);
    }

    const std::optional<size_t> hit = HitTestMenuItem(popup, pointDip);
    const bool changed              = (hit != popup.hoveredIndex);
    popup.hoveredIndex              = hit;
    if (pointerTakesKeyboardFocus)
    {
        popup.keyboardIndex.reset();
    }

    if (changed)
    {
        InvalidatePopup(popup);

        // Start/cancel submenu hover timer
        if (popup.hoverTimerId)
        {
            KillTimer(popup.hwnd, kSubmenuHoverTimerId);
            popup.hoverTimerId        = 0;
            popup.hoverTimerKind      = MenuPopup::SubmenuHoverTimerKind::None;
            popup.hoverTimerItemIndex = SIZE_MAX;
        }

        if (hit.has_value())
        {
            const size_t hitIndex = hit.value();
            if (hitIndex < popup.itemCount && ! popup.items[hitIndex].children.empty() && popup.items[hitIndex].enabled)
            {
                const std::optional<size_t> popupIndex = FindPopupIndex(controller, popup);
                if (popupIndex.has_value() && PopupOwnsOpenChildFromItem(controller, popupIndex.value(), hitIndex))
                {
                    CancelSubmenuHoverTimer(popup);
                }
                else
                {
                    ScheduleSubmenuOpenTimer(popup, hitIndex);
                }
            }
            else if (controller.popups.size() > 1 && controller.popups.back().get() != &popup)
            {
                ScheduleSubmenuCloseTimer(popup);
            }
        }
        else if (controller.popups.size() > 1 && controller.popups.back().get() != &popup)
        {
            ScheduleSubmenuCloseTimer(popup);
        }
    }
}

void DestroyPopupChain(MenuController& controller) noexcept
{
    for (auto it = controller.popups.rbegin(); it != controller.popups.rend(); ++it)
    {
        auto& popup = *it;
        DestroyMenuPopupWindow(*popup);
    }
    controller.popups.clear();
}

void ActivatePopupForKeyboard(MenuPopup& popup) noexcept
{
    if (! popup.hwnd || IsWindow(popup.hwnd) == FALSE)
    {
        return;
    }

    SetWindowPos(popup.hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOOWNERZORDER);
    SetActiveWindow(popup.hwnd);
    SetFocus(popup.hwnd);
}

std::vector<std::unique_ptr<MenuController>>& ActiveAsyncMenuControllers() noexcept
{
    thread_local std::vector<std::unique_ptr<MenuController>> controllers;
    return controllers;
}

[[nodiscard]] bool BeginAsyncMenuInteraction(MenuController& controller) noexcept
{
    MenuPopup* root = controller.GetRootPopup();
    if (! root || ! root->hwnd)
    {
        return false;
    }

    controller.previousCapture            = GetCapture();
    controller.previousFocus              = GetFocus();
    controller.ignoreInitialLeftButtonUp  = controller.sessionCallbacks.ignoreInitialLeftButtonUp || (GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0;
    controller.ignoreInitialRightButtonUp = controller.sessionCallbacks.ignoreInitialRightButtonUp || (GetAsyncKeyState(VK_RBUTTON) & 0x8000) != 0;
    controller.leftButtonDownInPopup      = false;
    controller.rightButtonDownInPopup     = false;

    SetCapture(root->hwnd);
    ActivatePopupForKeyboard(*root);
    controller.asyncInteractionActive = true;
    DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.async-start",
                              L"owner={:#x} root={:#x} previousCapture={:#x} previousFocus={:#x} currentCapture={:#x} popupCount={} items={} "
                              L"ignoreLeft={} ignoreRight={}",
                              TraceHwndValue(controller.ownerHwnd),
                              TraceHwndValue(root->hwnd),
                              TraceHwndValue(controller.previousCapture),
                              TraceHwndValue(controller.previousFocus),
                              TraceHwndValue(GetCapture()),
                              controller.popups.size(),
                              controller.rootItems.size(),
                              TraceBool(controller.ignoreInitialLeftButtonUp),
                              TraceBool(controller.ignoreInitialRightButtonUp));
    return true;
}

void EndAsyncMenuInteraction(MenuController& controller) noexcept
{
    controller.asyncInteractionActive = false;
    if (const HWND captured = GetCapture(); captured && controller.FindPopupForHwnd(captured))
    {
        DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.async-release-capture",
                                  L"captured={:#x} owner={:#x}",
                                  TraceHwndValue(captured),
                                  TraceHwndValue(controller.ownerHwnd));
        ReleaseCapture();
    }

    const HWND currentFocus = GetFocus();
    if (controller.previousFocus && IsWindow(controller.previousFocus) != FALSE && currentFocus && controller.FindPopupForHwnd(currentFocus))
    {
        SetFocus(controller.previousFocus);
        DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.async-restore-focus",
                                  L"previousFocus={:#x} focusAfter={:#x}",
                                  TraceHwndValue(controller.previousFocus),
                                  TraceHwndValue(GetFocus()));
    }
}

void FinalizeAsyncMenuController(MenuController& controller) noexcept
{
    if (! controller.asyncSession || controller.asyncFinalizing)
    {
        return;
    }

    controller.asyncFinalizing = true;

    std::unique_ptr<MenuController> ownedController;
    auto& controllers = ActiveAsyncMenuControllers();
    const auto it     = std::find_if(controllers.begin(), controllers.end(), [&controller](const std::unique_ptr<MenuController>& candidate) noexcept {
        return candidate.get() == &controller;
    });
    if (it != controllers.end())
    {
        ownedController = std::move(*it);
        controllers.erase(it);
    }

    MenuController& target               = ownedController ? *ownedController : controller;
    const std::optional<int> result      = target.result;
    ContextMenuClosedCallback onClosed   = std::move(target.asyncOnClosed);
    EndAsyncMenuInteraction(target);
    DestroyPopupChain(target);

    DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.async-end",
                              L"owner={:#x} result={} captureAfter={:#x} focusAfter={:#x}",
                              TraceHwndValue(target.ownerHwnd),
                              result.value_or(-1),
                              TraceHwndValue(GetCapture()),
                              TraceHwndValue(GetFocus()));

    if (onClosed)
    {
        onClosed(result);
    }
}

[[nodiscard]] bool SwitchRootPopup(MenuController& controller,
                                   ContextMenuRootSwitchRequest request,
                                   std::wstring_view popupDetail,
                                   bool focusFirstNavigableItem) noexcept
{
    DXUI_MENU_TRACE(L"DxUi::MenuTrace Popup switch-begin owner={:#x} capture={:#x} detail='{}' oldPopupCount={} newItems={} point=({}, {}) focusFirst={}",
                    reinterpret_cast<uintptr_t>(controller.ownerHwnd),
                    reinterpret_cast<uintptr_t>(GetCapture()),
                    popupDetail,
                    controller.popups.size(),
                    request.items.size(),
                    request.screenPoint.x,
                    request.screenPoint.y,
                    focusFirstNavigableItem ? L"true" : L"false");
    DestroyPopupChain(controller);

    controller.rootItems = std::move(request.items);
    if (controller.rootItems.empty())
    {
        DXUI_MENU_TRACE(L"DxUi::MenuTrace Popup switch-dismiss-empty owner={:#x}", reinterpret_cast<uintptr_t>(controller.ownerHwnd));
        controller.Dismiss();
        return false;
    }

    const auto startedAt = std::chrono::steady_clock::now();
    // Root switches repaint before returning to the modal loop so fast hover/keyboard
    // changes never expose a blank popup between the old root and new root.
    if (! CreateMenuPopupWindow(
            controller, controller.rootItems.data(), controller.rootItems.size(), request.screenPoint, false, nullptr, nullptr, true, focusFirstNavigableItem))
    {
        DXUI_MENU_TRACE(L"DxUi::MenuTrace Popup switch-create-failed owner={:#x}", reinterpret_cast<uintptr_t>(controller.ownerHwnd));
        controller.Dismiss();
        return false;
    }

    if (MenuPopup* root = controller.GetRootPopup(); root && root->hwnd)
    {
#if defined(ENABLE_TESTS)
        if (root->host.DebugGetRenderCount() > 0u)
        {
            ++controller.rootSwitchImmediateRenderCount;
        }
#endif
        const HWND previousCapture = GetCapture();
        SetCapture(root->hwnd);
        ActivatePopupForKeyboard(*root);
        DXUI_MENU_TRACE(L"DxUi::MenuTrace Popup switch-capture root={:#x} previousCapture={:#x} currentCapture={:#x}",
                        reinterpret_cast<uintptr_t>(root->hwnd),
                        reinterpret_cast<uintptr_t>(previousCapture),
                        reinterpret_cast<uintptr_t>(GetCapture()));
    }
    Debug::Perf::Emit(L"DxUI::PopupShow", popupDetail, Debug::Perf::ElapsedUs(startedAt), static_cast<uint64_t>(controller.rootItems.size()), 0u);
    return true;
}

[[nodiscard]] MenuPopup* FindPointerTargetPopup(MenuController& controller, POINT screenPt) noexcept
{
    for (auto it = controller.popups.rbegin(); it != controller.popups.rend(); ++it)
    {
        const RECT wr = (*it)->GetInteractiveScreenRect();
        if (PtInRect(&wr, screenPt))
        {
            return it->get();
        }
    }

    for (auto& popup : controller.popups)
    {
        if (popup->draggingScrollbarThumb)
        {
            return popup.get();
        }
    }

    return nullptr;
}

[[nodiscard]] bool IsPointInAnyPopupInteractiveRect(const MenuController& controller, POINT screenPt) noexcept
{
    for (const auto& popup : controller.popups)
    {
        const RECT interactiveRect = popup->GetInteractiveScreenRect();
        if (PtInRect(&interactiveRect, screenPt))
        {
            return true;
        }
    }

    return false;
}

[[nodiscard]] std::optional<D2D1_POINT_2F> ScreenPointToPopupDip(const MenuPopup& popup, POINT screenPt) noexcept
{
    RECT wr{};
    if (! popup.hwnd || GetWindowRect(popup.hwnd, &wr) == FALSE)
    {
        return std::nullopt;
    }

    return D2D1::Point2F(popup.PixelToDip(static_cast<float>(screenPt.x - wr.left)), popup.PixelToDip(static_cast<float>(screenPt.y - wr.top)));
}

[[nodiscard]] D2D1_POINT_2F PopupClientPointToDip(const MenuPopup& popup, POINT clientPoint) noexcept
{
    return D2D1::Point2F(popup.PixelToDip(static_cast<float>(clientPoint.x)), popup.PixelToDip(static_cast<float>(clientPoint.y)));
}

[[nodiscard]] bool IsDeliveredPopupSurfacePoint(const MenuPopup& popup, const MenuPointerEvent& event) noexcept
{
    return event.source == MenuInputSource::PopupWndProc && event.hwnd == popup.hwnd && event.deliveredPopupPointDip.has_value() &&
           PointInRect(popup.GetSurfaceRect(), event.deliveredPopupPointDip.value());
}

[[nodiscard]] MenuPopup* FindPointerTargetPopup(MenuController& controller, const MenuPointerEvent& event) noexcept
{
    if (MenuPopup* popup = controller.FindPopupForHwnd(event.hwnd); popup && IsDeliveredPopupSurfacePoint(*popup, event))
    {
        return popup;
    }

    return FindPointerTargetPopup(controller, event.screenPoint);
}

[[nodiscard]] std::optional<D2D1_POINT_2F> ResolvePointerPopupPointDip(const MenuPopup& popup, const MenuPointerEvent& event) noexcept
{
    if (event.hwnd == popup.hwnd && event.deliveredPopupPointDip.has_value())
    {
        return event.deliveredPopupPointDip;
    }

    return ScreenPointToPopupDip(popup, event.screenPoint);
}

struct MenuPointerTargetTrace
{
    int popupIndex = -1;
    uintptr_t popupHwnd = 0u;
    std::optional<size_t> hitIndex;
    std::optional<size_t> hoveredIndex;
    std::optional<size_t> keyboardIndex;
};

[[nodiscard]] MenuPointerTargetTrace CaptureMenuPointerTargetTrace(MenuController& controller, const MenuPointerEvent& event) noexcept
{
    MenuPointerTargetTrace trace;
    MenuPopup* popup = FindPointerTargetPopup(controller, event);
    if (! popup)
    {
        return trace;
    }

    trace.popupIndex    = TracePopupIndex(controller, popup);
    trace.popupHwnd     = TraceHwndValue(popup->hwnd);
    trace.hoveredIndex  = popup->hoveredIndex;
    trace.keyboardIndex = popup->keyboardIndex;
    const std::optional<D2D1_POINT_2F> pointDip = ResolvePointerPopupPointDip(*popup, event);
    if (pointDip.has_value())
    {
        trace.hitIndex = HitTestMenuItem(*popup, pointDip.value());
    }
    return trace;
}

void TraceMenuPointerRoute(
    MenuController& controller, const MenuPointerEvent& event, const MenuPointerTargetTrace& before, MenuInputDisposition disposition) noexcept
{
    static_cast<void>(before);
    static_cast<void>(disposition);

    const MenuPointerTargetTrace after = CaptureMenuPointerTargetTrace(controller, event);
    DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.pointer",
                              L"source={} kind={} hwnd={:#x} point=({}, {}) messageTime={} keyboardRootTimeHas={} keyboardRootTime={} deliveredPopupPointHas={} deliveredPopupPoint=({:.1f}, {:.1f}) "
                              L"hitWindow={:#x} mayInvoke={} mayDismiss={} maySwitchRoot={} "
                              L"mayKeyboardFocus={} beforePopup={} beforePopupHwnd={:#x} beforeHitHas={} beforeHit={} beforeHoverHas={} "
                              L"beforeHover={} beforeKeyboardHas={} beforeKeyboard={} afterPopup={} afterPopupHwnd={:#x} afterHitHas={} "
                              L"afterHit={} afterHoverHas={} afterHover={} afterKeyboardHas={} afterKeyboard={} popups={} running={} result={} disposition={}",
                              TraceInputSourceName(event.source),
                              TracePointerKindName(event.kind),
                              TraceHwndValue(event.hwnd),
                              event.screenPoint.x,
                              event.screenPoint.y,
                              static_cast<unsigned int>(event.messageTime),
                              TraceBool(controller.hasLastKeyboardRootSwitchMessageTime),
                              static_cast<unsigned int>(controller.lastKeyboardRootSwitchMessageTime),
                              TraceBool(event.deliveredPopupPointDip.has_value()),
                              event.deliveredPopupPointDip.has_value() ? event.deliveredPopupPointDip->x : -1.0f,
                              event.deliveredPopupPointDip.has_value() ? event.deliveredPopupPointDip->y : -1.0f,
                              TraceHwndValue(WindowFromPoint(event.screenPoint)),
                              TraceBool(event.mayInvoke),
                              TraceBool(event.mayDismiss),
                              TraceBool(event.maySwitchRoot),
                              TraceBool(event.mayTakeKeyboardFocus),
                              before.popupIndex,
                              before.popupHwnd,
                              TraceBool(before.hitIndex.has_value()),
                              TraceOptionalSizeValue(before.hitIndex),
                              TraceBool(before.hoveredIndex.has_value()),
                              TraceOptionalSizeValue(before.hoveredIndex),
                              TraceBool(before.keyboardIndex.has_value()),
                              TraceOptionalSizeValue(before.keyboardIndex),
                              after.popupIndex,
                              after.popupHwnd,
                              TraceBool(after.hitIndex.has_value()),
                              TraceOptionalSizeValue(after.hitIndex),
                              TraceBool(after.hoveredIndex.has_value()),
                              TraceOptionalSizeValue(after.hoveredIndex),
                              TraceBool(after.keyboardIndex.has_value()),
                              TraceOptionalSizeValue(after.keyboardIndex),
                              controller.popups.size(),
                              TraceBool(controller.running),
                              controller.result.value_or(-1),
                              TraceDispositionName(disposition));
}

[[nodiscard]] MenuInputDisposition RouteMenuPointerHover(MenuController& controller, const MenuPointerEvent& event) noexcept
{
    MenuPopup* targetPopup = FindPointerTargetPopup(controller, event);
    if (! targetPopup && ! event.maySwitchRoot)
    {
        return MenuInputDisposition::Ignored;
    }

    POINT routedRootSwitchPoint = event.screenPoint;

    bool pointerMoved = false;
    if (! targetPopup && event.maySwitchRoot)
    {
        pointerMoved = ! controller.hasLastRootSwitchPointerScreenPoint || routedRootSwitchPoint.x != controller.lastRootSwitchPointerScreenPoint.x ||
                       routedRootSwitchPoint.y != controller.lastRootSwitchPointerScreenPoint.y;
    }
    else
    {
        pointerMoved = ! controller.hasLastPointerScreenPoint || event.screenPoint.x != controller.lastPointerScreenPoint.x ||
                       event.screenPoint.y != controller.lastPointerScreenPoint.y;
        controller.lastPointerScreenPoint    = event.screenPoint;
        controller.hasLastPointerScreenPoint = true;
    }

    DXUI_MENU_TRACE(
        L"DxUi::MenuTrace Popup route-pointer source={} kind=move hwnd={:#x} capture={:#x} screen=({}, {}) moved={} targetPopup={} popupCount={}",
        static_cast<int>(event.source),
        reinterpret_cast<uintptr_t>(event.hwnd),
        reinterpret_cast<uintptr_t>(GetCapture()),
        event.screenPoint.x,
        event.screenPoint.y,
        pointerMoved ? L"true" : L"false",
        TracePopupIndex(controller, targetPopup),
        controller.popups.size());

    const bool initializeStationaryPopupHover =
        ! pointerMoved && targetPopup && ! targetPopup->hoveredIndex.has_value() && ! targetPopup->keyboardIndex.has_value();
    if (! pointerMoved && ! initializeStationaryPopupHover)
    {
        return MenuInputDisposition::Ignored;
    }

    if (! targetPopup && event.maySwitchRoot && controller.sessionCallbacks.switchRootFromPointer)
    {
        if (IsStaleRootSwitchPointerMoveAfterKeyboard(controller, event))
        {
            if (ShouldRememberRejectedRootSwitchPointer(event))
            {
                RememberRootSwitchPointerPoint(controller, routedRootSwitchPoint);
            }
            DXUI_MENU_TRACE(L"DxUi::MenuTrace Popup root-switch-stale-after-keyboard point=({}, {}) messageTime={} keyboardTime={}",
                            routedRootSwitchPoint.x,
                            routedRootSwitchPoint.y,
                            static_cast<unsigned int>(event.messageTime),
                            static_cast<unsigned int>(controller.lastKeyboardRootSwitchMessageTime));
            return MenuInputDisposition::Consumed;
        }
        if (IsPointInAnyPopupInteractiveRect(controller, routedRootSwitchPoint))
        {
            return MenuInputDisposition::Consumed;
        }
        DXUI_MENU_TRACE(L"DxUi::MenuTrace Popup root-switch-probe point=({}, {}) targetPopup={} activeItems={}",
                    routedRootSwitchPoint.x,
                    routedRootSwitchPoint.y,
                    TracePopupIndex(controller, targetPopup),
                    controller.rootItems.size());
        if (auto request = controller.sessionCallbacks.switchRootFromPointer(routedRootSwitchPoint); request.has_value())
        {
            DXUI_MENU_TRACE(L"DxUi::MenuTrace Popup root-switch-accepted point=({}, {}) newItems={} newPoint=({}, {})",
                        routedRootSwitchPoint.x,
                        routedRootSwitchPoint.y,
                        request->items.size(),
                        request->screenPoint.x,
                        request->screenPoint.y);
#if defined(ENABLE_TESTS)
            ++controller.rootPointerSwitchCount;
#endif
            RememberRootSwitchPointerPoint(controller, routedRootSwitchPoint);
            ClearKeyboardRootSwitchMessageTime(controller);
            static_cast<void>(SwitchRootPopup(controller, std::move(request.value()), L"root-switch-pointer", false));
            return MenuInputDisposition::SwitchedRoot;
        }
        DXUI_MENU_TRACE(L"DxUi::MenuTrace Popup root-switch-rejected point=({}, {})", routedRootSwitchPoint.x, routedRootSwitchPoint.y);
        if (ShouldRememberRejectedRootSwitchPointer(event))
        {
            RememberRootSwitchPointerPoint(controller, routedRootSwitchPoint);
        }
    }

    if (targetPopup)
    {
        if (const std::optional<D2D1_POINT_2F> pointDip = ResolvePointerPopupPointDip(*targetPopup, event); pointDip.has_value())
        {
            HandleMenuMouseMoveAtPointDip(controller, *targetPopup, pointDip.value(), event.mayTakeKeyboardFocus);
        }
        size_t targetPopupIndex = controller.popups.size();
        for (size_t i = 0; i < controller.popups.size(); ++i)
        {
            if (controller.popups[i].get() == targetPopup)
            {
                targetPopupIndex = i;
                break;
            }
        }

        for (size_t i = 0; i < targetPopupIndex && i < controller.popups.size(); ++i)
        {
            CancelSubmenuHoverTimer(*controller.popups[i]);
        }

        if (targetPopupIndex + 1u < controller.popups.size())
        {
            if (PopupHoverOwnsOpenChild(controller, targetPopupIndex, *targetPopup))
            {
                CancelSubmenuHoverTimer(*targetPopup);
            }
            else if (! HasPendingSubmenuOpenTimer(*targetPopup))
            {
                ScheduleSubmenuCloseTimer(*targetPopup);
            }
        }

        for (size_t i = 0; i < controller.popups.size(); ++i)
        {
            auto& popup = controller.popups[i];
            if (popup.get() == targetPopup)
            {
                continue;
            }

            bool invalidate = false;
            if (i < targetPopupIndex)
            {
                const std::optional<size_t> cascadeHighlight = controller.popups[i + 1u]->openedFromItemIndex;
                if (popup->hoveredIndex != cascadeHighlight)
                {
                    popup->hoveredIndex = cascadeHighlight;
                    invalidate          = true;
                }
                if (event.mayTakeKeyboardFocus && popup->keyboardIndex.has_value())
                {
                    popup->keyboardIndex.reset();
                    invalidate = true;
                }
            }
            else
            {
                if (popup->hoveredIndex.has_value())
                {
                    popup->hoveredIndex.reset();
                    invalidate = true;
                }
                if (event.mayTakeKeyboardFocus && popup->keyboardIndex.has_value())
                {
                    popup->keyboardIndex.reset();
                    invalidate = true;
                }
            }

            if (invalidate && popup->hwnd)
            {
                InvalidatePopup(*popup);
            }
        }

        return MenuInputDisposition::HoverChanged;
    }

    for (auto& popup : controller.popups)
    {
        if (popup->hoveredIndex.has_value())
        {
            popup->hoveredIndex.reset();
            InvalidatePopup(*popup);
        }
    }

    if (controller.popups.size() > 1)
    {
        for (size_t i = 1; i < controller.popups.size(); ++i)
        {
            CancelSubmenuHoverTimer(*controller.popups[i]);
        }
        ScheduleSubmenuCloseTimer(*controller.popups.front());
    }
    else if (! controller.popups.empty())
    {
        CancelSubmenuHoverTimer(*controller.popups.front());
    }

    return MenuInputDisposition::Consumed;
}

[[nodiscard]] MenuInputDisposition RouteMenuPointerEvent(MenuController& controller, const MenuPointerEvent& event) noexcept
{
    const bool traceEnabled = IsMenuDiagnosticsTraceEnabled();
    const MenuPointerTargetTrace traceBefore = traceEnabled ? CaptureMenuPointerTargetTrace(controller, event) : MenuPointerTargetTrace{};
    const auto finish = [&](MenuInputDisposition disposition) noexcept
    {
        if (traceEnabled)
        {
            TraceMenuPointerRoute(controller, event, traceBefore, disposition);
        }
        return disposition;
    };

    if (! controller.running)
    {
        return finish(MenuInputDisposition::Ignored);
    }

    if (event.kind == MenuPointerKind::Move)
    {
        return finish(RouteMenuPointerHover(controller, event));
    }

    MenuPopup* targetPopup = FindPointerTargetPopup(controller, event);
    if (event.kind == MenuPointerKind::Wheel)
    {
        if (targetPopup && targetPopup->NeedsScrollbar())
        {
            const short wheelDelta = GET_WHEEL_DELTA_WPARAM(event.wParam);
            const float steps      = static_cast<float>(wheelDelta) / static_cast<float>(WHEEL_DELTA);
            targetPopup->scrollOffsetDip -= steps * ResolveMenuItemHeightDip(targetPopup->host.GetTheme());
            targetPopup->ClampScrollOffset();
            InvalidatePopup(*targetPopup);
            return finish(MenuInputDisposition::Consumed);
        }

        return finish(targetPopup ? MenuInputDisposition::Consumed : MenuInputDisposition::Ignored);
    }

    if (event.kind == MenuPointerKind::LeftDown || event.kind == MenuPointerKind::RightDown)
    {
        if (! targetPopup)
        {
            if (event.mayDismiss)
            {
                controller.Dismiss();
                return finish(MenuInputDisposition::Dismissed);
            }
            return finish(MenuInputDisposition::Ignored);
        }

        if (event.kind == MenuPointerKind::LeftDown)
        {
            controller.leftButtonDownInPopup = true;
        }
        else
        {
            controller.rightButtonDownInPopup = true;
        }

        const std::optional<D2D1_POINT_2F> pointDip = ResolvePointerPopupPointDip(*targetPopup, event);
        if (pointDip.has_value())
        {
            HandleMenuMouseMoveAtPointDip(controller, *targetPopup, pointDip.value(), event.mayTakeKeyboardFocus);
        }
        if (event.kind == MenuPointerKind::LeftDown)
        {
            if (pointDip.has_value() && targetPopup->NeedsScrollbar() && PointInRect(targetPopup->GetScrollbarTrackRect(), pointDip.value()))
            {
                const D2D1_RECT_F thumb = targetPopup->GetScrollbarThumbRect();
                if (PointInRect(thumb, pointDip.value()))
                {
                    targetPopup->draggingScrollbarThumb = true;
                    targetPopup->scrollbarDragOffsetDip = pointDip->y - thumb.top;
                    targetPopup->scrollbarHotPart       = MenuPopup::ScrollbarHotPart::Thumb;
                }
                else
                {
                    const float pageStep = targetPopup->menuHeightDip * 0.8f;
                    targetPopup->scrollOffsetDip += pointDip->y < thumb.top ? -pageStep : pageStep;
                    targetPopup->ClampScrollOffset();
                    targetPopup->scrollbarHotPart = MenuPopup::ScrollbarHotPart::Track;
                }
                InvalidatePopup(*targetPopup);
            }
        }
        return finish(MenuInputDisposition::Consumed);
    }

    if (event.kind == MenuPointerKind::LeftUp || event.kind == MenuPointerKind::RightUp)
    {
        const bool isLeftButton      = event.kind == MenuPointerKind::LeftUp;
        const bool buttonDownInPopup = isLeftButton ? controller.leftButtonDownInPopup : controller.rightButtonDownInPopup;
        if (isLeftButton && controller.ignoreInitialLeftButtonUp && ! buttonDownInPopup)
        {
            controller.ignoreInitialLeftButtonUp = false;
            return finish(MenuInputDisposition::Consumed);
        }
        if (! isLeftButton && controller.ignoreInitialRightButtonUp && ! buttonDownInPopup)
        {
            controller.ignoreInitialRightButtonUp = false;
            return finish(MenuInputDisposition::Consumed);
        }

        if (isLeftButton)
        {
            controller.leftButtonDownInPopup = false;
        }
        else
        {
            controller.rightButtonDownInPopup = false;
        }

        if (targetPopup && targetPopup->draggingScrollbarThumb)
        {
            targetPopup->draggingScrollbarThumb = false;
            return finish(MenuInputDisposition::Consumed);
        }

        std::optional<size_t> targetIndex = targetPopup ? targetPopup->hoveredIndex : std::nullopt;
        if (targetPopup && ! targetIndex.has_value())
        {
            const std::optional<D2D1_POINT_2F> pointDip = ResolvePointerPopupPointDip(*targetPopup, event);
            if (pointDip.has_value())
            {
                targetIndex = HitTestMenuItem(*targetPopup, pointDip.value());
                if (targetIndex != targetPopup->hoveredIndex)
                {
                    targetPopup->hoveredIndex = targetIndex;
                    if (event.mayTakeKeyboardFocus)
                    {
                        targetPopup->keyboardIndex.reset();
                    }
                    InvalidatePopup(*targetPopup);
                }
            }
        }

        if (targetPopup && targetIndex.has_value())
        {
            const size_t idx = targetIndex.value();
            const auto& item = targetPopup->items[idx];
            if (item.kind == MenuItemKind::Slider && item.enabled)
            {
                const std::optional<D2D1_POINT_2F> pointDip = ResolvePointerPopupPointDip(*targetPopup, event);
                if (pointDip.has_value())
                {
                    const D2D1_RECT_F itemRect = GetVisibleItemRect(*targetPopup, idx);
                    const uint32_t stopIndex   = HitTestSliderStop(item, itemRect, pointDip.value());
                    const int commandId        = GetSliderCommandId(item, stopIndex);
                    if (commandId != 0)
                    {
                        controller.InvokeItem(commandId);
                        return finish(MenuInputDisposition::Invoked);
                    }
                }
            }
            else if (! item.children.empty() && item.enabled)
            {
                OpenSubmenu(controller, *targetPopup, idx, false);
            }
            else if (event.mayInvoke && IsInvokableMenuItem(item))
            {
                controller.InvokeItem(item.commandId);
                return finish(MenuInputDisposition::Invoked);
            }
        }
        else if (! targetPopup && event.mayDismiss)
        {
            controller.Dismiss();
            return finish(MenuInputDisposition::Dismissed);
        }

        return finish(MenuInputDisposition::Consumed);
    }

    return finish(MenuInputDisposition::Ignored);
}

[[nodiscard]] std::optional<MenuPointerKind> MenuPointerKindFromMessage(UINT message) noexcept
{
    switch (message)
    {
        case WM_MOUSEMOVE: return MenuPointerKind::Move;
        case WM_LBUTTONDOWN: return MenuPointerKind::LeftDown;
        case WM_LBUTTONUP: return MenuPointerKind::LeftUp;
        case WM_RBUTTONDOWN: return MenuPointerKind::RightDown;
        case WM_RBUTTONUP: return MenuPointerKind::RightUp;
        case WM_MOUSEWHEEL:
        case WM_MOUSEHWHEEL: return MenuPointerKind::Wheel;
        default: return std::nullopt;
    }
}

[[nodiscard]] MenuInputDisposition RouteMenuKeyboardEvent(MenuController& controller, const MenuKeyboardEvent& event) noexcept
{
    MenuPopup* topmost = controller.GetTopmostPopup();
    if (! topmost || ! controller.running)
    {
        return MenuInputDisposition::Ignored;
    }

    const auto focusedIndex = [&]() noexcept -> std::optional<size_t>
    {
        auto idx = topmost->keyboardIndex.has_value() ? topmost->keyboardIndex : topmost->hoveredIndex;
        if (! idx.has_value() || idx.value() >= topmost->itemCount)
        {
            return std::nullopt;
        }
        return idx;
    };
    const auto invokeFocusedSliderStop = [&](uint32_t stopIndex) -> bool
    {
        const auto idx = focusedIndex();
        if (! idx.has_value())
        {
            return false;
        }
        const auto& item = topmost->items[idx.value()];
        if (item.kind != MenuItemKind::Slider || ! item.enabled)
        {
            return false;
        }
        const int commandId = GetSliderCommandId(item, stopIndex);
        if (commandId == 0)
        {
            return true;
        }
        controller.InvokeItem(commandId);
        return true;
    };
    const auto invokeFocusedSliderDelta = [&](int delta) -> bool
    {
        const auto idx = focusedIndex();
        if (! idx.has_value())
        {
            return false;
        }
        const auto& item = topmost->items[idx.value()];
        if (item.kind != MenuItemKind::Slider || ! item.enabled || item.sliderStops.empty())
        {
            return false;
        }
        const int current = static_cast<int>(ClampSliderValue(item));
        const int maxStop = static_cast<int>(item.sliderStops.size() - 1u);
        return invokeFocusedSliderStop(static_cast<uint32_t>(std::clamp(current + delta, 0, maxStop)));
    };

    switch (event.kind)
    {
        case MenuKeyboardKind::Up:
        {
            auto cur  = topmost->keyboardIndex.value_or(topmost->hoveredIndex.value_or(SIZE_MAX));
            auto next = (cur == SIZE_MAX) ? topmost->FindLastNavigableItem() : topmost->FindNextNavigableItem(cur, false);
            if (next.has_value())
            {
                topmost->keyboardIndex = next;
                topmost->hoveredIndex.reset();
                topmost->EnsureItemVisible(next.value());
                InvalidatePopup(*topmost);
            }
            return MenuInputDisposition::Consumed;
        }
        case MenuKeyboardKind::Down:
        {
            auto cur  = topmost->keyboardIndex.value_or(topmost->hoveredIndex.value_or(SIZE_MAX));
            auto next = (cur == SIZE_MAX) ? topmost->FindFirstNavigableItem() : topmost->FindNextNavigableItem(cur, true);
            if (next.has_value())
            {
                topmost->keyboardIndex = next;
                topmost->hoveredIndex.reset();
                topmost->EnsureItemVisible(next.value());
                InvalidatePopup(*topmost);
            }
            return MenuInputDisposition::Consumed;
        }
        case MenuKeyboardKind::Home:
        {
            if (invokeFocusedSliderStop(0u))
            {
                return controller.running ? MenuInputDisposition::Consumed : MenuInputDisposition::Invoked;
            }
            auto first = topmost->FindFirstNavigableItem();
            if (first.has_value())
            {
                topmost->keyboardIndex = first;
                topmost->hoveredIndex.reset();
                topmost->EnsureItemVisible(first.value());
                InvalidatePopup(*topmost);
            }
            return MenuInputDisposition::Consumed;
        }
        case MenuKeyboardKind::End:
        {
            const auto idx = focusedIndex();
            if (idx.has_value())
            {
                const auto& item = topmost->items[idx.value()];
                if (item.kind == MenuItemKind::Slider && ! item.sliderStops.empty() && item.enabled)
                {
                    static_cast<void>(invokeFocusedSliderStop(static_cast<uint32_t>(item.sliderStops.size() - 1u)));
                    return controller.running ? MenuInputDisposition::Consumed : MenuInputDisposition::Invoked;
                }
            }
            auto last = topmost->FindLastNavigableItem();
            if (last.has_value())
            {
                topmost->keyboardIndex = last;
                topmost->hoveredIndex.reset();
                topmost->EnsureItemVisible(last.value());
                InvalidatePopup(*topmost);
            }
            return MenuInputDisposition::Consumed;
        }
        case MenuKeyboardKind::Enter:
        case MenuKeyboardKind::Space:
        {
            auto idx = topmost->keyboardIndex.has_value() ? topmost->keyboardIndex : topmost->hoveredIndex;
            if (idx.has_value() && idx.value() < topmost->itemCount)
            {
                const auto& item = topmost->items[idx.value()];
                if (item.kind == MenuItemKind::Slider && item.enabled)
                {
                    static_cast<void>(invokeFocusedSliderStop(ClampSliderValue(item)));
                    return controller.running ? MenuInputDisposition::Consumed : MenuInputDisposition::Invoked;
                }
                if (! item.children.empty() && item.enabled)
                {
                    OpenSubmenu(controller, *topmost, idx.value(), true);
                }
                else if (IsInvokableMenuItem(item))
                {
                    controller.InvokeItem(item.commandId);
                    return MenuInputDisposition::Invoked;
                }
            }
            return MenuInputDisposition::Consumed;
        }
        case MenuKeyboardKind::Escape:
            if (controller.popups.size() > 1)
            {
                controller.CloseTopmostSubmenu();
                return MenuInputDisposition::Consumed;
            }
            controller.Dismiss();
            return MenuInputDisposition::Dismissed;
        case MenuKeyboardKind::Tab:
        case MenuKeyboardKind::F10:
        case MenuKeyboardKind::Alt:
            controller.Dismiss();
            return MenuInputDisposition::Dismissed;
        case MenuKeyboardKind::Right:
        {
            if (invokeFocusedSliderDelta(1))
            {
                return controller.running ? MenuInputDisposition::Consumed : MenuInputDisposition::Invoked;
            }
            auto idx = topmost->keyboardIndex.has_value() ? topmost->keyboardIndex : topmost->hoveredIndex;
            if (idx.has_value() && idx.value() < topmost->itemCount && ! topmost->items[idx.value()].children.empty() &&
                topmost->items[idx.value()].enabled)
            {
                OpenSubmenu(controller, *topmost, idx.value(), true);
                return MenuInputDisposition::Consumed;
            }

            if (controller.sessionCallbacks.switchRootFromDirection)
            {
                if (auto request = controller.sessionCallbacks.switchRootFromDirection(true); request.has_value())
                {
                    if (SwitchRootPopup(controller, std::move(request.value()), L"root-switch-keyboard", true))
                    {
                        RememberKeyboardRootSwitchMessageTime(controller, event.messageTime);
                        return MenuInputDisposition::SwitchedRoot;
                    }
                }
            }
            return MenuInputDisposition::Consumed;
        }
        case MenuKeyboardKind::Left:
        {
            if (invokeFocusedSliderDelta(-1))
            {
                return controller.running ? MenuInputDisposition::Consumed : MenuInputDisposition::Invoked;
            }
            if (controller.popups.size() > 1)
            {
                controller.CloseTopmostSubmenu();
            }
            else if (controller.sessionCallbacks.switchRootFromDirection)
            {
                if (auto request = controller.sessionCallbacks.switchRootFromDirection(false); request.has_value())
                {
                    if (SwitchRootPopup(controller, std::move(request.value()), L"root-switch-keyboard", true))
                    {
                        RememberKeyboardRootSwitchMessageTime(controller, event.messageTime);
                        return MenuInputDisposition::SwitchedRoot;
                    }
                }
            }
            return MenuInputDisposition::Consumed;
        }
        case MenuKeyboardKind::Mnemonic:
        {
            auto match = topmost->FindMnemonicItem(event.mnemonic);
            if (match.has_value())
            {
                const auto& item = topmost->items[match.value()];
                if (! item.children.empty() && item.enabled)
                {
                    topmost->keyboardIndex = match;
                    InvalidatePopup(*topmost);
                    OpenSubmenu(controller, *topmost, match.value(), true);
                }
                else if (IsInvokableMenuItem(item))
                {
                    controller.InvokeItem(item.commandId);
                    return MenuInputDisposition::Invoked;
                }
            }
            return MenuInputDisposition::Consumed;
        }
    }

    return MenuInputDisposition::Ignored;
}

[[nodiscard]] std::optional<MenuKeyboardKind> MenuKeyboardKindFromVirtualKey(UINT virtualKey) noexcept
{
    switch (virtualKey)
    {
        case VK_UP: return MenuKeyboardKind::Up;
        case VK_DOWN: return MenuKeyboardKind::Down;
        case VK_LEFT: return MenuKeyboardKind::Left;
        case VK_RIGHT: return MenuKeyboardKind::Right;
        case VK_HOME: return MenuKeyboardKind::Home;
        case VK_END: return MenuKeyboardKind::End;
        case VK_RETURN: return MenuKeyboardKind::Enter;
        case VK_SPACE: return MenuKeyboardKind::Space;
        case VK_ESCAPE: return MenuKeyboardKind::Escape;
        case VK_TAB: return MenuKeyboardKind::Tab;
        case VK_F10: return MenuKeyboardKind::F10;
        case VK_MENU: return MenuKeyboardKind::Alt;
        default:
            if (virtualKey >= 'A' && virtualKey <= 'Z')
            {
                return MenuKeyboardKind::Mnemonic;
            }
            return std::nullopt;
    }
}

struct MenuKeyboardTargetTrace
{
    int popupIndex = -1;
    uintptr_t popupHwnd = 0u;
    std::optional<size_t> hoveredIndex;
    std::optional<size_t> keyboardIndex;
};

[[nodiscard]] MenuKeyboardTargetTrace CaptureMenuKeyboardTargetTrace(const MenuController& controller) noexcept
{
    MenuKeyboardTargetTrace trace;
    const MenuPopup* popup = controller.GetTopmostPopup();
    if (! popup)
    {
        return trace;
    }

    trace.popupIndex    = TracePopupIndex(controller, popup);
    trace.popupHwnd     = TraceHwndValue(popup->hwnd);
    trace.hoveredIndex  = popup->hoveredIndex;
    trace.keyboardIndex = popup->keyboardIndex;
    return trace;
}

void TraceMenuKeyboardRoute(
    const MenuController& controller, const MenuKeyboardEvent& event, const MenuKeyboardTargetTrace& before, MenuInputDisposition disposition) noexcept
{
    static_cast<void>(event);
    static_cast<void>(before);
    static_cast<void>(disposition);

    const MenuKeyboardTargetTrace after = CaptureMenuKeyboardTargetTrace(controller);
    DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.keyboard",
                              L"source={} kind={} hwnd={:#x} virtualKey={:#x} mnemonic={:#x} messageTime={} keyboardRootTimeHas={} keyboardRootTime={} beforePopup={} beforePopupHwnd={:#x} "
                              L"beforeHoverHas={} beforeHover={} beforeKeyboardHas={} beforeKeyboard={} afterPopup={} afterPopupHwnd={:#x} "
                              L"afterHoverHas={} afterHover={} afterKeyboardHas={} afterKeyboard={} popups={} running={} result={} disposition={}",
                              TraceInputSourceName(event.source),
                              TraceKeyboardKindName(event.kind),
                              TraceHwndValue(event.hwnd),
                              static_cast<unsigned int>(event.virtualKey),
                              static_cast<unsigned int>(event.mnemonic),
                              static_cast<unsigned int>(event.messageTime),
                              TraceBool(controller.hasLastKeyboardRootSwitchMessageTime),
                              static_cast<unsigned int>(controller.lastKeyboardRootSwitchMessageTime),
                              before.popupIndex,
                              before.popupHwnd,
                              TraceBool(before.hoveredIndex.has_value()),
                              TraceOptionalSizeValue(before.hoveredIndex),
                              TraceBool(before.keyboardIndex.has_value()),
                              TraceOptionalSizeValue(before.keyboardIndex),
                              after.popupIndex,
                              after.popupHwnd,
                              TraceBool(after.hoveredIndex.has_value()),
                              TraceOptionalSizeValue(after.hoveredIndex),
                              TraceBool(after.keyboardIndex.has_value()),
                              TraceOptionalSizeValue(after.keyboardIndex),
                              controller.popups.size(),
                              TraceBool(controller.running),
                              controller.result.value_or(-1),
                              TraceDispositionName(disposition));
}

[[nodiscard]] bool ProcessMenuPopupMessage(MenuController& controller, HWND hwnd, UINT msg, WPARAM wp, LPARAM lp)
{
    if (! controller.running)
    {
        DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.popup-message",
                                  L"msg={} msgId={:#x} hwnd={:#x} wParam={:#x} lParam={:#x} running=false handled=false",
                                  TraceMenuMessageName(msg),
                                  static_cast<unsigned int>(msg),
                                  TraceHwndValue(hwnd),
                                  static_cast<uintptr_t>(wp),
                                  static_cast<uintptr_t>(lp));
        return false;
    }

    if (msg == WM_TIMER && wp == kSubmenuHoverTimerId)
    {
        if (MenuPopup* popup = controller.FindPopupForHwnd(hwnd))
        {
            HandleSubmenuHoverTimer(controller, *popup);
        }

        return true;
    }

    if (msg >= WM_MOUSEFIRST && msg <= WM_MOUSELAST)
    {
        const std::optional<MenuPointerKind> kind = MenuPointerKindFromMessage(msg);
        if (! kind.has_value())
        {
            DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.popup-message",
                                      L"msg={} msgId={:#x} hwnd={:#x} wParam={:#x} lParam={:#x} supported=false handled=false",
                                      TraceMenuMessageName(msg),
                                      static_cast<unsigned int>(msg),
                                      TraceHwndValue(hwnd),
                                      static_cast<uintptr_t>(wp),
                                      static_cast<uintptr_t>(lp));
            return false;
        }

        std::optional<D2D1_POINT_2F> deliveredPopupPointDip;
        if (msg != WM_MOUSEWHEEL && msg != WM_MOUSEHWHEEL)
        {
            if (const MenuPopup* deliveredPopup = controller.FindPopupForHwnd(hwnd))
            {
                deliveredPopupPointDip = PopupClientPointToDip(*deliveredPopup, UnpackMousePoint(lp));
            }
        }

        const MSG syntheticMessage{.hwnd = hwnd, .message = msg, .wParam = wp, .lParam = lp, .time = CurrentMessageTime()};
        const MenuPointerEvent event{
            .source                = MenuInputSource::PopupWndProc,
            .kind                  = kind.value(),
            .hwnd                  = hwnd,
            .screenPoint           = ResolveMouseScreenPoint(syntheticMessage),
            .messageTime           = syntheticMessage.time,
            .deliveredPopupPointDip = deliveredPopupPointDip,
            .wParam                = wp,
            .lParam                = lp,
            .mayInvoke             = true,
            .mayDismiss            = true,
            .maySwitchRoot         = msg == WM_MOUSEMOVE,
            .mayTakeKeyboardFocus  = true,
        };
        const MenuInputDisposition disposition = RouteMenuPointerEvent(controller, event);
        DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.popup-message",
                                  L"msg={} msgId={:#x} hwnd={:#x} wParam={:#x} lParam={:#x} handled={} disposition={}",
                                  TraceMenuMessageName(msg),
                                  static_cast<unsigned int>(msg),
                                  TraceHwndValue(hwnd),
                                  static_cast<uintptr_t>(wp),
                                  static_cast<uintptr_t>(lp),
                                  TraceBool(disposition != MenuInputDisposition::Ignored),
                                  TraceDispositionName(disposition));
        return disposition != MenuInputDisposition::Ignored;
    }

    if (msg == WM_KEYDOWN || msg == WM_SYSKEYDOWN)
    {
        const UINT virtualKey = static_cast<UINT>(wp);
        const std::optional<MenuKeyboardKind> kind = MenuKeyboardKindFromVirtualKey(virtualKey);
        if (! kind.has_value())
        {
            DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.popup-message",
                                      L"msg={} msgId={:#x} hwnd={:#x} virtualKey={:#x} supported=false handled=false",
                                      TraceMenuMessageName(msg),
                                      static_cast<unsigned int>(msg),
                                      TraceHwndValue(hwnd),
                                      static_cast<unsigned int>(virtualKey));
            return false;
        }

        const MenuKeyboardEvent event{
            .source     = MenuInputSource::PopupWndProc,
            .kind       = kind.value(),
            .hwnd       = hwnd,
            .virtualKey = virtualKey,
            .mnemonic   = static_cast<wchar_t>(virtualKey),
            .messageTime = CurrentMessageTime(),
            .lParam     = lp,
        };
        const bool traceEnabled = IsMenuDiagnosticsTraceEnabled();
        const MenuKeyboardTargetTrace traceBefore = traceEnabled ? CaptureMenuKeyboardTargetTrace(controller) : MenuKeyboardTargetTrace{};
        const MenuInputDisposition disposition    = RouteMenuKeyboardEvent(controller, event);
        if (traceEnabled)
        {
            TraceMenuKeyboardRoute(controller, event, traceBefore, disposition);
        }
        DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.popup-message",
                                  L"msg={} msgId={:#x} hwnd={:#x} virtualKey={:#x} handled={} disposition={}",
                                  TraceMenuMessageName(msg),
                                  static_cast<unsigned int>(msg),
                                  TraceHwndValue(hwnd),
                                  static_cast<unsigned int>(virtualKey),
                                  TraceBool(disposition != MenuInputDisposition::Ignored),
                                  TraceDispositionName(disposition));
        return disposition != MenuInputDisposition::Ignored;
    }

    DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.popup-message",
                              L"msg={} msgId={:#x} hwnd={:#x} wParam={:#x} lParam={:#x} handled=false",
                              TraceMenuMessageName(msg),
                              static_cast<unsigned int>(msg),
                              TraceHwndValue(hwnd),
                              static_cast<uintptr_t>(wp),
                              static_cast<uintptr_t>(lp));
    return false;
}

// ---------------------------------------------------------------------------
// Modal message loop
// ---------------------------------------------------------------------------

void RunMenuModalLoop(MenuController& controller)
{
    // Capture mouse on root popup for light-dismiss
    MenuPopup* root = controller.GetRootPopup();
    if (! root || ! root->hwnd)
        return;

    const HWND previousCapture = GetCapture();
    const HWND previousFocus   = GetFocus();
    SetCapture(root->hwnd);
    ActivatePopupForKeyboard(*root);
    DXUI_MENU_TRACE(L"DxUi::MenuTrace Popup loop-start owner={:#x} root={:#x} previousCapture={:#x} currentCapture={:#x} popupCount={} items={}",
                reinterpret_cast<uintptr_t>(controller.ownerHwnd),
                reinterpret_cast<uintptr_t>(root->hwnd),
                reinterpret_cast<uintptr_t>(previousCapture),
                reinterpret_cast<uintptr_t>(GetCapture()),
                controller.popups.size(),
                controller.rootItems.size());
    DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.loop-start",
                              L"owner={:#x} root={:#x} previousCapture={:#x} previousFocus={:#x} currentCapture={:#x} popupCount={} items={}",
                              TraceHwndValue(controller.ownerHwnd),
                              TraceHwndValue(root->hwnd),
                              TraceHwndValue(previousCapture),
                              TraceHwndValue(previousFocus),
                              TraceHwndValue(GetCapture()),
                              controller.popups.size(),
                              controller.rootItems.size());

    controller.ignoreInitialLeftButtonUp  = controller.sessionCallbacks.ignoreInitialLeftButtonUp || (GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0;
    controller.ignoreInitialRightButtonUp = controller.sessionCallbacks.ignoreInitialRightButtonUp || (GetAsyncKeyState(VK_RBUTTON) & 0x8000) != 0;
    controller.leftButtonDownInPopup      = false;
    controller.rightButtonDownInPopup     = false;
    DXUI_MENU_TRACE(L"DxUi::MenuTrace Popup loop-initial-up-filter left={} right={} asyncLeft={} asyncRight={}",
                controller.ignoreInitialLeftButtonUp ? L"true" : L"false",
                controller.ignoreInitialRightButtonUp ? L"true" : L"false",
                (GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0 ? L"down" : L"up",
                (GetAsyncKeyState(VK_RBUTTON) & 0x8000) != 0 ? L"down" : L"up");
    DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.loop-initial-up-filter",
                              L"ignoreLeft={} ignoreRight={} asyncLeftDown={} asyncRightDown={}",
                              TraceBool(controller.ignoreInitialLeftButtonUp),
                              TraceBool(controller.ignoreInitialRightButtonUp),
                              TraceBool((GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0),
                              TraceBool((GetAsyncKeyState(VK_RBUTTON) & 0x8000) != 0));
    MSG msg{};
#if DXUI_MENU_PERSISTENT_DIAGNOSTICS
    struct MenuLoopRepeatedMessageTrace
    {
        UINT message              = 0;
        HWND hwnd                 = nullptr;
        WPARAM wParam             = 0;
        LPARAM lParam             = 0;
        bool popupMessage         = false;
        bool ownerOrPopupMessage  = false;
        uint64_t count            = 0u;
    };
    MenuLoopRepeatedMessageTrace repeatedMessageTrace{};
    const auto flushRepeatedMessageTrace = [&]() noexcept
    {
        if (repeatedMessageTrace.count == 0u)
        {
            return;
        }
        DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.loop-message-repeat",
                                  L"msg={} msgId={:#x} hwnd={:#x} wParam={:#x} lParam={:#x} popup={} ownerOrPopup={} count={}",
                                  TraceMenuMessageName(repeatedMessageTrace.message),
                                  static_cast<unsigned int>(repeatedMessageTrace.message),
                                  TraceHwndValue(repeatedMessageTrace.hwnd),
                                  static_cast<uintptr_t>(repeatedMessageTrace.wParam),
                                  static_cast<uintptr_t>(repeatedMessageTrace.lParam),
                                  TraceBool(repeatedMessageTrace.popupMessage),
                                  TraceBool(repeatedMessageTrace.ownerOrPopupMessage),
                                  repeatedMessageTrace.count);
        repeatedMessageTrace = {};
    };
    const auto traceLoopMessage = [&](const MSG& loopMessage, bool popupMessage, bool ownerOrPopupMessage) noexcept
    {
        if (ShouldTraceMenuLoopMessageImmediately(loopMessage.message, popupMessage))
        {
            flushRepeatedMessageTrace();
            DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.loop-message",
                                      L"msg={} msgId={:#x} hwnd={:#x} wParam={:#x} lParam={:#x} popup={} ownerOrPopup={}",
                                      TraceMenuMessageName(loopMessage.message),
                                      static_cast<unsigned int>(loopMessage.message),
                                      TraceHwndValue(loopMessage.hwnd),
                                      static_cast<uintptr_t>(loopMessage.wParam),
                                      static_cast<uintptr_t>(loopMessage.lParam),
                                      TraceBool(popupMessage),
                                      TraceBool(ownerOrPopupMessage));
            return;
        }

        if (repeatedMessageTrace.count == 0u || repeatedMessageTrace.message != loopMessage.message || repeatedMessageTrace.hwnd != loopMessage.hwnd ||
            repeatedMessageTrace.wParam != loopMessage.wParam || repeatedMessageTrace.lParam != loopMessage.lParam ||
            repeatedMessageTrace.popupMessage != popupMessage || repeatedMessageTrace.ownerOrPopupMessage != ownerOrPopupMessage)
        {
            flushRepeatedMessageTrace();
            repeatedMessageTrace.message             = loopMessage.message;
            repeatedMessageTrace.hwnd                = loopMessage.hwnd;
            repeatedMessageTrace.wParam              = loopMessage.wParam;
            repeatedMessageTrace.lParam              = loopMessage.lParam;
            repeatedMessageTrace.popupMessage        = popupMessage;
            repeatedMessageTrace.ownerOrPopupMessage = ownerOrPopupMessage;
        }
        ++repeatedMessageTrace.count;
        if (repeatedMessageTrace.count >= kMenuLoopTraceRepeatFlushCount)
        {
            flushRepeatedMessageTrace();
        }
    };
#else
    const auto flushRepeatedMessageTrace = []() noexcept {};
    const auto traceLoopMessage = [](const MSG& /*loopMessage*/, bool /*popupMessage*/, bool /*ownerOrPopupMessage*/) noexcept {};
#endif
    while (controller.running)
    {
        if (! PeekMenuPriorityMessage(msg) && PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE) == FALSE)
        {
            MenuPopup* currentRoot = controller.GetRootPopup();
            if (currentRoot && currentRoot->hwnd && IsWindow(currentRoot->hwnd) != FALSE && GetCapture() != currentRoot->hwnd)
            {
                const HWND previousPopupCapture = GetCapture();
                SetCapture(currentRoot->hwnd);
                DXUI_MENU_TRACE(L"DxUi::MenuTrace Popup loop-recapture root={:#x} previousCapture={:#x} currentCapture={:#x}",
                            reinterpret_cast<uintptr_t>(currentRoot->hwnd),
                            reinterpret_cast<uintptr_t>(previousPopupCapture),
                            reinterpret_cast<uintptr_t>(GetCapture()));
                DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.loop-recapture",
                                          L"root={:#x} previousCapture={:#x} currentCapture={:#x}",
                                          TraceHwndValue(currentRoot->hwnd),
                                          TraceHwndValue(previousPopupCapture),
                                          TraceHwndValue(GetCapture()));
            }

            static_cast<void>(MsgWaitForMultipleObjectsEx(0, nullptr, INFINITE, QS_ALLINPUT, MWMO_INPUTAVAILABLE));
            continue;
        }

        if (msg.message == WM_QUIT)
        {
            flushRepeatedMessageTrace();
            PostQuitMessage(static_cast<int>(msg.wParam));
            break;
        }

        const HWND originalMessageHwnd = msg.hwnd;
        if (originalMessageHwnd && IsWindow(originalMessageHwnd) == FALSE)
        {
            flushRepeatedMessageTrace();
            DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.loop-message-stale",
                                      L"msg={} msgId={:#x} hwnd={:#x} wParam={:#x} lParam={:#x}",
                                      TraceMenuMessageName(msg.message),
                                      static_cast<unsigned int>(msg.message),
                                      TraceHwndValue(originalMessageHwnd),
                                      static_cast<uintptr_t>(msg.wParam),
                                      static_cast<uintptr_t>(msg.lParam));
            continue;
        }

        if (msg.message == WndMsg::kDxUiContextMenuRootHoverChanged)
        {
            flushRepeatedMessageTrace();
            if (controller.sessionCallbacks.switchRootFromMenuBarHover)
            {
                if (auto request = controller.sessionCallbacks.switchRootFromMenuBarHover(); request.has_value())
                {
                    DXUI_MENU_TRACE(L"DxUi::MenuTrace Popup root-switch-menu-bar-hover-accepted newItems={} newPoint=({}, {})",
                                    request->items.size(),
                                    request->screenPoint.x,
                                    request->screenPoint.y);
#if defined(ENABLE_TESTS)
                    ++controller.rootPointerSwitchCount;
#endif
                    ClearKeyboardRootSwitchMessageTime(controller);
                    static_cast<void>(SwitchRootPopup(controller, std::move(request.value()), L"root-switch-menu-bar-hover", false));
                }
            }
            continue;
        }

        if (msg.message == WM_CAPTURECHANGED || msg.message == WM_CANCELMODE)
        {
            DXUI_MENU_TRACE(L"DxUi::MenuTrace Popup message msg={} hwnd={:#x} wParam={:#x} lParam={:#x} capture={:#x}",
                            TraceMenuMessageName(msg.message),
                            reinterpret_cast<uintptr_t>(msg.hwnd),
                            static_cast<uintptr_t>(msg.wParam),
                            static_cast<uintptr_t>(msg.lParam),
                            reinterpret_cast<uintptr_t>(GetCapture()));
        }

        const bool popupMessage        = controller.FindPopupForHwnd(msg.hwnd) != nullptr;
        const bool ownerOrPopupMessage = msg.hwnd == controller.ownerHwnd || popupMessage;
        traceLoopMessage(msg, popupMessage, ownerOrPopupMessage);
        if (ownerOrPopupMessage)
        {
            if ((msg.message == WM_ACTIVATEAPP && msg.wParam == FALSE) ||
                (msg.message == WM_ACTIVATE && LOWORD(static_cast<DWORD_PTR>(msg.wParam)) == WA_INACTIVE) ||
                (msg.message == WM_NCACTIVATE && msg.wParam == FALSE) || (msg.message == WM_DESTROY && msg.hwnd == controller.ownerHwnd))
            {
                DXUI_MENU_TRACE(L"DxUi::MenuTrace Popup dismiss-activation msg={} hwnd={:#x} owner={:#x} capture={:#x}",
                            TraceMenuMessageName(msg.message),
                            reinterpret_cast<uintptr_t>(msg.hwnd),
                            reinterpret_cast<uintptr_t>(controller.ownerHwnd),
                            reinterpret_cast<uintptr_t>(GetCapture()));
                DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.loop-dismiss-activation",
                                          L"msg={} msgId={:#x} hwnd={:#x} owner={:#x} capture={:#x}",
                                          TraceMenuMessageName(msg.message),
                                          static_cast<unsigned int>(msg.message),
                                          TraceHwndValue(msg.hwnd),
                                          TraceHwndValue(controller.ownerHwnd),
                                          TraceHwndValue(GetCapture()));
                controller.Dismiss();
                break;
            }
        }

        // Timer messages for submenu hover delay stay in the menu loop because
        // they mutate the popup cascade, not the individual popup window.
        if (msg.message == WM_TIMER && msg.wParam == kSubmenuHoverTimerId)
        {
            if (MenuPopup* popup = controller.FindPopupForHwnd(msg.hwnd))
            {
                HandleSubmenuHoverTimer(controller, *popup);
            }
            continue;
        }

        if (popupMessage)
        {
            DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.loop-dispatch-popup",
                                      L"msg={} msgId={:#x} hwnd={:#x}",
                                      TraceMenuMessageName(msg.message),
                                      static_cast<unsigned int>(msg.message),
                                      TraceHwndValue(msg.hwnd));
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
            continue;
        }

        // Route non-popup mouse messages through the shared menu input router.
        if (IsMenuClientMouseMessage(msg.message))
        {
            const std::optional<MenuPointerKind> pointerKind = MenuPointerKindFromMessage(msg.message);
            if (! pointerKind.has_value())
            {
                continue;
            }
            const MenuPointerEvent event{
                .source               = MenuInputSource::ModalMessage,
                .kind                 = pointerKind.value(),
                .hwnd                 = msg.hwnd,
                .screenPoint          = ResolveMouseScreenPoint(msg),
                .messageTime          = msg.time,
                .wParam               = msg.wParam,
                .lParam               = msg.lParam,
                .mayInvoke            = true,
                .mayDismiss           = true,
                .maySwitchRoot        = msg.message == WM_MOUSEMOVE,
                .mayTakeKeyboardFocus = true,
            };
            const MenuInputDisposition disposition = RouteMenuPointerEvent(controller, event);
            DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.loop-nonpopup-pointer",
                                      L"msg={} msgId={:#x} hwnd={:#x} disposition={} running={}",
                                      TraceMenuMessageName(msg.message),
                                      static_cast<unsigned int>(msg.message),
                                      TraceHwndValue(msg.hwnd),
                                      TraceDispositionName(disposition),
                                      TraceBool(controller.running));
            if (disposition == MenuInputDisposition::Dismissed || disposition == MenuInputDisposition::Invoked)
            {
                break;
            }
            continue;
        }

        // Keyboard messages — route to topmost popup
        if (msg.message == WM_KEYDOWN || msg.message == WM_SYSKEYDOWN)
        {
            const UINT virtualKey = static_cast<UINT>(msg.wParam);
            const std::optional<MenuKeyboardKind> keyboardKind = MenuKeyboardKindFromVirtualKey(virtualKey);
            if (keyboardKind.has_value())
            {
                const MenuKeyboardEvent event{
                    .source     = MenuInputSource::ModalMessage,
                    .kind       = keyboardKind.value(),
                    .hwnd       = msg.hwnd,
                    .virtualKey = virtualKey,
                    .mnemonic   = static_cast<wchar_t>(virtualKey),
                    .messageTime = msg.time,
                    .lParam     = msg.lParam,
                };
                const bool traceEnabled = IsMenuDiagnosticsTraceEnabled();
                const MenuKeyboardTargetTrace traceBefore = traceEnabled ? CaptureMenuKeyboardTargetTrace(controller) : MenuKeyboardTargetTrace{};
                const MenuInputDisposition disposition = RouteMenuKeyboardEvent(controller, event);
                if (traceEnabled)
                {
                    TraceMenuKeyboardRoute(controller, event, traceBefore, disposition);
                }
                DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.loop-keyboard",
                                          L"msg={} msgId={:#x} hwnd={:#x} virtualKey={:#x} disposition={} running={}",
                                          TraceMenuMessageName(msg.message),
                                          static_cast<unsigned int>(msg.message),
                                          TraceHwndValue(msg.hwnd),
                                          static_cast<unsigned int>(virtualKey),
                                          TraceDispositionName(disposition),
                                          TraceBool(controller.running));
                if (disposition == MenuInputDisposition::Dismissed || disposition == MenuInputDisposition::Invoked)
                {
                    break;
                }
                if (disposition != MenuInputDisposition::Ignored)
                {
                    continue;
                }
            }

            // Fall through for unhandled keys
        }

        // All WM_PAINT/WM_ERASEBKGND handled by MenuWndProc → WindowHost::HandleMessage
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    flushRepeatedMessageTrace();

    // Release capture
    if (const HWND captured = GetCapture(); captured && controller.FindPopupForHwnd(captured))
    {
        DXUI_MENU_TRACE(L"DxUi::MenuTrace Popup loop-release-capture captured={:#x} owner={:#x}",
                    reinterpret_cast<uintptr_t>(captured),
                    reinterpret_cast<uintptr_t>(controller.ownerHwnd));
        DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.loop-release-capture",
                                  L"captured={:#x} owner={:#x}",
                                  TraceHwndValue(captured),
                                  TraceHwndValue(controller.ownerHwnd));
        ReleaseCapture();
    }
    DXUI_MENU_TRACE(L"DxUi::MenuTrace Popup loop-end owner={:#x} result={} capture={:#x}",
                reinterpret_cast<uintptr_t>(controller.ownerHwnd),
                controller.result.value_or(-1),
                reinterpret_cast<uintptr_t>(GetCapture()));
    DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.loop-end",
                              L"owner={:#x} result={} capture={:#x} focus={:#x} previousFocus={:#x}",
                              TraceHwndValue(controller.ownerHwnd),
                              controller.result.value_or(-1),
                              TraceHwndValue(GetCapture()),
                              TraceHwndValue(GetFocus()),
                              TraceHwndValue(previousFocus));

    const HWND currentFocus = GetFocus();
    if (previousFocus && IsWindow(previousFocus) != FALSE && currentFocus && controller.FindPopupForHwnd(currentFocus))
    {
        SetFocus(previousFocus);
        DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.loop-restore-focus",
                                  L"previousFocus={:#x} focusAfter={:#x}",
                                  TraceHwndValue(previousFocus),
                                  TraceHwndValue(GetFocus()));
    }
}

} // anonymous namespace

// ---------------------------------------------------------------------------
// ContextMenu::Show — public API
// ---------------------------------------------------------------------------

bool IsContextMenuDiagnosticsEnabled() noexcept
{
    return IsMenuDiagnosticsTraceEnabled();
}

void TraceContextMenuDiagnostics(std::wstring_view eventName, std::wstring_view details) noexcept
{
    WriteMenuDiagnosticsTraceLine(eventName, details);
}

std::optional<int> ContextMenu::Show(
    HWND ownerHwnd, POINT screenPoint, std::span<const MenuFlyoutItem> items, const ThemePalette& theme, const ContextMenuSessionCallbacks& sessionCallbacks)
{
    if (items.empty() || ! ownerHwnd)
        return std::nullopt;

    DXUI_MENU_TRACE(
        L"DxUi::MenuTrace ContextMenu show-begin owner={:#x} point=({}, {}) items={} focusFirst={} ignoreInitialUp=({}, {}) captureBefore={:#x}",
        reinterpret_cast<uintptr_t>(ownerHwnd),
        screenPoint.x,
        screenPoint.y,
        items.size(),
        sessionCallbacks.focusFirstNavigableItem ? L"true" : L"false",
        sessionCallbacks.ignoreInitialLeftButtonUp ? L"true" : L"false",
        sessionCallbacks.ignoreInitialRightButtonUp ? L"true" : L"false",
        reinterpret_cast<uintptr_t>(GetCapture()));
    DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.show-begin",
                              L"owner={:#x} point=({}, {}) items={} focusFirst={} ignoreInitialLeftUp={} ignoreInitialRightUp={} maxRootHeightDip={:.1f} "
                              L"captureBefore={:#x} focusBefore={:#x} activeBefore={:#x} hitWindow={:#x}",
                              TraceHwndValue(ownerHwnd),
                              screenPoint.x,
                              screenPoint.y,
                              items.size(),
                              TraceBool(sessionCallbacks.focusFirstNavigableItem),
                              TraceBool(sessionCallbacks.ignoreInitialLeftButtonUp),
                              TraceBool(sessionCallbacks.ignoreInitialRightButtonUp),
                              sessionCallbacks.maxRootHeightDip,
                              TraceHwndValue(GetCapture()),
                              TraceHwndValue(GetFocus()),
                              TraceHwndValue(GetActiveWindow()),
                              TraceHwndValue(WindowFromPoint(screenPoint)));

    MenuController controller;
    controller.ownerHwnd        = ownerHwnd;
    controller.theme            = theme; // Copy
    controller.style            = ResolveMenuVisualStyle(theme);
    controller.sessionCallbacks = sessionCallbacks;
    if (sessionCallbacks.focusFirstNavigableItem)
    {
        RememberKeyboardRootSwitchMessageTime(controller, CurrentMessageTime());
    }

    // Copy items so they outlive the caller's temporaries
    controller.rootItems.assign(items.begin(), items.end());

    const auto startedAt = std::chrono::steady_clock::now();
    if (! CreateMenuPopupWindow(controller,
                                controller.rootItems.data(),
                                controller.rootItems.size(),
                                screenPoint,
                                false,
                                nullptr,
                                nullptr,
                                true,
                                sessionCallbacks.focusFirstNavigableItem))
    {
        DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.show-create-failed",
                                  L"owner={:#x} point=({}, {}) items={}",
                                  TraceHwndValue(ownerHwnd),
                                  screenPoint.x,
                                  screenPoint.y,
                                  items.size());
        return std::nullopt;
    }
    Debug::Perf::Emit(L"DxUI::PopupShow", L"root", Debug::Perf::ElapsedUs(startedAt), static_cast<uint64_t>(controller.rootItems.size()), 0u);

    // Run the modal loop
    RunMenuModalLoop(controller);

    DestroyPopupChain(controller);

    DXUI_MENU_TRACE(L"DxUi::MenuTrace ContextMenu show-end owner={:#x} result={} captureAfter={:#x}",
                reinterpret_cast<uintptr_t>(ownerHwnd),
                controller.result.value_or(-1),
                reinterpret_cast<uintptr_t>(GetCapture()));
    DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.show-end",
                              L"owner={:#x} result={} captureAfter={:#x} focusAfter={:#x}",
                              TraceHwndValue(ownerHwnd),
                              controller.result.value_or(-1),
                              TraceHwndValue(GetCapture()),
                              TraceHwndValue(GetFocus()));

    return controller.result;
}

bool ContextMenu::ShowAsync(HWND ownerHwnd,
                            POINT screenPoint,
                            std::span<const MenuFlyoutItem> items,
                            const ThemePalette& theme,
                            ContextMenuClosedCallback onClosed,
                            const ContextMenuSessionCallbacks& sessionCallbacks)
{
    if (items.empty() || ! ownerHwnd)
    {
        return false;
    }

    DXUI_MENU_TRACE(
        L"DxUi::MenuTrace ContextMenu show-async-begin owner={:#x} point=({}, {}) items={} focusFirst={} ignoreInitialUp=({}, {}) captureBefore={:#x}",
        reinterpret_cast<uintptr_t>(ownerHwnd),
        screenPoint.x,
        screenPoint.y,
        items.size(),
        sessionCallbacks.focusFirstNavigableItem ? L"true" : L"false",
        sessionCallbacks.ignoreInitialLeftButtonUp ? L"true" : L"false",
        sessionCallbacks.ignoreInitialRightButtonUp ? L"true" : L"false",
        reinterpret_cast<uintptr_t>(GetCapture()));
    DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.show-async-begin",
                              L"owner={:#x} point=({}, {}) items={} focusFirst={} ignoreInitialLeftUp={} ignoreInitialRightUp={} maxRootHeightDip={:.1f} "
                              L"captureBefore={:#x} focusBefore={:#x} activeBefore={:#x} hitWindow={:#x}",
                              TraceHwndValue(ownerHwnd),
                              screenPoint.x,
                              screenPoint.y,
                              items.size(),
                              TraceBool(sessionCallbacks.focusFirstNavigableItem),
                              TraceBool(sessionCallbacks.ignoreInitialLeftButtonUp),
                              TraceBool(sessionCallbacks.ignoreInitialRightButtonUp),
                              sessionCallbacks.maxRootHeightDip,
                              TraceHwndValue(GetCapture()),
                              TraceHwndValue(GetFocus()),
                              TraceHwndValue(GetActiveWindow()),
                              TraceHwndValue(WindowFromPoint(screenPoint)));

    auto controller = std::make_unique<MenuController>();
    controller->ownerHwnd        = ownerHwnd;
    controller->theme            = theme;
    controller->style            = ResolveMenuVisualStyle(theme);
    controller->sessionCallbacks = sessionCallbacks;
    controller->asyncOnClosed    = std::move(onClosed);
    controller->asyncSession     = true;
    if (sessionCallbacks.focusFirstNavigableItem)
    {
        RememberKeyboardRootSwitchMessageTime(*controller, CurrentMessageTime());
    }

    controller->rootItems.assign(items.begin(), items.end());

    const auto startedAt = std::chrono::steady_clock::now();
    if (! CreateMenuPopupWindow(*controller,
                                controller->rootItems.data(),
                                controller->rootItems.size(),
                                screenPoint,
                                false,
                                nullptr,
                                nullptr,
                                true,
                                sessionCallbacks.focusFirstNavigableItem))
    {
        DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.show-async-create-failed",
                                  L"owner={:#x} point=({}, {}) items={}",
                                  TraceHwndValue(ownerHwnd),
                                  screenPoint.x,
                                  screenPoint.y,
                                  items.size());
        return false;
    }
    Debug::Perf::Emit(L"DxUI::PopupShow", L"root_async", Debug::Perf::ElapsedUs(startedAt), static_cast<uint64_t>(controller->rootItems.size()), 0u);

    MenuController& controllerRef = *controller;
    ActiveAsyncMenuControllers().push_back(std::move(controller));
    if (! BeginAsyncMenuInteraction(controllerRef))
    {
        // Contract: returning false means the closed callback never fires. Drop
        // it before finalizing so FinalizeAsyncMenuController cannot invoke it.
        controllerRef.asyncOnClosed = nullptr;
        FinalizeAsyncMenuController(controllerRef);
        return false;
    }

    DXUI_MENU_DIAGNOSTICS_TRACE(L"menu.show-async-end",
                              L"owner={:#x} root={:#x} captureAfter={:#x} focusAfter={:#x}",
                              TraceHwndValue(ownerHwnd),
                              TraceHwndValue(controllerRef.GetRootPopup() ? controllerRef.GetRootPopup()->hwnd : nullptr),
                              TraceHwndValue(GetCapture()),
                              TraceHwndValue(GetFocus()));
    return true;
}

#if defined(ENABLE_TESTS)
bool DebugGetContextMenuItemDisplayText(const MenuFlyoutItem& item, std::wstring& outText)
{
    outText = ParseMenuLabel(DecodeMenuItemText(item).labelText).displayText;
    return true;
}

bool TryGetMenuPopupState(const MenuPopup& popup, ContextMenuPopupDebugState& outState) noexcept
{
    outState          = {};

    outState.hasScrollbar          = popup.NeedsScrollbar();
    outState.usesSystemBackdrop    = popup.usesSystemBackdrop;
    outState.usesAppBackdropBlur   = popup.usesAppBackdropBlur;
    outState.dpi                   = popup.dpi;
    outState.visibleWidthDip       = popup.menuWidthDip;
    outState.visibleHeightDip      = popup.menuHeightDip;
    outState.contentHeightDip      = popup.contentHeightDip;
    outState.scrollOffsetDip       = popup.scrollOffsetDip;
    outState.viewportRectDip       = popup.GetViewportRect();
    outState.scrollbarTrackRectDip = popup.NeedsScrollbar() ? popup.GetScrollbarTrackRect() : D2D1::RectF();
    outState.scrollbarThumbRectDip = popup.NeedsScrollbar() ? popup.GetScrollbarThumbRect() : D2D1::RectF();
    outState.surfaceRectPx         = popup.surfaceRectPx;
    outState.windowRectPx          = popup.windowRectPx;
    outState.hoveredIndex          = popup.hoveredIndex;
    outState.keyboardIndex         = popup.keyboardIndex;
    outState.hoverTimerActive      = popup.hoverTimerId != 0;
    outState.hoverTimerPendingOpen = popup.hoverTimerKind == MenuPopup::SubmenuHoverTimerKind::PendingOpen;
    outState.hoverTimerPendingClose = popup.hoverTimerKind == MenuPopup::SubmenuHoverTimerKind::PendingClose;
    outState.hoverTimerItemIndex =
        popup.hoverTimerItemIndex != SIZE_MAX ? std::optional<size_t>{popup.hoverTimerItemIndex} : std::nullopt;
    outState.itemTexts.reserve(popup.itemCount);
    outState.itemAcceleratorTexts.reserve(popup.itemCount);
    outState.itemKinds.reserve(popup.itemCount);
    outState.itemEnabled.reserve(popup.itemCount);
    outState.sliderValues.reserve(popup.itemCount);
    outState.sliderStopCounts.reserve(popup.itemCount);
    for (size_t itemIndex = 0; itemIndex < popup.itemCount; ++itemIndex)
    {
        std::wstring text;
        static_cast<void>(DebugGetContextMenuItemDisplayText(popup.items[itemIndex], text));
        outState.itemTexts.push_back(std::move(text));
        outState.itemAcceleratorTexts.push_back(std::wstring(DecodeMenuItemText(popup.items[itemIndex]).acceleratorText));
        outState.itemKinds.push_back(popup.items[itemIndex].kind);
        outState.itemEnabled.push_back(popup.items[itemIndex].enabled);
        if (popup.items[itemIndex].kind == MenuItemKind::Slider)
        {
            outState.sliderValues.push_back(ClampSliderValue(popup.items[itemIndex]));
            outState.sliderStopCounts.push_back(static_cast<uint32_t>(popup.items[itemIndex].sliderStops.size()));
        }
        else
        {
            outState.sliderValues.push_back(0u);
            outState.sliderStopCounts.push_back(0u);
        }
    }
    outState.renderCount = popup.host.DebugGetRenderCount();
    if (popup.controller)
    {
        outState.rootPointerSwitchCount          = popup.controller->rootPointerSwitchCount;
        outState.rootSwitchImmediateRenderCount = popup.controller->rootSwitchImmediateRenderCount;
    }
    return true;
}

bool DebugGetContextMenuPopupState(HWND hwnd, ContextMenuPopupDebugState& outState) noexcept
{
    const auto* popup = reinterpret_cast<const MenuPopup*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! popup || popup->hwnd != hwnd)
    {
        outState = {};
        return false;
    }

    if (GetWindowThreadProcessId(hwnd, nullptr) == GetCurrentThreadId())
    {
        return TryGetMenuPopupState(*popup, outState);
    }

    return SendMessageW(hwnd, kMenuDebugGetStateMessage, 0, reinterpret_cast<LPARAM>(&outState)) != FALSE;
}

bool DebugFireContextMenuPopupHoverTimer(HWND hwnd) noexcept
{
    const auto* popup = reinterpret_cast<const MenuPopup*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! popup || popup->hwnd != hwnd)
    {
        return false;
    }

    return PostMessageW(hwnd, WM_TIMER, kSubmenuHoverTimerId, 0) != FALSE;
}

bool TryGetMenuPopupItemRect(const MenuPopup& popup, size_t itemIndex, D2D1_RECT_F& outRectDip) noexcept
{
    outRectDip = D2D1::RectF();
    if (itemIndex >= popup.itemCount)
    {
        return false;
    }

    outRectDip = GetVisibleItemRect(popup, itemIndex);
    return outRectDip.right > outRectDip.left && outRectDip.bottom > outRectDip.top;
}

bool DebugGetContextMenuPopupItemRect(HWND hwnd, size_t itemIndex, D2D1_RECT_F& outRectDip) noexcept
{
    const auto* popup = reinterpret_cast<const MenuPopup*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! popup || popup->hwnd != hwnd)
    {
        outRectDip = D2D1::RectF();
        return false;
    }

    if (GetWindowThreadProcessId(hwnd, nullptr) == GetCurrentThreadId())
    {
        return TryGetMenuPopupItemRect(*popup, itemIndex, outRectDip);
    }

    MenuDebugGetItemRectRequest request{.itemIndex = itemIndex, .outRectDip = &outRectDip};
    return SendMessageW(hwnd, kMenuDebugGetItemRectMessage, 0, reinterpret_cast<LPARAM>(&request)) != FALSE;
}

bool TryGetMenuPopupItemPaint(const MenuPopup& popup, size_t itemIndex, ContextMenuPopupItemPaintDebugState& outState) noexcept
{
    outState = {};
    if (itemIndex >= popup.itemCount || ! popup.controller)
    {
        return false;
    }

    const auto& item              = popup.items[itemIndex];
    const bool isHovered          = (popup.hoveredIndex.has_value() && popup.hoveredIndex.value() == itemIndex) ||
                                    (popup.keyboardIndex.has_value() && popup.keyboardIndex.value() == itemIndex);
    const ParsedMenuLabel label   = ParseMenuLabel(DecodeMenuItemText(item).labelText);
    const auto itemPaint          = ResolveMenuItemPaintStyle(popup.controller->theme, popup.controller->style, item, label.displayText, isHovered);
    outState.hovered              = isHovered;
    outState.disabled             = ! item.enabled;
    outState.usesHighlightFill    = itemPaint.showHighlightFill;
    outState.usesRainbowHighlight = itemPaint.usesRainbowHighlight;
    outState.fillColor            = itemPaint.fill;
    outState.compositeFillColor   = itemPaint.compositeFill;
    outState.textColor            = itemPaint.text;
    outState.acceleratorColor     = itemPaint.accelText;
    outState.iconColor            = itemPaint.iconColor;
    outState.checkColor           = itemPaint.checkColor;
    outState.chevronColor         = itemPaint.chevronColor;
    return true;
}

bool DebugGetContextMenuPopupItemPaint(HWND hwnd, size_t itemIndex, ContextMenuPopupItemPaintDebugState& outState) noexcept
{
    const auto* popup = reinterpret_cast<const MenuPopup*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! popup || popup->hwnd != hwnd)
    {
        outState = {};
        return false;
    }

    if (GetWindowThreadProcessId(hwnd, nullptr) == GetCurrentThreadId())
    {
        return TryGetMenuPopupItemPaint(*popup, itemIndex, outState);
    }

    MenuDebugGetItemPaintRequest request{.itemIndex = itemIndex, .outState = &outState};
    return SendMessageW(hwnd, kMenuDebugGetItemPaintMessage, 0, reinterpret_cast<LPARAM>(&request)) != FALSE;
}

bool DebugGetContextMenuPopupItemText(HWND hwnd, size_t itemIndex, std::wstring& outText) noexcept
{
    outText.clear();
    const auto* popup = reinterpret_cast<const MenuPopup*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! popup || popup->hwnd != hwnd)
    {
        return false;
    }

    if (GetWindowThreadProcessId(hwnd, nullptr) == GetCurrentThreadId())
    {
        return TryGetMenuPopupItemText(*popup, itemIndex, outText);
    }

    MenuDebugGetItemTextRequest request{.itemIndex = itemIndex, .outText = &outText};
    return SendMessageW(hwnd, kMenuDebugGetItemTextMessage, 0, reinterpret_cast<LPARAM>(&request)) != FALSE;
}

bool DebugGetContextMenuPopupItemLayout(HWND hwnd, size_t itemIndex, ContextMenuPopupItemLayoutDebugState& outState) noexcept
{
    outState          = {};
    const auto* popup = reinterpret_cast<const MenuPopup*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! popup || popup->hwnd != hwnd || itemIndex >= popup->itemCount)
    {
        return false;
    }

    const MenuItemLayoutRects layout = GetMenuItemLayoutRects(*popup, itemIndex);
    outState.itemRectDip             = layout.itemRectDip;
    outState.iconRectDip             = layout.iconRectDip;
    outState.textRectDip             = layout.textRectDip;
    outState.acceleratorRectDip      = layout.acceleratorRectDip;
    outState.chevronRectDip          = layout.chevronRectDip;
    outState.hasBitmapIcon           = popup->items[itemIndex].iconBitmap != nullptr;
    return outState.itemRectDip.right > outState.itemRectDip.left && outState.itemRectDip.bottom > outState.itemRectDip.top;
}

bool DebugSetContextMenuPopupBackdropCapture(HWND hwnd, const WindowHostBitmapCapture& capture) noexcept
{
    if (capture.widthPx == 0u || capture.heightPx == 0u || capture.bgraPixels.empty())
    {
        return false;
    }

    auto* popup = reinterpret_cast<MenuPopup*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! popup || popup->hwnd != hwnd)
    {
        return false;
    }

    if (GetWindowThreadProcessId(hwnd, nullptr) == GetCurrentThreadId())
    {
        popup->backdropSnapshot.capture      = capture;
        popup->backdropSnapshot.cachedBitmap = nullptr;
        popup->backdropSnapshot.cachedDevice = nullptr;
        popup->usesAppBackdropBlur           = true;
        popup->host.Invalidate();
        return true;
    }

    return SendMessageW(hwnd, kMenuDebugSetBackdropMessage, 0, reinterpret_cast<LPARAM>(&capture)) != FALSE;
}

bool DebugCaptureContextMenuPopupBitmap(HWND hwnd, WindowHostBitmapCapture& outCapture) noexcept
{
    outCapture  = {};
    auto* popup = reinterpret_cast<MenuPopup*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (! popup || popup->hwnd != hwnd)
    {
        return false;
    }

    if (GetWindowThreadProcessId(hwnd, nullptr) == GetCurrentThreadId())
    {
        return popup->host.DebugCaptureBitmap(outCapture);
    }

    return SendMessageW(hwnd, kMenuDebugCaptureBitmapMessage, 0, reinterpret_cast<LPARAM>(&outCapture)) != FALSE;
}

bool DebugComputeContextMenuPopupPosition(POINT screenPoint,
                                          float widthDip,
                                          float heightDip,
                                          UINT dpi,
                                          bool isSubmenu,
                                          const RECT* parentRect,
                                          const RECT* parentItemRect,
                                          ContextMenuRootHorizontalAlignment rootHorizontalAlignment,
                                          RECT& outRect) noexcept
{
    return DebugComputeContextMenuPopupPosition(screenPoint,
                                                widthDip,
                                                heightDip,
                                                dpi,
                                                isSubmenu,
                                                parentRect,
                                                parentItemRect,
                                                rootHorizontalAlignment,
                                                ContextMenuRootVerticalPlacement::Below,
                                                outRect);
}

bool DebugComputeContextMenuPopupPosition(POINT screenPoint,
                                          float widthDip,
                                          float heightDip,
                                          UINT dpi,
                                          bool isSubmenu,
                                          const RECT* parentRect,
                                          const RECT* parentItemRect,
                                          ContextMenuRootHorizontalAlignment rootHorizontalAlignment,
                                          ContextMenuRootVerticalPlacement rootVerticalPlacement,
                                          RECT& outRect) noexcept
{
    outRect = RECT{};
    if (dpi == 0 || widthDip <= 0.0f || heightDip <= 0.0f)
    {
        return false;
    }

    outRect =
        ComputePopupPosition(screenPoint, widthDip, heightDip, dpi, isSubmenu, rootHorizontalAlignment, rootVerticalPlacement, parentRect, parentItemRect);
    return outRect.right > outRect.left && outRect.bottom > outRect.top;
}
#endif

} // namespace RedSalamander::DxUi
