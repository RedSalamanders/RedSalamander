#include "ViewerWeb.h"
#include "HandleIo.h"
#include "PathUtils.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <cctype>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <cwctype>
#include <filesystem>
#include <format>
#include <iterator>
#include <limits>
#include <memory>
#include <new>
#include <optional>
#include <span>
#include <string_view>
#include <type_traits>
#include <utility>
#include <vector>

#include <commdlg.h>
#include <d2d1.h>
#include <dwmapi.h>
#include <shellapi.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <shobjidl.h>
#include <uxtheme.h>

#pragma warning(push)
// (C6297) Arithmetic overflow. Results might not be an expected value.
// (C28182) Dereferencing NULL pointer.
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

#pragma comment(lib, "d2d1")
#pragma comment(lib, "Dwmapi.lib")
#pragma comment(lib, "dwrite")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "uxtheme")
#pragma comment(lib, "WebView2Loader.dll.lib")

#include "DxUi/DxUi.Typography.h"
#include "Helpers.h"
#include "JsStringEscape.h"
#include "LocalizationManager.h"
#include "UnicodeClipboard.h"
#include "ViewerWebCleanupTracker.h"
#include "ViewerFileComboHost.h"
#include "ViewerTitleBarTheme.h"
#include "WindowMessages.h"
#include "WindowSizing.h"

#include "resource.h"

extern HINSTANCE g_hInstance;

namespace Typography = RedSalamander::DxUi::Typography;

using RedSalamander::DxUi::ComboBox;
using RedSalamander::DxUi::ComboBoxVariant;
using RedSalamander::DxUi::MakeDefaultThemePalette;
using RedSalamander::DxUi::MakeThemePaletteFromViewerTheme;

// Defined in JsStringEscape.h; hoisted to global scope so the unqualified call sites inside the
// anonymous namespaces below (and the ViewerWeb:: members) resolve to the shared, hardened escapers.
using ViewerWebDetail::EscapeJavaScriptString;
using ViewerWebDetail::EscapeJavaScriptStringUtf8;

namespace
{
constexpr UINT kAsyncLoadCompleteMessage          = WndMsg::kViewerWebAsyncLoadComplete;
constexpr UINT kAsyncSaveCompleteMessage          = WM_APP + 0x634u;
constexpr UINT kAsyncPostFailureMessage           = WM_APP + 0x635u;
constexpr UINT kAsyncPostFailureLoad              = 1u;
constexpr UINT kAsyncPostFailureSave              = 2u;
constexpr int kHeaderHeightDip                    = 28;
static const int kViewerWebModuleAnchor = 0;

std::atomic<uint64_t> g_liveComCallbackCount{0u};
std::atomic<uint64_t> g_comCallbackReleaseEpoch{0u};
std::atomic<uint64_t> g_observedComCallbackReleaseEpoch{0u};
std::atomic<uint64_t> g_activeAsyncWorkerCount{0u};
std::atomic<DWORD> g_comCallbackOwnerThreadId{0u};
std::atomic<bool> g_comCallbackThreadViolation{false};
std::atomic<bool> g_shutdownCleanupComplete{false};
#ifdef ENABLE_TESTS
std::atomic<bool> g_failNextAsyncLoadCompletionPost{false};
std::atomic<bool> g_failNextAsyncSaveCompletionPost{false};
#endif

constexpr wchar_t kFileComboHostOriginalWndProcProp[] = L"RS.ViewerWeb.FileComboHostOriginalWndProc";
constexpr wchar_t kFileComboHostStateProp[]           = L"RS.ViewerWeb.FileComboHostState";

LRESULT CALLBACK FileComboHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

void UnhookFileComboHostWindow(HWND hwnd) noexcept
{
    RedSalamander::ViewerFileComboHost::UnhookFileComboHostWindow(hwnd, kFileComboHostStateProp, kFileComboHostOriginalWndProcProp, FileComboHostWndProc);
}

[[nodiscard]] size_t CountOwnerDrawMenuItems(HMENU menu) noexcept
{
    if (! menu)
    {
        return 0u;
    }

    const int itemCount = GetMenuItemCount(menu);
    if (itemCount <= 0)
    {
        return 0u;
    }

    size_t ownerDrawCount = 0u;
    for (UINT position = 0; position < static_cast<UINT>(itemCount); ++position)
    {
        MENUITEMINFOW itemInfo{};
        itemInfo.cbSize = sizeof(itemInfo);
        itemInfo.fMask  = MIIM_FTYPE | MIIM_SUBMENU;
        if (GetMenuItemInfoW(menu, position, TRUE, &itemInfo) == 0)
        {
            continue;
        }

        if ((itemInfo.fType & MFT_OWNERDRAW) != 0)
        {
            ++ownerDrawCount;
        }

        if (itemInfo.hSubMenu)
        {
            ownerDrawCount += CountOwnerDrawMenuItems(itemInfo.hSubMenu);
        }
    }

    return ownerDrawCount;
}

[[nodiscard]] D2D1_COLOR_F D2DColor(COLORREF color) noexcept
{
    return D2D1::ColorF(
        static_cast<float>(GetRValue(color)) / 255.0f, static_cast<float>(GetGValue(color)) / 255.0f, static_cast<float>(GetBValue(color)) / 255.0f, 1.0f);
}

[[nodiscard]] float PixelsToDips(float pixels, UINT dpi) noexcept
{
    const UINT effectiveDpi = std::max<UINT>(dpi, USER_DEFAULT_SCREEN_DPI);
    return (pixels * static_cast<float>(USER_DEFAULT_SCREEN_DPI)) / static_cast<float>(effectiveDpi);
}

[[nodiscard]] ID2D1Factory* GetSharedStatusD2DFactory() noexcept
{
    static const auto resources = []() noexcept
    {
        wil::com_ptr<ID2D1Factory> factory;
        static_cast<void>(D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, factory.put()));
        return factory;
    }();

    return resources.get();
}

[[nodiscard]] bool DrawStatusMessageWithDirectWrite(HDC hdc, HWND hwnd, const RECT& rcPx, std::wstring_view text, COLORREF textColor) noexcept
{
    if (! hdc || ! hwnd || text.empty() || rcPx.right <= rcPx.left || rcPx.bottom <= rcPx.top ||
        text.size() > static_cast<size_t>((std::numeric_limits<UINT32>::max)()))
    {
        return false;
    }

    ID2D1Factory* const d2dFactory      = GetSharedStatusD2DFactory();
    IDWriteFactory* const dwriteFactory = Typography::GetSharedMeasurementFactory();
    if (! d2dFactory || ! dwriteFactory)
    {
        return false;
    }

    const UINT dpi = std::max<UINT>(GetDpiForWindow(hwnd), USER_DEFAULT_SCREEN_DPI);
    wil::com_ptr<IDWriteTextFormat> format;
    if (FAILED(Typography::CreateTextFormat(dwriteFactory, Typography::MakeUiTextSpec(12.0f), format.put())) || ! format)
    {
        return false;
    }
    static_cast<void>(format->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_TRAILING));
    static_cast<void>(format->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER));
    static_cast<void>(format->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP));

    D2D1_RENDER_TARGET_PROPERTIES props = D2D1::RenderTargetProperties();
    props.dpiX                          = static_cast<float>(dpi);
    props.dpiY                          = static_cast<float>(dpi);

    wil::com_ptr<ID2D1DCRenderTarget> target;
    if (FAILED(d2dFactory->CreateDCRenderTarget(&props, target.put())) || ! target)
    {
        return false;
    }
    if (FAILED(target->BindDC(hdc, &rcPx)))
    {
        return false;
    }

    wil::com_ptr<ID2D1SolidColorBrush> brush;
    if (FAILED(target->CreateSolidColorBrush(D2DColor(textColor), brush.put())) || ! brush)
    {
        return false;
    }

    const D2D1_RECT_F textRect =
        D2D1::RectF(0.0f, 0.0f, PixelsToDips(static_cast<float>(rcPx.right - rcPx.left), dpi), PixelsToDips(static_cast<float>(rcPx.bottom - rcPx.top), dpi));

    target->BeginDraw();
    target->DrawText(
        text.data(), static_cast<UINT32>(text.size()), format.get(), textRect, brush.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP, DWRITE_MEASURING_MODE_NATURAL);
    const HRESULT endHr = target->EndDraw();
    if (FAILED(endHr))
    {
        Debug::Warning(L"ViewerWeb: DirectWrite status rendering EndDraw failed: 0x{:08X}", endHr);
        return false;
    }
    return true;
}

LRESULT CALLBACK FileComboHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    return RedSalamander::ViewerFileComboHost::DispatchFileComboHostWndProc<ViewerWeb>(
        hwnd, msg, wp, lp, kFileComboHostStateProp, kFileComboHostOriginalWndProcProp, FileComboHostWndProc);
}

[[maybe_unused]] [[nodiscard]] std::wstring KeyGlyphFromVirtualKey(UINT vk, HKL keyboardLayout) noexcept
{
    if (! keyboardLayout)
    {
        return {};
    }

    const UINT scanCode = MapVirtualKeyExW(vk, MAPVK_VK_TO_VSC, keyboardLayout);
    if (scanCode == 0)
    {
        return {};
    }

    std::array<BYTE, 256> keyboardState{};
    std::array<wchar_t, 8> buffer{};
    const int result = ToUnicodeEx(vk, scanCode, keyboardState.data(), buffer.data(), static_cast<int>(buffer.size() - 1), 0, keyboardLayout);

    if (result > 0)
    {
        std::wstring out(buffer.data(), buffer.data() + result);
        if (! out.empty() && ! std::iswcntrl(out.front()))
        {
            return out;
        }
    }
    else if (result < 0)
    {
        std::array<wchar_t, 8> clearBuf{};
        static_cast<void>(ToUnicodeEx(vk, scanCode, keyboardState.data(), clearBuf.data(), static_cast<int>(clearBuf.size() - 1), 0, keyboardLayout));
    }

    wchar_t nameBuf[64]{};
    const LONG lParam = static_cast<LONG>(scanCode << 16);
    const int nameLen = GetKeyNameTextW(lParam, nameBuf, static_cast<int>(std::size(nameBuf)));
    if (nameLen > 0)
    {
        return std::wstring(nameBuf, nameBuf + nameLen);
    }

    return {};
}

[[nodiscard]] UINT VkFromScanCode(UINT scanCode, HKL keyboardLayout) noexcept
{
    if (! keyboardLayout)
    {
        return 0;
    }

    return MapVirtualKeyExW(scanCode, MAPVK_VSC_TO_VK_EX, keyboardLayout);
}

[[nodiscard]] UINT ZoomInVirtualKeyForLayout(HKL keyboardLayout) noexcept
{
    // Number row: key left of Backspace (US: =/+)
    constexpr UINT kScanCode = 0x0Du;
    if (const UINT vk = VkFromScanCode(kScanCode, keyboardLayout); vk != 0)
    {
        return vk;
    }

    return VK_OEM_PLUS;
}

[[nodiscard]] UINT ZoomOutVirtualKeyForLayout(HKL keyboardLayout) noexcept
{
    // Number row: key right of '0' (US: -/_)
    constexpr UINT kScanCode = 0x0Cu;
    if (const UINT vk = VkFromScanCode(kScanCode, keyboardLayout); vk != 0)
    {
        return vk;
    }

    return VK_OEM_MINUS;
}

using unique_yyjson_doc     = wil::unique_any<yyjson_doc*, decltype(&yyjson_doc_free), yyjson_doc_free>;
using unique_yyjson_mut_doc = wil::unique_any<yyjson_mut_doc*, decltype(&yyjson_mut_doc_free), yyjson_mut_doc_free>;
using unique_malloc_string  = wil::unique_any<char*, decltype(&::free), ::free>;

template <typename TInterface, typename TFn, typename... TInvokeArgs> class ComCallback final : public TInterface
{
public:
    explicit ComCallback(TFn fn) noexcept : _fn(std::move(fn))
    {
        const DWORD currentThread = GetCurrentThreadId();
        DWORD ownerThread         = 0u;
        if (! g_comCallbackOwnerThreadId.compare_exchange_strong(
                ownerThread, currentThread, std::memory_order_acq_rel, std::memory_order_acquire) &&
            ownerThread != currentThread)
        {
            // WebView2 callbacks are STA-affine. If that contract is ever
            // violated, retain the module rather than claim a safe runtime
            // unload from a thread that cannot prove the Release stack is gone.
            g_comCallbackThreadViolation.store(true, std::memory_order_release);
        }
        g_liveComCallbackCount.fetch_add(1u, std::memory_order_relaxed);
    }

    ComCallback(const ComCallback&)            = delete;
    ComCallback& operator=(const ComCallback&) = delete;
    ComCallback(ComCallback&&)                 = delete;
    ComCallback& operator=(ComCallback&&)      = delete;

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) noexcept override
    {
        if (ppvObject == nullptr)
        {
            return E_POINTER;
        }

        *ppvObject = nullptr;

        if (riid == __uuidof(IUnknown) || riid == __uuidof(TInterface))
        {
            *ppvObject = static_cast<TInterface*>(this);
            AddRef();
            return S_OK;
        }

        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() noexcept override
    {
        return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
    }

    ULONG STDMETHODCALLTYPE Release() noexcept override
    {
        const ULONG refs = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
        if (refs == 0)
        {
            if (GetCurrentThreadId() != g_comCallbackOwnerThreadId.load(std::memory_order_acquire))
            {
                g_comCallbackThreadViolation.store(true, std::memory_order_release);
            }
            delete this;
            // WebView2 creates, invokes, and releases these callbacks on the
            // recorded owner STA. Publish a new release epoch before dropping
            // the live count; CanUnloadNow also requires that same STA and one
            // later poll, so the final Release stack has returned first.
            g_comCallbackReleaseEpoch.fetch_add(1u, std::memory_order_release);
            g_liveComCallbackCount.fetch_sub(1u, std::memory_order_release);
        }
        return refs;
    }

    HRESULT STDMETHODCALLTYPE Invoke(TInvokeArgs... args) noexcept override
    {
        if (GetCurrentThreadId() != g_comCallbackOwnerThreadId.load(std::memory_order_acquire))
        {
            g_comCallbackThreadViolation.store(true, std::memory_order_release);
            return RPC_E_WRONG_THREAD;
        }
        return _fn(args...);
    }

private:
    std::atomic<ULONG> _refCount{1};
    TFn _fn;
};

template <typename TInterface, typename... TInvokeArgs, typename TFn> wil::com_ptr<TInterface> MakeComCallback(TFn&& fn) noexcept
{
    using FnT      = std::decay_t<TFn>;
    auto* callback = new (std::nothrow) ComCallback<TInterface, FnT, TInvokeArgs...>(FnT(std::forward<TFn>(fn)));

    wil::com_ptr<TInterface> out;
    if (callback)
    {
        out.attach(callback);
    }
    return out;
}

using Common::Colors::BlendColorRefTruncate;
using Common::Colors::ColorRefFromArgb;
using Common::Colors::ColorRefFromHsvWrappedHue;
using Common::Colors::StableVisualHash32Utf16V1;

struct HsvColor
{
    float hue = 0.0f;
    float sat = 0.0f;
    float val = 0.0f;
};

[[nodiscard]] HsvColor ColorToHSV(COLORREF color) noexcept
{
    const float r = static_cast<float>(GetRValue(color)) / 255.0f;
    const float g = static_cast<float>(GetGValue(color)) / 255.0f;
    const float b = static_cast<float>(GetBValue(color)) / 255.0f;

    const float maxValue = std::max({r, g, b});
    const float minValue = std::min({r, g, b});
    const float delta    = maxValue - minValue;

    HsvColor out{};
    out.val = maxValue;
    out.sat = maxValue <= 0.0f ? 0.0f : delta / maxValue;

    if (delta <= 0.0001f)
    {
        out.hue = 0.0f;
        return out;
    }

    if (maxValue == r)
    {
        out.hue = 60.0f * std::fmod(((g - b) / delta), 6.0f);
    }
    else if (maxValue == g)
    {
        out.hue = 60.0f * (((b - r) / delta) + 2.0f);
    }
    else
    {
        out.hue = 60.0f * (((r - g) / delta) + 4.0f);
    }

    if (out.hue < 0.0f)
    {
        out.hue += 360.0f;
    }

    return out;
}

struct JsonTokenColors
{
    COLORREF key          = RGB(0, 0, 0);
    COLORREF stringValue  = RGB(0, 0, 0);
    COLORREF numberValue  = RGB(0, 0, 0);
    COLORREF literalValue = RGB(0, 0, 0);
};

[[nodiscard]] COLORREF ThemeAwareSemanticColor(
    float hue, float saturation, float value, COLORREF accent, COLORREF fg, uint8_t accentAlpha, uint8_t fgAlpha) noexcept
{
    const COLORREF semantic   = ColorRefFromHsvWrappedHue(hue, saturation, value);
    const COLORREF harmonized = BlendColorRefTruncate(semantic, accent, accentAlpha);
    return BlendColorRefTruncate(harmonized, fg, fgAlpha);
}

[[nodiscard]] JsonTokenColors BuildJsonTokenColors(COLORREF accent, COLORREF fg, bool darkMode) noexcept
{
    const HsvColor accentHsv = ColorToHSV(accent);
    const float keyHue       = accentHsv.sat > 0.05f ? accentHsv.hue : 195.0f;
    const float keySat       = std::clamp(std::max(accentHsv.sat, darkMode ? 0.60f : 0.72f), 0.0f, 1.0f);
    const float keyVal       = darkMode ? std::max(accentHsv.val, 0.95f) : std::clamp(std::max(accentHsv.val * 0.62f, 0.46f), 0.46f, 0.70f);
    const COLORREF keyBase   = ColorRefFromHsvWrappedHue(keyHue, keySat, keyVal);

    return {
        .key = BlendColorRefTruncate(keyBase, fg, darkMode ? 12u : 18u),
        .stringValue =
            ThemeAwareSemanticColor(145.0f, darkMode ? 0.58f : 0.76f, darkMode ? 0.94f : 0.44f, accent, fg, darkMode ? 24u : 18u, darkMode ? 16u : 18u),
        .numberValue =
            ThemeAwareSemanticColor(32.0f, darkMode ? 0.78f : 0.82f, darkMode ? 0.98f : 0.42f, accent, fg, darkMode ? 20u : 14u, darkMode ? 12u : 14u),
        .literalValue =
            ThemeAwareSemanticColor(282.0f, darkMode ? 0.64f : 0.74f, darkMode ? 0.96f : 0.46f, accent, fg, darkMode ? 18u : 12u, darkMode ? 12u : 16u),
    };
}

COLORREF ResolveAccentColor(const ViewerTheme& theme, std::wstring_view seed) noexcept
{
    if (theme.rainbowMode)
    {
        const uint32_t h = StableVisualHash32Utf16V1(seed);
        const float hue  = static_cast<float>(h % 360u);
        const float sat  = theme.darkBase ? 0.70f : 0.55f;
        const float val  = theme.darkBase ? 0.95f : 0.85f;
        return ColorRefFromHsvWrappedHue(hue, sat, val);
    }

    return ColorRefFromArgb(theme.accentArgb);
}

std::wstring LeafNameFromPath(std::wstring_view path)
{
    if (path.empty())
    {
        return {};
    }

    const size_t slash = path.find_last_of(L"\\/");
    if (slash == std::wstring_view::npos)
    {
        return std::wstring(path);
    }

    return std::wstring(path.substr(slash + 1));
}

// Forward declarations for file-scope helpers defined later in this file.
std::optional<std::filesystem::path> ShowSaveAsDialog(HWND hwnd, std::wstring_view suggestedFileName) noexcept;

struct ViewerWebClassBackgroundBrushState
{
    ViewerWebClassBackgroundBrushState()                                                     = default;
    ViewerWebClassBackgroundBrushState(const ViewerWebClassBackgroundBrushState&)            = delete;
    ViewerWebClassBackgroundBrushState& operator=(const ViewerWebClassBackgroundBrushState&) = delete;
    ViewerWebClassBackgroundBrushState(ViewerWebClassBackgroundBrushState&&)                 = delete;
    ViewerWebClassBackgroundBrushState& operator=(ViewerWebClassBackgroundBrushState&&)      = delete;

    wil::unique_hbrush activeBrush;
    COLORREF activeColor = CLR_INVALID;

    wil::unique_hbrush pendingBrush;
    COLORREF pendingColor = CLR_INVALID;

    bool classRegistered = false;
};

ViewerWebClassBackgroundBrushState g_viewerWebClassBackgroundBrush;

HBRUSH GetActiveViewerWebClassBackgroundBrush() noexcept
{
    if (! g_viewerWebClassBackgroundBrush.activeBrush)
    {
        const COLORREF sys = GetSysColor(COLOR_WINDOW);
        g_viewerWebClassBackgroundBrush.activeBrush.reset(CreateSolidBrush(sys));
        g_viewerWebClassBackgroundBrush.activeColor = sys;
    }
    return g_viewerWebClassBackgroundBrush.activeBrush.get();
}

void RequestViewerWebClassBackgroundColor(COLORREF color) noexcept
{
    if (color == CLR_INVALID)
    {
        return;
    }

    if (g_viewerWebClassBackgroundBrush.activeBrush && g_viewerWebClassBackgroundBrush.activeColor == color)
    {
        return;
    }

    if (g_viewerWebClassBackgroundBrush.pendingBrush && g_viewerWebClassBackgroundBrush.pendingColor == color)
    {
        return;
    }

    g_viewerWebClassBackgroundBrush.pendingBrush.reset(CreateSolidBrush(color));
    g_viewerWebClassBackgroundBrush.pendingColor = color;
}

void ApplyPendingViewerWebClassBackgroundBrush(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    if (! g_viewerWebClassBackgroundBrush.classRegistered)
    {
        return;
    }

    if (! g_viewerWebClassBackgroundBrush.pendingBrush)
    {
        return;
    }

    g_viewerWebClassBackgroundBrush.activeBrush  = std::move(g_viewerWebClassBackgroundBrush.pendingBrush);
    g_viewerWebClassBackgroundBrush.activeColor  = g_viewerWebClassBackgroundBrush.pendingColor;
    g_viewerWebClassBackgroundBrush.pendingColor = CLR_INVALID;

    SetClassLongPtrW(hwnd, GCLP_HBRBACKGROUND, reinterpret_cast<LONG_PTR>(g_viewerWebClassBackgroundBrush.activeBrush.get()));
}

[[nodiscard]] bool ResetViewerWebClassBackgroundBrushAtQuietPoint() noexcept
{
    if (g_viewerWebClassBackgroundBrush.classRegistered)
    {
        if (UnregisterClassW(L"RedSalamander.ViewerWeb", g_hInstance) == FALSE)
        {
            const DWORD error = GetLastError();
            if (error != ERROR_CLASS_DOES_NOT_EXIST)
            {
                Debug::Warning(L"ViewerWeb: class/brush quiet-point cleanup deferred (error={}).", error);
                return false;
            }
        }
        g_viewerWebClassBackgroundBrush.classRegistered = false;
    }

    g_viewerWebClassBackgroundBrush.pendingBrush.reset();
    g_viewerWebClassBackgroundBrush.pendingColor = CLR_INVALID;
    g_viewerWebClassBackgroundBrush.activeBrush.reset();
    g_viewerWebClassBackgroundBrush.activeColor = CLR_INVALID;
    return true;
}

std::string ResourceBytesString(HINSTANCE hinst, UINT id) noexcept
{
    std::vector<std::byte> bytes;
    if (! Localization::LoadResourceBytes(hinst, MAKEINTRESOURCEW(id), RT_RCDATA, bytes))
    {
        return {};
    }

    return {reinterpret_cast<const char*>(bytes.data()), bytes.size()};
}

// ----- Shared WebView2 Environment -----
// A DLL-global environment shared across all viewer instances. Creating a
// WebView2 environment is the most expensive operation (spawns browser
// processes). Sharing it means only the first viewer pays the cost.

struct SharedEnvironmentState
{
    wil::com_ptr<ICoreWebView2Environment> environment;
    bool createInProgress = false;
    bool shuttingDown     = false;
    uint64_t generation   = 1;
    // Pending consumers waiting for the first environment creation to complete.
    struct PendingConsumer
    {
        ViewerWeb* viewer = nullptr;
        HWND hwnd         = nullptr;
    };
    std::vector<PendingConsumer> pendingConsumers;
};

SharedEnvironmentState g_sharedEnvironment;
ViewerWebSecurity::StagedCleanupTracker g_stagedCleanupTracker;

void RetryStagedFileCleanup() noexcept;

[[nodiscard]] uint64_t BeginSharedEnvironmentUse() noexcept
{
    g_shutdownCleanupComplete.store(false, std::memory_order_release);
    g_sharedEnvironment.shuttingDown = false;
    return g_sharedEnvironment.generation;
}

[[nodiscard]] bool IsSharedEnvironmentGenerationCurrent(uint64_t generation) noexcept
{
    return ! g_sharedEnvironment.shuttingDown && g_sharedEnvironment.generation == generation;
}

std::wstring GetWebView2UserDataFolder() noexcept
{
    wchar_t appData[MAX_PATH]{};
    if (FAILED(SHGetFolderPathW(nullptr, CSIDL_LOCAL_APPDATA, nullptr, 0, appData)))
    {
        return {};
    }
    return std::format(L"{}\\RedSalamander\\WebView2UserData", appData);
}

void ResetSharedEnvironmentImpl() noexcept
{
    g_sharedEnvironment.shuttingDown = true;
    ++g_sharedEnvironment.generation;
    if (g_sharedEnvironment.generation == 0)
    {
        g_sharedEnvironment.generation = 1;
    }

    auto pendingConsumers = std::move(g_sharedEnvironment.pendingConsumers);
    g_sharedEnvironment.pendingConsumers.clear();
    for (auto& pending : pendingConsumers)
    {
        if (pending.viewer)
        {
            pending.viewer->CancelPendingWebView2Initialization();
            pending.viewer->Release();
        }
    }

    g_sharedEnvironment.createInProgress = false;
    g_sharedEnvironment.environment.reset();
    RetryStagedFileCleanup();
}

std::wstring UrlFromFilePath(std::wstring_view path)
{
    if (path.empty())
    {
        return {};
    }

    const std::wstring pathCopy(path);
    DWORD capacity = static_cast<DWORD>(pathCopy.size() * 3 + 64);
    capacity       = std::max<DWORD>(capacity, 256u);

    for (int attempt = 0; attempt < 4; ++attempt)
    {
        std::wstring url(static_cast<size_t>(capacity), L'\0');
        DWORD written = capacity;

        const HRESULT hr = UrlCreateFromPathW(pathCopy.c_str(), url.data(), &written, 0);
        if (SUCCEEDED(hr))
        {
            if (! url.empty())
            {
                const size_t len = wcsnlen(url.c_str(), url.size());
                url.resize(len);
            }
            return url;
        }

        if ((hr == E_POINTER || hr == HRESULT_FROM_WIN32(ERROR_INSUFFICIENT_BUFFER)) && written > capacity)
        {
            capacity = written;
            continue;
        }

        break;
    }

    return {};
}

[[nodiscard]] bool TryDeleteStagedFile([[maybe_unused]] void* context, std::wstring_view path) noexcept
{
    if (path.empty())
    {
        return true;
    }

    const std::wstring pathCopy(path);
    if (DeleteFileW(pathCopy.c_str()) != FALSE)
    {
        return true;
    }

    const DWORD deleteError = GetLastError();
    return deleteError == ERROR_FILE_NOT_FOUND || deleteError == ERROR_PATH_NOT_FOUND;
}

[[nodiscard]] bool TryScheduleStagedFileCleanup([[maybe_unused]] void* context, std::wstring_view path) noexcept
{
    if (path.empty())
    {
        return true;
    }

    const std::wstring pathCopy(path);
    return MoveFileExW(pathCopy.c_str(), nullptr, MOVEFILE_DELAY_UNTIL_REBOOT) != FALSE;
}

void RetryStagedFileCleanup() noexcept
{
    const ViewerWebSecurity::StagedCleanupTracker::Operations operations{
        .context       = nullptr,
        .deleteNow     = &TryDeleteStagedFile,
        .scheduleLater = &TryScheduleStagedFileCleanup,
    };
    g_stagedCleanupTracker.Retry(operations);
}

void DeleteStagedFileOrSchedule(const std::filesystem::path& path) noexcept
{
    if (path.empty())
    {
        return;
    }

    g_stagedCleanupTracker.Track(path.wstring());
    RetryStagedFileCleanup();
    if (g_stagedCleanupTracker.PendingCount() != 0u)
    {
        Debug::Warning(L"ViewerWeb: staged PDF cleanup remains queued for a later quiet-point retry.");
    }
}

constexpr unsigned long kProviderReadChunkBytes = 256u * 1024u;

template <typename TConsume>
[[nodiscard]] HRESULT ReadProviderExactly(
    IFileReader* reader, uint64_t expectedBytes, TConsume&& consume, uint64_t& consumedBytes, bool seekToStart = true) noexcept
{
    consumedBytes = 0u;
    if (! reader)
    {
        return E_POINTER;
    }

    if (seekToStart)
    {
        uint64_t position    = 0u;
        const HRESULT seekHr = reader->Seek(0, FILE_BEGIN, &position);
        if (FAILED(seekHr) || position != 0u)
        {
            return FAILED(seekHr) ? seekHr : HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
    }

    std::array<uint8_t, kProviderReadChunkBytes> buffer{};
    while (consumedBytes < expectedBytes)
    {
        const uint64_t remaining = expectedBytes - consumedBytes;
        const unsigned long requested = static_cast<unsigned long>(std::min<uint64_t>(remaining, buffer.size()));
        unsigned long returned         = 0u;
        const HRESULT readHr           = reader->Read(buffer.data(), requested, &returned);
        if (FAILED(readHr))
        {
            return readHr;
        }
        if (! ViewerWebSecurity::IsProviderReadCountValid(requested, returned))
        {
            return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
        }
        if (returned == 0u)
        {
            return HRESULT_FROM_WIN32(ERROR_HANDLE_EOF);
        }

        const HRESULT consumeHr = consume(std::span<const uint8_t>(buffer.data(), returned));
        if (FAILED(consumeHr))
        {
            return consumeHr;
        }
        consumedBytes += static_cast<uint64_t>(returned);
    }

    uint8_t trailingByte      = 0u;
    unsigned long trailingRead = 0u;
    const HRESULT trailingHr   = reader->Read(&trailingByte, 1u, &trailingRead);
    if (FAILED(trailingHr))
    {
        return trailingHr;
    }
    if (! ViewerWebSecurity::IsProviderReadCountValid(1u, trailingRead) || trailingRead != 0u)
    {
        return HRESULT_FROM_WIN32(ERROR_INVALID_DATA);
    }
    return S_OK;
}

constexpr char kInternalHtmlHead[] = "<!doctype html><html><head><meta charset=\"utf-8\">";

[[nodiscard]] std::wstring_view InternalDocumentKindSegment(ViewerWebKind kind) noexcept
{
    switch (kind)
    {
        case ViewerWebKind::Json: return L"json";
        case ViewerWebKind::Markdown: return L"markdown";
        case ViewerWebKind::Web:
        default: return L"web";
    }
}

[[nodiscard]] std::wstring BuildInternalDocumentUrl(ViewerWebKind kind, uint64_t requestId)
{
    return std::format(L"{}/{}/{}.html", ViewerWebSecurity::kInternalDocumentOrigin, InternalDocumentKindSegment(kind), requestId);
}

[[nodiscard]] bool IsInternalDocumentUrl(std::wstring_view url) noexcept
{
    return ViewerWebSecurity::StartsWithNoCase(url, ViewerWebSecurity::kInternalDocumentOrigin);
}

[[nodiscard]] std::string_view TrimAsciiWhitespace(std::string_view text) noexcept
{
    while (! text.empty() && std::isspace(static_cast<unsigned char>(text.front())) != 0)
    {
        text.remove_prefix(1);
    }

    while (! text.empty() && std::isspace(static_cast<unsigned char>(text.back())) != 0)
    {
        text.remove_suffix(1);
    }

    return text;
}

[[nodiscard]] bool IsJsonLinesPath(std::wstring_view path) noexcept
{
    if (path.empty())
    {
        return false;
    }

    const std::wstring extension = std::filesystem::path(path).extension().wstring();
    return OrdinalString::EqualsNoCase(extension, L".jsonl") || OrdinalString::EqualsNoCase(extension, L".ndjson");
}

[[nodiscard]] bool IsPdfPath(std::wstring_view path) noexcept
{
    if (path.empty())
    {
        return false;
    }
    return OrdinalString::EqualsNoCase(std::filesystem::path(path).extension().wstring(), L".pdf");
}

[[nodiscard]] std::string CssRgb(COLORREF c)
{
    return std::format("rgb({},{},{})", GetRValue(c), GetGValue(c), GetBValue(c));
}

struct ScrollbarColors
{
    COLORREF track      = RGB(0, 0, 0);
    COLORREF thumb      = RGB(0, 0, 0);
    COLORREF thumbHover = RGB(0, 0, 0);
    COLORREF corner     = RGB(0, 0, 0);
};

[[nodiscard]] ScrollbarColors BuildScrollbarColors(COLORREF bg, COLORREF fg, COLORREF accent, bool darkMode) noexcept
{
    const COLORREF thumbBase = BlendColorRefTruncate(bg, accent, darkMode ? 52u : 36u);
    return {
        .track      = BlendColorRefTruncate(bg, fg, darkMode ? 20u : 12u),
        .thumb      = BlendColorRefTruncate(thumbBase, fg, darkMode ? 92u : 120u),
        .thumbHover = BlendColorRefTruncate(thumbBase, fg, darkMode ? 128u : 152u),
        .corner     = BlendColorRefTruncate(bg, fg, darkMode ? 16u : 8u),
    };
}

struct JsonLinesEntry
{
    size_t lineNumber = 0;
    std::string timestampText;
    std::string levelText;
    std::string categoryText;
    std::string messageText;
    std::string summaryText;
    std::string prettyJson;
};

[[nodiscard]] std::string Utf8FromWide(std::wstring_view text)
{
    return Common::Strings::Utf8FromUtf16StrictOrEmpty(text);
}

struct JsonLinesUiStrings final
{
    std::string badge;
    std::string title;
    std::string recordSummary;
    std::string expandAll;
    std::string collapseAll;
    std::string genericValue;
};

struct LimitedYyjsonAllocatorState final
{
    size_t committedBytes = 0u;
    size_t limitBytes     = 0u;
};

void* LimitedYyjsonMalloc(void* context, size_t size) noexcept
{
    auto* state = static_cast<LimitedYyjsonAllocatorState*>(context);
    if (! state || state->committedBytes > state->limitBytes || size > state->limitBytes - state->committedBytes)
    {
        return nullptr;
    }
    void* memory = ::malloc(size);
    if (memory)
    {
        state->committedBytes += size;
    }
    return memory;
}

void* LimitedYyjsonRealloc(void* context, void* memory, size_t oldSize, size_t newSize) noexcept
{
    auto* state = static_cast<LimitedYyjsonAllocatorState*>(context);
    if (! state || state->committedBytes > state->limitBytes || oldSize > state->committedBytes)
    {
        return nullptr;
    }
    if (newSize > oldSize && newSize - oldSize > state->limitBytes - state->committedBytes)
    {
        return nullptr;
    }

    void* resized = ::realloc(memory, newSize);
    if (! resized && newSize != 0u)
    {
        return nullptr;
    }
    state->committedBytes = state->committedBytes - oldSize + newSize;
    return resized;
}

void LimitedYyjsonFree(void* /*context*/, void* memory) noexcept
{
    ::free(memory);
}

[[nodiscard]] yyjson_alc MakeLimitedYyjsonAllocator(LimitedYyjsonAllocatorState& state) noexcept
{
    return yyjson_alc{&LimitedYyjsonMalloc, &LimitedYyjsonRealloc, &LimitedYyjsonFree, &state};
}

[[nodiscard]] bool TryJsonStringFromScalar(yyjson_val* value, size_t maxBytes, std::string& output)
{
    output.clear();
    if (! value)
    {
        return true;
    }
    if (yyjson_is_str(value))
    {
        const char* text = yyjson_get_str(value);
        const size_t length = yyjson_get_len(value);
        if (! text || length > maxBytes)
        {
            return false;
        }
        output.assign(text, length);
        return true;
    }

    std::string text;
    if (yyjson_is_int(value))
    {
        text = std::to_string(yyjson_get_int(value));
    }
    else if (yyjson_is_uint(value))
    {
        text = std::to_string(yyjson_get_uint(value));
    }
    else if (yyjson_is_real(value))
    {
        text = std::format("{}", yyjson_get_real(value));
    }
    else if (yyjson_is_bool(value))
    {
        text = yyjson_get_bool(value) ? "true" : "false";
    }
    else if (yyjson_is_null(value))
    {
        text = "null";
    }

    if (text.size() > maxBytes)
    {
        return false;
    }
    output = std::move(text);
    return true;
}

[[nodiscard]] std::string DescribeJsonValue(yyjson_val* value)
{
    if (! value)
    {
        return Utf8FromWide(LoadStringResource(g_hInstance, IDS_VIEWERWEB_JSONL_VALUE_GENERIC));
    }
    if (yyjson_is_obj(value))
    {
        const size_t fieldCount = yyjson_obj_size(value);
        return Utf8FromWide(FormatStringResource(g_hInstance,
                                               fieldCount == 1u ? IDS_VIEWERWEB_JSONL_OBJECT_ONE_FMT : IDS_VIEWERWEB_JSONL_OBJECT_MANY_FMT,
                                               fieldCount));
    }
    if (yyjson_is_arr(value))
    {
        const size_t itemCount = yyjson_arr_size(value);
        return Utf8FromWide(FormatStringResource(g_hInstance,
                                               itemCount == 1u ? IDS_VIEWERWEB_JSONL_ARRAY_ONE_FMT : IDS_VIEWERWEB_JSONL_ARRAY_MANY_FMT,
                                               itemCount));
    }
    if (yyjson_is_str(value))
    {
        return Utf8FromWide(LoadStringResource(g_hInstance, IDS_VIEWERWEB_JSONL_VALUE_STRING));
    }
    if (yyjson_is_bool(value))
    {
        return Utf8FromWide(LoadStringResource(
            g_hInstance, yyjson_get_bool(value) ? IDS_VIEWERWEB_JSONL_VALUE_BOOLEAN_TRUE : IDS_VIEWERWEB_JSONL_VALUE_BOOLEAN_FALSE));
    }
    if (yyjson_is_null(value))
    {
        return Utf8FromWide(LoadStringResource(g_hInstance, IDS_VIEWERWEB_JSONL_VALUE_NULL));
    }
    if (yyjson_is_num(value))
    {
        std::string number;
        static_cast<void>(TryJsonStringFromScalar(value, 64u, number));
        return Utf8FromWide(FormatStringResource(
            g_hInstance, IDS_VIEWERWEB_JSONL_VALUE_NUMERIC_FMT, std::wstring(number.begin(), number.end())));
    }
    return Utf8FromWide(LoadStringResource(g_hInstance, IDS_VIEWERWEB_JSONL_VALUE_GENERIC));
}

[[nodiscard]] bool TryReadJsonObjectSummaryValue(
    yyjson_val* value, std::initializer_list<const char*> keys, size_t maxBytes, std::string& output)
{
    output.clear();
    if (! yyjson_is_obj(value))
    {
        return true;
    }

    for (const char* key : keys)
    {
        std::string text;
        if (! TryJsonStringFromScalar(yyjson_obj_get(value, key), maxBytes, text))
        {
            return false;
        }
        if (! text.empty())
        {
            output = std::move(text);
            return true;
        }
    }

    return true;
}

[[nodiscard]] bool TryParseJsonLinesEntries(
    std::string_view textUtf8, bool allowSingleEntry, size_t publicationLimit, std::vector<JsonLinesEntry>& outEntries)
{
    outEntries.clear();
    constexpr size_t kAbsoluteEntryLimit = 100'000u;
    constexpr size_t kMinimumPublishedBytesPerEntry = 128u;
    const size_t maxEntries = std::min(kAbsoluteEntryLimit, publicationLimit / kMinimumPublishedBytesPerEntry);
    const size_t retainedLimit = publicationLimit / 2u;
    if (maxEntries == 0u || retainedLimit == 0u)
    {
        return false;
    }
    outEntries.reserve(std::min<size_t>(maxEntries, 4096u));

    size_t lineNumber = 1;
    size_t offset     = 0;
    size_t retainedBytes = 0u;
    while (offset < textUtf8.size())
    {
        size_t lineEnd = textUtf8.find('\n', offset);
        if (lineEnd == std::string::npos)
        {
            lineEnd = textUtf8.size();
        }

        std::string_view line(textUtf8.data() + offset, lineEnd - offset);
        if (! line.empty() && line.back() == '\r')
        {
            line.remove_suffix(1);
        }
        line = TrimAsciiWhitespace(line);

        if (! line.empty())
        {
            if (outEntries.size() >= maxEntries || retainedBytes > retainedLimit || line.size() > retainedLimit - retainedBytes)
            {
                outEntries.clear();
                return false;
            }

            std::string mutableLine(line);
            yyjson_read_err lineErr{};
            LimitedYyjsonAllocatorState readAllocatorState{.limitBytes = retainedLimit - retainedBytes};
            yyjson_alc readAllocator = MakeLimitedYyjsonAllocator(readAllocatorState);
            unique_yyjson_doc lineDoc(
                yyjson_read_opts(mutableLine.data(), mutableLine.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, &readAllocator, &lineErr));
            if (! lineDoc)
            {
                outEntries.clear();
                return false;
            }

            yyjson_val* root = yyjson_doc_get_root(lineDoc.get());
            if (! root)
            {
                outEntries.clear();
                return false;
            }

            size_t prettyLen = 0;
            LimitedYyjsonAllocatorState writeAllocatorState{.limitBytes = retainedLimit - retainedBytes};
            yyjson_alc writeAllocator = MakeLimitedYyjsonAllocator(writeAllocatorState);
            unique_malloc_string pretty(
                yyjson_write_opts(lineDoc.get(), YYJSON_WRITE_PRETTY | YYJSON_WRITE_ESCAPE_UNICODE, &writeAllocator, &prettyLen, nullptr));
            if (! pretty)
            {
                outEntries.clear();
                return false;
            }

            JsonLinesEntry entry{};
            entry.lineNumber = lineNumber;
            if (prettyLen > retainedLimit - retainedBytes)
            {
                outEntries.clear();
                return false;
            }
            entry.prettyJson.assign(pretty.get(), prettyLen);
            retainedBytes += prettyLen;

            const auto readSummary = [&](std::initializer_list<const char*> keys, std::string& value) -> bool
            {
                const size_t remaining = retainedLimit - retainedBytes;
                if (! TryReadJsonObjectSummaryValue(root, keys, remaining, value) || value.size() > remaining)
                {
                    return false;
                }
                retainedBytes += value.size();
                return true;
            };
            if (! readSummary({"ts", "timestamp", "@timestamp", "time"}, entry.timestampText) ||
                ! readSummary({"level", "severity", "lvl", "type"}, entry.levelText) ||
                ! readSummary({"category", "op", "operation", "event", "logger"}, entry.categoryText) ||
                ! readSummary({"message", "msg", "text", "summary", "description"}, entry.messageText))
            {
                outEntries.clear();
                return false;
            }
            entry.summaryText   = DescribeJsonValue(root);
            if (entry.summaryText.size() > retainedLimit - retainedBytes)
            {
                outEntries.clear();
                return false;
            }
            retainedBytes += entry.summaryText.size();
            if (entry.messageText.empty())
            {
                if (! readSummary({"name", "srcLeaf", "dstLeaf", "path"}, entry.messageText))
                {
                    outEntries.clear();
                    return false;
                }
            }

            outEntries.push_back(std::move(entry));
        }

        offset = lineEnd < textUtf8.size() ? lineEnd + 1 : lineEnd;
        ++lineNumber;
    }

    const size_t minimumEntries = allowSingleEntry ? 1u : 2u;
    if (outEntries.size() < minimumEntries)
    {
        outEntries.clear();
        return false;
    }

    return true;
}

// Common theme helper JS shared by all internal HTML templates.
constexpr char kCommonThemeJs[] =
    "function parseRgb(s){const m=/rgb\\((\\d+),(\\d+),(\\d+)\\)/.exec(s.replace(/\\s+/g,''));return m?{r:+m[1],g:+m[2],b:+m[3]}:{r:0,g:0,b:0};}"
    "function rgb(c){return `rgb(${c.r},${c.g},${c.b})`;}"
    "function clamp(v,min,max){return Math.min(max,Math.max(min,v));}"
    "function blend(u,o,a){const inv=255-a;return {r:Math.round((u.r*inv+o.r*a)/255),g:Math.round((u.g*inv+o.g*a)/255),b:Math.round((u.b*inv+o.b*a)/255)};}"
    "function luma(c){return (c.r*299+c.g*587+c.b*114)/1000;}"
    "function hsv(h,s,v){const hue=((h%360)+360)%360;const c=v*s;const x=c*(1-Math.abs((hue/60)%2-1));const m=v-c;let rf=0,gf=0,bf=0;"
    "if(hue<60){rf=c;gf=x;}else if(hue<120){rf=x;gf=c;}else if(hue<180){gf=c;bf=x;}else if(hue<240){gf=x;bf=c;}else if(hue<300){rf=x;bf=c;}else{rf=c;bf=x;}"
    "return {r:Math.round((rf+m)*255),g:Math.round((gf+m)*255),b:Math.round((bf+m)*255)};}"
    "function rgbToHsv(c){const r=c.r/255,g=c.g/255,b=c.b/255;const max=Math.max(r,g,b),min=Math.min(r,g,b),delta=max-min;let h=0;"
    "if(delta>0){if(max===r){h=60*(((g-b)/delta)%6);}else if(max===g){h=60*(((b-r)/delta)+2);}else{h=60*(((r-g)/delta)+4);}if(h<0){h+=360;}}"
    "return {h,s:max<=0?0:delta/max,v:max};}"
    "function setScrollbarVars(r,bg,fg,acc){const dark=luma(bg)<128;const thumbBase=blend(bg,acc,dark?52:36);"
    "r.setProperty('--rs-scroll-track',rgb(blend(bg,fg,dark?20:12)));r.setProperty('--rs-scroll-thumb',rgb(blend(thumbBase,fg,dark?92:120)));"
    "r.setProperty('--rs-scroll-thumb-hover',rgb(blend(thumbBase,fg,dark?128:152)));r.setProperty('--rs-scroll-corner',rgb(blend(bg,fg,dark?16:8)));}"
    "function setJsonTokenVars(r,bg,fg,acc){const dark=luma(bg)<128;"
    "const accentHsv=rgbToHsv(acc);const keyHue=accentHsv.s>0.05?accentHsv.h:195;const keySat=clamp(Math.max(accentHsv.s,dark?0.60:0.72),0,1);"
    "const keyVal=dark?Math.max(accentHsv.v,0.95):clamp(Math.max(accentHsv.v*0.62,0.46),0.46,0.70);const key=blend(hsv(keyHue,keySat,keyVal),fg,dark?12:18);"
    "const str=blend(blend(hsv(145,dark?0.58:0.76,dark?0.94:0.44),acc,dark?24:18),fg,dark?16:18);"
    "const num=blend(blend(hsv(32,dark?0.78:0.82,dark?0.98:0.42),acc,dark?20:14),fg,dark?12:14);"
    "const lit=blend(blend(hsv(282,dark?0.64:0.74,dark?0.96:0.46),acc,dark?18:12),fg,dark?12:16);"
    "r.setProperty('--rs-key',rgb(key));r.setProperty('--rs-string',rgb(str));r.setProperty('--rs-number',rgb(num));r.setProperty('--rs-literal',rgb(lit));}";

constexpr char kCommonScrollbarCss[] =
    "html{scrollbar-color:var(--rs-scroll-thumb) var(--rs-scroll-track);}"
    "*{scrollbar-color:var(--rs-scroll-thumb) var(--rs-scroll-track);scrollbar-width:thin;}"
    "*::-webkit-scrollbar{width:12px;height:12px;background:var(--rs-scroll-track);}"
    "*::-webkit-scrollbar-track{background:var(--rs-scroll-track);}"
    "*::-webkit-scrollbar-thumb{background:var(--rs-scroll-thumb);border-radius:999px;border:3px solid var(--rs-scroll-track);}"
    "*::-webkit-scrollbar-thumb:hover{background:var(--rs-scroll-thumb-hover);}"
    "*::-webkit-scrollbar-button{display:none;width:0;height:0;}"
    "*::-webkit-scrollbar-corner{background:var(--rs-scroll-corner);}";

[[nodiscard]] bool BuildJsonLinesHtml(const std::vector<JsonLinesEntry>& entries,
                                      const JsonLinesUiStrings& strings,
                                      std::string_view highlightJs,
                                      std::string_view themeObj,
                                      COLORREF bg,
                                      COLORREF fg,
                                      COLORREF selBg,
                                      COLORREF selFg,
                                      COLORREF accent,
                                      bool darkMode,
                                      size_t publicationLimit,
                                      std::string& html)
{
    const COLORREF codeBg             = BlendColorRefTruncate(bg, fg, darkMode ? 18u : 8u);
    const COLORREF cardBg             = BlendColorRefTruncate(bg, fg, darkMode ? 12u : 5u);
    const COLORREF cardBgOpen         = BlendColorRefTruncate(bg, fg, darkMode ? 18u : 10u);
    const COLORREF border             = BlendColorRefTruncate(bg, fg, darkMode ? 36u : 58u);
    const COLORREF mutedFg            = BlendColorRefTruncate(bg, fg, 140u);
    const JsonTokenColors tokenColors = BuildJsonTokenColors(accent, fg, darkMode);
    const ScrollbarColors scrollbar   = BuildScrollbarColors(bg, fg, accent, darkMode);

    html.clear();
    if (highlightJs.size() > publicationLimit || publicationLimit - highlightJs.size() < 16384u)
    {
        return false;
    }
    html.reserve(std::min(publicationLimit, highlightJs.size() + 16384u));
    html += kInternalHtmlHead;
    html += "<style>";
    html += ":root{--rs-bg:" + CssRgb(bg) + ";--rs-fg:" + CssRgb(fg) + ";--rs-sel-bg:" + CssRgb(selBg) + ";--rs-sel-fg:" + CssRgb(selFg) +
            ";--rs-accent:" + CssRgb(accent) + ";--rs-code-bg:" + CssRgb(codeBg) + ";--rs-card-bg:" + CssRgb(cardBg) +
            ";--rs-card-bg-open:" + CssRgb(cardBgOpen) + ";--rs-border:" + CssRgb(border) + ";--rs-muted-fg:" + CssRgb(mutedFg) +
            ";--rs-key:" + CssRgb(tokenColors.key) + ";--rs-string:" + CssRgb(tokenColors.stringValue) + ";--rs-number:" + CssRgb(tokenColors.numberValue) +
            ";--rs-literal:" + CssRgb(tokenColors.literalValue) + ";--rs-scroll-track:" + CssRgb(scrollbar.track) +
            ";--rs-scroll-thumb:" + CssRgb(scrollbar.thumb) + ";--rs-scroll-thumb-hover:" + CssRgb(scrollbar.thumbHover) +
            ";--rs-scroll-corner:" + CssRgb(scrollbar.corner) + ";}";
    html += "html,body{height:100%;margin:0;}*,*::before,*::after{box-sizing:border-box;}body{background:var(--rs-bg);color:var(--rs-fg);"
            "font-family:\"Segoe UI Variable Text\",\"Segoe UI Variable Small\",\"Segoe UI\",sans-serif;overflow:hidden;}";
    html += kCommonScrollbarCss;
    html += "::selection{background:var(--rs-sel-bg);color:var(--rs-sel-fg);}.rs-shell{height:100%;min-height:0;display:flex;flex-direction:column;}";
    html += ".rs-toolbar{flex:none;display:flex;align-items:center;justify-content:space-between;gap:12px;padding:12px 14px;"
            "background:linear-gradient(180deg,var(--rs-bg),var(--rs-card-bg));border-bottom:1px solid var(--rs-border);}";
    html += ".rs-toolbar-left{display:flex;align-items:center;gap:10px;min-width:0;}.rs-toolbar-title{font-size:13px;font-weight:600;letter-spacing:0.02em;}"
            ".rs-toolbar-meta{font-size:12px;color:var(--rs-muted-fg);white-space:nowrap;}.rs-toolbar-actions{display:flex;align-items:center;gap:8px;}";
    html += ".rs-btn,.rs-pill{border:1px solid var(--rs-border);background:var(--rs-card-bg);color:var(--rs-fg);border-radius:999px;}"
            ".rs-btn{padding:6px 10px;font:inherit;cursor:pointer;}.rs-btn:hover{background:var(--rs-card-bg-open);}."
            "rs-pill{padding:4px 9px;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.05em;}";
    html += ".rs-list{flex:1 1 auto;min-height:0;padding:14px;display:flex;flex-direction:column;gap:10px;overflow:auto;"
            "overscroll-behavior:contain;scrollbar-gutter:stable both-edges;}.rs-entry{flex:none;border:1px solid var(--rs-border);"
            "border-radius:10px;background:var(--rs-card-bg);overflow:hidden;}.rs-entry.is-open{background:var(--rs-card-bg-open);}";
    html += ".rs-entry-summary{width:100%;display:flex;align-items:center;gap:8px;padding:10px 12px;border:0;background:transparent;color:inherit;"
            "cursor:pointer;text-align:left;font:inherit;line-height:1.25;appearance:none;}.rs-entry-summary:hover{background:rgba(255,255,255,0.02);}"
            ".rs-entry-summary:focus-visible{outline:2px solid var(--rs-accent);outline-offset:-2px;}.rs-entry-summary>*{min-width:0;}"
            ".rs-entry-summary::before{content:'+';display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;"
            "border-radius:999px;background:var(--rs-code-bg);color:var(--rs-muted-fg);font-size:12px;font-weight:700;flex:none;}";
    html += ".rs-entry.is-open .rs-entry-summary::before{content:'-';color:var(--rs-accent);}.rs-line-pill{min-width:56px;text-align:center;}";
    html += ".rs-badge{padding:4px 8px;border-radius:999px;font-size:11px;font-weight:700;border:1px solid var(--rs-border);background:var(--rs-code-bg);"
            "color:var(--rs-muted-fg);text-transform:uppercase;letter-spacing:0.04em;}";
    html += ".rs-badge[data-kind='level'][data-level='error']{background:rgba(190,64,64,0.22);color:rgb(255,210,210);}."
            "rs-badge[data-kind='level'][data-level='warning']{background:rgba(190,140,40,0.24);color:rgb(255,230,170);}."
            "rs-badge[data-kind='level'][data-level='info']{background:rgba(64,116,190,0.22);color:rgb(210,225,255);}."
            "rs-badge[data-kind='level'][data-level='debug']{background:rgba(120,120,140,0.22);color:rgb(220,220,230);}";
    html += ".rs-summary-text{min-width:0;flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-size:13px;}"
            ".rs-entry-body{padding:0 12px 12px 42px;}.rs-entry-body[hidden]{display:none;}.rs-entry-body pre{margin:0;max-width:100%;"
            "background:var(--rs-code-bg);border:1px solid var(--rs-border);padding:12px;"
            "overflow:auto;border-radius:8px;}.rs-entry-body code{font-family:Consolas,ui-monospace,monospace;font-size:13px;line-height:1.45;}";
    html += ".hljs{background:transparent;}.hljs-attr{color:var(--rs-key);} .hljs-string{color:var(--rs-string);} .hljs-number{color:var(--rs-number);} "
            ".hljs-literal{color:var(--rs-literal);} .hljs-punctuation,.hljs-brace{color:var(--rs-muted-fg);} .hljs-comment{opacity:0.8;}";
    html += "@media (max-width:720px){.rs-toolbar{flex-direction:column;align-items:stretch;}.rs-toolbar-actions{justify-content:flex-end;}.rs-entry-summary"
            "{flex-wrap:wrap;}.rs-entry-body{padding-left:12px;}}";
    html += "</style></head><body><div class=\"rs-shell\"><div class=\"rs-toolbar\">";
    html += "<div class=\"rs-toolbar-left\"><span id=\"typeBadge\" class=\"rs-pill\"></span><span id=\"viewTitle\" class=\"rs-toolbar-title\"></span>"
            "<span id=\"summary\" class=\"rs-toolbar-meta\"></span></div>";
    html += "<div class=\"rs-toolbar-actions\"><button id=\"expandAll\" type=\"button\" class=\"rs-btn\"></button>"
            "<button id=\"collapseAll\" type=\"button\" class=\"rs-btn\"></button></div>";
    html += "</div><div id=\"list\" class=\"rs-list\"></div></div>";
    html += "<script>";
    html.append(highlightJs);
    html += "</script><script>";
    html += "(() => {const initialTheme=" + std::string(themeObj) + ";";
    html += kCommonThemeJs;
    html += "function applyTheme(t){const r=document.documentElement.style;r.setProperty('--rs-bg',t.bg);r.setProperty('--rs-fg',t.fg);"
            "r.setProperty('--rs-sel-bg',t.selBg);r.setProperty('--rs-sel-fg',t.selFg);r.setProperty('--rs-accent',t.accent);"
            "const bg=parseRgb(t.bg),fg=parseRgb(t.fg),acc=parseRgb(t.accent);const dark=luma(bg)<128;"
            "r.setProperty('--rs-code-bg',rgb(blend(bg,fg,dark?18:8)));r.setProperty('--rs-card-bg',rgb(blend(bg,fg,dark?12:5)));"
            "r.setProperty('--rs-card-bg-open',rgb(blend(bg,fg,dark?18:10)));r.setProperty('--rs-border',rgb(blend(bg,fg,dark?36:58)));"
            "r.setProperty('--rs-muted-fg',rgb(blend(bg,fg,140)));setJsonTokenVars(r,bg,fg,acc);setScrollbarVars(r,bg,fg,acc);}";
    const auto appendPublished = [&](std::string_view value) noexcept
    { return ViewerWebDetail::TryAppendWithinLimit(html, value, publicationLimit); };
    const auto appendEscaped = [&](std::string_view value) noexcept
    { return ViewerWebDetail::TryAppendEscapedJavaScriptStringUtf8(value, html, publicationLimit); };

    if (! appendPublished("const labels={badge:'") || ! appendEscaped(strings.badge) || ! appendPublished("',title:'") ||
        ! appendEscaped(strings.title) || ! appendPublished("',summary:'") || ! appendEscaped(strings.recordSummary) ||
        ! appendPublished("',expand:'") || ! appendEscaped(strings.expandAll) || ! appendPublished("',collapse:'") ||
        ! appendEscaped(strings.collapseAll) || ! appendPublished("',value:'") || ! appendEscaped(strings.genericValue) ||
        ! appendPublished("'};document.getElementById('typeBadge').textContent=labels.badge;"
                          "document.getElementById('viewTitle').textContent=labels.title;"
                          "document.getElementById('summary').textContent=labels.summary;"
                          "document.getElementById('expandAll').textContent=labels.expand;"
                          "document.getElementById('collapseAll').textContent=labels.collapse;const entries=["))
    {
        html.clear();
        return false;
    }
    for (size_t index = 0u; index < entries.size(); ++index)
    {
        const JsonLinesEntry& entry = entries[index];
        const std::string lineNumber = std::to_string(entry.lineNumber);
        if ((index != 0u && ! appendPublished(",")) || ! appendPublished("{line:") || ! appendPublished(lineNumber) || ! appendPublished(",ts:'") ||
            ! appendEscaped(entry.timestampText) || ! appendPublished("',level:'") || ! appendEscaped(entry.levelText) ||
            ! appendPublished("',category:'") || ! appendEscaped(entry.categoryText) || ! appendPublished("',message:'") ||
            ! appendEscaped(entry.messageText) || ! appendPublished("',summary:'") || ! appendEscaped(entry.summaryText) ||
            ! appendPublished("',json:'") || ! appendEscaped(entry.prettyJson) || ! appendPublished("'}"))
        {
            html.clear();
            return false;
        }
    }
    if (! appendPublished("];const list=document.getElementById('list');"))
    {
        html.clear();
        return false;
    }
    if (! appendPublished(
            "function makeBadge(text,kind,level){if(!text){return null;}const el=document.createElement('span');el.className='rs-badge';"
            "el.textContent=text;el.dataset.kind=kind;if(level){el.dataset.level=String(level).toLowerCase();}return el;}"
            "function renderCode(code,text){code.textContent=text;try{hljs.highlightElement(code);}catch(e){}}"
            "function setEntryOpen(entryEl,open,scrollIntoView){if(!entryEl){return;}entryEl.classList.toggle('is-open',open);"
            "const header=entryEl.querySelector('.rs-entry-summary');const body=entryEl.querySelector('.rs-entry-body');"
            "if(header){header.setAttribute('aria-expanded',open?'true':'false');}if(body){body.hidden=!open;}"
            "if(open&&typeof entryEl._ensureRendered==='function'){entryEl._ensureRendered();if(scrollIntoView){entryEl.scrollIntoView({block:'nearest'});}}}"
            "function makeEntry(entry,index){const entryEl=document.createElement('article');entryEl.className='rs-entry';"
            "const summary=document.createElement('button');summary.type='button';summary.className='rs-entry-summary';summary.setAttribute('aria-expanded','false');"
            "const linePill=document.createElement('span');linePill.className='rs-pill rs-line-pill';linePill.textContent=`#${entry.line}`;summary.appendChild(linePill);"
            "if(entry.ts){const ts=document.createElement('span');ts.className='rs-toolbar-meta';ts.textContent=entry.ts;summary.appendChild(ts);}"
            "const levelBadge=makeBadge(entry.level,'level',entry.level);if(levelBadge){summary.appendChild(levelBadge);}"
            "const categoryBadge=makeBadge(entry.category,'category','');if(categoryBadge){summary.appendChild(categoryBadge);}"
            "const text=document.createElement('span');text.className='rs-summary-text';text.textContent=entry.message||entry.summary||labels.value;"
            "summary.appendChild(text);entryEl.appendChild(summary);const body=document.createElement('div');body.className='rs-entry-body';body.hidden=true;"
            "const pre=document.createElement('pre');const code=document.createElement('code');code.className='language-json';pre.appendChild(code);"
            "body.appendChild(pre);entryEl.appendChild(body);let rendered=false;entryEl._ensureRendered=()=>{if(rendered){return;}renderCode(code,entry.json);rendered=true;};"
            "summary.addEventListener('click',()=>setEntryOpen(entryEl,!entryEl.classList.contains('is-open'),true));"
            "if(index<2){setEntryOpen(entryEl,true,false);}return entryEl;}"
            "function renderList(){const frag=document.createDocumentFragment();entries.forEach((entry,index)=>frag.appendChild(makeEntry(entry,index)));"
            "list.replaceChildren(frag);}function expandAll(){document.querySelectorAll('.rs-entry').forEach((entryEl)=>setEntryOpen(entryEl,true,false));}"
            "function collapseAll(){document.querySelectorAll('.rs-entry').forEach((entryEl)=>setEntryOpen(entryEl,false,false));}"
            "document.getElementById('expandAll').addEventListener('click',expandAll);document.getElementById('collapseAll').addEventListener('click',collapseAll);"
            "window.RS={applyTheme:applyTheme,expandAll:expandAll,collapseAll:collapseAll};applyTheme(initialTheme);renderList();})();"
            "</script></body></html>"))
    {
        html.clear();
        return false;
    }
    return true;
}

constexpr char kViewerWebSchemaJson[] = R"json({
    "version": 1,
    "title": "Web Viewer",
    "fields": [
        {
            "key": "maxDocumentMiB",
            "type": "value",
            "label": "Max document size (MiB)",
            "description": "Maximum size for private-origin HTML and staged PDF loads.",
            "default": 32,
            "min": 1,
            "max": 64
        },
        {
            "key": "allowExternalNavigation",
            "type": "option",
            "label": "External navigation",
            "description": "Allow user-activated http/https links to open in the system browser.",
            "default": "0",
            "options": [
                { "value": "0", "label": "Block" },
                { "value": "1", "label": "Allow" }
            ]
        },
        {
            "key": "devToolsEnabled",
            "type": "option",
            "label": "DevTools",
            "description": "Allow opening DevTools for the viewer WebView2 instance.",
            "default": "0",
            "options": [
                { "value": "0", "label": "Off" },
                { "value": "1", "label": "On" }
            ]
        }
    ]
})json";

constexpr char kViewerJsonSchemaJson[] = R"json({
    "version": 1,
    "title": "JSON Viewer",
    "fields": [
        {
            "key": "maxDocumentMiB",
            "type": "value",
            "label": "Max document size (MiB)",
            "description": "Maximum size for in-memory loads.",
            "default": 32,
            "min": 1,
            "max": 64
        },
        {
            "key": "viewMode",
            "type": "option",
            "label": "View mode",
            "description": "Pretty highlighted text, interactive tree view, or JSONL log cards.",
            "default": "pretty",
            "options": [
                { "value": "pretty", "label": "Pretty" },
                { "value": "tree", "label": "Tree" },
                { "value": "jsonl", "label": "JSONL" }
            ]
        },
        {
            "key": "devToolsEnabled",
            "type": "option",
            "label": "DevTools",
            "description": "Allow opening DevTools for the viewer WebView2 instance.",
            "default": "0",
            "options": [
                { "value": "0", "label": "Off" },
                { "value": "1", "label": "On" }
            ]
        }
    ]
})json";

constexpr char kViewerMarkdownSchemaJson[] = R"json({
    "version": 1,
    "title": "Markdown Viewer",
    "fields": [
        {
            "key": "maxDocumentMiB",
            "type": "value",
            "label": "Max document size (MiB)",
            "description": "Maximum size for in-memory loads.",
            "default": 32,
            "min": 1,
            "max": 64
        },
        {
            "key": "allowExternalNavigation",
            "type": "option",
            "label": "External navigation",
            "description": "Allow user-activated http/https links to open in the system browser.",
            "default": "0",
            "options": [
                { "value": "0", "label": "Block" },
                { "value": "1", "label": "Allow" }
            ]
        },
        {
            "key": "devToolsEnabled",
            "type": "option",
            "label": "DevTools",
            "description": "Allow opening DevTools for the viewer WebView2 instance.",
            "default": "0",
            "options": [
                { "value": "0", "label": "Off" },
                { "value": "1", "label": "On" }
            ]
        }
    ]
})json";

[[nodiscard]] const char* GetViewerWebStaticConfigurationSchemaImpl(ViewerWebKind kind) noexcept
{
    switch (kind)
    {
        case ViewerWebKind::Json: return kViewerJsonSchemaJson;
        case ViewerWebKind::Markdown: return kViewerMarkdownSchemaJson;
        case ViewerWebKind::Web:
        default: return kViewerWebSchemaJson;
    }
}

// ----- Cached Resource Helpers -----
// JS/CSS resource data is immutable after load. The helpers below avoid
// repeated allocations by caching derived data (base64-encoded icons, CSS with
// icons inlined, and the common theme helper JS shared by all three HTML
// templates).

std::string_view GetHighlightJs() noexcept
{
    static const std::string s = ResourceBytesString(g_hInstance, IDR_VIEWERWEB_HIGHLIGHT_JS);
    return {s.data(), s.size()};
}

std::string_view GetMarkdownItJs() noexcept
{
    static const std::string s = ResourceBytesString(g_hInstance, IDR_VIEWERWEB_MARKDOWNIT_JS);
    return {s.data(), s.size()};
}

std::string_view GetJsonEditorJs() noexcept
{
    static const std::string s = ResourceBytesString(g_hInstance, IDR_VIEWERWEB_JSONEDITOR_JS);
    return {s.data(), s.size()};
}

// Returns the jsoneditor CSS with SVG icon URLs already inlined as data-URIs.
const std::string& GetJsonEditorCssWithIcons() noexcept
{
    static const std::string s = []
    {
        auto base64Encode = [](std::string_view bytes) noexcept -> std::string
        {
            static constexpr char kTable[] = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
            std::string out;
            out.reserve(((bytes.size() + 2) / 3) * 4);
            size_t i = 0;
            for (; i + 3 <= bytes.size(); i += 3)
            {
                const uint32_t n = (static_cast<uint32_t>(static_cast<uint8_t>(bytes[i])) << 16) |
                                   (static_cast<uint32_t>(static_cast<uint8_t>(bytes[i + 1])) << 8) | static_cast<uint32_t>(static_cast<uint8_t>(bytes[i + 2]));
                out.push_back(kTable[(n >> 18) & 0x3F]);
                out.push_back(kTable[(n >> 12) & 0x3F]);
                out.push_back(kTable[(n >> 6) & 0x3F]);
                out.push_back(kTable[n & 0x3F]);
            }
            const size_t rem = bytes.size() - i;
            if (rem == 1)
            {
                const uint32_t n = (static_cast<uint32_t>(static_cast<uint8_t>(bytes[i])) << 16);
                out.push_back(kTable[(n >> 18) & 0x3F]);
                out.push_back(kTable[(n >> 12) & 0x3F]);
                out.push_back('=');
                out.push_back('=');
            }
            else if (rem == 2)
            {
                const uint32_t n =
                    (static_cast<uint32_t>(static_cast<uint8_t>(bytes[i])) << 16) | (static_cast<uint32_t>(static_cast<uint8_t>(bytes[i + 1])) << 8);
                out.push_back(kTable[(n >> 18) & 0x3F]);
                out.push_back(kTable[(n >> 12) & 0x3F]);
                out.push_back(kTable[(n >> 6) & 0x3F]);
                out.push_back('=');
            }
            return out;
        };

        const std::string iconsSvg = ResourceBytesString(g_hInstance, IDR_VIEWERWEB_JSONEDITOR_ICONS_SVG);
        const std::string iconsUrl = std::string("data:image/svg+xml;base64,") + base64Encode(iconsSvg);

        std::string css = ResourceBytesString(g_hInstance, IDR_VIEWERWEB_JSONEDITOR_CSS);

        auto replaceAll = [](std::string& text, std::string_view needle, std::string_view replacement) noexcept
        {
            if (needle.empty())
                return;
            size_t pos = 0;
            while ((pos = text.find(needle, pos)) != std::string::npos)
            {
                text.replace(pos, needle.size(), replacement);
                pos += replacement.size();
            }
        };

        replaceAll(css, "./img/jsoneditor-icons.svg", iconsUrl);
        replaceAll(css, "img/jsoneditor-icons.svg", iconsUrl);
        return css;
    }();
    return s;
}

} // namespace

const char* GetViewerWebStaticConfigurationSchema(ViewerWebKind kind) noexcept
{
    return GetViewerWebStaticConfigurationSchemaImpl(kind);
}

LRESULT ViewerWeb::HandleFileComboHostMessage(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, bool& handled) noexcept
{
    const bool popupWasOpen      = _fileComboControl && _fileComboControl->DebugIsPopupOpen();
    const bool preExpandForPopup =
        ! popupWasOpen && _fileComboControl && RedSalamander::ViewerFileComboHost::MessageMayOpenWindowComboPopup(msg, wp);
    if (preExpandForPopup)
    {
        _fileComboHostPreExpandPopup = true;
        if (_hWnd)
        {
            Layout(_hWnd.get());
        }
    }

    const LRESULT dxResult = _fileComboHost.HandleMessage(hwnd, msg, wp, lp, handled);
    if (msg == WM_NCDESTROY)
    {
        handled = true;
        _fileComboHost.ReleaseMouseCapture();
        _fileComboControl            = nullptr;
        _fileComboHostPreExpandPopup = false;
        _hFileComboHost.release();
        return dxResult;
    }

    const bool popupIsOpen = _fileComboControl && _fileComboControl->DebugIsPopupOpen();
    if (popupIsOpen != popupWasOpen || (preExpandForPopup && ! popupIsOpen))
    {
        _fileComboHostPreExpandPopup = false;
        if (_hWnd)
        {
            Layout(_hWnd.get());
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
        }
    }
    return dxResult;
}

void ViewerWeb::FocusMainSurfaceFromFileCombo(HWND hwnd) noexcept
{
    if (_embeddedMode)
    {
        return;
    }

    if (hwnd && IsWindow(hwnd) != FALSE)
    {
        SetFocus(hwnd);
    }
}

void ViewerWeb::OnCreate(HWND hwnd)
{
    const DWORD comboHostStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP | SS_NOTIFY;
    _hFileComboHost.reset(CreateWindowExW(
        0, L"Static", L"", comboHostStyle, 0, 0, 0, 0, hwnd, reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_VIEWERWEB_FILE_COMBO)), g_hInstance, nullptr));
    if (! _hFileComboHost)
    {
        Debug::ErrorWithLastError(L"ViewerWeb: CreateWindowExW failed for DxUi file combo host.");
    }
    else if (! _fileComboHost.Attach(_hFileComboHost.get()))
    {
        Debug::Error(L"ViewerWeb: failed to attach DxUi host for file combo.");
        _hFileComboHost.reset();
    }
    else if (! RedSalamander::ViewerFileComboHost::InstallFileComboHostWindow(
                 _hFileComboHost.get(), this, kFileComboHostStateProp, kFileComboHostOriginalWndProcProp, FileComboHostWndProc))
    {
        Debug::ErrorWithLastError(L"ViewerWeb: failed to install WNDPROC hook for DxUi file combo host.");
        _fileComboHost.Detach();
        _hFileComboHost.reset();
    }
    else
    {
        auto combo        = std::make_unique<ComboBox>();
        _fileComboControl = combo.get();
        _fileComboControl->SetVariant(ComboBoxVariant::Window);
        _fileComboControl->SetOnSelectionChanged([this](size_t selectedIndex)
        {
            if (_syncingFileCombo || selectedIndex >= _otherFiles.size() || ! _hWnd)
            {
                return;
            }

            const HWND hwnd = _hWnd.get();
            _otherIndex     = selectedIndex;
            static_cast<void>(OpenPath(hwnd, _otherFiles[_otherIndex], false));
            if (! _embeddedMode)
            {
                SetFocus(hwnd);
            }
        });
        RedSalamander::ViewerFileComboHost::ConfigureFileComboKeyboard(_fileComboHost,
                                                                       [this]() noexcept
        {
            if (_hWnd)
            {
                FocusMainSurfaceFromFileCombo(_hWnd.get());
            }
        });
        _fileComboHost.SetTheme(_hasTheme ? MakeThemePaletteFromViewerTheme(_theme) : MakeDefaultThemePalette(false));
        _fileComboHost.SetRoot(std::move(combo));
    }

    if (! _menuHandle)
    {
        _menuHandle.reset(GetMenu(hwnd));
    }
    if (_menuHandle)
    {
        _menuBarHost.SetTheme(_hasTheme ? MakeThemePaletteFromViewerTheme(_theme) : MakeDefaultThemePalette(false));
        _menuBarHost.SetRefreshMenuStateCallback([this]
        {
            if (_hWnd)
            {
                UpdateMenuState(_hWnd.get(), false);
            }
        });
        _menuBarHost.SetOnTabBoundary([this](bool) noexcept
        {
            if (_hWnd)
            {
                FocusMainSurfaceFromFileCombo(_hWnd.get());
            }
            return true;
        });
        _menuBarHost.SetOnEscape([this]() noexcept
        {
            if (_hWnd)
            {
                FocusMainSurfaceFromFileCombo(_hWnd.get());
            }
            return true;
        });
        static_cast<void>(_menuBarHost.Attach(g_hInstance, hwnd, _menuHandle.get()));
    }

    ApplyTheme(hwnd);
    RefreshFileCombo(hwnd);
    Layout(hwnd);
    static_cast<void>(EnsureWebView2(hwnd));
}

void ViewerWeb::OnDestroy() noexcept
{
    _saveRequestId.fetch_add(1u, std::memory_order_acq_rel);
    _loadPostFailureTerminal = false;
    _saveInProgress = false;
    _hFindDialog.reset();
    DiscardWebView2();

    if (_tempExtractedPath.has_value())
    {
        DeleteStagedFileOrSchedule(_tempExtractedPath.value());
    }
    _tempExtractedPath.reset();

    _pendingPath.reset();
    _pendingWebContent.reset();
    _pendingDocumentUtf8.reset();
    _internalDocumentUrl.reset();
    _allowedDocumentUrl.clear();
    _documentRoute          = ViewerWebSecurity::DocumentRoute::None;
    _loadedSourceBytes      = 0u;
    _documentScriptsEnabled = false;
    _navigationCompleted    = false;
    _navigationSucceeded    = false;

    NotifyViewerClosed();
}

void ViewerWeb::OnSize(UINT /*width*/, UINT /*height*/) noexcept
{
    if (_hWnd)
    {
        Layout(_hWnd.get());
    }
}

void ViewerWeb::OnCommand(HWND hwnd, UINT commandId, [[maybe_unused]] UINT code, HWND /*control*/) noexcept
{
    switch (commandId)
    {
        case IDM_VIEWERWEB_FILE_SAVE_AS: static_cast<void>(CommandSaveAs(hwnd)); break;
        case IDM_VIEWERWEB_FILE_REFRESH: static_cast<void>(OpenPath(hwnd, _currentPath, false)); break;
        case IDM_VIEWERWEB_FILE_EXIT: static_cast<void>(Close()); break;

        case IDM_VIEWERWEB_OTHER_NEXT:
            if (_otherFiles.size() > 1)
            {
                _otherIndex = (_otherIndex + 1) % _otherFiles.size();
                static_cast<void>(OpenPath(hwnd, _otherFiles[_otherIndex], false));
                RefreshFileCombo(hwnd);
            }
            break;
        case IDM_VIEWERWEB_OTHER_PREVIOUS:
            if (_otherFiles.size() > 1)
            {
                _otherIndex = (_otherIndex + _otherFiles.size() - 1) % _otherFiles.size();
                static_cast<void>(OpenPath(hwnd, _otherFiles[_otherIndex], false));
                RefreshFileCombo(hwnd);
            }
            break;
        case IDM_VIEWERWEB_OTHER_FIRST:
            if (_otherFiles.size() > 1)
            {
                _otherIndex = 0;
                static_cast<void>(OpenPath(hwnd, _otherFiles[_otherIndex], false));
                RefreshFileCombo(hwnd);
            }
            break;
        case IDM_VIEWERWEB_OTHER_LAST:
            if (_otherFiles.size() > 1)
            {
                _otherIndex = _otherFiles.size() - 1;
                static_cast<void>(OpenPath(hwnd, _otherFiles[_otherIndex], false));
                RefreshFileCombo(hwnd);
            }
            break;

        case IDM_VIEWERWEB_SEARCH_FIND: CommandFind(hwnd); break;
        case IDM_VIEWERWEB_SEARCH_FIND_NEXT: CommandFindNext(hwnd); break;
        case IDM_VIEWERWEB_SEARCH_FIND_PREVIOUS: CommandFindPrevious(hwnd); break;

        case IDM_VIEWERWEB_VIEW_ZOOM_IN: CommandZoomIn(); break;
        case IDM_VIEWERWEB_VIEW_ZOOM_OUT: CommandZoomOut(); break;
        case IDM_VIEWERWEB_VIEW_ZOOM_RESET: CommandZoomReset(); break;
        case IDM_VIEWERWEB_VIEW_DEVTOOLS: CommandToggleDevTools(); break;

        case IDM_VIEWERWEB_TOOLS_COPY_URL: CommandCopyUrl(hwnd); break;
        case IDM_VIEWERWEB_TOOLS_OPEN_EXTERNAL: CommandOpenExternal(hwnd); break;
        case IDM_VIEWERWEB_TOOLS_JSON_EXPAND_ALL: CommandJsonExpandAll(); break;
        case IDM_VIEWERWEB_TOOLS_JSON_COLLAPSE_ALL: CommandJsonCollapseAll(); break;
        case IDM_VIEWERWEB_TOOLS_MARKDOWN_TOGGLE_SOURCE: CommandMarkdownToggleSource(); break;

        default: break;
    }
}

void ViewerWeb::OnContextMenu(HWND hwnd, POINT screenPt) noexcept
{
    if (! _menuHandle)
    {
        _menuHandle.reset(Localization::LoadMenuResource(g_hInstance, IDR_VIEWERWEB_MENU));
    }

    HMENU menu = _menuHandle ? _menuHandle.get() : GetMenu(hwnd);
    if (! menu)
    {
        return;
    }

    UpdateMenuState(hwnd, false);
    static constexpr std::array<int, 5> kPreviewContextMenuExcludedCommandIds{{
        IDM_VIEWERWEB_FILE_EXIT,
        IDM_VIEWERWEB_OTHER_NEXT,
        IDM_VIEWERWEB_OTHER_PREVIOUS,
        IDM_VIEWERWEB_OTHER_FIRST,
        IDM_VIEWERWEB_OTHER_LAST,
    }};
    RedSalamander::DxUi::NativeMenuFlyoutOptions previewMenuOptions{};
    previewMenuOptions.includeAcceleratorText = false;
    previewMenuOptions.omitEmptySubmenus      = true;
    previewMenuOptions.trimSeparators         = true;
    previewMenuOptions.excludedCommandIds     = kPreviewContextMenuExcludedCommandIds;

    const auto result = RedSalamander::DxUi::ShowNativeHMenuContextMenu(
        hwnd, screenPt, menu, _hasTheme ? MakeThemePaletteFromViewerTheme(_theme) : MakeDefaultThemePalette(false), previewMenuOptions);
    if (result.has_value() && result.value() > 0)
    {
        OnCommand(hwnd, static_cast<UINT>(result.value()), 0, nullptr);
    }
}

void ViewerWeb::OnKeyDown(HWND hwnd, UINT vk) noexcept
{
    if (! hwnd)
    {
        return;
    }

    const bool ctrl          = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
    const bool shift         = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
    const HKL keyboardLayout = GetKeyboardLayout(0);
    const UINT zoomInVk      = ZoomInVirtualKeyForLayout(keyboardLayout);
    const UINT zoomOutVk     = ZoomOutVirtualKeyForLayout(keyboardLayout);

    if (vk == VK_ESCAPE)
    {
        PostMessageW(hwnd, WM_CLOSE, 0, 0);
        return;
    }

    if (vk == VK_F5)
    {
        static_cast<void>(OpenPath(hwnd, _currentPath, false));
        return;
    }

    if (vk == VK_F12)
    {
        CommandToggleDevTools();
        return;
    }

    if (vk == VK_F3)
    {
        shift ? CommandFindPrevious(hwnd) : CommandFindNext(hwnd);
        return;
    }

    if (ctrl && (vk == 'F' || vk == 'f'))
    {
        CommandFind(hwnd);
        return;
    }

    if (ctrl && (vk == 'S' || vk == 's'))
    {
        static_cast<void>(CommandSaveAs(hwnd));
        return;
    }

    if (ctrl && (vk == 'L' || vk == 'l'))
    {
        CommandCopyUrl(hwnd);
        return;
    }

    if (ctrl && vk == VK_RETURN)
    {
        CommandOpenExternal(hwnd);
        return;
    }

    if (ctrl && (vk == VK_ADD || vk == zoomInVk || vk == '='))
    {
        CommandZoomIn();
        return;
    }

    if (ctrl && (vk == VK_SUBTRACT || vk == zoomOutVk || vk == '-'))
    {
        CommandZoomOut();
        return;
    }

    if (ctrl && (vk == '0'))
    {
        CommandZoomReset();
        return;
    }

    if (ctrl && vk == VK_OEM_3)
    {
        CommandMarkdownToggleSource();
        return;
    }

    if (ctrl && vk == VK_UP)
    {
        SendMessageW(hwnd, WM_COMMAND, IDM_VIEWERWEB_OTHER_PREVIOUS, 0);
        return;
    }

    if (ctrl && vk == VK_DOWN)
    {
        SendMessageW(hwnd, WM_COMMAND, IDM_VIEWERWEB_OTHER_NEXT, 0);
        return;
    }

    if (ctrl && vk == VK_HOME)
    {
        SendMessageW(hwnd, WM_COMMAND, IDM_VIEWERWEB_OTHER_FIRST, 0);
        return;
    }

    if (ctrl && vk == VK_END)
    {
        SendMessageW(hwnd, WM_COMMAND, IDM_VIEWERWEB_OTHER_LAST, 0);
        return;
    }

    if (_kind != ViewerWebKind::Web && vk == VK_SPACE)
    {
        SendMessageW(hwnd, WM_COMMAND, IDM_VIEWERWEB_OTHER_NEXT, 0);
        return;
    }

    if (_kind != ViewerWebKind::Web && vk == VK_BACK)
    {
        SendMessageW(hwnd, WM_COMMAND, IDM_VIEWERWEB_OTHER_PREVIOUS, 0);
        return;
    }
}

void ViewerWeb::OnPaint(HWND hwnd) noexcept
{
    PAINTSTRUCT ps{};
    wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);

    FillRect(hdc.get(), &ps.rcPaint, GetActiveViewerWebClassBackgroundBrush());

    if (_headerBrush)
    {
        FillRect(hdc.get(), &_headerRect, _headerBrush.get());
    }

    if (_hasTheme)
    {
        const UINT dpi          = GetDpiForWindow(hwnd);
        const COLORREF accent   = ResolveAccentColor(_theme, _currentPath.empty() ? _metaId : _currentPath);
        const int lineThickness = std::max(1, static_cast<int>(Common::WindowSizing::DipToPixelRounded(dpi, 1)));
        RECT line               = _headerRect;
        line.top                = std::max(line.top, line.bottom - lineThickness);
        line.bottom             = std::max(line.bottom, line.top);
        wil::unique_hbrush brush(CreateSolidBrush(accent));
        FillRect(hdc.get(), &line, brush.get());
    }

    if (! _statusMessage.empty())
    {
        const UINT dpi    = GetDpiForWindow(hwnd);
        const int padding = static_cast<int>(Common::WindowSizing::DipToPixelRounded(dpi, 8));
        RECT rc           = _headerRect;
        rc.left           = std::min(rc.right, rc.left + padding);
        rc.right          = std::max(rc.left, rc.right - padding);

        const COLORREF textColor = _hasTheme ? ColorRefFromArgb(_theme.textArgb) : GetSysColor(COLOR_WINDOWTEXT);
        if (! DrawStatusMessageWithDirectWrite(hdc.get(), hwnd, rc, _statusMessage, textColor))
        {
            static std::atomic<bool> loggedStatusRenderFailure{false};
            if (! loggedStatusRenderFailure.exchange(true, std::memory_order_relaxed))
            {
                Debug::Error(L"ViewerWeb: DirectWrite/D2D status rendering failed; status text is unavailable without a DX render path.");
            }
        }
    }
}

LRESULT ViewerWeb::OnEraseBkgnd(HWND /*hwnd*/, HDC /*hdc*/) noexcept
{
    return 1;
}

void ViewerWeb::OnDpiChanged(HWND hwnd, UINT newDpi, const RECT* suggested) noexcept
{
    if (suggested)
    {
        SetWindowPos(hwnd,
                     nullptr,
                     suggested->left,
                     suggested->top,
                     std::max(1L, suggested->right - suggested->left),
                     std::max(1L, suggested->bottom - suggested->top),
                     SWP_NOZORDER | SWP_NOACTIVATE);
    }

    static_cast<void>(newDpi);
    ApplyTheme(hwnd);
    Layout(hwnd);
}

LRESULT ViewerWeb::OnNcDestroy(HWND hwnd, WPARAM wp, LPARAM lp) noexcept
{
    OnDestroy();
    static_cast<void>(DrainPostedPayloadsForWindow(hwnd));

    _menuBarHost.Detach();
    _menuHandle.reset();
    UnhookFileComboHostWindow(_hFileComboHost.get());
    _fileComboHost.Detach();
    _hFileComboHost.release();
    _hWnd.release();
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);

    Release();
    return DefWindowProcW(hwnd, WM_NCDESTROY, wp, lp);
}

void ViewerWeb::OnFindMessage(const FINDREPLACEW* findReplace) noexcept
{
    if (! findReplace)
    {
        return;
    }

    if ((findReplace->Flags & FR_DIALOGTERM) != 0)
    {
        _hFindDialog.release();
        return;
    }

    if ((findReplace->Flags & FR_FINDNEXT) == 0)
    {
        return;
    }

    if (findReplace->lpstrFindWhat)
    {
        _findQuery.assign(findReplace->lpstrFindWhat);
    }

    if (_findQuery.empty() || ! _webView)
    {
        return;
    }

    const bool backwards        = (findReplace->Flags & FR_DOWN) == 0;
    const std::wstring queryEsc = EscapeJavaScriptString(_findQuery);
    const std::wstring script   = std::format(L"(function(){{try{{return window.find('{}',false,{},true,false,true,false);}}catch(e){{return false;}}}})();",
                                              queryEsc,
                                              backwards ? L"true" : L"false");
    static_cast<void>(_webView->ExecuteScript(script.c_str(), nullptr));
}

ViewerWeb::ViewerWeb(ViewerWebKind kind) noexcept : _kind(kind)
{
    switch (_kind)
    {
        case ViewerWebKind::Web:
            _metaId          = L"builtin/viewer-web";
            _metaShortId     = L"web";
            _metaName        = LoadStringResource(g_hInstance, IDS_VIEWERWEB_NAME);
            _metaDescription = LoadStringResource(g_hInstance, IDS_VIEWERWEB_DESCRIPTION);
            break;
        case ViewerWebKind::Json:
            _metaId          = L"builtin/viewer-json";
            _metaShortId     = L"json";
            _metaName        = LoadStringResource(g_hInstance, IDS_VIEWERJSON_NAME);
            _metaDescription = LoadStringResource(g_hInstance, IDS_VIEWERJSON_DESCRIPTION);
            break;
        case ViewerWebKind::Markdown:
            _metaId          = L"builtin/viewer-markdown";
            _metaShortId     = L"md";
            _metaName        = LoadStringResource(g_hInstance, IDS_VIEWERMARKDOWN_NAME);
            _metaDescription = LoadStringResource(g_hInstance, IDS_VIEWERMARKDOWN_DESCRIPTION);
            break;
        default:
            _metaId          = L"builtin/viewer-web";
            _metaShortId     = L"web";
            _metaName        = LoadStringResource(g_hInstance, IDS_VIEWERWEB_NAME);
            _metaDescription = LoadStringResource(g_hInstance, IDS_VIEWERWEB_DESCRIPTION);
            break;
    }
}

ViewerWeb::~ViewerWeb()
{
    // WebView2 cleanup is handled in OnDestroy() to allow async shutdown to complete
    // before the object is destroyed. Do not call DiscardWebView2() here.

    if (_hFileComboHost)
    {
        _fileComboControl            = nullptr;
        _fileComboHostPreExpandPopup = false;
        UnhookFileComboHostWindow(_hFileComboHost.get());
        _fileComboHost.Detach();
        _hFileComboHost.reset();
    }

    if (_tempExtractedPath.has_value())
    {
        DeleteStagedFileOrSchedule(_tempExtractedPath.value());
        _tempExtractedPath.reset();
    }
}

void ViewerWeb::SetHost(IHost* host) noexcept
{
    _host       = host;
    _hostAlerts = nullptr;

    if (! _host)
    {
        return;
    }

    wil::com_ptr<IHostAlerts> alerts;
    const HRESULT hr = _host->QueryInterface(__uuidof(IHostAlerts), alerts.put_void());
    if (SUCCEEDED(hr) && alerts)
    {
        _hostAlerts = std::move(alerts);
    }
}

HRESULT STDMETHODCALLTYPE ViewerWeb::QueryInterface(REFIID riid, void** ppvObject) noexcept
{
    if (ppvObject == nullptr)
    {
        return E_POINTER;
    }

    *ppvObject = nullptr;

    if (riid == __uuidof(IUnknown) || riid == __uuidof(IViewer))
    {
        *ppvObject = static_cast<IViewer*>(this);
        AddRef();
        return S_OK;
    }

    if (riid == __uuidof(IInformations))
    {
        *ppvObject = static_cast<IInformations*>(this);
        AddRef();
        return S_OK;
    }

    return E_NOINTERFACE;
}

ULONG STDMETHODCALLTYPE ViewerWeb::AddRef() noexcept
{
    return _refCount.fetch_add(1, std::memory_order_relaxed) + 1;
}

ULONG STDMETHODCALLTYPE ViewerWeb::Release() noexcept
{
    const ULONG refs = _refCount.fetch_sub(1, std::memory_order_acq_rel) - 1;
    if (refs == 0)
    {
        delete this;
    }
    return refs;
}

HRESULT STDMETHODCALLTYPE ViewerWeb::GetMetaData(const PluginMetaData** metaData) noexcept
{
    if (metaData == nullptr)
    {
        return E_POINTER;
    }

    _metaData.id          = _metaId.c_str();
    _metaData.shortId     = _metaShortId.c_str();
    _metaData.name        = _metaName.empty() ? nullptr : _metaName.c_str();
    _metaData.description = _metaDescription.empty() ? nullptr : _metaDescription.c_str();
    _metaData.author      = nullptr;
    _metaData.version     = VERSINFO_PLUGIN_VERSION;

    *metaData = &_metaData;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerWeb::GetConfigurationSchema(const char** schemaJsonUtf8) noexcept
{
    if (schemaJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *schemaJsonUtf8 = GetViewerWebStaticConfigurationSchema(_kind);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerWeb::SetConfiguration(const char* configurationJsonUtf8) noexcept
{
    uint32_t maxDocumentMiB      = ViewerWebSecurity::kDefaultMaxDocumentMiB;
    bool allowExternalNavigation = ViewerWebSecurity::kDefaultAllowExternalNavigation;
    bool devToolsEnabled         = false;
    JsonViewMode jsonViewMode    = JsonViewMode::Pretty;

    if (configurationJsonUtf8 != nullptr && configurationJsonUtf8[0] != '\0')
    {
        const std::string_view utf8(configurationJsonUtf8);
        if (! utf8.empty())
        {
            std::string mutableJson(utf8);
            yyjson_read_err err{};
            unique_yyjson_doc doc(yyjson_read_opts(mutableJson.data(), mutableJson.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, &err));
            if (doc)
            {
                yyjson_val* root = yyjson_doc_get_root(doc.get());
                if (yyjson_is_obj(root))
                {
                    yyjson_val* maxDoc = yyjson_obj_get(root, "maxDocumentMiB");
                    if (yyjson_is_int(maxDoc))
                    {
                        const int64_t value = yyjson_get_int(maxDoc);
                        if (value >= 1)
                        {
                            maxDocumentMiB = static_cast<uint32_t>(std::min<int64_t>(value, ViewerWebSecurity::kMaximumDocumentMiB));
                        }
                    }
                    else if (yyjson_is_uint(maxDoc))
                    {
                        maxDocumentMiB =
                            static_cast<uint32_t>(std::min<uint64_t>(yyjson_get_uint(maxDoc), ViewerWebSecurity::kMaximumDocumentMiB));
                        maxDocumentMiB = std::max(maxDocumentMiB, 1u);
                    }

                    const auto readBool = [&](const char* key, bool defaultValue) -> bool
                    {
                        yyjson_val* v = yyjson_obj_get(root, key);
                        if (yyjson_is_bool(v))
                        {
                            return yyjson_get_bool(v);
                        }
                        if (yyjson_is_str(v))
                        {
                            const char* s = yyjson_get_str(v);
                            if (s && strcmp(s, "1") == 0)
                            {
                                return true;
                            }
                            if (s && strcmp(s, "0") == 0)
                            {
                                return false;
                            }
                        }
                        return defaultValue;
                    };

                    allowExternalNavigation = readBool("allowExternalNavigation", allowExternalNavigation);
                    devToolsEnabled         = readBool("devToolsEnabled", devToolsEnabled);

                    yyjson_val* modeVal = yyjson_obj_get(root, "viewMode");
                    if (yyjson_is_str(modeVal))
                    {
                        const char* s = yyjson_get_str(modeVal);
                        if (s && (strcmp(s, "tree") == 0 || strcmp(s, "1") == 0))
                        {
                            jsonViewMode = JsonViewMode::Tree;
                        }
                        else if (s && (strcmp(s, "jsonl") == 0 || strcmp(s, "json-lines") == 0 || strcmp(s, "2") == 0))
                        {
                            jsonViewMode = JsonViewMode::JsonLines;
                        }
                        else if (s && (strcmp(s, "pretty") == 0 || strcmp(s, "0") == 0))
                        {
                            jsonViewMode = JsonViewMode::Pretty;
                        }
                    }
                    else if (yyjson_is_int(modeVal))
                    {
                        const int64_t value = yyjson_get_int(modeVal);
                        jsonViewMode        = value <= 0 ? JsonViewMode::Pretty : value == 1 ? JsonViewMode::Tree : JsonViewMode::JsonLines;
                    }
                    else if (yyjson_is_uint(modeVal))
                    {
                        const uint64_t value = yyjson_get_uint(modeVal);
                        jsonViewMode         = value == 0u ? JsonViewMode::Pretty : value == 1u ? JsonViewMode::Tree : JsonViewMode::JsonLines;
                    }
                }
            }
        }
    }

    _config.maxDocumentMiB          = maxDocumentMiB;
    _config.allowExternalNavigation = allowExternalNavigation;
    _config.devToolsEnabled         = devToolsEnabled;
    _config.jsonViewMode            = jsonViewMode;

    switch (_kind)
    {
        case ViewerWebKind::Json:
            _configurationJson = std::format(
                R"json({{
    "maxDocumentMiB": {},
    "viewMode": "{}",
    "devToolsEnabled": {}
}})json",
                _config.maxDocumentMiB,
                _config.jsonViewMode == JsonViewMode::Tree        ? "tree"
                : _config.jsonViewMode == JsonViewMode::JsonLines ? "jsonl"
                                                                  : "pretty",
                _config.devToolsEnabled ? "true" : "false");
            break;
        case ViewerWebKind::Markdown:
            _configurationJson = std::format(
                R"json({{
    "maxDocumentMiB": {},
    "allowExternalNavigation": {},
    "devToolsEnabled": {}
}})json",
                _config.maxDocumentMiB,
                _config.allowExternalNavigation ? "true" : "false",
                _config.devToolsEnabled ? "true" : "false");
            break;
        case ViewerWebKind::Web:
        default:
            _configurationJson = std::format(
                R"json({{
    "maxDocumentMiB": {},
    "allowExternalNavigation": {},
    "devToolsEnabled": {}
}})json",
                _config.maxDocumentMiB,
                _config.allowExternalNavigation ? "true" : "false",
                _config.devToolsEnabled ? "true" : "false");
            break;
    }

    if (_webView)
    {
        const HRESULT settingsHr = ConfigureWebViewSettings();
        if (FAILED(settingsHr))
        {
            _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_WEBVIEW2_INIT_FAILED);
            DiscardWebView2();
            return settingsHr;
        }
    }

    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerWeb::GetConfiguration(const char** configurationJsonUtf8) noexcept
{
    if (configurationJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    if (_configurationJson.empty())
    {
        *configurationJsonUtf8 = nullptr;
        return S_OK;
    }

    *configurationJsonUtf8 = _configurationJson.c_str();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerWeb::SomethingToSave(BOOL* pSomethingToSave) noexcept
{
    if (pSomethingToSave == nullptr)
    {
        return E_POINTER;
    }

    bool isDefault = false;
    switch (_kind)
    {
        case ViewerWebKind::Json:
            isDefault = _config.maxDocumentMiB == ViewerWebSecurity::kDefaultMaxDocumentMiB && _config.jsonViewMode == JsonViewMode::Pretty &&
                        ! _config.devToolsEnabled;
            break;
        case ViewerWebKind::Markdown:
            isDefault = _config.maxDocumentMiB == ViewerWebSecurity::kDefaultMaxDocumentMiB && ! _config.allowExternalNavigation &&
                        ! _config.devToolsEnabled;
            break;
        case ViewerWebKind::Web:
        default:
            isDefault = _config.maxDocumentMiB == ViewerWebSecurity::kDefaultMaxDocumentMiB && ! _config.allowExternalNavigation &&
                        ! _config.devToolsEnabled;
            break;
    }
    *pSomethingToSave = isDefault ? FALSE : TRUE;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerWeb::Open(const ViewerOpenContext* context) noexcept
{
    if (! context || ! context->focusedPath || context->focusedPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    if (! context->fileSystem)
    {
        return E_INVALIDARG;
    }

    _fileSystem = context->fileSystem;
    _fileSystemName.assign(context->fileSystemName ? context->fileSystemName : L"");

    _otherFiles.clear();
    if (context->otherFiles && context->otherFileCount > 0)
    {
        _otherFiles.reserve(context->otherFileCount);
        for (unsigned long i = 0; i < context->otherFileCount; ++i)
        {
            const wchar_t* p = context->otherFiles[i];
            if (p && p[0] != L'\0')
            {
                _otherFiles.emplace_back(p);
            }
        }
    }
    if (_otherFiles.empty())
    {
        _otherFiles.emplace_back(context->focusedPath);
    }

    _otherIndex = 0;
    if (context->focusedOtherFileIndex < _otherFiles.size())
    {
        _otherIndex = static_cast<size_t>(context->focusedOtherFileIndex);
    }

    const std::wstring focusedPath(context->focusedPath);

    const bool embeddedMode   = IsEmbeddedOpen(*context);
    const HWND embeddedParent = embeddedMode ? context->ownerWindow : nullptr;
    if (embeddedMode && (embeddedParent == nullptr || IsWindow(embeddedParent) == FALSE))
    {
        Debug::Error(L"ViewerWeb: embedded Open requires a valid ownerWindow parent.");
        return E_INVALIDARG;
    }
    if (ShouldRecreateViewerWindow(embeddedMode, embeddedParent))
    {
        static_cast<void>(Close());
    }

    if (! _hWnd)
    {
        if (! RegisterWndClass(g_hInstance))
        {
            return E_FAIL;
        }

        RECT ownerRect{};
        int x = CW_USEDEFAULT;
        int y = CW_USEDEFAULT;
        int w = 1000;
        int h = 700;
        if (embeddedMode)
        {
            RECT client{};
            if (GetClientRect(embeddedParent, &client) == 0)
            {
                const DWORD lastError = Debug::ErrorWithLastError(L"ViewerWeb: GetClientRect failed for embedded preview parent.");
                return HRESULT_FROM_WIN32(lastError);
            }
            x = 0;
            y = 0;
            w = std::max(1L, client.right - client.left);
            h = std::max(1L, client.bottom - client.top);
        }
        else if (context->ownerWindow && GetWindowRect(context->ownerWindow, &ownerRect) != 0)
        {
            x = static_cast<int>(ownerRect.left);
            y = static_cast<int>(ownerRect.top);
            w = static_cast<int>(std::max<LONG>(1, ownerRect.right - ownerRect.left));
            h = static_cast<int>(std::max<LONG>(1, ownerRect.bottom - ownerRect.top));
        }

        wil::unique_any<HMENU, decltype(&::DestroyMenu), ::DestroyMenu> menu(embeddedMode ? nullptr
                                                                                          : Localization::LoadMenuResource(g_hInstance, IDR_VIEWERWEB_MENU));
        const DWORD style               = embeddedMode ? (WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS) : (WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN);
        const std::wstring initialTitle = embeddedMode ? std::wstring{} : (_metaName.empty() ? LoadStringResource(g_hInstance, IDS_VIEWERWEB_NAME) : _metaName);
        HWND window =
            CreateWindowExW(0, kClassName, initialTitle.c_str(), style, x, y, w, h, embeddedMode ? embeddedParent : nullptr, menu.get(), g_hInstance, this);
        if (! window)
        {
            const DWORD lastError = Debug::ErrorWithLastError(L"ViewerWeb: CreateWindowExW failed.");
            return HRESULT_FROM_WIN32(lastError);
        }

        if (! embeddedMode)
        {
            menu.release();
        }
        _hWnd.reset(window);
        _embeddedMode = embeddedMode;

        ApplyTheme(_hWnd.get());
        ApplyPendingViewerWebClassBackgroundBrush(_hWnd.get());

        AddRef(); // Self-reference for window lifetime (released in WM_NCDESTROY)
        ShowWindow(_hWnd.get(), embeddedMode ? SW_SHOWNA : SW_SHOWNORMAL);
        if (! embeddedMode)
        {
            static_cast<void>(SetForegroundWindow(_hWnd.get()));
        }
    }
    else
    {
        ApplyPendingViewerWebClassBackgroundBrush(_hWnd.get());
        if (_embeddedMode)
        {
            RECT client{};
            if (embeddedParent != nullptr && GetClientRect(embeddedParent, &client) != 0)
            {
                SetWindowPos(_hWnd.get(),
                             nullptr,
                             0,
                             0,
                             std::max(1L, client.right - client.left),
                             std::max(1L, client.bottom - client.top),
                             SWP_NOZORDER | SWP_NOACTIVATE);
            }
            ShowWindow(_hWnd.get(), SW_SHOWNA);
        }
        else
        {
            ShowWindow(_hWnd.get(), SW_SHOWNORMAL);
            static_cast<void>(SetForegroundWindow(_hWnd.get()));
        }
    }

    if (! _hWnd)
    {
        Debug::Error(L"ViewerWeb: Open failed because viewer window is missing after creation.");
        return E_FAIL;
    }

    RefreshFileCombo(_hWnd.get());
    return OpenPath(_hWnd.get(), focusedPath, false);
}

HRESULT STDMETHODCALLTYPE ViewerWeb::Close() noexcept
{
    AddRef();
    const auto releaseSelf = wil::scope_exit([&]() noexcept { Release(); });
    _hWnd.reset();
    _embeddedMode = false;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerWeb::SetTheme(const ViewerTheme* theme) noexcept
{
    if (! theme || theme->version < 2u || theme->version > 4u)
    {
        return E_INVALIDARG;
    }

    _theme    = *theme;
    _hasTheme = true;

    RequestViewerWebClassBackgroundColor(ColorRefFromArgb(_theme.backgroundArgb));
    ApplyPendingViewerWebClassBackgroundBrush(_hWnd.get());

    if (_hWnd)
    {
        ApplyTheme(_hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, TRUE);
    }

    return S_OK;
}

ATOM ViewerWeb::RegisterWndClass(HINSTANCE instance) noexcept
{
    if (g_viewerWebClassBackgroundBrush.classRegistered)
    {
        return 1;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = WndProcThunk;
    wc.cbClsExtra    = 0;
    wc.cbWndExtra    = 0;
    wc.hInstance     = instance;
    wc.hIcon         = LoadIconW(nullptr, IDI_APPLICATION);
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = GetActiveViewerWebClassBackgroundBrush();
    wc.lpszMenuName  = nullptr;
    wc.lpszClassName = kClassName;
    wc.hIconSm       = wc.hIcon;

    const ATOM atom = RegisterClassExW(&wc);
    if (atom != 0)
    {
        g_viewerWebClassBackgroundBrush.classRegistered = true;
    }

    return atom;
}

LRESULT CALLBACK ViewerWeb::WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    ViewerWeb* self = reinterpret_cast<ViewerWeb*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (msg == WM_NCCREATE)
    {
        const auto* cs = reinterpret_cast<const CREATESTRUCTW*>(lp);
        self           = reinterpret_cast<ViewerWeb*>(cs ? cs->lpCreateParams : nullptr);
        if (self)
        {
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
            InitPostedPayloadWindow(hwnd);
        }
    }

    if (self)
    {
        return self->WndProc(hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT ViewerWeb::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    static const UINT findMsg = RegisterWindowMessageW(FINDMSGSTRINGW);
    if (findMsg != 0 && msg == findMsg)
    {
        OnFindMessage(reinterpret_cast<const FINDREPLACEW*>(lp));
        return 0;
    }

    const UINT debugSnapshotMessage = ViewerWebSecurity::GetDebugSnapshotMessage();
    if (debugSnapshotMessage != 0u && msg == debugSnapshotMessage)
    {
        auto* snapshot = reinterpret_cast<ViewerWebSecurity::DebugSnapshot*>(lp);
        if (! snapshot || snapshot->sizeBytes < sizeof(ViewerWebSecurity::DebugSnapshot))
        {
            return FALSE;
        }

        *snapshot                     = {};
        snapshot->route               = _documentRoute;
        snapshot->privateOrigin       = _internalDocumentUrl.has_value() ? TRUE : FALSE;
        snapshot->stagedFileTracked   = _tempExtractedPath.has_value() ? TRUE : FALSE;
        snapshot->navigationCompleted = _navigationCompleted ? TRUE : FALSE;
        snapshot->navigationSucceeded = _navigationSucceeded ? TRUE : FALSE;
        snapshot->generatedOutputRejected = _generatedOutputRejected ? TRUE : FALSE;
        snapshot->loadedSourceBytes   = _loadedSourceBytes;
        snapshot->pendingCleanupCount = g_stagedCleanupTracker.PendingCount();
        snapshot->generatedOutputBytes = _generatedOutputBytes;
        snapshot->generatedOutputLimit = _generatedOutputLimit;
        snapshot->asyncLoadPostFailures = _asyncLoadPostFailureCount.load(std::memory_order_acquire);
        snapshot->asyncSavePostFailures = _asyncSavePostFailureCount.load(std::memory_order_acquire);
        snapshot->loadPostFailureTerminal = _loadPostFailureTerminal ? TRUE : FALSE;
        snapshot->saveInProgress        = _saveInProgress ? TRUE : FALSE;
        const size_t copyLength       = std::min(_allowedDocumentUrl.size(), snapshot->allowedDocumentUrl.size() - 1u);
        std::copy_n(_allowedDocumentUrl.begin(), copyLength, snapshot->allowedDocumentUrl.begin());
        snapshot->allowedDocumentUrl[copyLength] = L'\0';
        snapshot->scriptsEnabled = _documentScriptsEnabled ? TRUE : FALSE;
        if (_webView)
        {
            wil::com_ptr<ICoreWebView2Settings> settings;
            if (SUCCEEDED(_webView->get_Settings(settings.put())) && settings)
            {
                static_cast<void>(settings->get_IsScriptEnabled(&snapshot->scriptsEnabled));
            }

            wil::unique_cotaskmem_string source;
            if (SUCCEEDED(_webView->get_Source(source.put())) && source)
            {
                const std::wstring_view sourceView(source.get());
                const size_t sourceLength = std::min(sourceView.size(), snapshot->webViewSourceUrl.size() - 1u);
                std::copy_n(sourceView.begin(), sourceLength, snapshot->webViewSourceUrl.begin());
                snapshot->webViewSourceUrl[sourceLength] = L'\0';
            }
        }
        return TRUE;
    }

#ifdef ENABLE_TESTS
    const UINT debugControlMessage = ViewerWebSecurity::GetDebugControlMessage();
    if (debugControlMessage != 0u && msg == debugControlMessage)
    {
        switch (static_cast<ViewerWebSecurity::DebugControlAction>(wp))
        {
            case ViewerWebSecurity::DebugControlAction::FailNextAsyncLoadCompletionPost:
                g_failNextAsyncLoadCompletionPost.store(true, std::memory_order_release);
                return TRUE;
            case ViewerWebSecurity::DebugControlAction::FailNextAsyncSaveCompletionPost:
                g_failNextAsyncSaveCompletionPost.store(true, std::memory_order_release);
                return TRUE;
            case ViewerWebSecurity::DebugControlAction::SaveAsToPath:
            {
                auto* request = reinterpret_cast<ViewerWebSecurity::DebugSaveAsRequest*>(lp);
                if (! request || request->sizeBytes < sizeof(ViewerWebSecurity::DebugSaveAsRequest) || ! request->destinationPath ||
                    request->destinationPath[0] == L'\0')
                {
                    return FALSE;
                }
                request->submissionHr = StartAsyncSave(hwnd, std::filesystem::path(request->destinationPath), request->faultMask);
                return TRUE;
            }
            default: return FALSE;
        }
    }
#endif

    switch (msg)
    {
#ifdef ENABLE_TESTS
        case WndMsg::kViewerDebugGetNativeMenuModelSnapshot:
        {
            auto* snapshot = reinterpret_cast<WndMsg::ViewerNativeMenuModelDebugSnapshot*>(lp);
            if (! snapshot)
            {
                return FALSE;
            }

            *snapshot                    = {};
            snapshot->hasHiddenMenuModel = _menuHandle != nullptr;
            snapshot->ownerDrawItemCount = CountOwnerDrawMenuItems(_menuHandle.get());
            return TRUE;
        }
#endif
        case WM_CREATE: OnCreate(hwnd); return 0;
        case WM_SIZE: OnSize(static_cast<UINT>(LOWORD(lp)), static_cast<UINT>(HIWORD(lp))); return 0;
        case WM_GETMINMAXINFO:
            if (auto* info = reinterpret_cast<MINMAXINFO*>(lp))
            {
                Common::WindowSizing::ApplyMinimumClientTrackSizeForDips(hwnd, *info, 520, 320);
            }
            return 0;
        case WM_COMMAND: OnCommand(hwnd, static_cast<UINT>(LOWORD(wp)), static_cast<UINT>(HIWORD(wp)), reinterpret_cast<HWND>(lp)); return 0;
        case WM_CONTEXTMENU: OnContextMenu(hwnd, RedSalamander::DxUi::ResolveNativeContextMenuScreenPoint(hwnd, lp)); return 0;
        case WM_KEYDOWN: OnKeyDown(hwnd, static_cast<UINT>(wp)); return 0;
        case WM_SYSKEYDOWN:
            if ((wp == VK_F10 || wp == VK_MENU) && _menuBarHost.FocusFirstItem())
            {
                return 0;
            }
            if ((GetKeyState(VK_CONTROL) & 0x8000) != 0)
            {
                OnKeyDown(hwnd, static_cast<UINT>(wp));
                return 0;
            }
            return DefWindowProcW(hwnd, msg, wp, lp);
        case WM_SYSCHAR:
            if (wp >= 0x20u && _menuBarHost.ActivateMnemonic(static_cast<wchar_t>(wp)))
            {
                return 0;
            }
            return DefWindowProcW(hwnd, msg, wp, lp);
        case WM_PAINT:
            if (_asyncLoadPostFailureRequestId.load(std::memory_order_acquire) != 0u)
            {
                OnAsyncPostFailure(kAsyncPostFailureLoad);
            }
            if (_asyncSavePostFailureRequestId.load(std::memory_order_acquire) != 0u)
            {
                OnAsyncPostFailure(kAsyncPostFailureSave);
            }
            OnPaint(hwnd);
            return 0;
        case WM_ERASEBKGND: return OnEraseBkgnd(hwnd, reinterpret_cast<HDC>(wp));
        case WM_DPICHANGED: OnDpiChanged(hwnd, HIWORD(wp), reinterpret_cast<const RECT*>(lp)); return 0;
        case kAsyncLoadCompleteMessage:
        {
            auto result = TakeMessagePayload<AsyncLoadResult>(lp);
            OnAsyncLoadComplete(std::move(result));
            return 0;
        }
        case kAsyncSaveCompleteMessage:
        {
            auto result = TakeMessagePayload<AsyncSaveWorkItem>(lp);
            OnAsyncSaveComplete(std::move(result));
            return 0;
        }
        case kAsyncPostFailureMessage: OnAsyncPostFailure(static_cast<UINT>(wp)); return 0;
        case WM_CLOSE: static_cast<void>(Close()); return 0;
        case WM_NCACTIVATE:
        {
            const bool windowActive = wp != FALSE;
            ApplyTitleBarTheme(windowActive);
            return DefWindowProcW(hwnd, msg, wp, lp);
        }
        case WM_NCDESTROY: return OnNcDestroy(hwnd, wp, lp);
        default: return DefWindowProcW(hwnd, msg, wp, lp);
    }
}

void ViewerWeb::OnAsyncLoadComplete(std::unique_ptr<AsyncLoadResult> result) noexcept
{
    if (! result)
    {
        return;
    }
    if (result->viewer != this)
    {
        if (result->extractedWin32Path.has_value())
        {
            DeleteStagedFileOrSchedule(result->extractedWin32Path.value());
        }
        return;
    }

    if (result->hwnd != _hWnd.get() || result->requestId != _openRequestId)
    {
        if (result->extractedWin32Path.has_value())
        {
            DeleteStagedFileOrSchedule(result->extractedWin32Path.value());
        }
        return;
    }

    _statusMessage               = result->statusMessage;
    _generatedOutputRejected     = result->generatedOutputRejected;
    _generatedOutputBytes        = result->generatedOutputBytes;
    _generatedOutputLimit        = result->generatedOutputLimit;
    _jsonExpandCollapseAvailable = SUCCEEDED(result->hr) && result->jsonExpandCollapseAvailable;
    UpdateMenuState(_hWnd.get());

    if (FAILED(result->hr))
    {
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
        }

        if (result->offerTextViewerFallback && OfferTextViewerFallbackPrompt())
        {
            return;
        }

        if (! _statusMessage.empty())
        {
            ShowHostAlert(_hWnd.get(), HOST_ALERT_ERROR, _statusMessage);
        }
        return;
    }

    if (_hWnd && ! result->title.empty())
    {
        SetWindowTextW(_hWnd.get(), result->title.c_str());
    }

    _pendingPath.reset();
    _pendingWebContent.reset();
    _generatedOutputRejected = false;
    _generatedOutputBytes    = 0u;
    _generatedOutputLimit    = 0u;
    _pendingDocumentUtf8.reset();
    _internalDocumentUrl.reset();
    _allowedDocumentUrl.clear();
    if (_tempExtractedPath.has_value())
    {
        DeleteStagedFileOrSchedule(_tempExtractedPath.value());
    }
    _tempExtractedPath.reset();
    _documentRoute          = result->documentRoute;
    _loadedSourceBytes      = result->loadedSourceBytes;
    _documentScriptsEnabled = result->documentRoute == ViewerWebSecurity::DocumentRoute::GeneratedPrivateOrigin;

    if (result->documentRoute == ViewerWebSecurity::DocumentRoute::RawHtmlPrivateOrigin ||
        result->documentRoute == ViewerWebSecurity::DocumentRoute::GeneratedPrivateOrigin)
    {
        if (result->documentRoute == ViewerWebSecurity::DocumentRoute::GeneratedPrivateOrigin && result->utf8.empty())
        {
            _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_BUILD_HTML_DOCUMENT);
            ShowHostAlert(_hWnd.get(), HOST_ALERT_ERROR, _statusMessage);
            if (_hWnd)
            {
                InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
            }
            return;
        }

        _pendingDocumentUtf8 = std::move(result->utf8);
        _internalDocumentUrl = BuildInternalDocumentUrl(_kind, result->requestId);
        _allowedDocumentUrl  = _internalDocumentUrl.value();
        _pendingPath         = _internalDocumentUrl;
    }
    else if (result->documentRoute == ViewerWebSecurity::DocumentRoute::StagedPdf)
    {
        if (! result->extractedWin32Path.has_value() || result->extractedWin32Path->empty())
        {
            _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_WRITE_TEMP_FILE_FAILED);
            ShowHostAlert(_hWnd.get(), HOST_ALERT_ERROR, _statusMessage);
            return;
        }

        const std::filesystem::path navigationPath = result->extractedWin32Path.value();
        _tempExtractedPath                         = navigationPath;
        _allowedDocumentUrl = UrlFromFilePath(navigationPath.wstring());
        if (_allowedDocumentUrl.empty())
        {
            _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_BUILD_FILE_URL);
            ShowHostAlert(_hWnd.get(), HOST_ALERT_ERROR, _statusMessage);
            return;
        }
        _pendingPath = _allowedDocumentUrl;
    }
    else
    {
        _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_BUILD_HTML_DOCUMENT);
        ShowHostAlert(_hWnd.get(), HOST_ALERT_ERROR, _statusMessage);
        return;
    }

    _loadPostFailureTerminal = false;

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
    }

    const HRESULT ensureHr = EnsureWebView2(_hWnd.get());
    if (SUCCEEDED(ensureHr) && _webView)
    {
        const HRESULT settingsHr = ConfigureWebViewSettings();
        if (FAILED(settingsHr))
        {
            _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_WEBVIEW2_INIT_FAILED);
            Debug::Error(std::format(L"ViewerWeb: security settings could not be applied (hr=0x{:08X}).",
                                     static_cast<unsigned long>(settingsHr)));
            DiscardWebView2();
            ShowHostAlert(_hWnd.get(), HOST_ALERT_ERROR, _statusMessage);
            return;
        }
        const HRESULT navHr = NavigatePendingContent(_hWnd.get());
        if (FAILED(navHr))
        {
            const std::wstring message = _statusMessage.empty()
                                             ? FormatStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_NAVIGATE_FAILED_FMT, static_cast<unsigned long>(navHr))
                                             : _statusMessage;
            ShowHostAlert(_hWnd.get(), HOST_ALERT_ERROR, message);
        }
    }
}

void ViewerWeb::OnAsyncSaveComplete(std::unique_ptr<AsyncSaveWorkItem> result) noexcept
{
    if (! result || result->viewer != this || result->hwnd != _hWnd.get() ||
        result->requestId != _saveRequestId.load(std::memory_order_acquire))
    {
        return;
    }

    _saveInProgress = false;
    if (FAILED(result->hr))
    {
        _statusMessage = result->statusMessage.empty() ? LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_SAVE_AS_FAILED)
                                                        : std::move(result->statusMessage);
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
        }
        ShowHostAlert(_hWnd.get(), HOST_ALERT_ERROR, _statusMessage);
    }
    else
    {
        _statusMessage.clear();
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
        }
    }
}

void ViewerWeb::OnAsyncPostFailure(UINT operationKind) noexcept
{
    if (operationKind == kAsyncPostFailureLoad)
    {
        const uint64_t requestId = _asyncLoadPostFailureRequestId.exchange(0u, std::memory_order_acq_rel);
        if (requestId == 0u || requestId != _openRequestId)
        {
            return;
        }
        _jsonExpandCollapseAvailable = false;
        _loadPostFailureTerminal = true;
        _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_READ_FILE_FAILED);
    }
    else if (operationKind == kAsyncPostFailureSave)
    {
        const uint64_t requestId = _asyncSavePostFailureRequestId.exchange(0u, std::memory_order_acq_rel);
        if (requestId == 0u || requestId != _saveRequestId.load(std::memory_order_acquire))
        {
            return;
        }
        _saveInProgress = false;
        const HRESULT saveHr = _asyncSavePostFailureHr.load(std::memory_order_relaxed);
        if (SUCCEEDED(saveHr))
        {
            return;
        }
        _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_SAVE_AS_FAILED);
    }
    else
    {
        return;
    }

    UpdateMenuState(_hWnd.get());
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
        ShowHostAlert(_hWnd.get(), HOST_ALERT_ERROR, _statusMessage);
    }
}

void ViewerWeb::Layout(HWND hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    _menuBarHost.UpdateLayout();
    ComputeLayoutRects(hwnd);

    const UINT dpi       = GetDpiForWindow(hwnd);
    const int minPadding = static_cast<int>(Common::WindowSizing::DipToPixelRounded(dpi, RedSalamander::ViewerFileComboHost::kStandaloneComboChromePaddingDip));
    const int accentH    = std::max(1, static_cast<int>(Common::WindowSizing::DipToPixelRounded(dpi, RedSalamander::ViewerFileComboHost::kStandaloneComboAccentHeightDip)));
    const int accentGap  = std::max(1, static_cast<int>(Common::WindowSizing::DipToPixelRounded(dpi, RedSalamander::ViewerFileComboHost::kStandaloneComboAccentGapDip)));
    const bool showCombo = _hFileComboHost && ! _embeddedMode && _otherFiles.size() > 1;

    RECT headerContentRect{};
    headerContentRect        = _headerRect;
    headerContentRect.top    = std::min(headerContentRect.bottom, headerContentRect.top + minPadding);
    headerContentRect.bottom = std::max(headerContentRect.top, headerContentRect.bottom - accentH - accentGap - minPadding);
    const int headerContentH = std::max(0L, headerContentRect.bottom - headerContentRect.top);

    if (_hFileComboHost)
    {
        ShowWindow(_hFileComboHost.get(), showCombo ? SW_SHOW : SW_HIDE);
        EnableWindow(_hFileComboHost.get(), showCombo ? TRUE : FALSE);
        if (_fileComboControl)
        {
            _fileComboControl->SetEnabled(showCombo);
            _fileComboControl->SetVisible(showCombo);
        }
        if (! showCombo)
        {
            _fileComboHostPreExpandPopup = false;
        }

        if (showCombo)
        {
            const int statusReserveW = _statusMessage.empty() ? 0 : static_cast<int>(Common::WindowSizing::DipToPixelRounded(dpi, 160));
            const int margin         = static_cast<int>(Common::WindowSizing::DipToPixelRounded(dpi, 10));

            const int comboX = headerContentRect.left + margin;
            int rightLimit   = std::max<LONG>(headerContentRect.left, headerContentRect.right) - margin;
            if (statusReserveW)
            {
                rightLimit = std::max(comboX, rightLimit - statusReserveW - margin);
            }
            const int comboW = std::max(0, rightLimit - comboX);

            int comboH = std::max(1L, Common::WindowSizing::DipToPixelRounded(dpi, 32));
            comboH     = std::clamp(comboH, 1, std::max(1, headerContentH));

            int comboY          = headerContentRect.top + std::max(0, (headerContentH - comboH) / 2);
            const int maxBottom = std::max(static_cast<int>(headerContentRect.top), static_cast<int>(headerContentRect.bottom));
            if (comboY + comboH > maxBottom)
            {
                comboY = std::max(static_cast<int>(headerContentRect.top), maxBottom - comboH);
            }

            const bool expandPopupHost = _fileComboHostPreExpandPopup || (_fileComboControl && _fileComboControl->DebugIsPopupOpen());
            const int hostHeight = comboH +
                                   (expandPopupHost ? RedSalamander::ViewerFileComboHost::ComputeStandaloneComboPopupHeightPx(
                                                          _otherFiles.size(), dpi)
                                                    : 0);
            SetWindowPos(_hFileComboHost.get(), HWND_TOP, comboX, comboY, comboW, hostHeight, SWP_NOACTIVATE);
            if (_fileComboControl)
            {
                _fileComboControl->SetBounds(D2D1::RectF(0.0f,
                                                         0.0f,
                                                         Common::WindowSizing::PixelToDip(static_cast<float>(comboW), static_cast<float>(dpi)),
                                                         Common::WindowSizing::PixelToDip(static_cast<float>(comboH), static_cast<float>(dpi))));
                _fileComboHost.Invalidate();
            }
        }
    }

    if (_webViewController)
    {
        RECT bounds   = _contentRect;
        bounds.right  = std::max(bounds.right, bounds.left);
        bounds.bottom = std::max(bounds.bottom, bounds.top);
        static_cast<void>(_webViewController->put_Bounds(bounds));
    }
}

void ViewerWeb::ComputeLayoutRects(HWND hwnd) noexcept
{
    RECT client{};
    if (! hwnd || GetClientRect(hwnd, &client) == 0)
    {
        _headerRect  = {};
        _contentRect = {};
        return;
    }

    client.top += _menuBarHost.GetHwnd() ? _menuBarHost.GetVisibleHeightPx() : 0;

    const UINT dpi             = GetDpiForWindow(hwnd);
    const int baseHeaderHeight = _embeddedMode ? 0 : Common::WindowSizing::DipToPixelRounded(dpi, kHeaderHeightDip);
    const int accentH =
        std::max(1L, Common::WindowSizing::DipToPixelRounded(dpi, RedSalamander::ViewerFileComboHost::kStandaloneComboAccentHeightDip));
    const int accentGap =
        std::max(1L, Common::WindowSizing::DipToPixelRounded(dpi, RedSalamander::ViewerFileComboHost::kStandaloneComboAccentGapDip));
    const int minPadding = Common::WindowSizing::DipToPixelRounded(dpi, RedSalamander::ViewerFileComboHost::kStandaloneComboChromePaddingDip);
    const bool showCombo         = _hFileComboHost && ! _embeddedMode && _otherFiles.size() > 1;
    const int desiredComboHeight = std::max(1, static_cast<int>(Common::WindowSizing::DipToPixelRounded(dpi, RedSalamander::ViewerFileComboHost::kStandaloneComboHeightDip)));

    const int minChromeHeight = Common::WindowSizing::DipToPixelRounded(dpi, 22) + accentH + accentGap + 2 * minPadding;
    int headerH               = _embeddedMode ? 0 : std::max(baseHeaderHeight, minChromeHeight);
    if (showCombo)
    {
        headerH = std::max(headerH, desiredComboHeight + accentH + accentGap + 2 * minPadding);
    }

    _headerRect        = client;
    _headerRect.bottom = std::min(client.bottom, client.top + headerH);

    _contentRect     = client;
    _contentRect.top = _headerRect.bottom;

    _headerRect.left   = std::max<LONG>(0, _headerRect.left);
    _headerRect.top    = std::max<LONG>(0, _headerRect.top);
    _headerRect.right  = std::max(_headerRect.right, _headerRect.left);
    _headerRect.bottom = std::max(_headerRect.bottom, _headerRect.top);

    _contentRect.left   = std::max<LONG>(0, _contentRect.left);
    _contentRect.top    = std::max<LONG>(0, _contentRect.top);
    _contentRect.right  = std::max(_contentRect.right, _contentRect.left);
    _contentRect.bottom = std::max(_contentRect.bottom, _contentRect.top);
}

void ViewerWeb::ApplyTheme(HWND hwnd) noexcept
{
    const COLORREF bg = _hasTheme ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);

    COLORREF headerBg = bg;
    if (_hasTheme && _theme.darkMode)
    {
        headerBg = RGB(std::max(0, GetRValue(bg) - 10), std::max(0, GetGValue(bg) - 10), std::max(0, GetBValue(bg) - 10));
    }
    else
    {
        headerBg = RGB(std::max(0, GetRValue(bg) - 5), std::max(0, GetGValue(bg) - 5), std::max(0, GetBValue(bg) - 5));
    }

    _headerBrush.reset(CreateSolidBrush(headerBg));

    if (_hasTheme && _hWnd)
    {
        const bool windowActive = GetActiveWindow() == _hWnd.get();
        ApplyTitleBarTheme(windowActive);
    }

    _fileComboHost.SetTheme(_hasTheme ? MakeThemePaletteFromViewerTheme(_theme) : MakeDefaultThemePalette(false));

    ApplyMenuTheme(hwnd);
    UpdateMenuState(hwnd);
    ApplyWebViewThemeScript();
}

void ViewerWeb::UpdateMenuState(HWND hwnd, bool syncDxMenuBar) noexcept
{
    HMENU menu = _menuHandle ? _menuHandle.get() : (hwnd ? GetMenu(hwnd) : nullptr);
    if (! menu)
    {
        return;
    }

    const bool jsonInteractiveMode = _kind == ViewerWebKind::Json && _jsonExpandCollapseAvailable;

    EnableMenuItem(menu,
                   IDM_VIEWERWEB_VIEW_DEVTOOLS,
                   static_cast<UINT>(MF_BYCOMMAND | (_documentScriptsEnabled && _config.devToolsEnabled ? MF_ENABLED : MF_GRAYED)));
    EnableMenuItem(menu,
                   IDM_VIEWERWEB_TOOLS_OPEN_EXTERNAL,
                   static_cast<UINT>(MF_BYCOMMAND | (_config.allowExternalNavigation ? MF_ENABLED : MF_GRAYED)));

    EnableMenuItem(menu, IDM_VIEWERWEB_TOOLS_JSON_EXPAND_ALL, static_cast<UINT>(MF_BYCOMMAND | (jsonInteractiveMode ? MF_ENABLED : MF_GRAYED)));
    EnableMenuItem(menu, IDM_VIEWERWEB_TOOLS_JSON_COLLAPSE_ALL, static_cast<UINT>(MF_BYCOMMAND | (jsonInteractiveMode ? MF_ENABLED : MF_GRAYED)));
    EnableMenuItem(
        menu, IDM_VIEWERWEB_TOOLS_MARKDOWN_TOGGLE_SOURCE, static_cast<UINT>(MF_BYCOMMAND | (_kind == ViewerWebKind::Markdown ? MF_ENABLED : MF_GRAYED)));

    CheckMenuItem(menu,
                  IDM_VIEWERWEB_TOOLS_MARKDOWN_TOGGLE_SOURCE,
                  static_cast<UINT>(MF_BYCOMMAND | ((_kind == ViewerWebKind::Markdown && _markdownShowSource) ? MF_CHECKED : MF_UNCHECKED)));

    if (syncDxMenuBar && _menuBarHost.GetHwnd())
    {
        _menuBarHost.SyncMenuModel();
    }
    else if (hwnd)
    {
        DrawMenuBar(hwnd);
    }
}

void ViewerWeb::ApplyTitleBarTheme(bool windowActive) noexcept
{
    if (! _hasTheme || ! _hWnd)
    {
        return;
    }

    RedSalamander::ViewerChrome::ApplyTitleBarTheme(_hWnd.get(), _theme, windowActive, L"title");
}

void ViewerWeb::ApplyMenuTheme(HWND hwnd) noexcept
{
    static_cast<void>(hwnd);
    _menuBarHost.SetTheme(_hasTheme ? MakeThemePaletteFromViewerTheme(_theme) : MakeDefaultThemePalette(false));
    if (_menuBarHost.GetHwnd())
    {
        _menuBarHost.SyncMenuModel();
    }
}

void ViewerWeb::ShowHostAlert(HWND targetWindow, HostAlertSeverity severity, const std::wstring& message) noexcept
{
    if (message.empty())
    {
        return;
    }

    const std::wstring title = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_TITLE);

    if (! _hostAlerts && _host)
    {
        wil::com_ptr<IHostAlerts> alerts;
        const HRESULT hr = _host->QueryInterface(__uuidof(IHostAlerts), alerts.put_void());
        if (SUCCEEDED(hr) && alerts)
        {
            _hostAlerts = std::move(alerts);
        }
    }

    if (! _hostAlerts)
    {
        const std::wstring_view normalizedTitle = title.empty() ? L"ViewerWeb" : std::wstring_view(title);
        switch (severity)
        {
            case HOST_ALERT_WARNING: Debug::Warning(L"ViewerWeb: Host alert dropped ({}): {}", normalizedTitle, message); break;
            case HOST_ALERT_INFO:
            case HOST_ALERT_BUSY: Debug::Info(L"ViewerWeb: Host alert dropped ({}): {}", normalizedTitle, message); break;
            case HOST_ALERT_ERROR:
            default: Debug::Error(L"ViewerWeb: Host alert dropped ({}): {}", normalizedTitle, message); break;
        }
        return;
    }

    HostAlertRequest request{};
    request.version      = 1;
    request.sizeBytes    = sizeof(request);
    request.scope        = (targetWindow && IsWindow(targetWindow)) ? HOST_ALERT_SCOPE_WINDOW : HOST_ALERT_SCOPE_APPLICATION;
    request.modality     = HOST_ALERT_MODELESS;
    request.severity     = severity;
    request.targetWindow = request.scope == HOST_ALERT_SCOPE_WINDOW ? targetWindow : nullptr;
    request.title        = title.empty() ? nullptr : title.c_str();
    request.message      = message.c_str();
    request.closable     = TRUE;

    static_cast<void>(_hostAlerts->ShowAlert(&request, targetWindow));
}

bool ViewerWeb::OfferTextViewerFallbackPrompt() noexcept
{
    if (! _host || _statusMessage.empty())
    {
        return false;
    }

    wil::com_ptr<IHostPrompts> prompts;
    const HRESULT promptsHr = _host->QueryInterface(__uuidof(IHostPrompts), prompts.put_void());
    if (FAILED(promptsHr) || ! prompts)
    {
        return false;
    }

    const std::wstring title   = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_TITLE);
    std::wstring promptMessage = FormatStringResource(g_hInstance, IDS_VIEWERWEB_PROMPT_OPEN_TEXT_VIEWER_FMT, _statusMessage);
    if (promptMessage.empty())
    {
        promptMessage = _statusMessage;
    }

    HostPromptRequest request{};
    request.version       = 1;
    request.sizeBytes     = sizeof(request);
    request.scope         = (_hWnd && IsWindow(_hWnd.get())) ? HOST_ALERT_SCOPE_WINDOW : HOST_ALERT_SCOPE_APPLICATION;
    request.severity      = HOST_ALERT_WARNING;
    request.buttons       = HOST_PROMPT_BUTTONS_YES_NO;
    request.targetWindow  = request.scope == HOST_ALERT_SCOPE_WINDOW ? _hWnd.get() : nullptr;
    request.title         = title.empty() ? nullptr : title.c_str();
    request.message       = promptMessage.c_str();
    request.defaultResult = HOST_PROMPT_RESULT_YES;

    HostPromptResult promptResult = HOST_PROMPT_RESULT_NONE;
    const HRESULT promptHr        = prompts->ShowPrompt(&request, _hWnd.get(), &promptResult);
    if (FAILED(promptHr))
    {
        return false;
    }

    if (promptResult != HOST_PROMPT_RESULT_YES)
    {
        return true;
    }

    const HRESULT openHr = OpenCurrentDocumentInTextViewer();
    if (FAILED(openHr))
    {
        ShowHostAlert(_hWnd.get(), HOST_ALERT_ERROR, LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_OPEN_TEXT_VIEWER_FAILED));
        return true;
    }

    if (_hWnd)
    {
        PostMessageW(_hWnd.get(), WM_CLOSE, 0, 0);
    }

    return true;
}

HRESULT ViewerWeb::OpenCurrentDocumentInTextViewer() noexcept
{
    if (! _host || ! _fileSystem || _currentPath.empty())
    {
        return E_UNEXPECTED;
    }

    wil::com_ptr<IHostViewers> hostViewers;
    const HRESULT viewersHr = _host->QueryInterface(__uuidof(IHostViewers), hostViewers.put_void());
    if (FAILED(viewersHr) || ! hostViewers)
    {
        return FAILED(viewersHr) ? viewersHr : E_NOINTERFACE;
    }

    std::vector<const wchar_t*> otherFilePointers;
    otherFilePointers.reserve(std::max<size_t>(_otherFiles.size(), 1u));
    for (const auto& otherPath : _otherFiles)
    {
        if (! otherPath.empty())
        {
            otherFilePointers.push_back(otherPath.c_str());
        }
    }
    if (otherFilePointers.empty())
    {
        otherFilePointers.push_back(_currentPath.c_str());
    }

    HostViewerOpenRequest request{};
    request.version               = 1;
    request.sizeBytes             = sizeof(request);
    request.pluginId              = L"builtin/viewer-text";
    request.ownerWindow           = _hWnd.get();
    request.fileSystem            = _fileSystem.get();
    request.fileSystemName        = _fileSystemName.empty() ? nullptr : _fileSystemName.c_str();
    request.focusedPath           = _currentPath.c_str();
    request.selectionPaths        = nullptr;
    request.selectionCount        = 0;
    request.otherFiles            = otherFilePointers.data();
    request.otherFileCount        = static_cast<unsigned long>(otherFilePointers.size());
    request.focusedOtherFileIndex = _otherIndex < otherFilePointers.size() ? static_cast<unsigned long>(_otherIndex) : 0;
    request.viewerFlags           = VIEWER_OPEN_FLAG_NONE;

    return hostViewers->OpenViewer(&request);
}

HRESULT ViewerWeb::HandleNavigationStarting(ICoreWebView2NavigationStartingEventArgs* args,
                                            ViewerWebSecurity::NavigationSurface surface) noexcept
{
    if (! args)
    {
        return E_POINTER;
    }

    wil::unique_cotaskmem_string uri;
    const HRESULT uriHr = args->get_Uri(uri.put());
    BOOL userInitiated  = FALSE;
    static_cast<void>(args->get_IsUserInitiated(&userInitiated));
    if (FAILED(uriHr) || ! uri)
    {
        const HRESULT cancelHr = args->put_Cancel(TRUE);
        if (FAILED(cancelHr))
        {
            _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_WEBVIEW2_INIT_FAILED);
            Debug::Error(std::format(L"ViewerWeb: failed to cancel an unreadable navigation request (hr=0x{:08X}).",
                                     static_cast<unsigned long>(cancelHr)));
            DiscardWebView2();
            ShowHostAlert(_hWnd.get(), HOST_ALERT_ERROR, _statusMessage);
            return cancelHr;
        }
        return S_OK;
    }

    const ViewerWebSecurity::NavigationAction action = ViewerWebSecurity::EvaluateNavigation(
        uri.get(), surface, userInitiated != FALSE, _config.allowExternalNavigation, _allowedDocumentUrl);
    if (action == ViewerWebSecurity::NavigationAction::AllowInViewer)
    {
        return S_OK;
    }

    const HRESULT cancelHr = args->put_Cancel(TRUE);
    if (FAILED(cancelHr))
    {
        // Do not launch an external browser unless the in-view request was
        // successfully canceled first. Close the partially trusted controller
        // so a failed Cancel call cannot continue the navigation in-view.
        _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_WEBVIEW2_INIT_FAILED);
        Debug::Error(std::format(L"ViewerWeb: failed to cancel a blocked navigation (hr=0x{:08X}).", static_cast<unsigned long>(cancelHr)));
        DiscardWebView2();
        ShowHostAlert(_hWnd.get(), HOST_ALERT_ERROR, _statusMessage);
        return cancelHr;
    }
    if (action == ViewerWebSecurity::NavigationAction::OpenExternal)
    {
        const HINSTANCE shellResult = ShellExecuteW(_hWnd.get(), L"open", uri.get(), nullptr, nullptr, SW_SHOWNORMAL);
        if (reinterpret_cast<INT_PTR>(shellResult) <= 32)
        {
            Debug::Warning(L"ViewerWeb: the system browser rejected an allowlisted external navigation.");
        }
    }
    return S_OK;
}

HRESULT ViewerWeb::HandleNewWindowRequested(ICoreWebView2NewWindowRequestedEventArgs* args) noexcept
{
    if (! args)
    {
        return E_POINTER;
    }

    wil::unique_cotaskmem_string uri;
    const HRESULT uriHr = args->get_Uri(uri.put());
    BOOL userInitiated  = FALSE;
    static_cast<void>(args->get_IsUserInitiated(&userInitiated));
    const HRESULT handledHr = args->put_Handled(TRUE);
    if (FAILED(handledHr))
    {
        // Failing to suppress the WebView popup is a terminal policy failure;
        // close the controller and never launch a second external copy.
        _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_WEBVIEW2_INIT_FAILED);
        Debug::Error(std::format(L"ViewerWeb: failed to suppress a new-window request (hr=0x{:08X}).", static_cast<unsigned long>(handledHr)));
        DiscardWebView2();
        ShowHostAlert(_hWnd.get(), HOST_ALERT_ERROR, _statusMessage);
        return handledHr;
    }
    if (FAILED(uriHr) || ! uri)
    {
        return S_OK;
    }

    const ViewerWebSecurity::NavigationAction action = ViewerWebSecurity::EvaluateNavigation(uri.get(),
                                                                                              ViewerWebSecurity::NavigationSurface::NewWindow,
                                                                                              userInitiated != FALSE,
                                                                                              _config.allowExternalNavigation,
                                                                                              _allowedDocumentUrl);
    if (action == ViewerWebSecurity::NavigationAction::OpenExternal)
    {
        const HINSTANCE shellResult = ShellExecuteW(_hWnd.get(), L"open", uri.get(), nullptr, nullptr, SW_SHOWNORMAL);
        if (reinterpret_cast<INT_PTR>(shellResult) <= 32)
        {
            Debug::Warning(L"ViewerWeb: the system browser rejected an allowlisted new-window navigation.");
        }
    }
    return S_OK;
}

HRESULT ViewerWeb::CreateControllerFromEnvironment(HWND hwnd, ICoreWebView2Environment* environment, uint64_t sharedEnvironmentGeneration) noexcept
{
    if (! environment)
    {
        _webViewInitInProgress = false;
        Release();
        return E_POINTER;
    }

    wil::com_ptr<ICoreWebView2Environment> controllerEnvironment = environment;
    auto callback        = MakeComCallback<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler, HRESULT, ICoreWebView2Controller*>(
        [this, hwnd, sharedEnvironmentGeneration, controllerEnvironment = std::move(controllerEnvironment)](
            HRESULT controllerResult, ICoreWebView2Controller* controller) -> HRESULT
    {
        _webViewInitInProgress = false;
        auto releaseSelf       = wil::scope_exit([&] { Release(); });

        if (! IsSharedEnvironmentGenerationCurrent(sharedEnvironmentGeneration))
        {
            if (controller)
            {
                controller->Close();
            }
            return S_OK;
        }

        if (FAILED(controllerResult) || ! controller)
        {
            ShowHostAlert(hwnd, HOST_ALERT_ERROR, LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_WEBVIEW2_INIT_FAILED));
            return S_OK;
        }

        if (! _hWnd || _hWnd.get() != hwnd)
        {
            controller->Close();
            return S_OK;
        }

        _webViewController = controller;

        wil::com_ptr<ICoreWebView2> webView;
        const HRESULT webViewHr = controller->get_CoreWebView2(webView.put());
        if (FAILED(webViewHr) || ! webView)
        {
            ShowHostAlert(hwnd, HOST_ALERT_ERROR, LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_WEBVIEW2_INIT_FAILED));
            DiscardWebView2();
            return S_OK;
        }

        _webView = std::move(webView);

        const auto failSecuritySetup = [&](HRESULT failureHr) noexcept -> HRESULT
        {
            Debug::Error(std::format(L"ViewerWeb: fail-closed WebView2 security setup failed (hr=0x{:08X}).",
                                     static_cast<unsigned long>(failureHr)));
            _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_WEBVIEW2_INIT_FAILED);
            DiscardWebView2();
            ShowHostAlert(hwnd, HOST_ALERT_ERROR, _statusMessage);
            if (_hWnd)
            {
                InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
            }
            return S_OK;
        };

        // Security-critical settings and event hooks must all exist before the
        // first navigation. A partially configured controller is discarded.
        HRESULT securityHr = ConfigureWebViewSettings();
        if (FAILED(securityHr))
        {
            return failSecuritySetup(securityHr);
        }

        auto topNavigationHandler =
            MakeComCallback<ICoreWebView2NavigationStartingEventHandler, ICoreWebView2*, ICoreWebView2NavigationStartingEventArgs*>(
                [this](ICoreWebView2* /*sender*/, ICoreWebView2NavigationStartingEventArgs* args) -> HRESULT
        { return HandleNavigationStarting(args, ViewerWebSecurity::NavigationSurface::TopLevel); });
        if (! topNavigationHandler)
        {
            return failSecuritySetup(E_OUTOFMEMORY);
        }
        securityHr = _webView->add_NavigationStarting(topNavigationHandler.get(), &_navStartingToken);
        if (FAILED(securityHr))
        {
            return failSecuritySetup(securityHr);
        }

        auto frameNavigationHandler =
            MakeComCallback<ICoreWebView2NavigationStartingEventHandler, ICoreWebView2*, ICoreWebView2NavigationStartingEventArgs*>(
                [this](ICoreWebView2* /*sender*/, ICoreWebView2NavigationStartingEventArgs* args) -> HRESULT
        { return HandleNavigationStarting(args, ViewerWebSecurity::NavigationSurface::Frame); });
        if (! frameNavigationHandler)
        {
            return failSecuritySetup(E_OUTOFMEMORY);
        }
        securityHr = _webView->add_FrameNavigationStarting(frameNavigationHandler.get(), &_frameNavStartingToken);
        if (FAILED(securityHr))
        {
            return failSecuritySetup(securityHr);
        }

        auto newWindowHandler =
            MakeComCallback<ICoreWebView2NewWindowRequestedEventHandler, ICoreWebView2*, ICoreWebView2NewWindowRequestedEventArgs*>(
                [this](ICoreWebView2* /*sender*/, ICoreWebView2NewWindowRequestedEventArgs* args) -> HRESULT
        { return HandleNewWindowRequested(args); });
        if (! newWindowHandler)
        {
            return failSecuritySetup(E_OUTOFMEMORY);
        }
        securityHr = _webView->add_NewWindowRequested(newWindowHandler.get(), &_newWindowRequestedToken);
        if (FAILED(securityHr))
        {
            return failSecuritySetup(securityHr);
        }

        wil::com_ptr<ICoreWebView2Environment> environmentKeepAlive = controllerEnvironment;
        auto webResourceHandler = MakeComCallback<ICoreWebView2WebResourceRequestedEventHandler, ICoreWebView2*, ICoreWebView2WebResourceRequestedEventArgs*>(
            [this, environmentKeepAlive = std::move(environmentKeepAlive)](ICoreWebView2* /*sender*/,
                                                                           ICoreWebView2WebResourceRequestedEventArgs* args) -> HRESULT
        {
            const auto failInternalRequest = [&](HRESULT failureHr) noexcept -> HRESULT
            {
                _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_WEBVIEW2_INIT_FAILED);
                Debug::Error(std::format(L"ViewerWeb: failed to intercept the private-origin document request (hr=0x{:08X}).",
                                         static_cast<unsigned long>(failureHr)));
                DiscardWebView2();
                ShowHostAlert(_hWnd.get(), HOST_ALERT_ERROR, _statusMessage);
                return failureHr;
            };

            if (! args || ! environmentKeepAlive)
            {
                return failInternalRequest(E_POINTER);
            }

            wil::com_ptr<ICoreWebView2WebResourceRequest> request;
            const HRESULT requestHr = args->get_Request(request.put());
            if (FAILED(requestHr) || ! request)
            {
                return failInternalRequest(FAILED(requestHr) ? requestHr : E_NOINTERFACE);
            }

            wil::unique_cotaskmem_string uri;
            const HRESULT uriHr = request->get_Uri(uri.put());
            if (FAILED(uriHr) || ! uri)
            {
                return failInternalRequest(FAILED(uriHr) ? uriHr : E_INVALIDARG);
            }

            const std::wstring_view requestUrl(uri.get());
            if (! IsInternalDocumentUrl(requestUrl))
            {
                return failInternalRequest(E_ACCESSDENIED);
            }

            wil::com_ptr<ICoreWebView2WebResourceResponse> response;
            if (_internalDocumentUrl.has_value() && OrdinalString::EqualsNoCase(requestUrl, _internalDocumentUrl.value()) && _pendingDocumentUtf8.has_value())
            {
                const auto* data = reinterpret_cast<const BYTE*>(_pendingDocumentUtf8->data());
                const UINT size  = static_cast<UINT>(std::min<size_t>(_pendingDocumentUtf8->size(), static_cast<size_t>(std::numeric_limits<UINT>::max())));
                wil::com_ptr<IStream> stream;
                stream.attach(SHCreateMemStream(data, size));
                if (! stream)
                {
                    return failInternalRequest(E_OUTOFMEMORY);
                }

                const wchar_t* responseHeaders = _documentRoute == ViewerWebSecurity::DocumentRoute::RawHtmlPrivateOrigin
                                                     ? ViewerWebSecurity::kRawHtmlResponseHeaders
                                                     : ViewerWebSecurity::kGeneratedDocumentResponseHeaders;
                const HRESULT responseHr =
                    environmentKeepAlive->CreateWebResourceResponse(stream.get(), 200, L"OK", responseHeaders, response.put());
                if (FAILED(responseHr) || ! response)
                {
                    return failInternalRequest(FAILED(responseHr) ? responseHr : E_NOINTERFACE);
                }
            }
            else
            {
                const HRESULT responseHr = environmentKeepAlive->CreateWebResourceResponse(
                    nullptr, 404, L"Not Found", L"Content-Type: text/plain; charset=utf-8\r\nCache-Control: no-store\r\n", response.put());
                if (FAILED(responseHr) || ! response)
                {
                    return failInternalRequest(FAILED(responseHr) ? responseHr : E_NOINTERFACE);
                }
            }

            const HRESULT installResponseHr = args->put_Response(response.get());
            return FAILED(installResponseHr) ? failInternalRequest(installResponseHr) : S_OK;
        });
        if (! webResourceHandler)
        {
            return failSecuritySetup(E_OUTOFMEMORY);
        }
        securityHr = _webView->add_WebResourceRequested(webResourceHandler.get(), &_webResourceRequestedToken);
        if (FAILED(securityHr))
        {
            return failSecuritySetup(securityHr);
        }
        securityHr =
            _webView->AddWebResourceRequestedFilter(ViewerWebSecurity::kInternalDocumentFilter.data(), COREWEBVIEW2_WEB_RESOURCE_CONTEXT_DOCUMENT);
        if (FAILED(securityHr))
        {
            return failSecuritySetup(securityHr);
        }

        static_cast<void>(_webView->add_NavigationCompleted(
            MakeComCallback<ICoreWebView2NavigationCompletedEventHandler, ICoreWebView2*, ICoreWebView2NavigationCompletedEventArgs*>(
                [this](ICoreWebView2* /*sender*/, ICoreWebView2NavigationCompletedEventArgs* args) -> HRESULT
        {
            BOOL succeeded = FALSE;
            if (args)
            {
                static_cast<void>(args->get_IsSuccess(&succeeded));
            }
            _navigationCompleted = true;
            _navigationSucceeded = succeeded != FALSE;
            ApplyWebViewThemeScript();
            return S_OK;
        }).get(),
            &_navCompletedToken));

        static_cast<void>(_webViewController->add_AcceleratorKeyPressed(
            MakeComCallback<ICoreWebView2AcceleratorKeyPressedEventHandler, ICoreWebView2Controller*, ICoreWebView2AcceleratorKeyPressedEventArgs*>(
                [this](ICoreWebView2Controller* /*sender*/, ICoreWebView2AcceleratorKeyPressedEventArgs* args) -> HRESULT
        {
            if (! args || ! _hWnd)
            {
                return S_OK;
            }

            AddRef();
            auto releaseSelf = wil::scope_exit([&] { Release(); });

            COREWEBVIEW2_KEY_EVENT_KIND kind{};
            if (FAILED(args->get_KeyEventKind(&kind)))
            {
                return S_OK;
            }

            if (kind != COREWEBVIEW2_KEY_EVENT_KIND_KEY_DOWN && kind != COREWEBVIEW2_KEY_EVENT_KIND_SYSTEM_KEY_DOWN)
            {
                return S_OK;
            }

            UINT vk = 0;
            static_cast<void>(args->get_VirtualKey(&vk));

            const bool ctrl          = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
            const bool shift         = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
            const HKL keyboardLayout = GetKeyboardLayout(0);
            const UINT zoomInVk      = ZoomInVirtualKeyForLayout(keyboardLayout);
            const UINT zoomOutVk     = ZoomOutVirtualKeyForLayout(keyboardLayout);

            const auto handle = [&](bool handled) noexcept
            {
                if (handled)
                {
                    static_cast<void>(args->put_Handled(TRUE));
                }
            };

            if (vk == VK_ESCAPE)
            {
                const HWND hwndToClose = _hWnd.get();
                if (hwndToClose)
                {
                    PostMessageW(hwndToClose, WM_CLOSE, 0, 0);
                }
                handle(true);
                return S_OK;
            }

            if (vk == VK_F5)
            {
                static_cast<void>(OpenPath(_hWnd.get(), _currentPath, false));
                handle(true);
                return S_OK;
            }

            if (vk == VK_F12)
            {
                CommandToggleDevTools();
                handle(true);
                return S_OK;
            }

            if (vk == VK_F3)
            {
                shift ? CommandFindPrevious(_hWnd.get()) : CommandFindNext(_hWnd.get());
                handle(true);
                return S_OK;
            }

            if (ctrl && (vk == 'F' || vk == 'f'))
            {
                CommandFind(_hWnd.get());
                handle(true);
                return S_OK;
            }

            if (ctrl && (vk == 'S' || vk == 's'))
            {
                static_cast<void>(CommandSaveAs(_hWnd.get()));
                handle(true);
                return S_OK;
            }

            if (ctrl && (vk == 'L' || vk == 'l'))
            {
                CommandCopyUrl(_hWnd.get());
                handle(true);
                return S_OK;
            }

            if (ctrl && vk == VK_RETURN)
            {
                CommandOpenExternal(_hWnd.get());
                handle(true);
                return S_OK;
            }

            if (ctrl && (vk == VK_ADD || vk == zoomInVk || vk == '='))
            {
                CommandZoomIn();
                handle(true);
                return S_OK;
            }

            if (ctrl && (vk == VK_SUBTRACT || vk == zoomOutVk || vk == '-'))
            {
                CommandZoomOut();
                handle(true);
                return S_OK;
            }

            if (ctrl && vk == '0')
            {
                CommandZoomReset();
                handle(true);
                return S_OK;
            }

            if (ctrl && vk == VK_OEM_3)
            {
                CommandMarkdownToggleSource();
                handle(true);
                return S_OK;
            }

            if (ctrl && vk == VK_UP)
            {
                SendMessageW(_hWnd.get(), WM_COMMAND, IDM_VIEWERWEB_OTHER_PREVIOUS, 0);
                handle(true);
                return S_OK;
            }

            if (ctrl && vk == VK_DOWN)
            {
                SendMessageW(_hWnd.get(), WM_COMMAND, IDM_VIEWERWEB_OTHER_NEXT, 0);
                handle(true);
                return S_OK;
            }

            if (ctrl && vk == VK_HOME)
            {
                SendMessageW(_hWnd.get(), WM_COMMAND, IDM_VIEWERWEB_OTHER_FIRST, 0);
                handle(true);
                return S_OK;
            }

            if (ctrl && vk == VK_END)
            {
                SendMessageW(_hWnd.get(), WM_COMMAND, IDM_VIEWERWEB_OTHER_LAST, 0);
                handle(true);
                return S_OK;
            }

            if (_kind != ViewerWebKind::Web && vk == VK_SPACE)
            {
                SendMessageW(_hWnd.get(), WM_COMMAND, IDM_VIEWERWEB_OTHER_NEXT, 0);
                handle(true);
                return S_OK;
            }

            if (_kind != ViewerWebKind::Web && vk == VK_BACK)
            {
                SendMessageW(_hWnd.get(), WM_COMMAND, IDM_VIEWERWEB_OTHER_PREVIOUS, 0);
                handle(true);
                return S_OK;
            }

            return S_OK;
        }).get(),
            &_accelToken));

        Layout(hwnd);
        ApplyWebViewThemeScript();

        const HRESULT navHr = NavigatePendingContent(hwnd);
        if (FAILED(navHr))
        {
            const std::wstring message = _statusMessage.empty()
                                             ? FormatStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_NAVIGATE_FAILED_FMT, static_cast<unsigned long>(navHr))
                                             : _statusMessage;
            ShowHostAlert(hwnd, HOST_ALERT_ERROR, message);
        }

        return S_OK;
    });

    if (! callback)
    {
        _webViewInitInProgress = false;
        Release();
        return E_OUTOFMEMORY;
    }

    const HRESULT hr = environment->CreateCoreWebView2Controller(hwnd, callback.get());

    if (FAILED(hr))
    {
        _webViewInitInProgress = false;
        ShowHostAlert(hwnd, HOST_ALERT_ERROR, LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_WEBVIEW2_INIT_FAILED));
        Release();
    }

    return hr;
}

HRESULT ViewerWeb::NavigatePendingContent(HWND /*hwnd*/) noexcept
{
    if (! _webView)
    {
        return E_UNEXPECTED;
    }

    if (_pendingPath.has_value())
    {
        const std::wstring url = std::move(_pendingPath.value());
        _pendingPath.reset();
        _navigationCompleted = false;
        _navigationSucceeded = false;
        return _webView->Navigate(url.c_str());
    }

    if (_pendingWebContent.has_value())
    {
        const std::wstring html = std::move(_pendingWebContent.value());
        _pendingWebContent.reset();
        if (html.empty())
        {
            _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_BUILD_HTML_DOCUMENT);
            return E_INVALIDARG;
        }

        _navigationCompleted = false;
        _navigationSucceeded = false;
        const HRESULT navHr = _webView->NavigateToString(html.c_str());
        if (FAILED(navHr))
        {
            _statusMessage = FormatStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_NAVIGATE_TO_STRING_FAILED_FMT, static_cast<unsigned long>(navHr));
        }
        return navHr;
    }

    if (_pendingDocumentUtf8.has_value() && _internalDocumentUrl.has_value())
    {
        _navigationCompleted = false;
        _navigationSucceeded = false;
        return _webView->Navigate(_internalDocumentUrl->c_str());
    }

    return S_FALSE;
}

HRESULT ViewerWeb::EnsureWebView2(HWND hwnd) noexcept
{
    if (_webView)
    {
        return S_OK;
    }

    if (! hwnd)
    {
        return E_INVALIDARG;
    }

    if (_webViewInitInProgress)
    {
        return S_FALSE;
    }

    _webViewInitInProgress = true;
    AddRef();
    const uint64_t sharedEnvironmentGeneration = BeginSharedEnvironmentUse();

    // Fast path: shared environment already exists — create controller immediately.
    if (g_sharedEnvironment.environment)
    {
        return CreateControllerFromEnvironment(hwnd, g_sharedEnvironment.environment.get(), sharedEnvironmentGeneration);
    }

    // Another instance is already creating the environment — queue ourselves.
    if (g_sharedEnvironment.createInProgress)
    {
        g_sharedEnvironment.pendingConsumers.push_back({this, hwnd});
        return S_FALSE;
    }

    // First caller — create the shared environment.
    g_sharedEnvironment.createInProgress = true;

    const std::wstring userDataFolder = GetWebView2UserDataFolder();

    auto callback        = MakeComCallback<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler, HRESULT, ICoreWebView2Environment*>(
        [this, hwnd, sharedEnvironmentGeneration](HRESULT result, ICoreWebView2Environment* environment) -> HRESULT
    {
        if (! IsSharedEnvironmentGenerationCurrent(sharedEnvironmentGeneration))
        {
            _webViewInitInProgress = false;
            Release();
            return S_OK;
        }

        g_sharedEnvironment.createInProgress = false;

        if (FAILED(result) || ! environment)
        {
            _webViewInitInProgress = false;

            const UINT msgId = (result == HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND) || result == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND))
                                   ? IDS_VIEWERWEB_ERROR_WEBVIEW2_RUNTIME_MISSING
                                   : IDS_VIEWERWEB_ERROR_WEBVIEW2_INIT_FAILED;
            ShowHostAlert(hwnd, HOST_ALERT_ERROR, LoadStringResource(g_hInstance, msgId));

            // Notify all pending consumers about the failure.
            for (auto& pending : g_sharedEnvironment.pendingConsumers)
            {
                if (pending.viewer)
                {
                    pending.viewer->_webViewInitInProgress = false;
                    pending.viewer->ShowHostAlert(pending.hwnd, HOST_ALERT_ERROR, LoadStringResource(g_hInstance, msgId));
                    pending.viewer->Release();
                }
            }
            g_sharedEnvironment.pendingConsumers.clear();

            Release();
            return S_OK;
        }

        g_sharedEnvironment.environment = environment;

        // Create our own controller.
        static_cast<void>(CreateControllerFromEnvironment(hwnd, environment, sharedEnvironmentGeneration));

        // Drain pending consumers — each creates its own controller from the shared environment.
        auto pendingConsumers = std::move(g_sharedEnvironment.pendingConsumers);
        for (auto& pending : pendingConsumers)
        {
            if (pending.viewer && pending.hwnd)
            {
                static_cast<void>(pending.viewer->CreateControllerFromEnvironment(pending.hwnd, environment, sharedEnvironmentGeneration));
            }
        }

        return S_OK;
    });

    if (! callback)
    {
        g_sharedEnvironment.createInProgress = false;
        _webViewInitInProgress               = false;
        _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_WEBVIEW2_INIT_FAILED);
        ShowHostAlert(hwnd, HOST_ALERT_ERROR, _statusMessage);
        Release();
        return E_OUTOFMEMORY;
    }

    const HRESULT hr = CreateCoreWebView2EnvironmentWithOptions(nullptr, userDataFolder.empty() ? nullptr : userDataFolder.c_str(), nullptr, callback.get());

    if (FAILED(hr))
    {
        g_sharedEnvironment.createInProgress = false;
        _webViewInitInProgress               = false;
        ShowHostAlert(hwnd, HOST_ALERT_ERROR, LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_WEBVIEW2_INIT_FAILED));
        Release();
        return hr;
    }

    return S_OK;
}

void ViewerWeb::DiscardWebView2() noexcept
{
    _webViewInitInProgress = false;

    // Unregister event handlers before closing
    if (_webViewController)
    {
        static_cast<void>(_webViewController->remove_AcceleratorKeyPressed(_accelToken));
    }
    if (_webView)
    {
        static_cast<void>(_webView->remove_NavigationStarting(_navStartingToken));
        static_cast<void>(_webView->remove_FrameNavigationStarting(_frameNavStartingToken));
        static_cast<void>(_webView->remove_NewWindowRequested(_newWindowRequestedToken));
        static_cast<void>(_webView->remove_NavigationCompleted(_navCompletedToken));
        static_cast<void>(_webView->remove_WebResourceRequested(_webResourceRequestedToken));
    }

    _navStartingToken          = {};
    _frameNavStartingToken     = {};
    _newWindowRequestedToken   = {};
    _navCompletedToken         = {};
    _accelToken                = {};
    _webResourceRequestedToken = {};

    // Close the WebView2 controller. Note: Close() is asynchronous and may have
    // pending I/O operations that complete on thread pool threads. This is why
    // we must call DiscardWebView2() early in OnDestroy() rather than in the
    // destructor - to give WebView2 time to complete its shutdown before the
    // plugin DLL is unloaded.
    if (_webViewController)
    {
        _webViewController->Close();
    }

    // Release per-instance COM pointers after initiating close.
    // The shared environment is NOT released here — it lives for the DLL lifetime.
    _webView.reset();
    _webViewController.reset();
}

void ViewerWeb::CancelPendingWebView2Initialization() noexcept
{
    _webViewInitInProgress = false;
}

HRESULT ViewerWeb::ConfigureWebViewSettings() noexcept
{
    if (! _webView)
    {
        return E_UNEXPECTED;
    }

    wil::com_ptr<ICoreWebView2Settings> settings;
    const HRESULT settingsHr = _webView->get_Settings(settings.put());
    if (FAILED(settingsHr) || ! settings)
    {
        return FAILED(settingsHr) ? settingsHr : E_NOINTERFACE;
    }

    // These settings are part of the document isolation boundary. A prior
    // generated page may have enabled scripting, so every route transition
    // must apply all three successfully before navigation can proceed.
    HRESULT hr = settings->put_IsScriptEnabled(_documentScriptsEnabled ? TRUE : FALSE);
    if (FAILED(hr))
    {
        return hr;
    }
    hr = settings->put_IsWebMessageEnabled(_documentScriptsEnabled ? TRUE : FALSE);
    if (FAILED(hr))
    {
        return hr;
    }
    hr = settings->put_AreDevToolsEnabled(_documentScriptsEnabled && _config.devToolsEnabled ? TRUE : FALSE);
    if (FAILED(hr))
    {
        return hr;
    }

    static_cast<void>(settings->put_AreDefaultContextMenusEnabled(TRUE));
    static_cast<void>(settings->put_IsZoomControlEnabled(TRUE));

    wil::com_ptr<ICoreWebView2Settings3> settings3;
    if (SUCCEEDED(settings->QueryInterface(IID_PPV_ARGS(settings3.put()))) && settings3)
    {
        static_cast<void>(settings3->put_AreBrowserAcceleratorKeysEnabled(TRUE));
    }

    return S_OK;
}

void ViewerWeb::ApplyWebViewThemeScript() noexcept
{
    if (_webViewController)
    {
        wil::com_ptr<ICoreWebView2Controller2> controller2;
        if (SUCCEEDED(_webViewController->QueryInterface(IID_PPV_ARGS(controller2.put()))) && controller2)
        {
            const COLORREF bg = _hasTheme ? ColorRefFromArgb(_theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);
            COREWEBVIEW2_COLOR color{};
            color.A = 255;
            color.R = GetRValue(bg);
            color.G = GetGValue(bg);
            color.B = GetBValue(bg);
            static_cast<void>(controller2->put_DefaultBackgroundColor(color));
        }
    }

    if (_webView && _documentScriptsEnabled)
    {
        if (_hasTheme)
        {
            const COLORREF bg     = ColorRefFromArgb(_theme.backgroundArgb);
            const COLORREF fg     = ColorRefFromArgb(_theme.textArgb);
            const COLORREF selBg  = ColorRefFromArgb(_theme.selectionBackgroundArgb);
            const COLORREF selFg  = ColorRefFromArgb(_theme.selectionTextArgb);
            const COLORREF accent = ResolveAccentColor(_theme, _currentPath.empty() ? _metaId : _currentPath);

            const auto cssRgb = [](COLORREF c) -> std::wstring { return std::format(L"rgb({},{},{})", GetRValue(c), GetGValue(c), GetBValue(c)); };

            const std::wstring script = std::format(L"(function(){{try{{if(window.RS&&window.RS.applyTheme){{window.RS.applyTheme({{bg:'{}',fg:'{}',selBg:'{}',"
                                                    L"selFg:'{}',accent:'{}'}});}}}}catch(e){{}}}})();",
                                                    cssRgb(bg),
                                                    cssRgb(fg),
                                                    cssRgb(selBg),
                                                    cssRgb(selFg),
                                                    cssRgb(accent));
            static_cast<void>(_webView->ExecuteScript(script.c_str(), nullptr));
        }
    }
}

HRESULT ViewerWeb::OpenPath(HWND hwnd, const std::wstring& path, bool updateOtherFiles) noexcept
{
    if (! hwnd)
    {
        return E_INVALIDARG;
    }

    if (path.empty())
    {
        Debug::Error(L"ViewerWeb: OpenPath called with an empty path.");
        return E_INVALIDARG;
    }

    if (! _fileSystem)
    {
        Debug::Error(L"ViewerWeb: OpenPath failed because file system is missing.");
        return E_FAIL;
    }

    _currentPath = path;

    if (updateOtherFiles)
    {
        _otherFiles.clear();
        _otherFiles.push_back(path);
        _otherIndex = 0;
        RefreshFileCombo(hwnd);
    }
    else if (! _otherFiles.empty())
    {
        for (size_t i = 0; i < _otherFiles.size(); ++i)
        {
            const std::wstring_view a(_otherFiles[i]);
            const std::wstring_view b(path);
            if (OrdinalString::EqualsNoCase(a, b))
            {
                _otherIndex = i;
                break;
            }
        }

        if (_fileComboControl)
        {
            _syncingFileCombo = true;
            auto restore      = wil::scope_exit([&] { _syncingFileCombo = false; });
            _fileComboControl->SetSelectedIndex(_otherIndex);
            _fileComboHost.Invalidate();
        }
    }

    const std::wstring leaf = LeafNameFromPath(path);
    std::wstring title      = leaf;
    if (! _metaName.empty())
    {
        title = leaf.empty() ? _metaName : std::format(L"{} - {}", leaf, _metaName);
    }
    if (_hWnd && ! title.empty())
    {
        SetWindowTextW(_hWnd.get(), title.c_str());
    }

    _statusMessage               = LoadStringResource(g_hInstance, IDS_VIEWERWEB_STATUS_LOADING);
    _jsonExpandCollapseAvailable = false;
    _generatedOutputRejected     = false;
    _generatedOutputBytes        = 0u;
    _generatedOutputLimit        = 0u;
    _pendingPath.reset();
    _pendingWebContent.reset();

    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
        UpdateMenuState(_hWnd.get());
    }

    return StartAsyncLoad(hwnd, path);
}

void ViewerWeb::RefreshFileCombo(HWND hwnd) noexcept
{
    if (! _fileComboControl)
    {
        return;
    }

    _syncingFileCombo = true;
    auto restore      = wil::scope_exit([&] { _syncingFileCombo = false; });

    std::vector<ComboBox::Item> items;
    items.reserve(_otherFiles.size());
    for (const auto& path : _otherFiles)
    {
        std::wstring leaf = LeafNameFromPath(path);
        if (leaf.empty())
        {
            leaf = path;
        }

        items.push_back(ComboBox::Item{path, std::move(leaf)});
    }
    _fileComboControl->SetItems(std::move(items));

    if (_otherIndex < _otherFiles.size())
    {
        _fileComboControl->SetSelectedIndex(_otherIndex);
    }
    else
    {
        _fileComboControl->SetSelectedIndex(std::nullopt);
    }

    _fileComboHost.Invalidate();
    Layout(hwnd);
}

HRESULT ViewerWeb::StartAsyncLoad(HWND hwnd, const std::wstring& path) noexcept
{
    RetryStagedFileCleanup();

    if (! hwnd)
    {
        return E_INVALIDARG;
    }

    if (path.empty())
    {
        return E_INVALIDARG;
    }

    if (! _fileSystem)
    {
        return E_FAIL;
    }

    const auto failSubmission = [&](HRESULT failureHr) noexcept -> HRESULT
    {
        _statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_READ_FILE_FAILED);
        Debug::Error(std::format(L"ViewerWeb: async load submission failed (hr=0x{:08X}).", static_cast<unsigned long>(failureHr)));
        ShowHostAlert(hwnd, HOST_ALERT_ERROR, _statusMessage);
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
        }
        return failureHr;
    };

    _openRequestId += 1u;
    const uint64_t requestId = _openRequestId;
    _loadPostFailureTerminal = false;

    std::unique_ptr<AsyncLoadResult> payload(new (std::nothrow) AsyncLoadResult{});
    if (! payload)
    {
        return failSubmission(E_OUTOFMEMORY);
    }

    payload->viewer    = this;
    payload->hwnd      = hwnd;
    payload->requestId = requestId;
    payload->path      = path;
    payload->hr        = E_FAIL;

    payload->kindSnapshot               = _kind;
    payload->configSnapshot             = _config;
    payload->hasThemeSnapshot           = _hasTheme;
    payload->themeSnapshot              = _theme;
    payload->markdownShowSourceSnapshot = _markdownShowSource;
    payload->fileSystemSnapshot         = _fileSystem;
    payload->metaIdSnapshot             = _metaId;
    payload->metaNameSnapshot           = _metaName;

    AddRef();

    struct AsyncLoadWorkItem final
    {
        AsyncLoadWorkItem()                                    = default;
        AsyncLoadWorkItem(const AsyncLoadWorkItem&)            = delete;
        AsyncLoadWorkItem& operator=(const AsyncLoadWorkItem&) = delete;

        std::unique_ptr<AsyncLoadResult> payload;
        wil::unique_hmodule moduleKeepAlive;
    };

    auto ctx = std::unique_ptr<AsyncLoadWorkItem>(new (std::nothrow) AsyncLoadWorkItem{});
    if (! ctx)
    {
        Release();
        return failSubmission(E_OUTOFMEMORY);
    }

    ctx->payload         = std::move(payload);
    ctx->moduleKeepAlive = AcquireModuleReferenceFromAddress(&kViewerWebModuleAnchor);
    if (! ctx->moduleKeepAlive)
    {
        Release();
        return failSubmission(E_FAIL);
    }

    g_activeAsyncWorkerCount.fetch_add(1u, std::memory_order_acq_rel);
    const BOOL queued = TrySubmitThreadpoolCallback(
        [](PTP_CALLBACK_INSTANCE instance, void* context) noexcept
    {
        std::unique_ptr<AsyncLoadWorkItem> ctx(static_cast<AsyncLoadWorkItem*>(context));
        auto releaseWorkerGate = wil::scope_exit(
            []() noexcept { g_activeAsyncWorkerCount.fetch_sub(1u, std::memory_order_acq_rel); });
        if (ctx && ctx->moduleKeepAlive)
        {
            TransferModulePinToCallbackReturn(instance, ctx->moduleKeepAlive);
        }
        if (! ctx || ! ctx->payload)
        {
            return;
        }

        AsyncLoadProc(ctx->payload.release());
    },
        ctx.get(),
        nullptr);

    if (queued == 0)
    {
        g_activeAsyncWorkerCount.fetch_sub(1u, std::memory_order_acq_rel);
        Release();
        return failSubmission(E_FAIL);
    }

    ctx.release();
    return S_OK;
}

void ViewerWeb::AsyncLoadProc(AsyncLoadResult* payload) noexcept
{
    std::unique_ptr<AsyncLoadResult> result(payload);
    if (! result || ! result->viewer)
    {
        return;
    }

    ViewerWeb* self  = result->viewer;
    auto releaseSelf = wil::scope_exit([&] { self->Release(); });

    const ViewerWebKind kind      = result->kindSnapshot;
    const ViewerWebConfig config  = result->configSnapshot;
    const bool hasTheme           = result->hasThemeSnapshot;
    const ViewerTheme theme       = result->themeSnapshot;
    const bool markdownShowSource = result->markdownShowSourceSnapshot;

    wil::com_ptr<IFileSystem> fileSystem = result->fileSystemSnapshot;

    const auto cssRgb = [](COLORREF c) { return std::format("rgb({},{},{})", GetRValue(c), GetGValue(c), GetBValue(c)); };

    const COLORREF bg    = hasTheme ? ColorRefFromArgb(theme.backgroundArgb) : GetSysColor(COLOR_WINDOW);
    const COLORREF fg    = hasTheme ? ColorRefFromArgb(theme.textArgb) : GetSysColor(COLOR_WINDOWTEXT);
    const COLORREF selBg = hasTheme ? ColorRefFromArgb(theme.selectionBackgroundArgb) : GetSysColor(COLOR_HIGHLIGHT);
    const COLORREF selFg = hasTheme ? ColorRefFromArgb(theme.selectionTextArgb) : GetSysColor(COLOR_HIGHLIGHTTEXT);
    const COLORREF accent =
        hasTheme ? ResolveAccentColor(theme, result->path.empty() ? result->metaIdSnapshot : std::wstring_view(result->path)) : GetSysColor(COLOR_HIGHLIGHT);

    const std::string themeObj =
        std::format("{{bg:'{}',fg:'{}',selBg:'{}',selFg:'{}',accent:'{}'}}", cssRgb(bg), cssRgb(fg), cssRgb(selBg), cssRgb(selFg), cssRgb(accent));

    const std::wstring leafW  = LeafNameFromPath(result->path);
    const std::wstring titleW = leafW.empty() ? result->metaNameSnapshot : std::format(L"{} - {}", leafW, result->metaNameSnapshot);
    result->title             = titleW;

    result->statusMessage.clear();
    const uint64_t configuredMaxBytes = static_cast<uint64_t>(config.maxDocumentMiB) * 1024ull * 1024ull;

    auto postBack = [&]([[maybe_unused]] bool cleanupTempOnFailure) noexcept
    {
        const HWND hwnd = result->hwnd;
        const uint64_t requestId = result->requestId;
        const std::optional<std::filesystem::path> stagedPath = result->extractedWin32Path;
        const uint64_t generatedOutputBytes = static_cast<uint64_t>(result->utf8.size());
        const uint64_t generatedOutputLimit = ViewerWebSecurity::GeneratedOutputLimit(configuredMaxBytes);
        if (result->documentRoute == ViewerWebSecurity::DocumentRoute::GeneratedPrivateOrigin)
        {
            if (! result->generatedOutputRejected)
            {
                result->generatedOutputBytes = generatedOutputBytes;
                result->generatedOutputLimit = generatedOutputLimit;
            }
        }
        if (result->documentRoute == ViewerWebSecurity::DocumentRoute::GeneratedPrivateOrigin && SUCCEEDED(result->hr) &&
            ! ViewerWebSecurity::IsGeneratedOutputWithinLimit(generatedOutputBytes, configuredMaxBytes))
        {
            result->hr                      = HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
            result->offerTextViewerFallback = true;
            result->generatedOutputRejected = true;
            result->jsonExpandCollapseAvailable = false;
            result->statusMessage = FormatStringResource(g_hInstance,
                                                         IDS_VIEWERWEB_ERROR_FILE_TOO_LARGE_FMT,
                                                         FormatBytesCompact(generatedOutputBytes),
                                                         FormatBytesCompact(generatedOutputLimit));
            std::string{}.swap(result->utf8);
        }
        std::wstring_view detail = L"unknown";
        switch (result->documentRoute)
        {
            case ViewerWebSecurity::DocumentRoute::RawHtmlPrivateOrigin: detail = L"raw-html-private-origin"; break;
            case ViewerWebSecurity::DocumentRoute::GeneratedPrivateOrigin: detail = L"generated-private-origin"; break;
            case ViewerWebSecurity::DocumentRoute::StagedPdf: detail = L"staged-pdf"; break;
            case ViewerWebSecurity::DocumentRoute::None:
            default: detail = kind == ViewerWebKind::Web ? L"web" : kind == ViewerWebKind::Json ? L"json" : L"markdown"; break;
        }
        Debug::Perf::Emit(L"viewer.web.load_bytes", detail, 0u, result->loadedSourceBytes, configuredMaxBytes, result->hr);
        if (result->documentRoute == ViewerWebSecurity::DocumentRoute::GeneratedPrivateOrigin)
        {
            Debug::Perf::Emit(
                L"viewer.web.output_bytes", detail, 0u, result->generatedOutputBytes, result->generatedOutputLimit, result->hr);
        }
        if (result->hr == HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE))
        {
            Debug::Perf::Emit(L"viewer.web.rejected_bytes", detail, 0u, result->loadedSourceBytes, configuredMaxBytes, result->hr);
        }

        bool forcePostFailure = false;
#ifdef ENABLE_TESTS
        forcePostFailure = g_failNextAsyncLoadCompletionPost.exchange(false, std::memory_order_acq_rel);
#endif
        const bool posted = hwnd && ! forcePostFailure && PostMessagePayload(hwnd, kAsyncLoadCompleteMessage, 0, std::move(result));
        if (! posted)
        {
            if (stagedPath.has_value())
            {
                DeleteStagedFileOrSchedule(stagedPath.value());
            }
            self->_asyncLoadPostFailureRequestId.store(requestId, std::memory_order_release);
            self->_asyncLoadPostFailureCount.fetch_add(1u, std::memory_order_relaxed);
            if (hwnd)
            {
                static_cast<void>(PostMessageW(hwnd, kAsyncPostFailureMessage, kAsyncPostFailureLoad, 0));
                static_cast<void>(InvalidateRect(hwnd, nullptr, FALSE));
            }
            return;
        }
    };

    if (! fileSystem)
    {
        result->hr            = E_FAIL;
        result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_FILE_SYSTEM_UNAVAILABLE);
        postBack(false);
        return;
    }

    wil::com_ptr<IFileSystemIO> fileIo;
    const HRESULT fileIoHr = fileSystem->QueryInterface(__uuidof(IFileSystemIO), fileIo.put_void());
    if (FAILED(fileIoHr) || ! fileIo)
    {
        result->hr            = FAILED(fileIoHr) ? fileIoHr : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED);
        result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_NO_FILE_IO);
        postBack(false);
        return;
    }

    if (kind == ViewerWebKind::Web)
    {
        wil::com_ptr<IFileReader> reader;
        const HRESULT openHr = fileIo->CreateFileReader(result->path.c_str(), reader.put());
        if (FAILED(openHr) || ! reader)
        {
            result->hr            = FAILED(openHr) ? openHr : E_FAIL;
            result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_OPEN_FOR_VIEWING_FAILED);
            postBack(false);
            return;
        }

        uint64_t advertisedBytes = 0u;
        const HRESULT sizeHr = reader->GetSize(&advertisedBytes);
        if (FAILED(sizeHr))
        {
            result->hr            = sizeHr;
            result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_READ_FILE_SIZE_FAILED);
            postBack(false);
            return;
        }

        result->loadedSourceBytes = advertisedBytes;
        if (advertisedBytes > configuredMaxBytes)
        {
            result->hr = HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
            result->statusMessage = FormatStringResource(
                g_hInstance, IDS_VIEWERWEB_ERROR_FILE_TOO_LARGE_FMT, FormatBytesCompact(advertisedBytes), FormatBytesCompact(configuredMaxBytes));
            postBack(false);
            return;
        }

        const bool pdfDocument = IsPdfPath(result->path);
        if (! pdfDocument)
        {
            std::string bytes;
            bytes.reserve(static_cast<size_t>(advertisedBytes));
            uint64_t consumedBytes = 0u;
            const HRESULT readHr = ReadProviderExactly(
                reader.get(),
                advertisedBytes,
                [&](std::span<const uint8_t> chunk) noexcept -> HRESULT
                {
                    bytes.append(reinterpret_cast<const char*>(chunk.data()), chunk.size());
                    return S_OK;
                },
                consumedBytes);
            result->loadedSourceBytes = consumedBytes;
            if (FAILED(readHr))
            {
                result->hr            = readHr;
                result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_READ_FILE_FAILED);
                postBack(false);
                return;
            }

            result->utf8              = std::move(bytes);
            result->documentRoute     = ViewerWebSecurity::DocumentRoute::RawHtmlPrivateOrigin;
            result->hr                = S_OK;
            postBack(false);
            return;
        }

        // Every PDF is staged through the provider reader. Path syntax is not
        // proof of provider identity and must never authorize host-local file
        // access through a direct file:// navigation.
        wchar_t tempDir[MAX_PATH]{};
        const DWORD tempDirLen = GetTempPathW(static_cast<DWORD>(std::size(tempDir)), tempDir);
        if (tempDirLen == 0 || tempDirLen >= std::size(tempDir))
        {
            const DWORD error = tempDirLen == 0u ? GetLastError() : ERROR_INSUFFICIENT_BUFFER;
            result->hr            = HRESULT_FROM_WIN32(error);
            result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_GET_TEMP_FOLDER_FAILED);
            postBack(false);
            return;
        }

        std::wstring tempPathText;
        wil::unique_hfile outFile;
        const Common::Paths::UniqueSiblingFileOptions options{.prefix = L"rsw-", .suffix = L".pdf"};
        const HRESULT createHr = Common::Paths::CreateUniqueFileInDirectory(tempDir, options, tempPathText, outFile);
        if (FAILED(createHr) || ! outFile || tempPathText.empty())
        {
            result->hr            = FAILED(createHr) ? createHr : E_FAIL;
            result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_CREATE_TEMP_FILE_FAILED);
            postBack(false);
            return;
        }
        const std::filesystem::path tempPath(tempPathText);

        auto cleanupUntrackedTemp = wil::scope_exit([&]() noexcept
        {
            DeleteStagedFileOrSchedule(tempPath);
        });

        bool writeFailed = false;
        uint64_t consumedBytes = 0u;
        const HRESULT readHr = ReadProviderExactly(
            reader.get(),
            advertisedBytes,
            [&](std::span<const uint8_t> chunk) noexcept -> HRESULT
            {
                const HRESULT writeHr = Common::HandleIo::WriteAll(
                    outFile.get(), std::as_bytes(std::span<const uint8_t>(chunk.data(), chunk.size())));
                writeFailed           = FAILED(writeHr);
                return writeHr;
            },
            consumedBytes);
        result->loadedSourceBytes = consumedBytes;
        if (FAILED(readHr))
        {
            result->hr            = readHr;
            result->statusMessage = LoadStringResource(
                g_hInstance, writeFailed ? IDS_VIEWERWEB_ERROR_WRITE_TEMP_FILE_FAILED : IDS_VIEWERWEB_ERROR_READ_FILE_FAILED);
            postBack(false);
            return;
        }

        if (FlushFileBuffers(outFile.get()) == FALSE)
        {
            result->hr            = HRESULT_FROM_WIN32(GetLastError());
            result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_WRITE_TEMP_FILE_FAILED);
            postBack(false);
            return;
        }
        outFile.reset();
        result->extractedWin32Path = tempPath;
        result->documentRoute      = ViewerWebSecurity::DocumentRoute::StagedPdf;
        result->hr                 = S_OK;
        cleanupUntrackedTemp.release();
        postBack(true);
        return;
    }

    // JSON/Markdown: load file into memory (UTF-8/UTF-16 with BOM supported).
    wil::com_ptr<IFileReader> reader;
    const HRESULT openHr = fileIo->CreateFileReader(result->path.c_str(), reader.put());
    if (FAILED(openHr) || ! reader)
    {
        result->hr            = FAILED(openHr) ? openHr : E_FAIL;
        result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_OPEN_FOR_VIEWING_FAILED);
        postBack(false);
        return;
    }

    uint64_t sizeBytes   = 0;
    const HRESULT sizeHr = reader->GetSize(&sizeBytes);
    if (FAILED(sizeHr))
    {
        result->hr            = sizeHr;
        result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_READ_FILE_SIZE_FAILED);
        postBack(false);
        return;
    }

    const uint64_t maxBytes = configuredMaxBytes;
    if (sizeBytes > maxBytes)
    {
        result->loadedSourceBytes       = sizeBytes;
        result->hr                      = HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
        result->offerTextViewerFallback = true;
        result->statusMessage =
            FormatStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_FILE_TOO_LARGE_FMT, FormatBytesCompact(sizeBytes), FormatBytesCompact(maxBytes));
        postBack(false);
        return;
    }

    std::string bytes;
    bytes.reserve(static_cast<size_t>(sizeBytes));
    uint64_t consumedBytes = 0u;
    const HRESULT readHr = ReadProviderExactly(
        reader.get(),
        sizeBytes,
        [&](std::span<const uint8_t> chunk) noexcept -> HRESULT
        {
            bytes.append(reinterpret_cast<const char*>(chunk.data()), chunk.size());
            return S_OK;
        },
        consumedBytes);
    result->loadedSourceBytes = consumedBytes;
    if (FAILED(readHr))
    {
        result->hr                      = readHr;
        result->offerTextViewerFallback = true;
        result->statusMessage            = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_READ_FILE_FAILED);
        postBack(false);
        return;
    }

    std::string textUtf8;
    const ViewerWebSecurity::NormalizeTextResult normalizeResult =
        ViewerWebSecurity::NormalizeTextUtf8Bounded(bytes, static_cast<size_t>(maxBytes), textUtf8);
    std::string{}.swap(bytes);
    if (normalizeResult != ViewerWebSecurity::NormalizeTextResult::Ok)
    {
        result->hr = normalizeResult == ViewerWebSecurity::NormalizeTextResult::TooLarge ? HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE)
                                                                                          : HRESULT_FROM_WIN32(ERROR_NO_UNICODE_TRANSLATION);
        result->offerTextViewerFallback = true;
        result->statusMessage = normalizeResult == ViewerWebSecurity::NormalizeTextResult::TooLarge
                                    ? FormatStringResource(g_hInstance,
                                                           IDS_VIEWERWEB_ERROR_FILE_TOO_LARGE_FMT,
                                                           FormatBytesCompact(result->loadedSourceBytes),
                                                           FormatBytesCompact(maxBytes))
                                    : LoadStringResource(g_hInstance,
                                                         kind == ViewerWebKind::Json ? IDS_VIEWERWEB_ERROR_PARSE_JSON_FAILED
                                                                                     : IDS_VIEWERWEB_ERROR_READ_FILE_FAILED);
        postBack(false);
        return;
    }

    result->documentRoute = ViewerWebSecurity::DocumentRoute::GeneratedPrivateOrigin;
    Debug::Perf::Emit(L"viewer.web.normalized_bytes", kind == ViewerWebKind::Json ? L"json" : L"markdown", 0u, textUtf8.size(), maxBytes, S_OK);

    if (kind == ViewerWebKind::Json)
    {
        std::vector<JsonLinesEntry> jsonLinesEntries;
        const bool forceJsonLines     = config.jsonViewMode == JsonViewMode::JsonLines;
        const bool pathLooksJsonLines = IsJsonLinesPath(result->path);
        const size_t publicationLimit = static_cast<size_t>(ViewerWebSecurity::GeneratedOutputLimit(maxBytes));
        JsonLinesUiStrings jsonLinesStrings{
            .badge        = Utf8FromWide(LoadStringResource(g_hInstance, IDS_VIEWERWEB_JSONL_BADGE)),
            .title        = Utf8FromWide(LoadStringResource(g_hInstance, IDS_VIEWERWEB_JSONL_TITLE)),
            .expandAll    = Utf8FromWide(LoadStringResource(g_hInstance, IDS_VIEWERWEB_JSONL_EXPAND_ALL)),
            .collapseAll  = Utf8FromWide(LoadStringResource(g_hInstance, IDS_VIEWERWEB_JSONL_COLLAPSE_ALL)),
            .genericValue = Utf8FromWide(LoadStringResource(g_hInstance, IDS_VIEWERWEB_JSONL_VALUE_GENERIC)),
        };
        const auto publishJsonLines = [&]() noexcept -> bool
        {
            std::string{}.swap(textUtf8);
            jsonLinesStrings.recordSummary = Utf8FromWide(FormatStringResource(
                g_hInstance,
                jsonLinesEntries.size() == 1u ? IDS_VIEWERWEB_JSONL_RECORD_ONE_FMT : IDS_VIEWERWEB_JSONL_RECORD_MANY_FMT,
                jsonLinesEntries.size()));
            if (! BuildJsonLinesHtml(jsonLinesEntries,
                                     jsonLinesStrings,
                                     GetHighlightJs(),
                                     themeObj,
                                     bg,
                                     fg,
                                     selBg,
                                     selFg,
                                     accent,
                                     theme.darkMode != FALSE,
                                     publicationLimit,
                                     result->utf8))
            {
                result->hr                      = HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
                result->generatedOutputRejected = true;
                result->generatedOutputBytes    = publicationLimit + 1u;
                result->generatedOutputLimit    = publicationLimit;
                result->statusMessage = FormatStringResource(g_hInstance,
                                                             IDS_VIEWERWEB_ERROR_FILE_TOO_LARGE_FMT,
                                                             FormatBytesCompact(result->generatedOutputBytes),
                                                             FormatBytesCompact(publicationLimit));
                postBack(false);
                return false;
            }
            result->jsonExpandCollapseAvailable = true;
            result->hr                          = S_OK;
            postBack(false);
            return true;
        };

        if (forceJsonLines)
        {
            if (! TryParseJsonLinesEntries(textUtf8, true, publicationLimit, jsonLinesEntries))
            {
                result->hr            = E_FAIL;
                result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_PARSE_JSON_FAILED);
                postBack(false);
                return;
            }

            static_cast<void>(publishJsonLines());
            return;
        }

        if (pathLooksJsonLines && TryParseJsonLinesEntries(textUtf8, true, publicationLimit, jsonLinesEntries))
        {
            static_cast<void>(publishJsonLines());
            return;
        }

        std::string jsonMutable(textUtf8);
        yyjson_read_err err{};
        unique_yyjson_doc doc(yyjson_read_opts(jsonMutable.data(), jsonMutable.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM, nullptr, &err));
        if (! doc)
        {
            if (TryParseJsonLinesEntries(textUtf8, false, publicationLimit, jsonLinesEntries))
            {
                std::string{}.swap(jsonMutable);
                static_cast<void>(publishJsonLines());
                return;
            }

            result->hr            = E_FAIL;
            result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_PARSE_JSON_FAILED);
            postBack(false);
            return;
        }

        size_t prettyLen = 0;
        unique_malloc_string pretty(yyjson_write_opts(doc.get(), YYJSON_WRITE_PRETTY | YYJSON_WRITE_ESCAPE_UNICODE, nullptr, &prettyLen, nullptr));
        if (! pretty)
        {
            result->hr            = E_OUTOFMEMORY;
            result->statusMessage = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_FORMAT_JSON_FAILED);
            postBack(false);
            return;
        }

        const std::string_view prettyJson(pretty.get(), prettyLen);
        const std::string escapedJson = EscapeJavaScriptStringUtf8(prettyJson);

        if (config.jsonViewMode == JsonViewMode::Pretty)
        {
            const std::string_view highlightJs = GetHighlightJs();

            const COLORREF codeBg             = BlendColorRefTruncate(bg, fg, theme.darkMode ? 20u : 10u);
            const COLORREF border             = BlendColorRefTruncate(bg, fg, theme.darkMode ? 35u : 45u);
            const COLORREF mutedFg            = BlendColorRefTruncate(bg, fg, 140u);
            const JsonTokenColors tokenColors = BuildJsonTokenColors(accent, fg, theme.darkMode != FALSE);
            const ScrollbarColors scrollbar   = BuildScrollbarColors(bg, fg, accent, theme.darkMode != FALSE);

            std::string html;
            html.reserve(highlightJs.size() + escapedJson.size() + 8192);
            html += kInternalHtmlHead;
            html += "<style>";
            html += ":root{--rs-bg:" + cssRgb(bg) + ";--rs-fg:" + cssRgb(fg) + ";--rs-sel-bg:" + cssRgb(selBg) + ";--rs-sel-fg:" + cssRgb(selFg) +
                    ";--rs-accent:" + cssRgb(accent) + ";--rs-code-bg:" + cssRgb(codeBg) + ";--rs-border:" + cssRgb(border) +
                    ";--rs-muted-fg:" + cssRgb(mutedFg) + ";--rs-key:" + cssRgb(tokenColors.key) + ";--rs-string:" + cssRgb(tokenColors.stringValue) +
                    ";--rs-number:" + cssRgb(tokenColors.numberValue) + ";--rs-literal:" + cssRgb(tokenColors.literalValue) +
                    ";--rs-scroll-track:" + cssRgb(scrollbar.track) + ";--rs-scroll-thumb:" + cssRgb(scrollbar.thumb) +
                    ";--rs-scroll-thumb-hover:" + cssRgb(scrollbar.thumbHover) + ";--rs-scroll-corner:" + cssRgb(scrollbar.corner) + ";}";
            html += "html,body{height:100%;margin:0;}body{background:var(--rs-bg);color:var(--rs-fg);font-family:\"Segoe UI Variable Text\",\"Segoe UI "
                    "Variable Small\",\"Segoe UI\",sans-serif;}";
            html += kCommonScrollbarCss;
            html += "::selection{background:var(--rs-sel-bg);color:var(--rs-sel-fg);}#app{height:100%;box-sizing:border-box;padding:12px;display:flex;}";
            html += "pre{flex:1;margin:0;background:var(--rs-code-bg);border:1px solid var(--rs-border);padding:12px;overflow:auto;border-radius:6px;}";
            html += "code{font-family:Consolas,ui-monospace,monospace;font-size:13px;line-height:1.45;}";
            html += ".hljs{background:transparent;}";
            html += ".hljs-attr{color:var(--rs-key);} .hljs-string{color:var(--rs-string);} .hljs-number{color:var(--rs-number);} "
                    ".hljs-literal{color:var(--rs-literal);}";
            html += ".hljs-punctuation,.hljs-brace{color:var(--rs-muted-fg);} .hljs-comment{opacity:0.8;}";
            html += "</style></head><body><div id=\"app\"><pre><code id=\"code\" class=\"language-json\"></code></pre></div>";
            html += "<script>";
            html.append(highlightJs);
            html += "</script><script>";
            html += "(() => {";
            html += "const initialTheme=" + themeObj + ";";
            html += kCommonThemeJs;
            html += "function applyTheme(t){const "
                    "r=document.documentElement.style;r.setProperty('--rs-bg',t.bg);r.setProperty('--rs-fg',t.fg);r.setProperty('--rs-sel-bg',t.selBg);r."
                    "setProperty('--rs-sel-fg',t.selFg);r.setProperty('--rs-accent',t.accent);const "
                    "bg=parseRgb(t.bg),fg=parseRgb(t.fg),acc=parseRgb(t.accent);const "
                    "dark=luma(bg)<128;r.setProperty('--rs-code-bg',rgb(blend(bg,fg,dark?20:10)));r.setProperty('--rs-border',rgb(blend(bg,fg,dark?35:45)));r."
                    "setProperty('--rs-muted-fg',rgb(blend(bg,fg,140)));setJsonTokenVars(r,bg,fg,acc);setScrollbarVars(r,bg,fg,acc);}";
            html += "const code=document.getElementById('code');";
            html += "code.textContent='" + escapedJson + "';";
            html += "window.RS={applyTheme:applyTheme};";
            html += "applyTheme(initialTheme);";
            html += "try{hljs.highlightElement(code);}catch(e){}";
            html += "})();";
            html += "</script></body></html>";

            result->utf8                        = std::move(html);
            result->jsonExpandCollapseAvailable = false;
            result->hr                          = S_OK;
            postBack(false);
            return;
        }

        const std::string_view jsonEditorJs = GetJsonEditorJs();
        const std::string& jsonEditorCss    = GetJsonEditorCssWithIcons();

        const COLORREF border             = BlendColorRefTruncate(bg, fg, theme.darkMode ? 45u : 80u);
        const COLORREF mutedFg            = BlendColorRefTruncate(bg, fg, 140u);
        const JsonTokenColors tokenColors = BuildJsonTokenColors(accent, fg, theme.darkMode != FALSE);
        const ScrollbarColors scrollbar   = BuildScrollbarColors(bg, fg, accent, theme.darkMode != FALSE);

        std::string html;
        html.reserve(jsonEditorJs.size() + jsonEditorCss.size() + escapedJson.size() + 8192);
        html += kInternalHtmlHead;
        html += "<style>";
        html += ":root{--rs-bg:" + cssRgb(bg) + ";--rs-fg:" + cssRgb(fg) + ";--rs-sel-bg:" + cssRgb(selBg) + ";--rs-sel-fg:" + cssRgb(selFg) +
                ";--rs-accent:" + cssRgb(accent) + ";--rs-border:" + cssRgb(border) + ";--rs-muted-fg:" + cssRgb(mutedFg) +
                ";--rs-scroll-track:" + cssRgb(scrollbar.track) + ";--rs-scroll-thumb:" + cssRgb(scrollbar.thumb) +
                ";--rs-scroll-thumb-hover:" + cssRgb(scrollbar.thumbHover) + ";--rs-scroll-corner:" + cssRgb(scrollbar.corner) +
                ";--rs-key:" + cssRgb(tokenColors.key) + ";--rs-string:" + cssRgb(tokenColors.stringValue) + ";--rs-number:" + cssRgb(tokenColors.numberValue) +
                ";--rs-literal:" + cssRgb(tokenColors.literalValue) + ";}";
        html += "html,body{height:100%;margin:0;}body{background:var(--rs-bg);color:var(--rs-fg);font-family:\"Segoe UI Variable Text\",\"Segoe UI Variable "
                "Small\",\"Segoe UI\",sans-serif;}#app{height:100%;}";
        html += kCommonScrollbarCss;
        html += jsonEditorCss;
        html += "html,body{background:var(--rs-bg)!important;color:var(--rs-fg)!important;}#app{height:100%!important;}";
        html += ".jsoneditor{border:none!important;height:100%!important;background:var(--rs-bg)!important;color:var(--rs-fg)!important;}";
        html += ".jsoneditor-frame{background:var(--rs-bg)!important;border:1px solid var(--rs-border)!important;}";
        html += ".jsoneditor-outer,.jsoneditor-inner,.jsoneditor-tree,.jsoneditor-tree-inner,.jsoneditor-text,.jsoneditor-text "
                "textarea{background:var(--rs-bg)!important;color:var(--rs-fg)!important;}";
        html += ".jsoneditor-field{color:var(--rs-key)!important;}";
        html += ".jsoneditor-value.jsoneditor-string{color:var(--rs-string)!important;}";
        html += ".jsoneditor-value.jsoneditor-number{color:var(--rs-number)!important;}";
        html += ".jsoneditor-value.jsoneditor-boolean,.jsoneditor-value.jsoneditor-null{color:var(--rs-literal)!important;}";
        html += ".jsoneditor-value.jsoneditor-object,.jsoneditor-value.jsoneditor-array{color:var(--rs-muted-fg)!important;}";
        html += ".jsoneditor-selected,.jsoneditor-highlight-active{background-color:var(--rs-sel-bg)!important;color:var(--rs-sel-fg)!important;}";
        html += ".jsoneditor-highlight{background-color:var(--rs-sel-bg)!important;}";
        html += ".jsoneditor .autocomplete.dropdown{background:var(--rs-bg)!important;border:1px solid var(--rs-border)!important;}";
        html += ".jsoneditor .autocomplete.dropdown .item{color:var(--rs-fg)!important;}";
        html += ".jsoneditor .autocomplete.dropdown .item.hover{background-color:var(--rs-sel-bg)!important;color:var(--rs-sel-fg)!important;}";
        html += ".jsoneditor-contextmenu .jsoneditor-menu{background:var(--rs-bg)!important;border:1px solid var(--rs-border)!important;}";
        html += ".jsoneditor-contextmenu .jsoneditor-menu button{color:var(--rs-fg)!important;}";
        html += ".jsoneditor-contextmenu .jsoneditor-menu button:hover{background-color:var(--rs-sel-bg)!important;color:var(--rs-sel-fg)!important;}";
        html += ".jsoneditor-contextmenu .jsoneditor-separator{border-top:1px solid var(--rs-border)!important;}";
        html += ".jsoneditor-contextmenu .jsoneditor-menu button.jsoneditor-expand{border-left:1px solid var(--rs-border)!important;}";
        html += "</style></head><body><div id=\"app\"></div>";
        html += "<script>";
        html.append(jsonEditorJs);
        html += "</script><script>";
        html += "(() => {";
        html += "const initialTheme=" + themeObj + ";";
        html += kCommonThemeJs;
        html +=
            "function applyTheme(t){const "
            "r=document.documentElement.style;r.setProperty('--rs-bg',t.bg);r.setProperty('--rs-fg',t.fg);r.setProperty('--rs-sel-bg',t.selBg);r.setProperty('-"
            "-rs-sel-fg',t.selFg);r.setProperty('--rs-accent',t.accent);const bg=parseRgb(t.bg),fg=parseRgb(t.fg),acc=parseRgb(t.accent);const "
            "dark=luma(bg)<128;r.setProperty('--rs-border',rgb(blend(bg,fg,dark?45:80)));r.setProperty('--rs-muted-fg',rgb(blend(bg,fg,140)));"
            "setJsonTokenVars(r,bg,fg,acc);setScrollbarVars(r,bg,fg,acc);}";
        html += "const jsonText='" + escapedJson + "';";
        html += "const container=document.getElementById('app');";
        html += "const options={mode:'tree',modes:['tree','view'],onEditable:()=>false,mainMenuBar:false,navigationBar:false,statusBar:false};";
        html += "const editor=new JSONEditor(container,options);";
        html += "window.RS={applyTheme:applyTheme,expandAll:()=>editor.expandAll(),collapseAll:()=>editor.collapseAll()};";
        html += "applyTheme(initialTheme);";
        html += "try{editor.set(JSON.parse(jsonText));}catch(e){editor.set({error:String(e)});}";
        html += "})();";
        html += "</script></body></html>";

        result->utf8                        = std::move(html);
        result->jsonExpandCollapseAvailable = true;
        result->hr                          = S_OK;
        postBack(false);
        return;
    }

    // Markdown
    const std::string escapedMarkdown   = EscapeJavaScriptStringUtf8(textUtf8);
    const std::string_view markdownItJs = GetMarkdownItJs();
    const std::string_view highlightJs  = GetHighlightJs();

    const COLORREF codeBg           = BlendColorRefTruncate(bg, fg, theme.darkMode ? 20u : 10u);
    const COLORREF border           = BlendColorRefTruncate(bg, fg, theme.darkMode ? 35u : 45u);
    const COLORREF mutedFg          = BlendColorRefTruncate(bg, fg, 140u);
    const COLORREF stringColor      = BlendColorRefTruncate(accent, fg, 60u);
    const COLORREF numberColor      = BlendColorRefTruncate(accent, fg, 90u);
    const ScrollbarColors scrollbar = BuildScrollbarColors(bg, fg, accent, theme.darkMode != FALSE);

    std::string html;
    html.reserve(markdownItJs.size() + highlightJs.size() + escapedMarkdown.size() + 8192);
    html += kInternalHtmlHead;
    html += "<style>";
    html += ":root{--rs-bg:" + cssRgb(bg) + ";--rs-fg:" + cssRgb(fg) + ";--rs-sel-bg:" + cssRgb(selBg) + ";--rs-sel-fg:" + cssRgb(selFg) +
            ";--rs-accent:" + cssRgb(accent) + ";--rs-code-bg:" + cssRgb(codeBg) + ";--rs-border:" + cssRgb(border) + ";--rs-muted-fg:" + cssRgb(mutedFg) +
            ";--rs-string:" + cssRgb(stringColor) + ";--rs-number:" + cssRgb(numberColor) + ";--rs-scroll-track:" + cssRgb(scrollbar.track) +
            ";--rs-scroll-thumb:" + cssRgb(scrollbar.thumb) + ";--rs-scroll-thumb-hover:" + cssRgb(scrollbar.thumbHover) +
            ";--rs-scroll-corner:" + cssRgb(scrollbar.corner) + ";}";
    html += "html,body{height:100%;margin:0;}body{background:var(--rs-bg);color:var(--rs-fg);font-family:\"Segoe UI Variable Text\",\"Segoe UI Variable "
            "Small\",\"Segoe UI\",sans-serif;}";
    html += kCommonScrollbarCss;
    html += "#app{max-width:100%;padding:16px;box-sizing:border-box;}a{color:var(--rs-accent);}";
    html += "pre{background:var(--rs-code-bg);border:1px solid var(--rs-border);padding:12px;overflow:auto;border-radius:6px;}";
    html += "code{font-family:Consolas,ui-monospace,monospace;}";
    html += "table{border-collapse:collapse;}th,td{border:1px solid var(--rs-border);padding:6px 10px;}";
    html += ".rs-source{white-space:pre;overflow:auto;font-family:Consolas,ui-monospace,monospace;}";
    html += ".hljs-comment{opacity:0.8;}.hljs-keyword,.hljs-selector-tag{color:var(--rs-accent);}";
    html += ".hljs-string{color:var(--rs-string);}.hljs-number{color:var(--rs-number);}.hljs-punctuation,.hljs-brace{color:var(--rs-muted-fg);}";
    html += "</style></head><body><div id=\"app\"></div>";
    html += "<script>";
    html.append(markdownItJs);
    html += "</script><script>";
    html.append(highlightJs);
    html += "</script><script>";
    html += "(() => {";
    html += "const initialTheme=" + themeObj + ";";
    html += kCommonThemeJs;
    html +=
        "function applyTheme(t){const "
        "r=document.documentElement.style;r.setProperty('--rs-bg',t.bg);r.setProperty('--rs-fg',t.fg);r.setProperty('--rs-sel-bg',t.selBg);r.setProperty('--rs-"
        "sel-fg',t.selFg);r.setProperty('--rs-accent',t.accent);const bg=parseRgb(t.bg),fg=parseRgb(t.fg),acc=parseRgb(t.accent);const "
        "dark=luma(bg)<128;r.setProperty('--rs-code-bg',rgb(blend(bg,fg,dark?20:10)));r.setProperty('--rs-border',rgb(blend(bg,fg,dark?35:45)));r.setProperty('"
        "--rs-muted-fg',rgb(blend(bg,fg,140)));r.setProperty('--rs-string',rgb(blend(acc,fg,60)));r.setProperty('--rs-number',rgb(blend(acc,fg,90)));"
        "setScrollbarVars(r,bg,fg,acc);}";
    html += "const src='" + escapedMarkdown + "';";
    html += "const container=document.getElementById('app');";
    html += "let showSource=" + std::string(markdownShowSource ? "true" : "false") + ";";
    html += "const md=window.markdownit({html:false,linkify:true,typographer:true});";
    html += "function "
            "render(){if(showSource){container.className='rs-source';container.textContent=src;return;}container.className='';container.innerHTML=md.render("
            "src);document.querySelectorAll('pre code').forEach((el)=>{try{hljs.highlightElement(el);}catch(e){}});}";
    html += "window.RS={applyTheme:applyTheme,setShowSource:(v)=>{showSource=!!v;render();}};";
    html += "applyTheme(initialTheme);render();";
    html += "})();";
    html += "</script></body></html>";

    result->utf8 = std::move(html);
    result->hr   = S_OK;
    postBack(false);
}

void ViewerWeb::AsyncSaveProc(AsyncSaveWorkItem* payload) noexcept
{
    std::unique_ptr<AsyncSaveWorkItem> work(payload);
    if (! work || ! work->viewer)
    {
        return;
    }

    ViewerWeb* self = work->viewer;
    auto releaseSelf = wil::scope_exit([&]() noexcept { self->Release(); });
    const auto isCurrent = [&]() noexcept
    { return self->_saveRequestId.load(std::memory_order_acquire) == work->requestId; };

    wil::com_ptr<IFileSystem> fileSystem;
    wil::com_ptr<IFileSystemIO> fileIo;
    wil::com_ptr<IFileReader> reader;
    wil::unique_hfile tempFile;
    std::wstring tempPath;
    bool committed = false;

    const auto closeOpenResources = [&]() noexcept
    {
        tempFile.reset();
        reader.reset();
        fileIo.reset();
        fileSystem.reset();
    };
    const auto cleanupUncommitted = [&]() noexcept
    {
        closeOpenResources();
        if (committed || tempPath.empty())
        {
            return;
        }
        if (DeleteFileW(tempPath.c_str()) != FALSE)
        {
            tempPath.clear();
            return;
        }

        const DWORD error = GetLastError();
        if (error == ERROR_FILE_NOT_FOUND || error == ERROR_PATH_NOT_FOUND)
        {
            tempPath.clear();
            return;
        }
        Debug::Warning(L"ViewerWeb: failed to remove an uncommitted Save As temp (error={}).", error);
    };
    auto cleanupAtExit = wil::scope_exit([&]() noexcept { cleanupUncommitted(); });

    const auto postResult = [&]() noexcept
    {
        // A visible completion is also the cleanup boundary: provider COM
        // references and any uncommitted sibling temp are gone before the UI
        // can observe that Save As has finished.
        cleanupUncommitted();
        work->fileSystemSnapshot.reset();
        if (! isCurrent())
        {
            return;
        }

        const HWND hwnd          = work->hwnd;
        const uint64_t requestId = work->requestId;
        const HRESULT resultHr   = work->hr;
        bool forcePostFailure    = false;
#ifdef ENABLE_TESTS
        forcePostFailure = g_failNextAsyncSaveCompletionPost.exchange(false, std::memory_order_acq_rel);
#endif
        const bool posted = hwnd && ! forcePostFailure && PostMessagePayload(hwnd, kAsyncSaveCompleteMessage, 0, std::move(work));
        if (! posted)
        {
            self->_asyncSavePostFailureHr.store(resultHr, std::memory_order_relaxed);
            self->_asyncSavePostFailureRequestId.store(requestId, std::memory_order_release);
            self->_asyncSavePostFailureCount.fetch_add(1u, std::memory_order_relaxed);
            if (hwnd)
            {
                static_cast<void>(PostMessageW(hwnd, kAsyncPostFailureMessage, kAsyncPostFailureSave, 0));
                static_cast<void>(InvalidateRect(hwnd, nullptr, FALSE));
            }
        }
    };

    const auto fail = [&](HRESULT hr, UINT messageId) noexcept
    {
        work->hr            = hr;
        work->statusMessage = LoadStringResource(g_hInstance, messageId);
        postResult();
    };

    if (! isCurrent())
    {
        return;
    }

    fileSystem = work->fileSystemSnapshot;
    if (! fileSystem)
    {
        fail(E_FAIL, IDS_VIEWERWEB_ERROR_SAVE_AS_NO_FILE_IO);
        return;
    }

    const HRESULT ioHr = fileSystem->QueryInterface(__uuidof(IFileSystemIO), fileIo.put_void());
    if (FAILED(ioHr) || ! fileIo)
    {
        fail(FAILED(ioHr) ? ioHr : HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED), IDS_VIEWERWEB_ERROR_SAVE_AS_NO_FILE_IO);
        return;
    }

    const HRESULT openHr = fileIo->CreateFileReader(work->sourcePath.c_str(), reader.put());
    if (FAILED(openHr) || ! reader)
    {
        fail(FAILED(openHr) ? openHr : E_FAIL, IDS_VIEWERWEB_ERROR_SAVE_AS_OPEN_FAILED);
        return;
    }

    uint64_t expectedBytes = 0u;
    const HRESULT sizeHr = reader->GetSize(&expectedBytes);
    if (FAILED(sizeHr))
    {
        fail(sizeHr, IDS_VIEWERWEB_ERROR_SAVE_AS_READ_FAILED);
        return;
    }
    if (expectedBytes > work->maxBytes)
    {
        work->hr = HRESULT_FROM_WIN32(ERROR_FILE_TOO_LARGE);
        work->statusMessage = FormatStringResource(g_hInstance,
                                                   IDS_VIEWERWEB_ERROR_FILE_TOO_LARGE_FMT,
                                                   FormatBytesCompact(expectedBytes),
                                                   FormatBytesCompact(work->maxBytes));
        postResult();
        return;
    }

    uint64_t position = 0u;
    const HRESULT seekHr = reader->Seek(0, FILE_BEGIN, &position);
    if (FAILED(seekHr) || position != 0u)
    {
        fail(FAILED(seekHr) ? seekHr : HRESULT_FROM_WIN32(ERROR_INVALID_DATA), IDS_VIEWERWEB_ERROR_SAVE_AS_SEEK_FAILED);
        return;
    }
    if (! isCurrent())
    {
        return;
    }

    const std::wstring destination = Common::Paths::ToExtendedWin32Path(work->destinationPath);
    const Common::Paths::UniqueSiblingFileOptions options{.prefix = L".rsw-save-", .suffix = L".tmp"};
    const HRESULT tempHr = Common::Paths::CreateUniqueSiblingFile(destination, options, tempPath, tempFile);
    if (FAILED(tempHr) || ! tempFile)
    {
        fail(FAILED(tempHr) ? tempHr : E_FAIL, IDS_VIEWERWEB_ERROR_SAVE_AS_WRITE_FAILED);
        return;
    }

    bool writeFailed = false;
    uint64_t copiedBytes = 0u;
    const HRESULT copyHr = ReadProviderExactly(
        reader.get(),
        expectedBytes,
        [&](std::span<const uint8_t> chunk) noexcept -> HRESULT
        {
            if (! isCurrent())
            {
                return HRESULT_FROM_WIN32(ERROR_CANCELLED);
            }
            if ((work->testFaultMask & ViewerWebSecurity::DebugSaveFaultWrite) != 0u)
            {
                writeFailed = true;
                return HRESULT_FROM_WIN32(ERROR_WRITE_FAULT);
            }
            const HRESULT writeHr = Common::HandleIo::WriteAll(
                tempFile.get(), std::as_bytes(std::span<const uint8_t>(chunk.data(), chunk.size())));
            writeFailed           = FAILED(writeHr);
            return writeHr;
        },
        copiedBytes,
        false);
    if (FAILED(copyHr))
    {
        if (copyHr == HRESULT_FROM_WIN32(ERROR_CANCELLED) || ! isCurrent())
        {
            return;
        }
        fail(copyHr, writeFailed ? IDS_VIEWERWEB_ERROR_SAVE_AS_WRITE_FAILED : IDS_VIEWERWEB_ERROR_SAVE_AS_READ_FAILED);
        return;
    }

    if (! isCurrent())
    {
        return;
    }
    if ((work->testFaultMask & ViewerWebSecurity::DebugSaveFaultFlush) != 0u || FlushFileBuffers(tempFile.get()) == FALSE)
    {
        const DWORD error = (work->testFaultMask & ViewerWebSecurity::DebugSaveFaultFlush) != 0u ? ERROR_WRITE_FAULT : GetLastError();
        fail(HRESULT_FROM_WIN32(error), IDS_VIEWERWEB_ERROR_SAVE_AS_WRITE_FAILED);
        return;
    }

    closeOpenResources();
    work->fileSystemSnapshot.reset();
    if (! isCurrent())
    {
        return;
    }

    if ((work->testFaultMask & ViewerWebSecurity::DebugSaveFaultCommit) != 0u ||
        MoveFileExW(tempPath.c_str(), destination.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH) == FALSE)
    {
        const DWORD error = (work->testFaultMask & ViewerWebSecurity::DebugSaveFaultCommit) != 0u ? ERROR_ACCESS_DENIED : GetLastError();
        fail(HRESULT_FROM_WIN32(error), IDS_VIEWERWEB_ERROR_SAVE_AS_WRITE_FAILED);
        return;
    }

    committed = true;
    work->hr   = S_OK;
    Debug::Perf::Emit(L"viewer.web.save_as_bytes", L"committed", 0u, copiedBytes, work->maxBytes, S_OK);
    postResult();
}

HRESULT ViewerWeb::CommandSaveAs(HWND hwnd) noexcept
{
    if (_currentPath.empty() || ! _fileSystem)
    {
        return S_FALSE;
    }

    const std::wstring suggested = LeafNameFromPath(_currentPath);
    const auto dest              = ShowSaveAsDialog(hwnd, suggested);
    if (! dest.has_value())
    {
        return S_FALSE;
    }

    return StartAsyncSave(hwnd, dest.value());
}

HRESULT ViewerWeb::StartAsyncSave(HWND hwnd, const std::filesystem::path& destination, uint32_t testFaultMask) noexcept
{
    constexpr uint32_t kSupportedTestFaults = ViewerWebSecurity::DebugSaveFaultWrite | ViewerWebSecurity::DebugSaveFaultFlush |
                                               ViewerWebSecurity::DebugSaveFaultCommit;
    if (! hwnd || ! _hWnd || hwnd != _hWnd.get() || _currentPath.empty() || ! _fileSystem || destination.empty() ||
        (testFaultMask & ~kSupportedTestFaults) != 0u)
    {
        return E_INVALIDARG;
    }

    const auto failSubmission = [&](HRESULT failureHr) noexcept -> HRESULT
    {
        _saveInProgress = false;
        _statusMessage  = LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_SAVE_AS_FAILED);
        Debug::Error(L"ViewerWeb: asynchronous Save As submission failed (hr=0x{:08X}).", static_cast<unsigned long>(failureHr));
        ShowHostAlert(hwnd, HOST_ALERT_ERROR, LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_SAVE_AS_FAILED));
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
        }
        return failureHr;
    };

    uint64_t requestId = _saveRequestId.fetch_add(1u, std::memory_order_acq_rel) + 1u;
    if (requestId == 0u)
    {
        requestId = _saveRequestId.fetch_add(1u, std::memory_order_acq_rel) + 1u;
    }

    auto work = std::unique_ptr<AsyncSaveWorkItem>(new (std::nothrow) AsyncSaveWorkItem{});
    if (! work)
    {
        return failSubmission(E_OUTOFMEMORY);
    }
    work->viewer             = this;
    work->hwnd               = hwnd;
    work->requestId          = requestId;
    work->maxBytes = static_cast<uint64_t>(std::min(_config.maxDocumentMiB, ViewerWebSecurity::kMaximumDocumentMiB)) * 1024ull * 1024ull;
    work->testFaultMask      = testFaultMask;
    work->sourcePath         = _currentPath;
    work->destinationPath    = destination.wstring();
    work->fileSystemSnapshot = _fileSystem;
    work->moduleKeepAlive    = AcquireModuleReferenceFromAddress(&kViewerWebModuleAnchor);
    if (! work->moduleKeepAlive)
    {
        return failSubmission(E_FAIL);
    }

    _saveInProgress = true;
    _statusMessage.clear();
    if (_hWnd)
    {
        InvalidateRect(_hWnd.get(), &_headerRect, FALSE);
    }
    AddRef();
    g_activeAsyncWorkerCount.fetch_add(1u, std::memory_order_acq_rel);
    const BOOL queued = TrySubmitThreadpoolCallback(
        [](PTP_CALLBACK_INSTANCE instance, void* context) noexcept
        {
            std::unique_ptr<AsyncSaveWorkItem> work(static_cast<AsyncSaveWorkItem*>(context));
            auto releaseWorkerGate = wil::scope_exit(
                []() noexcept { g_activeAsyncWorkerCount.fetch_sub(1u, std::memory_order_acq_rel); });
            if (work && work->moduleKeepAlive)
            {
                TransferModulePinToCallbackReturn(instance, work->moduleKeepAlive);
            }
            if (! work)
            {
                return;
            }
            AsyncSaveProc(work.release());
        },
        work.get(),
        nullptr);
    if (queued == FALSE)
    {
        g_activeAsyncWorkerCount.fetch_sub(1u, std::memory_order_acq_rel);
        Release();
        return failSubmission(E_FAIL);
    }

    work.release();
    return S_OK;
}

void ViewerWeb::CommandFind(HWND hwnd) noexcept
{
    if (_hFindDialog && IsWindow(_hFindDialog.get()))
    {
        ShowWindow(_hFindDialog.get(), SW_SHOWNORMAL);
        static_cast<void>(SetForegroundWindow(_hFindDialog.get()));
        return;
    }

    if (! _findQuery.empty())
    {
        wcsncpy_s(_findBuffer.data(), _findBuffer.size(), _findQuery.c_str(), _TRUNCATE);
    }
    else
    {
        _findBuffer[0] = L'\0';
    }

    _findReplace               = {};
    _findReplace.lStructSize   = sizeof(_findReplace);
    _findReplace.hwndOwner     = hwnd;
    _findReplace.lpstrFindWhat = _findBuffer.data();
    _findReplace.wFindWhatLen  = static_cast<WORD>(_findBuffer.size());
    _findReplace.Flags         = FR_DOWN;

    HWND dlg = FindTextW(&_findReplace);
    if (! dlg)
    {
        return;
    }

    _hFindDialog.reset(dlg);
}

void ViewerWeb::CommandFindNext(HWND hwnd) noexcept
{
    if (_findQuery.empty())
    {
        CommandFind(hwnd);
        return;
    }

    if (! _webView)
    {
        return;
    }

    const std::wstring queryEsc = EscapeJavaScriptString(_findQuery);
    const std::wstring script =
        std::format(L"(function(){{try{{return window.find('{}',false,false,true,false,true,false);}}catch(e){{return false;}}}})();", queryEsc);
    static_cast<void>(_webView->ExecuteScript(script.c_str(), nullptr));
}

void ViewerWeb::CommandFindPrevious(HWND hwnd) noexcept
{
    if (_findQuery.empty())
    {
        CommandFind(hwnd);
        return;
    }

    if (! _webView)
    {
        return;
    }

    const std::wstring queryEsc = EscapeJavaScriptString(_findQuery);
    const std::wstring script =
        std::format(L"(function(){{try{{return window.find('{}',false,true,true,false,true,false);}}catch(e){{return false;}}}})();", queryEsc);
    static_cast<void>(_webView->ExecuteScript(script.c_str(), nullptr));
}

void ViewerWeb::CommandCopyUrl(HWND hwnd) noexcept
{
    std::wstring toCopy;

    if (_webView)
    {
        wil::unique_cotaskmem_string source;
        if (SUCCEEDED(_webView->get_Source(source.put())) && source && source.get()[0] != L'\0')
        {
            const std::wstring_view src(source.get());
            if (! OrdinalString::StartsWithNoCase(src, L"about:") && ! IsInternalDocumentUrl(src))
            {
                toCopy.assign(src);
            }
        }
    }

    if (toCopy.empty())
    {
        if (_kind == ViewerWebKind::Web)
        {
            if (_tempExtractedPath.has_value() && ! _tempExtractedPath.value().empty())
            {
                toCopy = UrlFromFilePath(_tempExtractedPath.value().wstring());
            }
        }
    }

    if (toCopy.empty())
    {
        toCopy = _currentPath;
    }

    static_cast<void>(Common::Clipboard::TrySetUnicodeText(hwnd, toCopy, Common::Clipboard::EmptyUnicodeTextPolicy::Reject));
}

void ViewerWeb::CommandOpenExternal(HWND hwnd) noexcept
{
    if (! _config.allowExternalNavigation)
    {
        return;
    }

    std::wstring url;

    if (_webView)
    {
        wil::unique_cotaskmem_string source;
        if (SUCCEEDED(_webView->get_Source(source.put())) && source && source.get()[0] != L'\0')
        {
            const std::wstring_view src(source.get());
            if (ViewerWebSecurity::IsHttpOrHttps(src))
            {
                url.assign(src);
            }
        }
    }

    if (url.empty())
    {
        return;
    }

    const HINSTANCE res = ShellExecuteW(hwnd, L"open", url.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
    if (reinterpret_cast<INT_PTR>(res) <= 32)
    {
        ShowHostAlert(hwnd, HOST_ALERT_ERROR, LoadStringResource(g_hInstance, IDS_VIEWERWEB_ERROR_OPEN_BROWSER_FAILED));
    }
}

void ViewerWeb::CommandZoom(double factor) noexcept
{
    if (! _webViewController || factor <= 0.0)
    {
        return;
    }

    double current = 1.0;
    static_cast<void>(_webViewController->get_ZoomFactor(&current));

    const double next = std::clamp(current * factor, 0.25, 5.0);
    static_cast<void>(_webViewController->put_ZoomFactor(next));
}

void ViewerWeb::CommandZoomIn() noexcept
{
    CommandZoom(1.10);
}

void ViewerWeb::CommandZoomOut() noexcept
{
    CommandZoom(1.0 / 1.10);
}

void ViewerWeb::CommandZoomReset() noexcept
{
    if (_webViewController)
    {
        static_cast<void>(_webViewController->put_ZoomFactor(1.0));
    }
}

void ViewerWeb::CommandToggleDevTools() noexcept
{
    if (! _documentScriptsEnabled || ! _config.devToolsEnabled)
    {
        ShowHostAlert(_hWnd.get(), HOST_ALERT_WARNING, LoadStringResource(g_hInstance, IDS_VIEWERWEB_WARNING_DEVTOOLS_DISABLED));
        return;
    }

    if (_webView)
    {
        static_cast<void>(_webView->OpenDevToolsWindow());
    }
}

void ViewerWeb::CommandJsonExpandAll() noexcept
{
    if (_kind != ViewerWebKind::Json || ! _jsonExpandCollapseAvailable || ! _webView)
    {
        return;
    }

    static_cast<void>(_webView->ExecuteScript(L"(function(){try{if(window.RS&&window.RS.expandAll){window.RS.expandAll();}}catch(e){}})();", nullptr));
}

void ViewerWeb::CommandJsonCollapseAll() noexcept
{
    if (_kind != ViewerWebKind::Json || ! _jsonExpandCollapseAvailable || ! _webView)
    {
        return;
    }

    static_cast<void>(_webView->ExecuteScript(L"(function(){try{if(window.RS&&window.RS.collapseAll){window.RS.collapseAll();}}catch(e){}})();", nullptr));
}

void ViewerWeb::CommandMarkdownToggleSource() noexcept
{
    if (_kind != ViewerWebKind::Markdown)
    {
        return;
    }

    _markdownShowSource = ! _markdownShowSource;

    if (_hWnd)
    {
        HMENU menu = _menuHandle ? _menuHandle.get() : GetMenu(_hWnd.get());
        if (menu)
        {
            CheckMenuItem(
                menu, IDM_VIEWERWEB_TOOLS_MARKDOWN_TOGGLE_SOURCE, static_cast<UINT>(MF_BYCOMMAND | (_markdownShowSource ? MF_CHECKED : MF_UNCHECKED)));
            if (_menuBarHost.GetHwnd())
            {
                _menuBarHost.SyncMenuModel();
            }
            else
            {
                DrawMenuBar(_hWnd.get());
            }
        }
    }

    if (_webView)
    {
        const std::wstring script = std::format(L"(function(){{try{{if(window.RS&&window.RS.setShowSource){{window.RS.setShowSource({});}}}}catch(e){{}}}})();",
                                                _markdownShowSource ? L"true" : L"false");
        static_cast<void>(_webView->ExecuteScript(script.c_str(), nullptr));
    }
}

namespace
{
std::optional<std::filesystem::path> ShowSaveAsDialog(HWND hwnd, std::wstring_view suggestedFileName) noexcept
{
    wil::com_ptr<IFileSaveDialog> dialog;
    const HRESULT hr = CoCreateInstance(CLSID_FileSaveDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(dialog.put()));
    if (FAILED(hr) || ! dialog)
    {
        return std::nullopt;
    }

    DWORD options = 0;
    static_cast<void>(dialog->GetOptions(&options));
    options |= FOS_FORCEFILESYSTEM | FOS_OVERWRITEPROMPT | FOS_PATHMUSTEXIST;
    static_cast<void>(dialog->SetOptions(options));

    if (! suggestedFileName.empty())
    {
        static_cast<void>(dialog->SetFileName(std::wstring(suggestedFileName).c_str()));
    }

    const HRESULT showHr = dialog->Show(hwnd);
    if (showHr == HRESULT_FROM_WIN32(ERROR_CANCELLED))
    {
        return std::nullopt;
    }
    if (FAILED(showHr))
    {
        return std::nullopt;
    }

    wil::com_ptr<IShellItem> item;
    const HRESULT itemHr = dialog->GetResult(item.put());
    if (FAILED(itemHr) || ! item)
    {
        return std::nullopt;
    }

    wil::unique_cotaskmem_string path;
    const HRESULT nameHr = item->GetDisplayName(SIGDN_FILESYSPATH, path.put());
    if (FAILED(nameHr) || ! path)
    {
        return std::nullopt;
    }

    return std::filesystem::path(path.get());
}
} // namespace

void ResetSharedEnvironment() noexcept
{
    g_shutdownCleanupComplete.store(false, std::memory_order_release);
    ResetSharedEnvironmentImpl();
    g_shutdownCleanupComplete.store(ResetViewerWebClassBackgroundBrushAtQuietPoint(), std::memory_order_release);
}

[[nodiscard]] bool CanUnloadViewerWebModuleNow() noexcept
{
    RetryStagedFileCleanup();
    const DWORD callbackOwnerThread = g_comCallbackOwnerThreadId.load(std::memory_order_acquire);
    if (! g_shutdownCleanupComplete.load(std::memory_order_acquire) || g_activeAsyncWorkerCount.load(std::memory_order_acquire) != 0u ||
        g_liveComCallbackCount.load(std::memory_order_acquire) != 0u || g_stagedCleanupTracker.PendingCount() != 0u ||
        g_comCallbackThreadViolation.load(std::memory_order_acquire) ||
        (callbackOwnerThread != 0u && callbackOwnerThread != GetCurrentThreadId()))
    {
        return false;
    }

    const uint64_t releaseEpoch = g_comCallbackReleaseEpoch.load(std::memory_order_acquire);
    const uint64_t observedEpoch = g_observedComCallbackReleaseEpoch.load(std::memory_order_acquire);
    if (releaseEpoch != observedEpoch)
    {
        g_observedComCallbackReleaseEpoch.store(releaseEpoch, std::memory_order_release);
        return false;
    }
    return true;
}
