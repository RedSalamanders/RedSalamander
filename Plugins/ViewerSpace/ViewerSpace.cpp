#include "ViewerSpace.h"

#include "LocalizationManager.h"

#include <algorithm>
#include <array>
#include <chrono>
#include <cmath>
#include <condition_variable>
#include <cstring>
#include <cwchar>
#include <exception>
#include <filesystem>
#include <format>
#include <limits>
#include <memory>
#include <mutex>
#include <numeric>
#include <psapi.h>
#include <shellapi.h>
#include <string_view>
#include <system_error>

#pragma warning(push)
// (C6297) Arithmetic overflow. Results might not be an expected value.
// (C28182) Dereferencing NULL pointer.
#pragma warning(disable : 6297 28182)
#include <yyjson.h>
#pragma warning(pop)

#include "DxUi/DxUi.Typography.h"
#include "DxUi/DxUi.h"
#include "FluentIcons.h"
#include "Helpers.h"
#include "WindowMessages.h"
#include "WindowSizing.h"
#include "resource.h"

#pragma comment(lib, "d3d11.lib")
#pragma comment(lib, "dxgi.lib")

extern HINSTANCE g_hInstance;

namespace Typography = RedSalamander::DxUi::Typography;

namespace
{
constexpr UINT_PTR kTimerAnimationId = 1;
constexpr UINT kAnimationIntervalMs  = 16;

constexpr int kHostFolderViewContextMenuResourceId = 138;

constexpr UINT kCmdTreemapContextFocusInPane = 0xC100u;
constexpr UINT kCmdTreemapContextZoomIn      = 0xC101u;
constexpr UINT kCmdTreemapContextZoomOut     = 0xC102u;

constexpr UINT kCmdFolderViewContextOpen       = 33280u;
constexpr UINT kCmdFolderViewContextOpenWith   = 33281u;
constexpr UINT kCmdFolderViewContextDelete     = 33282u;
constexpr UINT kCmdFolderViewContextRename     = 33283u;
constexpr UINT kCmdFolderViewContextCopy       = 33284u;
constexpr UINT kCmdFolderViewContextPaste      = 33285u;
constexpr UINT kCmdFolderViewContextProperties = 33286u;
constexpr UINT kCmdFolderViewContextMove       = 33287u;
constexpr UINT kCmdFolderViewContextViewSpace  = 33288u;

constexpr UINT kFolderViewDebugCommandIdBase = 60000u;

constexpr float kHeaderHeightDip      = 72.0f;
constexpr float kHeaderButtonWidthDip = 52.0f;
constexpr float kPaddingDip           = 8.0f;
constexpr float kItemGapDip           = 1.0f;
constexpr float kMinHitAreaDip2       = 4.0f * 4.0f;
constexpr float kTooltipMaxWidthDip   = 420.0f;
constexpr float kTooltipMaxHeightDip  = 320.0f;

constexpr double kAnimationDurationSeconds = 0.18;
constexpr double kScanCompletedOverlaySeconds = 1.35;

constexpr size_t kDefaultLayoutItems = 4096;
constexpr size_t kMaxLayoutItems = 50000;
constexpr size_t kMaxDrawItems = 50000;
constexpr size_t kMaxDrawItemsWhileScanning = 20000;
constexpr float kMinVisibleTileAreaDip2 = 64.0f;
constexpr uint32_t kAdaptiveFileCandidateCeiling = 20000u;
constexpr uint64_t kMaxScanCacheSnapshotBytes = 64ull * 1024ull * 1024ull;
#ifdef ENABLE_TESTS
constexpr uint32_t kRendererFailureD3DDevice = 1u;
constexpr uint32_t kRendererFailureD2DContext = 2u;
constexpr uint32_t kRendererFailureSwapChain = 3u;
constexpr uint32_t kRendererFailureResizeBuffers = 4u;
constexpr uint32_t kRendererFailureGetBuffer = 5u;
constexpr uint32_t kRendererFailureTargetBitmap = 6u;
constexpr uint32_t kRendererFailureRenderTarget = 7u;
constexpr uint32_t kRendererDiscardDestroy = 20u;
constexpr uint32_t kRendererDiscardDpiChanged = 21u;
constexpr uint32_t kRendererDiscardStaticDeviceLost = 22u;
constexpr uint32_t kRendererDiscardDeviceLost = 23u;
constexpr uint32_t kRendererDiscardTheme = 25u;
constexpr uint32_t kRendererDiscardUnknown = 26u;
#endif

[[nodiscard]] uint64_t MaxScanCacheSnapshotBytes() noexcept
{
#ifdef ENABLE_TESTS
    std::array<wchar_t, 32> value{};
    const DWORD chars = GetEnvironmentVariableW(L"REDSALAMANDER_VIEWERSPACE_TEST_CACHE_CAP_BYTES", value.data(), static_cast<DWORD>(value.size()));
    if (chars > 0 && chars < value.size())
    {
        wchar_t* end = nullptr;
        const unsigned long long parsed = wcstoull(value.data(), &end, 10);
        if (parsed > 0 && end != value.data())
        {
            return std::min<uint64_t>(static_cast<uint64_t>(parsed), kMaxScanCacheSnapshotBytes);
        }
    }
#endif
    return kMaxScanCacheSnapshotBytes;
}

[[nodiscard]] uint64_t SampleCurrentWorkingSetBytes() noexcept
{
    PROCESS_MEMORY_COUNTERS_EX counters{};
    counters.cb = static_cast<DWORD>(sizeof(counters));
    if (K32GetProcessMemoryInfo(GetCurrentProcess(), reinterpret_cast<PROCESS_MEMORY_COUNTERS*>(&counters), counters.cb) == 0)
    {
        return 0u;
    }

    return static_cast<uint64_t>(counters.WorkingSetSize);
}

[[nodiscard]] uint32_t DefaultWin32ScanThreads() noexcept
{
    const uint32_t hardware = std::max<uint32_t>(1u, std::thread::hardware_concurrency());
    return std::clamp<uint32_t>(hardware, 1u, 8u);
}

[[nodiscard]] uint32_t PixelDrivenLayoutItemLimit(float projectedAreaDip2) noexcept
{
    if (projectedAreaDip2 <= 0.0f || ! std::isfinite(projectedAreaDip2))
    {
        return static_cast<uint32_t>(kDefaultLayoutItems);
    }

    const auto pixelBudget = static_cast<uint32_t>(std::ceil(projectedAreaDip2 / kMinVisibleTileAreaDip2));
    const auto conservativeBudget = std::max<uint32_t>(static_cast<uint32_t>(kDefaultLayoutItems), pixelBudget);
    return std::clamp<uint32_t>(conservativeBudget, 32u, static_cast<uint32_t>(kMaxLayoutItems));
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

[[nodiscard]] std::wstring StripMenuMnemonicMarkers(std::wstring_view text)
{
    std::wstring result;
    result.reserve(text.size());

    for (size_t index = 0; index < text.size(); ++index)
    {
        if (text[index] != L'&')
        {
            result.push_back(text[index]);
            continue;
        }

        if ((index + 1u) < text.size() && text[index + 1u] == L'&')
        {
            result.push_back(L'&');
            ++index;
        }
    }

    return result;
}

[[nodiscard]] std::wstring GetViewerSpaceContextMenuIconGlyph(UINT commandId) noexcept
{
    switch (commandId)
    {
        case kCmdFolderViewContextOpen: return std::wstring(1u, FluentIcons::kOpenFile);
        case kCmdFolderViewContextCopy: return std::wstring(1u, FluentIcons::kCopy);
        case kCmdFolderViewContextPaste: return std::wstring(1u, FluentIcons::kPaste);
        case kCmdFolderViewContextDelete: return std::wstring(1u, FluentIcons::kDelete);
        case kCmdFolderViewContextRename: return std::wstring(1u, FluentIcons::kRename);
        case kCmdFolderViewContextProperties: return std::wstring(1u, FluentIcons::kInfo);
        default: return {};
    }
}

[[nodiscard]] std::vector<RedSalamander::DxUi::MenuFlyoutItem> ConvertHMenuToDxFlyoutItems(HMENU menu) noexcept
{
    using RedSalamander::DxUi::MenuFlyoutItem;
    using RedSalamander::DxUi::MenuItemKind;

    std::vector<MenuFlyoutItem> items;
    if (! menu)
    {
        return items;
    }

    const int itemCount = GetMenuItemCount(menu);
    if (itemCount <= 0)
    {
        return items;
    }

    items.reserve(static_cast<size_t>(itemCount));
    for (UINT position = 0; position < static_cast<UINT>(itemCount); ++position)
    {
        MENUITEMINFOW itemInfo{};
        itemInfo.cbSize = sizeof(itemInfo);
        itemInfo.fMask  = MIIM_FTYPE | MIIM_ID | MIIM_STATE | MIIM_SUBMENU;
        if (! GetMenuItemInfoW(menu, position, TRUE, &itemInfo))
        {
            continue;
        }

        MenuFlyoutItem item{};
        if ((itemInfo.fType & MFT_SEPARATOR) != 0)
        {
            item.kind = MenuItemKind::Separator;
            items.push_back(std::move(item));
            continue;
        }

        std::array<wchar_t, 512> buffer{};
        const int textLength = GetMenuStringW(menu, position, buffer.data(), static_cast<int>(buffer.size()), MF_BYPOSITION);
        if (textLength > 0)
        {
            std::wstring_view fullText(buffer.data(), static_cast<size_t>(textLength));
            const size_t tabPos               = fullText.find(L'\t');
            const std::wstring_view labelText = tabPos == std::wstring_view::npos ? fullText : fullText.substr(0, tabPos);
            item.text                         = StripMenuMnemonicMarkers(labelText);
            if (tabPos != std::wstring_view::npos && (tabPos + 1u) < fullText.size())
            {
                item.acceleratorText.assign(fullText.substr(tabPos + 1u));
            }
        }

        item.commandId = static_cast<int>(itemInfo.wID);
        item.enabled   = (itemInfo.fState & MFS_GRAYED) == 0;
        item.checked   = (itemInfo.fState & MFS_CHECKED) != 0;

        if ((itemInfo.fType & MFT_RADIOCHECK) != 0)
        {
            item.kind = MenuItemKind::Radio;
        }
        else if (item.checked)
        {
            item.kind = MenuItemKind::Toggle;
        }

        item.iconGlyph = GetViewerSpaceContextMenuIconGlyph(itemInfo.wID);

        if (itemInfo.hSubMenu)
        {
            item.children = ConvertHMenuToDxFlyoutItems(itemInfo.hSubMenu);
        }

        items.push_back(std::move(item));
    }

    return items;
}

struct ViewerSpaceClassBackgroundBrushState
{
    ViewerSpaceClassBackgroundBrushState()                                                       = default;
    ViewerSpaceClassBackgroundBrushState(const ViewerSpaceClassBackgroundBrushState&)            = delete;
    ViewerSpaceClassBackgroundBrushState& operator=(const ViewerSpaceClassBackgroundBrushState&) = delete;
    ViewerSpaceClassBackgroundBrushState(ViewerSpaceClassBackgroundBrushState&&)                 = delete;
    ViewerSpaceClassBackgroundBrushState& operator=(ViewerSpaceClassBackgroundBrushState&&)      = delete;

    wil::unique_hbrush activeBrush;
    COLORREF activeColor = CLR_INVALID;

    wil::unique_hbrush pendingBrush;
    COLORREF pendingColor = CLR_INVALID;

    bool classRegistered = false;
};

ViewerSpaceClassBackgroundBrushState* g_viewerSpaceClassBackgroundBrush = []() noexcept
{
    auto* state = new (std::nothrow) ViewerSpaceClassBackgroundBrushState();
    if (state == nullptr)
    {
        std::terminate();
    }
    return state;
}();
std::atomic_uint32_t g_viewerSpaceWindowCount{0};
ATOM g_viewerSpaceClassAtom = 0;

void DestroyViewerSpaceClassBackgroundBrushState() noexcept
{
    if (g_viewerSpaceClassBackgroundBrush == nullptr)
    {
        return;
    }

    g_viewerSpaceClassBackgroundBrush->classRegistered = false;
    g_viewerSpaceClassBackgroundBrush->pendingBrush.reset();
    g_viewerSpaceClassBackgroundBrush->pendingColor = CLR_INVALID;
    g_viewerSpaceClassBackgroundBrush->activeBrush.reset();
    g_viewerSpaceClassBackgroundBrush->activeColor = CLR_INVALID;
    delete g_viewerSpaceClassBackgroundBrush;
    g_viewerSpaceClassBackgroundBrush = nullptr;
}

HBRUSH GetActiveViewerSpaceClassBackgroundBrush() noexcept
{
    if (g_viewerSpaceClassBackgroundBrush->pendingBrush)
    {
        return g_viewerSpaceClassBackgroundBrush->pendingBrush.get();
    }

    if (! g_viewerSpaceClassBackgroundBrush->activeBrush)
    {
        const COLORREF fallback                        = GetSysColor(COLOR_WINDOW);
        g_viewerSpaceClassBackgroundBrush->activeColor = fallback;
        g_viewerSpaceClassBackgroundBrush->activeBrush.reset(CreateSolidBrush(fallback));
    }

    return g_viewerSpaceClassBackgroundBrush->activeBrush.get();
}

void RequestViewerSpaceClassBackgroundColor(COLORREF color) noexcept
{
    if (color == CLR_INVALID)
    {
        return;
    }

    if (g_viewerSpaceClassBackgroundBrush->pendingBrush && g_viewerSpaceClassBackgroundBrush->pendingColor == color)
    {
        return;
    }

    wil::unique_hbrush brush(CreateSolidBrush(color));
    if (! brush)
    {
        return;
    }

    g_viewerSpaceClassBackgroundBrush->pendingColor = color;
    g_viewerSpaceClassBackgroundBrush->pendingBrush = std::move(brush);
}

void ApplyPendingViewerSpaceClassBackgroundBrush(HWND hwnd) noexcept
{
    if (! hwnd || ! g_viewerSpaceClassBackgroundBrush->pendingBrush || ! g_viewerSpaceClassBackgroundBrush->classRegistered)
    {
        return;
    }

    const LONG_PTR newBrush = reinterpret_cast<LONG_PTR>(g_viewerSpaceClassBackgroundBrush->pendingBrush.get());
    SetClassLongPtrW(hwnd, GCLP_HBRBACKGROUND, newBrush);

    g_viewerSpaceClassBackgroundBrush->activeBrush  = std::move(g_viewerSpaceClassBackgroundBrush->pendingBrush);
    g_viewerSpaceClassBackgroundBrush->activeColor  = g_viewerSpaceClassBackgroundBrush->pendingColor;
    g_viewerSpaceClassBackgroundBrush->pendingColor = CLR_INVALID;
}

constexpr char kViewerSpaceSchemaJson[] = R"json({
    "version": 1,
    "title": "Space Viewer",
	    "fields": [
	        {
	            "key": "topFilesPerDirectory",
	            "type": "value",
	            "label": "Top files per directory",
	            "description": "Compatibility floor for largest files retained per directory; the renderer may retain more candidates when pixel/detail budgets allow it.",
	            "default": 96,
	            "min": 0,
	            "max": 20000
	        },
	        {
	            "key": "scanThreads",
	            "type": "value",
	            "label": "Scan threads",
	            "description": "Number of background threads used to scan subfolders in parallel.",
	            "default": 1,
	            "min": 1,
	            "max": 16
	        },
	        {
	            "key": "maxConcurrentScansPerVolume",
	            "type": "value",
	            "label": "Max concurrent scans per volume",
	            "description": "Limits how many Space viewers scan the same drive at once (reduces disk thrash when opening multiple viewers).",
            "default": 1,
            "min": 1,
            "max": 8
        },
        {
            "key": "cacheEnabled",
            "type": "option",
            "label": "Scan cache",
            "description": "Reuse recent scan results when opening another Space viewer on the same root.",
            "default": "1",
            "options": [
                { "value": "0", "label": "Off" },
                { "value": "1", "label": "On" }
            ]
        },
        {
            "key": "cacheTtlSeconds",
            "type": "value",
            "label": "Cache TTL (seconds)",
            "description": "How long a scan result remains reusable.",
            "default": 60,
            "min": 0,
            "max": 3600
        },
        {
            "key": "cacheMaxEntries",
            "type": "value",
            "label": "Cache max entries",
            "description": "Maximum number of cached roots kept in memory.",
            "default": 1,
            "min": 0,
            "max": 16
        }
    ]
})json";

[[nodiscard]] const char* GetViewerSpaceStaticConfigurationSchemaImpl() noexcept
{
    return kViewerSpaceSchemaJson;
}

std::atomic_uint32_t g_maxConcurrentScansPerVolume{1};
std::atomic_bool g_cacheEnabled{true};
std::atomic_uint32_t g_cacheTtlSeconds{60};
std::atomic_uint32_t g_cacheMaxEntries{1};

std::wstring_view CopyToArena(std::pmr::monotonic_buffer_resource& arena, std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return {};
    }

    const size_t count = text.size();
    const size_t bytes = (count + 1u) * sizeof(wchar_t);
    void* memory       = arena.allocate(bytes, alignof(wchar_t));
    auto* buffer       = static_cast<wchar_t*>(memory);
    std::copy_n(text.data(), count, buffer);
    buffer[count] = L'\0';
    return std::wstring_view(buffer, count);
}

uint32_t GetMaxConcurrentScansPerVolume() noexcept
{
    const uint32_t value = g_maxConcurrentScansPerVolume.load(std::memory_order_acquire);
    return std::clamp(value, 1u, 8u);
}

D2D1::ColorF ColorFFromArgb(uint32_t argb) noexcept
{
    const float a = static_cast<float>((argb >> 24) & 0xFFu) / 255.0f;
    const float r = static_cast<float>((argb >> 16) & 0xFFu) / 255.0f;
    const float g = static_cast<float>((argb >> 8) & 0xFFu) / 255.0f;
    const float b = static_cast<float>((argb) & 0xFFu) / 255.0f;
    return D2D1::ColorF(r, g, b, a);
}

COLORREF ColorRefFromArgb(uint32_t argb) noexcept
{
    const uint8_t r = static_cast<uint8_t>((argb >> 16) & 0xFFu);
    const uint8_t g = static_cast<uint8_t>((argb >> 8) & 0xFFu);
    const uint8_t b = static_cast<uint8_t>(argb & 0xFFu);
    return RGB(r, g, b);
}

COLORREF BlendColor(COLORREF under, COLORREF over, uint8_t alpha) noexcept
{
    const uint32_t inv = 255u - alpha;
    const uint32_t r   = (static_cast<uint32_t>(GetRValue(under)) * inv + static_cast<uint32_t>(GetRValue(over)) * alpha) / 255u;
    const uint32_t g   = (static_cast<uint32_t>(GetGValue(under)) * inv + static_cast<uint32_t>(GetGValue(over)) * alpha) / 255u;
    const uint32_t b   = (static_cast<uint32_t>(GetBValue(under)) * inv + static_cast<uint32_t>(GetBValue(over)) * alpha) / 255u;
    return RGB(static_cast<BYTE>(r), static_cast<BYTE>(g), static_cast<BYTE>(b));
}

COLORREF ChooseContrastingTextColor(COLORREF background) noexcept
{
    const float r   = static_cast<float>(GetRValue(background)) / 255.0f;
    const float g   = static_cast<float>(GetGValue(background)) / 255.0f;
    const float b   = static_cast<float>(GetBValue(background)) / 255.0f;
    const float lum = 0.2126f * r + 0.7152f * g + 0.0722f * b;
    return lum > 0.60f ? RGB(0, 0, 0) : RGB(255, 255, 255);
}

D2D1::ColorF ColorFFromHsv(double hue01, double saturation, double value, float alpha) noexcept;

bool IsAsciiAlpha(wchar_t ch) noexcept
{
    return (ch >= L'A' && ch <= L'Z') || (ch >= L'a' && ch <= L'z');
}

wchar_t DeterminePreferredPathSeparator(std::wstring_view path, bool fileSystemIsWin32) noexcept;

bool IsPathSeparator(wchar_t ch) noexcept
{
    return ch == L'\\' || ch == L'/';
}

bool LooksLikeWin32Path(std::wstring_view path) noexcept
{
    if (path.size() >= 2 && IsAsciiAlpha(path[0]) && path[1] == L':')
    {
        return true;
    }

    if (path.size() >= 2 && path[0] == L'\\' && path[1] == L'\\')
    {
        return true;
    }

    return false;
}

float MeasureTextWidthDip(IDWriteFactory* factory, IDWriteTextFormat* format, std::wstring_view text) noexcept
{
    if (factory == nullptr || format == nullptr || text.empty())
    {
        return 0.0f;
    }

    wil::com_ptr<IDWriteTextLayout> layout;
    static constexpr float kMeasureWidthDip  = 8192.0f;
    static constexpr float kMeasureHeightDip = 256.0f;
    const HRESULT hr = factory->CreateTextLayout(text.data(), static_cast<UINT32>(text.size()), format, kMeasureWidthDip, kMeasureHeightDip, layout.put());
    if (FAILED(hr) || ! layout)
    {
        return 0.0f;
    }

    DWRITE_TEXT_METRICS metrics{};
    if (FAILED(layout->GetMetrics(&metrics)))
    {
        return 0.0f;
    }

    return metrics.widthIncludingTrailingWhitespace;
}

bool FitsTextWidthDip(IDWriteFactory* factory, IDWriteTextFormat* format, std::wstring_view text, float maxWidthDip) noexcept
{
    if (maxWidthDip <= 0.0f)
    {
        return false;
    }

    const float width = MeasureTextWidthDip(factory, format, text);
    return width <= maxWidthDip + 0.25f;
}

std::wstring BuildTailEllipsisText(std::wstring_view text, IDWriteFactory* factory, IDWriteTextFormat* format, float maxWidthDip) noexcept
{
    if (text.empty() || factory == nullptr || format == nullptr)
    {
        return std::wstring(text);
    }

    if (FitsTextWidthDip(factory, format, text, maxWidthDip))
    {
        return std::wstring(text);
    }

    static constexpr wchar_t kEllipsis[] = L"\u2026";
    if (FitsTextWidthDip(factory, format, kEllipsis, maxWidthDip))
    {
        std::wstring best = kEllipsis;

        size_t low  = 0;
        size_t high = text.size();
        while (low < high)
        {
            const size_t mid = low + (high - low + 1) / 2;
            std::wstring cand;
            cand.reserve(1u + mid);
            cand.append(kEllipsis);
            cand.append(text.substr(text.size() - mid));
            if (FitsTextWidthDip(factory, format, cand, maxWidthDip))
            {
                best = std::move(cand);
                low  = mid;
            }
            else
            {
                high = mid - 1;
            }
        }

        return best;
    }

    return kEllipsis;
}

std::wstring BuildTrailingEllipsisText(std::wstring_view text, IDWriteFactory* factory, IDWriteTextFormat* format, float maxWidthDip) noexcept
{
    if (text.empty() || factory == nullptr || format == nullptr)
    {
        return std::wstring(text);
    }

    if (FitsTextWidthDip(factory, format, text, maxWidthDip))
    {
        return std::wstring(text);
    }

    static constexpr wchar_t kEllipsis[] = L"\u2026";
    if (FitsTextWidthDip(factory, format, kEllipsis, maxWidthDip))
    {
        std::wstring best = kEllipsis;

        size_t low  = 0;
        size_t high = text.size();
        while (low < high)
        {
            const size_t mid = low + (high - low + 1) / 2;
            size_t prefixLen = mid;

            if (prefixLen > 0 && prefixLen < text.size())
            {
                const wchar_t last             = text[prefixLen - 1];
                const wchar_t next             = text[prefixLen];
                const bool lastIsHighSurrogate = (last >= 0xD800 && last <= 0xDBFF);
                const bool nextIsLowSurrogate  = (next >= 0xDC00 && next <= 0xDFFF);
                if (lastIsHighSurrogate && nextIsLowSurrogate)
                {
                    prefixLen -= 1;
                }
            }

            std::wstring cand;
            cand.reserve(prefixLen + 1u);
            cand.append(text.substr(0, prefixLen));
            cand.append(kEllipsis);
            if (FitsTextWidthDip(factory, format, cand, maxWidthDip))
            {
                best = std::move(cand);
                low  = mid;
            }
            else
            {
                high = mid - 1;
            }
        }

        return best;
    }

    return kEllipsis;
}

struct PathEllipsisParts final
{
    std::wstring_view root;
    std::vector<std::wstring_view> segments;
    wchar_t separator = L'\\';
};

PathEllipsisParts SplitPathForEllipsis(std::wstring_view path, bool fileSystemIsWin32) noexcept
{
    PathEllipsisParts parts;
    parts.separator = DeterminePreferredPathSeparator(path, fileSystemIsWin32);

    size_t rootLen = 0;
    if (path.size() >= 2 && IsPathSeparator(path[0]) && IsPathSeparator(path[1]))
    {
        // UNC: \\server\share\...
        auto findSep = [&](size_t start) noexcept -> size_t
        {
            for (size_t i = start; i < path.size(); ++i)
            {
                if (IsPathSeparator(path[i]))
                {
                    return i;
                }
            }
            return std::wstring_view::npos;
        };

        const size_t serverEnd = findSep(2);
        if (serverEnd == std::wstring_view::npos)
        {
            parts.root = path;
            return parts;
        }

        const size_t shareEnd = findSep(serverEnd + 1);
        if (shareEnd == std::wstring_view::npos)
        {
            parts.root = path;
            return parts;
        }

        rootLen    = shareEnd + 1;
        parts.root = path.substr(0, rootLen);
    }
    else if (path.size() >= 2 && IsAsciiAlpha(path[0]) && path[1] == L':')
    {
        // Drive root: C:\...
        rootLen = 2;
        if (path.size() >= 3 && IsPathSeparator(path[2]))
        {
            rootLen = 3;
        }
        parts.root = path.substr(0, rootLen);
    }
    else if (! path.empty() && IsPathSeparator(path[0]))
    {
        rootLen    = 1;
        parts.root = path.substr(0, rootLen);
    }

    size_t pos = rootLen;
    while (pos < path.size() && IsPathSeparator(path[pos]))
    {
        pos += 1;
    }

    while (pos < path.size())
    {
        const size_t start = pos;
        while (pos < path.size() && ! IsPathSeparator(path[pos]))
        {
            pos += 1;
        }
        if (pos > start)
        {
            parts.segments.push_back(path.substr(start, pos - start));
        }

        while (pos < path.size() && IsPathSeparator(path[pos]))
        {
            pos += 1;
        }
    }

    return parts;
}

std::wstring BuildMiddleEllipsisPathText(
    std::wstring_view fullText, bool fileSystemIsWin32, IDWriteFactory* factory, IDWriteTextFormat* format, float maxWidthDip) noexcept
{
    if (fullText.empty() || factory == nullptr || format == nullptr)
    {
        return std::wstring(fullText);
    }

    if (FitsTextWidthDip(factory, format, fullText, maxWidthDip))
    {
        return std::wstring(fullText);
    }

    const bool hasSeparator = fullText.find_first_of(L"\\/") != std::wstring_view::npos;
    if (! hasSeparator)
    {
        return BuildTailEllipsisText(fullText, factory, format, maxWidthDip);
    }

    const PathEllipsisParts parts = SplitPathForEllipsis(fullText, fileSystemIsWin32);
    if (parts.segments.empty())
    {
        return BuildTailEllipsisText(fullText, factory, format, maxWidthDip);
    }

    auto buildCandidate = [&](std::wstring_view root, size_t prefixCount, size_t suffixCount) -> std::wstring
    {
        static constexpr wchar_t kEllipsis[] = L"\u2026";

        std::wstring out;
        out.reserve(fullText.size() + 4);
        out.append(root);

        auto appendSeg = [&](std::wstring_view seg) noexcept
        {
            if (seg.empty())
            {
                return;
            }
            if (! out.empty() && ! IsPathSeparator(out.back()))
            {
                out.push_back(parts.separator);
            }
            out.append(seg.data(), seg.size());
        };

        for (size_t i = 0; i < prefixCount; ++i)
        {
            appendSeg(parts.segments[i]);
        }

        if (prefixCount > 0 && ! out.empty() && ! IsPathSeparator(out.back()))
        {
            out.push_back(parts.separator);
        }

        out.append(kEllipsis);

        if (suffixCount > 0)
        {
            if (! out.empty() && ! IsPathSeparator(out.back()))
            {
                out.push_back(parts.separator);
            }

            const size_t start = parts.segments.size() - suffixCount;
            for (size_t i = start; i < parts.segments.size(); ++i)
            {
                if (i != start)
                {
                    out.push_back(parts.separator);
                }
                out.append(parts.segments[i].data(), parts.segments[i].size());
            }
        }

        return out;
    };

    auto bestForRoot = [&](std::wstring_view root) -> std::wstring
    {
        const size_t total = parts.segments.size();

        size_t suffixCount = 1;
        std::wstring best  = buildCandidate(root, 0, suffixCount);
        if (! FitsTextWidthDip(factory, format, best, maxWidthDip))
        {
            return {};
        }

        while (suffixCount + 1 <= total)
        {
            const std::wstring cand = buildCandidate(root, 0, suffixCount + 1);
            if (! FitsTextWidthDip(factory, format, cand, maxWidthDip))
            {
                break;
            }
            suffixCount += 1;
            best = cand;
        }

        size_t prefixCount = 0;
        while (prefixCount + 1 + suffixCount <= total)
        {
            const std::wstring cand = buildCandidate(root, prefixCount + 1, suffixCount);
            if (! FitsTextWidthDip(factory, format, cand, maxWidthDip))
            {
                break;
            }
            prefixCount += 1;
            best = cand;
        }

        return best;
    };

    std::wstring best = bestForRoot(parts.root);
    if (! best.empty())
    {
        return best;
    }

    best = bestForRoot({});
    if (! best.empty())
    {
        return best;
    }

    const std::wstring_view leaf = parts.segments.back();

    auto tryTrimLeaf = [&](std::wstring_view root) noexcept -> std::wstring
    {
        static constexpr wchar_t kEllipsis[] = L"\u2026";

        std::wstring prefix;
        prefix.reserve(root.size() + 4);
        prefix.append(root);
        prefix.append(kEllipsis);
        prefix.push_back(parts.separator);

        const float prefixWidth = MeasureTextWidthDip(factory, format, prefix);
        if (prefixWidth <= 0.0f || prefixWidth >= maxWidthDip)
        {
            return {};
        }

        float leafWidthDip = maxWidthDip - prefixWidth;
        if (leafWidthDip <= 1.0f)
        {
            return {};
        }

        for (int attempt = 0; attempt < 4 && leafWidthDip > 0.0f; ++attempt)
        {
            const std::wstring trimmedLeaf = BuildTrailingEllipsisText(leaf, factory, format, leafWidthDip);
            std::wstring cand;
            cand.reserve(prefix.size() + trimmedLeaf.size());
            cand.append(prefix);
            cand.append(trimmedLeaf);

            if (FitsTextWidthDip(factory, format, cand, maxWidthDip))
            {
                return cand;
            }

            leafWidthDip -= 1.0f;
        }

        return {};
    };

    best = tryTrimLeaf(parts.root);
    if (! best.empty())
    {
        return best;
    }

    best = tryTrimLeaf({});
    if (! best.empty())
    {
        return best;
    }

    if (FitsTextWidthDip(factory, format, leaf, maxWidthDip))
    {
        return std::wstring(leaf);
    }

    return BuildTrailingEllipsisText(leaf, factory, format, maxWidthDip);
}

std::wstring_view TrimTrailingPathSeparators(std::wstring_view path) noexcept
{
    while (path.size() > 1)
    {
        const wchar_t last = path.back();
        if (last != L'/' && last != L'\\')
        {
            break;
        }
        path.remove_suffix(1);
    }
    return path;
}

std::optional<std::wstring> TryGetParentPathForNavigationGeneric(std::wstring_view path) noexcept
{
    if (path.empty())
    {
        return std::nullopt;
    }

    std::wstring_view trimmed = TrimTrailingPathSeparators(path);
    if (trimmed.empty() || trimmed == L"/" || trimmed == L"\\")
    {
        return std::nullopt;
    }

    const size_t lastSep = trimmed.find_last_of(L"/\\");
    if (lastSep == std::wstring_view::npos)
    {
        return std::nullopt;
    }

    if (lastSep == 0)
    {
        return std::wstring(trimmed.substr(0, 1));
    }

    if (lastSep > 0 && trimmed[lastSep - 1] == L':')
    {
        // For plugin paths like "sftp:/home", the parent should be "sftp:/", not "sftp:" (which has special meaning).
        return std::wstring(trimmed.substr(0, lastSep + 1));
    }

    return std::wstring(trimmed.substr(0, lastSep));
}

wchar_t DeterminePreferredPathSeparator(std::wstring_view path, bool fileSystemIsWin32) noexcept
{
    if (fileSystemIsWin32)
    {
        return L'\\';
    }

    const bool hasForward = path.find(L'/') != std::wstring_view::npos;
    const bool hasBack    = path.find(L'\\') != std::wstring_view::npos;

    if (hasForward && ! hasBack)
    {
        return L'/';
    }

    if (hasBack && ! hasForward)
    {
        return L'\\';
    }

    if (LooksLikeWin32Path(path))
    {
        return L'\\';
    }

    return L'/';
}

std::wstring JoinPath(std::wstring_view parent, std::wstring_view leaf, wchar_t separator) noexcept
{
    if (parent.empty())
    {
        return std::wstring(leaf);
    }

    std::wstring result(parent);
    const wchar_t last = result.back();
    if (last != L'/' && last != L'\\')
    {
        result.push_back(separator);
    }

    result.append(leaf.data(), leaf.size());
    return result;
}

std::optional<std::filesystem::path> TryGetParentPathForNavigation(const std::filesystem::path& path) noexcept
{
    if (path.empty())
    {
        return std::nullopt;
    }

    const std::filesystem::path normalized = path.lexically_normal();
    const std::filesystem::path root       = normalized.root_path();
    if (! root.empty() && normalized == root)
    {
        return std::nullopt;
    }

    const std::filesystem::path parent = normalized.parent_path();
    if (parent.empty() || parent == normalized)
    {
        return std::nullopt;
    }

    if (! parent.has_root_directory())
    {
        // For drive roots, parent_path() can produce "C:" which is not a navigable folder.
        return std::nullopt;
    }

    return parent;
}

D2D1::ColorF Mix(const D2D1::ColorF& a, const D2D1::ColorF& b, float t) noexcept
{
    const float clamped = std::clamp(t, 0.0f, 1.0f);
    return D2D1::ColorF(a.r + (b.r - a.r) * clamped, a.g + (b.g - a.g) * clamped, a.b + (b.b - a.b) * clamped, a.a + (b.a - a.a) * clamped);
}

double Fract(double value) noexcept
{
    return value - std::floor(value);
}

double EaseOutCubic(double t) noexcept
{
    const double clamped = std::clamp(t, 0.0, 1.0);
    const double inv     = 1.0 - clamped;
    return 1.0 - (inv * inv * inv);
}

uint32_t HashU32(uint32_t value) noexcept
{
    // SplitMix32-style mixing (stable across runs).
    value ^= value >> 16;
    value *= 0x7feb352du;
    value ^= value >> 15;
    value *= 0x846ca68bu;
    value ^= value >> 16;
    return value;
}

std::wstring FormatAggregateCountsLine(uint32_t folderCount, uint32_t fileCount) noexcept
{
    if (folderCount > 0 && fileCount > 0)
    {
        return FormatStringResource(g_hInstance, IDS_VIEWERSPACE_AGGREGATE_FOLDERS_FILES_FORMAT, folderCount, fileCount);
    }

    if (folderCount > 0)
    {
        return FormatStringResource(g_hInstance, IDS_VIEWERSPACE_AGGREGATE_FOLDERS_FORMAT, folderCount);
    }

    return FormatStringResource(g_hInstance, IDS_VIEWERSPACE_AGGREGATE_FILES_FORMAT, fileCount);
}

struct CompactBytesText final
{
    std::array<wchar_t, 64> buffer{};
    UINT32 length = 0;
};

CompactBytesText FormatBytesCompactInline(uint64_t bytes) noexcept
{
    CompactBytesText text;

    double value       = static_cast<double>(bytes);
    size_t suffixIndex = 0;
#pragma warning(push)
// 5246: the initialization of a subobject should be wrapped in braces
#pragma warning(disable : 5246)
    static constexpr std::array<std::wstring_view, 5> suffixes = {
        L"B",
        L"KB",
        L"MB",
        L"GB",
        L"TB",
    };
#pragma warning(pop)

    while (value >= 1024.0 && (suffixIndex + 1) < suffixes.size())
    {
        value /= 1024.0;
        suffixIndex += 1;
    }

    const auto& locale   = LocaleFormatting::GetFormatLocale();
    wchar_t* const begin = text.buffer.data();
    constexpr size_t max = (sizeof(text.buffer) / sizeof(text.buffer[0])) - 1;

    if (suffixIndex == 0)
    {
        const auto r         = std::format_to_n(begin, max, locale, L"{:L} {}", bytes, suffixes[suffixIndex]);
        const size_t written = (r.size < max) ? r.size : max;
        begin[written]       = L'\0';
        text.length          = static_cast<UINT32>(written);
    }
    else
    {
        int decimals = 0;
        if (value < 10.0)
        {
            decimals = (value >= 9.995) ? 1 : 2;
        }
        else if (value < 100.0)
        {
            decimals = (value >= 99.95) ? 0 : 1;
        }

        const auto r         = std::format_to_n(begin, max, locale, L"{:.{}Lf} {}", value, decimals, suffixes[suffixIndex]);
        const size_t written = (r.size < max) ? r.size : max;
        begin[written]       = L'\0';
        text.length          = static_cast<UINT32>(written);
    }

    return text;
}

D2D1::ColorF ColorFFromHsv(double hue01, double saturation, double value, float alpha) noexcept
{
    const double h   = Fract(hue01) * 6.0;
    const int sector = static_cast<int>(std::floor(h));
    const double f   = h - static_cast<double>(sector);

    const double s = std::clamp(saturation, 0.0, 1.0);
    const double v = std::clamp(value, 0.0, 1.0);

    const double p = v * (1.0 - s);
    const double q = v * (1.0 - s * f);
    const double t = v * (1.0 - s * (1.0 - f));

    double r = v;
    double g = t;
    double b = p;

    switch (sector)
    {
        case 0:
            r = v;
            g = t;
            b = p;
            break;
        case 1:
            r = q;
            g = v;
            b = p;
            break;
        case 2:
            r = p;
            g = v;
            b = t;
            break;
        case 3:
            r = p;
            g = q;
            b = v;
            break;
        case 4:
            r = t;
            g = p;
            b = v;
            break;
        default:
            r = v;
            g = p;
            b = q;
            break;
    }

    return D2D1::ColorF(static_cast<float>(r), static_cast<float>(g), static_cast<float>(b), alpha);
}

float RectArea(const D2D1_RECT_F& rc) noexcept
{
    const float w = std::max(0.0f, rc.right - rc.left);
    const float h = std::max(0.0f, rc.bottom - rc.top);
    return w * h;
}

class ScanScheduler final
{
public:
    ScanScheduler()  = default;
    ~ScanScheduler() = default;

    ScanScheduler(const ScanScheduler&)             = delete;
    ScanScheduler(const ScanScheduler&&)            = delete;
    ScanScheduler& operator=(const ScanScheduler&)  = delete;
    ScanScheduler& operator=(const ScanScheduler&&) = delete;

    struct VolumeEntry final
    {
        VolumeEntry()  = default;
        ~VolumeEntry() = default;

        VolumeEntry(const VolumeEntry&)            = delete;
        VolumeEntry(VolumeEntry&&)                 = delete;
        VolumeEntry& operator=(const VolumeEntry&) = delete;
        VolumeEntry& operator=(VolumeEntry&&)      = delete;

        std::mutex mutex;
        std::condition_variable cv;
        uint32_t inUse = 0;
    };

    class Permit final
    {
    public:
        Permit() = default;

        explicit Permit(std::shared_ptr<VolumeEntry> entry) noexcept : _entry(std::move(entry))
        {
        }

        Permit(const Permit&)            = delete;
        Permit& operator=(const Permit&) = delete;

        Permit(Permit&& other) noexcept
        {
            _entry = std::move(other._entry);
        }

        Permit& operator=(Permit&& other) noexcept
        {
            if (this != &other)
            {
                Release();
                _entry = std::move(other._entry);
            }
            return *this;
        }

        ~Permit()
        {
            Release();
        }

        explicit operator bool() const noexcept
        {
            return static_cast<bool>(_entry);
        }

    private:
        void Release() noexcept
        {
            if (! _entry)
            {
                return;
            }

            {
                std::scoped_lock lock(_entry->mutex);
                if (_entry->inUse > 0)
                {
                    _entry->inUse -= 1;
                }
            }

            _entry->cv.notify_one();
            _entry.reset();
        }

        std::shared_ptr<VolumeEntry> _entry;
    };

    Permit AcquireForPath(const std::filesystem::path& path, std::stop_token stopToken) noexcept
    {
        const std::wstring volumeKey = TryGetVolumeKey(path);
        return AcquireForKey(volumeKey.empty() ? L"*" : volumeKey, stopToken);
    }

    Permit AcquireForKey(const std::wstring& key, std::stop_token stopToken) noexcept
    {
        std::shared_ptr<VolumeEntry> entry;
        {
            std::scoped_lock lock(_mutex);
            auto it = _byVolume.find(key);
            if (it == _byVolume.end())
            {
                entry = std::make_shared<VolumeEntry>();
                _byVolume.emplace(key, entry);
            }
            else
            {
                entry = it->second;
            }
        }

        if (! entry)
        {
            return {};
        }

        std::unique_lock lock(entry->mutex);
        while (! stopToken.stop_requested())
        {
            const uint32_t limit = GetMaxConcurrentScansPerVolume();
            if (entry->inUse < limit)
            {
                entry->inUse += 1;
                return Permit(entry);
            }

            entry->cv.wait_for(lock, std::chrono::milliseconds(50));
        }

        return {};
    }

    void Shutdown() noexcept
    {
        std::vector<std::shared_ptr<VolumeEntry>> entries;
        {
            std::scoped_lock lock(_mutex);
            entries.reserve(_byVolume.size());
            for (const auto& item : _byVolume)
            {
                if (item.second)
                {
                    entries.push_back(item.second);
                }
            }
            _byVolume.clear();
        }

        for (const auto& entry : entries)
        {
            entry->cv.notify_all();
        }
    }

private:
    static std::wstring TryGetVolumeKey(const std::filesystem::path& path) noexcept
    {
        if (path.empty())
        {
            return {};
        }

        std::array<wchar_t, 1024> buffer{};
        if (GetVolumePathNameW(path.c_str(), buffer.data(), static_cast<DWORD>(buffer.size())) != 0)
        {
            return std::wstring(buffer.data());
        }

        const std::wstring pathText = path.wstring();
        if (pathText.size() >= 2u && pathText[1] == L':')
        {
            wchar_t drive[4]{};
            drive[0] = pathText[0];
            drive[1] = L':';
            drive[2] = L'\\';
            drive[3] = L'\0';
            return std::wstring(drive);
        }

        return {};
    }

private:
    std::mutex _mutex;
    std::unordered_map<std::wstring, std::shared_ptr<VolumeEntry>> _byVolume;
};

std::mutex g_scanSchedulerMutex;
std::unique_ptr<ScanScheduler> g_scanScheduler;

ScanScheduler& GetScanScheduler() noexcept
{
    std::scoped_lock lock(g_scanSchedulerMutex);
    if (! g_scanScheduler)
    {
        g_scanScheduler = std::make_unique<ScanScheduler>();
    }
    return *g_scanScheduler;
}

void ShutdownScanScheduler() noexcept
{
    std::scoped_lock lock(g_scanSchedulerMutex);
    if (! g_scanScheduler)
    {
        return;
    }

    g_scanScheduler->Shutdown();
    g_scanScheduler.reset();
}

struct ScanResultCacheKey final
{
    std::wstring rootKey;
    uint32_t topFilesPerDirectory = 0;
};

struct ScanResultCacheNode final
{
    uint32_t id               = 0;
    uint32_t parentId         = 0;
    bool isDirectory          = false;
    bool isSynthetic          = false;
    uint8_t scanState         = 0;
    uint64_t totalBytes       = 0;
    uint32_t childrenStart    = 0;
    uint32_t childrenCount    = 0;
    uint32_t childrenCapacity = 0;
    uint32_t aggregateFolders = 0;
    uint32_t aggregateFiles   = 0;
    std::wstring name;
};

struct ScanResultCacheFileRecord final
{
    uint32_t id       = 0;
    uint32_t parentId = 0;
    uint64_t bytes    = 0;
    std::wstring name;
};

struct ScanResultSnapshot final
{
    std::vector<ScanResultCacheNode> nodes;
    std::vector<ScanResultCacheFileRecord> fileRecords;
    std::vector<uint32_t> childrenArena;
};

std::wstring NormalizeRootPathForScanCache(const std::filesystem::path& rootPath) noexcept
{
    if (rootPath.empty())
    {
        return {};
    }

    std::error_code ec;
    const std::filesystem::path absolutePath   = std::filesystem::absolute(rootPath, ec);
    const std::filesystem::path normalizedPath = (ec ? rootPath : absolutePath).lexically_normal();

    std::wstring key = normalizedPath.wstring();
    if (key.empty())
    {
        key = rootPath.wstring();
    }

    std::replace(key.begin(), key.end(), L'/', L'\\');

    while (key.size() > 3 && key.back() == L'\\')
    {
        key.pop_back();
    }

    if (key.size() >= 2 && key[1] == L':' && key[0] >= L'a' && key[0] <= L'z')
    {
        key[0] = static_cast<wchar_t>(key[0] - L'a' + L'A');
    }

    return key;
}

class ScanResultCache final
{
public:
    ScanResultCache()  = default;
    ~ScanResultCache() = default;

    ScanResultCache& operator=(ScanResultCache&)  = delete;
    ScanResultCache(ScanResultCache&)             = delete;
    ScanResultCache& operator=(ScanResultCache&&) = delete;
    ScanResultCache(ScanResultCache&&)            = delete;

    void Clear() noexcept
    {
        std::scoped_lock lock(_mutex);
        _entries.clear();
    }

    void Shutdown() noexcept
    {
        Clear();
    }

    void TrimTo(uint32_t maxEntries) noexcept
    {
        const size_t limit = static_cast<size_t>(std::min<uint32_t>(maxEntries, 64u));
        std::scoped_lock lock(_mutex);
        while (_entries.size() > limit)
        {
            _entries.pop_back();
        }
    }

    std::shared_ptr<const ScanResultSnapshot> TryGet(const ScanResultCacheKey& key) noexcept
    {
        if (! g_cacheEnabled.load(std::memory_order_acquire))
        {
            return {};
        }

        const uint32_t maxEntries = g_cacheMaxEntries.load(std::memory_order_acquire);
        if (maxEntries == 0)
        {
            return {};
        }

        const uint32_t ttlSeconds = g_cacheTtlSeconds.load(std::memory_order_acquire);
        const auto now            = std::chrono::steady_clock::now();

        std::scoped_lock lock(_mutex);
        PurgeExpiredLocked(now, ttlSeconds);

        for (size_t i = 0; i < _entries.size(); ++i)
        {
            Entry& entry = _entries[i];
            if (entry.key.topFilesPerDirectory == key.topFilesPerDirectory && entry.key.rootKey == key.rootKey)
            {
                entry.lastUsed                                     = now;
                std::shared_ptr<const ScanResultSnapshot> snapshot = entry.snapshot;

                if (i != 0)
                {
                    std::rotate(_entries.begin(), _entries.begin() + static_cast<ptrdiff_t>(i), _entries.begin() + static_cast<ptrdiff_t>(i) + 1);
                }

                return snapshot;
            }
        }

        return {};
    }

    void Store(ScanResultCacheKey key, std::shared_ptr<ScanResultSnapshot> snapshot) noexcept
    {
        if (! snapshot)
        {
            return;
        }

        if (! g_cacheEnabled.load(std::memory_order_acquire))
        {
            return;
        }

        const uint32_t maxEntries = g_cacheMaxEntries.load(std::memory_order_acquire);
        if (maxEntries == 0)
        {
            return;
        }

        const uint32_t ttlSeconds = g_cacheTtlSeconds.load(std::memory_order_acquire);
        const auto now            = std::chrono::steady_clock::now();

        std::scoped_lock lock(_mutex);
        PurgeExpiredLocked(now, ttlSeconds);

        for (size_t i = 0; i < _entries.size(); ++i)
        {
            if (_entries[i].key.topFilesPerDirectory == key.topFilesPerDirectory && _entries[i].key.rootKey == key.rootKey)
            {
                _entries.erase(_entries.begin() + static_cast<ptrdiff_t>(i));
                break;
            }
        }

        Entry entry;
        entry.key      = std::move(key);
        entry.snapshot = std::move(snapshot);
        entry.inserted = now;
        entry.lastUsed = now;
        _entries.push_back(std::move(entry));
        std::rotate(_entries.begin(), _entries.end() - 1, _entries.end());

        const size_t limit = static_cast<size_t>(std::min<uint32_t>(maxEntries, 64u));
        while (_entries.size() > limit)
        {
            _entries.pop_back();
        }
    }

private:
    struct Entry final
    {
        ScanResultCacheKey key;
        std::shared_ptr<const ScanResultSnapshot> snapshot;
        std::chrono::steady_clock::time_point inserted;
        std::chrono::steady_clock::time_point lastUsed;
    };

    void PurgeExpiredLocked(const std::chrono::steady_clock::time_point now, uint32_t ttlSeconds)
    {
        if (ttlSeconds == 0)
        {
            return;
        }

        const auto ttl = std::chrono::seconds(ttlSeconds);
        for (size_t i = 0; i < _entries.size();)
        {
            if ((now - _entries[i].inserted) > ttl)
            {
                _entries.erase(_entries.begin() + static_cast<ptrdiff_t>(i));
                continue;
            }
            i += 1;
        }
    }

    std::mutex _mutex;
    std::vector<Entry> _entries;
};

std::mutex g_scanResultCacheMutex;
std::unique_ptr<ScanResultCache> g_scanResultCache;

ScanResultCache& GetScanResultCache() noexcept
{
    std::scoped_lock lock(g_scanResultCacheMutex);
    if (! g_scanResultCache)
    {
        g_scanResultCache = std::make_unique<ScanResultCache>();
    }
    return *g_scanResultCache;
}

void ShutdownScanResultCache() noexcept
{
    std::scoped_lock lock(g_scanResultCacheMutex);
    if (! g_scanResultCache)
    {
        return;
    }

    g_scanResultCache->Shutdown();
    g_scanResultCache.reset();
}

void ShutdownViewerSpaceModuleStateImpl() noexcept
{
    ShutdownScanResultCache();
    ShutdownScanScheduler();
    DestroyViewerSpaceClassBackgroundBrushState();
}

} // namespace

const char* GetViewerSpaceStaticConfigurationSchema() noexcept
{
    return GetViewerSpaceStaticConfigurationSchemaImpl();
}

void ShutdownViewerSpaceModuleState() noexcept
{
    ShutdownViewerSpaceModuleStateImpl();
}

ViewerSpace::ViewerSpace()
{
    _metaId                      = L"builtin/viewer-space";
    _metaShortId                 = L"viewspace";
    _metaName                    = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_NAME);
    _metaDescription             = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_DESCRIPTION);
    _scanInProgressWatermarkText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_WATERMARK_IN_PROGRESS);
    _scanIncompleteWatermarkText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_WATERMARK_SCAN_INCOMPLETE);

    _metaData.id          = _metaId.c_str();
    _metaData.shortId     = _metaShortId.c_str();
    _metaData.name        = _metaName.empty() ? nullptr : _metaName.c_str();
    _metaData.description = _metaDescription.empty() ? nullptr : _metaDescription.c_str();
    _metaData.author      = nullptr;
    _metaData.version     = VERSINFO_PLUGIN_VERSION;

    static_cast<void>(SetConfiguration(nullptr));
}

ViewerSpace::~ViewerSpace()
{
    CancelScanAndWait();
    CancelScanCacheBuild();

    {
        std::scoped_lock lock(_updateMutex);
        _pendingUpdates.clear();
    }

    _fileSystem.reset();
    _hostPaneExecute.reset();
}

void ViewerSpace::SetHost(IHost* host) noexcept
{
    _hostPaneExecute.reset();

    if (! host)
    {
        return;
    }

    wil::com_ptr<IHostPaneExecute> paneExecute;
    const HRESULT hr = host->QueryInterface(__uuidof(IHostPaneExecute), paneExecute.put_void());
    if (SUCCEEDED(hr) && paneExecute)
    {
        _hostPaneExecute = std::move(paneExecute);
    }
}

HRESULT STDMETHODCALLTYPE ViewerSpace::QueryInterface(REFIID riid, void** ppvObject) noexcept
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

ULONG STDMETHODCALLTYPE ViewerSpace::AddRef() noexcept
{
    return _refCount.fetch_add(1) + 1;
}

ULONG STDMETHODCALLTYPE ViewerSpace::Release() noexcept
{
    const ULONG remaining = _refCount.fetch_sub(1) - 1;
    if (remaining == 0)
    {
        delete this;
    }
    return remaining;
}

HRESULT STDMETHODCALLTYPE ViewerSpace::GetMetaData(const PluginMetaData** metaData) noexcept
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

HRESULT STDMETHODCALLTYPE ViewerSpace::GetConfigurationSchema(const char** schemaJsonUtf8) noexcept
{
    if (schemaJsonUtf8 == nullptr)
    {
        return E_POINTER;
    }

    *schemaJsonUtf8 = GetViewerSpaceStaticConfigurationSchema();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerSpace::SetConfiguration(const char* configurationJsonUtf8) noexcept
{
    uint32_t topFilesPerDirectory        = 96;
    uint32_t scanThreads                 = 1;
    uint32_t maxConcurrentScansPerVolume = 1;
    bool cacheEnabled                    = true;
    uint32_t cacheTtlSeconds             = 60;
    uint32_t cacheMaxEntries             = 1;

    if (configurationJsonUtf8 != nullptr && configurationJsonUtf8[0] != '\0')
    {
        const std::string_view utf8(configurationJsonUtf8);
        if (! utf8.empty())
        {
            yyjson_doc* doc = yyjson_read(utf8.data(), utf8.size(), YYJSON_READ_JSON5 | YYJSON_READ_ALLOW_BOM);
            if (doc)
            {
                auto freeDoc = wil::scope_exit([&] { yyjson_doc_free(doc); });

                yyjson_val* root = yyjson_doc_get_root(doc);
                if (root && yyjson_is_obj(root))
                {
                    yyjson_val* topFiles = yyjson_obj_get(root, "topFilesPerDirectory");
                    if (topFiles && yyjson_is_int(topFiles))
                    {
                        const int64_t value = yyjson_get_int(topFiles);
                        if (value >= 0)
                        {
                            topFilesPerDirectory = static_cast<uint32_t>(std::min<int64_t>(value, kAdaptiveFileCandidateCeiling));
                        }
                    }

                    yyjson_val* threads = yyjson_obj_get(root, "scanThreads");
                    if (threads && yyjson_is_int(threads))
                    {
                        const int64_t value = yyjson_get_int(threads);
                        if (value > 0)
                        {
                            scanThreads = static_cast<uint32_t>(std::clamp<int64_t>(value, 1, 16));
                        }
                    }

                    yyjson_val* concurrent = yyjson_obj_get(root, "maxConcurrentScansPerVolume");
                    if (concurrent && yyjson_is_int(concurrent))
                    {
                        const int64_t value = yyjson_get_int(concurrent);
                        if (value > 0)
                        {
                            maxConcurrentScansPerVolume = static_cast<uint32_t>(std::clamp<int64_t>(value, 1, 8));
                        }
                    }

                    yyjson_val* cache = yyjson_obj_get(root, "cacheEnabled");
                    if (cache)
                    {
                        if (yyjson_is_str(cache))
                        {
                            const char* value = yyjson_get_str(cache);
                            if (value != nullptr)
                            {
                                cacheEnabled = (strcmp(value, "1") == 0) || (strcmp(value, "true") == 0) || (strcmp(value, "on") == 0);
                            }
                        }
                        else if (yyjson_is_bool(cache))
                        {
                            cacheEnabled = yyjson_get_bool(cache);
                        }
                    }

                    yyjson_val* ttl = yyjson_obj_get(root, "cacheTtlSeconds");
                    if (ttl && yyjson_is_int(ttl))
                    {
                        const int64_t value = yyjson_get_int(ttl);
                        if (value >= 0)
                        {
                            cacheTtlSeconds = static_cast<uint32_t>(std::min<int64_t>(value, 3600));
                        }
                    }

                    yyjson_val* maxEntries = yyjson_obj_get(root, "cacheMaxEntries");
                    if (maxEntries && yyjson_is_int(maxEntries))
                    {
                        const int64_t value = yyjson_get_int(maxEntries);
                        if (value >= 0)
                        {
                            cacheMaxEntries = static_cast<uint32_t>(std::min<int64_t>(value, 16));
                        }
                    }
                }
            }
        }
    }

    _config.topFilesPerDirectory        = topFilesPerDirectory;
    _config.scanThreads                 = scanThreads;
    _config.maxConcurrentScansPerVolume = maxConcurrentScansPerVolume;
    _config.cacheEnabled                = cacheEnabled;
    _config.cacheTtlSeconds             = cacheTtlSeconds;
    _config.cacheMaxEntries             = cacheMaxEntries;
    _layoutCandidateCache.clear();

    g_maxConcurrentScansPerVolume.store(_config.maxConcurrentScansPerVolume, std::memory_order_release);
    g_cacheEnabled.store(_config.cacheEnabled, std::memory_order_release);
    g_cacheTtlSeconds.store(_config.cacheTtlSeconds, std::memory_order_release);
    g_cacheMaxEntries.store(_config.cacheMaxEntries, std::memory_order_release);

    if (! _config.cacheEnabled || _config.cacheMaxEntries == 0)
    {
        GetScanResultCache().Clear();
    }
    else
    {
        GetScanResultCache().TrimTo(_config.cacheMaxEntries);
    }

    _configurationJson = std::format("{{\"topFilesPerDirectory\":{},\"scanThreads\":{},\"maxConcurrentScansPerVolume\":{},\"cacheEnabled\":\"{}\","
                                     "\"cacheTtlSeconds\":{},\"cacheMaxEntries\":{}}}",
                                     _config.topFilesPerDirectory,
                                     _config.scanThreads,
                                     _config.maxConcurrentScansPerVolume,
                                     _config.cacheEnabled ? "1" : "0",
                                     _config.cacheTtlSeconds,
                                     _config.cacheMaxEntries);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerSpace::GetConfiguration(const char** configurationJsonUtf8) noexcept
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

HRESULT STDMETHODCALLTYPE ViewerSpace::SomethingToSave(BOOL* pSomethingToSave) noexcept
{
    if (pSomethingToSave == nullptr)
    {
        return E_POINTER;
    }

    const bool isDefault = _config.topFilesPerDirectory == 96u && _config.scanThreads == 1u && _config.maxConcurrentScansPerVolume == 1u &&
                           _config.cacheEnabled && _config.cacheTtlSeconds == 60u && _config.cacheMaxEntries == 1u;
    *pSomethingToSave    = isDefault ? FALSE : TRUE;
    return S_OK;
}

ATOM ViewerSpace::RegisterWndClass(HINSTANCE instance) noexcept
{
    if (g_viewerSpaceClassAtom)
    {
        g_viewerSpaceClassBackgroundBrush->classRegistered = true;
        return g_viewerSpaceClassAtom;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.lpfnWndProc   = &ViewerSpace::WndProcThunk;
    wc.hInstance     = instance;
    wc.lpszClassName = kClassName;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = GetActiveViewerSpaceClassBackgroundBrush();
    wc.style         = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
    g_viewerSpaceClassAtom = RegisterClassExW(&wc);
    if (g_viewerSpaceClassAtom == 0)
    {
        const DWORD lastError = GetLastError();
        if (lastError == ERROR_CLASS_ALREADY_EXISTS)
        {
            g_viewerSpaceClassAtom = 1;
        }
    }
    if (g_viewerSpaceClassAtom)
    {
        g_viewerSpaceClassBackgroundBrush->classRegistered = true;
    }
    return g_viewerSpaceClassAtom;
}

void ViewerSpace::UnregisterWndClassIfIdle() noexcept
{
    if (g_viewerSpaceWindowCount.load(std::memory_order_acquire) != 0 || ! g_viewerSpaceClassAtom)
    {
        return;
    }

    if (UnregisterClassW(kClassName, g_hInstance) == 0)
    {
        return;
    }

    g_viewerSpaceClassAtom = 0;
    g_viewerSpaceClassBackgroundBrush->classRegistered = false;
    g_viewerSpaceClassBackgroundBrush->pendingBrush.reset();
    g_viewerSpaceClassBackgroundBrush->pendingColor = CLR_INVALID;
    g_viewerSpaceClassBackgroundBrush->activeBrush.reset();
    g_viewerSpaceClassBackgroundBrush->activeColor = CLR_INVALID;
}

LRESULT CALLBACK ViewerSpace::WndProcThunk(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    if (msg == WM_NCCREATE)
    {
        const auto* cs = reinterpret_cast<const CREATESTRUCTW*>(lp);
        auto* self     = static_cast<ViewerSpace*>(cs ? cs->lpCreateParams : nullptr);
        if (self)
        {
            g_viewerSpaceWindowCount.fetch_add(1, std::memory_order_acq_rel);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        }
    }

    auto* self = reinterpret_cast<ViewerSpace*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
    if (self)
    {
        return self->WndProc(hwnd, msg, wp, lp);
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

LRESULT ViewerSpace::WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
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
        case WndMsg::kViewerSpaceDebugGetTooltipSnapshot:
        {
            auto* snapshot = reinterpret_cast<WndMsg::ViewerSpaceTooltipDebugSnapshot*>(lp);
            if (! snapshot)
            {
                return FALSE;
            }

            *snapshot                     = {};
            snapshot->tooltipNodeId       = _tooltipNodeId;
            snapshot->tooltipTextLength   = _tooltipText.size();
            snapshot->tooltipAnchorXDip   = _tooltipAnchorDip.x;
            snapshot->tooltipAnchorYDip   = _tooltipAnchorDip.y;
            snapshot->tooltipMaxWidthDip  = kTooltipMaxWidthDip;
            snapshot->tooltipMaxHeightDip = kTooltipMaxHeightDip;
            snapshot->tooltipPaintCount = _tooltipPaintCount;
            snapshot->hasRenderTarget = _renderTarget != nullptr;
            snapshot->hasTooltipFormat = _tooltipFormat != nullptr;
            snapshot->hoverNodeId = _hoverNodeId;
            snapshot->hasLastMouseMoveClientPoint = _debugHasLastMouseMoveClientPoint;
            snapshot->lastMouseMoveClientX = _debugLastMouseMoveClientPoint.x;
            snapshot->lastMouseMoveClientY = _debugLastMouseMoveClientPoint.y;
            snapshot->hasLastContextMenuScreenPoint = _debugHasLastContextMenuScreenPoint;
            snapshot->lastContextMenuScreenX = _debugLastContextMenuScreenPoint.x;
            snapshot->lastContextMenuScreenY = _debugLastContextMenuScreenPoint.y;
            snapshot->lastContextMenuHitNodeId = _debugLastContextMenuHitNodeId;
            return TRUE;
        }
        case WndMsg::kViewerSpaceDebugCompareHitTesting:
        {
            auto* snapshot = reinterpret_cast<WndMsg::ViewerSpaceHitTestDebugSnapshot*>(lp);
            if (! snapshot)
            {
                return FALSE;
            }

            *snapshot = {};
            static constexpr uint32_t kMaxHitTestSamples = 256u;
            const size_t stride = _drawItems.size() > kMaxHitTestSamples ? ((_drawItems.size() + kMaxHitTestSamples - 1u) / kMaxHitTestSamples) : 1u;
            for (size_t index = 0; index < _drawItems.size() && snapshot->sampleCount < kMaxHitTestSamples; index += stride)
            {
                const DrawItem& item = _drawItems[index];
                if (item.lod == DrawItem::Lod::Culled)
                {
                    continue;
                }

                float gap = kItemGapDip - static_cast<float>(item.depth) * 0.15f;
                gap       = std::clamp(gap, 0.5f, kItemGapDip);

                D2D1_RECT_F rc = item.currentRect;
                rc.left += gap;
                rc.top += gap;
                rc.right -= gap;
                rc.bottom -= gap;
                if (rc.right <= rc.left || rc.bottom <= rc.top || RectArea(rc) < kMinHitAreaDip2)
                {
                    continue;
                }

                const float xDip = (rc.left + rc.right) * 0.5f;
                const float yDip = (rc.top + rc.bottom) * 0.5f;
                const std::optional<uint32_t> gridHit = HitTestTreemap(xDip, yDip);
                const uint32_t gridCandidates         = _lastHitTestCandidatesChecked;
                uint32_t linearCandidates             = 0u;
                const std::optional<uint32_t> linearHit = HitTestTreemapLinear(xDip, yDip, linearCandidates);

                snapshot->sampleCount += 1u;
                snapshot->maxGridCandidatesChecked = std::max(snapshot->maxGridCandidatesChecked, gridCandidates);
                snapshot->maxLinearCandidatesChecked = std::max(snapshot->maxLinearCandidatesChecked, linearCandidates);
                if (gridHit.has_value())
                {
                    snapshot->gridHitCount += 1u;
                }
                if (linearHit.has_value())
                {
                    snapshot->linearHitCount += 1u;
                }

                const bool mismatch =
                    gridHit.has_value() != linearHit.has_value() || (gridHit.has_value() && linearHit.has_value() && gridHit.value() != linearHit.value());
                if (mismatch)
                {
                    snapshot->mismatchCount += 1u;
                    snapshot->lastMismatchGridNodeId   = gridHit.value_or(0u);
                    snapshot->lastMismatchLinearNodeId = linearHit.value_or(0u);
                }
            }

            Debug::Perf::Emit(L"viewer.space.hit_test.debug_samples",
                              L"grid-vs-linear",
                              0u,
                              snapshot->sampleCount,
                              snapshot->mismatchCount,
                              snapshot->mismatchCount == 0u ? S_OK : E_FAIL);
            return TRUE;
        }
        case WndMsg::kViewerSpaceDebugGetPerfSnapshot:
        case WndMsg::kViewerSpaceDebugForcePerfSample:
        {
            auto* snapshot = reinterpret_cast<WndMsg::ViewerSpacePerfDebugSnapshot*>(lp);
            if (! snapshot)
            {
                return FALSE;
            }

            *snapshot             = {};
            snapshot->rendererMode = _renderer.d2dContext ? WndMsg::ViewerSpacePerfRendererMode::DeviceContext
                                                           : (_renderTarget ? WndMsg::ViewerSpacePerfRendererMode::HwndRenderTarget
                                                                            : WndMsg::ViewerSpacePerfRendererMode::None);
            snapshot->scanState    = [&]() noexcept
            {
                switch (_overallState)
                {
                    case ScanState::Queued: return WndMsg::ViewerSpacePerfScanState::Queued;
                    case ScanState::Scanning: return WndMsg::ViewerSpacePerfScanState::Scanning;
                    case ScanState::Done: return WndMsg::ViewerSpacePerfScanState::Done;
                    case ScanState::Error: return WndMsg::ViewerSpacePerfScanState::Error;
                    case ScanState::Canceled: return WndMsg::ViewerSpacePerfScanState::Canceled;
                    case ScanState::NotStarted:
                    default: return WndMsg::ViewerSpacePerfScanState::NotStarted;
                }
            }();
            snapshot->hasRenderTarget = _renderTarget != nullptr && (_renderer.d2dContext == nullptr || _renderer.targetBitmap != nullptr);

            for (const Node& node : _nodes)
            {
                if (node.id == 0)
                {
                    continue;
                }

                if (node.isSynthetic)
                {
                    snapshot->syntheticCount += 1u;
                    snapshot->syntheticBytes += node.totalBytes;
                }
                else if (node.isDirectory)
                {
                    snapshot->realDirectoryCount += 1u;
                }
            }
            snapshot->fileCandidateCount = _fileRecords.size();
            for (const FileRecord& record : _fileRecords)
            {
                snapshot->fileCandidateBytes += record.bytes;
            }
            if (const Node* root = TryGetRealNode(_rootNodeId))
            {
                snapshot->rootTotalBytes = root->totalBytes;
            }
            for (const auto& [id, syntheticNode] : _syntheticNodes)
            {
                (void)id;
                snapshot->syntheticBytes += syntheticNode.totalBytes;
            }
            snapshot->syntheticCount += _syntheticNodes.size();

            {
                std::scoped_lock lock(_updateMutex);
                snapshot->pendingQueueCount = _pendingUpdates.size();
                snapshot->pendingQueueBytes = _pendingUpdateBytes.load(std::memory_order_relaxed);
            }

            snapshot->drawItemCount = _drawItems.size();
            for (const DrawItem& item : _drawItems)
            {
                if (item.lod == DrawItem::Lod::Culled)
                {
                    snapshot->culledTileCount += 1u;
                }
            }
            snapshot->visibleTileCount               = snapshot->drawItemCount - snapshot->culledTileCount;
            snapshot->staticCacheGeneration          = _staticTreemapCache.generation;
            snapshot->staticCacheHits                = _staticTreemapCache.hits;
            snapshot->staticCacheMisses              = _staticTreemapCache.misses;
            snapshot->staticCacheRecordCount         = _staticTreemapCache.recordCount;
            snapshot->staticCacheBytes               = _staticTreemapCache.bytes;
            snapshot->scanCacheSnapshotBytes         = _lastScanCacheSnapshotBytes;
            snapshot->lastStaticCacheRecordUs        = _staticTreemapCache.lastRecordUs;
            snapshot->lastPaintUs                    = _lastPaintUs;
            snapshot->lastLayoutUs                   = _lastLayoutUs;
            snapshot->lastDrainUs                    = _lastDrainUs;
            _lastWorkingSetBytes                     = SampleCurrentWorkingSetBytes();
            snapshot->lastWorkingSetBytes            = _lastWorkingSetBytes;
            snapshot->lastTileDrawCount              = _lastTileDrawCount;
            snapshot->lastTextDrawCount              = _lastTextDrawCount;
            snapshot->lastHitTestUs                  = _lastHitTestUs;
            snapshot->lastHitTestCandidatesChecked   = _lastHitTestCandidatesChecked;
            snapshot->effectiveFileCandidateBudget   = _scanTopFilesPerDirectory;
            snapshot->rendererDeviceCreateCount      = _rendererDeviceCreateCount;
            snapshot->swapChainResizeCount           = _swapChainResizeCount;
            snapshot->rendererBrushCreateCount       = _rendererBrushCreateCount;
            snapshot->rendererTextFormatCreateCount  = _rendererTextFormatCreateCount;
            snapshot->swapChainWidthPx               = _renderer.swapChainWidthPx;
            snapshot->swapChainHeightPx              = _renderer.swapChainHeightPx;
            snapshot->rendererFailureStage           = _rendererFailureStage;
            snapshot->rendererFailureHr              = _rendererFailureHr;
            snapshot->layoutGeneration               = _layoutGeneration;
            snapshot->hitGridCellCount               = static_cast<uint32_t>(std::min<size_t>(_hitGrid.cells.size(), std::numeric_limits<uint32_t>::max()));
            snapshot->hitGridMaxCandidatesPerCell    = _hitGrid.maxCandidatesPerCell;

            Debug::Perf::EmitValue(L"viewer.space.model.directory_count", static_cast<uint64_t>(snapshot->realDirectoryCount));
            Debug::Perf::EmitValue(L"viewer.space.model.file_candidate_count", static_cast<uint64_t>(snapshot->fileCandidateCount));
            Debug::Perf::EmitValue(L"viewer.space.model.synthetic_count", static_cast<uint64_t>(snapshot->syntheticCount));
            Debug::Perf::EmitValue(L"viewer.space.queue.pending_bytes", snapshot->pendingQueueBytes);
            return TRUE;
        }
        case WndMsg::kViewerSpaceDebugForceRendererFault:
        {
            const auto mode = static_cast<WndMsg::ViewerSpaceRendererFaultDebugMode>(wp);
            if (mode != WndMsg::ViewerSpaceRendererFaultDebugMode::D2DRecreateTarget && mode != WndMsg::ViewerSpaceRendererFaultDebugMode::DxgiDeviceRemoved)
            {
                return FALSE;
            }

            _debugForcedRendererFault = static_cast<uint32_t>(mode);
            if (_hWnd)
            {
                InvalidateRect(_hWnd.get(), nullptr, FALSE);
            }
            return TRUE;
        }
        case WndMsg::kViewerSpaceDebugShowTooltipOverlay:
        {
            _tooltipNodeId    = 1u;
            _tooltipText      = L"ViewerSpace debug tooltip overlay text that wraps through Direct2D.";
            _tooltipAnchorDip = D2D1::Point2F(40.0f, GetHeaderBottomDip() + 24.0f);
            if (_hWnd)
            {
                InvalidateRect(_hWnd.get(), nullptr, FALSE);
                UpdateWindow(_hWnd.get());
            }
            return TRUE;
        }
#endif
        case WM_CREATE: OnCreate(hwnd); return 0;
        case WM_DESTROY: OnDestroy(); return 0;
        case WM_SIZE: OnSize(static_cast<UINT>(LOWORD(lp)), static_cast<UINT>(HIWORD(lp))); return 0;
        case WM_DPICHANGED: OnDpiChanged(hwnd, static_cast<UINT>(HIWORD(wp)), reinterpret_cast<const RECT*>(lp)); return 0;
        case WM_DPICHANGED_AFTERPARENT: OnDpiChanged(hwnd, GetDpiForWindow(hwnd), nullptr); return 0;
        case WM_GETMINMAXINFO:
            if (! _embeddedMode)
            {
                if (auto* info = reinterpret_cast<MINMAXINFO*>(lp))
                {
                    Common::WindowSizing::ApplyMinimumClientTrackSizeForDips(hwnd, *info, 520, 360);
                }
            }
            return 0;
        case WM_PAINT: OnPaint(); return 0;
        case WM_ERASEBKGND: return _allowEraseBkgnd ? DefWindowProcW(hwnd, msg, wp, lp) : 1;
        case WM_CLOSE: static_cast<void>(Close()); return 0;
        case WM_COMMAND: OnCommand(hwnd, LOWORD(wp)); return 0;
        case WM_SYSKEYDOWN:
            if (! _embeddedMode && (wp == VK_F10 || wp == VK_MENU) && _menuBarHost.FocusFirstItem())
            {
                return 0;
            }
            break;
        case WM_SYSCHAR:
            if (! _embeddedMode && wp >= 0x20u && _menuBarHost.ActivateMnemonic(static_cast<wchar_t>(wp)))
            {
                return 0;
            }
            break;
        case WM_KEYDOWN:
        {
            const bool alt = (GetKeyState(VK_MENU) & 0x8000) != 0;
            OnKeyDown(wp, alt);
            return 0;
        }
        case WM_MOUSEMOVE:
        {
            const POINTS pt = {static_cast<short>(LOWORD(lp)), static_cast<short>(HIWORD(lp))};
            OnMouseMove(pt.x, pt.y);
            return 0;
        }
        case WM_MOUSELEAVE: OnMouseLeave(); return 0;
        case WM_LBUTTONDOWN:
        {
            const POINTS pt = {static_cast<short>(LOWORD(lp)), static_cast<short>(HIWORD(lp))};
            OnLButtonDown(pt.x, pt.y);
            return 0;
        }
        case WM_LBUTTONDBLCLK:
        {
            const POINTS pt = {static_cast<short>(LOWORD(lp)), static_cast<short>(HIWORD(lp))};
            OnLButtonDblClk(pt.x, pt.y);
            return 0;
        }
        case WM_CONTEXTMENU:
        {
            POINT screenPt{static_cast<LONG>(static_cast<short>(LOWORD(lp))), static_cast<LONG>(static_cast<short>(HIWORD(lp)))};
            if (screenPt.x == -1 && screenPt.y == -1)
            {
                RECT client{};
                if (GetClientRect(hwnd, &client) != 0)
                {
                    screenPt.x = client.left + ((client.right - client.left) / 2);
                    screenPt.y = client.top + ((client.bottom - client.top) / 2);
                    ClientToScreen(hwnd, &screenPt);
                }
                else
                {
                    screenPt = {};
                }
            }
            OnContextMenu(hwnd, screenPt);
            return 0;
        }
        case WM_TIMER: OnTimer(static_cast<UINT_PTR>(wp)); return 0;
        case WM_NCACTIVATE: OnNcActivate(hwnd, wp != FALSE); return DefWindowProcW(hwnd, msg, wp, lp);
        case WM_NCDESTROY: return OnNcDestroy(hwnd, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

void ViewerSpace::OnCreate(HWND hwnd)
{
    _allowEraseBkgnd = true;
    _dpi             = static_cast<float>(GetDpiForWindow(hwnd));

    if (! _embeddedMode && ! _menuHandle)
    {
        _menuHandle.reset(GetMenu(hwnd));
    }
    if (! _embeddedMode && _menuHandle)
    {
        _menuBarHost.SetTheme(_hasTheme ? RedSalamander::DxUi::MakeThemePaletteFromViewerTheme(_theme) : RedSalamander::DxUi::MakeDefaultThemePalette(false));
        _menuBarHost.SetRefreshMenuStateCallback([this]
        {
            if (_hWnd)
            {
                UpdateMenuState(_hWnd.get(), false);
            }
        });
        _menuBarHost.SetOnTabBoundary([this](bool) noexcept
        {
            if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
            {
                SetFocus(_hWnd.get());
            }
            return true;
        });
        _menuBarHost.SetOnEscape([this]() noexcept
        {
            if (_hWnd && IsWindow(_hWnd.get()) != FALSE)
            {
                SetFocus(_hWnd.get());
            }
            return true;
        });
        static_cast<void>(_menuBarHost.Attach(g_hInstance, hwnd, _menuHandle.get()));
    }

    SetTimer(hwnd, kTimerAnimationId, kAnimationIntervalMs, nullptr);
    ApplyThemeToWindow(hwnd);
    ApplyPendingViewerSpaceClassBackgroundBrush(hwnd);
    UpdateMenuState(hwnd);
    if (! _embeddedMode)
    {
        _menuBarHost.UpdateLayout();
    }
}

void ViewerSpace::OnNcActivate(HWND hwnd, bool windowActive) noexcept
{
    ApplyTitleBarTheme(hwnd, windowActive);
}

void ViewerSpace::OnDestroy()
{
    CancelScanAndWait();
    CancelScanCacheBuild();
#ifdef ENABLE_TESTS
    _rendererFailureStage = kRendererDiscardDestroy;
    _rendererFailureHr    = 0u;
#endif
    DiscardDirect2D();
    _fileSystem.reset();
    _fileSystemName.clear();
    _fileSystemShortId.clear();
    _fileSystemIsWin32 = true;

    NotifyViewerClosed();
}

void ViewerSpace::OnSize(UINT width, UINT height) noexcept
{
    _clientSize.cx = static_cast<LONG>(width);
    _clientSize.cy = static_cast<LONG>(height);

    if (! _embeddedMode)
    {
        _menuBarHost.UpdateLayout();
    }
    _layoutDirty = true;
    InvalidateStaticTreemapCache();
    InvalidateRect(_hWnd.get(), nullptr, FALSE);
}

void ViewerSpace::OnDpiChanged(HWND hwnd, UINT dpi, const RECT* suggestedRect) noexcept
{
    _dpi = static_cast<float>(dpi == 0 ? USER_DEFAULT_SCREEN_DPI : dpi);

    if (suggestedRect != nullptr && ! _embeddedMode)
    {
        SetWindowPos(hwnd,
                     nullptr,
                     suggestedRect->left,
                     suggestedRect->top,
                     suggestedRect->right - suggestedRect->left,
                     suggestedRect->bottom - suggestedRect->top,
                     SWP_NOZORDER | SWP_NOACTIVATE);
    }

#ifdef ENABLE_TESTS
    _rendererFailureStage = kRendererDiscardDpiChanged;
    _rendererFailureHr    = 0u;
#endif
    DiscardDirect2D();
    ApplyThemeToWindow(hwnd);
    ApplyPendingViewerSpaceClassBackgroundBrush(hwnd);
    if (! _embeddedMode)
    {
        _menuBarHost.UpdateLayout();
    }
    _layoutDirty = true;
    InvalidateRect(hwnd, nullptr, FALSE);
}

void ViewerSpace::OnPaint()
{
    const auto paintStartedAt = std::chrono::steady_clock::now();

    if (! _hWnd)
    {
        return;
    }

    PAINTSTRUCT ps{};
    wil::unique_hdc_paint hdc = wil::BeginPaint(_hWnd.get(), &ps);
    _allowEraseBkgnd          = false;
    static_cast<void>(hdc);

    if (! EnsureDirect2D(_hWnd.get()))
    {
        return;
    }

    EnsureLayoutForView();

    HRESULT drawHr = S_OK;
    _renderTarget->BeginDraw();
    auto endDraw = wil::scope_exit([&] { drawHr = _renderTarget->EndDraw(); });

    const D2D1::ColorF bg          = _hasTheme ? ColorFFromArgb(_theme.backgroundArgb) : D2D1::ColorF(D2D1::ColorF::White);
    const D2D1::ColorF textColor   = _hasTheme ? ColorFFromArgb(_theme.textArgb) : D2D1::ColorF(D2D1::ColorF::Black);
    const D2D1::ColorF accentColor = _hasTheme ? ColorFFromArgb(_theme.accentArgb) : D2D1::ColorF(D2D1::ColorF::CornflowerBlue);
    const D2D1::ColorF selectionBg = _hasTheme ? ColorFFromArgb(_theme.selectionBackgroundArgb) : accentColor;

    const bool highContrast = _hasTheme && _theme.highContrast != FALSE;
    const bool rainbowMode  = _hasTheme && _theme.rainbowMode != FALSE;
    const bool darkMode     = _hasTheme && _theme.darkMode != FALSE;

    const double nowSeconds = NowSeconds();

    _renderTarget->Clear(bg);

#ifdef ENABLE_TESTS
    auto applyForcedRendererFault = [&](HRESULT& forcedDrawHr, HRESULT& forcedPresentHr) noexcept
    {
        const uint32_t mode = _debugForcedRendererFault;
        if (mode == 0u)
        {
            return;
        }

        _debugForcedRendererFault = 0u;
        if (mode == static_cast<uint32_t>(WndMsg::ViewerSpaceRendererFaultDebugMode::D2DRecreateTarget))
        {
            forcedDrawHr = D2DERR_RECREATE_TARGET;
        }
        else if (mode == static_cast<uint32_t>(WndMsg::ViewerSpaceRendererFaultDebugMode::DxgiDeviceRemoved))
        {
            forcedDrawHr    = S_OK;
            forcedPresentHr = DXGI_ERROR_DEVICE_REMOVED;
        }
    };
#endif

    const float headerTopDip    = GetHeaderTopDip();
    const float headerBottomDip = GetHeaderBottomDip();
    const D2D1_RECT_F headerRc  = D2D1::RectF(0.0f, headerTopDip, DipFromPx(_clientSize.cx), headerBottomDip);

    const bool scanActive = _overallState == ScanState::Queued || _overallState == ScanState::Scanning;
    const bool completionOverlayActive =
        ! scanActive && _overallState == ScanState::Done && _scanCompletedSinceSeconds > 0.0 &&
        (nowSeconds - _scanCompletedSinceSeconds) >= 0.0 && (nowSeconds - _scanCompletedSinceSeconds) < kScanCompletedOverlaySeconds;
    uint64_t culledTileCount = 0u;
    uint64_t textDrawCandidateCount = 0u;
    for (const DrawItem& item : _drawItems)
    {
        if (item.lod == DrawItem::Lod::Culled)
        {
            culledTileCount += 1u;
        }
        if (item.lod >= DrawItem::Lod::Medium)
        {
            textDrawCandidateCount += 1u;
        }
    }
    const uint64_t tileDrawCount = static_cast<uint64_t>(_drawItems.size()) - culledTileCount;
#ifdef ENABLE_TESTS
    _lastTileDrawCount = tileDrawCount;
    _lastTextDrawCount = textDrawCandidateCount;
#endif

    if (CanReplayStaticTreemapCache(scanActive, completionOverlayActive))
    {
        _staticTreemapCache.hits += 1u;
        _renderTarget->DrawBitmap(_staticTreemapCache.bitmap.get(), nullptr, 1.0f, D2D1_BITMAP_INTERPOLATION_MODE_NEAREST_NEIGHBOR, nullptr);
        PaintHoverOverlay(accentColor);
        PaintTooltipOverlay();

        endDraw.reset();

        HRESULT presentHr = S_OK;
        if (SUCCEEDED(drawHr) && _renderer.swapChain)
        {
            presentHr = _renderer.swapChain->Present(1u, 0u);
        }
#ifdef ENABLE_TESTS
        applyForcedRendererFault(drawHr, presentHr);
#endif
        const HRESULT renderHr = FAILED(drawHr) ? drawHr : presentHr;

        const uint64_t paintUs = Debug::Perf::ElapsedUs(paintStartedAt);
#ifdef ENABLE_TESTS
        _lastPaintUs = paintUs;
#endif
        Debug::Perf::Emit(
            L"viewer.space.render.paint_us", L"static-cache-hit", paintUs, static_cast<uint64_t>(_drawItems.size()), _staticTreemapCache.hits, renderHr);
        Debug::Perf::Emit(L"viewer.space.render.static_cache_replay_us", L"bitmap", paintUs, static_cast<uint64_t>(_drawItems.size()), _staticTreemapCache.bytes, renderHr);
        Debug::Perf::Emit(L"viewer.space.render.tile_draw_count", L"static-cache-hit", 0u, 0u, culledTileCount, renderHr);
        Debug::Perf::Emit(L"viewer.space.render.text_draw_count", L"static-cache-hit", 0u, 0u, textDrawCandidateCount, renderHr);
        Debug::Perf::Emit(L"viewer.space.layout.visible_tiles", L"static-cache-hit", 0u, tileDrawCount, culledTileCount, renderHr);

        if (! _openToFirstPaintEmitted && _openStartedAt != std::chrono::steady_clock::time_point{})
        {
            _openToFirstPaintEmitted = true;
            Debug::Perf::Emit(L"viewer.space.open_to_first_paint_us",
                              L"static-cache-hit",
                              Debug::Perf::ElapsedUs(_openStartedAt),
                              static_cast<uint64_t>(_drawItems.size()),
                              0u,
                              renderHr);
        }

        if (drawHr == D2DERR_RECREATE_TARGET || presentHr == DXGI_ERROR_DEVICE_REMOVED || presentHr == DXGI_ERROR_DEVICE_RESET)
        {
#ifdef ENABLE_TESTS
            _rendererFailureStage = kRendererDiscardStaticDeviceLost;
            _rendererFailureHr    = static_cast<uint32_t>(FAILED(drawHr) ? drawHr : presentHr);
#endif
            DiscardDirect2D();
            if (_hWnd)
            {
                InvalidateRect(_hWnd.get(), nullptr, FALSE);
            }
        }
        else if (FAILED(drawHr))
        {
            Debug::Warning(L"ViewerSpace: EndDraw failed during static cache replay: 0x{:08X}", static_cast<unsigned long>(drawHr));
        }
        else if (FAILED(presentHr))
        {
            Debug::Warning(L"ViewerSpace: Present failed during static cache replay: 0x{:08X}", static_cast<unsigned long>(presentHr));
        }
        return;
    }

    if (_brushAccent)
    {
        const float barHeight = 5.0f;
        const D2D1_RECT_F track =
            D2D1::RectF(headerRc.left + kPaddingDip, headerRc.bottom - barHeight - 4.0f, headerRc.right - kPaddingDip, headerRc.bottom - 4.0f);
        if (track.right > track.left && track.bottom > track.top)
        {
            const float radius      = barHeight * 0.5f;
            const float trackAlpha  = highContrast ? 0.32f : (scanActive ? (darkMode ? 0.22f : 0.16f) : 0.10f);
            D2D1::ColorF trackColor = accentColor;
            if (highContrast)
            {
                trackColor = textColor;
            }
            else if (rainbowMode)
            {
                const double hue01 = Fract((nowSeconds - _animationStartSeconds) * 0.08);
                trackColor         = ColorFFromHsv(hue01, 0.90, darkMode ? 0.98 : 0.90, 1.0f);
            }

            _brushAccent->SetColor(D2D1::ColorF(trackColor.r, trackColor.g, trackColor.b, trackAlpha));
            _renderTarget->FillRoundedRectangle(D2D1::RoundedRect(track, radius, radius), _brushAccent.get());

            if (scanActive)
            {
                const float t           = static_cast<float>(Fract((nowSeconds - _animationStartSeconds) * 0.72));
                const float w           = std::max(42.0f, (track.right - track.left) * 0.26f);
                const float x           = track.left - w + t * ((track.right - track.left) + w * 2.0f);
                const D2D1_RECT_F chunk = D2D1::RectF(x, track.top, x + w, track.bottom);

                _renderTarget->PushAxisAlignedClip(track, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
                const float chunkAlpha = highContrast ? 0.85f : 0.70f;

                auto chunkColorFor = [&](double offset) noexcept -> D2D1::ColorF
                {
                    if (highContrast)
                    {
                        return textColor;
                    }

                    if (rainbowMode)
                    {
                        const double hue01 = Fract((nowSeconds - _animationStartSeconds) * 0.08 + static_cast<double>(t) * 0.24 + offset);
                        return ColorFFromHsv(hue01, 0.95, darkMode ? 0.99 : 0.92, 1.0f);
                    }

                    return accentColor;
                };

                const float inset1 = std::min(14.0f, w * 0.22f);
                const float inset2 = std::min(24.0f, w * 0.36f);

                auto fillRounded = [&](D2D1_RECT_F rc, float alpha, double hueOffset) noexcept
                {
                    if (rc.right <= rc.left || rc.bottom <= rc.top)
                    {
                        return;
                    }

                    const D2D1::ColorF c = chunkColorFor(hueOffset);
                    _brushAccent->SetColor(D2D1::ColorF(c.r, c.g, c.b, alpha));
                    const float r = std::min(radius, std::min((rc.right - rc.left) * 0.5f, (rc.bottom - rc.top) * 0.5f));
                    _renderTarget->FillRoundedRectangle(D2D1::RoundedRect(rc, r, r), _brushAccent.get());
                };

                fillRounded(chunk, chunkAlpha * 0.35f, 0.0);

                D2D1_RECT_F mid = chunk;
                mid.left += inset1;
                mid.right -= inset1;
                fillRounded(mid, chunkAlpha * 0.60f, 0.05);

                D2D1_RECT_F core = chunk;
                core.left += inset2;
                core.right -= inset2;
                fillRounded(core, chunkAlpha * 0.92f, 0.10);
                _renderTarget->PopAxisAlignedClip();
            }
        }
    }

    if (_brushText)
    {
        if (_headerStatusText.empty())
        {
            UpdateHeaderTextCache();
        }

        const std::wstring& status = _headerStatusText;

        const float buttonSide           = kHeaderButtonWidthDip;
        const D2D1_RECT_F upButtonRc     = D2D1::RectF(0.0f, headerTopDip, buttonSide, headerBottomDip);
        const bool showCancel            = scanActive;
        const D2D1_RECT_F cancelButtonRc = showCancel ? D2D1::RectF(headerRc.right - buttonSide, headerTopDip, headerRc.right, headerBottomDip)
                                                      : D2D1::RectF(headerRc.right, headerTopDip, headerRc.right, headerBottomDip);
        const bool canNavigateUp         = CanNavigateUp();

        if (_brushBackground)
        {
            const float hoverAlpha = 0.18f;
            if (canNavigateUp && _hoverHeaderHit == HeaderHit::Up)
            {
                _brushBackground->SetColor(D2D1::ColorF(selectionBg.r, selectionBg.g, selectionBg.b, hoverAlpha));
                _renderTarget->FillRectangle(upButtonRc, _brushBackground.get());
            }
            else if (_hoverHeaderHit == HeaderHit::Cancel)
            {
                _brushBackground->SetColor(D2D1::ColorF(selectionBg.r, selectionBg.g, selectionBg.b, hoverAlpha));
                _renderTarget->FillRectangle(cancelButtonRc, _brushBackground.get());
            }
        }

        if (_headerIconFormat)
        {
            const D2D1_COLOR_F oldBrushColor = _brushText->GetColor();

            if (! canNavigateUp)
            {
                const float disabledAlpha = highContrast ? 0.60f : (darkMode ? 0.38f : 0.48f);
                _brushText->SetColor(D2D1::ColorF(oldBrushColor.r, oldBrushColor.g, oldBrushColor.b, oldBrushColor.a * disabledAlpha));
            }

            const wchar_t upGlyph[] = L"↑";
            _renderTarget->DrawTextW(upGlyph, 1, _headerIconFormat.get(), upButtonRc, _brushText.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
            _brushText->SetColor(oldBrushColor);

            if (showCancel)
            {
                const wchar_t cancelGlyph[] = L"×";
                _renderTarget->DrawTextW(cancelGlyph, 1, _headerIconFormat.get(), cancelButtonRc, _brushText.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
            }
        }

        std::wstring pathFallback;
        std::wstring_view pathText = _viewPathText;
        if (pathText.empty() && ! _scanRootPath.empty())
        {
            pathFallback = _scanRootPath;
            pathText     = pathFallback;
        }
        if (pathText.empty())
        {
            pathText = _metaName;
        }

        const float contentTop    = headerTopDip + 4.0f;
        const float contentBottom = headerBottomDip - 10.0f;
        D2D1_RECT_F contentRc     = headerRc;
        contentRc.left += buttonSide + kPaddingDip;
        contentRc.right -= (showCancel ? buttonSide : 0.0f) + kPaddingDip;
        if (contentRc.right > contentRc.left && contentBottom > contentTop)
        {
            const float lineHeight  = (contentBottom - contentTop) / 3.0f;
            const D2D1_RECT_F line1 = D2D1::RectF(contentRc.left, contentTop, contentRc.right, contentTop + lineHeight);
            const D2D1_RECT_F line2 = D2D1::RectF(contentRc.left, line1.bottom, contentRc.right, line1.bottom + lineHeight);
            const D2D1_RECT_F line3 = D2D1::RectF(contentRc.left, line2.bottom, contentRc.right, contentBottom);

            const float availableWidth = std::max(0.0f, contentRc.right - contentRc.left);
            const float columnGapDip   = 12.0f;

            float statusWidthDip = 0.0f;
            if (_dwriteFactory && _headerStatusFormatRight && ! status.empty())
            {
                statusWidthDip = MeasureTextWidthDip(_dwriteFactory.get(), _headerStatusFormatRight.get(), status);
            }

            float statusColumnWidthDip = 0.0f;
            if (statusWidthDip > 0.0f)
            {
                statusColumnWidthDip = statusWidthDip + 10.0f;
            }
            else
            {
                statusColumnWidthDip = std::min(160.0f, availableWidth * 0.35f);
            }
            statusColumnWidthDip = std::clamp(statusColumnWidthDip, 48.0f, availableWidth);

            D2D1_RECT_F line1Right = line1;
            line1Right.left        = std::max(line1Right.left, line1Right.right - statusColumnWidthDip);

            D2D1_RECT_F line1Left = line1;
            line1Left.right       = std::max(line1Left.left, line1Right.left - columnGapDip);

            const float line2RightWidth = std::min(320.0f, availableWidth * 0.40f);
            D2D1_RECT_F line2Left       = line2;
            line2Left.right             = std::max(line2Left.left, line2Left.right - line2RightWidth);
            D2D1_RECT_F line2Right      = line2;
            line2Right.left             = line2Left.right;

            if (_headerFormat && ! pathText.empty())
            {
                std::wstring_view pathToDraw = pathText;
                const float maxWidthDip      = std::max(0.0f, line1Left.right - line1Left.left);
                if (_dwriteFactory && maxWidthDip > 1.0f)
                {
                    if (_headerPathSourceText != pathText || std::abs(maxWidthDip - _headerPathDisplayMaxWidthDip) > 0.5f)
                    {
                        _headerPathSourceText.assign(pathText.data(), pathText.size());
                        _headerPathDisplayMaxWidthDip = maxWidthDip;
                        _headerPathDisplayText =
                            BuildMiddleEllipsisPathText(_headerPathSourceText, _fileSystemIsWin32, _dwriteFactory.get(), _headerFormat.get(), maxWidthDip);
                    }

                    if (! _headerPathDisplayText.empty())
                    {
                        pathToDraw = _headerPathDisplayText;
                    }
                }

                _renderTarget->DrawTextW(
                    pathToDraw.data(), static_cast<UINT32>(pathToDraw.size()), _headerFormat.get(), line1Left, _brushText.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
            }

            if (_headerStatusFormatRight && ! status.empty())
            {
                _renderTarget->DrawTextW(status.c_str(),
                                         static_cast<UINT32>(status.size()),
                                         _headerStatusFormatRight.get(),
                                         line1Right,
                                         _brushText.get(),
                                         D2D1_DRAW_TEXT_OPTIONS_CLIP);
            }

            if (scanActive && _headerInfoFormatRight && ! _headerProcessingText.empty())
            {
                _renderTarget->DrawTextW(_headerProcessingText.c_str(),
                                         static_cast<UINT32>(_headerProcessingText.size()),
                                         _headerInfoFormatRight.get(),
                                         line2Right,
                                         _brushText.get(),
                                         D2D1_DRAW_TEXT_OPTIONS_CLIP);
            }

            if (_headerInfoFormat)
            {
                if (! _headerCountsText.empty() || ! _headerSizeText.empty())
                {
                    if (! _headerCountsText.empty())
                    {
                        _renderTarget->DrawTextW(_headerCountsText.c_str(),
                                                 static_cast<UINT32>(_headerCountsText.size()),
                                                 _headerInfoFormat.get(),
                                                 line2Left,
                                                 _brushText.get(),
                                                 D2D1_DRAW_TEXT_OPTIONS_CLIP);
                    }

                    if (! _headerSizeText.empty())
                    {
                        _renderTarget->DrawTextW(_headerSizeText.c_str(),
                                                 static_cast<UINT32>(_headerSizeText.size()),
                                                 _headerInfoFormat.get(),
                                                 line3,
                                                 _brushText.get(),
                                                 D2D1_DRAW_TEXT_OPTIONS_CLIP);
                    }
                }
            }
        }
    }

    const D2D1_RECT_F treemapRc                                = D2D1::RectF(0.0f, headerBottomDip, DipFromPx(_clientSize.cx), DipFromPx(_clientSize.cy));
    const D2D1_RECT_F treemapLayoutRc                          = D2D1::RectF(kPaddingDip,
                                                                             headerBottomDip + kPaddingDip,
                                                                             std::max(kPaddingDip, treemapRc.right - kPaddingDip),
                                                                             std::max(headerBottomDip + kPaddingDip, treemapRc.bottom - kPaddingDip));
    const float treemapLayoutAreaDip2                          = RectArea(treemapLayoutRc);
    static constexpr float kLargeTileAreaFractionWhileScanning = 0.10f;
    static constexpr float kLargeTileAreaFractionIdle          = 0.10f;
    const float largeTileAreaFractionThreshold                 = scanActive ? kLargeTileAreaFractionWhileScanning : kLargeTileAreaFractionIdle;

    uint64_t viewBytes   = 0;
    const Node* viewNode = TryGetRealNode(_viewNodeId);
    if (viewNode != nullptr)
    {
        viewBytes = viewNode->totalBytes;
    }

    auto resolveItem = [&](uint32_t itemId) noexcept -> std::optional<ResolvedItem>
    {
        return ResolveTreemapItem(itemId);
    };

    auto rectForItem = [&](const DrawItem& item) noexcept -> D2D1_RECT_F
    {
        float gap = kItemGapDip - static_cast<float>(item.depth) * 0.15f;
        gap       = std::clamp(gap, 0.5f, kItemGapDip);

        D2D1_RECT_F rc = item.currentRect;
        rc.left += gap;
        rc.top += gap;
        rc.right -= gap;
        rc.bottom -= gap;
        return rc;
    };

    auto baseColorForNode = [&](const ResolvedItem& nodeRef) noexcept -> D2D1::ColorF
    {
        double ratio = 0.0;
        if (viewBytes > 0)
        {
            ratio = static_cast<double>(nodeRef.totalBytes) / static_cast<double>(viewBytes);
            ratio = std::clamp(ratio, 0.0, 1.0);
        }

        const double sizeFactor = std::sqrt(ratio);
        const uint32_t hashed   = HashU32(nodeRef.id);

        D2D1::ColorF base = bg;
        if (highContrast)
        {
            base = Mix(bg, textColor, nodeRef.isDirectory ? 0.18f : 0.10f);
        }
        else if (rainbowMode)
        {
            const double hue01  = static_cast<double>(hashed) / 4294967296.0;
            const double jitter = static_cast<double>(HashU32(hashed ^ 0x68bc21ebu)) / 4294967296.0;

            double saturation = nodeRef.isDirectory ? 0.95 : 0.88;
            double value      = darkMode ? 0.92 : 0.80;
            if (! nodeRef.isDirectory)
            {
                value = darkMode ? 0.86 : 0.74;
            }

            saturation += (jitter - 0.5) * 0.08;
            value += (jitter - 0.5) * 0.06;

            const D2D1::ColorF rainbow = ColorFFromHsv(hue01, saturation, value, 1.0f);

            float mixT = nodeRef.isDirectory ? 0.82f : 0.74f;
            mixT += darkMode ? 0.08f : 0.04f;
            mixT += static_cast<float>(0.10 * sizeFactor);
            mixT = std::clamp(mixT, 0.55f, 1.0f);
            base = Mix(bg, rainbow, mixT);
        }
        else
        {
            const double variant01 = static_cast<double>(HashU32(hashed ^ 0x9e3779b9u)) / 4294967296.0;
            float accentMix        = nodeRef.isDirectory ? 0.55f : 0.35f;
            accentMix += static_cast<float>((variant01 - 0.5) * 0.30);
            accentMix = std::clamp(accentMix, 0.0f, 1.0f);

            const D2D1::ColorF nodeAccent = Mix(accentColor, selectionBg, accentMix);

            float mixT = nodeRef.isDirectory ? 0.24f : 0.14f;
            mixT += static_cast<float>((nodeRef.isDirectory ? 0.34 : 0.22) * sizeFactor);
            if (darkMode)
            {
                mixT += nodeRef.isDirectory ? 0.06f : 0.03f;
            }
            mixT = std::clamp(mixT, 0.0f, 0.90f);
            base = Mix(bg, nodeAccent, mixT);

            const double shade01      = static_cast<double>(HashU32(hashed ^ 0x85ebca6bu)) / 4294967296.0;
            const float shadeSigned   = static_cast<float>((shade01 - 0.5) * 2.0);
            const float shadeStrength = nodeRef.isDirectory ? 0.16f : 0.12f;

            const D2D1::ColorF lighter = darkMode ? Mix(base, textColor, 0.28f) : Mix(base, bg, 0.22f);
            const D2D1::ColorF darker  = darkMode ? Mix(base, bg, 0.22f) : Mix(base, textColor, 0.22f);

            if (shadeSigned >= 0.0f)
            {
                base = Mix(base, lighter, shadeSigned * shadeStrength);
            }
            else
            {
                base = Mix(base, darker, (-shadeSigned) * shadeStrength);
            }
        }

        return base;
    };

    const float labelAreaThresholdDip2 = scanActive ? (32.0f * 32.0f) : kMinHitAreaDip2;

    static constexpr double kMiniSpinnerPhaseSpeed = 1.8;
    static constexpr double kMiniSpinnerHueSpeed   = 0.10;
    static constexpr double kBigSpinnerPhaseSpeed  = 0.95;
    static constexpr double kBigSpinnerHueSpeed    = 0.06;

    auto drawMiniSpinner = [&](const D2D1_POINT_2F& center, float radius, uint32_t seed, double phaseSpeed, double hueSpeed) noexcept
    {
        if (! _brushAccent || radius <= 1.0f)
        {
            return;
        }

        static constexpr int kSegments = 12;
        static constexpr double kTwoPi = 6.28318530717958647692;

        const double seed01  = static_cast<double>(HashU32(seed)) / 4294967296.0;
        const double phase01 = Fract((nowSeconds - _animationStartSeconds) * phaseSpeed + seed01);
        const int head       = static_cast<int>(phase01 * static_cast<double>(kSegments));

        const float baseAlpha = highContrast ? 1.0f : 0.92f;
        const float stroke    = std::clamp(radius * 0.20f, 1.1f, 1.8f);

        for (int i = 0; i < kSegments; ++i)
        {
            const int dist    = (head - i + kSegments) % kSegments;
            const float fade  = 1.0f - static_cast<float>(dist) / static_cast<float>(kSegments);
            const float alpha = baseAlpha * (0.15f + 0.85f * fade);

            D2D1::ColorF c = accentColor;
            if (highContrast)
            {
                c = textColor;
            }
            else if (rainbowMode)
            {
                const double hue01 = Fract((nowSeconds - _animationStartSeconds) * hueSpeed + seed01 + static_cast<double>(i) / static_cast<double>(kSegments));
                c                  = ColorFFromHsv(hue01, 0.95, darkMode ? 0.99 : 0.92, 1.0f);
            }

            _brushAccent->SetColor(D2D1::ColorF(c.r, c.g, c.b, alpha));

            const double ang = (static_cast<double>(i) / static_cast<double>(kSegments)) * kTwoPi;
            const float cx   = static_cast<float>(std::cos(ang));
            const float cy   = static_cast<float>(std::sin(ang));

            const float r0 = radius * 0.52f;
            const float r1 = radius;

            const D2D1_POINT_2F p0 = D2D1::Point2F(center.x + cx * r0, center.y + cy * r0);
            const D2D1_POINT_2F p1 = D2D1::Point2F(center.x + cx * r1, center.y + cy * r1);
            _renderTarget->DrawLine(p0, p1, _brushAccent.get(), stroke);
        }
    };

    // Fill pass: parents first, children later.
    for (const auto& item : _drawItems)
    {
        const std::optional<ResolvedItem> node = resolveItem(item.nodeId);
        if (! node.has_value())
        {
            continue;
        }

        const ResolvedItem& nodeRef = node.value();

        D2D1_RECT_F rc = rectForItem(item);

        if (rc.right <= treemapRc.left || rc.left >= treemapRc.right || rc.bottom <= treemapRc.top || rc.top >= treemapRc.bottom)
        {
            continue;
        }

        const float area = RectArea(rc);
        if (item.lod == DrawItem::Lod::Culled)
        {
            continue;
        }

        D2D1::ColorF base              = baseColorForNode(nodeRef);
        const bool isOtherBucket       = nodeRef.isSynthetic;
        const bool isRealDirectory     = nodeRef.isDirectory && ! nodeRef.isSynthetic;
        const bool incompleteDirectory = isRealDirectory && (nodeRef.scanState == ScanState::NotStarted || nodeRef.scanState == ScanState::Queued ||
                                                             nodeRef.scanState == ScanState::Scanning);
        if (incompleteDirectory && ! highContrast)
        {
            float dimT = darkMode ? 0.38f : 0.28f;
            if (nodeRef.scanState == ScanState::Queued || nodeRef.scanState == ScanState::NotStarted)
            {
                dimT = darkMode ? 0.52f : 0.40f;
            }
            base = Mix(base, bg, dimT);
        }

        if (_brushBackground)
        {
            float fillAlpha = highContrast ? 1.0f : 0.96f;
            if (! nodeRef.isDirectory)
            {
                fillAlpha -= 0.04f;
            }
            if (isOtherBucket && ! highContrast)
            {
                fillAlpha = std::clamp(fillAlpha - (darkMode ? 0.05f : 0.07f), 0.55f, 0.96f);
            }
            _brushBackground->SetColor(D2D1::ColorF(base.r, base.g, base.b, fillAlpha));
            _renderTarget->FillRectangle(rc, _brushBackground.get());
        }

        if (_brushShading && ! highContrast && ! scanActive)
        {
            _brushShading->SetStartPoint(D2D1::Point2F(rc.left, rc.top));
            _brushShading->SetEndPoint(D2D1::Point2F(rc.right, rc.bottom));
            _renderTarget->FillRectangle(rc, _brushShading.get());
        }

        if (isOtherBucket && _brushOutline && ! highContrast)
        {
            const float w    = std::max(0.0f, rc.right - rc.left);
            const float h    = std::max(0.0f, rc.bottom - rc.top);
            const float side = std::min(w, h);
            if (side >= 18.0f)
            {
                float spacing         = std::clamp(side * 0.12f, 8.0f, 18.0f);
                const float thickness = std::clamp(side * 0.010f, 0.9f, 1.4f);
                float alpha           = darkMode ? 0.11f : 0.09f;
                if (rainbowMode)
                {
                    alpha += 0.04f;
                }

                const D2D1::ColorF hatch = rainbowMode ? textColor : Mix(accentColor, textColor, 0.55f);
                _brushOutline->SetColor(D2D1::ColorF(hatch.r, hatch.g, hatch.b, alpha));

                _renderTarget->PushAxisAlignedClip(rc, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
                const float diag = h;
                for (float x0 = rc.left - diag; x0 < rc.right; x0 += spacing)
                {
                    const D2D1_POINT_2F p0 = D2D1::Point2F(x0, rc.bottom);
                    const D2D1_POINT_2F p1 = D2D1::Point2F(x0 + diag, rc.top);
                    _renderTarget->DrawLine(p0, p1, _brushOutline.get(), thickness);
                }
                _renderTarget->PopAxisAlignedClip();
            }
        }

        const bool expanded = isRealDirectory && item.labelHeightDip > 0.0f;

        float directoryHeaderHeightDip = 0.0f;
        if (isRealDirectory)
        {
            const float areaFraction   = (treemapLayoutAreaDip2 > 1.0f) ? (area / treemapLayoutAreaDip2) : 0.0f;
            const bool wantsTallHeader = incompleteDirectory && areaFraction >= largeTileAreaFractionThreshold;

            const float h                   = std::max(0.0f, rc.bottom - rc.top);
            const float baseHeaderHeightDip = std::clamp(24.0f - static_cast<float>(item.depth) * 2.0f, 20.0f, 24.0f);
            float desiredHeaderHeightDip    = expanded ? std::max(item.labelHeightDip, baseHeaderHeightDip) : baseHeaderHeightDip;
            if (wantsTallHeader)
            {
                const float twoLine    = std::clamp(38.0f - static_cast<float>(item.depth) * 2.0f, 30.0f, 38.0f);
                desiredHeaderHeightDip = std::max(desiredHeaderHeightDip, twoLine);
            }
            const float maxHeaderHeightDip = wantsTallHeader ? 44.0f : 24.0f;
            desiredHeaderHeightDip         = std::clamp(desiredHeaderHeightDip, 20.0f, maxHeaderHeightDip);

            if (h >= desiredHeaderHeightDip)
            {
                directoryHeaderHeightDip = desiredHeaderHeightDip;
            }
        }

        if (directoryHeaderHeightDip > 0.0f && _brushBackground)
        {
            D2D1_RECT_F headerBar = rc;
            headerBar.bottom      = std::min(rc.bottom, rc.top + directoryHeaderHeightDip);

            D2D1::ColorF headerColor = Mix(base, bg, darkMode ? 0.10f : 0.22f);
            float headerAlpha        = 0.96f;
            if (rainbowMode)
            {
                headerColor = Mix(base, bg, darkMode ? 0.18f : 0.12f);
                headerAlpha = 0.96f;
            }
            _brushBackground->SetColor(D2D1::ColorF(headerColor.r, headerColor.g, headerColor.b, headerAlpha));
            _renderTarget->FillRectangle(headerBar, _brushBackground.get());

            if (_brushOutline && ! highContrast)
            {
                D2D1_RECT_F sep         = headerBar;
                sep.top                 = std::max(headerBar.top, headerBar.bottom - 1.0f);
                sep.bottom              = headerBar.bottom;
                const D2D1::ColorF line = Mix(base, textColor, 0.20f);
                _brushOutline->SetColor(D2D1::ColorF(line.r, line.g, line.b, 0.55f));
                _renderTarget->FillRectangle(sep, _brushOutline.get());
            }
        }
    }

    // Border + labels pass: draw children first, then parents (so parent borders stay visible).
    for (size_t idx = _drawItems.size(); idx-- > 0;)
    {
        const DrawItem& item = _drawItems[idx];

        const std::optional<ResolvedItem> node = resolveItem(item.nodeId);
        if (! node.has_value())
        {
            continue;
        }

        const ResolvedItem& nodeRef = node.value();

        D2D1_RECT_F rc = rectForItem(item);

        if (rc.right <= treemapRc.left || rc.left >= treemapRc.right || rc.bottom <= treemapRc.top || rc.top >= treemapRc.bottom)
        {
            continue;
        }

        const float area = RectArea(rc);
        if (item.lod == DrawItem::Lod::Culled || item.lod == DrawItem::Lod::Tiny)
        {
            continue;
        }

        const float tileW                = std::max(0.0f, rc.right - rc.left);
        const float tileH                = std::max(0.0f, rc.bottom - rc.top);
        const bool canShowAtLeastOneLine = tileW >= 28.0f && tileH >= 20.0f;
        const bool canShowTileLabels     = item.lod >= DrawItem::Lod::Medium && (area >= labelAreaThresholdDip2 || canShowAtLeastOneLine);

        D2D1::ColorF base         = baseColorForNode(nodeRef);
        const bool isScanningTile = nodeRef.isDirectory && ! nodeRef.isSynthetic && nodeRef.scanState == ScanState::Scanning;
        if (isScanningTile && ! highContrast)
        {
            const float dimT = darkMode ? 0.38f : 0.28f;
            base             = Mix(base, bg, dimT);
        }
        const bool isRealDirectory = nodeRef.isDirectory && ! nodeRef.isSynthetic;
        const bool expanded        = isRealDirectory && item.labelHeightDip > 0.0f;
        const bool isOtherBucket   = nodeRef.isSynthetic;
        const bool incomplete      = isRealDirectory && (nodeRef.scanState == ScanState::NotStarted || nodeRef.scanState == ScanState::Queued ||
                                                         nodeRef.scanState == ScanState::Scanning);

        const float areaFraction   = (treemapLayoutAreaDip2 > 1.0f) ? (area / treemapLayoutAreaDip2) : 0.0f;
        const bool wantsTallHeader = incomplete && areaFraction >= largeTileAreaFractionThreshold;

        float directoryHeaderHeightDip = 0.0f;
        if (isRealDirectory)
        {
            const float h                   = std::max(0.0f, rc.bottom - rc.top);
            const float baseHeaderHeightDip = std::clamp(24.0f - static_cast<float>(item.depth) * 2.0f, 20.0f, 24.0f);
            float desiredHeaderHeightDip    = expanded ? std::max(item.labelHeightDip, baseHeaderHeightDip) : baseHeaderHeightDip;
            if (wantsTallHeader)
            {
                const float twoLine    = std::clamp(38.0f - static_cast<float>(item.depth) * 2.0f, 30.0f, 38.0f);
                desiredHeaderHeightDip = std::max(desiredHeaderHeightDip, twoLine);
            }
            const float maxHeaderHeightDip = wantsTallHeader ? 44.0f : 24.0f;
            desiredHeaderHeightDip         = std::clamp(desiredHeaderHeightDip, 20.0f, maxHeaderHeightDip);

            if (h >= desiredHeaderHeightDip)
            {
                directoryHeaderHeightDip = desiredHeaderHeightDip;
            }
        }
        const bool hasDirectoryHeader = directoryHeaderHeightDip > 0.0f;

        const std::wstring_view watermarkText = scanActive ? std::wstring_view(_scanInProgressWatermarkText) : std::wstring_view(_scanIncompleteWatermarkText);

        const bool showStatusInHeader = wantsTallHeader && hasDirectoryHeader && directoryHeaderHeightDip >= 32.0f && ! watermarkText.empty();

        if (incomplete && _watermarkFormat && _brushWatermark && ! watermarkText.empty() && ! showStatusInHeader)
        {
            const float w    = std::max(0.0f, rc.right - rc.left);
            const float h    = std::max(0.0f, rc.bottom - rc.top);
            const float side = std::min(w, h);

            if (side >= 72.0f && area >= 3200.0f)
            {
                const D2D1_POINT_2F center = D2D1::Point2F((rc.left + rc.right) * 0.5f, (rc.top + rc.bottom) * 0.5f);
                const float diag           = std::sqrt(w * w + h * h);
                const float textW          = std::max(0.0f, diag - 12.0f);
                const float textH          = std::clamp(side * 0.24f, 24.0f, 56.0f);
                const D2D1_RECT_F textRc   = D2D1::RectF(center.x - textW * 0.5f, center.y - textH * 0.5f, center.x + textW * 0.5f, center.y + textH * 0.5f);

                float alpha = highContrast ? 0.78f : (darkMode ? 0.26f : 0.20f);
                if (rainbowMode)
                {
                    alpha += 0.04f;
                }
                alpha = std::clamp(alpha, 0.14f, 0.88f);

                D2D1::ColorF watermarkColor = textColor;
                if (! highContrast && ! rainbowMode)
                {
                    watermarkColor = Mix(base, textColor, darkMode ? 0.86f : 0.76f);
                }
                _brushWatermark->SetColor(D2D1::ColorF(watermarkColor.r, watermarkColor.g, watermarkColor.b, alpha));

                _renderTarget->PushAxisAlignedClip(rc, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE);
                auto popClip = wil::scope_exit([&] { _renderTarget->PopAxisAlignedClip(); });

                D2D1_MATRIX_3X2_F oldTransform{};
                _renderTarget->GetTransform(&oldTransform);
                _renderTarget->SetTransform(D2D1::Matrix3x2F::Rotation(-35.0f, center) * oldTransform);
                auto restoreTransform = wil::scope_exit([&] { _renderTarget->SetTransform(oldTransform); });

                _renderTarget->DrawTextW(watermarkText.data(),
                                         static_cast<UINT32>(watermarkText.size()),
                                         _watermarkFormat.get(),
                                         textRc,
                                         _brushWatermark.get(),
                                         D2D1_DRAW_TEXT_OPTIONS_CLIP);
            }
        }

        if (_brushOutline)
        {
            float stroke = nodeRef.isDirectory ? 1.25f : 1.0f;
            stroke -= static_cast<float>(item.depth) * 0.06f;
            stroke = std::clamp(stroke, 0.85f, 1.35f);

            ID2D1StrokeStyle* strokeStyle = nullptr;
            D2D1::ColorF outline          = highContrast ? textColor : Mix(base, textColor, nodeRef.isDirectory ? 0.20f : 0.14f);
            if (isOtherBucket)
            {
                stroke      = std::max(stroke, 1.85f);
                strokeStyle = (! highContrast && _otherStrokeStyle) ? _otherStrokeStyle.get() : nullptr;

                if (highContrast)
                {
                    outline = textColor;
                }
                else if (rainbowMode)
                {
                    outline = textColor;
                }
                else
                {
                    outline = Mix(accentColor, textColor, 0.30f);
                }
            }
            _brushOutline->SetColor(D2D1::ColorF(outline.r, outline.g, outline.b, 1.0f));

            const bool isFileTile = ! nodeRef.isDirectory && ! nodeRef.isSynthetic;
            if (isFileTile && canShowTileLabels && item.lod >= DrawItem::Lod::Large)
            {
                const float side    = std::min(tileW, tileH);
                const float dogSize = std::clamp(side * 0.18f, 8.0f, 14.0f);
                float cut           = dogSize + 2.0f;
                cut                 = std::clamp(cut, 6.0f, std::max(6.0f, side - 1.0f));

                if (cut > 1.0f && tileW > cut + 1.0f && tileH > cut + 1.0f)
                {
                    const D2D1_POINT_2F tl = D2D1::Point2F(rc.left, rc.top);
                    const D2D1_POINT_2F tr = D2D1::Point2F(rc.right, rc.top);
                    const D2D1_POINT_2F br = D2D1::Point2F(rc.right, rc.bottom);
                    const D2D1_POINT_2F bl = D2D1::Point2F(rc.left, rc.bottom);

                    const D2D1_POINT_2F cutA = D2D1::Point2F(rc.right - cut, rc.top);
                    const D2D1_POINT_2F cutB = D2D1::Point2F(rc.right, rc.top + cut);

                    _renderTarget->DrawLine(tl, cutA, _brushOutline.get(), stroke, strokeStyle);
                    _renderTarget->DrawLine(cutA, cutB, _brushOutline.get(), stroke, strokeStyle);
                    _renderTarget->DrawLine(cutB, br, _brushOutline.get(), stroke, strokeStyle);
                    _renderTarget->DrawLine(br, bl, _brushOutline.get(), stroke, strokeStyle);
                    _renderTarget->DrawLine(bl, tl, _brushOutline.get(), stroke, strokeStyle);
                }
                else
                {
                    _renderTarget->DrawRectangle(rc, _brushOutline.get(), stroke, strokeStyle);
                }
            }
            else
            {
                _renderTarget->DrawRectangle(rc, _brushOutline.get(), stroke, strokeStyle);
            }
        }

        if (_textFormat && _brushText && canShowTileLabels)
        {
            std::wstring fallbackName;
            std::wstring_view nameView = nodeRef.name;
            if (nameView.empty() && ! nodeRef.isSynthetic)
            {
                fallbackName = BuildItemPathText(nodeRef.id);
                nameView     = fallbackName;
            }

            D2D1_RECT_F labelRc = rc;
            labelRc.left += 6.0f;
            labelRc.top += 4.0f;
            labelRc.right -= 6.0f;
            labelRc.bottom -= 4.0f;

            D2D1_RECT_F spinnerRc{};
            bool showSpinner = false;

            if (hasDirectoryHeader)
            {
                const float headerBottom = std::min(rc.bottom, rc.top + directoryHeaderHeightDip);
                labelRc.bottom           = std::min(labelRc.bottom, headerBottom - 2.0f);
                labelRc.bottom           = std::max(labelRc.bottom, labelRc.top);

                if (incomplete)
                {
                    constexpr float kSpinnerBoxDip = 20.0f;
                    spinnerRc                      = labelRc;
                    spinnerRc.left                 = std::max(labelRc.left, labelRc.right - kSpinnerBoxDip);
                    const float spinnerW           = std::max(0.0f, spinnerRc.right - spinnerRc.left);
                    const float spinnerH           = std::max(0.0f, spinnerRc.bottom - spinnerRc.top);
                    showSpinner                    = spinnerW >= 8.0f && spinnerH >= 8.0f;
                    if (showSpinner)
                    {
                        labelRc.right = std::max(labelRc.left, spinnerRc.left - 2.0f);
                    }
                }
            }
            else
            {
                if (incomplete)
                {
                    constexpr float kSpinnerBoxDip = 20.0f;
                    spinnerRc                      = labelRc;
                    spinnerRc.left                 = std::max(labelRc.left, labelRc.right - kSpinnerBoxDip);
                    spinnerRc.bottom               = std::min(labelRc.bottom, labelRc.top + kSpinnerBoxDip);
                    const float spinnerW           = std::max(0.0f, spinnerRc.right - spinnerRc.left);
                    const float spinnerH           = std::max(0.0f, spinnerRc.bottom - spinnerRc.top);
                    showSpinner                    = spinnerW >= 8.0f && spinnerH >= 8.0f;
                    if (showSpinner)
                    {
                        labelRc.right = std::max(labelRc.left, spinnerRc.left - 2.0f);
                    }
                }

                if (rainbowMode && _brushBackground)
                {
                    const float scrimAlpha = darkMode ? 0.34f : 0.54f;
                    D2D1_RECT_F scrimRc    = labelRc;
                    scrimRc.bottom         = std::min(scrimRc.bottom, scrimRc.top + 34.0f);
                    _brushBackground->SetColor(D2D1::ColorF(bg.r, bg.g, bg.b, scrimAlpha));
                    _renderTarget->FillRectangle(scrimRc, _brushBackground.get());
                }
            }

            D2D1_RECT_F dogEarRc{};
            bool showDogEar = false;
            if (! nodeRef.isDirectory && ! nodeRef.isSynthetic && _brushOutline)
            {
                const float side = std::min(tileW, tileH);
                const float size = std::clamp(side * 0.18f, 8.0f, 14.0f);

                dogEarRc = rc;
                dogEarRc.top += 2.0f;
                dogEarRc.right -= 2.0f;
                dogEarRc.left   = std::max(dogEarRc.left, dogEarRc.right - size);
                dogEarRc.bottom = std::min(dogEarRc.bottom, dogEarRc.top + size);

                const float dogW = std::max(0.0f, dogEarRc.right - dogEarRc.left);
                const float dogH = std::max(0.0f, dogEarRc.bottom - dogEarRc.top);
                showDogEar       = dogW >= 6.0f && dogH >= 6.0f;
                if (showDogEar)
                {
                    labelRc.right = std::max(labelRc.left, std::min(labelRc.right, dogEarRc.left - 2.0f));

                    if (_brushBackground)
                    {
                        D2D1::ColorF revealFill = bg;
                        if (const std::optional<ResolvedItem> parent = ResolveTreemapItem(nodeRef.parentId); parent.has_value())
                        {
                            revealFill = baseColorForNode(parent.value());

                            const bool parentRealDir    = parent->isDirectory && ! parent->isSynthetic;
                            const bool parentIncomplete = parentRealDir && (parent->scanState == ScanState::NotStarted ||
                                                                            parent->scanState == ScanState::Queued || parent->scanState == ScanState::Scanning);
                            if (parentIncomplete && ! highContrast)
                            {
                                float dimT = darkMode ? 0.38f : 0.28f;
                                if (parent->scanState == ScanState::Queued || parent->scanState == ScanState::NotStarted)
                                {
                                    dimT = darkMode ? 0.52f : 0.40f;
                                }
                                revealFill = Mix(revealFill, bg, dimT);
                            }
                        }

                        const float revealAlpha = highContrast ? 1.0f : 0.96f;
                        _brushBackground->SetColor(D2D1::ColorF(revealFill.r, revealFill.g, revealFill.b, revealAlpha));
                        _renderTarget->FillRectangle(dogEarRc, _brushBackground.get());

                        if (_dogEarFlapGeometry)
                        {
                            const D2D1::ColorF flapFill = Mix(base, bg, darkMode ? 0.08f : 0.18f);
                            const float flapAlpha       = highContrast ? 1.0f : 0.88f;
                            _brushBackground->SetColor(D2D1::ColorF(flapFill.r, flapFill.g, flapFill.b, flapAlpha));

                            D2D1_MATRIX_3X2_F oldTransform{};
                            _renderTarget->GetTransform(&oldTransform);

                            const D2D1_MATRIX_3X2_F dogTransform =
                                D2D1::Matrix3x2F::Scale(dogW, dogH) * D2D1::Matrix3x2F::Translation(dogEarRc.left, dogEarRc.top) * oldTransform;
                            _renderTarget->SetTransform(dogTransform);
                            _renderTarget->FillGeometry(_dogEarFlapGeometry.get(), _brushBackground.get());
                            _renderTarget->SetTransform(oldTransform);
                        }
                    }

                    const D2D1::ColorF line = highContrast ? textColor : Mix(base, textColor, 0.52f);
                    _brushOutline->SetColor(D2D1::ColorF(line.r, line.g, line.b, highContrast ? 1.0f : 0.92f));
                    _renderTarget->DrawLine(
                        D2D1::Point2F(dogEarRc.left, dogEarRc.top), D2D1::Point2F(dogEarRc.right, dogEarRc.bottom), _brushOutline.get(), 1.1f);
                }
            }

            D2D1_RECT_F nameRc        = labelRc;
            D2D1_RECT_F statusRc      = {};
            D2D1_RECT_F sizeRc        = labelRc;
            const bool showStatusLine = showStatusInHeader && nodeRef.isDirectory;
            bool showSize             = false;

            if (showStatusLine)
            {
                const float availableH = std::max(0.0f, labelRc.bottom - labelRc.top);
                if (availableH >= 26.0f)
                {
                    const float nameH = std::clamp(availableH * 0.60f, 14.0f, availableH);
                    nameRc.bottom     = std::clamp(labelRc.top + nameH, labelRc.top, labelRc.bottom);

                    statusRc       = labelRc;
                    statusRc.top   = std::min(labelRc.bottom, nameRc.bottom + 1.0f);
                    statusRc.top   = std::min(statusRc.top, statusRc.bottom);
                    statusRc.left  = nameRc.left;
                    statusRc.right = nameRc.right;
                }
            }

            if (! nodeRef.isDirectory && item.lod >= DrawItem::Lod::Large)
            {
                static constexpr float kSizeLineHeightDip = 14.0f;
                const float availableH                    = std::max(0.0f, labelRc.bottom - labelRc.top);
                if (availableH >= 32.0f)
                {
                    showSize      = true;
                    sizeRc.top    = std::max(labelRc.top, labelRc.bottom - kSizeLineHeightDip);
                    nameRc.bottom = std::min(nameRc.bottom, sizeRc.top);
                }
            }

            if (! nameView.empty() && nameRc.right > nameRc.left && nameRc.bottom > nameRc.top)
            {
                _renderTarget->DrawTextW(
                    nameView.data(), static_cast<UINT32>(nameView.size()), _textFormat.get(), nameRc, _brushText.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
            }

            if (showStatusLine && statusRc.right > statusRc.left && statusRc.bottom > statusRc.top && ! watermarkText.empty())
            {
                const D2D1_COLOR_F old = _brushText->GetColor();
                const float alpha      = highContrast ? 1.0f : 0.72f;
                _brushText->SetColor(D2D1::ColorF(old.r, old.g, old.b, old.a * alpha));

                _renderTarget->DrawTextW(watermarkText.data(),
                                         static_cast<UINT32>(watermarkText.size()),
                                         _textFormat.get(),
                                         statusRc,
                                         _brushText.get(),
                                         D2D1_DRAW_TEXT_OPTIONS_CLIP);

                _brushText->SetColor(old);
            }

            if (showSize && sizeRc.right > sizeRc.left && sizeRc.bottom > sizeRc.top)
            {
                const CompactBytesText sizeText = FormatBytesCompactInline(nodeRef.totalBytes);
                if (sizeText.length > 0)
                {
                    _renderTarget->DrawTextW(sizeText.buffer.data(), sizeText.length, _textFormat.get(), sizeRc, _brushText.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                }
            }

            if (showSpinner)
            {
                const float w              = std::max(0.0f, spinnerRc.right - spinnerRc.left);
                const float h              = std::max(0.0f, spinnerRc.bottom - spinnerRc.top);
                const float radius         = std::clamp(std::min(w, h) * 0.34f, 3.0f, 7.0f);
                const D2D1_POINT_2F center = D2D1::Point2F((spinnerRc.left + spinnerRc.right) * 0.5f, (spinnerRc.top + spinnerRc.bottom) * 0.5f);
                drawMiniSpinner(center, radius, nodeRef.id, kMiniSpinnerPhaseSpeed, kMiniSpinnerHueSpeed);
            }
        }
    }

    PaintHoverOverlay(accentColor);

    if (scanActive && _brushBackground && _brushAccent)
    {
        struct SpinnerCandidate final
        {
            uint32_t nodeId = 0;
            D2D1_RECT_F rc{};
            float area = 0.0f;
        };

        static constexpr size_t kMaxBigSpinners      = 48;
        static constexpr float kMinBigSpinnerSideDip = 36.0f;

        std::array<SpinnerCandidate, kMaxBigSpinners> spinners{};
        size_t spinnerCount = 0;

        auto tryInsertSpinner = [&](uint32_t nodeId, const D2D1_RECT_F& rc, float area) noexcept
        {
            if (spinnerCount < spinners.size())
            {
                SpinnerCandidate cand;
                cand.nodeId            = nodeId;
                cand.rc                = rc;
                cand.area              = area;
                spinners[spinnerCount] = cand;
                spinnerCount += 1;
                return;
            }

            size_t smallest    = 0;
            float smallestArea = spinners[0].area;
            for (size_t i = 1; i < spinners.size(); ++i)
            {
                if (spinners[i].area < smallestArea)
                {
                    smallestArea = spinners[i].area;
                    smallest     = i;
                }
            }

            if (area <= smallestArea)
            {
                return;
            }

            spinners[smallest].nodeId = nodeId;
            spinners[smallest].rc     = rc;
            spinners[smallest].area   = area;
        };

        for (const DrawItem& item : _drawItems)
        {
            if (item.depth != 0)
            {
                continue;
            }

            const std::optional<ResolvedItem> node = resolveItem(item.nodeId);
            if (! node.has_value())
            {
                continue;
            }

            const ResolvedItem& nodeRef = node.value();
            if (! nodeRef.isDirectory || nodeRef.isSynthetic)
            {
                continue;
            }

            if (nodeRef.scanState != ScanState::Scanning)
            {
                continue;
            }

            const D2D1_RECT_F rc = rectForItem(item);
            const float w        = std::max(0.0f, rc.right - rc.left);
            const float h        = std::max(0.0f, rc.bottom - rc.top);
            const float side     = std::min(w, h);
            if (side < kMinBigSpinnerSideDip)
            {
                continue;
            }

            const float area = w * h;
            if (area <= 1.0f)
            {
                continue;
            }

            tryInsertSpinner(nodeRef.id, rc, area);
        }

        std::sort(spinners.begin(), spinners.begin() + static_cast<ptrdiff_t>(spinnerCount), [](const SpinnerCandidate& a, const SpinnerCandidate& b) noexcept {
            return a.area > b.area;
        });

        const size_t maxWanted = std::clamp<size_t>(static_cast<size_t>(std::clamp(_config.scanThreads, 1u, 16u)), 1u, kMaxBigSpinners);

        bool drewSpinner = false;
        for (size_t i = 0; i < std::min(spinnerCount, maxWanted); ++i)
        {
            const SpinnerCandidate& cand = spinners[i];
            const float w                = std::max(0.0f, cand.rc.right - cand.rc.left);
            const float h                = std::max(0.0f, cand.rc.bottom - cand.rc.top);
            const float side             = std::min(w, h);
            if (side < kMinBigSpinnerSideDip)
            {
                continue;
            }

            const D2D1_POINT_2F center = D2D1::Point2F((cand.rc.left + cand.rc.right) * 0.5f, (cand.rc.top + cand.rc.bottom) * 0.5f);
            const float radius         = std::clamp(side * 0.12f, 10.0f, 42.0f);
            const float scrimAlpha     = highContrast ? 0.86f : (darkMode ? 0.76f : 0.62f);
            _brushBackground->SetColor(D2D1::ColorF(bg.r, bg.g, bg.b, scrimAlpha));
            _renderTarget->FillEllipse(D2D1::Ellipse(center, radius * 1.55f, radius * 1.55f), _brushBackground.get());

            drawMiniSpinner(center, radius, cand.nodeId ^ 0x9e3779b9u, kBigSpinnerPhaseSpeed, kBigSpinnerHueSpeed);
            drewSpinner = true;
        }

        if (! drewSpinner)
        {
            const D2D1_RECT_F spinnerHostRc = D2D1::RectF(kPaddingDip,
                                                          headerBottomDip + kPaddingDip,
                                                          std::max(kPaddingDip, treemapRc.right - kPaddingDip),
                                                          std::max(headerBottomDip + kPaddingDip, treemapRc.bottom - kPaddingDip));

            const float hostW    = std::max(0.0f, spinnerHostRc.right - spinnerHostRc.left);
            const float hostH    = std::max(0.0f, spinnerHostRc.bottom - spinnerHostRc.top);
            const float hostSide = std::min(hostW, hostH);
            if (hostSide >= 26.0f)
            {
                const D2D1_POINT_2F center =
                    D2D1::Point2F((spinnerHostRc.left + spinnerHostRc.right) * 0.5f, (spinnerHostRc.top + spinnerHostRc.bottom) * 0.5f);

                const float radius     = std::clamp(hostSide * 0.10f, 12.0f, 42.0f);
                const float scrimAlpha = highContrast ? 0.86f : (darkMode ? 0.76f : 0.62f);
                _brushBackground->SetColor(D2D1::ColorF(bg.r, bg.g, bg.b, scrimAlpha));
                _renderTarget->FillEllipse(D2D1::Ellipse(center, radius * 1.55f, radius * 1.55f), _brushBackground.get());

                const uint32_t seed = _scanGeneration.load(std::memory_order_acquire);
                drawMiniSpinner(center, radius, seed ^ 0x9e3779b9u, kBigSpinnerPhaseSpeed, kBigSpinnerHueSpeed);

                if (_brushText && _headerInfoFormat && (! _headerCountsText.empty() || ! _headerSizeText.empty()))
                {
                    const D2D1_COLOR_F old = _brushText->GetColor();
                    const float alpha      = highContrast ? 0.95f : (darkMode ? 0.82f : 0.78f);
                    _brushText->SetColor(D2D1::ColorF(old.r, old.g, old.b, old.a * alpha));

                    auto drawCenteredLine = [&](std::wstring_view text, float top, float height) noexcept
                    {
                        if (text.empty() || ! _dwriteFactory)
                        {
                            return;
                        }

                        const float w     = MeasureTextWidthDip(_dwriteFactory.get(), _headerInfoFormat.get(), text);
                        const float maxW  = std::max(0.0f, spinnerHostRc.right - spinnerHostRc.left);
                        const float lineW = std::clamp(w, 0.0f, maxW);

                        D2D1_RECT_F rc{};
                        rc.left   = std::clamp(center.x - lineW * 0.5f, spinnerHostRc.left, spinnerHostRc.right);
                        rc.right  = std::clamp(rc.left + lineW, spinnerHostRc.left, spinnerHostRc.right);
                        rc.top    = top;
                        rc.bottom = top + height;

                        _renderTarget->DrawTextW(
                            text.data(), static_cast<UINT32>(text.size()), _headerInfoFormat.get(), rc, _brushText.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
                    };

                    const float lineHeight = std::clamp(radius * 0.85f, 14.0f, 18.0f);
                    float textTop          = center.y + radius * 2.05f;
                    if (textTop + lineHeight <= spinnerHostRc.bottom)
                    {
                        drawCenteredLine(_headerCountsText, textTop, lineHeight);
                        textTop += lineHeight;
                        if (textTop + lineHeight <= spinnerHostRc.bottom)
                        {
                            drawCenteredLine(_headerSizeText, textTop, lineHeight);
                        }
                    }

                    _brushText->SetColor(old);
                }
            }
        }
    }

    if (! scanActive && _overallState == ScanState::Done && _scanCompletedSinceSeconds > 0.0 && _watermarkFormat && _brushText && _brushBackground)
    {
        const double elapsed = nowSeconds - _scanCompletedSinceSeconds;
        if (elapsed >= 0.0 && elapsed < kScanCompletedOverlaySeconds)
        {
            const float t    = static_cast<float>(elapsed / kScanCompletedOverlaySeconds);
            const float fade = 1.0f - std::clamp(t, 0.0f, 1.0f);

            const std::wstring overlayText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_OVERLAY_SCAN_COMPLETED);
            if (! overlayText.empty() && _dwriteFactory)
            {
                const float textW = MeasureTextWidthDip(_dwriteFactory.get(), _watermarkFormat.get(), overlayText);
                const float maxW  = std::max(0.0f, treemapRc.right - treemapRc.left);
                const float boxW  = std::clamp(textW + 28.0f, 120.0f, maxW);

                const D2D1_POINT_2F center = D2D1::Point2F((treemapRc.left + treemapRc.right) * 0.5f, (treemapRc.top + treemapRc.bottom) * 0.5f);

                D2D1_RECT_F boxRc{};
                boxRc.left   = center.x - boxW * 0.5f;
                boxRc.right  = center.x + boxW * 0.5f;
                boxRc.top    = center.y - 22.0f;
                boxRc.bottom = center.y + 22.0f;

                const float scrimAlpha = (highContrast ? 0.88f : (darkMode ? 0.74f : 0.62f)) * fade;
                _brushBackground->SetColor(D2D1::ColorF(bg.r, bg.g, bg.b, scrimAlpha));
                _renderTarget->FillRoundedRectangle(D2D1::RoundedRect(boxRc, 10.0f, 10.0f), _brushBackground.get());

                const D2D1_COLOR_F oldTextColor = _brushText->GetColor();
                const float textAlpha           = (highContrast ? 1.0f : 0.92f) * fade;
                _brushText->SetColor(D2D1::ColorF(oldTextColor.r, oldTextColor.g, oldTextColor.b, oldTextColor.a * textAlpha));

                _renderTarget->DrawTextW(
                    overlayText.c_str(), static_cast<UINT32>(overlayText.size()), _watermarkFormat.get(), boxRc, _brushText.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);

                _brushText->SetColor(oldTextColor);
            }
        }
    }

    PaintTooltipOverlay();

    const bool shouldRecordStaticCache = CanRecordStaticTreemapCache(scanActive, completionOverlayActive, nowSeconds);
    const bool staticCacheWasEmpty     = _staticTreemapCache.bitmap == nullptr;

    endDraw.reset();

    if (SUCCEEDED(drawHr) && shouldRecordStaticCache)
    {
        if (staticCacheWasEmpty)
        {
            _staticTreemapCache.misses += 1u;
        }

        const HRESULT cacheHr = RecordStaticTreemapCacheFromCurrentTarget();
        if (FAILED(cacheHr))
        {
            Debug::Warning(L"ViewerSpace: static bitmap cache record failed: 0x{:08X}", static_cast<unsigned long>(cacheHr));
        }
    }

    HRESULT presentHr = S_OK;
    if (SUCCEEDED(drawHr) && _renderer.swapChain)
    {
        presentHr = _renderer.swapChain->Present(1u, 0u);
    }
#ifdef ENABLE_TESTS
    applyForcedRendererFault(drawHr, presentHr);
#endif
    const HRESULT renderHr = FAILED(drawHr) ? drawHr : presentHr;

    const uint64_t paintUs = Debug::Perf::ElapsedUs(paintStartedAt);
#ifdef ENABLE_TESTS
    _lastPaintUs = paintUs;
#endif
    Debug::Perf::Emit(
        L"viewer.space.render.paint_us", L"device-context", paintUs, static_cast<uint64_t>(_drawItems.size()), _scanActive.load() ? 1u : 0u, renderHr);
    Debug::Perf::Emit(L"viewer.space.render.tile_draw_count", L"lod", 0u, tileDrawCount, culledTileCount, renderHr);
    Debug::Perf::Emit(L"viewer.space.render.text_draw_count", L"lod", 0u, textDrawCandidateCount, tileDrawCount, renderHr);
    Debug::Perf::Emit(L"viewer.space.layout.visible_tiles", L"lod", 0u, tileDrawCount, culledTileCount, renderHr);

    if (! _openToFirstPaintEmitted && _openStartedAt != std::chrono::steady_clock::time_point{})
    {
        _openToFirstPaintEmitted = true;
        Debug::Perf::Emit(L"viewer.space.open_to_first_paint_us",
                          L"phase0-current",
                          Debug::Perf::ElapsedUs(_openStartedAt),
                          static_cast<uint64_t>(_drawItems.size()),
                          _scanActive.load() ? 1u : 0u,
                          renderHr);
    }

    if (drawHr == D2DERR_RECREATE_TARGET || presentHr == DXGI_ERROR_DEVICE_REMOVED || presentHr == DXGI_ERROR_DEVICE_RESET)
    {
#ifdef ENABLE_TESTS
        _rendererFailureStage = kRendererDiscardDeviceLost;
        _rendererFailureHr    = static_cast<uint32_t>(FAILED(drawHr) ? drawHr : presentHr);
#endif
        DiscardDirect2D();
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
        }
    }
    else if (FAILED(drawHr))
    {
        Debug::Warning(L"ViewerSpace: EndDraw failed: 0x{:08X}", static_cast<unsigned long>(drawHr));
    }
    else if (FAILED(presentHr))
    {
        Debug::Warning(L"ViewerSpace: Present failed: 0x{:08X}", static_cast<unsigned long>(presentHr));
    }
}

void ViewerSpace::OnCommand([[maybe_unused]] HWND hwnd, UINT commandId) noexcept
{
    switch (commandId)
    {
        case IDM_VIEWERSPACE_FILE_REFRESH: RefreshCurrent(); break;
        case IDM_VIEWERSPACE_NAV_UP: NavigateUp(); break;
        case IDM_VIEWERSPACE_FILE_EXIT: static_cast<void>(Close()); break;
    }
}

void ViewerSpace::OnKeyDown(WPARAM vk, bool alt) noexcept
{
    if (vk == VK_ESCAPE)
    {
        if (_overallState == ScanState::Queued || _overallState == ScanState::Scanning)
        {
            CancelScanByUser();
            return;
        }

        if (_hWnd)
        {
            PostMessageW(_hWnd.get(), WM_CLOSE, 0, 0);
        }
        return;
    }

    if (vk == VK_F5)
    {
        RefreshCurrent();
        return;
    }

    if (vk == VK_BACK || (alt && vk == VK_UP))
    {
        NavigateUp();
        return;
    }
}

void ViewerSpace::OnMouseMove(int x, int y) noexcept
{
    if (! _hWnd)
    {
        return;
    }

#if defined(ENABLE_TESTS)
    _debugHasLastMouseMoveClientPoint = true;
    _debugLastMouseMoveClientPoint    = POINT{x, y};
#endif

    if (! _trackingMouse)
    {
        TRACKMOUSEEVENT tme{};
        tme.cbSize    = sizeof(tme);
        tme.dwFlags   = TME_LEAVE;
        tme.hwndTrack = _hWnd.get();
        TrackMouseEvent(&tme);
        _trackingMouse = true;
    }

    const float xDip         = DipFromPx(x);
    const float yDip         = DipFromPx(y);
    const float headerTop    = GetHeaderTopDip();
    const float headerBottom = GetHeaderBottomDip();

    HeaderHit newHeaderHit = HeaderHit::None;
    if (yDip >= headerTop && yDip <= headerBottom)
    {
        if (xDip <= kHeaderButtonWidthDip && CanNavigateUp())
        {
            newHeaderHit = HeaderHit::Up;
        }
        else if ((_overallState == ScanState::Queued || _overallState == ScanState::Scanning) && xDip >= DipFromPx(_clientSize.cx) - kHeaderButtonWidthDip)
        {
            newHeaderHit = HeaderHit::Cancel;
        }
    }

    if (newHeaderHit != _hoverHeaderHit)
    {
        _hoverHeaderHit = newHeaderHit;
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }

    const std::optional<uint32_t> hit = HitTestTreemap(xDip, yDip);
    const uint32_t newHover           = hit.value_or(0);
    if (newHover != _hoverNodeId)
    {
        _hoverNodeId = newHover;
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }

    const uint32_t tooltipNodeId = yDip >= headerBottom ? newHover : 0u;

    if (tooltipNodeId == 0)
    {
        _tooltipCandidateNodeId       = 0;
        _tooltipCandidateSinceSeconds = 0.0;
        UpdateTooltipForHit(0);
        return;
    }

    if (tooltipNodeId == _tooltipNodeId)
    {
        _tooltipCandidateNodeId       = tooltipNodeId;
        _tooltipCandidateSinceSeconds = 0.0;
        UpdateTooltipPosition(x, y);
        return;
    }

    if (_tooltipNodeId != 0)
    {
        UpdateTooltipForHit(0);
    }

    const double nowSeconds = NowSeconds();
    if (tooltipNodeId != _tooltipCandidateNodeId)
    {
        _tooltipCandidateNodeId       = tooltipNodeId;
        _tooltipCandidateSinceSeconds = nowSeconds;
        return;
    }

    static constexpr double kTooltipHoverStabilityDelaySeconds = 0.12;
    if (_tooltipCandidateSinceSeconds <= 0.0)
    {
        _tooltipCandidateSinceSeconds = nowSeconds;
        return;
    }

    if ((nowSeconds - _tooltipCandidateSinceSeconds) >= kTooltipHoverStabilityDelaySeconds)
    {
        UpdateTooltipForHit(tooltipNodeId);
        UpdateTooltipPosition(x, y);
    }
}

void ViewerSpace::OnMouseLeave() noexcept
{
    _trackingMouse                = false;
    _hoverHeaderHit               = HeaderHit::None;
    _tooltipCandidateNodeId       = 0;
    _tooltipCandidateSinceSeconds = 0.0;
    UpdateTooltipForHit(0);
    if (_hoverNodeId != 0 && _hWnd)
    {
        _hoverNodeId = 0;
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void ViewerSpace::OnLButtonDown(int x, int y) noexcept
{
    const float xDip         = DipFromPx(x);
    const float yDip         = DipFromPx(y);
    const float headerTop    = GetHeaderTopDip();
    const float headerBottom = GetHeaderBottomDip();
    if (yDip >= headerTop && yDip <= headerBottom)
    {
        if (xDip <= kHeaderButtonWidthDip && CanNavigateUp())
        {
            NavigateUp();
        }
        else if ((_overallState == ScanState::Queued || _overallState == ScanState::Scanning) && xDip >= DipFromPx(_clientSize.cx) - kHeaderButtonWidthDip)
        {
            CancelScanByUser();
        }
        return;
    }

    const std::optional<uint32_t> hit = HitTestTreemap(xDip, yDip);
    if (! hit.has_value())
    {
        return;
    }

    const uint32_t nodeId = hit.value();
    const std::optional<ResolvedItem> node = ResolveTreemapItem(nodeId);
    if (node.has_value() && node->isDirectory && ! node->isSynthetic && ! node->isFileRecord)
    {
        NavigateTo(nodeId);
    }
}

void ViewerSpace::OnLButtonDblClk(int x, int y) noexcept
{
    const float xDip         = DipFromPx(x);
    const float yDip         = DipFromPx(y);
    const float headerTop    = GetHeaderTopDip();
    const float headerBottom = GetHeaderBottomDip();
    if (yDip >= headerTop && yDip <= headerBottom)
    {
        OnLButtonDown(x, y);
        return;
    }

    const std::optional<uint32_t> hit = HitTestTreemap(xDip, yDip);
    if (! hit.has_value())
    {
        return;
    }

    const uint32_t nodeId = hit.value();

    const Node* realNode = TryGetRealNode(nodeId);
    const bool isRealNode = realNode != nullptr;
    const std::optional<ResolvedItem> node = ResolveTreemapItem(nodeId);
    if (! node.has_value())
    {
        return;
    }

    if (node->isDirectory && ! node->isSynthetic && ! node->isFileRecord)
    {
        NavigateTo(nodeId);
        return;
    }

    if (! node->isSynthetic)
    {
        return;
    }

    const uint32_t parentId = node->parentId;
    if (parentId == 0)
    {
        return;
    }

    if (! isRealNode)
    {
        uint32_t currentLimit = static_cast<uint32_t>(kDefaultLayoutItems);
        const auto limitIt    = _layoutMaxItemsByNode.find(parentId);
        if (limitIt != _layoutMaxItemsByNode.end())
        {
            currentLimit = limitIt->second;
        }

        uint32_t nextLimit = currentLimit;
        if (nextLimit < kMaxLayoutItems)
        {
            nextLimit = nextLimit == 0 ? static_cast<uint32_t>(kDefaultLayoutItems) :
                                          std::min(static_cast<uint32_t>(kMaxLayoutItems), nextLimit * 2u);
        }

        if (nextLimit != currentLimit)
        {
            _layoutMaxItemsByNode[parentId] = nextLimit;
        }

        if (parentId != _viewNodeId)
        {
            NavigateTo(parentId);
            return;
        }

        _layoutDirty = true;
        if (_hWnd)
        {
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
        }
        return;
    }

    const std::wstring parentPath = BuildNodePathText(parentId);
    if (parentPath.empty())
    {
        return;
    }

    if (isRealNode && (_config.topFilesPerDirectory < kAdaptiveFileCandidateCeiling))
    {
        uint32_t nextTopK = _config.topFilesPerDirectory;
        if (nextTopK == 0)
        {
            nextTopK = 256;
        }
        else
        {
            nextTopK = std::min(kAdaptiveFileCandidateCeiling, nextTopK * 2u);
        }

        if (nextTopK != _config.topFilesPerDirectory)
        {
            const std::string cfg = std::format("{{\"topFilesPerDirectory\":{},\"scanThreads\":{},\"maxConcurrentScansPerVolume\":{},\"cacheEnabled\":\"{}\","
                                                "\"cacheTtlSeconds\":{},\"cacheMaxEntries\":{}}}",
                                                nextTopK,
                                                _config.scanThreads,
                                                _config.maxConcurrentScansPerVolume,
                                                _config.cacheEnabled ? "1" : "0",
                                                _config.cacheTtlSeconds,
                                                _config.cacheMaxEntries);
            static_cast<void>(SetConfiguration(cfg.c_str()));
        }

        StartScan(parentPath);
        return;
    }

    if (_fileSystemIsWin32)
    {
        ShellExecuteW(_hWnd.get(), L"open", parentPath.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
    }
}

void ViewerSpace::OnContextMenu(HWND hwnd, POINT screenPt) noexcept
{
#if defined(ENABLE_TESTS)
    _debugHasLastContextMenuScreenPoint = true;
    _debugLastContextMenuScreenPoint    = screenPt;
    _debugLastContextMenuHitNodeId      = 0;
#endif

    if (! hwnd)
    {
        return;
    }

    if (! _menuHandle)
    {
        _menuHandle.reset(Localization::LoadMenuResource(g_hInstance, IDR_VIEWERSPACE_MENU));
    }

    HMENU viewerMenu = _menuHandle ? _menuHandle.get() : GetMenu(hwnd);
    if (viewerMenu)
    {
        UpdateMenuState(hwnd, false);
    }

    static constexpr std::array<int, 1> kPreviewContextMenuExcludedCommandIds{{
        IDM_VIEWERSPACE_FILE_EXIT,
    }};
    RedSalamander::DxUi::NativeMenuFlyoutOptions previewMenuOptions{};
    previewMenuOptions.includeAcceleratorText = false;
    previewMenuOptions.omitEmptySubmenus      = true;
    previewMenuOptions.trimSeparators         = true;
    previewMenuOptions.excludedCommandIds     = kPreviewContextMenuExcludedCommandIds;

    POINT clientPt = screenPt;
    if (ScreenToClient(hwnd, &clientPt) == 0)
    {
        return;
    }

    const float xDip = DipFromPx(clientPt.x);
    const float yDip = DipFromPx(clientPt.y);

    const std::optional<uint32_t> hit = HitTestTreemap(xDip, yDip);
#if defined(ENABLE_TESTS)
    _debugLastContextMenuHitNodeId = hit.value_or(0);
#endif
    if (! hit.has_value())
    {
        HMENU menu = viewerMenu;
        if (! menu)
        {
            return;
        }

        const auto result = RedSalamander::DxUi::ShowNativeHMenuContextMenu(hwnd,
                                                                            screenPt,
                                                                            menu,
                                                                            _hasTheme ? RedSalamander::DxUi::MakeThemePaletteFromViewerTheme(_theme)
                                                                                      : RedSalamander::DxUi::MakeDefaultThemePalette(false),
                                                                            previewMenuOptions);
        if (result.has_value() && result.value() > 0)
        {
            OnCommand(hwnd, static_cast<UINT>(result.value()));
        }
        return;
    }

    const uint32_t nodeId = hit.value();

    const std::optional<ResolvedItem> node = ResolveTreemapItem(nodeId);
    if (! node.has_value())
    {
        return;
    }

    const std::wstring focusText   = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_CONTEXT_FOCUS_IN_PANE);
    const std::wstring zoomInText  = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_CONTEXT_ZOOM_IN);
    const std::wstring zoomOutText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_CONTEXT_ZOOM_OUT);

    HMENU rootMenu = Localization::LoadMenuResource(GetModuleHandleW(nullptr), kHostFolderViewContextMenuResourceId);
    if (! rootMenu)
    {
        return;
    }

    auto menuCleanup = wil::scope_exit([&]() noexcept { DestroyMenu(rootMenu); });

    HMENU menu = GetSubMenu(rootMenu, 0);
    if (! menu)
    {
        return;
    }

    // Remove View Space + any debug-only items.
    static_cast<void>(DeleteMenu(menu, kCmdFolderViewContextViewSpace, MF_BYCOMMAND));

    auto menuContainsDebugCommands = [&](HMENU m, auto&& self) noexcept -> bool
    {
        if (! m)
        {
            return false;
        }

        const int count = GetMenuItemCount(m);
        if (count <= 0)
        {
            return false;
        }

        for (int pos = 0; pos < count; ++pos)
        {
            const UINT id = GetMenuItemID(m, pos);
            if (id != static_cast<UINT>(-1) && id >= kFolderViewDebugCommandIdBase)
            {
                return true;
            }

            HMENU sub = GetSubMenu(m, pos);
            if (sub && self(sub, self))
            {
                return true;
            }
        }

        return false;
    };

    const int topCount = GetMenuItemCount(menu);
    if (topCount > 0)
    {
        for (int pos = 0; pos < topCount; ++pos)
        {
            HMENU sub = GetSubMenu(menu, pos);
            if (! sub)
            {
                continue;
            }

            if (! menuContainsDebugCommands(sub, menuContainsDebugCommands))
            {
                continue;
            }

            // Remove the debug popup and any separator directly above it.
            RemoveMenu(menu, static_cast<UINT>(pos), MF_BYPOSITION);

            const int sepPos = pos - 1;
            if (sepPos >= 0)
            {
                MENUITEMINFOW info{};
                info.cbSize = sizeof(info);
                info.fMask  = MIIM_FTYPE;
                if (GetMenuItemInfoW(menu, static_cast<UINT>(sepPos), TRUE, &info) != 0 && (info.fType & MFT_SEPARATOR) != 0)
                {
                    RemoveMenu(menu, static_cast<UINT>(sepPos), MF_BYPOSITION);
                }
            }
            break;
        }
    }

    // Insert our treemap commands at the top (reverse order).
    InsertMenuW(menu, 0, MF_BYPOSITION | MF_SEPARATOR, 0, nullptr);
    InsertMenuW(menu, 0, MF_BYPOSITION | MF_STRING, kCmdTreemapContextZoomOut, zoomOutText.c_str());
    InsertMenuW(menu, 0, MF_BYPOSITION | MF_STRING, kCmdTreemapContextZoomIn, zoomInText.c_str());
    InsertMenuW(menu, 0, MF_BYPOSITION | MF_STRING, kCmdTreemapContextFocusInPane, focusText.c_str());

    // Resolve host navigation target.
    std::wstring folderPath;
    std::wstring focusItemDisplayName;
    std::wstring folderPathForCommand;
    std::wstring focusItemForCommand;

    const std::wstring nodePath = BuildItemPathText(nodeId);
    const bool nodeHasPath      = ! nodePath.empty();

    if (nodeHasPath)
    {
        if (_fileSystemIsWin32)
        {
            const std::filesystem::path full(nodePath);
            const std::wstring leaf = full.filename().wstring();
            if (const std::optional<std::filesystem::path> parent = TryGetParentPathForNavigation(full); parent.has_value() && ! leaf.empty())
            {
                folderPathForCommand = parent.value().wstring();
                focusItemForCommand  = leaf;
            }
            else if (node->isDirectory)
            {
                folderPath = nodePath;
            }
        }
        else
        {
            std::wstring_view trimmed = TrimTrailingPathSeparators(nodePath);
            const size_t lastSep      = trimmed.find_last_of(L"/\\");
            if (lastSep != std::wstring_view::npos && (lastSep + 1) < trimmed.size())
            {
                const std::wstring_view leaf = trimmed.substr(lastSep + 1);
                if (const std::optional<std::wstring> parent = TryGetParentPathForNavigationGeneric(trimmed); parent.has_value())
                {
                    folderPathForCommand = parent.value();
                    focusItemForCommand.assign(leaf.data(), leaf.size());
                }
                else if (node->isDirectory)
                {
                    folderPath = nodePath;
                }
            }
            else if (node->isDirectory)
            {
                folderPath = nodePath;
            }
        }

        if (folderPath.empty())
        {
            folderPath           = folderPathForCommand;
            focusItemDisplayName = focusItemForCommand;
        }
    }

    const bool canZoomIn  = node->isDirectory && ! node->isSynthetic && ! node->isFileRecord;
    const bool canZoomOut = CanNavigateUp();

    const bool hostAvailable      = static_cast<bool>(_hostPaneExecute);
    const bool canFocusInPane     = hostAvailable && ! folderPath.empty();
    const bool canExecuteLeafCmds = hostAvailable && ! folderPathForCommand.empty() && ! focusItemForCommand.empty();
    const bool canExecutePaste    = hostAvailable && ! folderPath.empty();

    EnableMenuItem(menu, kCmdTreemapContextZoomIn, static_cast<UINT>(MF_BYCOMMAND | (canZoomIn ? MF_ENABLED : MF_GRAYED)));
    EnableMenuItem(menu, kCmdTreemapContextZoomOut, static_cast<UINT>(MF_BYCOMMAND | (canZoomOut ? MF_ENABLED : MF_GRAYED)));
    EnableMenuItem(menu, kCmdTreemapContextFocusInPane, static_cast<UINT>(MF_BYCOMMAND | (canFocusInPane ? MF_ENABLED : MF_GRAYED)));

    const bool isDirectory = node->isDirectory;

    auto enableFolderCmd = [&](UINT commandId, bool enabled) noexcept
    { EnableMenuItem(menu, commandId, static_cast<UINT>(MF_BYCOMMAND | (enabled ? MF_ENABLED : MF_GRAYED))); };

    enableFolderCmd(kCmdFolderViewContextOpen, canExecuteLeafCmds);
    enableFolderCmd(kCmdFolderViewContextOpenWith, canExecuteLeafCmds && ! isDirectory);
    enableFolderCmd(kCmdFolderViewContextDelete, canExecuteLeafCmds);
    enableFolderCmd(kCmdFolderViewContextMove, canExecuteLeafCmds);
    enableFolderCmd(kCmdFolderViewContextRename, canExecuteLeafCmds);
    enableFolderCmd(kCmdFolderViewContextCopy, canExecuteLeafCmds);
    enableFolderCmd(kCmdFolderViewContextProperties, canExecuteLeafCmds);
    enableFolderCmd(kCmdFolderViewContextPaste, canExecutePaste);

    std::vector<RedSalamander::DxUi::MenuFlyoutItem> popupItems = ConvertHMenuToDxFlyoutItems(menu);
    if (viewerMenu)
    {
        std::vector<RedSalamander::DxUi::MenuFlyoutItem> viewerItems = RedSalamander::DxUi::ConvertNativeHMenuToFlyoutItems(viewerMenu, previewMenuOptions);
        if (! viewerItems.empty())
        {
            if (! popupItems.empty())
            {
                RedSalamander::DxUi::MenuFlyoutItem separator{};
                separator.kind = RedSalamander::DxUi::MenuItemKind::Separator;
                popupItems.push_back(std::move(separator));
            }

            for (auto& viewerItem : viewerItems)
            {
                popupItems.push_back(std::move(viewerItem));
            }
        }
    }

    if (popupItems.empty())
    {
        return;
    }

    const auto result = RedSalamander::DxUi::ContextMenu::Show(hwnd,
                                                               screenPt,
                                                               popupItems,
                                                               _hasTheme ? RedSalamander::DxUi::MakeThemePaletteFromViewerTheme(_theme)
                                                                         : RedSalamander::DxUi::MakeDefaultThemePalette(false));
    if (! result.has_value())
    {
        return;
    }
    const UINT commandId = static_cast<UINT>(result.value());

    if (commandId == kCmdTreemapContextZoomIn)
    {
        if (canZoomIn)
        {
            NavigateTo(nodeId);
        }
        return;
    }

    if (commandId == kCmdTreemapContextZoomOut)
    {
        if (canZoomOut)
        {
            NavigateUp();
        }
        return;
    }

    if (commandId == kCmdTreemapContextFocusInPane)
    {
        if (! canFocusInPane)
        {
            return;
        }

        HostPaneExecuteRequest request{};
        request.version              = 1;
        request.sizeBytes            = sizeof(request);
        request.flags                = HOST_PANE_EXECUTE_FLAG_ACTIVATE_WINDOW;
        request.folderPath           = folderPath.c_str();
        request.focusItemDisplayName = focusItemDisplayName.empty() ? nullptr : focusItemDisplayName.c_str();
        request.folderViewCommandId  = 0;

        static_cast<void>(_hostPaneExecute->ExecuteInActivePane(&request));
        return;
    }

    switch (commandId)
    {
        case IDM_VIEWERSPACE_FILE_REFRESH:
        case IDM_VIEWERSPACE_NAV_UP:
        case IDM_VIEWERSPACE_FILE_EXIT: OnCommand(hwnd, commandId); return;
        default: break;
    }

    if (! canExecutePaste && commandId == kCmdFolderViewContextPaste)
    {
        return;
    }

    if (! canExecuteLeafCmds && commandId != kCmdFolderViewContextPaste)
    {
        return;
    }

    if (! _hostPaneExecute)
    {
        return;
    }

    HostPaneExecuteRequest request{};
    request.version              = 1;
    request.sizeBytes            = sizeof(request);
    request.flags                = HOST_PANE_EXECUTE_FLAG_ACTIVATE_WINDOW;
    request.folderPath           = (commandId == kCmdFolderViewContextPaste) ? folderPath.c_str() : folderPathForCommand.c_str();
    request.focusItemDisplayName = (commandId == kCmdFolderViewContextPaste || focusItemForCommand.empty()) ? nullptr : focusItemForCommand.c_str();
    request.folderViewCommandId  = commandId;

    static_cast<void>(_hostPaneExecute->ExecuteInActivePane(&request));
}

LRESULT ViewerSpace::OnNcDestroy(HWND hwnd, WPARAM wp, LPARAM lp) noexcept
{
    _menuBarHost.Detach();
    _menuHandle.reset();
    _embeddedMode = false;
    _hWnd.release();
    SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);

    const LRESULT result = DefWindowProcW(hwnd, WM_NCDESTROY, wp, lp);
    const uint32_t previousWindowCount = g_viewerSpaceWindowCount.load(std::memory_order_acquire);
    if (previousWindowCount != 0 &&
        g_viewerSpaceWindowCount.fetch_sub(1, std::memory_order_acq_rel) == 1u)
    {
        UnregisterWndClassIfIdle();
    }

    Release();
    return result;
}

void ViewerSpace::OnTimer(UINT_PTR timerId) noexcept
{
    if (timerId != kTimerAnimationId || ! _hWnd)
    {
        return;
    }

    ReapFinishedScanWorkers(false);
    DrainUpdates();
    MaybeRebuildLayout();

    const double now = NowSeconds();

    if (_scanActive.load() || _overallState == ScanState::Queued || _overallState == ScanState::Scanning)
    {
        static constexpr double kMinInvalidateIntervalSecondsWhileScanning = 1.0 / 30.0;

        const double sinceLast = now - _lastScanInvalidateSeconds;
        if (_lastScanInvalidateSeconds <= 0.0 || sinceLast >= kMinInvalidateIntervalSecondsWhileScanning)
        {
            _lastScanInvalidateSeconds = now;
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
        }
        return;
    }

    if (! _layoutDirty)
    {
        ContinueScanCacheBuild();
    }

    if (_scanCompletedSinceSeconds > 0.0)
    {
        const double elapsed = now - _scanCompletedSinceSeconds;
        if (elapsed < kScanCompletedOverlaySeconds)
        {
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
            return;
        }

        _scanCompletedSinceSeconds = 0.0;
    }

    bool animating = false;
    for (const auto& item : _drawItems)
    {
        if (now - item.animationStartSeconds < kAnimationDurationSeconds)
        {
            animating = true;
            break;
        }
    }

    if (animating)
    {
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

bool ViewerSpace::EnsureDirect2D(HWND hwnd) noexcept
{
    RECT rc{};
    GetClientRect(hwnd, &rc);
    const UINT32 width  = static_cast<UINT32>(std::max<LONG>(1, rc.right - rc.left));
    const UINT32 height = static_cast<UINT32>(std::max<LONG>(1, rc.bottom - rc.top));

    if (! _d2dFactory)
    {
        D2D1_FACTORY_OPTIONS options{};
        const HRESULT hr =
            D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, __uuidof(ID2D1Factory1), &options, reinterpret_cast<void**>(_d2dFactory.addressof()));
        if (FAILED(hr))
        {
            Debug::Warning(L"ViewerSpace: Failed to create D2D factory: 0x{:08X}", static_cast<unsigned long>(hr));
            return false;
        }
    }

    if (! _dwriteFactory)
    {
        const HRESULT hr = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(_dwriteFactory.put()));
        if (FAILED(hr))
        {
            return false;
        }
    }

    if (! _otherStrokeStyle && _d2dFactory)
    {
        D2D1_STROKE_STYLE_PROPERTIES props = D2D1::StrokeStyleProperties();
        props.startCap                     = D2D1_CAP_STYLE_ROUND;
        props.endCap                       = D2D1_CAP_STYLE_ROUND;
        props.dashCap                      = D2D1_CAP_STYLE_ROUND;
        props.lineJoin                     = D2D1_LINE_JOIN_ROUND;
        props.dashStyle                    = D2D1_DASH_STYLE_DASH;
        if (FAILED(_d2dFactory->CreateStrokeStyle(props, nullptr, 0, _otherStrokeStyle.put())))
        {
            Debug::Warning(L"ViewerSpace: Failed to create stroke style");
        }
    }

    if (! _dogEarFlapGeometry && _d2dFactory)
    {
        wil::com_ptr<ID2D1PathGeometry> geom;
        if (SUCCEEDED(_d2dFactory->CreatePathGeometry(geom.put())) && geom)
        {
            wil::com_ptr<ID2D1GeometrySink> sink;
            if (SUCCEEDED(geom->Open(sink.put())) && sink)
            {
                sink->BeginFigure(D2D1::Point2F(0.0f, 0.0f), D2D1_FIGURE_BEGIN_FILLED);
                const D2D1_POINT_2F points[] = {D2D1::Point2F(0.0f, 1.0f), D2D1::Point2F(1.0f, 1.0f)};
                sink->AddLines(points, static_cast<UINT32>(std::size(points)));
                sink->EndFigure(D2D1_FIGURE_END_CLOSED);

                if (SUCCEEDED(sink->Close()))
                {
                    _dogEarFlapGeometry = std::move(geom);
                }
            }
        }
    }

    if (_renderTarget && _renderer.swapChainWidthPx == width && _renderer.swapChainHeightPx == height)
    {
        return true;
    }
#ifdef ENABLE_TESTS
    auto rememberRendererFailure = [&](uint32_t stage, HRESULT hr) noexcept
    {
        _rendererFailureStage = stage;
        _rendererFailureHr    = static_cast<uint32_t>(hr);
    };
    auto clearRendererFailure = [&]() noexcept
    {
        _rendererFailureStage = 0u;
        _rendererFailureHr    = 0u;
    };
#endif

    if (! _renderer.d3dDevice || ! _renderer.d3dContext)
    {
        UINT creationFlags = D3D11_CREATE_DEVICE_BGRA_SUPPORT;
#if defined(_DEBUG)
        creationFlags |= D3D11_CREATE_DEVICE_DEBUG;
#endif
        D3D_FEATURE_LEVEL featureLevels[] = {D3D_FEATURE_LEVEL_11_1, D3D_FEATURE_LEVEL_11_0, D3D_FEATURE_LEVEL_10_1};
        HRESULT d3dHr                     = D3D11CreateDevice(nullptr,
                                         D3D_DRIVER_TYPE_HARDWARE,
                                         nullptr,
                                         creationFlags,
                                         featureLevels,
                                         static_cast<UINT>(std::size(featureLevels)),
                                         D3D11_SDK_VERSION,
                                         _renderer.d3dDevice.addressof(),
                                         &_renderer.featureLevel,
                                         _renderer.d3dContext.addressof());
        if (FAILED(d3dHr))
        {
            Debug::Warning(L"ViewerSpace: Hardware D3D11CreateDevice failed, falling back to WARP: 0x{:08X}", static_cast<unsigned long>(d3dHr));
            d3dHr = D3D11CreateDevice(nullptr,
                                      D3D_DRIVER_TYPE_WARP,
                                      nullptr,
                                      creationFlags,
                                      featureLevels,
                                      static_cast<UINT>(std::size(featureLevels)),
                                      D3D11_SDK_VERSION,
                                      _renderer.d3dDevice.addressof(),
                                      &_renderer.featureLevel,
                                      _renderer.d3dContext.addressof());
        }
        if (FAILED(d3dHr) || ! _renderer.d3dDevice || ! _renderer.d3dContext)
        {
            Debug::Warning(L"ViewerSpace: D3D11CreateDevice failed: 0x{:08X}", static_cast<unsigned long>(d3dHr));
#ifdef ENABLE_TESTS
            rememberRendererFailure(kRendererFailureD3DDevice, d3dHr);
#endif
            return false;
        }

        wil::com_ptr<IDXGIDevice> dxgiDevice;
        wil::com_ptr<IDXGIAdapter> adapter;
        const HRESULT dxgiHr        = _renderer.d3dDevice->QueryInterface(IID_PPV_ARGS(dxgiDevice.addressof()));
        const HRESULT adapterHr     = SUCCEEDED(dxgiHr) && dxgiDevice ? dxgiDevice->GetAdapter(adapter.addressof()) : E_FAIL;
        const HRESULT factoryHr     = SUCCEEDED(adapterHr) && adapter ? adapter->GetParent(IID_PPV_ARGS(_renderer.dxgiFactory.addressof())) : E_FAIL;
        const HRESULT d2dDeviceHr   = SUCCEEDED(dxgiHr) && dxgiDevice ? _d2dFactory->CreateDevice(dxgiDevice.get(), _renderer.d2dDevice.addressof()) : E_FAIL;
        const HRESULT d2dContextHr  = SUCCEEDED(d2dDeviceHr) && _renderer.d2dDevice
                                          ? _renderer.d2dDevice->CreateDeviceContext(D2D1_DEVICE_CONTEXT_OPTIONS_NONE, _renderer.d2dContext.addressof())
                                          : E_FAIL;
        if (FAILED(dxgiHr) || ! dxgiDevice || FAILED(adapterHr) || ! adapter || FAILED(factoryHr) || ! _renderer.dxgiFactory || FAILED(d2dDeviceHr) ||
            ! _renderer.d2dDevice || FAILED(d2dContextHr) || ! _renderer.d2dContext)
        {
            Debug::Warning(L"ViewerSpace: Device-context renderer creation failed ({:08X}, {:08X}, {:08X}, {:08X}, {:08X})",
                           static_cast<unsigned long>(dxgiHr),
                           static_cast<unsigned long>(adapterHr),
                           static_cast<unsigned long>(factoryHr),
                           static_cast<unsigned long>(d2dDeviceHr),
                           static_cast<unsigned long>(d2dContextHr));
#ifdef ENABLE_TESTS
            rememberRendererFailure(kRendererFailureD2DContext, FAILED(d2dContextHr) ? d2dContextHr : E_FAIL);
#endif
            DiscardDirect2D();
            return false;
        }

        _renderer.d2dContext->SetUnitMode(D2D1_UNIT_MODE_DIPS);
#ifdef ENABLE_TESTS
        _rendererDeviceCreateCount += 1u;
#endif
    }

    if (! _renderer.swapChain)
    {
        DXGI_SWAP_CHAIN_DESC1 desc{};
        desc.Width            = width;
        desc.Height           = height;
        desc.Format           = DXGI_FORMAT_B8G8R8A8_UNORM;
        desc.SampleDesc.Count = 1;
        desc.BufferUsage      = DXGI_USAGE_RENDER_TARGET_OUTPUT;
        desc.BufferCount      = 2;
        desc.Scaling          = DXGI_SCALING_STRETCH;
        desc.SwapEffect       = DXGI_SWAP_EFFECT_FLIP_SEQUENTIAL;
        desc.AlphaMode        = DXGI_ALPHA_MODE_IGNORE;

        const HRESULT swapHr = _renderer.dxgiFactory ? _renderer.dxgiFactory->CreateSwapChainForHwnd(
                                                           _renderer.d3dDevice.get(), hwnd, &desc, nullptr, nullptr, _renderer.swapChain.addressof())
                                                     : E_FAIL;
        if (FAILED(swapHr) || ! _renderer.swapChain)
        {
            Debug::Warning(L"ViewerSpace: CreateSwapChainForHwnd failed: 0x{:08X}", static_cast<unsigned long>(swapHr));
#ifdef ENABLE_TESTS
            rememberRendererFailure(kRendererFailureSwapChain, swapHr);
#endif
            DiscardDirect2D();
            return false;
        }
        _renderer.swapChainWidthPx  = width;
        _renderer.swapChainHeightPx = height;
        static_cast<void>(_renderer.dxgiFactory->MakeWindowAssociation(hwnd, DXGI_MWA_NO_ALT_ENTER | DXGI_MWA_NO_WINDOW_CHANGES));
    }

    const bool sizeChanged = _renderer.swapChainWidthPx != width || _renderer.swapChainHeightPx != height;
    if (sizeChanged)
    {
        InvalidateStaticTreemapCache();
        if (_renderer.d2dContext)
        {
            _renderer.d2dContext->SetTarget(nullptr);
            static_cast<void>(_renderer.d2dContext->Flush());
        }
        _renderTarget.reset();
        _renderer.targetBitmap.reset();
        _headerPathSourceText.clear();
        _headerPathDisplayText.clear();
        _headerPathDisplayMaxWidthDip = 0.0f;
        if (_renderer.d3dContext)
        {
            _renderer.d3dContext->ClearState();
            _renderer.d3dContext->Flush();
        }

        const HRESULT resizeHr = _renderer.swapChain->ResizeBuffers(0u, width, height, DXGI_FORMAT_UNKNOWN, 0u);
        if (FAILED(resizeHr))
        {
            Debug::Warning(L"ViewerSpace: ResizeBuffers failed: 0x{:08X}", static_cast<unsigned long>(resizeHr));
#ifdef ENABLE_TESTS
            rememberRendererFailure(kRendererFailureResizeBuffers, resizeHr);
#endif
            DiscardDirect2D();
            return false;
        }
        _renderer.swapChainWidthPx  = width;
        _renderer.swapChainHeightPx = height;
#ifdef ENABLE_TESTS
        _swapChainResizeCount += 1u;
#endif
    }

    if (! _renderer.targetBitmap)
    {
        wil::com_ptr<IDXGISurface> surface;
        const HRESULT bufferHr = _renderer.swapChain ? _renderer.swapChain->GetBuffer(0u, IID_PPV_ARGS(surface.addressof())) : E_FAIL;
        if (FAILED(bufferHr) || ! surface)
        {
            Debug::Warning(L"ViewerSpace: IDXGISwapChain1::GetBuffer failed: 0x{:08X}", static_cast<unsigned long>(bufferHr));
#ifdef ENABLE_TESTS
            rememberRendererFailure(kRendererFailureGetBuffer, bufferHr);
#endif
            DiscardDirect2D();
            return false;
        }

        const D2D1_BITMAP_PROPERTIES1 props =
            D2D1::BitmapProperties1(D2D1_BITMAP_OPTIONS_TARGET | D2D1_BITMAP_OPTIONS_CANNOT_DRAW,
                                    D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_IGNORE),
                                    _dpi,
                                    _dpi);
        const HRESULT bitmapHr = _renderer.d2dContext->CreateBitmapFromDxgiSurface(surface.get(), &props, _renderer.targetBitmap.addressof());
        if (FAILED(bitmapHr) || ! _renderer.targetBitmap)
        {
            Debug::Warning(L"ViewerSpace: CreateBitmapFromDxgiSurface failed: 0x{:08X}", static_cast<unsigned long>(bitmapHr));
#ifdef ENABLE_TESTS
            rememberRendererFailure(kRendererFailureTargetBitmap, bitmapHr);
#endif
            DiscardDirect2D();
            return false;
        }
        _renderer.d2dContext->SetTarget(_renderer.targetBitmap.get());
    }

    _renderer.d2dContext->SetDpi(_dpi, _dpi);
    _renderer.d2dContext->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);
    const HRESULT renderTargetHr = _renderTarget ? S_OK : _renderer.d2dContext->QueryInterface(IID_PPV_ARGS(_renderTarget.addressof()));
    if (FAILED(renderTargetHr) || ! _renderTarget)
    {
        Debug::Warning(L"ViewerSpace: Device context did not expose ID2D1RenderTarget.");
#ifdef ENABLE_TESTS
        rememberRendererFailure(kRendererFailureRenderTarget, renderTargetHr);
#endif
        DiscardDirect2D();
        return false;
    }

    const D2D1::ColorF bg  = _hasTheme ? ColorFFromArgb(_theme.backgroundArgb) : D2D1::ColorF(D2D1::ColorF::White);
    const D2D1::ColorF txt = _hasTheme ? ColorFFromArgb(_theme.textArgb) : D2D1::ColorF(D2D1::ColorF::Black);
    const D2D1::ColorF acc = _hasTheme ? ColorFFromArgb(_theme.accentArgb) : D2D1::ColorF(D2D1::ColorF::CornflowerBlue);

    if (! _brushBackground)
    {
        if (FAILED(_renderTarget->CreateSolidColorBrush(bg, _brushBackground.put())))
        {
            Debug::Warning(L"ViewerSpace: Failed to create background brush");
            return false;
        }
#ifdef ENABLE_TESTS
        _rendererBrushCreateCount += 1u;
#endif
    }
    if (! _brushText)
    {
        if (FAILED(_renderTarget->CreateSolidColorBrush(txt, _brushText.put())))
        {
            Debug::Warning(L"ViewerSpace: Failed to create text brush");
            return false;
        }
#ifdef ENABLE_TESTS
        _rendererBrushCreateCount += 1u;
#endif
    }
    if (! _brushOutline)
    {
        if (FAILED(_renderTarget->CreateSolidColorBrush(txt, _brushOutline.put())))
        {
            Debug::Warning(L"ViewerSpace: Failed to create outline brush");
            return false;
        }
#ifdef ENABLE_TESTS
        _rendererBrushCreateCount += 1u;
#endif
    }
    if (! _brushAccent)
    {
        if (FAILED(_renderTarget->CreateSolidColorBrush(acc, _brushAccent.put())))
        {
            Debug::Warning(L"ViewerSpace: Failed to create accent brush");
            return false;
        }
#ifdef ENABLE_TESTS
        _rendererBrushCreateCount += 1u;
#endif
    }
    if (! _brushWatermark)
    {
        if (FAILED(_renderTarget->CreateSolidColorBrush(txt, _brushWatermark.put())))
        {
            Debug::Warning(L"ViewerSpace: Failed to create watermark brush");
            return false;
        }
#ifdef ENABLE_TESTS
        _rendererBrushCreateCount += 1u;
#endif
    }

    const bool darkMode     = _hasTheme && _theme.darkMode != FALSE;
    const bool highContrast = _hasTheme && _theme.highContrast != FALSE;
    const bool rainbowMode  = _hasTheme && _theme.rainbowMode != FALSE;

    float highlightAlpha = darkMode ? 0.14f : 0.08f;
    float shadowAlpha    = darkMode ? 0.22f : 0.14f;
    if (rainbowMode)
    {
        highlightAlpha *= 0.45f;
        shadowAlpha *= 0.45f;
    }
    if (highContrast)
    {
        highlightAlpha = 0.0f;
        shadowAlpha    = 0.0f;
    }

    std::array<D2D1_GRADIENT_STOP, 3> stops{};
    stops[0].position = 0.0f;
    stops[0].color    = D2D1::ColorF(1.0f, 1.0f, 1.0f, highlightAlpha);
    stops[1].position = 0.55f;
    stops[1].color    = D2D1::ColorF(1.0f, 1.0f, 1.0f, 0.0f);
    stops[2].position = 1.0f;
    stops[2].color    = D2D1::ColorF(0.0f, 0.0f, 0.0f, shadowAlpha);

    if (! _shadingStops && FAILED(_renderTarget->CreateGradientStopCollection(stops.data(), static_cast<UINT32>(stops.size()), _shadingStops.put())))
    {
        Debug::Warning(L"ViewerSpace: Failed to create gradient stop collection");
        return false;
    }
    if (! _brushShading)
    {
        if (FAILED(_renderTarget->CreateLinearGradientBrush(
                D2D1::LinearGradientBrushProperties(D2D1::Point2F(0, 0), D2D1::Point2F(1, 1)), _shadingStops.get(), _brushShading.put())))
        {
            Debug::Warning(L"ViewerSpace: Failed to create shading brush");
            return false;
        }
#ifdef ENABLE_TESTS
        _rendererBrushCreateCount += 1u;
#endif
    }

    HRESULT textHr = S_OK;
    if (! _textFormat)
    {
        textHr = Typography::CreateTextFormat(_dwriteFactory.get(), Typography::MakeUiTextSpec(12.0f), _textFormat.put());
        if (SUCCEEDED(textHr))
        {
#ifdef ENABLE_TESTS
            _rendererTextFormatCreateCount += 1u;
#endif
            _textFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
            _textFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR);
        }
    }

    if (! _headerFormat)
    {
        textHr = Typography::CreateTextFormat(_dwriteFactory.get(), Typography::MakeUiTextSpec(13.0f, DWRITE_FONT_WEIGHT_SEMI_BOLD), _headerFormat.put());
        if (SUCCEEDED(textHr))
        {
#ifdef ENABLE_TESTS
            _rendererTextFormatCreateCount += 1u;
#endif
            _headerFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
            _headerFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
            _headerFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        }
    }

    if (! _headerStatusFormatRight)
    {
        textHr =
            Typography::CreateTextFormat(_dwriteFactory.get(), Typography::MakeUiTextSpec(13.0f, DWRITE_FONT_WEIGHT_SEMI_BOLD), _headerStatusFormatRight.put());
        if (SUCCEEDED(textHr))
        {
#ifdef ENABLE_TESTS
            _rendererTextFormatCreateCount += 1u;
#endif
            _headerStatusFormatRight->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
            _headerStatusFormatRight->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
            _headerStatusFormatRight->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_TRAILING);
        }
    }

    if (! _headerInfoFormat)
    {
        textHr = Typography::CreateTextFormat(_dwriteFactory.get(), Typography::MakeUiTextSpec(11.0f), _headerInfoFormat.put());
        if (SUCCEEDED(textHr))
        {
#ifdef ENABLE_TESTS
            _rendererTextFormatCreateCount += 1u;
#endif
            _headerInfoFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
            _headerInfoFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
            _headerInfoFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        }
    }

    if (! _headerInfoFormatRight)
    {
        textHr = Typography::CreateTextFormat(_dwriteFactory.get(), Typography::MakeUiTextSpec(11.0f), _headerInfoFormatRight.put());
        if (SUCCEEDED(textHr))
        {
#ifdef ENABLE_TESTS
            _rendererTextFormatCreateCount += 1u;
#endif
            _headerInfoFormatRight->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
            _headerInfoFormatRight->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
            _headerInfoFormatRight->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_TRAILING);
        }
    }

    if (! _headerIconFormat)
    {
        textHr = Typography::CreateTextFormat(_dwriteFactory.get(), Typography::MakeUiTextSpec(18.0f, DWRITE_FONT_WEIGHT_SEMI_BOLD), _headerIconFormat.put());
        if (SUCCEEDED(textHr))
        {
#ifdef ENABLE_TESTS
            _rendererTextFormatCreateCount += 1u;
#endif
            _headerIconFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
            _headerIconFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
            _headerIconFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
        }
    }

    if (! _watermarkFormat)
    {
        textHr = Typography::CreateTextFormat(_dwriteFactory.get(), Typography::MakeUiTextSpec(18.0f, DWRITE_FONT_WEIGHT_SEMI_BOLD), _watermarkFormat.put());
        if (SUCCEEDED(textHr))
        {
#ifdef ENABLE_TESTS
            _rendererTextFormatCreateCount += 1u;
#endif
            _watermarkFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
            _watermarkFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
            _watermarkFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_CENTER);
        }
    }

    if (! _tooltipFormat)
    {
        textHr = Typography::CreateTextFormat(_dwriteFactory.get(), Typography::MakeUiTextSpec(12.0f), _tooltipFormat.put());
        if (SUCCEEDED(textHr))
        {
#ifdef ENABLE_TESTS
            _rendererTextFormatCreateCount += 1u;
#endif
            _tooltipFormat->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP);
            _tooltipFormat->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR);
            _tooltipFormat->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING);
        }
    }

    auto applyEllipsisTrimming = [&](IDWriteTextFormat* format) noexcept
    {
        if (format == nullptr || ! _dwriteFactory)
        {
            return;
        }

        DWRITE_TRIMMING trimming{};
        trimming.granularity = DWRITE_TRIMMING_GRANULARITY_CHARACTER;

        wil::com_ptr<IDWriteInlineObject> ellipsis;
        if (SUCCEEDED(_dwriteFactory->CreateEllipsisTrimmingSign(format, ellipsis.put())))
        {
            format->SetTrimming(&trimming, ellipsis.get());
        }
    };

    applyEllipsisTrimming(_textFormat.get());
    applyEllipsisTrimming(_headerFormat.get());
    applyEllipsisTrimming(_headerStatusFormatRight.get());
    applyEllipsisTrimming(_headerInfoFormat.get());
    applyEllipsisTrimming(_headerInfoFormatRight.get());
    applyEllipsisTrimming(_watermarkFormat.get());

#ifdef ENABLE_TESTS
    clearRendererFailure();
#endif
    return true;
}

void ViewerSpace::DiscardDirect2D() noexcept
{
#ifdef ENABLE_TESTS
    if (_rendererFailureStage == 0u)
    {
        _rendererFailureStage = kRendererDiscardUnknown;
        _rendererFailureHr    = 0u;
    }
#endif
    InvalidateStaticTreemapCache();
    if (_renderer.d2dContext)
    {
        _renderer.d2dContext->SetTarget(nullptr);
    }
    _renderTarget.reset();
    _renderer.targetBitmap.reset();
    _renderer.swapChain.reset();
    _renderer.d2dContext.reset();
    _renderer.d2dDevice.reset();
    _renderer.dxgiFactory.reset();
    _renderer.d3dContext.reset();
    _renderer.d3dDevice.reset();
    _renderer.swapChainWidthPx  = 0;
    _renderer.swapChainHeightPx = 0;
    _brushBackground.reset();
    _brushText.reset();
    _brushOutline.reset();
    _brushAccent.reset();
    _brushWatermark.reset();
    _brushShading.reset();
    _shadingStops.reset();
    _textFormat.reset();
    _headerFormat.reset();
    _headerStatusFormatRight.reset();
    _headerInfoFormat.reset();
    _headerInfoFormatRight.reset();
    _headerIconFormat.reset();
    _watermarkFormat.reset();
    _tooltipFormat.reset();
    _headerPathSourceText.clear();
    _headerPathDisplayText.clear();
    _headerPathDisplayMaxWidthDip = 0.0f;
}

void ViewerSpace::InvalidateStaticTreemapCache() noexcept
{
    _staticTreemapCache.bitmap.reset();
    _staticTreemapCache.widthPx = 0u;
    _staticTreemapCache.heightPx = 0u;
    _staticTreemapCache.bytes = 0u;
    _staticTreemapCache.generation += 1u;
}

bool ViewerSpace::HasActiveLayoutAnimation(double nowSeconds) const noexcept
{
    static constexpr float kSettledRectToleranceDip = 0.25f;
    for (const DrawItem& item : _drawItems)
    {
        if ((nowSeconds - item.animationStartSeconds) < kAnimationDurationSeconds)
        {
            return true;
        }

        if (std::fabs(item.currentRect.left - item.targetRect.left) > kSettledRectToleranceDip ||
            std::fabs(item.currentRect.top - item.targetRect.top) > kSettledRectToleranceDip ||
            std::fabs(item.currentRect.right - item.targetRect.right) > kSettledRectToleranceDip ||
            std::fabs(item.currentRect.bottom - item.targetRect.bottom) > kSettledRectToleranceDip)
        {
            return true;
        }
    }

    return false;
}

bool ViewerSpace::CanReplayStaticTreemapCache(bool scanActive, bool completionOverlayActive) const noexcept
{
    if (! _staticTreemapCache.bitmap || ! _renderTarget)
    {
        return false;
    }

    if (_staticTreemapCache.widthPx != _renderer.swapChainWidthPx || _staticTreemapCache.heightPx != _renderer.swapChainHeightPx)
    {
        return false;
    }

    if (_layoutDirty || scanActive || completionOverlayActive || _hoverHeaderHit != HeaderHit::None)
    {
        return false;
    }

    return true;
}

bool ViewerSpace::CanRecordStaticTreemapCache(bool scanActive, bool completionOverlayActive, double nowSeconds) const noexcept
{
    if (! _renderTarget || _renderer.swapChainWidthPx == 0u || _renderer.swapChainHeightPx == 0u)
    {
        return false;
    }

    if (_layoutDirty || scanActive || completionOverlayActive || _hoverHeaderHit != HeaderHit::None || _hoverNodeId != 0 || _tooltipNodeId != 0)
    {
        return false;
    }

    return ! HasActiveLayoutAnimation(nowSeconds);
}

HRESULT ViewerSpace::EnsureStaticTreemapCacheBitmap() noexcept
{
    if (! _renderer.d2dContext)
    {
        return E_FAIL;
    }

    const UINT width = _renderer.swapChainWidthPx;
    const UINT height = _renderer.swapChainHeightPx;
    if (width == 0u || height == 0u)
    {
        return E_INVALIDARG;
    }

    if (_staticTreemapCache.bitmap && _staticTreemapCache.widthPx == width && _staticTreemapCache.heightPx == height)
    {
        return S_OK;
    }

    _staticTreemapCache.bitmap.reset();

    D2D1_BITMAP_PROPERTIES1 props = D2D1::BitmapProperties1(
        D2D1_BITMAP_OPTIONS_NONE,
        D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_IGNORE),
        _dpi,
        _dpi);

    const HRESULT hr = _renderer.d2dContext->CreateBitmap(D2D1::SizeU(width, height), nullptr, 0u, &props, _staticTreemapCache.bitmap.put());
    if (FAILED(hr))
    {
        _staticTreemapCache.widthPx = 0u;
        _staticTreemapCache.heightPx = 0u;
        _staticTreemapCache.bytes = 0u;
        return hr;
    }

    _staticTreemapCache.widthPx = width;
    _staticTreemapCache.heightPx = height;
    _staticTreemapCache.bytes = static_cast<uint64_t>(width) * static_cast<uint64_t>(height) * 4u;
    return S_OK;
}

HRESULT ViewerSpace::RecordStaticTreemapCacheFromCurrentTarget() noexcept
{
    const auto recordStartedAt = std::chrono::steady_clock::now();
    HRESULT hr = EnsureStaticTreemapCacheBitmap();
    if (FAILED(hr))
    {
        return hr;
    }

    const D2D1_RECT_U sourceRect = D2D1::RectU(0u, 0u, _staticTreemapCache.widthPx, _staticTreemapCache.heightPx);
    hr = _staticTreemapCache.bitmap->CopyFromRenderTarget(nullptr, _renderTarget.get(), &sourceRect);
    if (FAILED(hr))
    {
        _staticTreemapCache.bitmap.reset();
        _staticTreemapCache.bytes = 0u;
        return hr;
    }

    _staticTreemapCache.recordCount += 1u;
    _staticTreemapCache.lastRecordUs = Debug::Perf::ElapsedUs(recordStartedAt);
    Debug::Perf::Emit(L"viewer.space.render.static_cache_record_us",
                      L"bitmap",
                      _staticTreemapCache.lastRecordUs,
                      static_cast<uint64_t>(_drawItems.size()),
                      _staticTreemapCache.bytes,
                      S_OK);
    return S_OK;
}

void ViewerSpace::PaintHoverOverlay(const D2D1::ColorF& accentColor) noexcept
{
    if (_hoverNodeId == 0 || ! _brushOutline || ! _renderTarget)
    {
        return;
    }

    for (const auto& item : _drawItems)
    {
        if (item.nodeId != _hoverNodeId)
        {
            continue;
        }

        float gap = kItemGapDip - static_cast<float>(item.depth) * 0.15f;
        gap       = std::clamp(gap, 0.5f, kItemGapDip);

        D2D1_RECT_F rc = item.currentRect;
        rc.left += gap;
        rc.top += gap;
        rc.right -= gap;
        rc.bottom -= gap;
        if (RectArea(rc) < 1.0f)
        {
            break;
        }

        _brushOutline->SetColor(D2D1::ColorF(accentColor.r, accentColor.g, accentColor.b, 1.0f));
        const std::optional<ResolvedItem> node = ResolveTreemapItem(item.nodeId);
        const bool isFileTile = node.has_value() && ! node->isDirectory && ! node->isSynthetic;
        if (isFileTile)
        {
            const float tileW   = std::max(0.0f, rc.right - rc.left);
            const float tileH   = std::max(0.0f, rc.bottom - rc.top);
            const float side    = std::min(tileW, tileH);
            const float dogSize = std::clamp(side * 0.18f, 8.0f, 14.0f);
            float cut           = dogSize + 2.0f;
            cut                 = std::clamp(cut, 6.0f, std::max(6.0f, side - 1.0f));

            if (cut > 1.0f && tileW > cut + 1.0f && tileH > cut + 1.0f)
            {
                const D2D1_POINT_2F tl = D2D1::Point2F(rc.left, rc.top);
                const D2D1_POINT_2F br = D2D1::Point2F(rc.right, rc.bottom);
                const D2D1_POINT_2F bl = D2D1::Point2F(rc.left, rc.bottom);

                const D2D1_POINT_2F cutA = D2D1::Point2F(rc.right - cut, rc.top);
                const D2D1_POINT_2F cutB = D2D1::Point2F(rc.right, rc.top + cut);

                _renderTarget->DrawLine(tl, cutA, _brushOutline.get(), 2.25f, nullptr);
                _renderTarget->DrawLine(cutA, cutB, _brushOutline.get(), 2.25f, nullptr);
                _renderTarget->DrawLine(cutB, br, _brushOutline.get(), 2.25f, nullptr);
                _renderTarget->DrawLine(br, bl, _brushOutline.get(), 2.25f, nullptr);
                _renderTarget->DrawLine(bl, tl, _brushOutline.get(), 2.25f, nullptr);
            }
            else
            {
                _renderTarget->DrawRectangle(rc, _brushOutline.get(), 2.25f);
            }
        }
        else
        {
            _renderTarget->DrawRectangle(rc, _brushOutline.get(), 2.25f);
        }
        break;
    }
}

void ViewerSpace::ApplyThemeToWindow(HWND hwnd) noexcept
{
    if (! _hasTheme)
    {
        return;
    }

    const bool dark = _theme.darkMode != FALSE;
    MessageBoxCenteringDetail::ApplyImmersiveDarkMode(hwnd, dark);

    const bool windowActive = GetActiveWindow() == hwnd;
    ApplyTitleBarTheme(hwnd, windowActive);
    ApplyMenuTheme(hwnd);
}

void ViewerSpace::ApplyTitleBarTheme(HWND hwnd, bool windowActive) noexcept
{
    if (! _hasTheme || ! hwnd)
    {
        return;
    }

    using DwmSetWindowAttributeFunc          = HRESULT(WINAPI*)(HWND, DWORD, LPCVOID, DWORD);
    static DwmSetWindowAttributeFunc setAttr = []() noexcept -> DwmSetWindowAttributeFunc
    {
        HMODULE dwm = LoadLibraryW(L"dwmapi.dll");
        if (! dwm)
        {
            return nullptr;
        }
#pragma warning(push)
#pragma warning(disable : 4191) // 'reinterpret_cast': unsafe conversion from 'FARPROC'
        return reinterpret_cast<DwmSetWindowAttributeFunc>(GetProcAddress(dwm, "DwmSetWindowAttribute"));
#pragma warning(pop)
    }();

    if (! setAttr)
    {
        return;
    }

    static constexpr DWORD kDwmwaBorderColor  = 34u;
    static constexpr DWORD kDwmwaCaptionColor = 35u;
    static constexpr DWORD kDwmwaTextColor    = 36u;
    static constexpr DWORD kDwmColorDefault   = 0xFFFFFFFFu;

    DWORD borderValue  = kDwmColorDefault;
    DWORD captionValue = kDwmColorDefault;
    DWORD textValue    = kDwmColorDefault;

    if (_theme.highContrast == FALSE && _theme.rainbowMode != FALSE)
    {
        COLORREF accent = ColorRefFromArgb(_theme.accentArgb);
        if (! windowActive)
        {
            static constexpr uint8_t kInactiveTitleBlendAlpha = 223u; // ~7/8 toward background
            const COLORREF bg                                 = ColorRefFromArgb(_theme.backgroundArgb);
            accent                                            = BlendColor(accent, bg, kInactiveTitleBlendAlpha);
        }

        const COLORREF text = ChooseContrastingTextColor(accent);
        borderValue         = static_cast<DWORD>(accent);
        captionValue        = static_cast<DWORD>(accent);
        textValue           = static_cast<DWORD>(text);
    }

    setAttr(hwnd, kDwmwaBorderColor, &borderValue, sizeof(borderValue));
    setAttr(hwnd, kDwmwaCaptionColor, &captionValue, sizeof(captionValue));
    setAttr(hwnd, kDwmwaTextColor, &textValue, sizeof(textValue));
}

void ViewerSpace::UpdateWindowTitle(HWND hwnd) noexcept
{
    std::wstring pathText = _viewPathText;

    std::wstring title;
    if (! pathText.empty())
    {
        title = FormatStringResource(g_hInstance, IDS_VIEWERSPACE_TITLE_FORMAT, pathText);
        if (title.empty())
        {
            title = pathText;
        }
    }
    else
    {
        title = _metaName;
    }

    if (! title.empty())
    {
        SetWindowTextW(hwnd, title.c_str());
    }
}

void ViewerSpace::ApplyMenuTheme(HWND hwnd) noexcept
{
    static_cast<void>(hwnd);
    _menuBarHost.SetTheme(_hasTheme ? RedSalamander::DxUi::MakeThemePaletteFromViewerTheme(_theme) : RedSalamander::DxUi::MakeDefaultThemePalette(false));
    if (_menuBarHost.GetHwnd())
    {
        _menuBarHost.SyncMenuModel();
    }
    else if (hwnd)
    {
        DrawMenuBar(hwnd);
    }
}

void ViewerSpace::UpdateHeaderTextCache() noexcept
{
    UINT statusId = IDS_VIEWERSPACE_STATUS_DONE;
    switch (_overallState)
    {
        case ScanState::NotStarted: statusId = IDS_VIEWERSPACE_STATUS_NOT_STARTED; break;
        case ScanState::Queued: statusId = IDS_VIEWERSPACE_STATUS_QUEUED; break;
        case ScanState::Scanning: statusId = IDS_VIEWERSPACE_STATUS_SCANNING; break;
        case ScanState::Done: statusId = IDS_VIEWERSPACE_STATUS_DONE; break;
        case ScanState::Error: statusId = IDS_VIEWERSPACE_STATUS_ERROR; break;
        case ScanState::Canceled: statusId = IDS_VIEWERSPACE_STATUS_CANCELED; break;
    }

    if (_headerStatusId != statusId)
    {
        _headerStatusId   = statusId;
        _headerStatusText = LoadStringResource(g_hInstance, statusId);
    }

    const bool scanActive = _overallState == ScanState::Queued || _overallState == ScanState::Scanning;

    const uint64_t items = static_cast<uint64_t>(_scanProgressFolders) + static_cast<uint64_t>(_scanProgressFiles);
    if (items > 0 || scanActive)
    {
        _headerCountsText = FormatStringResource(g_hInstance, IDS_VIEWERSPACE_HEADER_COUNTS_FORMAT, items, _scanProgressFolders, _scanProgressFiles);

        const CompactBytesText sizeText = FormatBytesCompactInline(_scanProgressBytes);
        const std::wstring_view sizeView(sizeText.buffer.data(), sizeText.length);

        _headerSizeText = FormatStringResource(g_hInstance, IDS_VIEWERSPACE_HEADER_SIZE_FORMAT, sizeView, _scanProgressBytes);
    }
    else
    {
        _headerCountsText.clear();
        _headerSizeText.clear();
    }

    if (scanActive && ! _scanProcessingFolderName.empty())
    {
        _headerProcessingText = FormatStringResource(g_hInstance, IDS_VIEWERSPACE_HEADER_PROCESSING_FORMAT, _scanProcessingFolderName);
    }
    else
    {
        _headerProcessingText.clear();
    }
}

void ViewerSpace::UpdateTooltipForHit(uint32_t nodeId) noexcept
{
    if (! _hWnd)
    {
        _tooltipNodeId = 0;
        _tooltipText.clear();
        return;
    }

    if (nodeId == 0)
    {
        if (_tooltipNodeId != 0)
        {
            _tooltipNodeId = 0;
            _tooltipText.clear();
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
        }
        return;
    }

    if (nodeId != _tooltipNodeId)
    {
        _tooltipNodeId = nodeId;
        _tooltipText   = BuildTooltipText(nodeId);
        if (_tooltipText.empty())
        {
            _tooltipNodeId = 0;
        }
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void ViewerSpace::UpdateTooltipPosition(int x, int y) noexcept
{
    if (! _hWnd || _tooltipNodeId == 0)
    {
        return;
    }

    const D2D1_POINT_2F newAnchor = D2D1::Point2F(DipFromPx(x) + 14.0f, DipFromPx(y) + 18.0f);
    if (std::abs(newAnchor.x - _tooltipAnchorDip.x) <= 0.5f && std::abs(newAnchor.y - _tooltipAnchorDip.y) <= 0.5f)
    {
        return;
    }

    _tooltipAnchorDip = newAnchor;
    InvalidateRect(_hWnd.get(), nullptr, FALSE);
}

void ViewerSpace::PaintTooltipOverlay() noexcept
{
    if (_tooltipNodeId == 0 || _tooltipText.empty() || ! _renderTarget || ! _dwriteFactory || ! _tooltipFormat || ! _brushBackground || ! _brushText)
    {
        return;
    }

    const float maxWidthDip  = std::max(180.0f, kTooltipMaxWidthDip);
    const float maxHeightDip = kTooltipMaxHeightDip;

    wil::com_ptr<IDWriteTextLayout> layout;
    const HRESULT layoutHr = _dwriteFactory->CreateTextLayout(
        _tooltipText.c_str(), static_cast<UINT32>(_tooltipText.size()), _tooltipFormat.get(), maxWidthDip, maxHeightDip, layout.put());
    if (FAILED(layoutHr) || ! layout)
    {
        return;
    }

    DWRITE_TEXT_METRICS metrics{};
    if (FAILED(layout->GetMetrics(&metrics)))
    {
        return;
    }

    constexpr float kPadXDip      = 10.0f;
    constexpr float kPadYDip      = 7.0f;
    constexpr float kScreenGapDip = 8.0f;
    const float textWidthDip      = std::clamp(metrics.widthIncludingTrailingWhitespace, 48.0f, maxWidthDip);
    const float textHeightDip     = std::clamp(metrics.height, 16.0f, maxHeightDip);
    const float boxWidthDip       = textWidthDip + (kPadXDip * 2.0f);
    const float boxHeightDip      = textHeightDip + (kPadYDip * 2.0f);
    const float clientWidthDip    = DipFromPx(_clientSize.cx);
    const float clientHeightDip   = DipFromPx(_clientSize.cy);

    float left = _tooltipAnchorDip.x;
    float top  = _tooltipAnchorDip.y;
    if (left + boxWidthDip > clientWidthDip - kScreenGapDip)
    {
        left = _tooltipAnchorDip.x - boxWidthDip - 24.0f;
    }
    if (top + boxHeightDip > clientHeightDip - kScreenGapDip)
    {
        top = _tooltipAnchorDip.y - boxHeightDip - 24.0f;
    }
    left = std::clamp(left, kScreenGapDip, std::max(kScreenGapDip, clientWidthDip - boxWidthDip - kScreenGapDip));
    top  = std::clamp(top, kScreenGapDip, std::max(kScreenGapDip, clientHeightDip - boxHeightDip - kScreenGapDip));

    const D2D1_RECT_F boxRc  = D2D1::RectF(left, top, left + boxWidthDip, top + boxHeightDip);
    const D2D1_POINT_2F text = D2D1::Point2F(boxRc.left + kPadXDip, boxRc.top + kPadYDip);

    const bool highContrast = _hasTheme && _theme.highContrast != FALSE;
    const bool darkMode     = _hasTheme && _theme.darkMode != FALSE;
    D2D1::ColorF bg         = _hasTheme ? ColorFFromArgb(_theme.selectionBackgroundArgb)
                                        : D2D1::ColorF(GetRValue(GetSysColor(COLOR_INFOBK)) / 255.0f,
                                                       GetGValue(GetSysColor(COLOR_INFOBK)) / 255.0f,
                                                       GetBValue(GetSysColor(COLOR_INFOBK)) / 255.0f,
                                                       1.0f);
    D2D1::ColorF fg         = _hasTheme ? ColorFFromArgb(_theme.selectionTextArgb)
                                        : D2D1::ColorF(GetRValue(GetSysColor(COLOR_INFOTEXT)) / 255.0f,
                                                       GetGValue(GetSysColor(COLOR_INFOTEXT)) / 255.0f,
                                                       GetBValue(GetSysColor(COLOR_INFOTEXT)) / 255.0f,
                                                       1.0f);

    bg.a = highContrast ? 1.0f : 0.96f;
    fg.a = 1.0f;

    const D2D1_COLOR_F oldBg   = _brushBackground->GetColor();
    const D2D1_COLOR_F oldText = _brushText->GetColor();
    _brushBackground->SetColor(bg);
    _renderTarget->FillRoundedRectangle(D2D1::RoundedRect(boxRc, 4.0f, 4.0f), _brushBackground.get());

    if (_brushOutline)
    {
        const D2D1_COLOR_F oldOutline = _brushOutline->GetColor();
        const float alpha             = highContrast ? 1.0f : (darkMode ? 0.45f : 0.30f);
        _brushOutline->SetColor(D2D1::ColorF(fg.r, fg.g, fg.b, alpha));
        _renderTarget->DrawRoundedRectangle(D2D1::RoundedRect(boxRc, 4.0f, 4.0f), _brushOutline.get(), 1.0f);
        _brushOutline->SetColor(oldOutline);
    }

    _brushText->SetColor(fg);
    _renderTarget->DrawTextLayout(text, layout.get(), _brushText.get(), D2D1_DRAW_TEXT_OPTIONS_CLIP);
#ifdef ENABLE_TESTS
    ++_tooltipPaintCount;
#endif
    _brushText->SetColor(oldText);
    _brushBackground->SetColor(oldBg);
}

std::wstring ViewerSpace::BuildTooltipText(uint32_t nodeId) const
{
    if (nodeId == 0)
    {
        return {};
    }

    const std::optional<ResolvedItem> resolved = ResolveTreemapItem(nodeId);
    if (! resolved.has_value())
    {
        return {};
    }
    const ResolvedItem& node = resolved.value();

    const std::wstring pathText = BuildItemPathText(nodeId);

    std::wstring name(node.name.data(), node.name.size());
    if (name.empty())
    {
        name = pathText;
    }

    const std::wstring sizeText = FormatBytesCompact(node.totalBytes);

    uint64_t viewBytes = 0;
    const Node* view   = TryGetRealNode(_viewNodeId);
    if (view != nullptr)
    {
        viewBytes = view->totalBytes;
    }

    std::wstring shareText = LoadEmbeddedStringResource(g_hInstance, IDS_VIEWERSPACE_TOOLTIP_SHARE_UNKNOWN);
    if (viewBytes > 0)
    {
        const double percent = std::clamp(static_cast<double>(node.totalBytes) * 100.0 / static_cast<double>(viewBytes), 0.0, 100.0);
        shareText            = FormatStringResource(g_hInstance, IDS_VIEWERSPACE_TOOLTIP_PERCENT_FORMAT, percent);
        if (shareText.empty())
        {
            shareText = LoadEmbeddedStringResource(g_hInstance, IDS_VIEWERSPACE_TOOLTIP_SHARE_UNKNOWN);
        }
    }

    if (! pathText.empty())
    {
        if (node.scanState != ScanState::Done)
        {
            std::wstring stateText;
            switch (node.scanState)
            {
                case ScanState::NotStarted: stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_STATUS_NOT_STARTED); break;
                case ScanState::Queued: stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_STATUS_QUEUED); break;
                case ScanState::Scanning: stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_STATUS_SCANNING); break;
                case ScanState::Done: stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_STATUS_DONE); break;
                case ScanState::Error: stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_STATUS_ERROR); break;
                case ScanState::Canceled: stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_STATUS_CANCELED); break;
            }

            return FormatStringResource(g_hInstance, IDS_VIEWERSPACE_TOOLTIP_FORMAT_WITH_PATH, name, pathText, sizeText, shareText, stateText);
        }

        return FormatStringResource(g_hInstance, IDS_VIEWERSPACE_TOOLTIP_FORMAT_WITH_PATH_NO_STATE, name, pathText, sizeText, shareText);
    }

    if (node.scanState != ScanState::Done)
    {
        std::wstring stateText;
        switch (node.scanState)
        {
            case ScanState::NotStarted: stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_STATUS_NOT_STARTED); break;
            case ScanState::Queued: stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_STATUS_QUEUED); break;
            case ScanState::Scanning: stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_STATUS_SCANNING); break;
            case ScanState::Done: stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_STATUS_DONE); break;
            case ScanState::Error: stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_STATUS_ERROR); break;
            case ScanState::Canceled: stateText = LoadStringResource(g_hInstance, IDS_VIEWERSPACE_STATUS_CANCELED); break;
        }

        return FormatStringResource(g_hInstance, IDS_VIEWERSPACE_TOOLTIP_FORMAT_NO_PATH, name, sizeText, shareText, stateText);
    }

    return FormatStringResource(g_hInstance, IDS_VIEWERSPACE_TOOLTIP_FORMAT_NO_PATH_NO_STATE, name, sizeText, shareText);
}

void ViewerSpace::StartScan(std::wstring_view rootPath, bool allowCache)
{
    const uint32_t generation = _scanGeneration.fetch_add(1) + 1;
    CancelScan();
    ReapFinishedScanWorkers(false);
    CancelScanCacheBuild();
    _scanCacheLastStoredGeneration = 0;
    _scanCompletedSinceSeconds     = 0.0;
    _lastScanCacheSnapshotBytes    = 0u;
    _lastRectsByNode.clear();
    _pendingUpdatePostedCount.store(0u, std::memory_order_relaxed);
    _pendingUpdateCoalescedCount.store(0u, std::memory_order_relaxed);
    _pendingUpdateBytes.store(0u, std::memory_order_relaxed);
    InvalidateStaticTreemapCache();
    _staticTreemapCache.hits = 0u;
    _staticTreemapCache.misses = 0u;
    _staticTreemapCache.recordCount = 0u;
    _staticTreemapCache.lastRecordUs = 0u;
    _openStartedAt            = std::chrono::steady_clock::now();
    _openToFirstPaintEmitted  = false;
#ifdef ENABLE_TESTS
    _lastPaintUs              = 0u;
    _lastLayoutUs             = 0u;
    _lastDrainUs              = 0u;
    _lastWorkingSetBytes      = 0u;
    _lastTileDrawCount        = 0u;
    _lastTextDrawCount        = 0u;
    _layoutGeneration         = 0u;
    _lastHitTestUs            = 0u;
    _lastHitTestCandidatesChecked = 0u;
#endif

    std::wstring scanRootPath(rootPath);
    _scanRootPath = scanRootPath;
    _scanRootParentPath.reset();

    const bool rootLooksWin32 = LooksLikeWin32Path(_scanRootPath);
    if (_fileSystemIsWin32 || rootLooksWin32)
    {
        const std::optional<std::filesystem::path> parent = TryGetParentPathForNavigation(std::filesystem::path(_scanRootPath));
        if (parent.has_value())
        {
            _scanRootParentPath = parent.value().wstring();
        }
    }
    else
    {
        _scanRootParentPath = TryGetParentPathForNavigationGeneric(_scanRootPath);
    }

    _syntheticNodes.clear();
    _layoutMaxItemsByNode.clear();
    _autoExpandedOtherByNode.clear();
    _layoutCandidateCache.clear();
    _layoutCandidateCacheHits   = 0u;
    _layoutCandidateCacheMisses = 0u;

    std::destroy_at(&_childrenArena);
    std::destroy_at(&_fileRecords);
    std::destroy_at(&_nodes);
    _nodePool.release();
    _nameArena.release();
    _layoutNameArena.release();
    std::construct_at(&_nodes, &_nodePool);
    std::construct_at(&_fileRecords, &_nodePool);
    std::construct_at(&_childrenArena, &_nodePool);
    _nodes.resize(2u);
    _drawItems.clear();
    _hitGrid.Clear();
    _navStack.clear();
    _layoutDirty         = true;
    _hoverNodeId         = 0;
    _tooltipNodeId       = 0;
    UpdateTooltipForHit(0);

    {
        std::scoped_lock lock(_updateMutex);
        _pendingUpdates.clear();
    }
    _pendingUpdateBytes.store(0u, std::memory_order_relaxed);

    _scanProgressBytes    = 0;
    _scanProgressFolders  = 0;
    _scanProgressFiles    = 0;
    _scanProcessingNodeId = 0;
    _scanProcessingFolderName.clear();
    _headerStatusId = 0;
    _headerStatusText.clear();
    _headerCountsText.clear();
    _headerSizeText.clear();
    _headerProcessingText.clear();

    const uint32_t effectiveTopFilesPerDirectory = ComputeAdaptiveFileCandidateBudget();
    _scanTopFilesPerDirectory = effectiveTopFilesPerDirectory;
    const size_t topFilesPerDirectory = static_cast<size_t>(effectiveTopFilesPerDirectory);
    const uint32_t scanThreads =
        _fileSystemIsWin32 ? std::clamp<uint32_t>(std::max(_config.scanThreads, DefaultWin32ScanThreads()), 1u, 16u) : std::clamp(_config.scanThreads, 1u, 16u);

    if (allowCache && _config.cacheEnabled && _fileSystemIsWin32)
    {
        ScanResultCacheKey cacheKey;
        cacheKey.rootKey              = NormalizeRootPathForScanCache(std::filesystem::path(_scanRootPath));
        cacheKey.topFilesPerDirectory = effectiveTopFilesPerDirectory;

        if (! cacheKey.rootKey.empty())
        {
            const std::shared_ptr<const ScanResultSnapshot> snapshot = GetScanResultCache().TryGet(cacheKey);
            if (snapshot && snapshot->nodes.size() > 1 && snapshot->nodes[1].id == 1)
            {
                Debug::Perf::EmitCounter(L"viewer.space.render.static_cache_hit");
                _nodes.resize(std::max<size_t>(snapshot->nodes.size(), 2u));

                _childrenArena.resize(snapshot->childrenArena.size());
                std::copy(snapshot->childrenArena.begin(), snapshot->childrenArena.end(), _childrenArena.begin());
                _fileRecords.resize(snapshot->fileRecords.size());

                for (size_t i = 0; i < snapshot->nodes.size(); ++i)
                {
                    const ScanResultCacheNode& cachedNode = snapshot->nodes[i];
                    if (cachedNode.id == 0)
                    {
                        continue;
                    }

                    Node node;
                    node.id               = cachedNode.id;
                    node.parentId         = cachedNode.parentId;
                    node.isDirectory      = cachedNode.isDirectory;
                    node.isSynthetic      = cachedNode.isSynthetic;
                    node.scanState        = static_cast<ScanState>(cachedNode.scanState);
                    node.name             = CopyToArena(_nameArena, cachedNode.name);
                    node.totalBytes       = cachedNode.totalBytes;
                    node.childrenStart    = cachedNode.childrenStart;
                    node.childrenCount    = cachedNode.childrenCount;
                    node.childrenCapacity = cachedNode.childrenCapacity;
                    node.aggregateFolders = cachedNode.aggregateFolders;
                    node.aggregateFiles   = cachedNode.aggregateFiles;

                    const size_t index = static_cast<size_t>(cachedNode.id);
                    if (index < _nodes.size())
                    {
                        _nodes[index] = std::move(node);
                    }
                }

                for (size_t i = 0; i < snapshot->fileRecords.size(); ++i)
                {
                    const ScanResultCacheFileRecord& cachedFile = snapshot->fileRecords[i];
                    FileRecord record;
                    record.id       = cachedFile.id;
                    record.parentId = cachedFile.parentId;
                    record.name     = CopyToArena(_nameArena, cachedFile.name);
                    record.bytes    = cachedFile.bytes;

                    if (IsFileRecordId(record.id))
                    {
                        const uint32_t fileIndex = record.id - kFileRecordIdBase;
                        if (fileIndex < _fileRecords.size())
                        {
                            _fileRecords[fileIndex] = std::move(record);
                        }
                    }
                }

                _rootNodeId = 1;
                _viewNodeId = 1;
                UpdateViewPathText();

                const Node* rootCached = TryGetRealNode(_rootNodeId);
                _overallState          = rootCached ? rootCached->scanState : ScanState::Done;
                _scanActive.store(false);
                _animationStartSeconds    = NowSeconds();
                _lastLayoutRebuildSeconds = 0.0;

                if (_hWnd)
                {
                    UpdateWindowTitle(_hWnd.get());
                    UpdateMenuState(_hWnd.get());
                    InvalidateRect(_hWnd.get(), nullptr, FALSE);
                }

                _scanCacheLastStoredGeneration = generation;
                return;
            }
            Debug::Perf::EmitCounter(L"viewer.space.render.static_cache_miss");
        }
    }

    _overallState = ScanState::Queued;
    _scanActive.store(true);
    _animationStartSeconds    = NowSeconds();
    _lastLayoutRebuildSeconds = 0.0;

    Node root;
    root.id          = 1;
    root.parentId    = 0;
    root.isDirectory = true;
    root.scanState   = ScanState::Queued;
    std::wstring rootName;
    if (_fileSystemIsWin32 || rootLooksWin32)
    {
        const std::filesystem::path rootFs(_scanRootPath);
        rootName = rootFs.filename().wstring();
        if (rootName.empty())
        {
            rootName = rootFs.wstring();
        }
    }
    else
    {
        const std::wstring_view trimmed = TrimTrailingPathSeparators(_scanRootPath);
        const size_t lastSep            = trimmed.find_last_of(L"/\\");
        if (lastSep != std::wstring_view::npos && lastSep + 1 < trimmed.size())
        {
            rootName = std::wstring(trimmed.substr(lastSep + 1));
        }
        if (rootName.empty())
        {
            rootName = std::wstring(trimmed);
        }
        if (rootName.empty())
        {
            rootName = L"/";
        }
    }

    root.name = CopyToArena(_nameArena, rootName);
    if (root.name.empty() && ! _scanRootPath.empty())
    {
        root.name = CopyToArena(_nameArena, _scanRootPath);
    }

    _nodes[root.id] = std::move(root);
    _rootNodeId     = 1;
    _viewNodeId     = 1;
    UpdateViewPathText();

    UpdateHeaderTextCache();

    auto done                                = std::make_shared<std::atomic_bool>(false);
    _scanWorker.done                         = done;
    wil::com_ptr<IFileSystem> scanFileSystem = _fileSystem;
    const bool fileSystemIsWin32             = _fileSystemIsWin32;
    _scanWorker.thread =
        std::jthread([this, generation, done, scanFileSystem, fileSystemIsWin32, scanRootPath, topFilesPerDirectory, scanThreads](std::stop_token st) noexcept
    {
        auto markDone = wil::scope_exit([done] { done->store(true); });
        ScanMain(st, generation, scanFileSystem, fileSystemIsWin32, scanRootPath, 1, 2, topFilesPerDirectory, scanThreads);
    });

    if (_hWnd)
    {
        UpdateWindowTitle(_hWnd.get());
        UpdateMenuState(_hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void ViewerSpace::CancelScan() noexcept
{
    if (_scanWorker.thread.joinable())
    {
        _scanWorker.thread.request_stop();
        _retiredScanWorkers.emplace_back(std::move(_scanWorker));
        _scanWorker = {};
    }
    _scanActive.store(false);
}

void ViewerSpace::CancelScanByUser() noexcept
{
    if (! _scanWorker.thread.joinable())
    {
        return;
    }

    _scanGeneration.fetch_add(1);

    CancelScan();
    CancelScanCacheBuild();
    _scanCacheLastStoredGeneration = 0;

    {
        std::scoped_lock lock(_updateMutex);
        _pendingUpdates.clear();
    }

    Node* root = TryGetRealNode(_rootNodeId);
    if (root != nullptr)
    {
        root->scanState = ScanState::Canceled;
    }
    _overallState = ScanState::Canceled;
    _scanActive.store(false);
    UpdateHeaderTextCache();

    if (_hWnd)
    {
        UpdateWindowTitle(_hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void ViewerSpace::CancelScanAndWait() noexcept
{
    if (_scanWorker.thread.joinable())
    {
        _scanWorker.thread.request_stop();
    }

    for (auto& worker : _retiredScanWorkers)
    {
        if (worker.thread.joinable())
        {
            worker.thread.request_stop();
        }
    }

    if (_scanWorker.thread.joinable())
    {
        _scanWorker.thread.join();
    }
    _scanWorker = {};

    for (auto& worker : _retiredScanWorkers)
    {
        if (worker.thread.joinable())
        {
            worker.thread.join();
        }
    }
    _retiredScanWorkers.clear();

    _scanActive.store(false);
}

void ViewerSpace::ReapFinishedScanWorkers(bool wait) noexcept
{
    if (wait)
    {
        CancelScanAndWait();
        return;
    }

    if (_scanWorker.thread.joinable())
    {
        if (_scanWorker.done != nullptr && _scanWorker.done->load())
        {
            _scanWorker.thread.join();
            _scanWorker = {};
        }
    }

    for (size_t i = 0; i < _retiredScanWorkers.size();)
    {
        ScanWorker& worker = _retiredScanWorkers[i];
        if (worker.done != nullptr && worker.done->load() && worker.thread.joinable())
        {
            worker.thread.join();
            _retiredScanWorkers.erase(_retiredScanWorkers.begin() + static_cast<ptrdiff_t>(i));
            continue;
        }
        i += 1;
    }
}

void ViewerSpace::ScanMain(std::stop_token stopToken,
                           uint32_t generation,
                           wil::com_ptr<IFileSystem> fileSystem,
                           bool fileSystemIsWin32,
                           std::wstring rootPath,
                           uint32_t rootNodeId,
                           uint32_t nextNodeId,
                           size_t topFilesPerDirectory,
                           uint32_t scanThreads) noexcept
{
    static constexpr size_t kProgressUpdateStride = 384;
    static constexpr auto kProgressUpdateInterval = std::chrono::milliseconds(150);

    const uint32_t threadCount = std::clamp(scanThreads, 1u, 16u);

    std::atomic_uint32_t nextId(nextNodeId);
    std::atomic_uint32_t scannedFolders(0);
    std::atomic_uint32_t scannedFiles(0);
    std::atomic_uint64_t scannedBytes(0);

    const auto scanStartedAt = std::chrono::steady_clock::now();
    ScanState scanFinalState = ScanState::Done;
    HRESULT scanFinalHr      = S_OK;
    auto emitScanSummary     = wil::scope_exit([&]() noexcept
    {
        std::wstring_view detail = L"done";
        switch (scanFinalState)
        {
            case ScanState::Canceled: detail = L"canceled"; break;
            case ScanState::Error: detail = L"error"; break;
            case ScanState::Queued: detail = L"queued"; break;
            case ScanState::Scanning: detail = L"scanning"; break;
            case ScanState::Done: detail = L"done"; break;
            case ScanState::NotStarted:
            default: detail = L"not-started"; break;
        }

        const uint64_t folders = scannedFolders.load(std::memory_order_relaxed);
        const uint64_t files   = scannedFiles.load(std::memory_order_relaxed);
        const uint64_t bytes   = scannedBytes.load(std::memory_order_relaxed);
        Debug::Perf::Emit(L"viewer.space.scan.total_us", detail, Debug::Perf::ElapsedUs(scanStartedAt), files, folders, scanFinalHr);
        Debug::Perf::EmitValue(L"viewer.space.scan.files", files, scanFinalHr);
        Debug::Perf::EmitValue(L"viewer.space.scan.folders", folders, scanFinalHr);
        Debug::Perf::EmitValue(L"viewer.space.scan.bytes", bytes, scanFinalHr);
        Debug::Perf::EmitValue(L"viewer.space.scan.active_workers", threadCount, scanFinalHr);
    });

    struct ProgressState final
    {
        std::chrono::steady_clock::time_point lastProgress = std::chrono::steady_clock::now();
        uint32_t lastProgressNodeId                        = 0;
    };

    const wchar_t pathSeparator = DeterminePreferredPathSeparator(rootPath, fileSystemIsWin32);

    auto leafNameFromPath = [&](std::wstring_view path) noexcept -> std::wstring
    {
        if (path.empty())
        {
            return {};
        }

        const std::wstring_view trimmed = TrimTrailingPathSeparators(path);
        if (trimmed.size() == 2 && IsAsciiAlpha(trimmed[0]) && trimmed[1] == L':' && path.size() >= 3)
        {
            return std::wstring(path);
        }

        const size_t lastSep = trimmed.find_last_of(L"/\\");
        if (lastSep != std::wstring_view::npos && lastSep + 1 < trimmed.size())
        {
            return std::wstring(trimmed.substr(lastSep + 1));
        }

        return std::wstring(trimmed);
    };

    auto postProgress = [&](ProgressState& progress, uint32_t nodeId, std::wstring_view currentPath) noexcept
    {
        const auto now = std::chrono::steady_clock::now();
        if (nodeId == progress.lastProgressNodeId && (now - progress.lastProgress) < kProgressUpdateInterval)
        {
            return;
        }
        progress.lastProgress       = now;
        progress.lastProgressNodeId = nodeId;

        PendingUpdate up;
        up.kind           = PendingUpdate::Kind::Progress;
        up.generation     = generation;
        up.nodeId         = nodeId;
        up.bytes          = scannedBytes.load(std::memory_order_relaxed);
        up.scannedFolders = scannedFolders.load(std::memory_order_relaxed);
        up.scannedFiles   = scannedFiles.load(std::memory_order_relaxed);
        up.name           = leafNameFromPath(currentPath);
        PostUpdate(std::move(up));
    };

    PendingUpdate queuedUp;
    queuedUp.kind       = PendingUpdate::Kind::UpdateState;
    queuedUp.generation = generation;
    queuedUp.nodeId     = rootNodeId;
    queuedUp.state      = ScanState::Queued;
    PostUpdate(std::move(queuedUp));

    if (! fileSystem)
    {
        scanFinalState = ScanState::Error;
        scanFinalHr    = E_POINTER;
        PendingUpdate errorUp;
        errorUp.kind       = PendingUpdate::Kind::UpdateState;
        errorUp.generation = generation;
        errorUp.nodeId     = rootNodeId;
        errorUp.state      = ScanState::Error;
        PostUpdate(std::move(errorUp));
        return;
    }

    ScanScheduler& scheduler = GetScanScheduler();
    ScanScheduler::Permit permit;
    if (fileSystemIsWin32)
    {
        permit = scheduler.AcquireForPath(std::filesystem::path(rootPath), stopToken);
    }
    else
    {
        wchar_t key[32]{};
        constexpr size_t keyMax                  = (sizeof(key) / sizeof(key[0])) - 1;
        const auto r                             = std::format_to_n(key, keyMax, L"fs:{:p}", static_cast<const void*>(fileSystem.get()));
        key[(r.size < keyMax) ? r.size : keyMax] = L'\0';
        permit                                   = scheduler.AcquireForKey(key, stopToken);
    }
    if (! permit)
    {
        scanFinalState = ScanState::Canceled;
        scanFinalHr    = HRESULT_FROM_WIN32(ERROR_CANCELLED);
        PendingUpdate canceled;
        canceled.kind       = PendingUpdate::Kind::UpdateState;
        canceled.generation = generation;
        canceled.nodeId     = rootNodeId;
        canceled.state      = ScanState::Canceled;
        PostUpdate(std::move(canceled));
        return;
    }

    const BOOL backgroundMode  = SetThreadPriority(GetCurrentThread(), THREAD_MODE_BACKGROUND_BEGIN);
    auto restoreBackgroundMode = wil::scope_exit([backgroundMode]
    {
        if (backgroundMode != 0)
        {
            SetThreadPriority(GetCurrentThread(), THREAD_MODE_BACKGROUND_END);
        }
    });

    PendingUpdate scanningUp;
    scanningUp.kind       = PendingUpdate::Kind::UpdateState;
    scanningUp.generation = generation;
    scanningUp.nodeId     = rootNodeId;
    scanningUp.state      = ScanState::Scanning;
    PostUpdate(std::move(scanningUp));

    struct ChildDir final
    {
        uint32_t nodeId = 0;
        std::wstring name;
    };

    struct StackItem final
    {
        uint32_t nodeId = 0;
        std::wstring path;
        uint64_t bytes          = 0;
        bool enumerated         = false;
        bool failed             = false;
        size_t processedEntries = 0;
        size_t nextChildIndex   = 0;
        std::vector<ChildDir> children;
        std::vector<FileSummaryItem> topFiles;
        uint64_t otherBytes  = 0;
        uint32_t otherCount  = 0;
        uint32_t otherNodeId = 0;
    };

    auto minHeapByBytes = [](const FileSummaryItem& a, const FileSummaryItem& b) noexcept { return a.bytes > b.bytes; };

    auto enumerate = [&](StackItem& item, ProgressState& progress, uint32_t workerIndex) noexcept -> void
    {
        const auto enumerateStartedAt = std::chrono::steady_clock::now();
        HRESULT enumerateHr           = S_OK;
        auto emitEnumerateMetric      = wil::scope_exit([&]() noexcept
        {
            Debug::Perf::Emit(L"viewer.space.scan.enumerate_us",
                              L"phase0-current",
                              Debug::Perf::ElapsedUs(enumerateStartedAt),
                              item.processedEntries,
                              workerIndex,
                              enumerateHr);
        });

        item.enumerated = true;
        postProgress(progress, item.nodeId, item.path);

        wil::com_ptr<IFilesInformation> filesInformation;
        const HRESULT enumHr = fileSystem->ReadDirectoryInfo(item.path.c_str(), filesInformation.put());
        enumerateHr          = enumHr;
        if (FAILED(enumHr) || ! filesInformation)
        {
            Debug::Warning(L"ViewerSpace: Failed to enumerate directory '{}' (HRESULT: {:#x})", item.path, static_cast<uint32_t>(enumHr));
            item.failed = true;

            PendingUpdate stateUp;
            stateUp.kind       = PendingUpdate::Kind::UpdateState;
            stateUp.generation = generation;
            stateUp.nodeId     = item.nodeId;
            stateUp.state      = ScanState::Error;
            PostUpdate(std::move(stateUp));
            return;
        }

        PendingUpdate stateUp;
        stateUp.kind       = PendingUpdate::Kind::UpdateState;
        stateUp.generation = generation;
        stateUp.nodeId     = item.nodeId;
        stateUp.state      = ScanState::Scanning;
        PostUpdate(std::move(stateUp));

        FileInfo* buffer         = nullptr;
        unsigned long bufferSize = 0;
        const HRESULT bufferHr   = filesInformation->GetBuffer(&buffer);
        const HRESULT sizeHr     = filesInformation->GetBufferSize(&bufferSize);
        if (FAILED(bufferHr) || FAILED(sizeHr))
        {
            enumerateHr = FAILED(bufferHr) ? bufferHr : sizeHr;
            Debug::Warning(L"ViewerSpace: Failed to get buffer for directory '{}' (buffer: {:#x}, size: {:#x})",
                           item.path,
                           static_cast<uint32_t>(bufferHr),
                           static_cast<uint32_t>(sizeHr));
            item.failed = true;

            PendingUpdate errorUp;
            errorUp.kind       = PendingUpdate::Kind::UpdateState;
            errorUp.generation = generation;
            errorUp.nodeId     = item.nodeId;
            errorUp.state      = ScanState::Error;
            PostUpdate(std::move(errorUp));
            return;
        }

        if (buffer != nullptr && bufferSize > 0)
        {
            const unsigned char* const bufferBytes         = reinterpret_cast<const unsigned char*>(buffer);
            unsigned long offset                           = 0;
            static constexpr unsigned long kWcharSizeBytes = static_cast<unsigned long>(sizeof(wchar_t));
            const size_t nameOffset                        = offsetof(FileInfo, FileName);

            while (offset < bufferSize)
            {
                if (stopToken.stop_requested())
                {
                    return;
                }

                if (static_cast<size_t>(bufferSize - offset) < nameOffset)
                {
                    break;
                }

                const auto* entry              = reinterpret_cast<const FileInfo*>(bufferBytes + offset);
                const unsigned long nextOffset = entry->NextEntryOffset;

                if ((entry->FileNameSize % kWcharSizeBytes) == 0u &&
                    nameOffset + static_cast<size_t>(entry->FileNameSize) <= static_cast<size_t>(bufferSize - offset))
                {
                    const size_t nameChars = static_cast<size_t>(entry->FileNameSize / kWcharSizeBytes);
                    const std::wstring_view name(entry->FileName, nameChars);

                    if (name != L"." && name != L"..")
                    {
                        item.processedEntries += 1;

                        const bool isDirectory = (entry->FileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
                        const bool isReparse   = (entry->FileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;
                        if (isDirectory)
                        {
                            if (! isReparse)
                            {
                                scannedFolders.fetch_add(1u, std::memory_order_relaxed);

                                const uint32_t dirNodeId = nextId.fetch_add(1u, std::memory_order_relaxed);

                                PendingUpdate addDir;
                                addDir.kind       = PendingUpdate::Kind::AddChild;
                                addDir.generation = generation;
                                addDir.parentId   = item.nodeId;
                                addDir.nodeId     = dirNodeId;
                                addDir.name.assign(name);
                                addDir.isDirectory = true;
                                PostUpdate(std::move(addDir));

                                ChildDir childDir;
                                childDir.nodeId = dirNodeId;
                                childDir.name.assign(name);
                                item.children.push_back(std::move(childDir));
                            }
                        }
                        else
                        {
                            scannedFiles.fetch_add(1u, std::memory_order_relaxed);
                            const uint64_t fileBytes = entry->EndOfFile > 0 ? static_cast<uint64_t>(entry->EndOfFile) : 0u;
                            item.bytes += fileBytes;
                            scannedBytes.fetch_add(fileBytes, std::memory_order_relaxed);

                            FileSummaryItem candidate;
                            candidate.bytes = fileBytes;
                            candidate.name.assign(name);

                            if (item.topFiles.size() < topFilesPerDirectory)
                            {
                                item.topFiles.push_back(std::move(candidate));
                                std::push_heap(item.topFiles.begin(), item.topFiles.end(), minHeapByBytes);
                            }
                            else if (! item.topFiles.empty() && fileBytes > item.topFiles.front().bytes)
                            {
                                std::pop_heap(item.topFiles.begin(), item.topFiles.end(), minHeapByBytes);
                                FileSummaryItem dropped = std::move(item.topFiles.back());
                                item.topFiles.back()    = std::move(candidate);
                                std::push_heap(item.topFiles.begin(), item.topFiles.end(), minHeapByBytes);

                                item.otherBytes += dropped.bytes;
                                item.otherCount += 1;
                            }
                            else
                            {
                                item.otherBytes += fileBytes;
                                item.otherCount += 1;
                            }

                            if ((item.processedEntries % kProgressUpdateStride) == 0)
                            {
                                PendingUpdate progressUp;
                                progressUp.kind       = PendingUpdate::Kind::UpdateSize;
                                progressUp.generation = generation;
                                progressUp.nodeId     = item.nodeId;
                                progressUp.bytes      = item.bytes;
                                PostUpdate(std::move(progressUp));
                                postProgress(progress, item.nodeId, item.path);
                            }
                        }
                    }
                }

                if (nextOffset == 0)
                {
                    break;
                }

                if (nextOffset > bufferSize - offset)
                {
                    break;
                }

                offset += nextOffset;
            }
        }

        PendingUpdate sizeUp;
        sizeUp.kind       = PendingUpdate::Kind::UpdateSize;
        sizeUp.generation = generation;
        sizeUp.nodeId     = item.nodeId;
        sizeUp.bytes      = item.bytes;
        PostUpdate(std::move(sizeUp));

        std::sort(item.topFiles.begin(),
                  item.topFiles.end(),
                  [](const FileSummaryItem& a, const FileSummaryItem& b) noexcept
        {
            if (a.bytes != b.bytes)
            {
                return a.bytes > b.bytes;
            }
            return a.name < b.name;
        });

        if (item.otherCount > 0 || item.otherBytes > 0)
        {
            item.otherNodeId = nextId.fetch_add(1u, std::memory_order_relaxed);
        }

        PendingUpdate filesUp;
        filesUp.kind        = PendingUpdate::Kind::DirectoryFilesSummary;
        filesUp.generation  = generation;
        filesUp.nodeId      = item.nodeId;
        filesUp.otherBytes  = item.otherBytes;
        filesUp.otherCount  = item.otherCount;
        filesUp.otherNodeId = item.otherNodeId;
        filesUp.topFiles    = std::move(item.topFiles);
        PostUpdate(std::move(filesUp));

        postProgress(progress, item.nodeId, item.path);
    };

    struct ScanJob final
    {
        uint32_t nodeId = 0;
        uint32_t parentId = 0;
        std::wstring path;
    };

    struct DirectoryCompletion final
    {
        uint32_t parentId = 0;
        uint64_t bytes = 0;
        uint32_t pendingChildren = 0;
        bool enumerated = false;
        bool completed = false;
        bool failed = false;
    };

    std::mutex frontierMutex;
    std::condition_variable frontierCv;
    std::deque<ScanJob> frontier;
    std::unordered_map<uint32_t, DirectoryCompletion> completions;
    completions.reserve(1024u);
    completions.emplace(rootNodeId, DirectoryCompletion{});
    frontier.push_back(ScanJob{rootNodeId, 0u, std::move(rootPath)});

    bool frontierFinished = false;
    std::atomic_bool rootFailed(false);
    std::atomic_uint32_t activeWorkers(0);
    std::atomic_uint32_t idleWorkerWaits(0);
    std::atomic_uint32_t frontierDepthPeak(1u);

    auto updateFrontierDepthPeak = [&](uint32_t depth) noexcept
    {
        uint32_t observed = frontierDepthPeak.load(std::memory_order_relaxed);
        while (depth > observed && ! frontierDepthPeak.compare_exchange_weak(observed, depth, std::memory_order_relaxed))
        {
        }
    };

    auto postDirectoryCompletion = [&](uint32_t nodeId, uint64_t bytes, ScanState state) noexcept
    {
        PendingUpdate sizeUp;
        sizeUp.kind       = PendingUpdate::Kind::UpdateSize;
        sizeUp.generation = generation;
        sizeUp.nodeId     = nodeId;
        sizeUp.bytes      = bytes;
        PostUpdate(std::move(sizeUp));

        PendingUpdate doneUp;
        doneUp.kind       = PendingUpdate::Kind::UpdateState;
        doneUp.generation = generation;
        doneUp.nodeId     = nodeId;
        doneUp.state      = state;
        PostUpdate(std::move(doneUp));
    };

    auto completeReadyDirectories = [&](uint32_t initialNodeId) noexcept -> void
    {
        uint32_t nodeId = initialNodeId;
        for (;;)
        {
            uint32_t parentId = 0;
            uint64_t bytes = 0;
            bool failed = false;
            bool completedRoot = false;

            {
                std::scoped_lock lock(frontierMutex);
                auto it = completions.find(nodeId);
                if (it == completions.end())
                {
                    return;
                }

                DirectoryCompletion& completion = it->second;
                if (! completion.enumerated || completion.pendingChildren != 0u || completion.completed)
                {
                    return;
                }

                completion.completed = true;
                parentId             = completion.parentId;
                bytes                = completion.bytes;
                failed               = completion.failed;
                completedRoot        = nodeId == rootNodeId;
                if (completedRoot)
                {
                    frontierFinished = true;
                    rootFailed.store(failed, std::memory_order_relaxed);
                }
            }

            postDirectoryCompletion(nodeId, bytes, failed ? ScanState::Error : ScanState::Done);

            if (completedRoot)
            {
                frontierCv.notify_all();
                return;
            }

            uint64_t parentBytes = 0;
            bool parentReady = false;
            {
                std::scoped_lock lock(frontierMutex);
                auto parentIt = completions.find(parentId);
                if (parentIt == completions.end())
                {
                    return;
                }

                DirectoryCompletion& parent = parentIt->second;
                parent.bytes += bytes;
                parentBytes = parent.bytes;
                if (parent.pendingChildren > 0u)
                {
                    parent.pendingChildren -= 1u;
                }
                parentReady = parent.enumerated && parent.pendingChildren == 0u && ! parent.completed;
            }

            PendingUpdate parentSizeUp;
            parentSizeUp.kind       = PendingUpdate::Kind::UpdateSize;
            parentSizeUp.generation = generation;
            parentSizeUp.nodeId     = parentId;
            parentSizeUp.bytes      = parentBytes;
            PostUpdate(std::move(parentSizeUp));

            if (! parentReady)
            {
                return;
            }

            nodeId = parentId;
        }
    };

    auto processJob = [&](const ScanJob& job, uint32_t workerIndex) noexcept
    {
        ProgressState progress;

        StackItem item;
        item.nodeId = job.nodeId;
        item.path   = job.path;
        enumerate(item, progress, workerIndex);

        if (stopToken.stop_requested())
        {
            return;
        }

        std::vector<ScanJob> childJobs;
        if (! item.failed)
        {
            childJobs.reserve(item.children.size());
            for (const ChildDir& child : item.children)
            {
                ScanJob childJob;
                childJob.nodeId   = child.nodeId;
                childJob.parentId = item.nodeId;
                childJob.path     = JoinPath(item.path, child.name, pathSeparator);
                childJobs.push_back(std::move(childJob));
            }
        }

        size_t frontierDepth = 0u;
        {
            std::scoped_lock lock(frontierMutex);
            DirectoryCompletion& completion = completions[item.nodeId];
            completion.parentId             = job.parentId;
            completion.bytes += item.bytes;
            completion.pendingChildren += static_cast<uint32_t>(std::min<size_t>(childJobs.size(), std::numeric_limits<uint32_t>::max()));
            completion.enumerated = true;
            completion.failed     = item.failed;

            for (auto& childJob : childJobs)
            {
                DirectoryCompletion childCompletion;
                childCompletion.parentId = item.nodeId;
                completions.emplace(childJob.nodeId, childCompletion);
                frontier.push_back(std::move(childJob));
            }

            frontierDepth = frontier.size();
        }

        updateFrontierDepthPeak(static_cast<uint32_t>(std::min<size_t>(frontierDepth, std::numeric_limits<uint32_t>::max())));
        if (! childJobs.empty())
        {
            Debug::Perf::EmitValue(L"viewer.space.scan.frontier_depth", static_cast<uint64_t>(frontierDepth));
            frontierCv.notify_all();
        }

        completeReadyDirectories(item.nodeId);
    };

    auto runWorker = [&](bool setBackgroundMode, uint32_t workerIndex) noexcept
    {
        const BOOL threadBackgroundMode  = setBackgroundMode ? SetThreadPriority(GetCurrentThread(), THREAD_MODE_BACKGROUND_BEGIN) : 0;
        auto restoreThreadBackgroundMode = wil::scope_exit([setBackgroundMode, threadBackgroundMode]
        {
            if (setBackgroundMode && threadBackgroundMode != 0)
            {
                SetThreadPriority(GetCurrentThread(), THREAD_MODE_BACKGROUND_END);
            }
        });

        for (;;)
        {
            ScanJob job;
            uint32_t activeAfterStart = 0u;
            size_t frontierDepthAfterPop = 0u;
            {
                std::unique_lock lock(frontierMutex);
                while (frontier.empty() && ! frontierFinished && ! stopToken.stop_requested())
                {
                    idleWorkerWaits.fetch_add(1u, std::memory_order_relaxed);
                    frontierCv.wait_for(lock, std::chrono::milliseconds(25));
                }

                if (stopToken.stop_requested() || (frontier.empty() && frontierFinished))
                {
                    return;
                }

                if (frontier.empty())
                {
                    continue;
                }

                job = std::move(frontier.front());
                frontier.pop_front();
                frontierDepthAfterPop = frontier.size();
                activeAfterStart      = activeWorkers.fetch_add(1u, std::memory_order_relaxed) + 1u;
            }

            Debug::Perf::EmitValue(L"viewer.space.scan.active_workers", activeAfterStart);
            Debug::Perf::EmitValue(L"viewer.space.scan.frontier_depth", static_cast<uint64_t>(frontierDepthAfterPop));

            processJob(job, workerIndex);

            const uint32_t activeAfterFinish = activeWorkers.fetch_sub(1u, std::memory_order_relaxed) - 1u;
            Debug::Perf::EmitValue(L"viewer.space.scan.active_workers", activeAfterFinish);
            {
                std::scoped_lock lock(frontierMutex);
                if (frontier.empty() && activeAfterFinish == 0u)
                {
                    frontierCv.notify_all();
                }
            }
        }
    };

    {
        const size_t extraWorkers = threadCount > 0u ? (static_cast<size_t>(threadCount) - 1u) : 0u;

        std::vector<std::jthread> workers;
        workers.reserve(extraWorkers);
        for (size_t workerIndex = 0; workerIndex < extraWorkers; ++workerIndex)
        {
            const uint32_t metricWorkerIndex = static_cast<uint32_t>(workerIndex + 1u);
            workers.emplace_back([&runWorker, metricWorkerIndex](std::stop_token) noexcept { runWorker(true, metricWorkerIndex); });
        }

        runWorker(false, 0u);
    }

    if (stopToken.stop_requested())
    {
        scanFinalState = ScanState::Canceled;
        scanFinalHr    = HRESULT_FROM_WIN32(ERROR_CANCELLED);
        PendingUpdate canceled;
        canceled.kind       = PendingUpdate::Kind::UpdateState;
        canceled.generation = generation;
        canceled.nodeId     = rootNodeId;
        canceled.state      = ScanState::Canceled;
        PostUpdate(std::move(canceled));
        return;
    }

    Debug::Perf::EmitValue(L"viewer.space.scan.frontier_depth_peak", frontierDepthPeak.load(std::memory_order_relaxed), scanFinalHr);
    Debug::Perf::EmitValue(L"viewer.space.scan.idle_worker_waits", idleWorkerWaits.load(std::memory_order_relaxed), scanFinalHr);
    if (rootFailed.load(std::memory_order_relaxed))
    {
        scanFinalState = ScanState::Error;
        scanFinalHr    = E_FAIL;
    }
}

void ViewerSpace::PostUpdate(PendingUpdate&& update) noexcept
{
    if (update.generation != _scanGeneration.load())
    {
        return;
    }

    const auto lockStartedAt = std::chrono::steady_clock::now();
    const size_t updateBytes = EstimatePendingUpdateBytes(update);
    size_t pendingCount = 0;
    uint64_t pendingBytes = 0u;
    bool coalesced = false;
    {
        std::scoped_lock lock(_updateMutex);
        if (update.kind == PendingUpdate::Kind::UpdateSize || update.kind == PendingUpdate::Kind::Progress)
        {
            size_t examined = 0u;
            for (auto it = _pendingUpdates.rbegin(); it != _pendingUpdates.rend() && examined < 64u; ++it, ++examined)
            {
                if (it->kind != update.kind || it->generation != update.generation || it->nodeId != update.nodeId)
                {
                    continue;
                }

                const size_t previousBytes = EstimatePendingUpdateBytes(*it);
                *it = std::move(update);
                const size_t replacementBytes = EstimatePendingUpdateBytes(*it);
                if (replacementBytes >= previousBytes)
                {
                    pendingBytes = _pendingUpdateBytes.fetch_add(static_cast<uint64_t>(replacementBytes - previousBytes), std::memory_order_relaxed) +
                                   static_cast<uint64_t>(replacementBytes - previousBytes);
                }
                else
                {
                    const uint64_t delta = static_cast<uint64_t>(previousBytes - replacementBytes);
                    pendingBytes = SubtractAtomicFloor(_pendingUpdateBytes, delta);
                }
                coalesced = true;
                break;
            }
        }

        if (! coalesced)
        {
            _pendingUpdates.emplace_back(std::move(update));
            pendingBytes = _pendingUpdateBytes.fetch_add(static_cast<uint64_t>(updateBytes), std::memory_order_relaxed) + static_cast<uint64_t>(updateBytes);
        }
        pendingCount = _pendingUpdates.size();
    }

    _pendingUpdatePostedCount.fetch_add(1u, std::memory_order_relaxed);
    if (coalesced)
    {
        _pendingUpdateCoalescedCount.fetch_add(1u, std::memory_order_relaxed);
        Debug::Perf::EmitCounter(L"viewer.space.queue.coalesced_count");
    }
    Debug::Perf::EmitCounter(L"viewer.space.queue.posted_count");
    Debug::Perf::EmitValue(L"viewer.space.queue.pending_bytes", pendingBytes);
    Debug::Perf::EmitDurationUs(L"viewer.space.queue.lock_hold_us", Debug::Perf::ElapsedUs(lockStartedAt), pendingCount, pendingBytes);

    static constexpr uint64_t kPendingUpdateSoftBackpressureBytes = 16ull * 1024ull * 1024ull;
    static constexpr uint64_t kPendingUpdateHardBackpressureBytes = 64ull * 1024ull * 1024ull;
    if (pendingBytes > kPendingUpdateHardBackpressureBytes)
    {
        std::this_thread::sleep_for(std::chrono::milliseconds(5));
    }
    else if (pendingBytes > kPendingUpdateSoftBackpressureBytes)
    {
        std::this_thread::sleep_for(std::chrono::milliseconds(1));
    }
    else if (pendingBytes > (kPendingUpdateSoftBackpressureBytes / 4ull))
    {
        std::this_thread::yield();
    }
}

void ViewerSpace::DrainUpdates() noexcept
{
    static constexpr double kDrainBudgetSecondsWhileScanning = 0.002;
    static constexpr double kDrainBudgetSecondsIdle          = 0.004;

    const bool scanning             = _scanActive.load();
    const double budgetSeconds      = scanning ? kDrainBudgetSecondsWhileScanning : kDrainBudgetSecondsIdle;
    const size_t maxUpdatesPerDrain = scanning ? 512u : 2048u;

    const auto drainStartedAt = std::chrono::steady_clock::now();
    const double startSeconds = NowSeconds();
    size_t processed          = 0;
    uint64_t pendingBytesAfter = _pendingUpdateBytes.load(std::memory_order_relaxed);

    bool layoutChanged   = false;
    bool headerTextDirty = false;

    std::deque<PendingUpdate> batch;
    size_t batchBytes = 0u;
    {
        std::scoped_lock lock(_updateMutex);
        while (batch.size() < maxUpdatesPerDrain && ! _pendingUpdates.empty())
        {
            const size_t updateBytes = EstimatePendingUpdateBytes(_pendingUpdates.front());
            batchBytes += updateBytes;
            batch.emplace_back(std::move(_pendingUpdates.front()));
            _pendingUpdates.pop_front();
        }

        if (batchBytes > 0u)
        {
            pendingBytesAfter = SubtractAtomicFloor(_pendingUpdateBytes, static_cast<uint64_t>(batchBytes));
        }
    }

    size_t batchIndex = 0u;
    while (batchIndex < batch.size())
    {
        PendingUpdate update = std::move(batch[batchIndex]);
        batchIndex += 1u;
        processed += 1;
        switch (update.kind)
        {
            case PendingUpdate::Kind::AddChild:
            {
                Node node;
                node.id          = update.nodeId;
                node.parentId    = update.parentId;
                node.isDirectory = update.isDirectory;
                node.isSynthetic = update.isSynthetic;
                node.name        = CopyToArena(_nameArena, update.name);
                node.totalBytes  = update.bytes;
                node.scanState   = update.isDirectory ? ScanState::Queued : ScanState::Done;

                const size_t requiredSize = static_cast<size_t>(node.id) + 1u;
                if (_nodes.size() < requiredSize)
                {
                    _nodes.resize(requiredSize);
                }

                _nodes[node.id] = std::move(node);

                Node* parent = TryGetRealNode(update.parentId);
                if (parent != nullptr)
                {
                    AddRealNodeChild(*parent, update.nodeId);
                }

                layoutChanged = true;
                break;
            }
            case PendingUpdate::Kind::UpdateSize:
            {
                Node* node = TryGetRealNode(update.nodeId);
                if (node != nullptr)
                {
                    node->totalBytes = update.bytes;
                    layoutChanged    = true;
                }
                break;
            }
            case PendingUpdate::Kind::UpdateState:
            {
                Node* node = TryGetRealNode(update.nodeId);
                if (node != nullptr)
                {
                    node->scanState = update.state;
                    if (update.nodeId == _rootNodeId)
                    {
                        headerTextDirty = true;
                    }
                }
                break;
            }
            case PendingUpdate::Kind::DirectoryFilesSummary:
            {
                const size_t requiredSize = static_cast<size_t>(std::max(update.otherNodeId, update.nodeId)) + 1u;
                if (_nodes.size() < requiredSize)
                {
                    _nodes.resize(requiredSize);
                }

                Node* parent = TryGetRealNode(update.nodeId);
                if (parent == nullptr)
                {
                    break;
                }

                for (const auto& file : update.topFiles)
                {
                    const uint32_t fileId = FileRecordIdFromIndex(_fileRecords.size());
                    if (fileId == 0)
                    {
                        continue;
                    }

                    FileRecord record;
                    record.id       = fileId;
                    record.parentId = update.nodeId;
                    record.name     = CopyToArena(_nameArena, file.name);
                    record.bytes    = file.bytes;

                    _fileRecords.push_back(record);
                    AddRealNodeChild(*parent, fileId);
                }

                if (update.otherNodeId != 0 && (update.otherCount > 0 || update.otherBytes > 0))
                {
                    Node otherNode;
                    otherNode.id               = update.otherNodeId;
                    otherNode.parentId         = update.nodeId;
                    otherNode.isDirectory      = false;
                    otherNode.isSynthetic      = true;
                    otherNode.scanState        = ScanState::Done;
                    otherNode.totalBytes       = update.otherBytes;
                    otherNode.aggregateFolders = 0;
                    otherNode.aggregateFiles   = update.otherCount;

                    std::wstring otherName          = FormatStringResource(g_hInstance, IDS_VIEWERSPACE_OTHER_BUCKET_FORMAT, update.otherCount);
                    const std::wstring otherDetails = FormatAggregateCountsLine(otherNode.aggregateFolders, otherNode.aggregateFiles);
                    if (! otherDetails.empty())
                    {
                        otherName.append(L"\n");
                        otherName.append(otherDetails);
                    }

                    otherNode.name = CopyToArena(_nameArena, otherName);

                    _nodes[otherNode.id] = std::move(otherNode);
                    AddRealNodeChild(*parent, update.otherNodeId);
                }

                layoutChanged = true;
                break;
            }
            case PendingUpdate::Kind::Progress:
            {
                _scanProgressBytes        = update.bytes;
                _scanProgressFolders      = update.scannedFolders;
                _scanProgressFiles        = update.scannedFiles;
                _scanProcessingNodeId     = update.nodeId;
                _scanProcessingFolderName = std::move(update.name);

                Node* root = TryGetRealNode(_rootNodeId);
                if (root != nullptr && root->scanState != ScanState::Done)
                {
                    root->totalBytes = update.bytes;
                }

                headerTextDirty = true;
                break;
            }
        }

        if (processed >= 128u && (processed % 64u) == 0u)
        {
            const double elapsed = NowSeconds() - startSeconds;
            if (elapsed >= budgetSeconds)
            {
                break;
            }
        }
    }

    if (batchIndex < batch.size())
    {
        size_t remainingBytes = 0u;
        for (size_t index = batchIndex; index < batch.size(); ++index)
        {
            remainingBytes += EstimatePendingUpdateBytes(batch[index]);
        }

        {
            std::scoped_lock lock(_updateMutex);
            for (size_t index = batch.size(); index-- > batchIndex;)
            {
                _pendingUpdates.emplace_front(std::move(batch[index]));
            }
            pendingBytesAfter =
                _pendingUpdateBytes.fetch_add(static_cast<uint64_t>(remainingBytes), std::memory_order_relaxed) + static_cast<uint64_t>(remainingBytes);
        }
    }

    if (processed == 0)
    {
        return;
    }

    const uint64_t drainUs = Debug::Perf::ElapsedUs(drainStartedAt);
#ifdef ENABLE_TESTS
    _lastDrainUs = drainUs;
#endif
    Debug::Perf::Emit(L"viewer.space.queue.drain_us", L"phase0-current", drainUs, processed, pendingBytesAfter, S_OK);
    Debug::Perf::EmitValue(L"viewer.space.queue.coalesced_count", _pendingUpdateCoalescedCount.load(std::memory_order_relaxed));

    const ScanState previousOverallState = _overallState;

    const bool wasScanActive = _scanActive.load();
    bool isScanActiveNow     = false;

    const Node* root = TryGetRealNode(_rootNodeId);
    if (root != nullptr)
    {
        _overallState = root->scanState;
        if (root->scanState == ScanState::Queued || root->scanState == ScanState::Scanning || root->scanState == ScanState::NotStarted)
        {
            isScanActiveNow = true;
        }
        else
        {
            isScanActiveNow = false;
        }
    }
    else
    {
        _overallState   = ScanState::NotStarted;
        isScanActiveNow = false;
    }

    _scanActive.store(isScanActiveNow);
    if (previousOverallState != _overallState)
    {
        if (_overallState == ScanState::Done)
        {
            _scanCompletedSinceSeconds = NowSeconds();
        }
        else
        {
            _scanCompletedSinceSeconds = 0.0;
        }
    }
    if (wasScanActive != isScanActiveNow)
    {
        _layoutDirty              = true;
        _lastLayoutRebuildSeconds = 0.0;
        headerTextDirty           = true;
    }

    if (layoutChanged)
    {
        _layoutDirty = true;
        _layoutCandidateCache.clear();
    }

    if (layoutChanged || headerTextDirty || previousOverallState != _overallState || wasScanActive != isScanActiveNow)
    {
        InvalidateStaticTreemapCache();
    }

    if (headerTextDirty)
    {
        UpdateHeaderTextCache();
    }

    if (_hWnd && (layoutChanged || headerTextDirty || wasScanActive != isScanActiveNow))
    {
        if (! isScanActiveNow || wasScanActive != isScanActiveNow)
        {
            InvalidateRect(_hWnd.get(), nullptr, FALSE);
        }
    }
}

void ViewerSpace::CancelScanCacheBuild() noexcept
{
    _scanCacheBuildSnapshot.reset();
    _scanCacheBuildRootKey.clear();
    _scanCacheBuildTopFilesPerDirectory = 0;
    _scanCacheBuildGeneration           = 0;
    _scanCacheBuildChildrenNext         = 0;
    _scanCacheBuildNodesNext            = 0;
    _scanCacheBuildFileRecordsNext      = 0;
}

void ViewerSpace::ContinueScanCacheBuild() noexcept
{
    if (! _config.cacheEnabled || _scanRootPath.empty() || ! _fileSystemIsWin32)
    {
        CancelScanCacheBuild();
        return;
    }

    if (! g_cacheEnabled.load(std::memory_order_acquire) || g_cacheMaxEntries.load(std::memory_order_acquire) == 0)
    {
        CancelScanCacheBuild();
        return;
    }

    const uint32_t currentGeneration = _scanGeneration.load(std::memory_order_acquire);
    if (_overallState != ScanState::Done)
    {
        CancelScanCacheBuild();
        return;
    }

    if (! _scanCacheBuildSnapshot && _scanCacheLastStoredGeneration != currentGeneration)
    {
        ScanResultCacheKey cacheKey;
        cacheKey.rootKey              = NormalizeRootPathForScanCache(std::filesystem::path(_scanRootPath));
        cacheKey.topFilesPerDirectory = _scanTopFilesPerDirectory;

        if (! cacheKey.rootKey.empty())
        {
            const uint64_t maxSnapshotBytes = MaxScanCacheSnapshotBytes();
            uint64_t estimatedSnapshotBytes = static_cast<uint64_t>(
                _childrenArena.size() * sizeof(uint32_t) + _nodes.size() * sizeof(ScanResultCacheNode) +
                _fileRecords.size() * sizeof(ScanResultCacheFileRecord));
            for (const Node& node : _nodes)
            {
                estimatedSnapshotBytes += static_cast<uint64_t>(node.name.size() * sizeof(wchar_t));
                if (estimatedSnapshotBytes > maxSnapshotBytes)
                {
                    break;
                }
            }
            for (const FileRecord& file : _fileRecords)
            {
                estimatedSnapshotBytes += static_cast<uint64_t>(file.name.size() * sizeof(wchar_t));
                if (estimatedSnapshotBytes > maxSnapshotBytes)
                {
                    break;
                }
            }

            if (estimatedSnapshotBytes > maxSnapshotBytes)
            {
                _lastScanCacheSnapshotBytes = estimatedSnapshotBytes;
                Debug::Perf::EmitValue(L"viewer.space.model.cache_skipped_large", estimatedSnapshotBytes);
                _scanCacheLastStoredGeneration = currentGeneration;
                CancelScanCacheBuild();
                return;
            }

            auto snapshot = std::make_shared<ScanResultSnapshot>();
            snapshot->nodes.reserve(_nodes.size());
            snapshot->fileRecords.reserve(_fileRecords.size());
            snapshot->childrenArena.reserve(_childrenArena.size());

            _scanCacheBuildSnapshot             = snapshot;
            _scanCacheBuildRootKey              = std::move(cacheKey.rootKey);
            _scanCacheBuildTopFilesPerDirectory = cacheKey.topFilesPerDirectory;
            _scanCacheBuildGeneration           = currentGeneration;
            _scanCacheBuildChildrenNext         = 0;
            _scanCacheBuildNodesNext            = 0;
            _scanCacheBuildFileRecordsNext      = 0;
        }
    }

    if (! _scanCacheBuildSnapshot)
    {
        return;
    }

    if (_scanCacheBuildGeneration != currentGeneration)
    {
        CancelScanCacheBuild();
        return;
    }

    auto snapshot = std::static_pointer_cast<ScanResultSnapshot>(_scanCacheBuildSnapshot);
    if (! snapshot)
    {
        CancelScanCacheBuild();
        return;
    }

    static constexpr double kCacheBuildBudgetSeconds = 0.0012;
    const double startSeconds                        = NowSeconds();

    const size_t childCount = _childrenArena.size();
    while (_scanCacheBuildChildrenNext < childCount)
    {
        snapshot->childrenArena.push_back(_childrenArena[_scanCacheBuildChildrenNext]);
        _scanCacheBuildChildrenNext += 1;

        if ((_scanCacheBuildChildrenNext % 4096u) == 0u)
        {
            if ((NowSeconds() - startSeconds) >= kCacheBuildBudgetSeconds)
            {
                return;
            }
        }
    }

    const size_t nodeCount = _nodes.size();
    while (_scanCacheBuildNodesNext < nodeCount)
    {
        const Node& node = _nodes[_scanCacheBuildNodesNext];

        ScanResultCacheNode cachedNode;
        if (node.id != 0)
        {
            cachedNode.id               = node.id;
            cachedNode.parentId         = node.parentId;
            cachedNode.isDirectory      = node.isDirectory;
            cachedNode.isSynthetic      = node.isSynthetic;
            cachedNode.scanState        = static_cast<uint8_t>(node.scanState);
            cachedNode.totalBytes       = node.totalBytes;
            cachedNode.childrenStart    = node.childrenStart;
            cachedNode.childrenCount    = node.childrenCount;
            cachedNode.childrenCapacity = node.childrenCapacity;
            cachedNode.aggregateFolders = node.aggregateFolders;
            cachedNode.aggregateFiles   = node.aggregateFiles;
            cachedNode.name.assign(node.name.data(), node.name.size());
        }

        snapshot->nodes.push_back(std::move(cachedNode));
        _scanCacheBuildNodesNext += 1;

        if ((_scanCacheBuildNodesNext % 256u) == 0u)
        {
            if ((NowSeconds() - startSeconds) >= kCacheBuildBudgetSeconds)
            {
                return;
            }
        }
    }

    if (snapshot->nodes.size() != nodeCount)
    {
        CancelScanCacheBuild();
        return;
    }

    const size_t fileRecordCount = _fileRecords.size();
    while (_scanCacheBuildFileRecordsNext < fileRecordCount)
    {
        const FileRecord& file = _fileRecords[_scanCacheBuildFileRecordsNext];

        ScanResultCacheFileRecord cachedFile;
        if (file.id != 0)
        {
            cachedFile.id       = file.id;
            cachedFile.parentId = file.parentId;
            cachedFile.bytes    = file.bytes;
            cachedFile.name.assign(file.name.data(), file.name.size());
        }

        snapshot->fileRecords.push_back(std::move(cachedFile));
        _scanCacheBuildFileRecordsNext += 1;

        if ((_scanCacheBuildFileRecordsNext % 512u) == 0u)
        {
            if ((NowSeconds() - startSeconds) >= kCacheBuildBudgetSeconds)
            {
                return;
            }
        }
    }

    if (snapshot->fileRecords.size() != fileRecordCount)
    {
        CancelScanCacheBuild();
        return;
    }

    ScanResultCacheKey cacheKey;
    cacheKey.rootKey              = _scanCacheBuildRootKey;
    cacheKey.topFilesPerDirectory = _scanCacheBuildTopFilesPerDirectory;

    uint64_t snapshotBytes = static_cast<uint64_t>(
        snapshot->childrenArena.size() * sizeof(uint32_t) + snapshot->nodes.size() * sizeof(ScanResultCacheNode) +
        snapshot->fileRecords.size() * sizeof(ScanResultCacheFileRecord));
    for (const ScanResultCacheNode& node : snapshot->nodes)
    {
        snapshotBytes += static_cast<uint64_t>(node.name.size() * sizeof(wchar_t));
    }
    for (const ScanResultCacheFileRecord& file : snapshot->fileRecords)
    {
        snapshotBytes += static_cast<uint64_t>(file.name.size() * sizeof(wchar_t));
    }
    _lastScanCacheSnapshotBytes = snapshotBytes;
    Debug::Perf::EmitValue(L"viewer.space.model.cache_snapshot_bytes", snapshotBytes);
    if (snapshotBytes > MaxScanCacheSnapshotBytes())
    {
        Debug::Perf::EmitValue(L"viewer.space.model.cache_skipped_large", snapshotBytes);
        _scanCacheLastStoredGeneration = _scanCacheBuildGeneration;
        CancelScanCacheBuild();
        return;
    }

    GetScanResultCache().Store(std::move(cacheKey), std::move(snapshot));
    _scanCacheLastStoredGeneration = _scanCacheBuildGeneration;
    CancelScanCacheBuild();
}

size_t ViewerSpace::EstimatePendingUpdateBytes(const PendingUpdate& update) noexcept
{
    size_t bytes = sizeof(PendingUpdate);
    bytes += update.name.size() * sizeof(wchar_t);
    bytes += update.topFiles.capacity() * sizeof(FileSummaryItem);
    for (const FileSummaryItem& file : update.topFiles)
    {
        bytes += file.name.size() * sizeof(wchar_t);
    }
    return bytes;
}

uint64_t ViewerSpace::SubtractAtomicFloor(std::atomic_uint64_t& value, uint64_t delta) noexcept
{
    uint64_t current = value.load(std::memory_order_relaxed);
    for (;;)
    {
        const uint64_t next = current > delta ? current - delta : 0u;
        if (value.compare_exchange_weak(current, next, std::memory_order_relaxed, std::memory_order_relaxed))
        {
            return next;
        }
    }
}

const ViewerSpace::Node* ViewerSpace::TryGetRealNode(uint32_t nodeId) const noexcept
{
    if (nodeId == 0)
    {
        return nullptr;
    }

    const size_t index = static_cast<size_t>(nodeId);
    if (index >= _nodes.size())
    {
        return nullptr;
    }

    const Node& node = _nodes[index];
    if (node.id != nodeId)
    {
        return nullptr;
    }

    return &node;
}

ViewerSpace::Node* ViewerSpace::TryGetRealNode(uint32_t nodeId) noexcept
{
    if (nodeId == 0)
    {
        return nullptr;
    }

    const size_t index = static_cast<size_t>(nodeId);
    if (index >= _nodes.size())
    {
        return nullptr;
    }

    Node& node = _nodes[index];
    if (node.id != nodeId)
    {
        return nullptr;
    }

    return &node;
}

bool ViewerSpace::IsFileRecordId(uint32_t itemId) noexcept
{
    return itemId >= kFileRecordIdBase && itemId < kSyntheticNodeIdBase;
}

bool ViewerSpace::IsSyntheticNodeId(uint32_t itemId) noexcept
{
    return itemId >= kSyntheticNodeIdBase;
}

uint32_t ViewerSpace::FileRecordIdFromIndex(size_t index) noexcept
{
    if (index > static_cast<size_t>(kTreemapItemIdMask))
    {
        return 0u;
    }

    return kFileRecordIdBase + static_cast<uint32_t>(index);
}

uint32_t ViewerSpace::SyntheticOtherBucketIdForParent(uint32_t parentId) noexcept
{
    return kSyntheticNodeIdBase | (parentId & kTreemapItemIdMask);
}

const ViewerSpace::FileRecord* ViewerSpace::TryGetFileRecord(uint32_t itemId) const noexcept
{
    if (! IsFileRecordId(itemId))
    {
        return nullptr;
    }

    const size_t index = static_cast<size_t>(itemId - kFileRecordIdBase);
    if (index >= _fileRecords.size())
    {
        return nullptr;
    }

    const FileRecord& file = _fileRecords[index];
    if (file.id != itemId)
    {
        return nullptr;
    }

    return &file;
}

std::optional<ViewerSpace::ResolvedItem> ViewerSpace::ResolveTreemapItem(uint32_t itemId) const noexcept
{
    if (itemId == 0)
    {
        return std::nullopt;
    }

    if (const FileRecord* file = TryGetFileRecord(itemId))
    {
        return ResolvedItem{
            .id = file->id,
            .parentId = file->parentId,
            .isDirectory = false,
            .isSynthetic = false,
            .isFileRecord = true,
            .scanState = ScanState::Done,
            .name = file->name,
            .totalBytes = file->bytes,
            .aggregateFolders = 0,
            .aggregateFiles = 1,
        };
    }

    if (const Node* node = TryGetRealNode(itemId))
    {
        return ResolvedItem{
            .id = node->id,
            .parentId = node->parentId,
            .isDirectory = node->isDirectory,
            .isSynthetic = node->isSynthetic,
            .isFileRecord = false,
            .scanState = node->scanState,
            .name = node->name,
            .totalBytes = node->totalBytes,
            .aggregateFolders = node->aggregateFolders,
            .aggregateFiles = node->aggregateFiles,
        };
    }

    const auto synthIt = _syntheticNodes.find(itemId);
    if (synthIt != _syntheticNodes.end())
    {
        const Node& node = synthIt->second;
        return ResolvedItem{
            .id = node.id,
            .parentId = node.parentId,
            .isDirectory = node.isDirectory,
            .isSynthetic = node.isSynthetic,
            .isFileRecord = false,
            .scanState = node.scanState,
            .name = node.name,
            .totalBytes = node.totalBytes,
            .aggregateFolders = node.aggregateFolders,
            .aggregateFiles = node.aggregateFiles,
        };
    }

    return std::nullopt;
}

std::span<const uint32_t> ViewerSpace::GetRealNodeChildren(const Node& node) const noexcept
{
    if (node.childrenCount == 0)
    {
        return {};
    }

    const size_t start = static_cast<size_t>(node.childrenStart);
    const size_t count = static_cast<size_t>(node.childrenCount);
    if (start >= _childrenArena.size())
    {
        return {};
    }

    const size_t end = start + count;
    if (end > _childrenArena.size())
    {
        return {};
    }

    return std::span<const uint32_t>(_childrenArena.data() + start, count);
}

void ViewerSpace::AddRealNodeChild(Node& parent, uint32_t childItemId) noexcept
{
    static constexpr uint32_t kInitialCapacity = 8;

    if (parent.childrenCapacity == 0)
    {
        const size_t start = _childrenArena.size();
        if (start > static_cast<size_t>(std::numeric_limits<uint32_t>::max()))
        {
            return;
        }

        parent.childrenStart    = static_cast<uint32_t>(start);
        parent.childrenCount    = 0;
        parent.childrenCapacity = kInitialCapacity;
        _childrenArena.resize(start + static_cast<size_t>(kInitialCapacity));
    }

    if (parent.childrenCount >= parent.childrenCapacity)
    {
        const uint32_t oldStart    = parent.childrenStart;
        const uint32_t oldCount    = parent.childrenCount;
        const uint32_t oldCapacity = parent.childrenCapacity;

        uint32_t nextCapacity = oldCapacity;
        if (nextCapacity < kInitialCapacity)
        {
            nextCapacity = kInitialCapacity;
        }
        else if (nextCapacity > (std::numeric_limits<uint32_t>::max() / 2u))
        {
            nextCapacity = std::numeric_limits<uint32_t>::max();
        }
        else
        {
            nextCapacity = nextCapacity * 2u;
        }

        const size_t previousSize = _childrenArena.size();
        if (previousSize > static_cast<size_t>(std::numeric_limits<uint32_t>::max()))
        {
            return;
        }

        const uint32_t newStart = static_cast<uint32_t>(previousSize);
        _childrenArena.resize(previousSize + static_cast<size_t>(nextCapacity));

        if (oldCount > 0)
        {
            const size_t oldStartIndex = static_cast<size_t>(oldStart);
            const size_t oldCountIndex = static_cast<size_t>(oldCount);
            if (oldStartIndex < previousSize && (oldStartIndex + oldCountIndex) <= previousSize)
            {
                std::copy_n(_childrenArena.data() + oldStartIndex, oldCountIndex, _childrenArena.data() + previousSize);
            }
        }

        parent.childrenStart    = newStart;
        parent.childrenCapacity = nextCapacity;
        parent.childrenCount    = oldCount;
    }

    const size_t slot = static_cast<size_t>(parent.childrenStart) + static_cast<size_t>(parent.childrenCount);
    if (slot >= _childrenArena.size())
    {
        return;
    }

    _childrenArena[slot] = childItemId;
    parent.childrenCount += 1;
}

std::wstring ViewerSpace::BuildNodePathText(uint32_t nodeId) const
{
    const Node* node = TryGetRealNode(nodeId);
    if (node == nullptr || node->isSynthetic)
    {
        return {};
    }

    if (_scanRootPath.empty() || _rootNodeId == 0)
    {
        return {};
    }

    if (nodeId == _rootNodeId)
    {
        return _scanRootPath;
    }

    std::vector<const Node*> segments;
    segments.reserve(24);

    uint32_t currentId = nodeId;
    while (currentId != 0 && currentId != _rootNodeId)
    {
        const Node* current = TryGetRealNode(currentId);
        if (current == nullptr)
        {
            return {};
        }

        if (current->name.empty())
        {
            return {};
        }

        segments.push_back(current);
        currentId = current->parentId;
    }

    if (currentId != _rootNodeId)
    {
        return {};
    }

    const wchar_t separator = DeterminePreferredPathSeparator(_scanRootPath, _fileSystemIsWin32);

    std::wstring pathText = _scanRootPath;
    if (! pathText.empty() && pathText.back() != L'/' && pathText.back() != L'\\')
    {
        pathText.push_back(separator);
    }

    for (auto it = segments.rbegin(); it != segments.rend(); ++it)
    {
        const Node* segment = *it;
        if (segment == nullptr)
        {
            return {};
        }

        pathText.append(segment->name.data(), segment->name.size());
        if (it + 1 != segments.rend())
        {
            pathText.push_back(separator);
        }
    }

    return pathText;
}

std::wstring ViewerSpace::BuildItemPathText(uint32_t itemId) const
{
    if (const FileRecord* file = TryGetFileRecord(itemId))
    {
        std::wstring parentPath = BuildNodePathText(file->parentId);
        if (parentPath.empty() || file->name.empty())
        {
            return {};
        }

        const wchar_t separator = DeterminePreferredPathSeparator(parentPath, _fileSystemIsWin32);
        if (! parentPath.empty() && parentPath.back() != L'/' && parentPath.back() != L'\\')
        {
            parentPath.push_back(separator);
        }
        parentPath.append(file->name.data(), file->name.size());
        return parentPath;
    }

    return BuildNodePathText(itemId);
}

void ViewerSpace::UpdateViewPathText() noexcept
{
    _viewPathText = BuildNodePathText(_viewNodeId);
    _headerPathSourceText.clear();
    _headerPathDisplayText.clear();
    _headerPathDisplayMaxWidthDip = 0.0f;
}

uint32_t ViewerSpace::ComputeAdaptiveFileCandidateBudget() const noexcept
{
    const float widthDip        = DipFromPx(_clientSize.cx);
    const float heightDip       = DipFromPx(_clientSize.cy);
    const float headerBottomDip = GetHeaderBottomDip();
    const D2D1_RECT_F viewRc    = D2D1::RectF(kPaddingDip,
                                           headerBottomDip + kPaddingDip,
                                           std::max(kPaddingDip, widthDip - kPaddingDip),
                                           std::max(headerBottomDip + kPaddingDip, heightDip - kPaddingDip));

    const uint32_t pixelDrivenBudget = std::min<uint32_t>(PixelDrivenLayoutItemLimit(RectArea(viewRc)), kAdaptiveFileCandidateCeiling);
    return std::max<uint32_t>(_config.topFilesPerDirectory, pixelDrivenBudget);
}

void ViewerSpace::EnsureLayoutForView() noexcept
{
    const double now = NowSeconds();
    for (auto& item : _drawItems)
    {
        const double dt         = now - item.animationStartSeconds;
        const double t          = std::clamp(dt / kAnimationDurationSeconds, 0.0, 1.0);
        const double eased      = EaseOutCubic(t);
        item.currentRect.left   = static_cast<float>(item.startRect.left + (item.targetRect.left - item.startRect.left) * eased);
        item.currentRect.top    = static_cast<float>(item.startRect.top + (item.targetRect.top - item.startRect.top) * eased);
        item.currentRect.right  = static_cast<float>(item.startRect.right + (item.targetRect.right - item.startRect.right) * eased);
        item.currentRect.bottom = static_cast<float>(item.startRect.bottom + (item.targetRect.bottom - item.startRect.bottom) * eased);
    }
}

ViewerSpace::DrawItem::Lod ViewerSpace::ClassifyDrawItemLod(const D2D1_RECT_F& itemRc, uint8_t depth) noexcept
{
    float gap = kItemGapDip - static_cast<float>(depth) * 0.15f;
    gap       = std::clamp(gap, 0.5f, kItemGapDip);

    const float w = std::max(0.0f, (itemRc.right - itemRc.left) - (gap * 2.0f));
    const float h = std::max(0.0f, (itemRc.bottom - itemRc.top) - (gap * 2.0f));
    const float area = w * h;

    if (w < 1.0f || h < 1.0f || area < 1.0f)
    {
        return DrawItem::Lod::Culled;
    }

    if (w >= 110.0f && h >= 80.0f)
    {
        return DrawItem::Lod::Hero;
    }

    if (w >= 56.0f && h >= 36.0f && area >= 1800.0f)
    {
        return DrawItem::Lod::Large;
    }

    if (w >= 24.0f && h >= 12.0f && area >= 400.0f)
    {
        return DrawItem::Lod::Medium;
    }

    if (std::min(w, h) >= 3.0f && area >= 16.0f)
    {
        return DrawItem::Lod::Small;
    }

    return DrawItem::Lod::Tiny;
}

void ViewerSpace::MaybeRebuildLayout() noexcept
{
    if (! _layoutDirty)
    {
        return;
    }

    const double now                                                      = NowSeconds();
    static constexpr double kMinLayoutRebuildIntervalSecondsWhileScanning = 0.06;
    if (_scanActive.load())
    {
        const double sinceLast = now - _lastLayoutRebuildSeconds;
        if (_lastLayoutRebuildSeconds > 0.0 && sinceLast < kMinLayoutRebuildIntervalSecondsWhileScanning)
        {
            return;
        }
    }

    RebuildLayout();
    _lastLayoutRebuildSeconds = now;
    _layoutDirty              = false;
}

void ViewerSpace::RebuildLayout() noexcept
{
    const auto layoutStartedAt = std::chrono::steady_clock::now();
    InvalidateStaticTreemapCache();

    for (const auto& item : _drawItems)
    {
        _lastRectsByNode[item.nodeId] = item.currentRect;
    }

    _drawItems.clear();
    _hitGrid.Clear();
    _syntheticNodes.clear();
    _layoutNameArena.release();

    const Node* viewNode = TryGetRealNode(_viewNodeId);
    if (viewNode == nullptr)
    {
        return;
    }

    const Node& view = *viewNode;

    const float width  = DipFromPx(_clientSize.cx);
    const float height = DipFromPx(_clientSize.cy);
    if (width <= 1.0f || height <= 1.0f)
    {
        return;
    }

    const bool scanning                                         = _scanActive.load();
    static constexpr float kAutoExpandAreaFractionWhileScanning = 0.08f;
    static constexpr float kAutoExpandAreaFractionIdle          = 0.035f;
    const float autoExpandAreaFractionThreshold                 = scanning ? kAutoExpandAreaFractionWhileScanning : kAutoExpandAreaFractionIdle;

    const uint8_t maxAutoExpandDepth = static_cast<uint8_t>(scanning ? 8u : 10u);
    const size_t maxDrawItems        = scanning ? kMaxDrawItemsWhileScanning : kMaxDrawItems;

    constexpr float kNestedInsetDip         = 2.0f;
    constexpr float kMinExpandAreaDip2      = 140.0f * 110.0f;
    constexpr float kMinExpandChildAreaDip2 = 110.0f * 80.0f;
    constexpr float kMinExpandChildSideDip  = 60.0f;

    const float headerBottomDip = GetHeaderBottomDip();
    const D2D1_RECT_F rc        = D2D1::RectF(
        kPaddingDip, headerBottomDip + kPaddingDip, std::max(kPaddingDip, width - kPaddingDip), std::max(headerBottomDip + kPaddingDip, height - kPaddingDip));
    const float viewAreaDip2 = RectArea(rc);

    using Item       = LayoutItem;
    using ExpandTask = LayoutExpandTask;

    const double now                 = NowSeconds();
    size_t remaining                 = maxDrawItems;
    const double layoutStartSeconds  = now;
    const double layoutBudgetSeconds = scanning ? 0.010 : 0.030;

    auto getNode = [&](uint32_t nodeId) noexcept -> const Node*
    {
        const Node* node = TryGetRealNode(nodeId);
        if (node != nullptr)
        {
            return node;
        }

        const auto synthIt = _syntheticNodes.find(nodeId);
        if (synthIt != _syntheticNodes.end())
        {
            return &synthIt->second;
        }

        return nullptr;
    };

    auto pushDrawItem = [&](uint32_t nodeId, const D2D1_RECT_F& itemRc, uint8_t depth, float labelHeightDip) noexcept
    {
        if (remaining == 0)
        {
            return;
        }

        DrawItem di;
        di.nodeId         = nodeId;
        di.depth          = depth;
        di.labelHeightDip = labelHeightDip;
        di.targetRect     = itemRc;
        di.lod            = ClassifyDrawItemLod(itemRc, depth);
        di.areaDip2       = RectArea(itemRc);

        const auto previousRectIt = _lastRectsByNode.find(nodeId);
        if (previousRectIt != _lastRectsByNode.end())
        {
            di.startRect   = previousRectIt->second;
            di.currentRect = previousRectIt->second;
        }
        else
        {
            const float cx = (itemRc.left + itemRc.right) * 0.5f;
            const float cy = (itemRc.top + itemRc.bottom) * 0.5f;
            di.startRect   = D2D1::RectF(cx, cy, cx, cy);
            di.currentRect = di.startRect;
        }

        di.animationStartSeconds = now;
        _drawItems.push_back(di);
        remaining -= 1;
    };

    auto buildItemsForNode = [&](uint32_t parentId, float projectedAreaDip2, LayoutWorkspace& workspace, std::vector<Item>& out) noexcept
    {
        out.clear();

        const Node* parent = TryGetRealNode(parentId);
        if (parent == nullptr)
        {
            return;
        }

        const std::span<const uint32_t> children = GetRealNodeChildren(*parent);

        uint32_t maxLayoutItems  = PixelDrivenLayoutItemLimit(projectedAreaDip2);
        const auto layoutLimitIt = _layoutMaxItemsByNode.find(parentId);
        if (layoutLimitIt != _layoutMaxItemsByNode.end())
        {
            maxLayoutItems = std::max(maxLayoutItems, layoutLimitIt->second);
        }
        maxLayoutItems  = std::clamp<uint32_t>(maxLayoutItems, 32u, static_cast<uint32_t>(kMaxLayoutItems));
        size_t maxItems = static_cast<size_t>(maxLayoutItems);

        auto capMaxItemsToBudget = [&]() noexcept
        {
            if (remaining == 0)
            {
                return;
            }

            maxItems = std::min(maxItems, remaining);
            maxItems = std::max<size_t>(maxItems, 1u);
        };
        capMaxItemsToBudget();

        const size_t reserveCount = std::min(children.size(), maxItems) + 1;
        out.reserve(reserveCount);

        std::vector<Item>& topItems            = workspace.topItems;
        std::vector<Item>& forcedItems         = workspace.forcedItems;
        std::vector<uint32_t>& forcedChildIds  = workspace.forcedChildIds;

        uint64_t otherBytes   = 0;
        double otherWeight    = 0.0;
        size_t otherCount     = 0;
        uint64_t otherFolders = 0;
        uint64_t otherFiles   = 0;

        auto addUnderlyingCounts = [&](const ResolvedItem& nodeRef) noexcept
        {
            if (nodeRef.isSynthetic)
            {
                otherFolders += nodeRef.aggregateFolders;
                otherFiles += nodeRef.aggregateFiles;
                return;
            }

            if (nodeRef.isDirectory)
            {
                otherFolders += 1;
            }
            else
            {
                otherFiles += 1;
            }
        };

        auto appendOtherBucket = [&](uint64_t bucketBytes,
                                     double bucketWeight,
                                     size_t bucketCount,
                                     uint64_t bucketFolders,
                                     uint64_t bucketFiles) noexcept
        {
            if (bucketCount == 0 || bucketWeight <= 0.0)
            {
                return;
            }

            const uint32_t otherId = SyntheticOtherBucketIdForParent(parentId);

            Node other;
            other.id               = otherId;
            other.parentId         = parentId;
            other.isDirectory      = false;
            other.isSynthetic      = true;
            other.scanState        = ScanState::Done;
            other.totalBytes       = bucketBytes;
            other.aggregateFolders = static_cast<uint32_t>(std::min<uint64_t>(bucketFolders, static_cast<uint64_t>(std::numeric_limits<uint32_t>::max())));
            other.aggregateFiles   = static_cast<uint32_t>(std::min<uint64_t>(bucketFiles, static_cast<uint64_t>(std::numeric_limits<uint32_t>::max())));

            const uint64_t otherItemCount = bucketFolders + bucketFiles;
            std::wstring otherName = FormatStringResource(g_hInstance, IDS_VIEWERSPACE_OTHER_BUCKET_FORMAT, otherItemCount == 0 ? bucketCount : otherItemCount);

            const std::wstring otherDetails = FormatAggregateCountsLine(other.aggregateFolders, other.aggregateFiles);
            if (! otherDetails.empty())
            {
                otherName.append(L"\n");
                otherName.append(otherDetails);
            }

            other.name                = CopyToArena(_layoutNameArena, otherName);
            _syntheticNodes[other.id] = other;

            Item otherItem;
            otherItem.nodeId = other.id;
            otherItem.bytes  = bucketBytes;
            otherItem.weight = bucketWeight;
            out.push_back(otherItem);
        };

        auto mixCandidateSignature = [](uint64_t seed, uint64_t value) noexcept -> uint64_t
        {
            seed ^= value + 0x9E3779B97F4A7C15ull + (seed << 6) + (seed >> 2);
            return seed;
        };

        uint64_t candidateSignature = 0xCBF29CE484222325ull;
        candidateSignature = mixCandidateSignature(candidateSignature, static_cast<uint64_t>(children.size()));
        candidateSignature = mixCandidateSignature(candidateSignature, parent->totalBytes);
        candidateSignature = mixCandidateSignature(candidateSignature, static_cast<uint64_t>(maxItems));
        candidateSignature = mixCandidateSignature(candidateSignature, static_cast<uint64_t>(parent->scanState));

        const bool canUseCandidateCache = ! scanning && remaining >= maxItems;
        if (canUseCandidateCache)
        {
            const auto cacheIt = _layoutCandidateCache.find(parentId);
            if (cacheIt != _layoutCandidateCache.end())
            {
                const LayoutCandidateCacheEntry& entry = cacheIt->second;
                if (entry.signature == candidateSignature && entry.maxItems == static_cast<uint32_t>(maxItems))
                {
                    out.assign(entry.items.begin(), entry.items.end());
                    appendOtherBucket(entry.otherBytes, entry.otherWeight, entry.otherCount, entry.otherFolders, entry.otherFiles);
                    _layoutCandidateCacheHits += 1u;
                    return;
                }
            }
            _layoutCandidateCacheMisses += 1u;
        }

        auto minHeapByWeight = [](const Item& a, const Item& b) noexcept { return a.weight > b.weight; };

        bool autoExpanded = false;

        for (;;)
        {
            topItems.clear();
            topItems.reserve(maxItems);
            forcedItems.clear();
            forcedChildIds.clear();

            otherBytes   = 0;
            otherWeight  = 0.0;
            otherCount   = 0;
            otherFolders = 0;
            otherFiles   = 0;

            size_t normalSlots = maxItems;
            if (scanning && parentId == view.id)
            {
                const size_t forcedLimit = std::min<size_t>(maxItems, static_cast<size_t>(std::clamp(_config.scanThreads, 1u, 16u)));
                forcedChildIds.reserve(forcedLimit);

                for (const uint32_t childId : children)
                {
                    if (forcedChildIds.size() >= forcedLimit)
                    {
                        break;
                    }

                    const Node* child = TryGetRealNode(childId);
                    if (child == nullptr)
                    {
                        continue;
                    }

                    if (! child->isDirectory || child->isSynthetic || child->scanState != ScanState::Scanning)
                    {
                        continue;
                    }

                    forcedChildIds.push_back(childId);
                }

                forcedItems.reserve(forcedChildIds.size());
                normalSlots = maxItems > forcedChildIds.size() ? (maxItems - forcedChildIds.size()) : 0u;
            }

            auto isForcedChild = [&](uint32_t childId) noexcept -> bool
            {
                for (const uint32_t forcedId : forcedChildIds)
                {
                    if (forcedId == childId)
                    {
                        return true;
                    }
                }
                return false;
            };

            for (const uint32_t childId : children)
            {
                const std::optional<ResolvedItem> child = ResolveTreemapItem(childId);
                if (! child.has_value())
                {
                    continue;
                }

                Item item;
                item.nodeId = childId;
                item.bytes  = child->totalBytes;
                item.weight = static_cast<double>(std::max<uint64_t>(item.bytes, 1ull));

                if (isForcedChild(childId))
                {
                    forcedItems.push_back(item);
                    continue;
                }

                if (normalSlots == 0)
                {
                    otherBytes += item.bytes;
                    otherWeight += item.weight;
                    otherCount += 1;
                    addUnderlyingCounts(child.value());
                    continue;
                }

                if (topItems.size() < normalSlots)
                {
                    topItems.push_back(item);
                    std::push_heap(topItems.begin(), topItems.end(), minHeapByWeight);
                    continue;
                }

                if (! topItems.empty() && item.weight > topItems.front().weight)
                {
                    std::pop_heap(topItems.begin(), topItems.end(), minHeapByWeight);
                    const Item dropped = topItems.back();
                    topItems.back()    = item;
                    std::push_heap(topItems.begin(), topItems.end(), minHeapByWeight);

                    otherBytes += dropped.bytes;
                    otherWeight += dropped.weight;
                    otherCount += 1;
                    const std::optional<ResolvedItem> droppedNode = ResolveTreemapItem(dropped.nodeId);
                    if (droppedNode.has_value())
                    {
                        addUnderlyingCounts(droppedNode.value());
                    }
                    continue;
                }

                otherBytes += item.bytes;
                otherWeight += item.weight;
                otherCount += 1;
                addUnderlyingCounts(child.value());
            }

            if (! forcedItems.empty())
            {
                topItems.insert(topItems.end(), forcedItems.begin(), forcedItems.end());
            }

            if (parentId == view.id && ! autoExpanded && otherCount > 0 && otherWeight > 0.0 && maxItems < kMaxLayoutItems)
            {
                double totalWeight = otherWeight;
                for (const auto& top : topItems)
                {
                    totalWeight += top.weight;
                }

                const double ratio = totalWeight > 0.0 ? (otherWeight / totalWeight) : 0.0;
                if (ratio >= 0.62)
                {
                    if (_autoExpandedOtherByNode.emplace(parentId).second)
                    {
                        maxLayoutItems = std::min<uint32_t>(static_cast<uint32_t>(kMaxLayoutItems), maxLayoutItems * 2u);
                        maxItems       = static_cast<size_t>(maxLayoutItems);
                        capMaxItemsToBudget();
                        _layoutMaxItemsByNode[parentId] = maxLayoutItems;
                        autoExpanded                    = true;
                        continue;
                    }
                }
            }

            break;
        }

        if (topItems.empty())
        {
            return;
        }

        std::sort(topItems.begin(),
                  topItems.end(),
                  [](const Item& a, const Item& b) noexcept
        {
            if (a.bytes != b.bytes)
            {
                return a.bytes > b.bytes;
            }
            return a.nodeId < b.nodeId;
        });

        if (otherCount > 0 && remaining > 0)
        {
            while (! topItems.empty() && (topItems.size() + 1u) > remaining)
            {
                const Item dropped = topItems.back();
                topItems.pop_back();
                otherBytes += dropped.bytes;
                otherWeight += dropped.weight;
                otherCount += 1;
                const std::optional<ResolvedItem> droppedNode = ResolveTreemapItem(dropped.nodeId);
                if (droppedNode.has_value())
                {
                    addUnderlyingCounts(droppedNode.value());
                }
            }
        }

        if (canUseCandidateCache)
        {
            LayoutCandidateCacheEntry& entry = _layoutCandidateCache[parentId];
            entry.signature                  = candidateSignature;
            entry.maxItems                   = static_cast<uint32_t>(maxItems);
            entry.items                      = topItems;
            entry.otherBytes                 = otherBytes;
            entry.otherWeight                = otherWeight;
            entry.otherCount                 = otherCount;
            entry.otherFolders               = otherFolders;
            entry.otherFiles                 = otherFiles;
        }

        out.swap(topItems);
        appendOtherBucket(otherBytes, otherWeight, otherCount, otherFolders, otherFiles);
    };

    auto canExpand = [&](const Node& node, const D2D1_RECT_F& itemRc, uint8_t depth, float& outLabelHeight, D2D1_RECT_F& outChildrenRc) noexcept -> bool
    {
        if (! node.isDirectory || node.isSynthetic)
        {
            return false;
        }

        const float itemAreaDip2 = RectArea(itemRc);
        const float areaFraction = (viewAreaDip2 > 1.0f) ? (itemAreaDip2 / viewAreaDip2) : 0.0f;

        if (depth >= maxAutoExpandDepth)
        {
            return false;
        }

        if (areaFraction < autoExpandAreaFractionThreshold)
        {
            return false;
        }

        if (GetRealNodeChildren(node).empty())
        {
            return false;
        }

        const float h = std::max(0.0f, itemRc.bottom - itemRc.top);
        if (itemAreaDip2 < kMinExpandAreaDip2)
        {
            return false;
        }

        float label = 24.0f - static_cast<float>(depth) * 2.0f;
        label       = std::clamp(label, 20.0f, 24.0f);

        const bool incomplete = node.scanState == ScanState::NotStarted || node.scanState == ScanState::Queued || node.scanState == ScanState::Scanning;
        if (incomplete && areaFraction >= autoExpandAreaFractionThreshold)
        {
            const float twoLineLabel = std::clamp(38.0f - static_cast<float>(depth) * 2.0f, 30.0f, 38.0f);
            if (h >= twoLineLabel + kMinExpandChildSideDip)
            {
                label = std::max(label, twoLineLabel);
            }
        }

        if (h < label + kMinExpandChildSideDip)
        {
            return false;
        }

        D2D1_RECT_F content = itemRc;
        content.top += label;
        content.left += kNestedInsetDip;
        content.right -= kNestedInsetDip;
        content.top += kNestedInsetDip;
        content.bottom -= kNestedInsetDip;

        const float cw = std::max(0.0f, content.right - content.left);
        const float ch = std::max(0.0f, content.bottom - content.top);
        if (cw < kMinExpandChildSideDip || ch < kMinExpandChildSideDip)
        {
            return false;
        }

        if (RectArea(content) < kMinExpandChildAreaDip2)
        {
            return false;
        }

        outLabelHeight = label;
        outChildrenRc  = content;
        return true;
    };

    auto layoutNode = [&](auto&& self, uint32_t nodeId, const D2D1_RECT_F& bounds, uint8_t depth) noexcept -> void
    {
        if (remaining == 0)
        {
            return;
        }

        if (depth > 0)
        {
            const double elapsed = NowSeconds() - layoutStartSeconds;
            if (elapsed >= layoutBudgetSeconds)
            {
                return;
            }
        }

        const float w = std::max(0.0f, bounds.right - bounds.left);
        const float h = std::max(0.0f, bounds.bottom - bounds.top);
        if (w <= 1.0f || h <= 1.0f)
        {
            return;
        }

        const double boundsArea = static_cast<double>(w * h);

        if (_layoutWorkspaceByDepth.size() <= depth)
        {
            _layoutWorkspaceByDepth.resize(static_cast<size_t>(depth) + 1u);
        }

        LayoutWorkspace& workspace = _layoutWorkspaceByDepth[depth];

        std::vector<Item>& items = workspace.items;
        buildItemsForNode(nodeId, static_cast<float>(boundsArea), workspace, items);
        if (items.empty())
        {
            return;
        }

        const double totalWeight = std::accumulate(items.begin(), items.end(), 0.0, [](double sum, const Item& it) { return sum + it.weight; });
        if (totalWeight <= 0.0)
        {
            return;
        }

        const double scale      = boundsArea / totalWeight;

        std::vector<ExpandTask>& expandTasks = workspace.expandTasks;
        expandTasks.clear();

        auto worstAspectForWeights = [](double sumWeight, double minWeight, double maxWeight, double side, double scaleInner) noexcept -> double
        {
            if (sumWeight <= 0.0 || minWeight <= 0.0)
            {
                return std::numeric_limits<double>::infinity();
            }

            const double sumArea = sumWeight * scaleInner;
            const double maxArea = maxWeight * scaleInner;
            const double minArea = minWeight * scaleInner;

            const double side2 = side * side;
            return std::max((side2 * maxArea) / (sumArea * sumArea), (sumArea * sumArea) / (side2 * minArea));
        };

        auto layoutRow = [&](const std::vector<Item>& row, double rowWeight, D2D1_RECT_F& freeRc) noexcept
        {
            const float freeW     = std::max(0.0f, freeRc.right - freeRc.left);
            const float freeH     = std::max(0.0f, freeRc.bottom - freeRc.top);
            const bool horizontal = freeH > freeW;

            const double rowArea = rowWeight * scale;
            if (rowArea <= 0.0)
            {
                return;
            }

            if (horizontal)
            {
                const float rowH = static_cast<float>(rowArea / std::max(1.0f, freeW));
                float x          = freeRc.left;
                for (const auto& item : row)
                {
                    if (remaining == 0)
                    {
                        return;
                    }

                    const float itemW        = static_cast<float>((item.weight * scale) / std::max(1.0f, rowH));
                    const D2D1_RECT_F itemRc = D2D1::RectF(x, freeRc.top, x + itemW, freeRc.top + rowH);
                    x += itemW;

                    const Node* node  = getNode(item.nodeId);
                    float labelHeight = 0.0f;
                    D2D1_RECT_F childRc{};
                    const bool expanded = node != nullptr && canExpand(*node, itemRc, depth, labelHeight, childRc);
                    pushDrawItem(item.nodeId, itemRc, depth, expanded ? labelHeight : 0.0f);
                    if (expanded)
                    {
                        ExpandTask task;
                        task.nodeId = item.nodeId;
                        task.bounds = childRc;
                        task.area   = RectArea(childRc);
                        task.depth  = static_cast<uint8_t>(depth + 1);
                        expandTasks.push_back(task);
                    }
                }

                freeRc.top += rowH;
            }
            else
            {
                const float rowW = static_cast<float>(rowArea / std::max(1.0f, freeH));
                float y          = freeRc.top;
                for (const auto& item : row)
                {
                    if (remaining == 0)
                    {
                        return;
                    }

                    const float itemH        = static_cast<float>((item.weight * scale) / std::max(1.0f, rowW));
                    const D2D1_RECT_F itemRc = D2D1::RectF(freeRc.left, y, freeRc.left + rowW, y + itemH);
                    y += itemH;

                    const Node* node  = getNode(item.nodeId);
                    float labelHeight = 0.0f;
                    D2D1_RECT_F childRc{};
                    const bool expanded = node != nullptr && canExpand(*node, itemRc, depth, labelHeight, childRc);
                    pushDrawItem(item.nodeId, itemRc, depth, expanded ? labelHeight : 0.0f);
                    if (expanded)
                    {
                        ExpandTask task;
                        task.nodeId = item.nodeId;
                        task.bounds = childRc;
                        task.area   = RectArea(childRc);
                        task.depth  = static_cast<uint8_t>(depth + 1);
                        expandTasks.push_back(task);
                    }
                }

                freeRc.left += rowW;
            }
        };

        D2D1_RECT_F freeRc = bounds;
        std::vector<Item>& row = workspace.row;
        row.clear();
        double rowWeight    = 0.0;
        double rowMinWeight = 0.0;
        double rowMaxWeight = 0.0;
        size_t itemIndex    = 0;

        while (itemIndex < items.size())
        {
            const float freeW = std::max(0.0f, freeRc.right - freeRc.left);
            const float freeH = std::max(0.0f, freeRc.bottom - freeRc.top);
            const double side = static_cast<double>(std::min(freeW, freeH));
            if (side <= 1.0)
            {
                break;
            }

            const Item next = items[itemIndex];
            itemIndex += 1;

            if (row.empty())
            {
                row.push_back(next);
                rowWeight    = next.weight;
                rowMinWeight = next.weight;
                rowMaxWeight = next.weight;
                continue;
            }

            const double worstBefore   = worstAspectForWeights(rowWeight, rowMinWeight, rowMaxWeight, side, scale);
            const double nextRowWeight = rowWeight + next.weight;
            const double nextRowMin    = std::min(rowMinWeight, next.weight);
            const double nextRowMax    = std::max(rowMaxWeight, next.weight);
            const double worstAfter    = worstAspectForWeights(nextRowWeight, nextRowMin, nextRowMax, side, scale);

            if (worstAfter <= worstBefore)
            {
                row.push_back(next);
                rowWeight    = nextRowWeight;
                rowMinWeight = nextRowMin;
                rowMaxWeight = nextRowMax;
            }
            else
            {
                layoutRow(row, rowWeight, freeRc);
                row.clear();
                row.push_back(next);
                rowWeight    = next.weight;
                rowMinWeight = next.weight;
                rowMaxWeight = next.weight;
            }
        }

        if (! row.empty())
        {
            layoutRow(row, rowWeight, freeRc);
        }

        std::sort(expandTasks.begin(),
                  expandTasks.end(),
                  [](const ExpandTask& a, const ExpandTask& b) noexcept
        {
            if (a.area != b.area)
            {
                return a.area > b.area;
            }
            return a.nodeId < b.nodeId;
        });

        for (const ExpandTask& task : expandTasks)
        {
            if (remaining == 0)
            {
                break;
            }
            const double elapsed = NowSeconds() - layoutStartSeconds;
            if (elapsed >= layoutBudgetSeconds)
            {
                break;
            }
            self(self, task.nodeId, task.bounds, task.depth);
        }
    };

    layoutNode(layoutNode, view.id, rc, 0);
    BuildHitGrid(rc);
    if (! _lastRectsByNode.empty())
    {
        std::unordered_set<uint32_t> currentDrawItemIds;
        currentDrawItemIds.reserve(_drawItems.size());
        for (const DrawItem& item : _drawItems)
        {
            currentDrawItemIds.insert(item.nodeId);
        }

        for (auto it = _lastRectsByNode.begin(); it != _lastRectsByNode.end();)
        {
            if (currentDrawItemIds.contains(it->first))
            {
                ++it;
            }
            else
            {
                it = _lastRectsByNode.erase(it);
            }
        }
    }
#ifdef ENABLE_TESTS
    _layoutGeneration += 1u;
#endif

    if (_hoverNodeId != 0)
    {
        const bool stillVisible = std::any_of(_drawItems.begin(), _drawItems.end(), [this](const DrawItem& item) { return item.nodeId == _hoverNodeId; });
        if (! stillVisible)
        {
            _hoverNodeId = 0;
        }
    }

    const uint64_t layoutUs = Debug::Perf::ElapsedUs(layoutStartedAt);
    const uint64_t culledTiles =
        static_cast<uint64_t>(std::count_if(_drawItems.begin(), _drawItems.end(), [](const DrawItem& item) noexcept {
            return item.lod == DrawItem::Lod::Culled;
        }));
#ifdef ENABLE_TESTS
    _lastLayoutUs = layoutUs;
#endif
    Debug::Perf::Emit(L"viewer.space.layout.rebuild_us", L"phase0-current", layoutUs, static_cast<uint64_t>(_drawItems.size()), 0u, S_OK);
    Debug::Perf::EmitValue(L"viewer.space.layout.draw_items", static_cast<uint64_t>(_drawItems.size()));
    Debug::Perf::EmitValue(L"viewer.space.layout.visible_tiles", static_cast<uint64_t>(_drawItems.size()) - culledTiles);
    Debug::Perf::EmitValue(L"viewer.space.layout.culled_tiles", culledTiles);
    Debug::Perf::EmitValue(L"viewer.space.layout.candidate_cache_hits", _layoutCandidateCacheHits);
    Debug::Perf::EmitValue(L"viewer.space.layout.candidate_cache_misses", _layoutCandidateCacheMisses);
}

void ViewerSpace::BuildHitGrid(const D2D1_RECT_F& bounds) noexcept
{
    _hitGrid.Clear();
    const float boundsW = std::max(0.0f, bounds.right - bounds.left);
    const float boundsH = std::max(0.0f, bounds.bottom - bounds.top);
    if (boundsW <= 1.0f || boundsH <= 1.0f || _drawItems.size() < 256u)
    {
        return;
    }

    const uint32_t columns = static_cast<uint32_t>(std::clamp(std::ceil(static_cast<double>(boundsW) / 96.0), 8.0, 64.0));
    const uint32_t rows    = static_cast<uint32_t>(std::clamp(std::ceil(static_cast<double>(boundsH) / 96.0), 8.0, 64.0));
    _hitGrid.bounds        = bounds;
    _hitGrid.columns       = columns;
    _hitGrid.rows          = rows;
    _hitGrid.cells.resize(static_cast<size_t>(columns) * static_cast<size_t>(rows));

    auto cellIndex = [columns](uint32_t x, uint32_t y) noexcept -> size_t
    {
        return static_cast<size_t>(y) * static_cast<size_t>(columns) + static_cast<size_t>(x);
    };

    for (size_t itemIndex = 0; itemIndex < _drawItems.size(); ++itemIndex)
    {
        const DrawItem& item = _drawItems[itemIndex];
        if (item.lod == DrawItem::Lod::Culled)
        {
            continue;
        }

        float gap = kItemGapDip - static_cast<float>(item.depth) * 0.15f;
        gap       = std::clamp(gap, 0.5f, kItemGapDip);

        D2D1_RECT_F rc = item.targetRect;
        rc.left += gap;
        rc.top += gap;
        rc.right -= gap;
        rc.bottom -= gap;
        if (rc.right <= rc.left || rc.bottom <= rc.top || RectArea(rc) < kMinHitAreaDip2)
        {
            continue;
        }

        const auto clampCellX = [&](float x) noexcept -> uint32_t
        {
            const float t = (x - bounds.left) / std::max(1.0f, boundsW);
            return static_cast<uint32_t>(
                std::clamp(std::floor(static_cast<double>(t) * static_cast<double>(columns)), 0.0, static_cast<double>(columns - 1u)));
        };
        const auto clampCellY = [&](float y) noexcept -> uint32_t
        {
            const float t = (y - bounds.top) / std::max(1.0f, boundsH);
            return static_cast<uint32_t>(
                std::clamp(std::floor(static_cast<double>(t) * static_cast<double>(rows)), 0.0, static_cast<double>(rows - 1u)));
        };

        const uint32_t leftCell   = clampCellX(rc.left);
        const uint32_t rightCell  = clampCellX(rc.right);
        const uint32_t topCell    = clampCellY(rc.top);
        const uint32_t bottomCell = clampCellY(rc.bottom);

        for (uint32_t y = topCell; y <= bottomCell; ++y)
        {
            for (uint32_t x = leftCell; x <= rightCell; ++x)
            {
                std::vector<uint32_t>& cell = _hitGrid.cells[cellIndex(x, y)];
                cell.push_back(static_cast<uint32_t>(itemIndex));
                _hitGrid.maxCandidatesPerCell = std::max<uint32_t>(_hitGrid.maxCandidatesPerCell, static_cast<uint32_t>(cell.size()));
            }
        }
    }

    Debug::Perf::EmitValue(L"viewer.space.hit_grid.cells", static_cast<uint64_t>(_hitGrid.cells.size()));
    Debug::Perf::EmitValue(L"viewer.space.hit_grid.max_candidates", _hitGrid.maxCandidatesPerCell);
}

std::optional<uint32_t> ViewerSpace::HitTestTreemap(float xDip, float yDip) const noexcept
{
    const auto hitStartedAt = std::chrono::steady_clock::now();
    uint32_t candidatesChecked = 0u;
    std::wstring_view hitDetail = L"linear";
    auto emitHitMetric         = wil::scope_exit([&]() noexcept
    {
        const uint64_t hitUs = Debug::Perf::ElapsedUs(hitStartedAt);
#ifdef ENABLE_TESTS
        _lastHitTestUs                = hitUs;
        _lastHitTestCandidatesChecked = candidatesChecked;
#endif
        Debug::Perf::Emit(L"viewer.space.hit_test_us", hitDetail, hitUs, static_cast<uint64_t>(_drawItems.size()), candidatesChecked, S_OK);
    });

    if (yDip < GetHeaderBottomDip())
    {
        return std::nullopt;
    }

    if (HasActiveLayoutAnimation(NowSeconds()))
    {
        hitDetail = L"linear-animation";
        return HitTestTreemapLinear(xDip, yDip, candidatesChecked);
    }

    if (_hitGrid.columns > 0u && _hitGrid.rows > 0u && ! _hitGrid.cells.empty() && xDip >= _hitGrid.bounds.left && xDip <= _hitGrid.bounds.right &&
        yDip >= _hitGrid.bounds.top && yDip <= _hitGrid.bounds.bottom)
    {
        hitDetail = L"grid";
        const float boundsW = std::max(1.0f, _hitGrid.bounds.right - _hitGrid.bounds.left);
        const float boundsH = std::max(1.0f, _hitGrid.bounds.bottom - _hitGrid.bounds.top);
        const uint32_t cellX =
            static_cast<uint32_t>(std::clamp(
                std::floor((static_cast<double>(xDip - _hitGrid.bounds.left) / static_cast<double>(boundsW)) * static_cast<double>(_hitGrid.columns)),
                0.0,
                static_cast<double>(_hitGrid.columns - 1u)));
        const uint32_t cellY =
            static_cast<uint32_t>(std::clamp(
                std::floor((static_cast<double>(yDip - _hitGrid.bounds.top) / static_cast<double>(boundsH)) * static_cast<double>(_hitGrid.rows)),
                0.0,
                static_cast<double>(_hitGrid.rows - 1u)));
        const size_t index = static_cast<size_t>(cellY) * static_cast<size_t>(_hitGrid.columns) + static_cast<size_t>(cellX);
        if (index < _hitGrid.cells.size())
        {
            const std::vector<uint32_t>& candidates = _hitGrid.cells[index];
            for (size_t i = candidates.size(); i-- > 0;)
            {
                const uint32_t itemIndex = candidates[i];
                if (itemIndex >= _drawItems.size())
                {
                    continue;
                }

                const DrawItem& item = _drawItems[itemIndex];
                candidatesChecked += 1u;

                float gap = kItemGapDip - static_cast<float>(item.depth) * 0.15f;
                gap       = std::clamp(gap, 0.5f, kItemGapDip);

                D2D1_RECT_F rc = item.currentRect;
                rc.left += gap;
                rc.top += gap;
                rc.right -= gap;
                rc.bottom -= gap;
                if (rc.right <= rc.left || rc.bottom <= rc.top)
                {
                    continue;
                }

                if (xDip >= rc.left && xDip <= rc.right && yDip >= rc.top && yDip <= rc.bottom)
                {
                    const float area = RectArea(rc);
                    if (area >= kMinHitAreaDip2)
                    {
                        return item.nodeId;
                    }
                }
            }

            hitDetail = L"grid-fallback-linear";
            uint32_t linearCandidates = 0u;
            std::optional<uint32_t> linearHit = HitTestTreemapLinear(xDip, yDip, linearCandidates);
            candidatesChecked += linearCandidates;
            return linearHit;
        }
    }

    std::optional<uint32_t> linearHit = HitTestTreemapLinear(xDip, yDip, candidatesChecked);
    return linearHit;
}

std::optional<uint32_t> ViewerSpace::HitTestTreemapLinear(float xDip, float yDip, uint32_t& candidatesChecked) const noexcept
{
    candidatesChecked = 0u;
    if (yDip < GetHeaderBottomDip())
    {
        return std::nullopt;
    }

    for (size_t i = _drawItems.size(); i-- > 0;)
    {
        const DrawItem& item = _drawItems[i];
        candidatesChecked += 1u;

        float gap = kItemGapDip - static_cast<float>(item.depth) * 0.15f;
        gap       = std::clamp(gap, 0.5f, kItemGapDip);

        D2D1_RECT_F rc = item.currentRect;
        rc.left += gap;
        rc.top += gap;
        rc.right -= gap;
        rc.bottom -= gap;
        if (rc.right <= rc.left || rc.bottom <= rc.top)
        {
            continue;
        }

        if (xDip >= rc.left && xDip <= rc.right && yDip >= rc.top && yDip <= rc.bottom)
        {
            const float area = RectArea(rc);
            if (area >= kMinHitAreaDip2)
            {
                return item.nodeId;
            }
        }
    }

    return std::nullopt;
}

void ViewerSpace::NavigateTo(uint32_t nodeId) noexcept
{
    if (nodeId == 0 || nodeId == _viewNodeId)
    {
        return;
    }

    if (TryGetRealNode(nodeId) == nullptr)
    {
        return;
    }

    _navStack.push_back(_viewNodeId);
    _viewNodeId = nodeId;
    UpdateViewPathText();
    _layoutDirty = true;
    InvalidateStaticTreemapCache();

    if (_hWnd)
    {
        UpdateWindowTitle(_hWnd.get());
        UpdateMenuState(_hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

bool ViewerSpace::CanNavigateUp() const noexcept
{
    if (! _navStack.empty())
    {
        return true;
    }

    const Node* view = TryGetRealNode(_viewNodeId);
    if (view != nullptr && view->parentId != 0)
    {
        return true;
    }

    return _scanRootParentPath.has_value();
}

void ViewerSpace::UpdateMenuState(HWND hwnd, bool syncDxMenuBar) noexcept
{
    HMENU menu = _menuHandle ? _menuHandle.get() : (hwnd ? GetMenu(hwnd) : nullptr);
    if (! menu)
    {
        return;
    }

    const UINT state = CanNavigateUp() ? MF_ENABLED : MF_GRAYED;
    EnableMenuItem(menu, IDM_VIEWERSPACE_NAV_UP, MF_BYCOMMAND | state);
    if (syncDxMenuBar && _menuBarHost.GetHwnd())
    {
        _menuBarHost.SyncMenuModel();
    }
    else if (hwnd)
    {
        DrawMenuBar(hwnd);
    }
}

float ViewerSpace::GetMenuBarHeightDip() const noexcept
{
    if (! _menuBarHost.GetHwnd())
    {
        return 0.0f;
    }

    const float dpi = _dpi > 0.0f ? _dpi : static_cast<float>(USER_DEFAULT_SCREEN_DPI);
    return static_cast<float>(_menuBarHost.GetVisibleHeightPx()) * static_cast<float>(USER_DEFAULT_SCREEN_DPI) / dpi;
}

float ViewerSpace::GetHeaderTopDip() const noexcept
{
    return GetMenuBarHeightDip();
}

float ViewerSpace::GetHeaderBottomDip() const noexcept
{
    return GetHeaderTopDip() + kHeaderHeightDip;
}

void ViewerSpace::NavigateUp() noexcept
{
    if (! CanNavigateUp())
    {
        return;
    }

    uint32_t nextNode = 0;
    if (! _navStack.empty())
    {
        nextNode = _navStack.back();
        _navStack.pop_back();
    }
    else
    {
        const Node* view = TryGetRealNode(_viewNodeId);
        if (view != nullptr && view->parentId != 0)
        {
            nextNode = view->parentId;
        }
    }

    if (nextNode == 0)
    {
        if (_scanRootParentPath.has_value())
        {
            StartScan(_scanRootParentPath.value());
        }
        return;
    }

    if (nextNode == _viewNodeId)
    {
        return;
    }

    _viewNodeId = nextNode;
    UpdateViewPathText();
    _layoutDirty = true;
    InvalidateStaticTreemapCache();

    if (_hWnd)
    {
        UpdateWindowTitle(_hWnd.get());
        UpdateMenuState(_hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }
}

void ViewerSpace::RefreshCurrent() noexcept
{
    const std::wstring rootPath = BuildNodePathText(_viewNodeId);
    if (rootPath.empty())
    {
        return;
    }

    StartScan(rootPath, false);
}

float ViewerSpace::DipFromPx(int px) const noexcept
{
    return static_cast<float>(px) * (96.0f / _dpi);
}

int ViewerSpace::PxFromDip(float dip) const noexcept
{
    const double scale = static_cast<double>(_dpi) / 96.0;
    return static_cast<int>(std::lround(static_cast<double>(dip) * scale));
}

double ViewerSpace::NowSeconds() const noexcept
{
    using clock             = std::chrono::steady_clock;
    static const auto start = clock::now();
    const auto now          = clock::now();
    return std::chrono::duration<double>(now - start).count();
}

HRESULT STDMETHODCALLTYPE ViewerSpace::Open(const ViewerOpenContext* context) noexcept
{
    if (context == nullptr || context->fileSystem == nullptr || context->focusedPath == nullptr || context->focusedPath[0] == L'\0')
    {
        return E_INVALIDARG;
    }

    const bool embeddedMode   = IsEmbeddedOpen(*context);
    const HWND embeddedParent = embeddedMode ? context->ownerWindow : nullptr;
    if (embeddedMode && (! embeddedParent || IsWindow(embeddedParent) == FALSE))
    {
        return E_INVALIDARG;
    }
    if (ShouldRecreateViewerWindow(embeddedMode, embeddedParent))
    {
        static_cast<void>(Close());
    }

    _fileSystem = context->fileSystem;

    _fileSystemName = context->fileSystemName ? context->fileSystemName : L"";
    _fileSystemShortId.clear();
    _fileSystemIsWin32 = true;

    wil::com_ptr<IInformations> fileSystemInfo;
    if (_fileSystem.try_query_to(fileSystemInfo.put()) && fileSystemInfo)
    {
        const PluginMetaData* metaData = nullptr;
        if (SUCCEEDED(fileSystemInfo->GetMetaData(&metaData)) && metaData != nullptr && metaData->shortId != nullptr)
        {
            _fileSystemShortId = metaData->shortId;
            _fileSystemIsWin32 = _fileSystemShortId == L"file";
        }
    }

    if (! _hWnd)
    {
        if (! RegisterWndClass(g_hInstance))
        {
            return E_FAIL;
        }

        HWND ownerWindow = context->ownerWindow;
        RECT ownerRc{};
        int x = CW_USEDEFAULT;
        int y = CW_USEDEFAULT;
        int w = 1000;
        int h = 700;
        if (embeddedMode)
        {
            RECT client{};
            if (GetClientRect(embeddedParent, &client) != 0)
            {
                x = 0;
                y = 0;
                w = std::max(1L, client.right - client.left);
                h = std::max(1L, client.bottom - client.top);
            }
        }
        else if (ownerWindow && GetWindowRect(ownerWindow, &ownerRc) != 0)
        {
            x = ownerRc.left;
            y = ownerRc.top;
            w = std::max(1L, ownerRc.right - ownerRc.left);
            h = std::max(1L, ownerRc.bottom - ownerRc.top);
        }

        HMENU loadedMenu = embeddedMode ? nullptr : Localization::LoadMenuResource(g_hInstance, IDR_VIEWERSPACE_MENU);
        wil::unique_any<HMENU, decltype(&::DestroyMenu), ::DestroyMenu> menu(loadedMenu);
        const DWORD style               = embeddedMode ? (WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS) : WS_OVERLAPPEDWINDOW;
        const bool previousEmbeddedMode = _embeddedMode;
        _embeddedMode                   = embeddedMode;
        const std::wstring initialTitle =
            embeddedMode ? std::wstring{} : (_metaName.empty() ? LoadStringResource(g_hInstance, IDS_VIEWERSPACE_NAME) : _metaName);
        HWND window =
            CreateWindowExW(0, kClassName, initialTitle.c_str(), style, x, y, w, h, embeddedMode ? embeddedParent : nullptr, menu.get(), g_hInstance, this);
        if (! window)
        {
            _embeddedMode = previousEmbeddedMode;
            return HRESULT_FROM_WIN32(GetLastError());
        }

        if (! embeddedMode)
        {
            menu.release();
        }
        _hWnd.reset(window);

        ApplyThemeToWindow(_hWnd.get());
        AddRef(); // Self-reference for window lifetime (released in WM_NCDESTROY)
        ShowWindow(_hWnd.get(), embeddedMode ? SW_SHOWNA : SW_SHOWNORMAL);
        if (! embeddedMode)
        {
            static_cast<void>(SetForegroundWindow(_hWnd.get()));
        }
    }
    else
    {
        if (embeddedMode)
        {
            RECT client{};
            if (GetClientRect(embeddedParent, &client) != 0)
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

    StartScan(context->focusedPath);
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerSpace::Close() noexcept
{
    AddRef();
    const auto releaseSelf = wil::scope_exit([&]() noexcept { Release(); });
    _hWnd.reset();
    return S_OK;
}

HRESULT STDMETHODCALLTYPE ViewerSpace::SetTheme(const ViewerTheme* theme) noexcept
{
    if (theme == nullptr || theme->version < 2u || theme->version > 4u)
    {
        return E_INVALIDARG;
    }

    _theme    = *theme;
    _hasTheme = true;
    _dpi      = static_cast<float>(_theme.dpi == 0 ? USER_DEFAULT_SCREEN_DPI : _theme.dpi);

    RequestViewerSpaceClassBackgroundColor(ColorRefFromArgb(_theme.backgroundArgb));
    ApplyPendingViewerSpaceClassBackgroundBrush(_hWnd.get());

    _layoutDirty = true;
#ifdef ENABLE_TESTS
    _rendererFailureStage = kRendererDiscardTheme;
    _rendererFailureHr    = 0u;
#endif
    DiscardDirect2D();

    if (_hWnd)
    {
        ApplyThemeToWindow(_hWnd.get());
        InvalidateRect(_hWnd.get(), nullptr, FALSE);
    }

    return S_OK;
}
