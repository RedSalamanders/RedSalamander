// Preferences.Internal.cpp

#include "Framework.h"

#include "Preferences.Internal.h"

#include <algorithm>
#include <array>
#include <chrono>
#include <cwchar>
#include <cwctype>
#include <fstream>
#include <limits>
#include <string>

#include <uxtheme.h>

#include "FileSystemPluginManager.h"
#include "UiMetrics.h"
#include "ViewerPluginManager.h"
#include "WindowMessages.h"

namespace
{
constexpr wchar_t kPrefsDxHostProp[]                = L"RedSalamander.Preferences.DxHost";
constexpr wchar_t kPrefsDxHostOriginalWndProcProp[] = L"RedSalamander.Preferences.DxHostOriginalWndProc";

[[maybe_unused]] [[nodiscard]] WNDPROC GetStoredWndProc(HWND hwnd, const wchar_t* propName) noexcept
{
    return RedSalamander::Win32Callback::GetStoredWndProc(hwnd, propName);
}

[[nodiscard]] bool InstallWndProcHook(HWND hwnd, const wchar_t* originalWndProcProp, WNDPROC hookWndProc) noexcept
{
    return RedSalamander::Win32Callback::InstallWndProcHook(hwnd, originalWndProcProp, hookWndProc);
}

void RestoreWndProcHook(HWND hwnd, const wchar_t* originalWndProcProp) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return;
    }

    RedSalamander::Win32Callback::RestoreWndProcHook(hwnd, originalWndProcProp);
}

[[nodiscard]] LRESULT CallStoredWndProc(HWND hwnd, const wchar_t* originalWndProcProp, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    return RedSalamander::Win32Callback::CallStoredWndProc(hwnd, originalWndProcProp, msg, wp, lp);
}

void ApplyAncestorRedrawSuppression(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return;
    }

    for (HWND ancestor = hwnd; ancestor; ancestor = GetParent(ancestor))
    {
        const auto redrawBlockCount = reinterpret_cast<ULONG_PTR>(GetPropW(ancestor, PrefsUi::kPrefsTreeRedrawBlockProp));
        if (redrawBlockCount != 0u)
        {
            SendMessageW(hwnd, WM_SETREDRAW, FALSE, 0);
            break;
        }
    }
}

LRESULT CALLBACK PrefsDxHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    auto* host = reinterpret_cast<RedSalamander::DxUi::WindowHost*>(GetPropW(hwnd, kPrefsDxHostProp));
    if (! host)
    {
        return CallStoredWndProc(hwnd, kPrefsDxHostOriginalWndProcProp, msg, wp, lp);
    }

    if (msg == WM_GETDLGCODE)
    {
        LRESULT dlgCode = CallStoredWndProc(hwnd, kPrefsDxHostOriginalWndProcProp, msg, wp, lp) | DLGC_WANTARROWS | DLGC_WANTCHARS;
        if (PrefsKeyboardCaptureWantsAllKeys(PrefsUi::GetDialogState(hwnd)))
        {
            dlgCode |= DLGC_WANTALLKEYS;
        }
        return dlgCode;
    }

    if (PrefsHandleKeyboardCaptureMessage(hwnd, msg, wp, lp))
    {
        return 0;
    }

    if (msg == WM_NCDESTROY)
    {
        const auto originalWndProc = RedSalamander::Win32Callback::GetStoredWndProc(hwnd, kPrefsDxHostOriginalWndProcProp);
        RemovePropW(hwnd, kPrefsDxHostProp);
        RedSalamander::Win32Callback::RestoreWndProcHook(hwnd, kPrefsDxHostOriginalWndProcProp, PrefsDxHostWndProc);
        host->ReleaseMouseCapture();
        return originalWndProc ? RedSalamander::Win32Callback::CallWindowProcNoThrow(originalWndProc, hwnd, msg, wp, lp) : DefWindowProcW(hwnd, msg, wp, lp);
    }

    bool handled           = false;
    const LRESULT dxResult = host->HandleMessage(hwnd, msg, wp, lp, handled);

    return handled ? dxResult : CallStoredWndProc(hwnd, kPrefsDxHostOriginalWndProcProp, msg, wp, lp);
}

} // namespace

void PrefsReorderPanelChildren(RedSalamander::DxUi::Panel* root, std::span<RedSalamander::DxUi::Control* const> orderedControls)
{
    if (! root)
    {
        return;
    }

    auto children = root->GetChildren();
    if (children.empty())
    {
        return;
    }

    std::vector<std::unique_ptr<RedSalamander::DxUi::Control>> reordered;
    reordered.reserve(children.size());

    auto moveChild = [&](RedSalamander::DxUi::Control* wanted) noexcept
    {
        if (! wanted)
        {
            return;
        }

        for (auto& child : children)
        {
            if (child && child.get() == wanted)
            {
                reordered.push_back(std::move(child));
                return;
            }
        }
    };

    for (auto* control : orderedControls)
    {
        moveChild(control);
    }

    for (auto& child : children)
    {
        if (child)
        {
            reordered.push_back(std::move(child));
        }
    }

    if (reordered.size() != children.size())
    {
        return;
    }

    for (size_t index = 0; index < children.size(); ++index)
    {
        children[index] = std::move(reordered[index]);
    }
}

namespace PrefsPageHost
{
namespace
{
void SyncDxScrollPanelOffset(HWND pageHostWindow, PreferencesDialogState& state) noexcept
{
    if (! pageHostWindow || ! state.pageHostUsesDxUi)
    {
        return;
    }

    const UINT dpi = GetDpiForWindow(pageHostWindow);
    if (dpi == 0u)
    {
        return;
    }

    const float scrollDip = (static_cast<float>(state.pageScrollY) * 96.0f) / static_cast<float>(dpi);
    if (auto* scrollPanel = state.pageHostDxScrollPanelControl)
    {
        scrollPanel->SetScrollOffset(scrollDip);
    }
}
} // namespace

void ApplyScrollDelta(HWND pageHostWindow, int dy, bool syncScrollPanel) noexcept
{
    if (! pageHostWindow)
    {
        return;
    }

    PreferencesDialogState* state = reinterpret_cast<PreferencesDialogState*>(GetWindowLongPtrW(pageHostWindow, GWLP_USERDATA));
    if (dy == 0)
    {
        if (syncScrollPanel && state)
        {
            SyncDxScrollPanelOffset(pageHostWindow, *state);
        }
        return;
    }

    const auto scrollApplyStartedAt = std::chrono::steady_clock::now();
    Debug::Perf::Scope scrollApplyPerf(L"preferences.page_host.scroll_apply_us");
    int childCount = 0;
    for (HWND child = GetWindow(pageHostWindow, GW_CHILD); child; child = GetWindow(child, GW_HWNDNEXT))
    {
        ++childCount;
    }

#ifdef ENABLE_TESTS
    uint64_t dxMovedControlCount = 0u;
    const auto recordApply       = [&]() noexcept
    {
        if (! state)
        {
            return;
        }

        ++state->pageHostScrollApplyCount;
        state->pageHostScrollMovedChildCountTotal += static_cast<uint64_t>(childCount);
        state->pageHostScrollLastApplyUs = Debug::Perf::ElapsedUs(scrollApplyStartedAt);
    };
#else
    const auto recordApply = []() noexcept {};
#endif

    scrollApplyPerf.SetValue0(static_cast<uint64_t>(childCount));
    scrollApplyPerf.SetValue1(static_cast<uint64_t>(std::abs(dy)));

    HDWP hdwp = BeginDeferWindowPos(std::max(1, childCount));
    if (! hdwp)
    {
        for (HWND child = GetWindow(pageHostWindow, GW_CHILD); child; child = GetWindow(child, GW_HWNDNEXT))
        {
            RECT rc{};
            if (! GetWindowRect(child, &rc))
            {
                continue;
            }
            MapWindowPoints(nullptr, pageHostWindow, reinterpret_cast<POINT*>(&rc), 2);
            SetWindowPos(child, nullptr, rc.left, rc.top + dy, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOCOPYBITS);
        }
        recordApply();
        return;
    }

    for (HWND child = GetWindow(pageHostWindow, GW_CHILD); child; child = GetWindow(child, GW_HWNDNEXT))
    {
        RECT rc{};
        if (! GetWindowRect(child, &rc))
        {
            continue;
        }
        MapWindowPoints(nullptr, pageHostWindow, reinterpret_cast<POINT*>(&rc), 2);
        hdwp = DeferWindowPos(hdwp, child, nullptr, rc.left, rc.top + dy, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOCOPYBITS);
        if (! hdwp)
        {
            break;
        }
    }

    if (hdwp)
    {
        static_cast<void>(EndDeferWindowPos(hdwp));
    }

    recordApply();

    if (! syncScrollPanel || ! state || ! state->pageHostUsesDxUi)
    {
        return;
    }

    if (state->pageHostDxScrollPanelControl)
    {
        SyncDxScrollPanelOffset(pageHostWindow, *state);
#ifdef ENABLE_TESTS
        ++dxMovedControlCount;
#endif
    }
#ifdef ENABLE_TESTS
    state->pageHostDxScrollMovedControlCountTotal += dxMovedControlCount;
    state->pageHostDxScrollLastMovedControlCount = dxMovedControlCount;
#endif
}

void ApplyScrollToPosition(HWND pageHostWindow, PreferencesDialogState& state, int newScrollY) noexcept
{
    if (! pageHostWindow)
    {
        return;
    }

    newScrollY = std::clamp(newScrollY, 0, state.pageScrollMaxY);
    if (newScrollY == state.pageScrollY)
    {
        ApplyScrollDelta(pageHostWindow, 0);
        return;
    }

    const int oldScrollY = state.pageScrollY;
    state.pageScrollY    = newScrollY;

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_POS;
    si.nPos   = state.pageScrollY;
    SetScrollInfo(pageHostWindow, SB_VERT, &si, TRUE);

    const int dy = oldScrollY - state.pageScrollY;
    ApplyScrollDelta(pageHostWindow, dy);
    // Avoid RDW_UPDATENOW — let Windows coalesce into a single WM_PAINT.
    RedrawWindow(pageHostWindow, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_FRAME);
}

void ScrollTo(HWND pageHostWindow, PreferencesDialogState& state, int newScrollY) noexcept
{
    if (! pageHostWindow)
    {
        return;
    }

#ifdef ENABLE_TESTS
    ++state.pageHostScrollRequestCount;
#endif
    state.pageHostScrollApplyPending = false;
    ApplyScrollToPosition(pageHostWindow, state, newScrollY);
}

void RequestScrollTo(HWND pageHostWindow, PreferencesDialogState& state, int newScrollY) noexcept
{
    if (! pageHostWindow)
    {
        return;
    }

#ifdef ENABLE_TESTS
    ++state.pageHostScrollRequestCount;
#endif
    newScrollY = std::clamp(newScrollY, 0, state.pageScrollMaxY);
    if (state.pageHostScrollApplyPending)
    {
#ifdef ENABLE_TESTS
        ++state.pageHostScrollCoalescedRequestCount;
#endif
        state.pageHostPendingScrollY = newScrollY;
        return;
    }

    if (newScrollY == state.pageScrollY)
    {
        return;
    }

    state.pageHostPendingScrollY     = newScrollY;
    state.pageHostScrollApplyPending = true;
    if (PostMessageW(pageHostWindow, WndMsg::kPreferencesApplyPageHostScroll, 0, 0) == FALSE)
    {
        state.pageHostScrollApplyPending = false;
        ApplyScrollToPosition(pageHostWindow, state, newScrollY);
    }
}

void FlushPendingScroll(HWND pageHostWindow, PreferencesDialogState& state) noexcept
{
    if (! pageHostWindow || ! state.pageHostScrollApplyPending)
    {
        return;
    }

    const int pendingScrollY         = state.pageHostPendingScrollY;
    state.pageHostScrollApplyPending = false;
    ApplyScrollToPosition(pageHostWindow, state, pendingScrollY);
}

void EnsureControlVisible(HWND pageHostWindow, PreferencesDialogState& state, HWND control) noexcept
{
    if (! pageHostWindow || ! control || state.pageScrollMaxY <= 0)
    {
        return;
    }

    RECT rc{};
    if (! GetWindowRect(control, &rc))
    {
        return;
    }

    MapWindowPoints(nullptr, pageHostWindow, reinterpret_cast<POINT*>(&rc), 2);

    RECT client{};
    GetClientRect(pageHostWindow, &client);

    const UINT dpi          = GetDpiForWindow(pageHostWindow);
    const int padY          = UiMetrics::ScaleDip(dpi, 10);
    const int desiredTop    = client.top + padY;
    const int desiredBottom = client.bottom - padY;

    int newScrollY = state.pageScrollY;
    if (rc.top < desiredTop)
    {
        newScrollY = state.pageScrollY + (rc.top - desiredTop);
    }
    else if (rc.bottom > desiredBottom)
    {
        newScrollY = state.pageScrollY + (rc.bottom - desiredBottom);
    }

    ScrollTo(pageHostWindow, state, newScrollY);
}
} // namespace PrefsPageHost

namespace PrefsPlugins
{
[[nodiscard]] int CompareTextNoCase(std::wstring_view left, std::wstring_view right) noexcept
{
    if (left.empty() && right.empty())
    {
        return 0;
    }

    const int compare = CompareStringOrdinal(left.data(), static_cast<int>(left.size()), right.data(), static_cast<int>(right.size()), TRUE);
    if (compare == CSTR_LESS_THAN)
    {
        return -1;
    }
    if (compare == CSTR_GREATER_THAN)
    {
        return 1;
    }
    return 0;
}

void BuildListItems(std::vector<PrefsPluginListItem>& out) noexcept
{
    out.clear();

    const auto& fsPlugins = FileSystemPluginManager::GetInstance().GetPlugins();
    out.reserve(fsPlugins.size());
    for (size_t i = 0; i < fsPlugins.size(); ++i)
    {
        if (! fsPlugins[i].id.empty())
        {
            out.push_back(PrefsPluginListItem{PrefsPluginType::FileSystem, i});
        }
    }

    const auto& viewerPlugins = ViewerPluginManager::GetInstance().GetPlugins();
    out.reserve(out.size() + viewerPlugins.size());
    for (size_t i = 0; i < viewerPlugins.size(); ++i)
    {
        if (! viewerPlugins[i].id.empty())
        {
            out.push_back(PrefsPluginListItem{PrefsPluginType::Viewer, i});
        }
    }

    std::sort(out.begin(),
              out.end(),
              [](const PrefsPluginListItem& a, const PrefsPluginListItem& b) noexcept
    {
        const std::wstring_view aName = GetDisplayName(a);
        const std::wstring_view bName = GetDisplayName(b);
        const int nameCompare         = CompareTextNoCase(aName, bName);
        if (nameCompare != 0)
        {
            return nameCompare < 0;
        }

        const std::wstring_view aId = GetId(a);
        const std::wstring_view bId = GetId(b);
        const int idCompare         = CompareTextNoCase(aId, bId);
        if (idCompare != 0)
        {
            return idCompare < 0;
        }

        if (a.type != b.type)
        {
            return a.type < b.type;
        }

        const int aOrigin = GetOriginOrder(a);
        const int bOrigin = GetOriginOrder(b);
        if (aOrigin != bOrigin)
        {
            return aOrigin < bOrigin;
        }

        return a.index < b.index;
    });
}

[[nodiscard]] std::optional<PrefsPluginListItem> FindItemById(std::wstring_view pluginId) noexcept
{
    if (pluginId.empty())
    {
        return std::nullopt;
    }

    const auto& fsPlugins = FileSystemPluginManager::GetInstance().GetPlugins();
    for (size_t i = 0; i < fsPlugins.size(); ++i)
    {
        if (CompareTextNoCase(fsPlugins[i].id, pluginId) == 0)
        {
            return PrefsPluginListItem{PrefsPluginType::FileSystem, i};
        }
    }

    const auto& viewerPlugins = ViewerPluginManager::GetInstance().GetPlugins();
    for (size_t i = 0; i < viewerPlugins.size(); ++i)
    {
        if (CompareTextNoCase(viewerPlugins[i].id, pluginId) == 0)
        {
            return PrefsPluginListItem{PrefsPluginType::Viewer, i};
        }
    }

    return std::nullopt;
}

[[nodiscard]] std::wstring_view GetId(const PrefsPluginListItem& item) noexcept
{
    if (item.type == PrefsPluginType::FileSystem)
    {
        const auto& plugins = FileSystemPluginManager::GetInstance().GetPlugins();
        return (item.index < plugins.size()) ? std::wstring_view(plugins[item.index].id) : std::wstring_view{};
    }

    const auto& plugins = ViewerPluginManager::GetInstance().GetPlugins();
    return (item.index < plugins.size()) ? std::wstring_view(plugins[item.index].id) : std::wstring_view{};
}

[[nodiscard]] std::wstring_view GetDisplayName(const PrefsPluginListItem& item) noexcept
{
    if (item.type == PrefsPluginType::FileSystem)
    {
        const auto& plugins = FileSystemPluginManager::GetInstance().GetPlugins();
        if (item.index >= plugins.size())
        {
            return {};
        }

        const auto& plugin = plugins[item.index];
        return plugin.name.empty() ? std::wstring_view(plugin.id) : std::wstring_view(plugin.name);
    }

    const auto& plugins = ViewerPluginManager::GetInstance().GetPlugins();
    if (item.index >= plugins.size())
    {
        return {};
    }

    const auto& plugin = plugins[item.index];
    return plugin.name.empty() ? std::wstring_view(plugin.id) : std::wstring_view(plugin.name);
}

[[nodiscard]] std::wstring_view GetDescription(const PrefsPluginListItem& item) noexcept
{
    if (item.type == PrefsPluginType::FileSystem)
    {
        const auto& plugins = FileSystemPluginManager::GetInstance().GetPlugins();
        if (item.index >= plugins.size())
        {
            return {};
        }

        return std::wstring_view(plugins[item.index].description);
    }

    const auto& plugins = ViewerPluginManager::GetInstance().GetPlugins();
    if (item.index >= plugins.size())
    {
        return {};
    }

    return std::wstring_view(plugins[item.index].description);
}

[[nodiscard]] std::wstring_view GetShortIdOrId(const PrefsPluginListItem& item) noexcept
{
    if (item.type == PrefsPluginType::FileSystem)
    {
        const auto& plugins = FileSystemPluginManager::GetInstance().GetPlugins();
        if (item.index >= plugins.size())
        {
            return {};
        }

        const auto& plugin = plugins[item.index];
        return plugin.shortId.empty() ? std::wstring_view(plugin.id) : std::wstring_view(plugin.shortId);
    }

    const auto& plugins = ViewerPluginManager::GetInstance().GetPlugins();
    if (item.index >= plugins.size())
    {
        return {};
    }

    const auto& plugin = plugins[item.index];
    return plugin.shortId.empty() ? std::wstring_view(plugin.id) : std::wstring_view(plugin.shortId);
}

[[nodiscard]] bool IsLoadable(const PrefsPluginListItem& item) noexcept
{
    if (item.type == PrefsPluginType::FileSystem)
    {
        const auto& plugins = FileSystemPluginManager::GetInstance().GetPlugins();
        return (item.index < plugins.size()) ? plugins[item.index].loadable : false;
    }

    const auto& plugins = ViewerPluginManager::GetInstance().GetPlugins();
    return (item.index < plugins.size()) ? plugins[item.index].loadable : false;
}

[[nodiscard]] int GetOriginOrder(const PrefsPluginListItem& item) noexcept
{
    if (item.type == PrefsPluginType::FileSystem)
    {
        const auto& plugins = FileSystemPluginManager::GetInstance().GetPlugins();
        return (item.index < plugins.size()) ? static_cast<int>(plugins[item.index].origin) : 0;
    }

    const auto& plugins = ViewerPluginManager::GetInstance().GetPlugins();
    return (item.index < plugins.size()) ? static_cast<int>(plugins[item.index].origin) : 0;
}
} // namespace PrefsPlugins

namespace PrefsUi
{
PreferencesTypographyContext MakeTypographyContext(HWND hwnd) noexcept
{
    using RedSalamander::DxUi::FontRole;
    using RedSalamander::DxUi::Typography::GetDxUiTypographySpec;

    const UINT dpi = hwnd ? GetDpiForWindow(hwnd) : USER_DEFAULT_SCREEN_DPI;
    return PreferencesTypographyContext{
        .dpi     = (std::max<UINT>)(dpi, USER_DEFAULT_SCREEN_DPI),
        .body    = GetDxUiTypographySpec(FontRole::Body),
        .caption = GetDxUiTypographySpec(FontRole::Small),
        .title   = GetDxUiTypographySpec(FontRole::Title),
        .strong  = GetDxUiTypographySpec(FontRole::BodyStrong),
    };
}

int MeasureSingleLineTextWidthPx(const PreferencesTypographyContext& typography,
                                 const RedSalamander::DxUi::Typography::TypographySpec& spec,
                                 std::wstring_view text) noexcept
{
    using namespace RedSalamander::DxUi::Typography;

    IDWriteFactory* factory = GetSharedMeasurementFactory();
    if (! factory)
    {
        return 0;
    }

    wil::com_ptr<IDWriteTextFormat> format;
    if (FAILED(CreateTextFormat(factory, spec, format.put())) || ! format)
    {
        return 0;
    }
    static_cast<void>(format->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP));
    return MeasureSingleLineTextMetrics(factory, format.get(), typography.dpi, text).widthPx;
}

int MeasureWrappedTextHeightPx(const PreferencesTypographyContext& typography,
                               const RedSalamander::DxUi::Typography::TypographySpec& spec,
                               int width,
                               std::wstring_view text) noexcept
{
    using namespace RedSalamander::DxUi::Typography;

    if (width <= 0 || text.empty() || text.size() > static_cast<size_t>((std::numeric_limits<UINT32>::max)()))
    {
        return 0;
    }

    IDWriteFactory* factory = GetSharedMeasurementFactory();
    if (! factory)
    {
        return 0;
    }

    wil::com_ptr<IDWriteTextFormat> format;
    if (FAILED(CreateTextFormat(factory, spec, format.put())) || ! format)
    {
        return 0;
    }
    static_cast<void>(format->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP));
    static_cast<void>(format->SetParagraphAlignment(DWRITE_PARAGRAPH_ALIGNMENT_NEAR));
    static_cast<void>(format->SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING));

    const float layoutWidthDip = (std::max)(1.0f, (static_cast<float>(width) * 96.0f) / static_cast<float>(typography.dpi));
    wil::com_ptr<IDWriteTextLayout> layout;
    if (FAILED(factory->CreateTextLayout(text.data(), static_cast<UINT32>(text.size()), format.get(), layoutWidthDip, 4096.0f, layout.put())) || ! layout)
    {
        return 0;
    }

    DWRITE_TEXT_METRICS metrics{};
    if (FAILED(layout->GetMetrics(&metrics)))
    {
        return 0;
    }

    const int paddingY = UiMetrics::ScaleDip(typography.dpi, 6);
    return DipExtentToPixels(metrics.height, typography.dpi) + std::max(1, paddingY);
}

std::wstring_view TrimWhitespace(std::wstring_view text) noexcept
{
    return StringUtils::TrimWhitespace(text);
}

bool ContainsCaseInsensitive(std::wstring_view haystack, std::wstring_view needle) noexcept
{
    if (needle.empty())
    {
        return true;
    }

    const auto it = std::search(haystack.begin(), haystack.end(), needle.begin(), needle.end(), [](wchar_t a, wchar_t b) noexcept {
        return std::towlower(static_cast<wint_t>(a)) == std::towlower(static_cast<wint_t>(b));
    });
    return it != haystack.end();
}

void InvalidateComboBox(HWND combo) noexcept
{
    if (! combo)
    {
        return;
    }

    InvalidateRect(combo, nullptr, TRUE);

    COMBOBOXINFO cbi{};
    cbi.cbSize = sizeof(cbi);
    if (GetComboBoxInfo(combo, &cbi) && cbi.hwndItem)
    {
        InvalidateRect(cbi.hwndItem, nullptr, TRUE);
    }
}

PreferencesDialogState* GetDialogState(HWND childOrDialog) noexcept
{
    if (! childOrDialog)
    {
        return nullptr;
    }

    HWND dlg = GetAncestor(childOrDialog, GA_ROOT);
    if (! dlg)
    {
        dlg = childOrDialog;
    }

    if (! dlg || IsWindow(dlg) == FALSE)
    {
        return nullptr;
    }

    return reinterpret_cast<PreferencesDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));
}

bool PostDeferredAction(HWND hwnd, PreferencesDeferredActionKind kind) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return false;
    }

    auto payload  = std::make_unique<PreferencesDeferredActionPayload>();
    payload->kind = kind;
    return PostMessagePayload(hwnd, WndMsg::kPreferencesDeferredPaneAction, 0, std::move(payload));
}

std::optional<uint32_t> TryParseUInt32(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return std::nullopt;
    }

    uint32_t value = 0;
    for (const wchar_t ch : text)
    {
        if (ch < L'0' || ch > L'9')
        {
            return std::nullopt;
        }

        const uint32_t digit = static_cast<uint32_t>(ch - L'0');
        if (value > (std::numeric_limits<uint32_t>::max() - digit) / 10u)
        {
            return std::nullopt;
        }

        value = (value * 10u) + digit;
    }

    return value;
}

std::optional<uint64_t> TryParseUInt64(std::wstring_view text) noexcept
{
    if (text.empty())
    {
        return std::nullopt;
    }

    uint64_t value = 0;
    for (const wchar_t ch : text)
    {
        if (ch < L'0' || ch > L'9')
        {
            return std::nullopt;
        }

        const uint64_t digit = static_cast<uint64_t>(ch - L'0');
        if (value > (std::numeric_limits<uint64_t>::max() - digit) / 10ull)
        {
            return std::nullopt;
        }

        value = (value * 10ull) + digit;
    }

    return value;
}

bool EqualsNoCase(std::wstring_view a, std::wstring_view b) noexcept
{
    if (a.size() != b.size())
    {
        return false;
    }

    for (size_t i = 0; i < a.size(); ++i)
    {
        const wchar_t ca = static_cast<wchar_t>(std::towlower(static_cast<wint_t>(a[i])));
        const wchar_t cb = static_cast<wchar_t>(std::towlower(static_cast<wint_t>(b[i])));
        if (ca != cb)
        {
            return false;
        }
    }

    return true;
}

} // namespace PrefsUi

namespace PrefsUi
{
bool IsActuallyVisibleChildWindow(HWND hwnd) noexcept
{
    if (! hwnd || IsWindowVisible(hwnd) == FALSE)
    {
        return false;
    }

    // DxUi text bridges stay WS_VISIBLE for IME routing, but an empty region keeps them off-screen.
    wil::unique_hrgn region(CreateRectRgn(0, 0, 0, 0));
    if (region)
    {
        const int rgnType = GetWindowRgn(hwnd, region.get());
        if (rgnType == NULLREGION)
        {
            return false;
        }
    }

    return true;
}

void TryPushCard(std::vector<RECT>& cards, const RECT& card) noexcept
{
    cards.push_back(card);
}

} // namespace PrefsUi

namespace PrefsDxHost
{
bool Attach(HWND hwnd, RedSalamander::DxUi::WindowHost& host) noexcept
{
    if (! hwnd)
    {
        return false;
    }

    static_cast<void>(RedSalamander::Win32Callback::SetPropNoThrow(hwnd, kPrefsDxHostProp, &host));
    const bool attached = InstallWndProcHook(hwnd, kPrefsDxHostOriginalWndProcProp, PrefsDxHostWndProc);
    if (! attached)
    {
        RemovePropW(hwnd, kPrefsDxHostProp);
        return false;
    }

    ApplyAncestorRedrawSuppression(hwnd);
    return true;
}

void ResetOwnedHostWindow(wil::unique_hwnd& hwnd) noexcept
{
    if (! hwnd)
    {
        return;
    }

    if (IsWindow(hwnd.get()) == FALSE)
    {
        static_cast<void>(hwnd.release());
        return;
    }

    if (auto* host = reinterpret_cast<RedSalamander::DxUi::WindowHost*>(GetPropW(hwnd.get(), kPrefsDxHostProp)))
    {
        host->Detach();
    }

    RemovePropW(hwnd.get(), kPrefsDxHostProp);
    RestoreWndProcHook(hwnd.get(), kPrefsDxHostOriginalWndProcProp);
    hwnd.reset();
}

#ifdef ENABLE_TESTS
size_t CountVisibleRenderedHosts(HWND parent) noexcept
{
    if (! parent)
    {
        return 0u;
    }

    size_t count = 0u;
    for (HWND child = GetWindow(parent, GW_CHILD); child; child = GetWindow(child, GW_HWNDNEXT))
    {
        if (! PrefsUi::IsActuallyVisibleChildWindow(child))
        {
            continue;
        }

        auto* host = reinterpret_cast<RedSalamander::DxUi::WindowHost*>(GetPropW(child, kPrefsDxHostProp));
        if (! host)
        {
            continue;
        }

        if (host->DebugGetRenderCount() != 0u)
        {
            ++count;
        }
    }

    return count;
}

size_t CountVisibleHostsWithResizeFailures(HWND parent) noexcept
{
    if (! parent)
    {
        return 0u;
    }

    size_t count = 0u;
    for (HWND child = GetWindow(parent, GW_CHILD); child; child = GetWindow(child, GW_HWNDNEXT))
    {
        if (! PrefsUi::IsActuallyVisibleChildWindow(child))
        {
            continue;
        }

        auto* host = reinterpret_cast<RedSalamander::DxUi::WindowHost*>(GetPropW(child, kPrefsDxHostProp));
        if (! host)
        {
            continue;
        }

        if (host->DebugGetResizeFailureCount() != 0u)
        {
            ++count;
        }
    }

    return count;
}

uint64_t SumVisibleRenderedHostRenderCounts(HWND parent) noexcept
{
    if (! parent)
    {
        return 0u;
    }

    struct VisibleRenderCountAccumulator
    {
        uint64_t total = 0u;
    } accumulator{};

    EnumChildWindows(parent,
                     [](HWND child, LPARAM lParam) noexcept -> BOOL
    {
        auto& accumulatorRef = *reinterpret_cast<VisibleRenderCountAccumulator*>(lParam);
        if (! PrefsUi::IsActuallyVisibleChildWindow(child))
        {
            return TRUE;
        }

        auto* host = reinterpret_cast<RedSalamander::DxUi::WindowHost*>(GetPropW(child, kPrefsDxHostProp));
        if (! host)
        {
            return TRUE;
        }

        accumulatorRef.total += host->DebugGetRenderCount();
        return TRUE;
    },
                     reinterpret_cast<LPARAM>(&accumulator));

    return accumulator.total;
}

bool TryGetDirectHostMetrics(HWND hwnd, size_t& visibleHostCount, size_t& resizeFailureCount, uint64_t& renderCountTotal) noexcept
{
    visibleHostCount   = 0u;
    resizeFailureCount = 0u;
    renderCountTotal   = 0u;

    if (! hwnd || IsWindow(hwnd) == FALSE || ! PrefsUi::IsActuallyVisibleChildWindow(hwnd))
    {
        return false;
    }

    auto* host = reinterpret_cast<RedSalamander::DxUi::WindowHost*>(GetPropW(hwnd, kPrefsDxHostProp));
    if (! host)
    {
        return false;
    }

    visibleHostCount   = 1u;
    resizeFailureCount = host->DebugGetResizeFailureCount() != 0u ? 1u : 0u;
    renderCountTotal   = host->DebugGetRenderCount();
    return true;
}
#endif
} // namespace PrefsDxHost

namespace PrefsFile
{
bool TryReadFileToString(const std::filesystem::path& path, std::string& out) noexcept
{
    out.clear();

    std::ifstream file(path, std::ios::binary);
    if (! file)
    {
        return false;
    }

    file.seekg(0, std::ios::end);
    const std::streamoff end = file.tellg();
    if (end < 0)
    {
        return false;
    }

    file.seekg(0, std::ios::beg);
    out.resize(static_cast<size_t>(end));
    if (end > 0)
    {
        file.read(out.data(), static_cast<std::streamsize>(end));
        if (! file)
        {
            return false;
        }
    }
    return true;
}

bool TryWriteFileFromString(const std::filesystem::path& path, std::string_view text) noexcept
{
    std::ofstream file(path, std::ios::binary | std::ios::trunc);
    if (! file)
    {
        return false;
    }

    if (! text.empty())
    {
        file.write(text.data(), static_cast<std::streamsize>(text.size()));
        if (! file)
        {
            return false;
        }
    }

    file.flush();
    return static_cast<bool>(file);
}
} // namespace PrefsFile

namespace PrefsFolders
{
FolderPanePreferences GetFolderPanePreferences(const Common::Settings::Settings& settings, std::wstring_view slot) noexcept
{
    FolderPanePreferences prefs{};
    if (! settings.folders.has_value())
    {
        return prefs;
    }

    for (const auto& pane : settings.folders->items)
    {
        if (pane.slot != slot)
        {
            continue;
        }

        prefs.display          = pane.view.display;
        prefs.sortBy           = pane.view.sortBy;
        prefs.sortDirection    = pane.view.sortDirection;
        prefs.statusBarVisible = pane.view.statusBarVisible;
        break;
    }

    return prefs;
}

uint32_t GetFolderHistoryMax(const Common::Settings::Settings& settings) noexcept
{
    if (! settings.folders.has_value())
    {
        return Common::Settings::FoldersSettings{}.historyMax;
    }

    return std::clamp(settings.folders->historyMax, 1u, 50u);
}

bool GetFolderShowHiddenFiles(const Common::Settings::Settings& settings) noexcept
{
    if (! settings.folders.has_value())
    {
        return Common::Settings::FoldersSettings{}.showHiddenFiles;
    }

    return settings.folders->showHiddenFiles;
}

bool GetFolderShowSystemFiles(const Common::Settings::Settings& settings) noexcept
{
    if (! settings.folders.has_value())
    {
        return Common::Settings::FoldersSettings{}.showSystemFiles;
    }

    return settings.folders->showSystemFiles;
}

bool AreEquivalentFolderPreferences(const Common::Settings::Settings& a, const Common::Settings::Settings& b) noexcept
{
    if (GetFolderShowHiddenFiles(a) != GetFolderShowHiddenFiles(b) || GetFolderShowSystemFiles(a) != GetFolderShowSystemFiles(b))
    {
        return false;
    }

    if (GetFolderHistoryMax(a) != GetFolderHistoryMax(b))
    {
        return false;
    }

    const FolderPanePreferences leftA  = GetFolderPanePreferences(a, kLeftPaneSlot);
    const FolderPanePreferences leftB  = GetFolderPanePreferences(b, kLeftPaneSlot);
    const FolderPanePreferences rightA = GetFolderPanePreferences(a, kRightPaneSlot);
    const FolderPanePreferences rightB = GetFolderPanePreferences(b, kRightPaneSlot);

    if (leftA.display != leftB.display || leftA.sortBy != leftB.sortBy || leftA.sortDirection != leftB.sortDirection ||
        leftA.statusBarVisible != leftB.statusBarVisible)
    {
        return false;
    }

    if (rightA.display != rightB.display || rightA.sortBy != rightB.sortBy || rightA.sortDirection != rightB.sortDirection ||
        rightA.statusBarVisible != rightB.statusBarVisible)
    {
        return false;
    }

    return true;
}

Common::Settings::FolderSortDirection DefaultFolderSortDirection(Common::Settings::FolderSortBy sortBy) noexcept
{
    switch (sortBy)
    {
        case Common::Settings::FolderSortBy::Time:
        case Common::Settings::FolderSortBy::Size: return Common::Settings::FolderSortDirection::Descending;
        case Common::Settings::FolderSortBy::Name:
        case Common::Settings::FolderSortBy::Extension:
        case Common::Settings::FolderSortBy::Attributes:
        case Common::Settings::FolderSortBy::None: return Common::Settings::FolderSortDirection::Ascending;
    }
    return Common::Settings::FolderSortDirection::Ascending;
}

Common::Settings::FoldersSettings* EnsureWorkingFoldersSettings(Common::Settings::Settings& settings) noexcept
{
    if (settings.folders.has_value())
    {
        return &settings.folders.value();
    }

    settings.folders.emplace();
    return &settings.folders.value();
}

Common::Settings::FolderPane* EnsureWorkingFolderPane(Common::Settings::Settings& settings, std::wstring_view slot) noexcept
{
    Common::Settings::FoldersSettings* folders = settings.folders.has_value() ? &settings.folders.value() : EnsureWorkingFoldersSettings(settings);
    if (! folders)
    {
        return nullptr;
    }

    for (auto& pane : folders->items)
    {
        if (pane.slot == slot)
        {
            return &pane;
        }
    }

    Common::Settings::FolderPane pane{};
    pane.slot.assign(slot);
    folders->items.push_back(std::move(pane));
    return &folders->items.back();
}
} // namespace PrefsFolders

namespace PrefsMonitor
{
const Common::Settings::MonitorSettings& GetMonitorSettingsOrDefault(const Common::Settings::Settings& settings) noexcept
{
    static const Common::Settings::MonitorSettings kDefaults{};
    if (settings.monitor.has_value())
    {
        return settings.monitor.value();
    }
    return kDefaults;
}

Common::Settings::MonitorSettings* EnsureWorkingMonitorSettings(Common::Settings::Settings& settings) noexcept
{
    if (settings.monitor.has_value())
    {
        return &settings.monitor.value();
    }

    settings.monitor.emplace();
    return &settings.monitor.value();
}
} // namespace PrefsMonitor

namespace PrefsCache
{
const Common::Settings::CacheSettings& GetCacheSettingsOrDefault(const Common::Settings::Settings& settings) noexcept
{
    static const Common::Settings::CacheSettings kDefaults{};
    if (settings.cache.has_value())
    {
        return settings.cache.value();
    }
    return kDefaults;
}

Common::Settings::CacheSettings* EnsureWorkingCacheSettings(Common::Settings::Settings& settings) noexcept
{
    if (settings.cache.has_value())
    {
        return &settings.cache.value();
    }

    settings.cache.emplace();
    return &settings.cache.value();
}

void MaybeResetWorkingCacheSettingsIfEmpty(Common::Settings::Settings& settings) noexcept
{
    if (! settings.cache.has_value())
    {
        return;
    }

    const auto& directoryInfo     = settings.cache->directoryInfo;
    const bool wroteDirectoryInfo = (directoryInfo.maxBytes.has_value() && directoryInfo.maxBytes.value() > 0) || directoryInfo.maxWatchers.has_value() ||
                                    directoryInfo.mruWatched.has_value();
    if (! wroteDirectoryInfo)
    {
        settings.cache.reset();
    }
}

std::optional<uint64_t> TryParseCacheBytes(std::wstring_view text) noexcept
{
    std::wstring_view trimmed = PrefsUi::TrimWhitespace(text);
    if (trimmed.empty())
    {
        return std::nullopt;
    }

    size_t digitCount = 0;
    while (digitCount < trimmed.size() && trimmed[digitCount] >= L'0' && trimmed[digitCount] <= L'9')
    {
        ++digitCount;
    }
    if (digitCount == 0)
    {
        return std::nullopt;
    }

    const auto valueOpt = PrefsUi::TryParseUInt64(trimmed.substr(0, digitCount));
    if (! valueOpt.has_value())
    {
        return std::nullopt;
    }

    const uint64_t value   = valueOpt.value();
    std::wstring_view unit = PrefsUi::TrimWhitespace(trimmed.substr(digitCount));
    uint64_t multiplier    = 1024ull;

    if (! unit.empty())
    {
        if (PrefsUi::EqualsNoCase(unit, L"kb") || PrefsUi::EqualsNoCase(unit, L"kib"))
        {
            multiplier = 1024ull;
        }
        else if (PrefsUi::EqualsNoCase(unit, L"mb") || PrefsUi::EqualsNoCase(unit, L"mib"))
        {
            multiplier = 1024ull * 1024ull;
        }
        else if (PrefsUi::EqualsNoCase(unit, L"gb") || PrefsUi::EqualsNoCase(unit, L"gib"))
        {
            multiplier = 1024ull * 1024ull * 1024ull;
        }
        else
        {
            return std::nullopt;
        }
    }

    if (value == 0 || multiplier == 0)
    {
        return 0ull;
    }

    if (value > std::numeric_limits<uint64_t>::max() / multiplier)
    {
        return std::nullopt;
    }

    return value * multiplier;
}

std::wstring FormatCacheBytes(uint64_t bytes) noexcept
{
    constexpr uint64_t kKiB = 1024ull;
    constexpr uint64_t kMiB = 1024ull * 1024ull;
    constexpr uint64_t kGiB = 1024ull * 1024ull * 1024ull;

    if (bytes == 0)
    {
        return {};
    }

    if (bytes % kGiB == 0)
    {
        std::wstring text = std::to_wstring(bytes / kGiB);
        text.append(L" GB");
        return text;
    }
    if (bytes % kMiB == 0)
    {
        std::wstring text = std::to_wstring(bytes / kMiB);
        text.append(L" MB");
        return text;
    }
    if (bytes % kKiB == 0)
    {
        std::wstring text = std::to_wstring(bytes / kKiB);
        text.append(L" KB");
        return text;
    }

    return std::to_wstring(bytes);
}
} // namespace PrefsCache

namespace PrefsConnections
{
const Common::Settings::ConnectionsSettings& GetConnectionsSettingsOrDefault(const Common::Settings::Settings& settings) noexcept
{
    static const Common::Settings::ConnectionsSettings kDefaults{};
    if (settings.connections.has_value())
    {
        return settings.connections.value();
    }
    return kDefaults;
}

Common::Settings::ConnectionsSettings* EnsureWorkingConnectionsSettings(Common::Settings::Settings& settings) noexcept
{
    if (settings.connections.has_value())
    {
        return &settings.connections.value();
    }

    settings.connections.emplace();
    return &settings.connections.value();
}

void MaybeResetWorkingConnectionsSettingsIfEmpty(Common::Settings::Settings& settings) noexcept
{
    if (! settings.connections.has_value())
    {
        return;
    }

    if (! settings.connections->items.empty())
    {
        return;
    }

    const Common::Settings::ConnectionsSettings defaults{};
    const bool wroteGlobals = settings.connections->bypassWindowsHello != defaults.bypassWindowsHello ||
                              settings.connections->allowInsecureTlsInAutomation != defaults.allowInsecureTlsInAutomation ||
                              settings.connections->windowsHelloReauthTimeoutMinute != defaults.windowsHelloReauthTimeoutMinute;
    if (! wroteGlobals)
    {
        settings.connections.reset();
    }
}
} // namespace PrefsConnections

namespace PrefsFileOperations
{
const Common::Settings::FileOperationsSettings& GetFileOperationsSettingsOrDefault(const Common::Settings::Settings& settings) noexcept
{
    static const Common::Settings::FileOperationsSettings kDefaults{};
    if (settings.fileOperations.has_value())
    {
        return settings.fileOperations.value();
    }
    return kDefaults;
}

Common::Settings::FileOperationsSettings* EnsureWorkingFileOperationsSettings(Common::Settings::Settings& settings) noexcept
{
    if (settings.fileOperations.has_value())
    {
        return &settings.fileOperations.value();
    }

    settings.fileOperations.emplace();
    return &settings.fileOperations.value();
}

void MaybeResetWorkingFileOperationsSettingsIfEmpty(Common::Settings::Settings& settings) noexcept
{
    if (! settings.fileOperations.has_value())
    {
        return;
    }

    const Common::Settings::FileOperationsSettings defaults{};
    const auto& fileOperations = settings.fileOperations.value();
    const bool hasNonDefault =
        fileOperations.autoDismissSuccess != defaults.autoDismissSuccess || fileOperations.preCalcEnabled != defaults.preCalcEnabled ||
        fileOperations.preCalcMaxWorkers != defaults.preCalcMaxWorkers || fileOperations.crossFsBridgeBufferSizeKB != defaults.crossFsBridgeBufferSizeKB ||
        fileOperations.defaultBandwidthLimitBytesPerSecond != defaults.defaultBandwidthLimitBytesPerSecond ||
        fileOperations.maxDiagnosticsLogFiles != defaults.maxDiagnosticsLogFiles || fileOperations.diagnosticsInfoEnabled != defaults.diagnosticsInfoEnabled ||
        fileOperations.diagnosticsDebugEnabled != defaults.diagnosticsDebugEnabled || fileOperations.maxIssueReportFiles.has_value() ||
        fileOperations.maxDiagnosticsInMemory.has_value() || fileOperations.maxDiagnosticsPerFlush.has_value() ||
        fileOperations.diagnosticsFlushIntervalMs.has_value() || fileOperations.diagnosticsCleanupIntervalMs.has_value();

    if (! hasNonDefault)
    {
        settings.fileOperations.reset();
    }
}
} // namespace PrefsFileOperations

namespace PrefsCompareDirectories
{
const Common::Settings::CompareDirectoriesSettings& GetCompareDirectoriesSettingsOrDefault(const Common::Settings::Settings& settings) noexcept
{
    static const Common::Settings::CompareDirectoriesSettings kDefaults{};
    if (settings.compareDirectories.has_value())
    {
        return settings.compareDirectories.value();
    }
    return kDefaults;
}

Common::Settings::CompareDirectoriesSettings* EnsureWorkingCompareDirectoriesSettings(Common::Settings::Settings& settings) noexcept
{
    if (settings.compareDirectories.has_value())
    {
        return &settings.compareDirectories.value();
    }

    settings.compareDirectories.emplace();
    return &settings.compareDirectories.value();
}

void MaybeResetWorkingCompareDirectoriesSettingsIfEmpty(Common::Settings::Settings& settings) noexcept
{
    if (! settings.compareDirectories.has_value())
    {
        return;
    }

    const Common::Settings::CompareDirectoriesSettings defaults{};
    const auto& compare = settings.compareDirectories.value();

    const bool hasNonDefault =
        compare.compareSize != defaults.compareSize || compare.compareDateTime != defaults.compareDateTime ||
        compare.compareAttributes != defaults.compareAttributes || compare.compareContent != defaults.compareContent ||
        compare.compareSubdirectories != defaults.compareSubdirectories || compare.compareSubdirectoryAttributes != defaults.compareSubdirectoryAttributes ||
        compare.selectSubdirsOnlyInOnePane != defaults.selectSubdirsOnlyInOnePane || compare.ignoreFiles != defaults.ignoreFiles ||
        compare.ignoreFilesPatterns != defaults.ignoreFilesPatterns || compare.ignoreDirectories != defaults.ignoreDirectories ||
        compare.ignoreDirectoriesPatterns != defaults.ignoreDirectoriesPatterns || compare.keepIdenticalItems != defaults.keepIdenticalItems ||
        compare.showIdenticalItems != defaults.showIdenticalItems || compare.contentCompareWorkerCount != defaults.contentCompareWorkerCount;

    if (! hasNonDefault)
    {
        settings.compareDirectories.reset();
    }
}
} // namespace PrefsCompareDirectories
