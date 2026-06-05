#include "DxUi.Internal.h"

#include <algorithm>
#include <chrono>
#include <cmath>
#include <cstring>
#include <cwctype>
#include <format>
#include <limits>
#include <map>
#include <mutex>
#include <new>
#include <ranges>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

#include <UIAutomation.h>
#include <d3d11_4.h>
#include <dcomp.h>
#include <dwmapi.h>
#include <dxgi1_6.h>
#include <richedit.h>
#include <windowsx.h>

#include "DxUi.FrameRuntime.h"
#include "DxUi.Typography.h"
#include "Helpers.h"
#include "Ui/AnimationDispatcher.h"
#include "WindowMessages.h"

#pragma comment(lib, "d2d1.lib")
#pragma comment(lib, "d3d11.lib")
#pragma comment(lib, "dcomp.lib")
#pragma comment(lib, "dwmapi.lib")
#pragma comment(lib, "dxgi.lib")
#pragma comment(lib, "dwrite.lib")

namespace RedSalamander::DxUi
{
namespace
{
constexpr UINT kModifierAlt                    = 0x0100u;
constexpr wchar_t kDxUiDiagnosticsProp[]       = L"RedSalamander.Preferences.DxDiagnostics";
constexpr uint64_t kTooltipFallbackShowDelayMs = 500u;
constexpr uint64_t kTooltipMinShowDelayMs      = 100u;
constexpr uint64_t kTooltipMaxShowDelayMs      = 2500u;

struct SharedWindowHostGraphicsResources
{
    wil::com_ptr<ID3D11Device> d3dDevice;
    wil::com_ptr<ID3D11DeviceContext> d3dContext;
    wil::com_ptr<IDXGIFactory2> dxgiFactory;
    wil::com_ptr<IDCompositionDesktopDevice> dcompDevice;
    wil::com_ptr<ID2D1Factory1> d2dFactory;
    wil::com_ptr<ID2D1Device> d2dDevice;
    wil::com_ptr<IDWriteFactory> dwriteFactory;
    DWORD ownerThreadId            = 0u;
    D3D_FEATURE_LEVEL featureLevel = D3D_FEATURE_LEVEL_11_0;
    uint64_t generation            = 1u;
    uint32_t attachedHostCount     = 0u;
};

struct WindowHostDirtyRectMetrics
{
    bool isPartialDirty = false;
    uint64_t count      = 0u;
    uint64_t areaPx     = 0u;
};

[[nodiscard]] WindowHostDirtyRectMetrics ResolveWindowHostDirtyRectMetrics(const RECT* dirtyRectPx, UINT widthPx, UINT heightPx) noexcept
{
    if (! dirtyRectPx || dirtyRectPx->left >= dirtyRectPx->right || dirtyRectPx->top >= dirtyRectPx->bottom)
    {
        return {};
    }

    const auto dirtyWidthPx  = static_cast<UINT>(dirtyRectPx->right - dirtyRectPx->left);
    const auto dirtyHeightPx = static_cast<UINT>(dirtyRectPx->bottom - dirtyRectPx->top);
    if (dirtyWidthPx >= widthPx && dirtyHeightPx >= heightPx)
    {
        return {};
    }

    return WindowHostDirtyRectMetrics{
        .isPartialDirty = true,
        .count          = 1u,
        .areaPx         = static_cast<uint64_t>(dirtyWidthPx) * static_cast<uint64_t>(dirtyHeightPx),
    };
}

void EmitWindowHostFrameMetrics(
    uint64_t totalUs, uint64_t updateUs, uint64_t renderUs, uint64_t presentUs, const WindowHostDirtyRectMetrics& dirtyMetrics) noexcept
{
    EmitFrameMetric(L"dxui.frame.total_us", totalUs);
    EmitFrameMetric(L"dxui.frame.update_us", updateUs);
    EmitFrameMetric(L"dxui.frame.render_us", renderUs);
    EmitFrameMetric(L"dxui.frame.present_us", presentUs);
    Debug::Perf::EmitValue(L"dxui.frame.dirty_rect_count", dirtyMetrics.count);
    Debug::Perf::EmitValue(L"dxui.frame.dirty_rect_area_px", dirtyMetrics.areaPx);
}

[[nodiscard]] wchar_t NormalizeMnemonicChar(wchar_t ch) noexcept
{
    return static_cast<wchar_t>(std::towupper(static_cast<wint_t>(ch)));
}

[[nodiscard]] bool IsPointerButtonDownMessage(UINT msg) noexcept
{
    return msg == WM_LBUTTONDOWN || msg == WM_RBUTTONDOWN;
}

[[nodiscard]] bool IsPointerDoubleClickMessage(UINT msg) noexcept
{
    return msg == WM_LBUTTONDBLCLK || msg == WM_RBUTTONDBLCLK;
}

[[nodiscard]] const wchar_t* TraceWindowHostMessageName(UINT msg) noexcept
{
    switch (msg)
    {
        case WM_MOUSEMOVE: return L"WM_MOUSEMOVE";
        case WM_LBUTTONDOWN: return L"WM_LBUTTONDOWN";
        case WM_LBUTTONUP: return L"WM_LBUTTONUP";
        case WM_LBUTTONDBLCLK: return L"WM_LBUTTONDBLCLK";
        case WM_RBUTTONDOWN: return L"WM_RBUTTONDOWN";
        case WM_RBUTTONUP: return L"WM_RBUTTONUP";
        case WM_RBUTTONDBLCLK: return L"WM_RBUTTONDBLCLK";
        case WM_MOUSEWHEEL: return L"WM_MOUSEWHEEL";
        case WM_MOUSEHWHEEL: return L"WM_MOUSEHWHEEL";
        case WM_MOUSEACTIVATE: return L"WM_MOUSEACTIVATE";
        case WM_SETCURSOR: return L"WM_SETCURSOR";
        case WM_CAPTURECHANGED: return L"WM_CAPTURECHANGED";
        case WM_CANCELMODE: return L"WM_CANCELMODE";
        case WM_SETFOCUS: return L"WM_SETFOCUS";
        case WM_KILLFOCUS: return L"WM_KILLFOCUS";
        case WM_ACTIVATE: return L"WM_ACTIVATE";
        case WM_ACTIVATEAPP: return L"WM_ACTIVATEAPP";
        default: return L"message";
    }
}

[[nodiscard]] bool ShouldTraceWindowHostMessage(UINT msg) noexcept
{
    switch (msg)
    {
        case WM_MOUSEMOVE:
        case WM_LBUTTONDOWN:
        case WM_LBUTTONUP:
        case WM_LBUTTONDBLCLK:
        case WM_RBUTTONDOWN:
        case WM_RBUTTONUP:
        case WM_RBUTTONDBLCLK:
        case WM_MOUSEWHEEL:
        case WM_MOUSEHWHEEL:
        case WM_MOUSEACTIVATE:
        case WM_SETCURSOR:
        case WM_CAPTURECHANGED:
        case WM_CANCELMODE:
        case WM_SETFOCUS:
        case WM_KILLFOCUS:
        case WM_ACTIVATE:
        case WM_ACTIVATEAPP: return true;
        default: return false;
    }
}

[[nodiscard]] bool ShouldResolveScreenHitWindowForWindowHostTrace(UINT msg) noexcept
{
    switch (msg)
    {
        case WM_MOUSEMOVE:
        case WM_LBUTTONDOWN:
        case WM_LBUTTONUP:
        case WM_LBUTTONDBLCLK:
        case WM_RBUTTONDOWN:
        case WM_RBUTTONUP:
        case WM_RBUTTONDBLCLK:
        case WM_MOUSEACTIVATE: return true;
        default: return false;
    }
}

[[nodiscard]] bool IsClientMouseMessageForWindowHostTrace(UINT msg) noexcept
{
    return msg >= WM_MOUSEFIRST && msg <= WM_MOUSELAST && msg != WM_MOUSEWHEEL && msg != WM_MOUSEHWHEEL;
}

[[nodiscard]] std::wstring TraceLimitedText(std::wstring_view value)
{
    constexpr size_t kMaxTraceTextChars = 80u;
    std::wstring text;
    text.reserve((std::min)(value.size(), kMaxTraceTextChars + 3u));
    for (const wchar_t ch : value)
    {
        if (text.size() >= kMaxTraceTextChars)
        {
            text.append(L"...");
            break;
        }
        text.push_back((ch == L'\r' || ch == L'\n' || ch == L'\t') ? L' ' : ch);
    }
    return text;
}

[[nodiscard]] std::wstring TraceRect(const D2D1_RECT_F& rect)
{
    return std::format(L"({:.1f},{:.1f},{:.1f},{:.1f})", rect.left, rect.top, rect.right, rect.bottom);
}

[[nodiscard]] std::wstring TraceWindowClassName(HWND hwnd)
{
    if (! hwnd)
    {
        return L"null";
    }

    wchar_t className[128]{};
    const int length = GetClassNameW(hwnd, className, static_cast<int>(std::size(className)));
    if (length <= 0)
    {
        return L"<unknown>";
    }
    return std::wstring(className, static_cast<size_t>(length));
}

[[nodiscard]] const wchar_t* TraceButtonVariantName(ButtonVariant variant) noexcept
{
    switch (variant)
    {
        case ButtonVariant::Standard: return L"Standard";
        case ButtonVariant::DropDown: return L"DropDown";
        case ButtonVariant::Split: return L"Split";
        case ButtonVariant::Hyperlink: return L"Hyperlink";
        case ButtonVariant::IconOnly: return L"IconOnly";
        case ButtonVariant::Repeat: return L"Repeat";
        default: return L"Unknown";
    }
}

[[nodiscard]] const wchar_t* TraceComboBoxVariantName(ComboBoxVariant variant) noexcept
{
    switch (variant)
    {
        case ComboBoxVariant::Window: return L"Window";
        case ComboBoxVariant::Modern: return L"Modern";
        case ComboBoxVariant::Edit: return L"Edit";
        default: return L"Unknown";
    }
}

[[nodiscard]] std::wstring DescribeWindowHostTraceControl(const Control* control)
{
    if (! control)
    {
        return L"null";
    }

    std::wstring type = L"Control";
    std::wstring detail;
    if (const auto* const comboBox = dynamic_cast<const ComboBox*>(control))
    {
        type   = L"ComboBox";
        detail = std::format(L" variant={} text=\"{}\" display=\"{}\" editable={}",
                             TraceComboBoxVariantName(comboBox->GetVariant()),
                             TraceLimitedText(comboBox->GetText()),
                             TraceLimitedText(comboBox->GetDisplayedText()),
                             comboBox->IsEditable() ? 1 : 0);
    }
    else if (const auto* const textField = dynamic_cast<const TextField*>(control))
    {
        type   = L"TextField";
        detail = std::format(L" text=\"{}\" multiline={}", TraceLimitedText(textField->GetText()), textField->IsMultiline() ? 1 : 0);
    }
    else if (const auto* const checkbox = dynamic_cast<const Checkbox*>(control))
    {
        type   = L"Checkbox";
        detail = std::format(L" text=\"{}\" checked={} indeterminate={}",
                             TraceLimitedText(checkbox->GetDisplayedText()),
                             checkbox->IsChecked() ? 1 : 0,
                             checkbox->IsIndeterminate() ? 1 : 0);
    }
    else if (const auto* const toggle = dynamic_cast<const Toggle*>(control))
    {
        type   = L"Toggle";
        detail = std::format(L" text=\"{}\" checked={}", TraceLimitedText(toggle->GetDisplayedText()), toggle->IsChecked() ? 1 : 0);
    }
    else if (const auto* const radioButton = dynamic_cast<const RadioButton*>(control))
    {
        type   = L"RadioButton";
        detail = std::format(L" text=\"{}\" checked={}", TraceLimitedText(radioButton->GetText()), radioButton->IsChecked() ? 1 : 0);
    }
    else if (const auto* const button = dynamic_cast<const Button*>(control))
    {
        type   = L"Button";
        detail = std::format(
            L" text=\"{}\" variant={} primary={}", TraceLimitedText(button->GetText()), TraceButtonVariantName(button->GetVariant()), button->IsPrimary() ? 1 : 0);
    }
    else if (const auto* const label = dynamic_cast<const Label*>(control))
    {
        type   = L"Label";
        detail = std::format(L" text=\"{}\"", TraceLimitedText(label->GetText()));
    }
    else if (const auto* const statusStrip = dynamic_cast<const StatusStrip*>(control))
    {
        type   = L"StatusStrip";
        detail = std::format(L" text=\"{}\" sections={}", TraceLimitedText(statusStrip->GetText()), statusStrip->GetSectionCount());
    }
    else if (dynamic_cast<const Grid*>(control))
    {
        type = L"Grid";
    }
    else if (dynamic_cast<const PopupLayer*>(control))
    {
        type = L"PopupLayer";
    }
    else if (dynamic_cast<const ScrollPanel*>(control))
    {
        type = L"ScrollPanel";
    }
    else if (dynamic_cast<const PageHost*>(control))
    {
        type = L"PageHost";
    }
    else if (dynamic_cast<const Panel*>(control))
    {
        type = L"Panel";
    }

    return std::format(L"{}@{:#x} bounds={} hit={} visible={} enabled={} focus={} hover={} acc=\"{}\"{}",
                       type,
                       reinterpret_cast<uintptr_t>(control),
                       TraceRect(control->GetBounds()),
                       TraceRect(control->GetHitBounds()),
                       control->IsVisible() ? 1 : 0,
                       control->IsEnabled() ? 1 : 0,
                       control->HasFocus() ? 1 : 0,
                       control->IsHovered() ? 1 : 0,
                       TraceLimitedText(control->GetAccessibleName()),
                       detail);
}

template <typename... Args>
void TraceWindowHostDiagnostics(std::wstring_view eventName, std::wformat_string<Args...> format, Args&&... args) noexcept
{
    if (! IsContextMenuDiagnosticsEnabled())
    {
        return;
    }

    try
    {
        TraceContextMenuDiagnostics(eventName, std::format(format, std::forward<Args>(args)...));
    }
    catch (const std::bad_alloc&)
    {
        std::terminate();
    }
    catch (const std::format_error&)
    {
        // DxUi tracing is diagnostic only; formatting failure must not disturb input dispatch.
        TraceContextMenuDiagnostics(eventName, L"formatting failed");
    }
}

[[nodiscard]] UINT PointerButtonDownMessageFor(UINT msg) noexcept
{
    switch (msg)
    {
        case WM_LBUTTONDOWN:
        case WM_LBUTTONDBLCLK: return WM_LBUTTONDOWN;
        case WM_RBUTTONDOWN:
        case WM_RBUTTONDBLCLK: return WM_RBUTTONDOWN;
        default: return 0u;
    }
}

[[nodiscard]] uint64_t ResolveTooltipShowDelayMs() noexcept
{
    UINT hoverTimeMs = 0u;
    if (SystemParametersInfoW(SPI_GETMOUSEHOVERTIME, 0u, &hoverTimeMs, 0u) != FALSE && hoverTimeMs > 0u)
    {
        return std::clamp<uint64_t>(hoverTimeMs, kTooltipMinShowDelayMs, kTooltipMaxShowDelayMs);
    }

    return kTooltipFallbackShowDelayMs;
}

[[nodiscard]] bool IsWithinSystemDoubleClickBounds(POINT firstPointPx, POINT secondPointPx) noexcept
{
    const int halfWidthPx  = std::max(1, GetSystemMetrics(SM_CXDOUBLECLK) / 2);
    const int halfHeightPx = std::max(1, GetSystemMetrics(SM_CYDOUBLECLK) / 2);
    return std::abs(secondPointPx.x - firstPointPx.x) <= halfWidthPx && std::abs(secondPointPx.y - firstPointPx.y) <= halfHeightPx;
}

[[nodiscard]] std::map<DWORD, SharedWindowHostGraphicsResources>& GetSharedWindowHostGraphicsResourcesByThread() noexcept
{
    static auto* const resourcesByThread = []() noexcept
    {
        auto* resources = new (std::nothrow) std::map<DWORD, SharedWindowHostGraphicsResources>();
        if (resources == nullptr)
        {
            std::terminate();
        }
        return resources;
    }();
    return *resourcesByThread;
}

[[nodiscard]] bool IsInteractionDiagnosticsEnabled(HWND hwnd) noexcept
{
    return hwnd && GetPropW(hwnd, kDxUiDiagnosticsProp) != nullptr;
}

[[nodiscard]] std::mutex& GetSharedWindowHostGraphicsResourcesMutex() noexcept
{
    static std::mutex mutex;
    return mutex;
}

[[nodiscard]] std::mutex& GetAttachedWindowHostsMutex() noexcept
{
    static std::mutex mutex;
    return mutex;
}

[[nodiscard]] std::vector<WindowHost*>& GetAttachedWindowHosts() noexcept
{
    static std::vector<WindowHost*> hosts;
    return hosts;
}

[[nodiscard]] bool IsAttachedWindowHostRegistered(WindowHost* host) noexcept
{
    if (! host)
    {
        return false;
    }

    const std::scoped_lock lock(GetAttachedWindowHostsMutex());
    const auto& hosts = GetAttachedWindowHosts();
    return std::ranges::find(hosts, host) != hosts.end();
}

[[nodiscard]] SharedWindowHostGraphicsResources& GetSharedWindowHostGraphicsResourcesLocked(const DWORD threadId) noexcept
{
    auto& resourcesByThread   = GetSharedWindowHostGraphicsResourcesByThread();
    const auto [it, inserted] = resourcesByThread.try_emplace(threadId);
    if (inserted)
    {
        it->second.ownerThreadId = threadId;
    }
    return it->second;
}

[[nodiscard]] SharedWindowHostGraphicsResources& GetSharedWindowHostGraphicsResources() noexcept
{
    const std::scoped_lock lock(GetSharedWindowHostGraphicsResourcesMutex());
    return GetSharedWindowHostGraphicsResourcesLocked(GetCurrentThreadId());
}

void ResetSharedWindowHostGraphicsResourcesLocked(SharedWindowHostGraphicsResources& resources) noexcept
{
    resources.dwriteFactory.reset();
    resources.d2dDevice.reset();
    resources.d2dFactory.reset();
    resources.dcompDevice.reset();
    resources.dxgiFactory.reset();
    resources.d3dContext.reset();
    resources.d3dDevice.reset();
    ++resources.generation;
}

void ResetSharedWindowHostGraphicsResources() noexcept
{
    const std::scoped_lock lock(GetSharedWindowHostGraphicsResourcesMutex());
    auto& resourcesByThread = GetSharedWindowHostGraphicsResourcesByThread();
    if (const auto it = resourcesByThread.find(GetCurrentThreadId()); it != resourcesByThread.end())
    {
        ResetSharedWindowHostGraphicsResourcesLocked(it->second);
    }
}

void ResetAllSharedWindowHostGraphicsResourcesForProcessExit() noexcept
{
    const std::scoped_lock lock(GetSharedWindowHostGraphicsResourcesMutex());
    auto& resourcesByThread = GetSharedWindowHostGraphicsResourcesByThread();
    for (auto& [threadId, resources] : resourcesByThread)
    {
        static_cast<void>(threadId);
        ResetSharedWindowHostGraphicsResourcesLocked(resources);
    }
    resourcesByThread.clear();
}

void RegisterSharedWindowHostAttachment(WindowHost* host, DWORD ownerThreadId) noexcept
{
    {
        const std::scoped_lock hostLock(GetAttachedWindowHostsMutex());
        auto& hosts = GetAttachedWindowHosts();
        if (std::ranges::find(hosts, host) == hosts.end())
        {
            hosts.push_back(host);
        }
    }

    const std::scoped_lock lock(GetSharedWindowHostGraphicsResourcesMutex());
    SharedWindowHostGraphicsResources& resources = GetSharedWindowHostGraphicsResourcesLocked(ownerThreadId);
    ++resources.attachedHostCount;
}

void ReleaseSharedWindowHostAttachment(WindowHost* host, DWORD ownerThreadId) noexcept
{
    {
        const std::scoped_lock hostLock(GetAttachedWindowHostsMutex());
        auto& hosts = GetAttachedWindowHosts();
        if (const auto it = std::ranges::find(hosts, host); it != hosts.end())
        {
            hosts.erase(it);
        }
    }

    const DWORD currentThreadId = GetCurrentThreadId();
    if (ownerThreadId == 0u)
    {
        ownerThreadId = currentThreadId;
    }
    else if (ownerThreadId != currentThreadId)
    {
        Debug::Warning(L"DxUi::WindowHost: cross-thread detach releasing thread-local graphics bucket owner={} current={}", ownerThreadId, currentThreadId);
    }

    const std::scoped_lock lock(GetSharedWindowHostGraphicsResourcesMutex());
    auto& resourcesByThread = GetSharedWindowHostGraphicsResourcesByThread();
    const auto it           = resourcesByThread.find(ownerThreadId);
    if (it == resourcesByThread.end() || it->second.attachedHostCount == 0u)
    {
        Debug::Warning(L"DxUi::WindowHost: missing shared graphics bucket during detach owner={} current={}", ownerThreadId, currentThreadId);
        return;
    }

    SharedWindowHostGraphicsResources& resources = it->second;
    --resources.attachedHostCount;
    if (resources.attachedHostCount == 0u)
    {
        resourcesByThread.erase(it);
    }
}

[[nodiscard]] bool EnsureSharedWindowHostGraphicsResources() noexcept
{
    const std::scoped_lock lock(GetSharedWindowHostGraphicsResourcesMutex());
    SharedWindowHostGraphicsResources& resources = GetSharedWindowHostGraphicsResourcesLocked(GetCurrentThreadId());

    if (resources.d3dDevice && resources.d3dContext && resources.dxgiFactory && resources.d2dFactory && resources.d2dDevice && resources.dwriteFactory)
    {
        return true;
    }

    if (resources.d3dDevice || resources.d3dContext || resources.dxgiFactory || resources.d2dFactory || resources.d2dDevice || resources.dwriteFactory)
    {
        ResetSharedWindowHostGraphicsResourcesLocked(resources);
    }

    UINT creationFlags = D3D11_CREATE_DEVICE_BGRA_SUPPORT;
#if defined(_DEBUG)
    creationFlags |= D3D11_CREATE_DEVICE_DEBUG;
#endif
    D3D_FEATURE_LEVEL levels[] = {D3D_FEATURE_LEVEL_11_1, D3D_FEATURE_LEVEL_11_0, D3D_FEATURE_LEVEL_10_1};
    HRESULT hr                 = D3D11CreateDevice(nullptr,
                                                   D3D_DRIVER_TYPE_HARDWARE,
                                                   nullptr,
                                                   creationFlags,
                                                   levels,
                                                   std::size(levels),
                                                   D3D11_SDK_VERSION,
                                                   resources.d3dDevice.addressof(),
                                                   &resources.featureLevel,
                                                   resources.d3dContext.addressof());
    if (FAILED(hr))
    {
        Debug::Warning(L"DxUi::WindowHost: shared hardware D3D11CreateDevice failed, falling back to WARP: 0x{:08X}", hr);
        hr = D3D11CreateDevice(nullptr,
                               D3D_DRIVER_TYPE_WARP,
                               nullptr,
                               creationFlags,
                               levels,
                               std::size(levels),
                               D3D11_SDK_VERSION,
                               resources.d3dDevice.addressof(),
                               &resources.featureLevel,
                               resources.d3dContext.addressof());
    }
    if (FAILED(hr) || ! resources.d3dDevice || ! resources.d3dContext)
    {
        Debug::Error(L"DxUi::WindowHost: shared D3D11CreateDevice failed: 0x{:08X}", hr);
        ResetSharedWindowHostGraphicsResourcesLocked(resources);
        return false;
    }

    wil::com_ptr<IDXGIDevice> dxgiDevice;
    wil::com_ptr<IDXGIAdapter> adapter;
    const HRESULT hrDxgi        = resources.d3dDevice->QueryInterface(IID_PPV_ARGS(dxgiDevice.addressof()));
    const HRESULT hrAdapter     = SUCCEEDED(hrDxgi) && dxgiDevice ? dxgiDevice->GetAdapter(adapter.addressof()) : E_FAIL;
    const HRESULT hrDxgiFactory = SUCCEEDED(hrAdapter) && adapter ? adapter->GetParent(IID_PPV_ARGS(resources.dxgiFactory.addressof())) : E_FAIL;
    if (FAILED(hrDxgi) || ! dxgiDevice || FAILED(hrAdapter) || ! adapter || FAILED(hrDxgiFactory) || ! resources.dxgiFactory)
    {
        Debug::Error(L"DxUi::WindowHost: shared DXGI factory chain failed ({:08X}, {:08X}, {:08X})", hrDxgi, hrAdapter, hrDxgiFactory);
        ResetSharedWindowHostGraphicsResourcesLocked(resources);
        return false;
    }

    D2D1_FACTORY_OPTIONS options{};
#if defined(_DEBUG)
    // The D2D debug layer calls DebugBreak from process teardown when graphics
    // plugins are intentionally retained until ExitProcess to avoid driver
    // unload hangs. Keep D2D diagnostics opt-in; D3D debug remains enabled.
    options.debugLevel = D2D1_DEBUG_LEVEL_NONE;
#endif
    const HRESULT hrD2dFactory =
        D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, __uuidof(ID2D1Factory1), &options, reinterpret_cast<void**>(resources.d2dFactory.addressof()));
    if (FAILED(hrD2dFactory) || ! resources.d2dFactory)
    {
        Debug::Error(L"DxUi::WindowHost: shared D2D1CreateFactory failed: 0x{:08X}", hrD2dFactory);
        ResetSharedWindowHostGraphicsResourcesLocked(resources);
        return false;
    }

    const HRESULT hrDevice = resources.d2dFactory->CreateDevice(dxgiDevice.get(), resources.d2dDevice.addressof());
    if (FAILED(hrDevice) || ! resources.d2dDevice)
    {
        Debug::Error(L"DxUi::WindowHost: shared ID2D1Factory1::CreateDevice failed: 0x{:08X}", hrDevice);
        ResetSharedWindowHostGraphicsResourcesLocked(resources);
        return false;
    }

    const HRESULT hrWrite =
        DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory), reinterpret_cast<IUnknown**>(resources.dwriteFactory.addressof()));
    if (FAILED(hrWrite) || ! resources.dwriteFactory)
    {
        Debug::Error(L"DxUi::WindowHost: shared DWriteCreateFactory failed: 0x{:08X}", hrWrite);
        ResetSharedWindowHostGraphicsResourcesLocked(resources);
        return false;
    }

    return true;
}

[[nodiscard]] bool EnsureSharedWindowHostCompositionDevice() noexcept
{
    const std::scoped_lock lock(GetSharedWindowHostGraphicsResourcesMutex());
    SharedWindowHostGraphicsResources& resources = GetSharedWindowHostGraphicsResourcesLocked(GetCurrentThreadId());
    if (resources.dcompDevice)
    {
        return true;
    }
    if (! resources.d3dDevice)
    {
        return false;
    }

    wil::com_ptr<IDXGIDevice> dxgiDevice;
    const HRESULT hrDxgiDevice = resources.d3dDevice->QueryInterface(IID_PPV_ARGS(dxgiDevice.addressof()));
    if (FAILED(hrDxgiDevice) || ! dxgiDevice)
    {
        Debug::Error(L"DxUi::WindowHost: shared QueryInterface(IDXGIDevice) for DirectComposition failed: 0x{:08X}", hrDxgiDevice);
        return false;
    }

    const HRESULT hrDComp =
        DCompositionCreateDevice2(dxgiDevice.get(), __uuidof(IDCompositionDesktopDevice), reinterpret_cast<void**>(resources.dcompDevice.addressof()));
    if (FAILED(hrDComp) || ! resources.dcompDevice)
    {
        Debug::Error(L"DxUi::WindowHost: DCompositionCreateDevice2 failed: 0x{:08X}", hrDComp);
        resources.dcompDevice.reset();
        return false;
    }

    return true;
}

[[nodiscard]] bool IsHostWindowEffectivelyVisible(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE || IsWindowVisible(hwnd) == FALSE)
    {
        return false;
    }

    const HWND root = GetAncestor(hwnd, GA_ROOT);
    return ! root || IsIconic(root) == FALSE;
}

[[nodiscard]] bool IsSwapChainActiveHostWindow(HWND hwnd, UINT widthPx, UINT heightPx) noexcept
{
    return widthPx != 0u && heightPx != 0u && IsHostWindowEffectivelyVisible(hwnd);
}

[[nodiscard]] uint64_t PackTextFormatKey(FontRole role,
                                         DWRITE_TEXT_ALIGNMENT alignment,
                                         DWRITE_PARAGRAPH_ALIGNMENT paragraphAlignment,
                                         bool wrap,
                                         DWRITE_READING_DIRECTION readingDirection) noexcept
{
    return (static_cast<uint64_t>(static_cast<uint8_t>(role)) << 32u) | (static_cast<uint64_t>(alignment) << 24u) |
           (static_cast<uint64_t>(paragraphAlignment) << 16u) | (static_cast<uint64_t>(wrap ? 1u : 0u) << 8u) | static_cast<uint64_t>(readingDirection);
}

[[nodiscard]] bool HasFontFamily(IDWriteFactory* factory, PCWSTR familyName) noexcept
{
    return Typography::IsFontFamilyAvailable(factory, familyName);
}

[[nodiscard]] bool ModifiersContainCtrl(UINT modifiers) noexcept
{
    return (modifiers & MK_CONTROL) != 0u;
}

[[nodiscard]] bool ModifiersContainShift(UINT modifiers) noexcept
{
    return (modifiers & MK_SHIFT) != 0u;
}

[[nodiscard]] bool ModifiersContainAlt(UINT modifiers) noexcept
{
    return (modifiers & kModifierAlt) != 0u;
}

[[nodiscard]] bool NativeImeCompositionOwnsKey(UINT virtualKey, UINT modifiers) noexcept
{
    if (ModifiersContainCtrl(modifiers) || ModifiersContainShift(modifiers) || ModifiersContainAlt(modifiers))
    {
        return false;
    }

    switch (virtualKey)
    {
        case VK_RETURN:
        case VK_ESCAPE:
        case VK_TAB:
        case VK_UP:
        case VK_DOWN:
        case VK_LEFT:
        case VK_RIGHT:
        case VK_HOME:
        case VK_END:
        case VK_PRIOR:
        case VK_NEXT:
        case VK_BACK:
        case VK_DELETE: return true;
        default: return false;
    }
}

[[nodiscard]] UINT ComputeModifierMask() noexcept
{
    UINT mask = 0u;
    if ((GetKeyState(VK_SHIFT) & 0x8000) != 0)
    {
        mask |= MK_SHIFT;
    }
    if ((GetKeyState(VK_CONTROL) & 0x8000) != 0)
    {
        mask |= MK_CONTROL;
    }
    if ((GetKeyState(VK_MENU) & 0x8000) != 0)
    {
        mask |= kModifierAlt;
    }
    if ((GetKeyState(VK_LBUTTON) & 0x8000) != 0)
    {
        mask |= MK_LBUTTON;
    }
    if ((GetKeyState(VK_RBUTTON) & 0x8000) != 0)
    {
        mask |= MK_RBUTTON;
    }
    return mask;
}

[[nodiscard]] bool IsWindowOrDescendant(HWND root, HWND candidate) noexcept
{
    return root && candidate && (root == candidate || IsChild(root, candidate) != FALSE);
}

[[nodiscard]] bool OpenClipboardWithRetries(HWND ownerWindow) noexcept
{
    if (! ownerWindow)
    {
        return false;
    }

    constexpr int kClipboardOpenRetryCount = 20;
    constexpr DWORD kClipboardRetryDelayMs = 10;

    for (int attempt = 0; attempt < kClipboardOpenRetryCount; ++attempt)
    {
        if (OpenClipboard(ownerWindow) != 0)
        {
            return true;
        }

        if ((attempt + 1) < kClipboardOpenRetryCount)
        {
            if (GetOpenClipboardWindow() == nullptr)
            {
                static_cast<void>(CloseClipboard());
            }
            Sleep(kClipboardRetryDelayMs);
        }
    }

    return false;
}

#if defined(ENABLE_TESTS)
[[nodiscard]] std::mutex& GetClipboardFallbackMutex() noexcept
{
    static std::mutex mutex;
    return mutex;
}

[[nodiscard]] std::optional<std::wstring>& GetClipboardFallbackTextStorage() noexcept
{
    static std::optional<std::wstring> text;
    return text;
}
#endif

[[nodiscard]] bool SetClipboardUnicodeText(HWND ownerWindow, std::wstring_view text) noexcept
{
    if (! ownerWindow)
    {
        return false;
    }

    if (text.size() >= (std::numeric_limits<size_t>::max() / sizeof(wchar_t)))
    {
        return false;
    }

    const SIZE_T bytes = (text.size() + 1u) * sizeof(wchar_t);
    wil::unique_hglobal memory(GlobalAlloc(GMEM_MOVEABLE, bytes));
    if (! memory)
    {
        return false;
    }

    auto* out = static_cast<wchar_t*>(GlobalLock(memory.get()));
    if (! out)
    {
        return false;
    }

    if (! text.empty())
    {
        std::memcpy(out, text.data(), text.size() * sizeof(wchar_t));
    }
    out[text.size()] = L'\0';
    GlobalUnlock(memory.get());

#if defined(ENABLE_TESTS)
    const auto setFallbackText = [&]() noexcept
    {
        static_cast<void>(DebugSetClipboardFallbackText(text));
        return true;
    };
#endif

    if (! OpenClipboardWithRetries(ownerWindow))
    {
#if defined(ENABLE_TESTS)
        return setFallbackText();
#else
        return false;
#endif
    }
    const auto closeClipboard = wil::scope_exit([&] { CloseClipboard(); });

    if (EmptyClipboard() == 0)
    {
#if defined(ENABLE_TESTS)
        return setFallbackText();
#else
        return false;
#endif
    }
    if (SetClipboardData(CF_UNICODETEXT, memory.get()) == nullptr)
    {
#if defined(ENABLE_TESTS)
        return setFallbackText();
#else
        return false;
#endif
    }

    static_cast<void>(memory.release());
#if defined(ENABLE_TESTS)
    static_cast<void>(DebugSetClipboardFallbackText(text));
#endif
    return true;
}

[[nodiscard]] std::optional<std::wstring> TryReadClipboardUnicodeText(HWND ownerWindow) noexcept
{
    if (! OpenClipboardWithRetries(ownerWindow))
    {
#if defined(ENABLE_TESTS)
        return DebugReadClipboardFallbackText();
#else
        return std::nullopt;
#endif
    }
    const auto closeClipboard = wil::scope_exit([&] { CloseClipboard(); });

    HANDLE handle = GetClipboardData(CF_UNICODETEXT);
    if (! handle)
    {
#if defined(ENABLE_TESTS)
        return DebugReadClipboardFallbackText();
#else
        return std::nullopt;
#endif
    }

    const auto* text = static_cast<const wchar_t*>(GlobalLock(handle));
    if (! text)
    {
#if defined(ENABLE_TESTS)
        return DebugReadClipboardFallbackText();
#else
        return std::nullopt;
#endif
    }
    const auto unlock = wil::scope_exit([&] { GlobalUnlock(handle); });
    std::wstring clipboardText(text);
#if defined(ENABLE_TESTS)
    static_cast<void>(DebugSetClipboardFallbackText(clipboardText));
#endif
    return clipboardText;
}

void CollectFocusableControls(Control* control, std::vector<Control*>& out)
{
    if (! control || ! control->IsVisible() || ! control->IsEnabled())
    {
        return;
    }

    if (control->IsFocusable())
    {
        out.push_back(control);
    }

    if (auto* panel = dynamic_cast<Panel*>(control))
    {
        for (const auto& child : panel->GetChildren())
        {
            CollectFocusableControls(child.get(), out);
        }
    }
}

struct FocusAdvanceResult
{
    Control* control    = nullptr;
    bool wrapped        = false;
    size_t currentIndex = std::numeric_limits<size_t>::max();
    size_t nextIndex    = std::numeric_limits<size_t>::max();
};

[[nodiscard]] FocusAdvanceResult FindAdjacentFocusable(Control* root, Control* current, bool reverse)
{
    std::vector<Control*> focusableControls;
    CollectFocusableControls(root, focusableControls);
    if (focusableControls.empty())
    {
        return {};
    }

    if (! current)
    {
        return FocusAdvanceResult{.control   = reverse ? focusableControls.back() : focusableControls.front(),
                                  .nextIndex = reverse ? focusableControls.size() - 1u : 0u};
    }

    const auto it = std::ranges::find(focusableControls, current);
    if (it == focusableControls.end())
    {
        return FocusAdvanceResult{.control   = reverse ? focusableControls.back() : focusableControls.front(),
                                  .nextIndex = reverse ? focusableControls.size() - 1u : 0u};
    }

    const size_t currentIndex = static_cast<size_t>(std::distance(focusableControls.begin(), it));
    if (reverse)
    {
        const bool wrapped = it == focusableControls.begin();
        return FocusAdvanceResult{.control      = wrapped ? focusableControls.back() : *(it - 1),
                                  .wrapped      = wrapped,
                                  .currentIndex = currentIndex,
                                  .nextIndex    = wrapped ? focusableControls.size() - 1u : currentIndex - 1u};
    }

    const auto nextIt  = it + 1;
    const bool wrapped = nextIt == focusableControls.end();
    return FocusAdvanceResult{.control      = wrapped ? focusableControls.front() : *nextIt,
                              .wrapped      = wrapped,
                              .currentIndex = currentIndex,
                              .nextIndex    = wrapped ? 0u : currentIndex + 1u};
}

[[nodiscard]] Control* FindMnemonicControl(Control* control, wchar_t mnemonic) noexcept
{
    if (! control || ! control->IsVisible() || ! control->IsEnabled())
    {
        return nullptr;
    }

    if (NormalizeMnemonicChar(control->GetMnemonic()) == mnemonic)
    {
        return control;
    }

    for (size_t childIndex = 0u; childIndex < control->GetLogicalChildCount(); ++childIndex)
    {
        if (Control* child = control->GetLogicalChild(childIndex))
        {
            if (Control* match = FindMnemonicControl(child, mnemonic))
            {
                return match;
            }
        }
    }

    return nullptr;
}

[[nodiscard]] bool ControlBelongsToTree(const Control* root, const Control* target) noexcept
{
    if (! root || ! target)
    {
        return false;
    }
    if (root == target)
    {
        return true;
    }
    for (size_t childIndex = 0u; childIndex < root->GetLogicalChildCount(); ++childIndex)
    {
        if (const Control* const child = root->GetLogicalChild(childIndex))
        {
            if (ControlBelongsToTree(child, target))
            {
                return true;
            }
        }
    }
    return false;
}

struct ControlInteractionState
{
    bool inTree                 = false;
    bool effectivelyInteractive = false;
};

[[nodiscard]] ControlInteractionState ResolveControlInteractionState(const Control* root,
                                                                     const Control* target,
                                                                     bool ancestorsVisible = true,
                                                                     bool ancestorsEnabled = true) noexcept
{
    if (! root || ! target)
    {
        return {};
    }

    const bool visible = ancestorsVisible && root->IsVisible();
    const bool enabled = ancestorsEnabled && root->IsEnabled();
    if (root == target)
    {
        return ControlInteractionState{.inTree = true, .effectivelyInteractive = visible && enabled};
    }

    for (size_t childIndex = 0u; childIndex < root->GetLogicalChildCount(); ++childIndex)
    {
        if (const Control* const child = root->GetLogicalChild(childIndex))
        {
            const ControlInteractionState childState = ResolveControlInteractionState(child, target, visible, enabled);
            if (childState.inTree)
            {
                return childState;
            }
        }
    }

    return {};
}

[[nodiscard]] const Button* ResolveTreeButton(const std::unique_ptr<Control>& root, const Button* button) noexcept
{
    if (! button)
    {
        return nullptr;
    }

    return root && ControlBelongsToTree(root.get(), button) ? button : nullptr;
}

[[nodiscard]] Button* ResolveTreeButton(const std::unique_ptr<Control>& root, Button*& button) noexcept
{
    if (! button)
    {
        return nullptr;
    }

    if (! root || ! ControlBelongsToTree(root.get(), button))
    {
        button = nullptr;
        return nullptr;
    }

    return button;
}

struct ScopedPaint final
{
    HWND hwnd = nullptr;
    PAINTSTRUCT paint{};
    bool active = false;

    explicit ScopedPaint(HWND target) noexcept : hwnd(target)
    {
        if (hwnd)
        {
            BeginPaint(hwnd, &paint);
            active = true;
        }
    }

    ~ScopedPaint()
    {
        if (active && hwnd)
        {
            EndPaint(hwnd, &paint);
        }
    }

    ScopedPaint(const ScopedPaint&)            = delete;
    ScopedPaint& operator=(const ScopedPaint&) = delete;
};
} // namespace

void ShutdownAllWindowHostsForProcessExit() noexcept
{
    std::vector<WindowHost*> attachedHosts;
    {
        const std::scoped_lock lock(GetAttachedWindowHostsMutex());
        attachedHosts = GetAttachedWindowHosts();
    }

    for (WindowHost* const host : attachedHosts)
    {
        if (host)
        {
            host->Detach();
        }
    }

    ResetAllSharedWindowHostGraphicsResourcesForProcessExit();
}

#if defined(ENABLE_TESTS)
size_t DebugGetAttachedWindowHostCount() noexcept
{
    const std::scoped_lock lock(GetAttachedWindowHostsMutex());
    return GetAttachedWindowHosts().size();
}

uint32_t DebugGetSharedWindowHostAttachmentCountForThread(DWORD threadId) noexcept
{
    const std::scoped_lock lock(GetSharedWindowHostGraphicsResourcesMutex());
    const auto& resourcesByThread = GetSharedWindowHostGraphicsResourcesByThread();
    if (const auto it = resourcesByThread.find(threadId); it != resourcesByThread.end())
    {
        return it->second.attachedHostCount;
    }

    return 0u;
}
#endif

WindowHost::~WindowHost()
{
    // Real owners already detach during WM_NCDESTROY / dialog teardown. Keeping the
    // destructor passive avoids re-entering Win32 teardown from object destruction paths
    // such as plugin discovery probes, where no real host window may ever have existed.
}

bool WindowHost::Attach(HWND hwnd) noexcept
{
    return Attach(hwnd, AttachOptions{});
}

bool WindowHost::Attach(HWND hwnd, const AttachOptions& options) noexcept
{
    if (! hwnd)
    {
        return false;
    }

    if (_hwnd == nullptr)
    {
        _attachmentOwnerThreadId = GetCurrentThreadId();
        RegisterSharedWindowHostAttachment(this, _attachmentOwnerThreadId);
    }

    _hwnd             = hwnd;
    _presentationMode = options.presentationMode;
    RegisterWindowHostAccessibilityTarget(_hwnd, this);
    _dpi = GetDpiForWindow(hwnd);
    RECT client{};
    GetClientRect(hwnd, &client);
    _widthPx  = static_cast<UINT>(std::max(0L, client.right - client.left));
    _heightPx = static_cast<UINT>(std::max(0L, client.bottom - client.top));
    return true;
}

void WindowHost::Detach() noexcept
{
    if (_detachInProgress.exchange(true, std::memory_order_acq_rel))
    {
        return;
    }
    const auto clearDetachFlag = wil::scope_exit([&] { _detachInProgress.store(false, std::memory_order_release); });

    const HWND attachedHwnd  = _hwnd;
    const bool hadAttachment = attachedHwnd != nullptr;
    if (attachedHwnd && GetCapture() == attachedHwnd)
    {
        ReleaseCapture();
    }
    const DWORD attachmentOwnerThreadId = _attachmentOwnerThreadId;
    if (hadAttachment)
    {
        // Sever the Win32 attachment before tearing down retained state so any nested
        // message delivery cannot re-enter this host while its tree is partially cleared.
        UnregisterWindowHostAccessibilityTarget(attachedHwnd, this);
    }
    _attachmentOwnerThreadId = 0u;
    _hwnd                    = nullptr;

    if (_animationSubscriptionId != 0u)
    {
        Ui::AnimationDispatcher::GetInstance().Unsubscribe(_animationSubscriptionId);
        _animationSubscriptionId = 0u;
    }

    ResetRootInteractionState();
    if (_root)
    {
        _root->PropagateHost(nullptr);
    }
    _root.reset();
    _tooltipLayer.Clear();
    DiscardSizeDependentResources(L"detach");
    DiscardDeviceResources();
    _configuredTextFormats.clear();
    _defaultButton  = nullptr;
    _cancelButton   = nullptr;
    _onTabBoundary  = {};
    _onEscape       = {};
    _onFocusChanged = {};
    if (hadAttachment)
    {
        // Release the shared graphics bucket after this host's retained tree and device
        // resources are gone. The last host must not tear down shared D2D/DComp state
        // while its own controls/resources are still destructing.
        ReleaseSharedWindowHostAttachment(this, attachmentOwnerThreadId);
    }
}

void WindowHost::SetTheme(const ThemePalette& palette) noexcept
{
    const bool densityChanged = _palette.density != palette.density;
    _palette                  = palette;
    RecreateBrushCache();
    if (_root && densityChanged)
    {
        _root->OnDensityChanged();
    }
    Invalidate();
}

const ThemePalette& WindowHost::GetTheme() const noexcept
{
    return _palette;
}

void WindowHost::SetRoot(std::unique_ptr<Control> root)
{
    ResetRootInteractionState();
    if (_root)
    {
        _root->PropagateHost(nullptr);
    }
    _root          = std::move(root);
    _defaultButton = nullptr;
    _cancelButton  = nullptr;
    if (_root)
    {
        _root->SetParent(nullptr);
        _root->PropagateHost(this);
        _root->SetBounds(D2D1::RectF(0.0f, 0.0f, PixelsToDip(static_cast<float>(_widthPx)), PixelsToDip(static_cast<float>(_heightPx))));
    }
    Invalidate();
}

Control* WindowHost::GetRoot() noexcept
{
    return _root.get();
}

const Control* WindowHost::GetRoot() const noexcept
{
    return _root.get();
}

bool WindowHost::SetTooltip(std::wstring text, const D2D1_POINT_2F& originDip)
{
    const bool changed = _tooltipLayer.SetTooltip(std::move(text), originDip);
    if (changed)
    {
        Invalidate();
    }
    return changed;
}

bool WindowHost::SetTooltipDelayed(std::wstring text, const D2D1_POINT_2F& originDip)
{
    const uint64_t nowTickMs = _lastAnimationTickMs != 0u ? _lastAnimationTickMs : GetTickCount64();
    const bool changed       = _tooltipLayer.SetTooltipDelayed(std::move(text), originDip, nowTickMs, ResolveTooltipShowDelayMs());
    if (changed)
    {
        RequestAnimation();
        Invalidate();
    }
    return changed;
}

bool WindowHost::BeginTooltipHideDelay(uint64_t delayMs) noexcept
{
    const uint64_t nowTickMs = _lastAnimationTickMs != 0u ? _lastAnimationTickMs : GetTickCount64();
    const bool changed       = _tooltipLayer.BeginHideDelay(nowTickMs, delayMs);
    if (changed)
    {
        RequestAnimation();
    }
    return changed;
}

bool WindowHost::ClearTooltip() noexcept
{
    const bool changed = _tooltipLayer.Clear();
    if (changed)
    {
        Invalidate();
    }
    return changed;
}

bool WindowHost::HasTooltip() const noexcept
{
    return _tooltipLayer.HasTooltip();
}

std::wstring_view WindowHost::GetTooltipText() const noexcept
{
    return _tooltipLayer.GetTooltipText();
}

#if defined(ENABLE_TESTS)
std::wstring_view WindowHost::DebugGetPendingTooltipText() const noexcept
{
    return _tooltipLayer.DebugGetPendingTooltipText();
}

bool WindowHost::DebugAdvanceTooltipDelayForTest() noexcept
{
    const bool changed = _tooltipLayer.Tick(*this, GetTickCount64() + ResolveTooltipShowDelayMs() + 1u);
    if (changed)
    {
        Invalidate();
    }
    return changed;
}
#endif

void WindowHost::SetSmokeOverlayVisible(bool visible) noexcept
{
    if (_smokeOverlayVisible != visible)
    {
        _smokeOverlayVisible = visible;
        Invalidate();
    }
}

bool WindowHost::IsSmokeOverlayVisible() const noexcept
{
    return _smokeOverlayVisible;
}

bool WindowHost::SetSystemBackdrop(BackdropType type) noexcept
{
    if (! _hwnd)
    {
        return false;
    }

    // DWMWA_SYSTEMBACKDROP_TYPE = 38 — available on the supported Windows 11 baseline.
    constexpr DWORD kDwmwaSystemBackdropType = 38;
    const auto value                         = static_cast<int>(type);
    const HRESULT hr                         = DwmSetWindowAttribute(_hwnd, kDwmwaSystemBackdropType, &value, sizeof(value));
    return SUCCEEDED(hr);
}

#if defined(ENABLE_TESTS)
D2D1_RECT_F WindowHost::DebugGetTooltipBoundsDip() const noexcept
{
    return _tooltipLayer.DebugGetBoundsDip(*this);
}
#endif

void WindowHost::Invalidate() const noexcept
{
#if defined(ENABLE_TESTS)
    ++_debugInvalidateCount;
#endif
    if (_hwnd && IsHostWindowEffectivelyVisible(_hwnd))
    {
        InvalidateRect(_hwnd, nullptr, FALSE);
    }
}

bool WindowHost::PrimeForShow() noexcept
{
    return EnsureSizeDependentResources(true);
}

bool WindowHost::RenderInitialFrameForShow() noexcept
{
    if (! EnsureSizeDependentResources(true))
    {
        return false;
    }

    Render(nullptr, true);
    return true;
}

void WindowHost::RequestAnimation() noexcept
{
    if (_animationSubscriptionId == 0u)
    {
        _animationSubscriptionId = Ui::AnimationDispatcher::GetInstance().Subscribe(&WindowHost::AnimationTickThunk, this);
    }
}

void WindowHost::SetDefaultButton(Button* button) noexcept
{
    if (button && _root && ! ControlBelongsToTree(_root.get(), button))
    {
        return;
    }
    _defaultButton = button;
    Invalidate();
}

Button* WindowHost::GetDefaultButton() const noexcept
{
    return const_cast<Button*>(ResolveTreeButton(_root, _defaultButton));
}

void WindowHost::SetCancelButton(Button* button) noexcept
{
    if (button && _root && ! ControlBelongsToTree(_root.get(), button))
    {
        return;
    }
    _cancelButton = button;
}

Button* WindowHost::GetCancelButton() const noexcept
{
    return const_cast<Button*>(ResolveTreeButton(_root, _cancelButton));
}

void WindowHost::SetOnTabBoundary(std::function<bool(bool reverse)> onTabBoundary)
{
    _onTabBoundary = std::move(onTabBoundary);
}

void WindowHost::SetOnEscape(std::function<bool()> onEscape)
{
    _onEscape = std::move(onEscape);
}

void WindowHost::SetOnFocusChanged(std::function<void(Control* control)> onFocusChanged)
{
    _onFocusChanged = std::move(onFocusChanged);
}

void WindowHost::ResetInteractionState() noexcept
{
    ResetRootInteractionState();
    Invalidate();
}

void WindowHost::SetFocusControl(Control* control) noexcept
{
    PruneStaleInteractionState();
    if (control && _root)
    {
        const ControlInteractionState requestedState = ResolveControlInteractionState(_root.get(), control);
        if (! requestedState.inTree || ! requestedState.effectivelyInteractive)
        {
            return;
        }
    }
    if (control && ! control->IsFocusable())
    {
        control = nullptr;
    }
    if (_focusedControl == control)
    {
        bool restoredFocus = false;
        if (_focusedControl && ! _focusedControl->HasFocus())
        {
            _focusedControl->OnFocusChanged(*this, true);
            restoredFocus = true;
        }
        if (_focusedControl && _focusedControl->SupportsTextInput())
        {
            ActivateTextInput(_focusedControl);
        }
        if (restoredFocus)
        {
            Invalidate();
        }
        return;
    }
    if (_focusedControl && _focusedControl->SupportsTextInput())
    {
        DeactivateTextInput(false);
    }
    if (_focusedControl)
    {
        _focusedControl->OnFocusChanged(*this, false);
    }
    _focusedControl = control;
    if (_focusedControl)
    {
        _focusedControl->OnFocusChanged(*this, true);
    }
    if (_focusedControl && _focusedControl->SupportsTextInput())
    {
        ActivateTextInput(_focusedControl);
    }
    else if (_hwnd && _focusedControl && GetFocus() != _hwnd)
    {
        SetFocus(_hwnd);
    }
    if (IsInteractionDiagnosticsEnabled(_hwnd))
    {
#ifdef _DEBUG
        const uint64_t renderCount = DebugGetRenderCount();
        const uint64_t resizeCount = DebugGetResizeCount();
#else
        const uint64_t renderCount = 0;
        const uint64_t resizeCount = 0;
#endif
        Debug::Info(L"DxUi::WindowHost: focus-change hwnd={:#x} focus={} hovered={} captured={} textInput={} size={}x{} renderCount={} resizeCount={}",
                    reinterpret_cast<uintptr_t>(_hwnd),
                    static_cast<const void*>(_focusedControl),
                    static_cast<const void*>(_hoveredControl),
                    static_cast<const void*>(_capturedControl),
                    HasActiveTextInput() ? L"true" : L"false",
                    _widthPx,
                    _heightPx,
                    renderCount,
                    resizeCount);
    }
    Debug::Perf::Emit(L"DxUI::FocusChange", L"", 0u, _focusedControl ? 1u : 0u, HasActiveTextInput() ? 1u : 0u);
    if (_onFocusChanged)
    {
        _onFocusChanged(_focusedControl);
    }
    Invalidate();
}

Control* WindowHost::GetFocusControl() const noexcept
{
    return _focusedControl;
}

bool WindowHost::HandleMnemonic(wchar_t mnemonic) noexcept
{
    if (! _root)
    {
        return false;
    }

    const wchar_t normalizedMnemonic = NormalizeMnemonicChar(mnemonic);
    if (normalizedMnemonic == L'\0')
    {
        return false;
    }

    if (Control* target = FindMnemonicControl(_root.get(), normalizedMnemonic))
    {
        SetInputModality(InputModality::Keyboard);
        return target->OnMnemonic(*this);
    }

    return false;
}

void WindowHost::RecordNativeTextInputKeyToStateMetric(
    std::chrono::steady_clock::time_point inputStartedAt, const wchar_t* detail, uint64_t value0, uint64_t value1, bool armPaintMetric) noexcept
{
    Debug::Perf::Emit(L"dxui.textinput.key_to_state_us", detail, Debug::Perf::ElapsedUs(inputStartedAt), value0, value1, S_OK);
    if (! armPaintMetric || ! Debug::Perf::IsCaptureEnabled())
    {
        return;
    }

    _pendingNativeTextInputPaintMetric = PendingNativeTextInputPaintMetric{inputStartedAt, detail ? detail : L"", value0, value1};
}

void WindowHost::EmitPendingNativeTextInputPaintMetric(HRESULT hr) noexcept
{
    if (! _pendingNativeTextInputPaintMetric.has_value())
    {
        return;
    }

    const PendingNativeTextInputPaintMetric metric = _pendingNativeTextInputPaintMetric.value();
    _pendingNativeTextInputPaintMetric.reset();
    Debug::Perf::Emit(L"dxui.textinput.key_to_paint_us",
                      metric.detail ? metric.detail : L"",
                      Debug::Perf::ElapsedUs(metric.inputStartedAt),
                      metric.value0,
                      metric.value1,
                      hr);
}

bool WindowHost::RouteFocusedCharInput(wchar_t ch, UINT modifiers, const wchar_t* perfDetail) noexcept
{
    if (! _focusedControl)
    {
        return false;
    }

    Control* const charTarget = _focusedControl;
    const auto inputStartedAt = std::chrono::steady_clock::now();
    const bool controlHandled = charTarget->OnChar(*this, ch, modifiers);
    if (controlHandled && charTarget == _focusedControl && charTarget->SupportsTextInput())
    {
        SyncNativeTextInputSession(charTarget);
        RecordNativeTextInputKeyToStateMetric(inputStartedAt, perfDetail, static_cast<uint64_t>(ch), modifiers);
    }
    return controlHandled;
}

bool WindowHost::HandleTabNavigation(bool reverse) noexcept
{
    if (! _root)
    {
        return false;
    }

    const FocusAdvanceResult next = FindAdjacentFocusable(_root.get(), _focusedControl, reverse);
    Debug::Perf::Emit(L"dxui.focus.tab_navigation",
                      L"",
                      reverse ? 1u : 0u,
                      next.currentIndex == std::numeric_limits<size_t>::max() ? UINT64_MAX : static_cast<uint64_t>(next.currentIndex),
                      next.nextIndex == std::numeric_limits<size_t>::max() ? UINT64_MAX : static_cast<uint64_t>(next.nextIndex),
                      S_OK);
    if (next.wrapped && _focusedControl && _onTabBoundary && _onTabBoundary(reverse))
    {
        return true;
    }

    SetFocusControl(next.control);
    return true;
}

void WindowHost::CaptureMouse(Control* control) noexcept
{
    if (control && _root && ! ControlBelongsToTree(_root.get(), control))
    {
        return;
    }
    _capturedControl = control;
    // Some controls capture during their own mouse-down handling; avoid a
    // redundant SetCapture that can immediately re-enter capture-lost cleanup.
    if (_hwnd && GetCapture() != _hwnd)
    {
        SetCapture(_hwnd);
    }
    if (IsInteractionDiagnosticsEnabled(_hwnd))
    {
        Debug::Info(L"DxUi::WindowHost: capture hwnd={:#x} target={} focus={} textInput={}",
                    reinterpret_cast<uintptr_t>(_hwnd),
                    static_cast<const void*>(_capturedControl),
                    static_cast<const void*>(_focusedControl),
                    HasActiveTextInput() ? L"true" : L"false");
    }
}

void WindowHost::ReleaseMouseCapture() noexcept
{
    _capturedControl = nullptr;
    if (_hwnd && GetCapture() == _hwnd)
    {
        ReleaseCapture();
    }
    if (IsInteractionDiagnosticsEnabled(_hwnd))
    {
        Debug::Info(L"DxUi::WindowHost: release-capture hwnd={:#x} focus={} hovered={} textInput={}",
                    reinterpret_cast<uintptr_t>(_hwnd),
                    static_cast<const void*>(_focusedControl),
                    static_cast<const void*>(_hoveredControl),
                    HasActiveTextInput() ? L"true" : L"false");
    }
}

void WindowHost::ClearPendingPointerDoubleClick() noexcept
{
    _pendingPointerDoubleClick = {};
}

bool WindowHost::ShouldTreatButtonDownAsDoubleClick(Control* target, UINT buttonDownMessage, LPARAM lp) const noexcept
{
    if (! target || ! IsPointerButtonDownMessage(buttonDownMessage))
    {
        return false;
    }

    const auto& candidate = _pendingPointerDoubleClick;
    if (! candidate.target || candidate.target != target || candidate.downMessage != buttonDownMessage || candidate.tickMs == 0u)
    {
        return false;
    }

    const uint64_t nowTickMs = GetTickCount64();
    if (nowTickMs - candidate.tickMs > static_cast<uint64_t>(GetDoubleClickTime()))
    {
        return false;
    }

    const POINT pointPx{GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
    return IsWithinSystemDoubleClickBounds(candidate.pointPx, pointPx);
}

void WindowHost::RememberPointerButtonDown(Control* target, UINT buttonDownMessage, LPARAM lp) noexcept
{
    if (! target || ! IsPointerButtonDownMessage(buttonDownMessage))
    {
        ClearPendingPointerDoubleClick();
        return;
    }

    _pendingPointerDoubleClick.target      = target;
    _pendingPointerDoubleClick.downMessage = buttonDownMessage;
    _pendingPointerDoubleClick.pointPx     = POINT{GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
    _pendingPointerDoubleClick.tickMs      = GetTickCount64();
}

HWND WindowHost::GetHwnd() const noexcept
{
    return _hwnd;
}

D2D1_RECT_F WindowHost::GetClientBoundsDip() const noexcept
{
    return D2D1::RectF(0.0f, 0.0f, PixelsToDip(static_cast<float>(_widthPx)), PixelsToDip(static_cast<float>(_heightPx)));
}

float WindowHost::GetDpi() const noexcept
{
    return static_cast<float>(_dpi == 0 ? USER_DEFAULT_SCREEN_DPI : _dpi);
}

float WindowHost::PixelsToDip(float pixels) const noexcept
{
    return (pixels * USER_DEFAULT_SCREEN_DPI) / GetDpi();
}

float WindowHost::DipsToPixels(float dips) const noexcept
{
    return (dips * GetDpi()) / USER_DEFAULT_SCREEN_DPI;
}

PointDip WindowHost::PixelsToDipPoint(POINT pointPx) const noexcept
{
    return MakePointDip(PixelsToDip(static_cast<float>(pointPx.x)), PixelsToDip(static_cast<float>(pointPx.y)));
}

std::optional<PointDip> WindowHost::ScreenPointToDipPoint(POINT screenPointPx) const noexcept
{
    if (! _hwnd || IsWindow(_hwnd) == FALSE)
    {
        return std::nullopt;
    }

    POINT clientPointPx = screenPointPx;
    if (ScreenToClient(_hwnd, &clientPointPx) == FALSE)
    {
        return std::nullopt;
    }

    return PixelsToDipPoint(clientPointPx);
}

POINT WindowHost::DipPointToScreenPoint(D2D1_POINT_2F pointDip) const noexcept
{
    POINT pointPx{static_cast<LONG>(std::lround(DipsToPixels(pointDip.x))), static_cast<LONG>(std::lround(DipsToPixels(pointDip.y)))};
    if (_hwnd)
    {
        ClientToScreen(_hwnd, &pointPx);
    }
    return pointPx;
}

InputModality WindowHost::GetInputModality() const noexcept
{
    return _inputModality;
}

bool WindowHost::IsKeyboardFocusVisible() const noexcept
{
    return _inputModality == InputModality::Keyboard;
}

ID2D1DeviceContext* WindowHost::GetDeviceContext() const noexcept
{
    return _d2dContext.get();
}

IDWriteFactory* WindowHost::GetWriteFactory() const noexcept
{
    if (! _dwriteFactory)
    {
        if (! EnsureDeviceIndependentResources())
        {
            return nullptr;
        }
    }
    return _dwriteFactory.get();
}

IDWriteTextFormat* WindowHost::GetTextFormat(FontRole role) const noexcept
{
    if ((! _dwriteFactory || ! _bodyTextFormat || ! _bodyStrongTextFormat || ! _bodyLargeTextFormat || ! _listItemTextFormat || ! _titleTextFormat ||
         ! _subtitleTextFormat || ! _titleLargeTextFormat || ! _displayTextFormat || ! _headerTextFormat || ! _smallTextFormat || ! _monoTextFormat ||
         (_fluentIconFontAvailable && (! _iconTextFormat || ! _heroIconTextFormat))))
    {
        if (! EnsureDeviceIndependentResources())
        {
            return nullptr;
        }
    }

    switch (role)
    {
        case FontRole::BodyStrong: return _bodyStrongTextFormat.get();
        case FontRole::BodyLarge: return _bodyLargeTextFormat.get();
        case FontRole::ListItem: return _listItemTextFormat.get();
        case FontRole::Title: return _titleTextFormat.get();
        case FontRole::Subtitle: return _subtitleTextFormat.get();
        case FontRole::TitleLarge: return _titleLargeTextFormat.get();
        case FontRole::Display: return _displayTextFormat.get();
        case FontRole::Header: return _headerTextFormat.get();
        case FontRole::Small: return _smallTextFormat.get();
        case FontRole::Icon: return _iconTextFormat ? _iconTextFormat.get() : _smallTextFormat.get();
        case FontRole::HeroIcon: return _heroIconTextFormat ? _heroIconTextFormat.get() : (_iconTextFormat ? _iconTextFormat.get() : _smallTextFormat.get());
        case FontRole::Monospace: return _monoTextFormat.get();
        case FontRole::Body:
        default: return _bodyTextFormat.get();
    }
}

IDWriteTextFormat* WindowHost::GetTextFormat(FontRole role,
                                             DWRITE_TEXT_ALIGNMENT alignment,
                                             DWRITE_PARAGRAPH_ALIGNMENT paragraphAlignment,
                                             bool wrap,
                                             DWRITE_READING_DIRECTION readingDirection) const noexcept
{
    if (! _dwriteFactory)
    {
        if (! EnsureDeviceIndependentResources())
        {
            return nullptr;
        }
    }

    const uint64_t key = PackTextFormatKey(role, alignment, paragraphAlignment, wrap, readingDirection);
    const auto it      = _configuredTextFormats.find(key);
    if (it != _configuredTextFormats.end())
    {
        return it->second.get();
    }

    if ((role == FontRole::Icon || role == FontRole::HeroIcon) && ! HasFluentIconFont())
    {
        return GetTextFormat(FontRole::Small, alignment, paragraphAlignment, wrap);
    }

    if (! _dwriteFactory)
    {
        return GetTextFormat(role);
    }

    const Typography::TypographySpec spec = Typography::GetDxUiTypographySpec(role);
    wil::com_ptr<IDWriteTextFormat> format;
    const HRESULT hrCreate = Typography::CreateTextFormat(_dwriteFactory.get(), spec, format.addressof());
    if (FAILED(hrCreate) || ! format)
    {
        Debug::Warning(L"DxUi::WindowHost: CreateTextFormat failed for configured font role {}: 0x{:08X}", static_cast<uint32_t>(role), hrCreate);
        return GetTextFormat(role);
    }

    const HRESULT hrAlignment        = format->SetTextAlignment(alignment);
    const HRESULT hrParagraph        = format->SetParagraphAlignment(paragraphAlignment);
    const HRESULT hrWrapping         = format->SetWordWrapping(wrap ? DWRITE_WORD_WRAPPING_WRAP : DWRITE_WORD_WRAPPING_NO_WRAP);
    const HRESULT hrReadingDirection = format->SetReadingDirection(readingDirection);
    if (FAILED(hrAlignment) || FAILED(hrParagraph) || FAILED(hrWrapping) || FAILED(hrReadingDirection))
    {
        Debug::Warning(L"DxUi::WindowHost: failed to configure text format role {} alignment/paragraph/wrap/reading ({:08X}, {:08X}, {:08X}, {:08X})",
                       static_cast<uint32_t>(role),
                       hrAlignment,
                       hrParagraph,
                       hrWrapping,
                       hrReadingDirection);
        return GetTextFormat(role);
    }

    return _configuredTextFormats.emplace(key, std::move(format)).first->second.get();
}

bool WindowHost::HasFluentIconFont() const noexcept
{
    if (! _dwriteFactory || ! _fluentIconFontAvailabilityChecked)
    {
        if (! EnsureDeviceIndependentResources())
        {
            return false;
        }
    }

    return _fluentIconFontAvailable;
}

ID2D1SolidColorBrush* WindowHost::GetSolidBrush(const D2D1_COLOR_F& color) const
{
#if defined(ENABLE_TESTS)
    if (_debugForceNullSolidBrushes)
    {
        return nullptr;
    }
#endif

    const uint32_t key = PackColor(color);
    const auto it      = _brushCache.find(key);
    if (it != _brushCache.end())
    {
        return it->second.get();
    }
    if (! _d2dContext)
    {
        return nullptr;
    }

    wil::com_ptr<ID2D1SolidColorBrush> brush;
    const HRESULT hr = _d2dContext->CreateSolidColorBrush(color, brush.addressof());
    if (FAILED(hr) || ! brush)
    {
        if (std::ranges::find(_brushFailureLogKeys, key) == _brushFailureLogKeys.end())
        {
            _brushFailureLogKeys.push_back(key);
            Debug::Warning(L"DxUi::WindowHost: CreateSolidColorBrush failed for ARGB 0x{:08X}: 0x{:08X}", key, hr);
        }
        if (! _fallbackBrush && _d2dContext)
        {
            static_cast<void>(_d2dContext->CreateSolidColorBrush(D2D1::ColorF(1.0f, 1.0f, 1.0f, 1.0f), _fallbackBrush.addressof()));
        }
        if (_fallbackBrush)
        {
            _fallbackBrush->SetColor(color);
            return _fallbackBrush.get();
        }
        return nullptr;
    }
    _brushCache[key] = brush;
    return brush.get();
}

bool WindowHost::CopyTextToClipboard(std::wstring_view text) const noexcept
{
    return SetClipboardUnicodeText(_hwnd, text);
}

std::optional<std::wstring> WindowHost::ReadTextFromClipboard() const noexcept
{
    return TryReadClipboardUnicodeText(_hwnd);
}

#if defined(ENABLE_TESTS)
uint64_t WindowHost::DebugGetInvalidateCount() const noexcept
{
    return _debugInvalidateCount;
}

UINT WindowHost::DebugGetModifierState() const noexcept
{
    return _modifierState;
}

void DebugClearClipboardFallbackText() noexcept
{
    const std::lock_guard lock(GetClipboardFallbackMutex());
    GetClipboardFallbackTextStorage().reset();
}

bool DebugSetClipboardFallbackText(std::wstring_view text) noexcept
{
    const std::lock_guard lock(GetClipboardFallbackMutex());
    GetClipboardFallbackTextStorage() = std::wstring(text);
    return true;
}

std::optional<std::wstring> DebugReadClipboardFallbackText() noexcept
{
    const std::lock_guard lock(GetClipboardFallbackMutex());
    return GetClipboardFallbackTextStorage();
}

bool DebugWriteClipboardUnicodeText(HWND ownerWindow, std::wstring_view text) noexcept
{
    return SetClipboardUnicodeText(ownerWindow, text);
}

uint64_t WindowHost::DebugGetRenderCount() const noexcept
{
    return _debugRenderCount;
}

uint64_t WindowHost::DebugGetResizeCount() const noexcept
{
    return _debugResizeCount;
}

uint64_t WindowHost::DebugGetResizeFailureCount() const noexcept
{
    return _debugResizeFailureCount;
}

uint64_t WindowHost::DebugGetSwapChainPrepareD2DFlushFailureCount() const noexcept
{
    return _debugSwapChainPrepareD2DFlushFailureCount;
}

uint64_t WindowHost::DebugGetPresentFailureCount() const noexcept
{
    return _debugPresentFailureCount;
}

bool WindowHost::DebugHasActiveAnimationSubscription() const noexcept
{
    return _animationSubscriptionId != 0u;
}

bool WindowHost::DebugAnimationTickForTest(uint64_t nowTickMs) noexcept
{
    return OnAnimationTick(nowTickMs);
}

IRawElementProviderFragmentRoot* WindowHost::DebugCreateAccessibilityProvider() const noexcept
{
    return _hwnd ? CreateWindowHostAccessibilityProvider(_hwnd) : nullptr;
}

void WindowHost::DebugSimulateDeviceLoss() noexcept
{
    DiscardSizeDependentResources(L"debug-simulate-device-loss");
    DiscardDeviceResources();
    ResetSharedWindowHostGraphicsResources();
    Invalidate();
}

size_t WindowHost::DebugGetBrushCacheSize() const noexcept
{
    return _brushCache.size();
}

bool WindowHost::DebugHasFallbackBrush() const noexcept
{
    return _fallbackBrush != nullptr;
}

bool WindowHost::DebugHasD2DContext() const noexcept
{
    return _d2dContext != nullptr;
}

size_t WindowHost::DebugGetConfiguredTextFormatCount() const noexcept
{
    return _configuredTextFormats.size();
}

void WindowHost::DebugSetForceNullSolidBrushes(const bool force) noexcept
{
    _debugForceNullSolidBrushes = force;
}

const Control* WindowHost::DebugHitTestControl(D2D1_POINT_2F pointDip) noexcept
{
    return HitTestControl(pointDip);
}

WindowHostCursorKind WindowHost::DebugResolveCursorKindForPoint(D2D1_POINT_2F pointDip) noexcept
{
    return ResolveCursorKindForPoint(pointDip);
}

bool WindowHost::DebugCaptureBitmap(WindowHostBitmapCapture& out) noexcept
{
    out = {};
    if (! _hwnd)
    {
        return false;
    }

    Render(nullptr, &out);
    return out.widthPx > 0u && out.heightPx > 0u && ! out.bgraPixels.empty();
}
#endif

void WindowHost::OnDpiChanged(HWND hwnd, UINT newDpi, const RECT* suggestedRect) noexcept
{
    Debug::Perf::Scope perf(L"dxui.windowhost.dpi_change_us");

    const UINT resolvedDpi = newDpi != 0u ? newDpi : (hwnd ? GetDpiForWindow(hwnd) : USER_DEFAULT_SCREEN_DPI);
    if (resolvedDpi == 0u)
    {
        return;
    }

    _dpi = resolvedDpi;
    if (_d2dContext)
    {
        _d2dContext->SetDpi(static_cast<float>(_dpi), static_cast<float>(_dpi));
    }

    if (_root)
    {
        _root->OnHostDpiChanged(*this);
        _root->SetBounds(D2D1::RectF(0.0f, 0.0f, PixelsToDip(static_cast<float>(_widthPx)), PixelsToDip(static_cast<float>(_heightPx))));
    }

    _tooltipLayer.OnHostDpiChanged(*this);

    if (hwnd && ((GetWindowLongPtrW(hwnd, GWL_STYLE) & WS_CHILD) == 0) && suggestedRect)
    {
        SetWindowPos(hwnd,
                     nullptr,
                     suggestedRect->left,
                     suggestedRect->top,
                     suggestedRect->right - suggestedRect->left,
                     suggestedRect->bottom - suggestedRect->top,
                     SWP_NOZORDER | SWP_NOACTIVATE);
    }

    Invalidate();
}

LRESULT WindowHost::HandleMessage(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp, bool& handled) noexcept
{
    handled = false;
    if (hwnd != _hwnd)
    {
        return 0;
    }

    PruneStaleInteractionState();

    LRESULT accessibilityResult = 0;
    if (TryHandleWindowHostAccessibilityMessage(hwnd, msg, wp, lp, accessibilityResult))
    {
        handled = true;
        return accessibilityResult;
    }

    if (ShouldTraceWindowHostMessage(msg))
    {
        POINT cursorScreenPt{};
        const bool haveCursorScreenPt = GetCursorPos(&cursorScreenPt) != FALSE; // getcursorpos-allow: diagnostic-only
        POINT cursorClientPt          = cursorScreenPt;
        const bool haveCursorClientPt = haveCursorScreenPt && ScreenToClient(hwnd, &cursorClientPt) != FALSE;
        POINT messageClientPt{};
        bool haveMessageClientPt = false;
        POINT messageScreenPt{};
        bool haveMessageScreenPt = false;
        Control* messageHitControl = nullptr;
        if (IsClientMouseMessageForWindowHostTrace(msg))
        {
            messageClientPt     = POINT{GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
            haveMessageClientPt = true;
            messageScreenPt     = messageClientPt;
            haveMessageScreenPt = ClientToScreen(hwnd, &messageScreenPt) != FALSE;
            if (_root)
            {
                messageHitControl = HitTestControl(PointFromLParam(lp));
            }
        }
        const POINT hitScreenPt = haveMessageScreenPt ? messageScreenPt : cursorScreenPt;
        const bool haveHitScreenPt = haveMessageScreenPt || haveCursorScreenPt;
        const HWND windowAtPoint =
            (haveHitScreenPt && ShouldResolveScreenHitWindowForWindowHostTrace(msg)) ? WindowFromPoint(hitScreenPt) : nullptr;
        const HWND parentHwnd = GetParent(hwnd);
        const HWND ownerHwnd  = GetWindow(hwnd, GW_OWNER);
        TraceWindowHostDiagnostics(L"dxui.windowhost.raw",
                                   L"phase=enter hwnd={:#x} class=\"{}\" parent={:#x} owner={:#x} message={} msg=0x{:x} wParam={:#x} "
                                   L"lParam={:#x} cursorClientPt=({}, {}) haveCursorClient={} cursorScreenPt=({}, {}) haveCursorScreen={} "
                                   L"messageClientPt=({}, {}) haveMessageClient={} messageScreenPt=({}, {}) haveMessageScreen={} "
                                   L"windowAtPoint={:#x} windowAtPointClass=\"{}\" focus={:#x} active={:#x} foreground={:#x} capture={:#x} "
                                   L"messageHitControl={} focusedControl={} hoveredControl={} capturedControl={} root={} size={}x{}",
                                   reinterpret_cast<uintptr_t>(hwnd),
                                   TraceWindowClassName(hwnd),
                                   reinterpret_cast<uintptr_t>(parentHwnd),
                                   reinterpret_cast<uintptr_t>(ownerHwnd),
                                   TraceWindowHostMessageName(msg),
                                   msg,
                                   static_cast<uintptr_t>(wp),
                                   static_cast<uintptr_t>(lp),
                                   cursorClientPt.x,
                                   cursorClientPt.y,
                                   haveCursorClientPt ? 1 : 0,
                                   cursorScreenPt.x,
                                   cursorScreenPt.y,
                                   haveCursorScreenPt ? 1 : 0,
                                   messageClientPt.x,
                                   messageClientPt.y,
                                   haveMessageClientPt ? 1 : 0,
                                   messageScreenPt.x,
                                   messageScreenPt.y,
                                   haveMessageScreenPt ? 1 : 0,
                                   reinterpret_cast<uintptr_t>(windowAtPoint),
                                   TraceWindowClassName(windowAtPoint),
                                   reinterpret_cast<uintptr_t>(GetFocus()),
                                   reinterpret_cast<uintptr_t>(GetActiveWindow()),
                                   reinterpret_cast<uintptr_t>(GetForegroundWindow()),
                                   reinterpret_cast<uintptr_t>(GetCapture()),
                                   DescribeWindowHostTraceControl(messageHitControl),
                                   DescribeWindowHostTraceControl(_focusedControl),
                                   DescribeWindowHostTraceControl(_hoveredControl),
                                   DescribeWindowHostTraceControl(_capturedControl),
                                   DescribeWindowHostTraceControl(_root.get()),
                                   _widthPx,
                                   _heightPx);
    }

    switch (msg)
    {
        case WM_PAINT:
            handled = true;
            {
                ScopedPaint paint(hwnd);
                Render(&paint.paint.rcPaint);
            }
            return 0;
        case WM_ERASEBKGND: handled = true; return 1;
        case WM_SIZE:
            handled = true;
            OnSize(static_cast<UINT>(LOWORD(lp)), static_cast<UINT>(HIWORD(lp)));
            return 0;
        case WM_SHOWWINDOW:
            Debug::Warning(L"DxUi::WindowHost: WM_SHOWWINDOW hwnd=0x{:X}  show={}", reinterpret_cast<uintptr_t>(hwnd), static_cast<int>(wp));
            // Do NOT discard the swap chain when the window is hidden.
            // WM_SETREDRAW FALSE on a parent sends WM_SHOWWINDOW FALSE
            // to children, but WM_SETREDRAW TRUE does NOT reliably send
            // WM_SHOWWINDOW TRUE back.  Discarding here permanently
            // breaks the swap chain for ScopedWindowRedrawBlock users.
            // The swap chain is cleaned up on Detach() / WM_NCDESTROY.
            if (wp != FALSE)
            {
                Invalidate();
            }
            handled = false;
            return 0;
        case WM_DPICHANGED:
            OnDpiChanged(hwnd, static_cast<UINT>(HIWORD(wp)), reinterpret_cast<const RECT*>(lp));
            handled = true;
            return 0;
        case WM_DPICHANGED_AFTERPARENT:
            OnDpiChanged(hwnd, GetDpiForWindow(hwnd), nullptr);
            handled = true;
            return 0;
        case WM_SETFOCUS:
            handled = true;
            OnSetFocus();
            return 0;
        case WM_KILLFOCUS:
            handled = true;
            if (IsWindowOrDescendant(hwnd, reinterpret_cast<HWND>(wp)))
            {
                return 0;
            }
            OnKillFocus(false);
            return 0;
        case WM_ACTIVATE:
            if (LOWORD(static_cast<DWORD_PTR>(wp)) != WA_INACTIVE)
            {
                break;
            }
            handled = true;
            OnKillFocus(false);
            return 0;
        case WM_ACTIVATEAPP:
            if (wp != FALSE)
            {
                break;
            }
            handled = true;
            OnKillFocus(false);
            return 0;
        case WM_MOUSEACTIVATE: handled = true; return MA_ACTIVATE;
        case WM_MOUSEMOVE:
        {
            handled = true;
            TRACKMOUSEEVENT tme{};
            tme.cbSize    = sizeof(tme);
            tme.dwFlags   = TME_LEAVE;
            tme.hwndTrack = hwnd;
            TrackMouseEvent(&tme);
            UpdateHover(PointFromLParam(lp), static_cast<UINT>(wp));
            return 0;
        }
        case WM_SETCURSOR:
        {
            if (LOWORD(lp) != HTCLIENT)
            {
                handled = false;
                return 0;
            }

            POINT pointPx = {GET_X_LPARAM(GetMessagePos()), GET_Y_LPARAM(GetMessagePos())};
            ScreenToClient(hwnd, &pointPx);
            const D2D1_POINT_2F pointDip = D2D1::Point2F(PixelsToDip(static_cast<float>(pointPx.x)), PixelsToDip(static_cast<float>(pointPx.y)));
            SetCursor(ResolveCursorHandle(ResolveCursorKindForPoint(pointDip)));
            handled = true;
            return TRUE;
        }
        case WM_MOUSELEAVE:
        {
            handled = true;
            if (_hoveredControl)
            {
                _hoveredControl->OnHoverChanged(*this, false);
                _hoveredControl->OnMouseLeave(*this);
                _hoveredControl = nullptr;
            }
            ClearTooltip();
            Invalidate();
            return 0;
        }
        case WM_LBUTTONDOWN:
        case WM_RBUTTONDOWN:
        case WM_LBUTTONDBLCLK:
        case WM_RBUTTONDBLCLK:
        {
            handled = true;
            SetInputModality(InputModality::Pointer);
            const D2D1_POINT_2F point    = PointFromLParam(lp);
            Control* target              = _capturedControl;
            const UINT buttonDownMessage = PointerButtonDownMessageFor(msg);
            if (! target && IsPointerDoubleClickMessage(msg))
            {
                const auto& candidate = _pendingPointerDoubleClick;
                const POINT pointPx{GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
                if (candidate.target && candidate.downMessage == buttonDownMessage && candidate.tickMs != 0u &&
                    GetTickCount64() - candidate.tickMs <= static_cast<uint64_t>(GetDoubleClickTime()) &&
                    IsWithinSystemDoubleClickBounds(candidate.pointPx, pointPx) && ControlBelongsToTree(_root.get(), candidate.target))
                {
                    target = candidate.target;
                }
            }
            if (! target)
            {
                target = HitTestControl(point);
            }
            const bool rightButton = msg == WM_RBUTTONDOWN || msg == WM_RBUTTONDBLCLK;
            // Many DxUi hosts attach to existing STATIC/custom windows that never emit
            // WM_*BUTTONDBLCLK, so detect the second down here and route it through the
            // same control-level double-click path.
            const bool doubleClick       = IsPointerDoubleClickMessage(msg) || ShouldTreatButtonDownAsDoubleClick(target, buttonDownMessage, lp);
            const bool controlHandled    = target && (doubleClick ? target->OnMouseDoubleClick(*this, point, rightButton, static_cast<UINT>(wp))
                                                                  : target->OnMouseDown(*this, point, rightButton, static_cast<UINT>(wp)));
            if (IsContextMenuDiagnosticsEnabled())
            {
                const POINT messagePointPx{GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
                POINT liveScreenPx{};
                const bool haveLiveScreen = GetCursorPos(&liveScreenPx) != FALSE; // getcursorpos-allow: diagnostic-only
                POINT liveClientPx        = liveScreenPx;
                const bool haveLiveClient = haveLiveScreen && ScreenToClient(hwnd, &liveClientPx) != FALSE;
                RECT liveClientRectPx{};
                const bool haveClientRect = GetClientRect(hwnd, &liveClientRectPx) != FALSE;
                const bool liveInClient =
                    haveLiveClient && haveClientRect && PtInRect(&liveClientRectPx, liveClientPx) != FALSE;
                const HWND liveWindowAtPoint = haveLiveScreen ? WindowFromPoint(liveScreenPx) : nullptr;
                TraceWindowHostDiagnostics(
                    L"dxui.windowhost.pointer-down",
                    L"hwnd={:#x} message={} point=({:.1f}, {:.1f}) messageClientPx=({}, {}) liveClientPx=({}, {}) haveLiveClient={} "
                    L"liveInClient={} liveWindowAtPoint={:#x} target={} controlHandled={} doubleClick={} rightButton={} focusBefore={} hovered={} capturedBefore={}",
                    reinterpret_cast<uintptr_t>(hwnd),
                    TraceWindowHostMessageName(msg),
                    point.x,
                    point.y,
                    messagePointPx.x,
                    messagePointPx.y,
                    liveClientPx.x,
                    liveClientPx.y,
                    haveLiveClient ? 1 : 0,
                    liveInClient ? 1 : 0,
                    reinterpret_cast<uintptr_t>(liveWindowAtPoint),
                    DescribeWindowHostTraceControl(target),
                    controlHandled ? 1 : 0,
                    doubleClick ? 1 : 0,
                    rightButton ? 1 : 0,
                    DescribeWindowHostTraceControl(_focusedControl),
                    DescribeWindowHostTraceControl(_hoveredControl),
                    DescribeWindowHostTraceControl(_capturedControl));
            }
            if (controlHandled)
            {
                if (target->IsFocusable())
                {
                    if (target->SupportsTextInput())
                    {
                        SetFocusControl(target);
                    }
                    else
                    {
                        SetFocus(hwnd);
                        SetFocusControl(target);
                    }
                }
                CaptureMouse(target);
                if (doubleClick)
                {
                    ClearPendingPointerDoubleClick();
                }
                else
                {
                    RememberPointerButtonDown(target, buttonDownMessage, lp);
                }
            }
            else if (_focusedControl && target != _focusedControl)
            {
                SetFocusControl(nullptr);
                ClearPendingPointerDoubleClick();
            }
            else if (! controlHandled)
            {
                ClearPendingPointerDoubleClick();
            }
            return 0;
        }
        case WM_LBUTTONUP:
        case WM_RBUTTONUP:
        {
            handled                   = true;
            const D2D1_POINT_2F point = PointFromLParam(lp);
            Control* target           = _capturedControl;
            if (! target)
            {
                target = HitTestControl(point);
            }
            const bool rightButton = msg == WM_RBUTTONUP;
            if (target)
            {
                target->OnMouseUp(*this, point, rightButton, static_cast<UINT>(wp));
            }
            if (IsContextMenuDiagnosticsEnabled())
            {
                const POINT messagePointPx{GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
                POINT liveScreenPx{};
                const bool haveLiveScreen = GetCursorPos(&liveScreenPx) != FALSE; // getcursorpos-allow: diagnostic-only
                POINT liveClientPx        = liveScreenPx;
                const bool haveLiveClient = haveLiveScreen && ScreenToClient(hwnd, &liveClientPx) != FALSE;
                RECT liveClientRectPx{};
                const bool haveClientRect = GetClientRect(hwnd, &liveClientRectPx) != FALSE;
                const bool liveInClient =
                    haveLiveClient && haveClientRect && PtInRect(&liveClientRectPx, liveClientPx) != FALSE;
                const HWND liveWindowAtPoint = haveLiveScreen ? WindowFromPoint(liveScreenPx) : nullptr;
                TraceWindowHostDiagnostics(
                    L"dxui.windowhost.pointer-up",
                    L"hwnd={:#x} message={} point=({:.1f}, {:.1f}) messageClientPx=({}, {}) liveClientPx=({}, {}) haveLiveClient={} "
                    L"liveInClient={} liveWindowAtPoint={:#x} target={} rightButton={} focus={} hovered={} capturedBeforeRelease={}",
                    reinterpret_cast<uintptr_t>(hwnd),
                    TraceWindowHostMessageName(msg),
                    point.x,
                    point.y,
                    messagePointPx.x,
                    messagePointPx.y,
                    liveClientPx.x,
                    liveClientPx.y,
                    haveLiveClient ? 1 : 0,
                    liveInClient ? 1 : 0,
                    reinterpret_cast<uintptr_t>(liveWindowAtPoint),
                    DescribeWindowHostTraceControl(target),
                    rightButton ? 1 : 0,
                    DescribeWindowHostTraceControl(_focusedControl),
                    DescribeWindowHostTraceControl(_hoveredControl),
                    DescribeWindowHostTraceControl(_capturedControl));
            }
            ReleaseMouseCapture();
            return 0;
        }
        case WM_MOUSEWHEEL:
        {
            POINT pointPx{GET_X_LPARAM(lp), GET_Y_LPARAM(lp)};
            ScreenToClient(hwnd, &pointPx);
            const D2D1_POINT_2F point = D2D1::Point2F(PixelsToDip(static_cast<float>(pointPx.x)), PixelsToDip(static_cast<float>(pointPx.y)));
            Control* target           = _capturedControl;
            if (! target)
            {
                target = HitTestControl(point);
            }
            if (target)
            {
                handled = target->OnMouseWheel(*this, point, static_cast<float>(GET_WHEEL_DELTA_WPARAM(wp)), ComputeModifierMask());
                return 0;
            }
            handled = false;
            return 0;
        }
        case WM_KEYDOWN:
        case WM_SYSKEYDOWN:
        {
            handled = true;
            SetInputModality(InputModality::Keyboard);
            UpdateModifierStateForKey(static_cast<UINT>(wp), true, msg == WM_SYSKEYDOWN);
            UINT nativeImeModifiers = GetModifierState();
            if (msg == WM_SYSKEYDOWN && wp != VK_F10 && wp != VK_MENU)
            {
                nativeImeModifiers |= kModifierAlt;
            }
            if (_textInputBackend == TextInputBackend::Native && _nativeTextInputImeComposing && _focusedControl &&
                _focusedControl == _nativeTextInputControl && _focusedControl->SupportsTextInput() &&
                NativeImeCompositionOwnsKey(static_cast<UINT>(wp), nativeImeModifiers))
            {
                return 0;
            }
            if (wp == VK_TAB && _root)
            {
                Control* const keyTarget  = _focusedControl;
                const auto inputStartedAt = std::chrono::steady_clock::now();
                const UINT modifiers      = GetModifierState();
                if (keyTarget && keyTarget->OnKeyDown(*this, static_cast<UINT>(wp), modifiers))
                {
                    if (_textInputBackend == TextInputBackend::Native && keyTarget == _focusedControl && keyTarget->SupportsTextInput())
                    {
                        SyncNativeTextInputSession(keyTarget);
                        RecordNativeTextInputKeyToStateMetric(inputStartedAt, L"native-keydown", static_cast<uint64_t>(wp), modifiers);
                    }
                    return 0;
                }
                static_cast<void>(HandleTabNavigation(ModifiersContainShift(GetModifierState())));
                return 0;
            }
            {
                UINT modifiers = GetModifierState();
                if (msg == WM_SYSKEYDOWN && wp != VK_F10 && wp != VK_MENU)
                {
                    modifiers |= kModifierAlt;
                }
                const bool keyboardContextMenu = (wp == VK_APPS) || (wp == VK_F10 && ModifiersContainShift(modifiers) && ! ModifiersContainCtrl(modifiers));
                if (keyboardContextMenu && _focusedControl && _focusedControl->OnContextMenu(*this, true, D2D1::Point2F()))
                {
                    return 0;
                }
                Control* const keyTarget  = _focusedControl;
                const auto inputStartedAt = std::chrono::steady_clock::now();
                const bool controlHandled = keyTarget && keyTarget->OnKeyDown(*this, static_cast<UINT>(wp), modifiers);
                if (controlHandled)
                {
                    if (_textInputBackend == TextInputBackend::Native && keyTarget == _focusedControl && keyTarget->SupportsTextInput())
                    {
                        SyncNativeTextInputSession(keyTarget);
                        RecordNativeTextInputKeyToStateMetric(inputStartedAt, L"native-keydown", static_cast<uint64_t>(wp), modifiers);
                    }
                    return 0;
                }
                if (msg == WM_KEYDOWN && ! ModifiersContainAlt(modifiers) && ! ModifiersContainCtrl(modifiers))
                {
                    if (Button* const defaultButton = ResolveTreeButton(_root, _defaultButton);
                        wp == VK_RETURN && defaultButton && defaultButton->Invoke(*this, false))
                    {
                        return 0;
                    }
                    if (Button* const cancelButton = ResolveTreeButton(_root, _cancelButton);
                        wp == VK_ESCAPE && cancelButton && cancelButton->Invoke(*this, false))
                    {
                        return 0;
                    }
                    if (wp == VK_ESCAPE && _onEscape && _onEscape())
                    {
                        return 0;
                    }
                }
            }
            return 0;
        }
        case WM_KEYUP:
        case WM_SYSKEYUP:
            UpdateModifierStateForKey(static_cast<UINT>(wp), false, msg == WM_SYSKEYUP);
            {
                UINT modifiers = GetModifierState();
                if (msg == WM_SYSKEYUP && wp != VK_F10 && wp != VK_MENU)
                {
                    modifiers |= kModifierAlt;
                }
                Control* const keyTarget  = _focusedControl;
                const auto inputStartedAt = std::chrono::steady_clock::now();
                if (keyTarget && keyTarget->OnKeyUp(*this, static_cast<UINT>(wp), modifiers))
                {
                    handled = true;
                    if (_textInputBackend == TextInputBackend::Native && keyTarget == _focusedControl && keyTarget->SupportsTextInput())
                    {
                        SyncNativeTextInputSession(keyTarget);
                        RecordNativeTextInputKeyToStateMetric(inputStartedAt, L"native-keyup", static_cast<uint64_t>(wp), modifiers);
                    }
                    return 0;
                }
            }
            handled = false;
            return 0;
        case WM_GETTEXTLENGTH:
        case WM_GETTEXT:
        case WM_SETTEXT:
        case EM_GETSEL:
        case EM_SETSEL:
        case EM_REPLACESEL:
        case WM_COPY:
        case WM_CUT:
        case WM_PASTE:
        case WM_CLEAR:
        case WM_UNDO:
        {
            LRESULT result = 0;
            handled        = HandleNativeTextInputEditMessage(msg, wp, lp, result);
            if (handled)
            {
                return result;
            }
            handled = true;
            return DefWindowProcW(hwnd, msg, wp, lp);
        }
        case WM_IME_STARTCOMPOSITION:
        case WM_IME_COMPOSITION:
        case WM_IME_ENDCOMPOSITION: handled = HandleNativeTextInputImeMessage(msg, wp, lp); return 0;
        case WM_CHAR:
            handled = true;
            SetInputModality(InputModality::Keyboard);
            static_cast<void>(RouteFocusedCharInput(static_cast<wchar_t>(wp), GetModifierState(), L"native-char"));
            return 0;
        case WM_SYSCHAR:
            handled = true;
            SetInputModality(InputModality::Keyboard);
            if (HandleMnemonic(static_cast<wchar_t>(wp)))
            {
                return 0;
            }
            if (_textInputBackend == TextInputBackend::Native && _focusedControl && _focusedControl->SupportsTextInput() &&
                RouteFocusedCharInput(static_cast<wchar_t>(wp), GetModifierState() | kModifierAlt, L"native-syschar"))
            {
                return 0;
            }
            handled = false;
            return 0;
        case WM_CAPTURECHANGED:
            handled = true;
            if (_capturedControl)
            {
                _capturedControl->OnCaptureLost(*this);
            }
            ReleaseMouseCapture();
            if (_hoveredControl)
            {
                _hoveredControl->OnHoverChanged(*this, false);
                _hoveredControl->OnMouseLeave(*this);
                _hoveredControl = nullptr;
            }
            ClearTooltip();
            Invalidate();
            return 0;
        case WM_NCDESTROY:
            handled = true;
            ReleaseMouseCapture();
            Detach();
            return 0;
    }
    return 0;
}

// DWrite text formats are device-independent and do NOT need to be recreated on DPI changes.
// They are only cleared on Detach() when the WindowHost is fully torn down.
// DPI scaling is handled by the D2D device context, not by recreating text formats.
bool WindowHost::EnsureDeviceIndependentResources() const noexcept
{
    if (! _dwriteFactory)
    {
        if (! EnsureSharedWindowHostGraphicsResources())
        {
            return false;
        }
        _dwriteFactory = GetSharedWindowHostGraphicsResources().dwriteFactory;
    }
    if (! _dwriteFactory)
    {
        return false;
    }
    if (! _bodyTextFormat)
    {
        const Typography::TypographySpec spec = Typography::GetDxUiTypographySpec(FontRole::Body);
        const HRESULT hr                      = Typography::CreateTextFormat(_dwriteFactory.get(), spec, _bodyTextFormat.addressof());
        if (FAILED(hr) || ! _bodyTextFormat)
        {
            Debug::Error(L"DxUi::WindowHost: CreateTextFormat failed for body text: 0x{:08X}", hr);
            return false;
        }
    }
    if (! _bodyStrongTextFormat)
    {
        const Typography::TypographySpec spec = Typography::GetDxUiTypographySpec(FontRole::BodyStrong);
        const HRESULT hr                      = Typography::CreateTextFormat(_dwriteFactory.get(), spec, _bodyStrongTextFormat.addressof());
        if (FAILED(hr) || ! _bodyStrongTextFormat)
        {
            Debug::Error(L"DxUi::WindowHost: CreateTextFormat failed for body-strong text: 0x{:08X}", hr);
            return false;
        }
    }
    if (! _bodyLargeTextFormat)
    {
        const Typography::TypographySpec spec = Typography::GetDxUiTypographySpec(FontRole::BodyLarge);
        const HRESULT hr                      = Typography::CreateTextFormat(_dwriteFactory.get(), spec, _bodyLargeTextFormat.addressof());
        if (FAILED(hr) || ! _bodyLargeTextFormat)
        {
            Debug::Error(L"DxUi::WindowHost: CreateTextFormat failed for body-large text: 0x{:08X}", hr);
            return false;
        }
    }
    if (! _listItemTextFormat)
    {
        const Typography::TypographySpec spec = Typography::GetDxUiTypographySpec(FontRole::ListItem);
        const HRESULT hr                      = Typography::CreateTextFormat(_dwriteFactory.get(), spec, _listItemTextFormat.addressof());
        if (FAILED(hr) || ! _listItemTextFormat)
        {
            Debug::Error(L"DxUi::WindowHost: CreateTextFormat failed for list-item text: 0x{:08X}", hr);
            return false;
        }
    }
    if (! _titleTextFormat)
    {
        const Typography::TypographySpec spec = Typography::GetDxUiTypographySpec(FontRole::Title);
        const HRESULT hr                      = Typography::CreateTextFormat(_dwriteFactory.get(), spec, _titleTextFormat.addressof());
        if (FAILED(hr) || ! _titleTextFormat)
        {
            Debug::Error(L"DxUi::WindowHost: CreateTextFormat failed for title text: 0x{:08X}", hr);
            return false;
        }
    }
    if (! _subtitleTextFormat)
    {
        const Typography::TypographySpec spec = Typography::GetDxUiTypographySpec(FontRole::Subtitle);
        const HRESULT hr                      = Typography::CreateTextFormat(_dwriteFactory.get(), spec, _subtitleTextFormat.addressof());
        if (FAILED(hr) || ! _subtitleTextFormat)
        {
            Debug::Error(L"DxUi::WindowHost: CreateTextFormat failed for subtitle text: 0x{:08X}", hr);
            return false;
        }
    }
    if (! _titleLargeTextFormat)
    {
        const Typography::TypographySpec spec = Typography::GetDxUiTypographySpec(FontRole::TitleLarge);
        const HRESULT hr                      = Typography::CreateTextFormat(_dwriteFactory.get(), spec, _titleLargeTextFormat.addressof());
        if (FAILED(hr) || ! _titleLargeTextFormat)
        {
            Debug::Error(L"DxUi::WindowHost: CreateTextFormat failed for title-large text: 0x{:08X}", hr);
            return false;
        }
    }
    if (! _displayTextFormat)
    {
        const Typography::TypographySpec spec = Typography::GetDxUiTypographySpec(FontRole::Display);
        const HRESULT hr                      = Typography::CreateTextFormat(_dwriteFactory.get(), spec, _displayTextFormat.addressof());
        if (FAILED(hr) || ! _displayTextFormat)
        {
            Debug::Error(L"DxUi::WindowHost: CreateTextFormat failed for display text: 0x{:08X}", hr);
            return false;
        }
    }
    if (! _headerTextFormat)
    {
        const Typography::TypographySpec spec = Typography::GetDxUiTypographySpec(FontRole::Header);
        const HRESULT hr                      = Typography::CreateTextFormat(_dwriteFactory.get(), spec, _headerTextFormat.addressof());
        if (FAILED(hr) || ! _headerTextFormat)
        {
            Debug::Error(L"DxUi::WindowHost: CreateTextFormat failed for header text: 0x{:08X}", hr);
            return false;
        }
    }
    if (! _smallTextFormat)
    {
        const Typography::TypographySpec spec = Typography::GetDxUiTypographySpec(FontRole::Small);
        const HRESULT hr                      = Typography::CreateTextFormat(_dwriteFactory.get(), spec, _smallTextFormat.addressof());
        if (FAILED(hr) || ! _smallTextFormat)
        {
            Debug::Error(L"DxUi::WindowHost: CreateTextFormat failed for small text: 0x{:08X}", hr);
            return false;
        }
    }
    if (! _fluentIconFontAvailabilityChecked)
    {
        _fluentIconFontAvailable           = HasFontFamily(_dwriteFactory.get(), Typography::kSegoeFluentIconsFamily);
        _fluentIconFontAvailabilityChecked = true;
    }
    if (_fluentIconFontAvailable && ! _iconTextFormat)
    {
        const Typography::TypographySpec spec = Typography::GetDxUiTypographySpec(FontRole::Icon);
        const HRESULT hr                      = Typography::CreateTextFormat(_dwriteFactory.get(), spec, _iconTextFormat.addressof());
        if (FAILED(hr) || ! _iconTextFormat)
        {
            Debug::Warning(L"DxUi::WindowHost: CreateTextFormat failed for fluent icon text: 0x{:08X}", hr);
            _fluentIconFontAvailable = false;
            _iconTextFormat.reset();
        }
    }
    if (_fluentIconFontAvailable && ! _heroIconTextFormat)
    {
        const Typography::TypographySpec spec = Typography::GetDxUiTypographySpec(FontRole::HeroIcon);
        const HRESULT hr                      = Typography::CreateTextFormat(_dwriteFactory.get(), spec, _heroIconTextFormat.addressof());
        if (FAILED(hr) || ! _heroIconTextFormat)
        {
            Debug::Warning(L"DxUi::WindowHost: CreateTextFormat failed for fluent hero icon text: 0x{:08X}", hr);
            _heroIconTextFormat.reset();
        }
    }
    if (! _monoTextFormat)
    {
        const Typography::TypographySpec spec = Typography::GetDxUiTypographySpec(FontRole::Monospace);
        const HRESULT hr                      = Typography::CreateTextFormat(_dwriteFactory.get(), spec, _monoTextFormat.addressof());
        if (FAILED(hr) || ! _monoTextFormat)
        {
            Debug::Error(L"DxUi::WindowHost: CreateTextFormat failed for monospace text: 0x{:08X}", hr);
            return false;
        }
    }
    return true;
}

bool WindowHost::EnsureDeviceResources() noexcept
{
    SharedWindowHostGraphicsResources& shared = GetSharedWindowHostGraphicsResources();
    if (_sharedGraphicsGeneration != 0u && _sharedGraphicsGeneration != shared.generation)
    {
        DiscardSizeDependentResources(L"shared-generation-change");
        _dcompVisual.reset();
        _dcompTarget.reset();
        _dcompDevice.reset();
        _d2dContext.reset();
        _d2dDevice.reset();
        _d2dFactory.reset();
        _dwriteFactory.reset();
        _d3dContext.reset();
        _d3dDevice.reset();
        _sharedGraphicsGeneration = 0u;
    }

    if (_d3dDevice && _d2dContext && _d2dFactory && _sharedGraphicsGeneration == shared.generation)
    {
        return true;
    }

    if (! EnsureSharedWindowHostGraphicsResources())
    {
        return false;
    }

    _d3dDevice    = shared.d3dDevice;
    _d3dContext   = shared.d3dContext;
    _d2dFactory   = shared.d2dFactory;
    _d2dDevice    = shared.d2dDevice;
    _featureLevel = shared.featureLevel;

    if (_presentationMode == PresentationMode::CompositionSwapChain)
    {
        if (! EnsureSharedWindowHostCompositionDevice())
        {
            return false;
        }

        SharedWindowHostGraphicsResources& refreshedShared = GetSharedWindowHostGraphicsResources();
        _dcompDevice                                       = refreshedShared.dcompDevice;
        if (! _dcompDevice)
        {
            return false;
        }
    }

    const HRESULT hrContext = _d2dDevice->CreateDeviceContext(D2D1_DEVICE_CONTEXT_OPTIONS_NONE, _d2dContext.addressof());
    if (FAILED(hrContext) || ! _d2dContext)
    {
        Debug::Error(L"DxUi::WindowHost: ID2D1Device::CreateDeviceContext failed: 0x{:08X}", hrContext);
        _d2dDevice.reset();
        _d2dFactory.reset();
        _d3dContext.reset();
        _d3dDevice.reset();
        return false;
    }

    _sharedGraphicsGeneration = shared.generation;
    _d2dContext->SetUnitMode(D2D1_UNIT_MODE_DIPS);
    _d2dContext->SetDpi(static_cast<float>(_dpi), static_cast<float>(_dpi));
    RecreateBrushCache();
    return true;
}

bool WindowHost::EnsureSizeDependentResources(const bool allowHidden) noexcept
{
    if (! _hwnd || _widthPx == 0u || _heightPx == 0u)
    {
        return false;
    }

    if (! allowHidden && ! IsSwapChainActiveHostWindow(_hwnd, _widthPx, _heightPx))
    {
        // Window is invisible or zero-sized.  Skip rendering but keep
        // existing resources alive — WM_SETREDRAW FALSE temporarily
        // hides children without a matching WM_SHOWWINDOW TRUE on
        // restore.  Destroying the swap chain here would permanently
        // break rendering until a full recreation event.
        return false;
    }

    if (! EnsureDeviceIndependentResources() || ! EnsureDeviceResources())
    {
        return false;
    }

    if (! _swapChain)
    {
        SharedWindowHostGraphicsResources& shared = GetSharedWindowHostGraphicsResources();
        if (! shared.dxgiFactory)
        {
            return false;
        }

        DXGI_SWAP_CHAIN_DESC1 desc{};
        desc.Width            = std::max<UINT>(1u, _widthPx);
        desc.Height           = std::max<UINT>(1u, _heightPx);
        desc.Format           = DXGI_FORMAT_B8G8R8A8_UNORM;
        desc.SampleDesc.Count = 1;
        desc.BufferUsage      = DXGI_USAGE_RENDER_TARGET_OUTPUT;
        desc.BufferCount      = 2;
        desc.Scaling          = DXGI_SCALING_STRETCH;
        desc.SwapEffect       = DXGI_SWAP_EFFECT_FLIP_SEQUENTIAL;
        desc.AlphaMode        = _presentationMode == PresentationMode::CompositionSwapChain ? DXGI_ALPHA_MODE_PREMULTIPLIED : DXGI_ALPHA_MODE_IGNORE;

        HRESULT hrSwapChain = E_FAIL;
        if (_presentationMode == PresentationMode::CompositionSwapChain)
        {
            hrSwapChain = shared.dxgiFactory->CreateSwapChainForComposition(_d3dDevice.get(), &desc, nullptr, _swapChain.addressof());
            if (SUCCEEDED(hrSwapChain) && _swapChain)
            {
                if (! _dcompDevice)
                {
                    Debug::Error(L"DxUi::WindowHost: composition swap chain created without DirectComposition device.");
                    _swapChain.reset();
                    return false;
                }

                if (! _dcompTarget)
                {
                    const HRESULT hrTarget = _dcompDevice->CreateTargetForHwnd(_hwnd, TRUE, _dcompTarget.addressof());
                    if (FAILED(hrTarget) || ! _dcompTarget)
                    {
                        Debug::Error(L"DxUi::WindowHost: IDCompositionDesktopDevice::CreateTargetForHwnd failed: 0x{:08X}", hrTarget);
                        _swapChain.reset();
                        return false;
                    }
                }

                if (! _dcompVisual)
                {
                    const HRESULT hrVisual = _dcompDevice->CreateVisual(_dcompVisual.addressof());
                    if (FAILED(hrVisual) || ! _dcompVisual)
                    {
                        Debug::Error(L"DxUi::WindowHost: IDCompositionDesktopDevice::CreateVisual failed: 0x{:08X}", hrVisual);
                        _swapChain.reset();
                        return false;
                    }
                }

                const HRESULT hrSetContent = _dcompVisual->SetContent(_swapChain.get());
                const HRESULT hrSetRoot    = SUCCEEDED(hrSetContent) ? _dcompTarget->SetRoot(_dcompVisual.get()) : E_FAIL;
                const HRESULT hrCommit     = SUCCEEDED(hrSetRoot) ? _dcompDevice->Commit() : E_FAIL;
                if (FAILED(hrSetContent) || FAILED(hrSetRoot) || FAILED(hrCommit))
                {
                    Debug::Error(L"DxUi::WindowHost: DirectComposition visual setup failed ({:08X}, {:08X}, {:08X})", hrSetContent, hrSetRoot, hrCommit);
                    _dcompVisual.reset();
                    _dcompTarget.reset();
                    _swapChain.reset();
                    return false;
                }
            }
        }
        else
        {
            hrSwapChain = shared.dxgiFactory->CreateSwapChainForHwnd(_d3dDevice.get(), _hwnd, &desc, nullptr, nullptr, _swapChain.addressof());
        }
        if (FAILED(hrSwapChain) || ! _swapChain)
        {
            Debug::Error(L"DxUi::WindowHost: {} failed: 0x{:08X}",
                         _presentationMode == PresentationMode::CompositionSwapChain ? L"CreateSwapChainForComposition" : L"CreateSwapChainForHwnd",
                         hrSwapChain);
            return false;
        }

        _swapChainWidthPx  = desc.Width;
        _swapChainHeightPx = desc.Height;
    }

    const UINT desiredWidthPx  = std::max<UINT>(1u, _widthPx);
    const UINT desiredHeightPx = std::max<UINT>(1u, _heightPx);
    const bool sizeChanged     = _swapChainWidthPx != desiredWidthPx || _swapChainHeightPx != desiredHeightPx;
    if (! sizeChanged && _targetBitmap)
    {
        return true;
    }

    const auto createTargetBitmap = [&]() noexcept
    {
        wil::com_ptr<IDXGISurface> surface;
        const HRESULT hrBuffer = _swapChain->GetBuffer(0u, IID_PPV_ARGS(surface.addressof()));
        if (FAILED(hrBuffer) || ! surface)
        {
            Debug::Error(L"DxUi::WindowHost: IDXGISwapChain1::GetBuffer failed: 0x{:08X}", hrBuffer);
            return false;
        }
        const D2D1_BITMAP_PROPERTIES1 props = D2D1::BitmapProperties1(
            D2D1_BITMAP_OPTIONS_TARGET | D2D1_BITMAP_OPTIONS_CANNOT_DRAW,
            D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM,
                              _presentationMode == PresentationMode::CompositionSwapChain ? D2D1_ALPHA_MODE_PREMULTIPLIED : D2D1_ALPHA_MODE_IGNORE),
            static_cast<float>(_dpi),
            static_cast<float>(_dpi));
        const HRESULT hrBitmap = _d2dContext->CreateBitmapFromDxgiSurface(surface.get(), &props, _targetBitmap.addressof());
        if (FAILED(hrBitmap) || ! _targetBitmap)
        {
            Debug::Error(L"DxUi::WindowHost: CreateBitmapFromDxgiSurface failed: 0x{:08X}", hrBitmap);
            return false;
        }
        _d2dContext->SetTarget(_targetBitmap.get());
        RecreateBrushCache();
        return true;
    };

    if (sizeChanged)
    {
        Debug::Warning(L"DxUi::WindowHost: ResizeBuffers hwnd=0x{:X}  {}x{} -> {}x{}",
                       reinterpret_cast<uintptr_t>(_hwnd),
                       _swapChainWidthPx,
                       _swapChainHeightPx,
                       desiredWidthPx,
                       desiredHeightPx);
        PrepareForSwapChainResize();
        const HRESULT hrResize = _swapChain->ResizeBuffers(0u, desiredWidthPx, desiredHeightPx, DXGI_FORMAT_UNKNOWN, 0u);
        if (FAILED(hrResize))
        {
#if defined(ENABLE_TESTS)
            ++_debugResizeFailureCount;
#endif
            Debug::Warning(L"DxUi::WindowHost: ResizeBuffers failed; deferring swap-chain recreation to a later visible render: 0x{:08X}", hrResize);
            DiscardSizeDependentResources(L"resize-buffers-failed");
            return false;
        }

        _swapChainWidthPx  = desiredWidthPx;
        _swapChainHeightPx = desiredHeightPx;
#if defined(ENABLE_TESTS)
        ++_debugResizeCount;
#endif
    }

    return createTargetBitmap();
}

void WindowHost::PrepareForSwapChainResize() noexcept
{
    if (_d2dContext && _targetBitmap)
    {
        // EndDraw has already submitted D2D work.  Flushing the D2D context
        // after detaching the target leaves transient menu/preferences hosts in
        // D2DERR_WRONG_STATE; the D3D flush below releases the swap-chain buffer.
        _d2dContext->SetTarget(nullptr);
    }
    _targetBitmap.reset();
    _brushCache.clear();
    if (_d3dContext)
    {
        _d3dContext->ClearState();
        _d3dContext->Flush();
    }
}

void WindowHost::DiscardSizeDependentResources(const std::wstring_view reason) noexcept
{
    Debug::Warning(L"DxUi::WindowHost: DiscardSizeDependentResources hwnd=0x{:X} reason={} size={}x{} swap={}x{} visible={} focused={} hovered={} captured={} "
                   L"textInput={}",
                   reinterpret_cast<uintptr_t>(_hwnd),
                   reason.empty() ? L"(unspecified)" : reason,
                   _widthPx,
                   _heightPx,
                   _swapChainWidthPx,
                   _swapChainHeightPx,
                   (_hwnd && IsWindowVisible(_hwnd) != FALSE) ? L"true" : L"false",
                   static_cast<const void*>(_focusedControl),
                   static_cast<const void*>(_hoveredControl),
                   static_cast<const void*>(_capturedControl),
                   HasActiveTextInput() ? L"true" : L"false");
    PrepareForSwapChainResize();
    _swapChain.reset();
    _swapChainWidthPx  = 0u;
    _swapChainHeightPx = 0u;
}

void WindowHost::DiscardDeviceResources() noexcept
{
    if (_d2dContext || _targetBitmap || _swapChain || _d3dContext)
    {
        PrepareForSwapChainResize();
    }
    _brushCache.clear();
    _brushFailureLogKeys.clear();
    _fallbackBrush.reset();
    _targetBitmap.reset();
    _swapChain.reset();
    _swapChainWidthPx  = 0u;
    _swapChainHeightPx = 0u;
    _dcompVisual.reset();
    _dcompTarget.reset();
    _dcompDevice.reset();
    _d2dContext.reset();
    _d2dDevice.reset();
    _d2dFactory.reset();
    _d3dContext.reset();
    _d3dDevice.reset();
    _sharedGraphicsGeneration = 0u;
}

void WindowHost::RecreateBrushCache() const
{
    _brushCache.clear();
    _brushFailureLogKeys.clear();
    _fallbackBrush.reset();
    if (! _d2dContext)
    {
        return;
    }

    const HRESULT hr = _d2dContext->CreateSolidColorBrush(D2D1::ColorF(1.0f, 1.0f, 1.0f, 1.0f), _fallbackBrush.addressof());
    if (FAILED(hr) || ! _fallbackBrush)
    {
        Debug::Warning(L"DxUi::WindowHost: failed to create fallback solid brush: 0x{:08X}", hr);
    }
}

void WindowHost::Render(const RECT* dirtyRectPx, bool allowHidden) noexcept
{
#if defined(ENABLE_TESTS)
    Render(dirtyRectPx, nullptr, allowHidden);
#else
    FrameClock frameClock;
    FrameStage frameStage       = FrameStage::Idle;
    const auto frameStartedAt   = frameClock.Now();
    uint64_t updateUs           = 0u;
    uint64_t renderUs           = 0u;
    uint64_t presentUs          = 0u;
    const auto dirtyRectMetrics = ResolveWindowHostDirtyRectMetrics(dirtyRectPx, _widthPx, _heightPx);
    const auto emitFrameMetrics = wil::scope_exit([&]
    { EmitWindowHostFrameMetrics(frameClock.ElapsedUs(frameStartedAt, frameClock.Now()), updateUs, renderUs, presentUs, dirtyRectMetrics); });

    Debug::Perf::Scope paintPerf(L"DxUI::Paint");
    paintPerf.SetValue0(_widthPx);
    paintPerf.SetValue1(_heightPx);

    {
        FrameStageScope updateScope(frameStage, FrameStage::Update);
        const auto updateStartedAt = frameClock.Now();
        if (! EnsureSizeDependentResources(allowHidden))
        {
            updateUs = frameClock.ElapsedUs(updateStartedAt, frameClock.Now());
            paintPerf.SetHr(E_FAIL);
            EmitPendingNativeTextInputPaintMetric(E_FAIL);
            return;
        }
        updateUs = frameClock.ElapsedUs(updateStartedAt, frameClock.Now());
    }

    const bool isPartialDirty = dirtyRectMetrics.isPartialDirty;
    paintPerf.SetDetail(isPartialDirty ? L"partial" : L"full");

    D2D1_RECT_F clipDip{};
    if (isPartialDirty)
    {
        clipDip = D2D1::RectF(PixelsToDip(static_cast<float>(dirtyRectPx->left)),
                              PixelsToDip(static_cast<float>(dirtyRectPx->top)),
                              PixelsToDip(static_cast<float>(dirtyRectPx->right)),
                              PixelsToDip(static_cast<float>(dirtyRectPx->bottom)));
    }

    HRESULT hrDraw = S_OK;
    {
        FrameStageScope renderScope(frameStage, FrameStage::Render);
        const auto renderStartedAt = frameClock.Now();
        bool clipPushed            = false;
        _d2dContext->BeginDraw();
        auto endDrawGuard = wil::scope_exit([&]
        {
            if (clipPushed)
            {
                _d2dContext->PopAxisAlignedClip();
            }
            _d2dContext->EndDraw(nullptr, nullptr);
        });
        _d2dContext->SetTransform(D2D1::Matrix3x2F::Identity());

        if (isPartialDirty)
        {
            _d2dContext->PushAxisAlignedClip(clipDip, D2D1_ANTIALIAS_MODE_ALIASED);
            clipPushed = true;
        }

        _d2dContext->Clear(_presentationMode == PresentationMode::CompositionSwapChain ? D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f) : _palette.windowBackground);
        if (_root)
        {
            _root->Paint(*this);
        }
        if (_smokeOverlayVisible)
        {
            if (auto* brush = GetSolidBrush(_palette.smokeOverlay))
            {
                _d2dContext->FillRectangle(GetClientBoundsDip(), brush);
            }
        }
        if (_root)
        {
            _root->PaintOverlay(*this);
        }
        if (_tooltipLayer.HasTooltip())
        {
            _tooltipLayer.Paint(*this);
        }

        if (clipPushed)
        {
            _d2dContext->PopAxisAlignedClip();
            clipPushed = false;
        }

        endDrawGuard.release();
        hrDraw   = _d2dContext->EndDraw();
        renderUs = frameClock.ElapsedUs(renderStartedAt, frameClock.Now());
    }

    HRESULT hrPresent = S_OK;
    {
        FrameStageScope presentScope(frameStage, FrameStage::Present);
        const auto presentStartedAt = frameClock.Now();
        if (SUCCEEDED(hrDraw))
        {
            if (isPartialDirty)
            {
                RECT dirtyRect = *dirtyRectPx;
                DXGI_PRESENT_PARAMETERS params{};
                params.DirtyRectsCount = 1;
                params.pDirtyRects     = &dirtyRect;
                hrPresent              = _swapChain->Present1(1u, 0u, &params);
            }
            else
            {
                hrPresent = _swapChain->Present(1u, 0u);
            }
        }
        presentUs = frameClock.ElapsedUs(presentStartedAt, frameClock.Now());
    }

    if (hrDraw == D2DERR_RECREATE_TARGET || hrPresent == DXGI_ERROR_DEVICE_REMOVED || hrPresent == DXGI_ERROR_DEVICE_RESET)
    {
        const HRESULT renderHr = FAILED(hrDraw) ? hrDraw : hrPresent;
        paintPerf.SetHr(renderHr);
        EmitPendingNativeTextInputPaintMetric(renderHr);
        DiscardSizeDependentResources(L"render-device-lost");
        DiscardDeviceResources();
        ResetSharedWindowHostGraphicsResources();
        return;
    }
    const HRESULT renderHr = FAILED(hrDraw) ? hrDraw : hrPresent;
    paintPerf.SetHr(renderHr);
    EmitPendingNativeTextInputPaintMetric(renderHr);
    if (FAILED(hrDraw))
    {
        Debug::Error(L"DxUi::WindowHost: EndDraw failed: 0x{:08X}", hrDraw);
    }
    if (FAILED(hrPresent))
    {
#if defined(ENABLE_TESTS)
        ++_debugPresentFailureCount;
#endif
        Debug::Error(L"DxUi::WindowHost: Present failed: 0x{:08X}", hrPresent);
    }
#endif
}

#if defined(ENABLE_TESTS)
bool WindowHost::CaptureCurrentBackBuffer(WindowHostBitmapCapture& out) noexcept
{
    if (! _swapChain || ! _d3dDevice || ! _d3dContext)
    {
        return false;
    }

    wil::com_ptr<ID3D11Texture2D> backBuffer;
    if (FAILED(_swapChain->GetBuffer(0, IID_PPV_ARGS(backBuffer.addressof()))) || ! backBuffer)
    {
        return false;
    }

    D3D11_TEXTURE2D_DESC desc{};
    backBuffer->GetDesc(&desc);
    desc.BindFlags      = 0u;
    desc.MiscFlags      = 0u;
    desc.Usage          = D3D11_USAGE_STAGING;
    desc.CPUAccessFlags = D3D11_CPU_ACCESS_READ;

    wil::com_ptr<ID3D11Texture2D> staging;
    if (FAILED(_d3dDevice->CreateTexture2D(&desc, nullptr, staging.addressof())) || ! staging)
    {
        return false;
    }

    _d3dContext->CopyResource(staging.get(), backBuffer.get());

    D3D11_MAPPED_SUBRESOURCE mapped{};
    if (FAILED(_d3dContext->Map(staging.get(), 0u, D3D11_MAP_READ, 0u, &mapped)))
    {
        return false;
    }
    const auto unmapGuard = wil::scope_exit([&] { _d3dContext->Unmap(staging.get(), 0u); });

    out.widthPx  = desc.Width;
    out.heightPx = desc.Height;
    out.bgraPixels.resize(static_cast<size_t>(desc.Width) * static_cast<size_t>(desc.Height) * 4u);
    for (UINT rowIndex = 0u; rowIndex < desc.Height; ++rowIndex)
    {
        const auto* const sourceRow = static_cast<const uint8_t*>(mapped.pData) + (static_cast<size_t>(mapped.RowPitch) * rowIndex);
        auto* const targetRow       = out.bgraPixels.data() + (static_cast<size_t>(rowIndex) * static_cast<size_t>(desc.Width) * 4u);
        std::copy_n(sourceRow, static_cast<size_t>(desc.Width) * 4u, targetRow);
    }

    return true;
}

void WindowHost::Render(const RECT* dirtyRectPx, WindowHostBitmapCapture* capture, bool allowHidden) noexcept
{
    FrameClock frameClock;
    FrameStage frameStage       = FrameStage::Idle;
    const auto frameStartedAt   = frameClock.Now();
    uint64_t updateUs           = 0u;
    uint64_t renderUs           = 0u;
    uint64_t presentUs          = 0u;
    const auto dirtyRectMetrics = ResolveWindowHostDirtyRectMetrics(dirtyRectPx, _widthPx, _heightPx);
    const auto emitFrameMetrics = wil::scope_exit([&]
    { EmitWindowHostFrameMetrics(frameClock.ElapsedUs(frameStartedAt, frameClock.Now()), updateUs, renderUs, presentUs, dirtyRectMetrics); });

    Debug::Perf::Scope paintPerf(L"DxUI::Paint");
    paintPerf.SetValue0(_widthPx);
    paintPerf.SetValue1(_heightPx);

    {
        FrameStageScope updateScope(frameStage, FrameStage::Update);
        const auto updateStartedAt = frameClock.Now();
        if (! EnsureSizeDependentResources(allowHidden))
        {
            updateUs = frameClock.ElapsedUs(updateStartedAt, frameClock.Now());
            paintPerf.SetHr(E_FAIL);
            EmitPendingNativeTextInputPaintMetric(E_FAIL);
            return;
        }
        updateUs = frameClock.ElapsedUs(updateStartedAt, frameClock.Now());
    }

#if defined(ENABLE_TESTS)
    ++_debugRenderCount;
#endif

    // Determine whether rcPaint describes a partial dirty region.
    // FLIP_SEQUENTIAL preserves back buffer content between frames, so
    // clipping to the dirty rect and using Present1 with dirty-rect params
    // is safe: non-dirty regions retain previously-presented content.
    const bool isPartialDirty = dirtyRectMetrics.isPartialDirty;
    paintPerf.SetDetail(isPartialDirty ? L"partial" : L"full");

    D2D1_RECT_F clipDip{};
    if (isPartialDirty)
    {
        clipDip = D2D1::RectF(PixelsToDip(static_cast<float>(dirtyRectPx->left)),
                              PixelsToDip(static_cast<float>(dirtyRectPx->top)),
                              PixelsToDip(static_cast<float>(dirtyRectPx->right)),
                              PixelsToDip(static_cast<float>(dirtyRectPx->bottom)));
    }

    HRESULT hrDraw    = S_OK;
    HRESULT hrCapture = S_OK;
    {
        FrameStageScope renderScope(frameStage, FrameStage::Render);
        const auto renderStartedAt = frameClock.Now();
        bool clipPushed            = false;
        _d2dContext->BeginDraw();
        auto endDrawGuard = wil::scope_exit([&]
        {
            if (clipPushed)
            {
                _d2dContext->PopAxisAlignedClip();
            }
            _d2dContext->EndDraw(nullptr, nullptr);
        });
        _d2dContext->SetTransform(D2D1::Matrix3x2F::Identity());

        if (isPartialDirty)
        {
            _d2dContext->PushAxisAlignedClip(clipDip, D2D1_ANTIALIAS_MODE_ALIASED);
            clipPushed = true;
        }

        _d2dContext->Clear(_presentationMode == PresentationMode::CompositionSwapChain ? D2D1::ColorF(0.0f, 0.0f, 0.0f, 0.0f) : _palette.windowBackground);
        if (_root)
        {
            _root->Paint(*this);
        }
        if (_smokeOverlayVisible)
        {
            if (auto* brush = GetSolidBrush(_palette.smokeOverlay))
            {
                _d2dContext->FillRectangle(GetClientBoundsDip(), brush);
            }
        }
        if (_root)
        {
            _root->PaintOverlay(*this);
        }
        if (_tooltipLayer.HasTooltip())
        {
            _tooltipLayer.Paint(*this);
        }

        if (clipPushed)
        {
            _d2dContext->PopAxisAlignedClip();
            clipPushed = false;
        }

        endDrawGuard.release();
        hrDraw = _d2dContext->EndDraw();

        if (capture)
        {
            if (! CaptureCurrentBackBuffer(*capture))
            {
                hrCapture = E_FAIL;
            }
        }
        renderUs = frameClock.ElapsedUs(renderStartedAt, frameClock.Now());
    }

    HRESULT hrPresent = S_OK;
    {
        FrameStageScope presentScope(frameStage, FrameStage::Present);
        const auto presentStartedAt = frameClock.Now();
        if (SUCCEEDED(hrDraw))
        {
            if (isPartialDirty)
            {
                // Tell the compositor only the dirty region changed, so it can
                // skip recompositing the rest of the window.
                RECT dirtyRect = *dirtyRectPx;
                DXGI_PRESENT_PARAMETERS params{};
                params.DirtyRectsCount = 1;
                params.pDirtyRects     = &dirtyRect;
                hrPresent              = _swapChain->Present1(1u, 0u, &params);
            }
            else
            {
                hrPresent = _swapChain->Present(1u, 0u);
            }
        }
        presentUs = frameClock.ElapsedUs(presentStartedAt, frameClock.Now());
    }

    if (hrDraw == D2DERR_RECREATE_TARGET || hrPresent == DXGI_ERROR_DEVICE_REMOVED || hrPresent == DXGI_ERROR_DEVICE_RESET)
    {
        const HRESULT renderHr = FAILED(hrDraw) ? hrDraw : hrPresent;
        paintPerf.SetHr(renderHr);
        EmitPendingNativeTextInputPaintMetric(renderHr);
        DiscardSizeDependentResources(L"render-device-lost");
        DiscardDeviceResources();
        ResetSharedWindowHostGraphicsResources();
        return;
    }
    const HRESULT renderHr = FAILED(hrDraw) ? hrDraw : (FAILED(hrCapture) ? hrCapture : hrPresent);
    paintPerf.SetHr(renderHr);
    EmitPendingNativeTextInputPaintMetric(renderHr);
    if (FAILED(hrDraw))
    {
        Debug::Error(L"DxUi::WindowHost: EndDraw failed: 0x{:08X}", hrDraw);
    }
    if (FAILED(hrCapture))
    {
        Debug::Error(L"DxUi::WindowHost: Debug capture failed.");
    }
    if (FAILED(hrPresent))
    {
#if defined(ENABLE_TESTS)
        ++_debugPresentFailureCount;
#endif
        Debug::Error(L"DxUi::WindowHost: Present failed: 0x{:08X}", hrPresent);
    }
}
#endif

void WindowHost::OnSize(UINT widthPx, UINT heightPx) noexcept
{
    _widthPx  = widthPx;
    _heightPx = heightPx;
    if (_root)
    {
        _root->SetBounds(D2D1::RectF(0.0f, 0.0f, PixelsToDip(static_cast<float>(_widthPx)), PixelsToDip(static_cast<float>(_heightPx))));
    }
    if (_widthPx == 0u || _heightPx == 0u)
    {
        if (IsInteractionDiagnosticsEnabled(_hwnd))
        {
            Debug::Info(L"DxUi::WindowHost: on-size-zero hwnd={:#x} size={}x{} focus={} textInput={} visible={}",
                        reinterpret_cast<uintptr_t>(_hwnd),
                        _widthPx,
                        _heightPx,
                        static_cast<const void*>(_focusedControl),
                        HasActiveTextInput() ? L"true" : L"false",
                        (_hwnd && IsWindowVisible(_hwnd) != FALSE) ? L"true" : L"false");
        }
        DiscardSizeDependentResources(L"on-size-zero");
    }
    if (IsSwapChainActiveHostWindow(_hwnd, _widthPx, _heightPx))
    {
        Invalidate();
    }
}

void WindowHost::OnSetFocus() noexcept
{
    PruneStaleInteractionState();
    if (IsInteractionDiagnosticsEnabled(_hwnd))
    {
        Debug::Info(L"DxUi::WindowHost: on-set-focus hwnd={:#x} focus={} hovered={} captured={} textInput={}",
                    reinterpret_cast<uintptr_t>(_hwnd),
                    static_cast<const void*>(_focusedControl),
                    static_cast<const void*>(_hoveredControl),
                    static_cast<const void*>(_capturedControl),
                    HasActiveTextInput() ? L"true" : L"false");
    }
    if (_focusedControl)
    {
        bool restoredFocus = false;
        if (! _focusedControl->HasFocus())
        {
            _focusedControl->OnFocusChanged(*this, true);
            restoredFocus = true;
        }
        if (_focusedControl->SupportsTextInput())
        {
            ActivateTextInput(_focusedControl);
        }
        if (restoredFocus)
        {
            Invalidate();
        }
        return;
    }

    if (! _focusedControl && _root)
    {
        SetFocusControl(FindAdjacentFocusable(_root.get(), nullptr, false).control);
    }
}

void WindowHost::OnKillFocus(bool clearRetainedFocus) noexcept
{
    DeactivateTextInput(false);
    PruneStaleInteractionState();
    _modifierState = 0u;
    if (_focusedControl && ControlBelongsToTree(_root.get(), _focusedControl))
    {
        if (_focusedControl->HasFocus())
        {
            _focusedControl->OnFocusChanged(*this, false);
        }
    }
    if (clearRetainedFocus)
    {
        _focusedControl = nullptr;
    }
    if (IsInteractionDiagnosticsEnabled(_hwnd))
    {
        Debug::Info(L"DxUi::WindowHost: on-kill-focus hwnd={:#x} focus={} hovered={} captured={} textInput={}",
                    reinterpret_cast<uintptr_t>(_hwnd),
                    static_cast<const void*>(_focusedControl),
                    static_cast<const void*>(_hoveredControl),
                    static_cast<const void*>(_capturedControl),
                    HasActiveTextInput() ? L"true" : L"false");
    }
    Invalidate();
}

void WindowHost::PruneStaleInteractionState() noexcept
{
    const auto classifyInteractionState = [this](const Control* control) noexcept
    {
        if (! control || ! _root)
        {
            return ControlInteractionState{};
        }
        return ResolveControlInteractionState(_root.get(), control);
    };

    bool prunedCapture         = false;
    bool prunedHover           = false;
    bool prunedFocus           = false;
    bool prunedNativeTextInput = false;

    if (_capturedControl && ! classifyInteractionState(_capturedControl).effectivelyInteractive)
    {
        _capturedControl = nullptr;
        prunedCapture    = true;
        if (_hwnd && GetCapture() == _hwnd)
        {
            ReleaseCapture();
        }
    }

    const ControlInteractionState hoveredState = classifyInteractionState(_hoveredControl);
    if (_hoveredControl && ! hoveredState.effectivelyInteractive)
    {
        if (hoveredState.inTree)
        {
            _hoveredControl->OnHoverChanged(*this, false);
            _hoveredControl->OnMouseLeave(*this);
        }
        _hoveredControl = nullptr;
        prunedHover     = true;
        ClearTooltip();
    }

    const ControlInteractionState focusedState = classifyInteractionState(_focusedControl);
    if (_focusedControl && ! focusedState.effectivelyInteractive)
    {
        if (focusedState.inTree)
        {
            _focusedControl->OnFocusChanged(*this, false);
        }
        _focusedControl = nullptr;
        prunedFocus     = true;
    }

    if (_nativeTextInputControl && ! classifyInteractionState(_nativeTextInputControl).effectivelyInteractive)
    {
        DeactivateNativeTextInputSession(false);
        prunedNativeTextInput = true;
    }

    if (_pendingPointerDoubleClick.target && ! classifyInteractionState(_pendingPointerDoubleClick.target).effectivelyInteractive)
    {
        ClearPendingPointerDoubleClick();
    }

    if ((prunedCapture || prunedHover || prunedFocus || prunedNativeTextInput) && IsInteractionDiagnosticsEnabled(_hwnd))
    {
        Debug::Info(L"DxUi::WindowHost: prune hwnd={:#x} capture={} hover={} focus={} nativeTextInput={} remainingFocus={} remainingHover={} "
                    L"remainingCapture={} activeTextInput={}",
                    reinterpret_cast<uintptr_t>(_hwnd),
                    prunedCapture ? L"true" : L"false",
                    prunedHover ? L"true" : L"false",
                    prunedFocus ? L"true" : L"false",
                    prunedNativeTextInput ? L"true" : L"false",
                    static_cast<const void*>(_focusedControl),
                    static_cast<const void*>(_hoveredControl),
                    static_cast<const void*>(_capturedControl),
                    HasActiveTextInput() ? L"true" : L"false");
    }
}

void WindowHost::ResetRootInteractionState() noexcept
{
    ClearPendingPointerDoubleClick();
    ReleaseMouseCapture();
    DeactivateTextInput(true);
    _modifierState = 0u;
    if (_hoveredControl && ControlBelongsToTree(_root.get(), _hoveredControl))
    {
        _hoveredControl->OnHoverChanged(*this, false);
        _hoveredControl->OnMouseLeave(*this);
    }
    _hoveredControl = nullptr;
    if (_focusedControl && ControlBelongsToTree(_root.get(), _focusedControl))
    {
        _focusedControl->OnFocusChanged(*this, false);
    }
    _focusedControl = nullptr;
    ClearTooltip();
}

Control* WindowHost::HitTestControl(D2D1_POINT_2F pointDip) noexcept
{
    if (! _root)
    {
        return nullptr;
    }

    if (Control* overlayTarget = _root->HitTestOverlay(pointDip))
    {
        return overlayTarget;
    }

    return _root->HitTest(pointDip);
}

WindowHostCursorKind WindowHost::ResolveCursorKindForPoint(D2D1_POINT_2F pointDip) noexcept
{
    Control* target = _capturedControl;
    if (! target)
    {
        target = HitTestControl(pointDip);
    }
    return target ? target->ResolveCursorKind(*this, pointDip) : WindowHostCursorKind::Default;
}

HCURSOR WindowHost::ResolveCursorHandle(WindowHostCursorKind cursorKind) const noexcept
{
    static const HCURSOR arrowCursor      = LoadCursorW(nullptr, IDC_ARROW);
    static const HCURSOR horizontalCursor = LoadCursorW(nullptr, IDC_SIZEWE);

    switch (cursorKind)
    {
        case WindowHostCursorKind::HorizontalResize: return horizontalCursor ? horizontalCursor : arrowCursor;
        case WindowHostCursorKind::Default:
        default: return arrowCursor;
    }
}

void WindowHost::UpdateHover(D2D1_POINT_2F pointDip, UINT modifiers) noexcept
{
    PruneStaleInteractionState();
    Control* target = _capturedControl;
    if (! target)
    {
        target = HitTestControl(pointDip);
    }
    bool hoverChanged = false;
    if (_hoveredControl != target)
    {
        if (_hoveredControl)
        {
            _hoveredControl->OnHoverChanged(*this, false);
            _hoveredControl->OnMouseLeave(*this);
        }
        _hoveredControl = target;
        if (_hoveredControl)
        {
            _hoveredControl->OnHoverChanged(*this, true);
        }
        hoverChanged = true;
    }
    if (_hoveredControl)
    {
        _hoveredControl->OnMouseMove(*this, pointDip, modifiers);
    }
    if (hoverChanged)
    {
        Invalidate();
    }
}

void WindowHost::SetInputModality(InputModality modality) noexcept
{
    if (_inputModality == modality)
    {
        return;
    }

    _inputModality = modality;
    if (_focusedControl)
    {
        Invalidate();
    }
}

void WindowHost::UpdateModifierStateForKey(UINT virtualKey, bool keyDown, bool /*systemKey*/) noexcept
{
    UINT bit = 0u;
    switch (virtualKey)
    {
        case VK_SHIFT:
        case VK_LSHIFT:
        case VK_RSHIFT: bit = MK_SHIFT; break;
        case VK_CONTROL:
        case VK_LCONTROL:
        case VK_RCONTROL: bit = MK_CONTROL; break;
        case VK_MENU:
        case VK_LMENU:
        case VK_RMENU: bit = kModifierAlt; break;
        default: return;
    }

    if (keyDown)
    {
        _modifierState |= bit;
    }
    else
    {
        _modifierState &= ~bit;
    }
}

D2D1_POINT_2F WindowHost::PointFromLParam(LPARAM lp) const noexcept
{
    return D2D1::Point2F(PixelsToDip(static_cast<float>(GET_X_LPARAM(lp))), PixelsToDip(static_cast<float>(GET_Y_LPARAM(lp))));
}

UINT WindowHost::GetModifierState() const noexcept
{
    return _modifierState;
}

bool WindowHost::AnimationTickThunk(void* context, uint64_t nowTickMs) noexcept
{
    auto* self = static_cast<WindowHost*>(context);
    return (self && IsAttachedWindowHostRegistered(self)) ? self->OnAnimationTick(nowTickMs) : false;
}

bool WindowHost::OnAnimationTick(uint64_t nowTickMs) noexcept
{
    _lastAnimationTickMs = nowTickMs;
    if (_hwnd && ! IsHostWindowEffectivelyVisible(_hwnd))
    {
        return true;
    }

    if (! _root)
    {
        const bool tooltipTicking = _tooltipLayer.Tick(*this, nowTickMs);
        if (! tooltipTicking)
        {
            _animationSubscriptionId = 0u;
            return false;
        }
        Invalidate();
        return true;
    }

    const bool rootTicking    = _root->Tick(*this, nowTickMs);
    const bool tooltipTicking = _tooltipLayer.Tick(*this, nowTickMs);
    const bool keepTicking    = rootTicking || tooltipTicking;
    if (! keepTicking)
    {
        _animationSubscriptionId = 0u;
        return false;
    }
    Invalidate();
    return true;
}
} // namespace RedSalamander::DxUi
