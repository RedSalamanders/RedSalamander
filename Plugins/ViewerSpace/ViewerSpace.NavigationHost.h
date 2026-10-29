#pragma once

#include "ViewerSpace.Hosts.h"

#include <array>
#include <cstdint>
#include <functional>
#include <memory>
#include <optional>
#include <span>
#include <string>
#include <vector>

#include <DxUi/DxUi.h>

namespace ViewerSpaceNavigation
{
// Product-private because the app NavigationView resolves filesystem path strings, while ViewerSpace must
// navigate display-generation-qualified virtual model tokens. Standard DxUi controls keep targets keyboard/UIA reachable.
// Calls and callbacks stay on the owning UI thread; Detach clears callbacks before destroying the control tree.
class NavigationHost final
{
public:
    struct Captions final
    {
        std::wstring back;
        std::wstring forward;

        std::wstring path;
        std::wstring ancestors;
        std::wstring folders;
        std::wstring mode;
        std::wstring explorer;
        std::wstring analyst;
        std::wstring hidePanel;
        std::wstring showPanel;
        std::wstring mapOnly;
    };

    struct Crumb final
    {
        uint64_t id = 0u;
        std::wstring text;
    };

    struct Callbacks final
    {
        std::function<void(uint64_t token)> navigate;
        std::function<void()> back;
        std::function<void()> forward;

        std::function<std::vector<Crumb>(uint64_t token)> folders;
        std::function<void()> toggleAnalysis;
        std::function<void(bool analystMode)> modeChanged;
        std::function<bool(bool reverse)> focusBoundary;
    };

    NavigationHost() noexcept;
    ~NavigationHost() noexcept;

    NavigationHost(const NavigationHost&)            = delete;
    NavigationHost(NavigationHost&&)                 = delete;
    NavigationHost& operator=(const NavigationHost&) = delete;
    NavigationHost& operator=(NavigationHost&&)      = delete;

    [[nodiscard]] HRESULT Attach(HWND parent, const Captions& captions, Callbacks callbacks) noexcept;
    void Detach() noexcept;
    void SetBounds(const RECT& boundsPx) noexcept;
    void ApplyTheme(const DxUi::ThemePalette& palette) noexcept;
    void SetState(std::span<const Crumb> crumbs, bool canBack, bool canForward);
    void SetAnalysisVisible(bool visible) noexcept;
    void SetMode(bool analystMode) noexcept;
    [[nodiscard]] bool Focus(bool reverse) noexcept;
    [[nodiscard]] HWND GetHwnd() const noexcept;
    [[nodiscard]] const DxUi::ThemePalette& GetTheme() const noexcept
    {
        return _host.GetTheme();
    }
    [[nodiscard]] ViewerSpaceHosts::HostDiagnostics GetDiagnostics() const noexcept
    {
        return ViewerSpaceHosts::CaptureHostDiagnostics(_host, _hwnd.get());
    }

private:
    static constexpr size_t kMaxCrumbs           = 512u;
    static constexpr size_t kVisibleCrumbButtons = 6u;

    static LRESULT CALLBACK WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;
    [[nodiscard]] LRESULT HandleWindowMessage(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) noexcept;
    [[nodiscard]] bool RegisterChildClass() noexcept;
    void BuildControls();
    void Layout() noexcept;
    void RebuildCrumbs();
    void InvokeNavigate(uint64_t token) noexcept;
    void InvokeBack() noexcept;
    void InvokeForward() noexcept;
    void InvokeModeChanged(bool analystMode) noexcept;

    void ShowFolders(size_t slot) noexcept;
    void ShowModeMenu() noexcept;
    void ShowAncestorMenu() noexcept;
    [[nodiscard]] std::optional<uint64_t> ShowChoiceMenu(DxUi::Control& anchor,
                                                         std::span<const Crumb> choices,
                                                         std::optional<uint64_t> selected = std::nullopt) noexcept;
    void InvokeFocusBoundary(bool reverse, bool& handled) noexcept;
    void LogCallbackFailureOnce() noexcept;

    DxUi::WindowHost _host;
    wil::unique_hwnd _hwnd;
    std::wstring _windowClassName;
    ATOM _windowClassAtom = 0u;
    Captions _captions;
    Callbacks _callbacks;
    std::unique_ptr<DxUi::Panel> _root;
    DxUi::Button* _backButton    = nullptr;
    DxUi::Button* _forwardButton = nullptr;

    DxUi::Button* _analysisButton  = nullptr;
    DxUi::ComboBox* _modeCombo     = nullptr;
    DxUi::ComboBox* _ancestorCombo = nullptr;
    std::array<DxUi::Button*, kVisibleCrumbButtons> _crumbButtons{};
    std::array<DxUi::Button*, kVisibleCrumbButtons> _folderButtons{};
    std::array<uint64_t, kVisibleCrumbButtons> _visibleCrumbTokens{};
    std::vector<Crumb> _crumbs;
    std::vector<uint64_t> _ancestorTokens;
    RECT _boundsPx{};
    bool _canBack         = false;
    bool _canForward      = false;
    bool _analysisVisible = true;
    bool _analystMode     = false;

    bool _callbackFailureLogged = false;
};
} // namespace ViewerSpaceNavigation
