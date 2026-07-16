#include "Framework.h"

#include "ManagePluginsDialog.h"

#include <algorithm>
#include <array>
#include <atomic>
#include <cerrno>
#include <cwchar>
#include <filesystem>
#include <format>
#include <limits>
#include <new>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "DxUiThemePalette.h"
#include "FileSystemPluginManager.h"
#include "Helpers.h"
#include "HostServices.h"
#include "PluginConfiguration.h"
#include "SelfTestCommon.h"
#include "SettingsHotReload.h"
#include "ThemedInputFrames.h"
#include "UiMetrics.h"
#include "ViewerPluginManager.h"
#include "WindowMessages.h"
#include "WindowSizing.h"
#include "resource.h"

#pragma warning(push)
// WIL: C4625 (copy ctor deleted), C4626 (copy assign deleted), C5026 (move ctor deleted), C5027 (move assign deleted)
#pragma warning(disable : 4625 4626 5026 5027 28182)
#include <wil/resource.h>
#include <wil/win32_helpers.h>
#pragma warning(pop)

#include "DxUi/DxUi.Typography.h"
#include "DxUi/DxUi.h"
#include <uxtheme.h>

namespace
{
using RedSalamander::DxUi::Button;
using RedSalamander::DxUi::Checkbox;
using RedSalamander::DxUi::ComboBox;
using RedSalamander::DxUi::ComboBoxVariant;
using RedSalamander::DxUi::FontRole;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::TextField;
using RedSalamander::DxUi::ThemePalette;
using RedSalamander::DxUi::Toggle;
using RedSalamander::DxUi::WindowHost;
using UiMetrics::GetControlSurfaceColor;
using UiMetrics::ScaleDip;
namespace Typography = RedSalamander::DxUi::Typography;

constexpr wchar_t kPluginConfigDxButtonHostOriginalWndProcProp[] = L"RS.PluginConfigDxButtonHostOriginalWndProc";
constexpr wchar_t kPluginConfigDxButtonHostStateProp[]           = L"RS.PluginConfigDxButtonHostState";
constexpr wchar_t kPluginConfigDxFieldHostOriginalWndProcProp[]  = L"RS.PluginConfigDxFieldHostOriginalWndProc";
constexpr wchar_t kPluginConfigDxFieldHostStateProp[]            = L"RS.PluginConfigDxFieldHostState";
constexpr wchar_t kPluginConfigPanelOriginalWndProcProp[]        = L"RS.PluginConfigPanelOriginalWndProc";
constexpr wchar_t kPluginConfigPanelStateProp[]                  = L"RS.PluginConfigPanelState";
std::atomic<HWND> g_pluginConfigurationDialogWindow{nullptr};
// Keep the command row and the full visible schema-driven form surface
// on the shared DX path in the stabilized mixed-window state.
constexpr bool kEnableManagePluginsDxCommandButtons = true;
// The current validated state keeps both the visible schema-form statics
// and interactive inputs on the shared DX path.
constexpr bool kEnableManagePluginsDxFormStatics = true;
constexpr bool kEnableManagePluginsDxFormInputs  = true;

#ifdef ENABLE_TESTS
void TracePluginConfigDebug(std::wstring_view message)
{
    SelfTest::AppendSuiteTrace(SelfTest::SelfTestSuite::Commands, message);
}
#endif

enum class DxCommandButtonIndex : size_t
{
    Ok,
    Cancel,
    Count,
};

#ifdef ENABLE_TESTS
constexpr size_t kInvalidPluginConfigDebugFocusIndex = std::numeric_limits<size_t>::max();

enum class PluginConfigDebugCommand : WPARAM
{
    GetSnapshot = 1,
    ScrollByWheelDetents,
    FocusFirstInput,
    GetFirstVisibleToggleRect,
    GetVisibleToggleRectByLabel,
    GetFocusedHost,
    AdvanceTab,
    Cancel,
};

struct PluginConfigurationDialogDebugHostRect
{
    HWND host = nullptr;
    RECT rect{};
};

struct PluginConfigurationDialogDebugHost
{
    HWND host = nullptr;
};

struct PluginConfigurationDialogDebugLabeledHostRect
{
    const wchar_t* label = nullptr;
    HWND host            = nullptr;
    RECT rect{};
};
#endif

struct DxCommandButtonHost;
LRESULT CALLBACK PluginConfigDxButtonHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;
struct PluginConfigDxHostSlot;
LRESULT CALLBACK PluginConfigDxFieldHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept;

[[nodiscard]] WNDPROC GetStoredWndProc(HWND hwnd, const wchar_t* propName) noexcept
{
    return RedSalamander::Win32Callback::GetStoredWndProc(hwnd, propName);
}

[[nodiscard]] bool InstallWndProcHook(HWND hwnd, const wchar_t* originalWndProcProp, WNDPROC hookWndProc) noexcept
{
    if (! hwnd || ! originalWndProcProp || ! hookWndProc)
    {
        return false;
    }

    if (GetStoredWndProc(hwnd, originalWndProcProp))
    {
        return true;
    }

    const auto originalWndProc = reinterpret_cast<WNDPROC>(GetWindowLongPtrW(hwnd, GWLP_WNDPROC));
    if (! originalWndProc)
    {
        return false;
    }

    if (! RedSalamander::Win32Callback::SetPropNoThrow(hwnd, originalWndProcProp, reinterpret_cast<HANDLE>(originalWndProc)))
    {
        return false;
    }

    const auto previousWndProc =
        reinterpret_cast<WNDPROC>(RedSalamander::Win32Callback::SetWindowLongPtrNoThrow(hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(hookWndProc)));
    if (previousWndProc != originalWndProc)
    {
        RemovePropW(hwnd, originalWndProcProp);
        if (previousWndProc)
        {
            static_cast<void>(RedSalamander::Win32Callback::SetWindowLongPtrNoThrow(hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(previousWndProc)));
        }
        return false;
    }

    return true;
}

void RestoreWndProcHook(HWND hwnd, const wchar_t* originalWndProcProp) noexcept
{
    if (! hwnd || ! originalWndProcProp)
    {
        return;
    }

    if (const auto originalWndProc = GetStoredWndProc(hwnd, originalWndProcProp))
    {
        RemovePropW(hwnd, originalWndProcProp);
        static_cast<void>(RedSalamander::Win32Callback::SetWindowLongPtrNoThrow(hwnd, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(originalWndProc)));
    }
}

[[nodiscard]] LRESULT CallStoredWndProc(HWND hwnd, const wchar_t* originalWndProcProp, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    if (const auto originalWndProc = GetStoredWndProc(hwnd, originalWndProcProp))
    {
        return RedSalamander::Win32Callback::CallWindowProcNoThrow(originalWndProc, hwnd, msg, wp, lp);
    }

    return DefWindowProcW(hwnd, msg, wp, lp);
}

void ShowDialogAlert(HWND dlg, HostAlertSeverity severity, const std::wstring& title, const std::wstring& message) noexcept
{
    if (! dlg || message.empty())
    {
        return;
    }

    HostAlertRequest request{};
    request.version      = 1;
    request.sizeBytes    = sizeof(request);
    request.scope        = HOST_ALERT_SCOPE_WINDOW;
    request.modality     = HOST_ALERT_MODELESS;
    request.severity     = severity;
    request.targetWindow = dlg;
    request.title        = title.empty() ? nullptr : title.c_str();
    request.message      = message.c_str();
    request.closable     = TRUE;

    static_cast<void>(HostShowAlert(request));
}

using PluginConfigFieldType = Common::PluginConfiguration::FieldType;
using PluginConfigChoice    = Common::PluginConfiguration::Choice;
using PluginConfigField     = Common::PluginConfiguration::Field;

struct PluginConfigDxHostSlot
{
    PluginConfigDxHostSlot() = default;
    ~PluginConfigDxHostSlot() noexcept
    {
        Detach();
    }
    PluginConfigDxHostSlot(const PluginConfigDxHostSlot&)            = delete;
    PluginConfigDxHostSlot& operator=(const PluginConfigDxHostSlot&) = delete;
    PluginConfigDxHostSlot(PluginConfigDxHostSlot&&)                 = default;
    PluginConfigDxHostSlot& operator=(PluginConfigDxHostSlot&&)      = default;

    wil::unique_hwnd hostHwnd;
    std::unique_ptr<WindowHost> host;

    void Detach() noexcept
    {
        if (hostHwnd)
        {
            if (IsWindow(hostHwnd.get()) != FALSE)
            {
                RemovePropW(hostHwnd.get(), kPluginConfigDxFieldHostStateProp);
                RestoreWndProcHook(hostHwnd.get(), kPluginConfigDxFieldHostOriginalWndProcProp);
                if (host)
                {
                    host->Detach();
                    host.reset();
                }
                hostHwnd.reset();
                return;
            }

            static_cast<void>(hostHwnd.release());
        }
        if (host)
        {
            host->Detach();
            host.reset();
        }
    }
};

struct PluginConfigChoiceDxControl
{
    PluginConfigChoiceDxControl()                                              = default;
    PluginConfigChoiceDxControl(const PluginConfigChoiceDxControl&)            = delete;
    PluginConfigChoiceDxControl& operator=(const PluginConfigChoiceDxControl&) = delete;
    PluginConfigChoiceDxControl(PluginConfigChoiceDxControl&&)                 = default;
    PluginConfigChoiceDxControl& operator=(PluginConfigChoiceDxControl&&)      = default;

    PluginConfigDxHostSlot slot;
    Checkbox* checkbox = nullptr;
};

struct PluginConfigFieldControls
{
    PluginConfigFieldControls()                                            = default;
    PluginConfigFieldControls(const PluginConfigFieldControls&)            = delete;
    PluginConfigFieldControls& operator=(const PluginConfigFieldControls&) = delete;
    PluginConfigFieldControls(PluginConfigFieldControls&&)                 = default;
    PluginConfigFieldControls& operator=(PluginConfigFieldControls&&)      = default;

    PluginConfigField field;
    HWND hLabel                 = nullptr;
    HWND hEditFrame             = nullptr;
    HWND hEdit                  = nullptr;
    HWND hComboFrame            = nullptr;
    HWND hCombo                 = nullptr;
    HWND hToggle                = nullptr;
    HWND hComment               = nullptr;
    HWND hDefaults              = nullptr;
    size_t toggleOnChoiceIndex  = 0;
    size_t toggleOffChoiceIndex = 0;
    std::vector<HWND> choiceButtons;

    PluginConfigDxHostSlot dxLabelSlot;
    Label* dxLabelControl = nullptr;
    PluginConfigDxHostSlot dxDescriptionSlot;
    Label* dxDescriptionControl = nullptr;
    PluginConfigDxHostSlot dxDefaultsSlot;
    Label* dxDefaultsControl = nullptr;
    PluginConfigDxHostSlot dxEditSlot;
    TextField* dxEditControl = nullptr;
    PluginConfigDxHostSlot dxComboSlot;
    ComboBox* dxComboControl = nullptr;
    PluginConfigDxHostSlot dxToggleSlot;
    Toggle* dxToggleControl = nullptr;
    std::vector<PluginConfigChoiceDxControl> dxChoiceControls;
};

struct DxCommandButtonHost
{
    DxCommandButtonHost() = default;
    ~DxCommandButtonHost() noexcept
    {
        Detach();
    }
    DxCommandButtonHost(const DxCommandButtonHost&)            = delete;
    DxCommandButtonHost& operator=(const DxCommandButtonHost&) = delete;
    DxCommandButtonHost(DxCommandButtonHost&&)                 = delete;
    DxCommandButtonHost& operator=(DxCommandButtonHost&&)      = delete;

    wil::unique_hwnd hostHwnd;
    WindowHost host;
    Button* button = nullptr;
    UINT commandId = 0u;

    void Detach() noexcept
    {
        button    = nullptr;
        commandId = 0u;
        if (hostHwnd)
        {
            if (IsWindow(hostHwnd.get()) != FALSE)
            {
                RemovePropW(hostHwnd.get(), kPluginConfigDxButtonHostStateProp);
                RestoreWndProcHook(hostHwnd.get(), kPluginConfigDxButtonHostOriginalWndProcProp);
                host.Detach();
                hostHwnd.reset();
                return;
            }

            static_cast<void>(hostHwnd.release());
        }
        host.Detach();
    }
};

enum class PluginConfigCommitMode : uint8_t
{
    ApplyToPluginsAndPersist,
    UpdateSettingsOnly,
};

struct PluginConfigDialogState
{
    PluginConfigDialogState()                                          = default;
    PluginConfigDialogState(const PluginConfigDialogState&)            = delete;
    PluginConfigDialogState& operator=(const PluginConfigDialogState&) = delete;

    Common::Settings::Settings* settings             = nullptr;
    Common::Settings::Settings* reloadSourceSettings = nullptr;
    std::wstring appId;
    AppTheme theme{};
    PluginType pluginType = PluginType::FileSystem;
    std::wstring pluginId;
    std::wstring pluginName;
    std::string schemaJsonUtf8;
    std::string configurationJsonUtf8;
    std::string baselineConfigurationJsonUtf8;
    PluginConfigCommitMode commitMode = PluginConfigCommitMode::ApplyToPluginsAndPersist;
    bool staleFromExternalReload      = false;

    wil::unique_hbrush backgroundBrush;
    wil::unique_hbrush inputBrush;
    COLORREF inputBackgroundColor = RGB(255, 255, 255);
    ThemedInputFrames::FrameStyle inputFrameStyle{};
    HWND panel             = nullptr;
    int contentHeight      = 0;
    int scrollPosY         = 0;
    int fixedWindowWidthPx = 0;
    std::array<DxCommandButtonHost, static_cast<size_t>(DxCommandButtonIndex::Count)> dxCommandButtons;
    bool usesDxUiCommandButtons = false;
    std::vector<PluginConfigFieldControls> controls;
#ifdef ENABLE_TESTS
    HWND lastDebugFocusedHost = nullptr;
    size_t lastDebugFocusedHostIndex = kInvalidPluginConfigDebugFocusIndex;
#endif
};

INT_PTR OnPluginConfigDialogCtlColorStatic(PluginConfigDialogState* state, HDC hdc, HWND control);
INT_PTR OnPluginConfigDialogCtlColorButton(PluginConfigDialogState* state, HDC hdc, HWND control);
INT_PTR OnPluginConfigDialogCtlColorEdit(PluginConfigDialogState* state, HDC hdc);
INT_PTR OnPluginConfigDialogCtlColorListBox(PluginConfigDialogState* state, HDC hdc);
[[nodiscard]] HWND GetLegacyCommandButton(HWND dlg, DxCommandButtonIndex index) noexcept;

[[nodiscard]] HRESULT PersistSettings(HWND owner, Common::Settings::Settings& settings, std::wstring_view appId) noexcept
{
    if (appId.empty())
    {
        return E_INVALIDARG;
    }

    const HRESULT hr = SettingsHotReload::SaveSettingsAndSchema(appId, settings);
    if (SUCCEEDED(hr))
    {
        return S_OK;
    }

    const std::filesystem::path settingsPath = Common::Settings::GetSettingsPath(appId);
    Debug::Error(L"SaveSettings failed (hr=0x{:08X}) path={}", static_cast<unsigned long>(hr), settingsPath.wstring());

    if (! owner)
    {
        return hr;
    }

    const std::wstring message = FormatStringResource(nullptr, IDS_FMT_SETTINGS_SAVE_FAILED, settingsPath.wstring(), static_cast<unsigned long>(hr));
    const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
    ShowDialogAlert(owner, HOST_ALERT_ERROR, title, message);
    return hr;
}

[[nodiscard]] std::wstring GetPluginConfigEditorName(const PluginConfigDialogState& state) noexcept
{
    if (! state.pluginName.empty())
    {
        return state.pluginName;
    }
    if (! state.pluginId.empty())
    {
        return state.pluginId;
    }
    return LoadStringResource(nullptr, IDS_CAPTION_PLUGINS_MANAGER);
}

[[nodiscard]] std::wstring_view TryGetChoiceLabelForValue(const PluginConfigField& field, std::wstring_view value) noexcept
{
    for (const auto& choice : field.choices)
    {
        if (choice.value == value)
        {
            return choice.label.empty() ? std::wstring_view(choice.value) : std::wstring_view(choice.label);
        }
    }
    return {};
}

std::wstring BuildFieldDefaultsTextForDisplay(const PluginConfigField& field)
{
    std::wstring defaults;

    switch (field.type)
    {
        case PluginConfigFieldType::Text:
        {
            if (! field.defaultText.empty())
            {
                defaults = std::format(L"Default: {}", field.defaultText);
            }
            break;
        }
        case PluginConfigFieldType::Value:
        {
            defaults = std::format(LocaleFormatting::GetFormatLocale(), L"Default: {:L}", field.defaultInt);
            if (field.hasMin)
            {
                defaults.append(std::format(LocaleFormatting::GetFormatLocale(), L"   Min: {:L}", field.min));
            }
            if (field.hasMax)
            {
                defaults.append(std::format(LocaleFormatting::GetFormatLocale(), L"   Max: {:L}", field.max));
            }
            break;
        }
        case PluginConfigFieldType::Bool:
        {
            defaults = std::format(L"Default: {}", field.defaultBool ? L"True" : L"False");
            break;
        }
        case PluginConfigFieldType::Option:
        {
            if (! field.defaultOption.empty())
            {
                const std::wstring_view label = TryGetChoiceLabelForValue(field, field.defaultOption);
                if (! label.empty())
                {
                    defaults = std::format(L"Default: {}", label);
                }
                else
                {
                    defaults = std::format(L"Default: {}", field.defaultOption);
                }
            }
            break;
        }
        case PluginConfigFieldType::Selection:
        {
            if (! field.defaultSelection.empty())
            {
                std::wstring joined;
                for (size_t i = 0; i < field.defaultSelection.size(); ++i)
                {
                    if (i > 0)
                    {
                        joined.append(L", ");
                    }

                    const std::wstring_view label = TryGetChoiceLabelForValue(field, field.defaultSelection[i]);
                    joined.append(label.empty() ? field.defaultSelection[i] : std::wstring(label));
                }
                defaults = std::format(L"Default: {}", joined);
            }
            break;
        }
    }

    return defaults;
}

int MeasureInfoHeight(HWND dlg, int width, const std::wstring& text) noexcept
{
    if (text.empty() || width <= 0)
    {
        return 0;
    }

    return Typography::MeasureWrappedTextHeightPx(dlg, FontRole::Body, width, text);
}

[[nodiscard]] std::wstring GetWindowTextString(HWND hwnd) noexcept
{
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return {};
    }

    const int length = GetWindowTextLengthW(hwnd);
    if (length <= 0)
    {
        return {};
    }

    std::wstring text(static_cast<size_t>(length) + 1u, L'\0');
    const int written = GetWindowTextW(hwnd, text.data(), static_cast<int>(text.size()));
    if (written <= 0)
    {
        return {};
    }

    text.resize(static_cast<size_t>(written));
    return text;
}

[[nodiscard]] bool WindowOwnsFocus(const HWND hwnd, const HWND focused) noexcept
{
    return hwnd && focused && (hwnd == focused || IsChild(hwnd, focused) != FALSE);
}

#ifdef ENABLE_TESTS
template <typename ControlT>
[[nodiscard]] bool IsPluginConfigDebugFieldHostInteractive(const PluginConfigDxHostSlot& slot, ControlT* control) noexcept
{
    return control && slot.hostHwnd && IsWindowVisible(slot.hostHwnd.get()) != FALSE && control->IsVisible() && control->IsEnabled();
}

template <typename Visitor> [[nodiscard]] bool VisitPluginConfigDebugInteractiveHosts(const PluginConfigDialogState& state, Visitor&& visitor) noexcept
{
    size_t index = 0u;

    const auto visitFieldHost = [&](const PluginConfigDxHostSlot& slot, auto* control) noexcept
    {
        if (! IsPluginConfigDebugFieldHostInteractive(slot, control))
        {
            return false;
        }

        if (visitor(index, slot.hostHwnd.get()))
        {
            return true;
        }

        ++index;
        return false;
    };

    for (const auto& controls : state.controls)
    {
        if (visitFieldHost(controls.dxEditSlot, controls.dxEditControl) || visitFieldHost(controls.dxComboSlot, controls.dxComboControl) ||
            visitFieldHost(controls.dxToggleSlot, controls.dxToggleControl))
        {
            return true;
        }

        for (const auto& choiceControl : controls.dxChoiceControls)
        {
            if (visitFieldHost(choiceControl.slot, choiceControl.checkbox))
            {
                return true;
            }
        }
    }

    for (const DxCommandButtonHost& slot : state.dxCommandButtons)
    {
        if (! slot.button || ! slot.hostHwnd || IsWindowVisible(slot.hostHwnd.get()) == FALSE || ! slot.button->IsVisible() || ! slot.button->IsEnabled())
        {
            continue;
        }

        if (visitor(index, slot.hostHwnd.get()))
        {
            return true;
        }

        ++index;
    }

    return false;
}

[[nodiscard]] bool FindPluginConfigDebugInteractiveHostIndex(const PluginConfigDialogState& state, const HWND host, size_t& outIndex) noexcept
{
    outIndex = kInvalidPluginConfigDebugFocusIndex;
    if (! host)
    {
        return false;
    }

    return VisitPluginConfigDebugInteractiveHosts(
        state,
        [&](const size_t index, const HWND candidate) noexcept
    {
        if (candidate != host)
        {
            return false;
        }

        outIndex = index;
        return true;
    });
}

[[nodiscard]] bool TryResolvePluginConfigDebugFocusHostByIndex(const PluginConfigDialogState& state, HWND& outHost) noexcept
{
    outHost = nullptr;
    if (state.lastDebugFocusedHostIndex == kInvalidPluginConfigDebugFocusIndex)
    {
        return false;
    }

    return VisitPluginConfigDebugInteractiveHosts(
        state,
        [&](const size_t index, const HWND candidate) noexcept
    {
        if (index != state.lastDebugFocusedHostIndex)
        {
            return false;
        }

        outHost = candidate;
        return outHost != nullptr;
    });
}

void RememberPluginConfigDebugFocusedHost(PluginConfigDialogState& state, const HWND host, const size_t hostIndex) noexcept
{
    state.lastDebugFocusedHost      = host;
    state.lastDebugFocusedHostIndex = hostIndex;
}

void RememberPluginConfigDebugFocusedHost(PluginConfigDialogState& state, const HWND host) noexcept
{
    size_t hostIndex = kInvalidPluginConfigDebugFocusIndex;
    static_cast<void>(FindPluginConfigDebugInteractiveHostIndex(state, host, hostIndex));
    RememberPluginConfigDebugFocusedHost(state, host, hostIndex);
}

[[nodiscard]] bool TryRecoverPluginConfigDebugFocusedHost(const PluginConfigDialogState& state, HWND& outHost) noexcept
{
    outHost = nullptr;

    if (state.lastDebugFocusedHost && IsWindow(state.lastDebugFocusedHost) != FALSE && IsWindowVisible(state.lastDebugFocusedHost) != FALSE)
    {
        outHost = state.lastDebugFocusedHost;
        return true;
    }

    return TryResolvePluginConfigDebugFocusHostByIndex(state, outHost);
}

[[nodiscard]] bool TryFillPluginConfigFocusSnapshotFromDebugHost(const HWND dlg,
                                                                 const PluginConfigDialogState& state,
                                                                 const HWND host,
                                                                 PluginConfigurationDialogDebugSnapshot& snapshot) noexcept
{
    if (! host)
    {
        return false;
    }

    for (const auto& controls : state.controls)
    {
        if (controls.dxEditSlot.hostHwnd.get() == host)
        {
            snapshot.focusKind  = PluginConfigurationDialogDebugFocusKind::Edit;
            snapshot.focusLabel = controls.field.label;
            return true;
        }

        if (controls.dxComboSlot.hostHwnd.get() == host)
        {
            snapshot.focusKind  = PluginConfigurationDialogDebugFocusKind::Combo;
            snapshot.focusLabel = controls.field.label;
            return true;
        }

        if (controls.dxToggleSlot.hostHwnd.get() == host)
        {
            snapshot.focusKind  = PluginConfigurationDialogDebugFocusKind::Toggle;
            snapshot.focusLabel = controls.field.label;
            return true;
        }

        for (const auto& choiceControl : controls.dxChoiceControls)
        {
            if (choiceControl.slot.hostHwnd.get() == host)
            {
                snapshot.focusKind  = PluginConfigurationDialogDebugFocusKind::Choice;
                snapshot.focusLabel = controls.field.label;
                return true;
            }
        }
    }

    for (size_t i = 0; i < static_cast<size_t>(DxCommandButtonIndex::Count); ++i)
    {
        if (state.dxCommandButtons[i].hostHwnd.get() == host)
        {
            const auto index     = static_cast<DxCommandButtonIndex>(i);
            const HWND legacyBtn = GetLegacyCommandButton(dlg, index);
            snapshot.focusKind   = PluginConfigurationDialogDebugFocusKind::CommandButton;
            snapshot.focusLabel  = GetWindowTextString(legacyBtn);
            return true;
        }
    }

    return false;
}
#endif

[[nodiscard]] bool TryRoutePluginConfigDialogCommandKey(const HWND hostHwnd, const WPARAM vk) noexcept
{
    const HWND dlg = hostHwnd ? GetAncestor(hostHwnd, GA_ROOT) : nullptr;
    if (! dlg || IsWindow(dlg) == FALSE)
    {
        return false;
    }

    UINT commandId = 0u;
    switch (vk)
    {
        case VK_RETURN: commandId = IDOK; break;
        case VK_ESCAPE: commandId = IDCANCEL; break;
        default: return false;
    }

    return PostMessageW(dlg, WM_COMMAND, MAKEWPARAM(commandId, 0), 0) != FALSE;
}

#ifdef ENABLE_TESTS
[[nodiscard]] HWND GetFirstVisibleInteractiveTarget(const PluginConfigDialogState& state) noexcept
{
    for (const auto& controls : state.controls)
    {
        if (controls.dxEditSlot.hostHwnd && IsWindowVisible(controls.dxEditSlot.hostHwnd.get()) != FALSE)
        {
            return controls.dxEditSlot.hostHwnd.get();
        }
        if (controls.dxComboSlot.hostHwnd && IsWindowVisible(controls.dxComboSlot.hostHwnd.get()) != FALSE)
        {
            return controls.dxComboSlot.hostHwnd.get();
        }
        if (controls.dxToggleSlot.hostHwnd && IsWindowVisible(controls.dxToggleSlot.hostHwnd.get()) != FALSE)
        {
            return controls.dxToggleSlot.hostHwnd.get();
        }
        for (const auto& choiceControl : controls.dxChoiceControls)
        {
            if (choiceControl.slot.hostHwnd && IsWindowVisible(choiceControl.slot.hostHwnd.get()) != FALSE)
            {
                return choiceControl.slot.hostHwnd.get();
            }
        }
    }

    return nullptr;
}

[[nodiscard]] bool FocusFirstVisibleInteractiveTarget(PluginConfigDialogState& state) noexcept
{
    const auto focusDxControl = [](PluginConfigDxHostSlot& slot, auto* control) noexcept
    {
        if (! control || ! slot.host || ! slot.hostHwnd || IsWindowVisible(slot.hostHwnd.get()) == FALSE)
        {
            return false;
        }

        slot.host->SetFocusControl(control);
        slot.host->Invalidate();
        return slot.host->GetFocusControl() == control || WindowOwnsFocus(slot.hostHwnd.get(), GetFocus());
    };

    for (auto& controls : state.controls)
    {
        if (focusDxControl(controls.dxEditSlot, controls.dxEditControl))
        {
            return true;
        }
        if (focusDxControl(controls.dxComboSlot, controls.dxComboControl))
        {
            return true;
        }
        if (focusDxControl(controls.dxToggleSlot, controls.dxToggleControl))
        {
            return true;
        }
        for (auto& choiceControl : controls.dxChoiceControls)
        {
            if (focusDxControl(choiceControl.slot, choiceControl.checkbox))
            {
                return true;
            }
        }
    }

    const HWND fallbackTarget = GetFirstVisibleInteractiveTarget(state);
    if (fallbackTarget && IsWindowVisible(fallbackTarget) != FALSE)
    {
        SetFocus(fallbackTarget);
        return true;
    }

    return false;
}

[[nodiscard]] bool EditSlotOwnsFocus(const PluginConfigFieldControls& controls, const HWND focused) noexcept
{
    if (focused)
    {
        if (WindowOwnsFocus(controls.dxEditSlot.hostHwnd.get(), focused) || WindowOwnsFocus(controls.hEdit, focused))
        {
            return true;
        }

        return controls.dxEditSlot.host && WindowOwnsFocus(controls.dxEditSlot.host->GetTextInputHwnd(), focused);
    }

    return controls.dxEditControl && controls.dxEditSlot.host && controls.dxEditSlot.host->GetFocusControl() == controls.dxEditControl;
}

[[nodiscard]] bool ComboSlotOwnsFocus(const PluginConfigFieldControls& controls, const HWND focused) noexcept
{
    if (focused)
    {
        return WindowOwnsFocus(controls.dxComboSlot.hostHwnd.get(), focused) || WindowOwnsFocus(controls.hCombo, focused);
    }

    return controls.dxComboControl && controls.dxComboSlot.host && controls.dxComboSlot.host->GetFocusControl() == controls.dxComboControl;
}

[[nodiscard]] bool ToggleSlotOwnsFocus(const PluginConfigFieldControls& controls, const HWND focused) noexcept
{
    if (focused)
    {
        return WindowOwnsFocus(controls.dxToggleSlot.hostHwnd.get(), focused) || WindowOwnsFocus(controls.hToggle, focused);
    }

    return controls.dxToggleControl && controls.dxToggleSlot.host && controls.dxToggleSlot.host->GetFocusControl() == controls.dxToggleControl;
}

[[nodiscard]] bool ChoiceSlotOwnsFocus(const PluginConfigFieldControls& controls, const size_t choiceIndex, const HWND focused) noexcept
{
    if (choiceIndex >= controls.dxChoiceControls.size())
    {
        return false;
    }

    const auto& choiceControl = controls.dxChoiceControls[choiceIndex];
    const HWND legacyChoice   = choiceIndex < controls.choiceButtons.size() ? controls.choiceButtons[choiceIndex] : nullptr;
    if (focused)
    {
        return WindowOwnsFocus(choiceControl.slot.hostHwnd.get(), focused) || WindowOwnsFocus(legacyChoice, focused);
    }

    return choiceControl.checkbox && choiceControl.slot.host && choiceControl.slot.host->GetFocusControl() == choiceControl.checkbox;
}

[[nodiscard]] bool CommandButtonHostOwnsFocus(const DxCommandButtonHost& slot, const HWND legacyButton, const HWND focused) noexcept
{
    if (focused)
    {
        return WindowOwnsFocus(slot.hostHwnd.get(), focused) || WindowOwnsFocus(legacyButton, focused);
    }

    return slot.button && slot.host.GetFocusControl() == slot.button;
}

void FillPluginConfigFocusSnapshot(const HWND dlg, const PluginConfigDialogState& state, PluginConfigurationDialogDebugSnapshot& snapshot) noexcept
{
    snapshot.focusKind = PluginConfigurationDialogDebugFocusKind::None;
    snapshot.focusLabel.clear();

    const HWND focused = GetFocus();
    if (focused && WindowOwnsFocus(state.panel, focused))
    {
        snapshot.focusKind  = PluginConfigurationDialogDebugFocusKind::Panel;
        snapshot.focusLabel = L"Scroll panel";
    }

    for (const auto& controls : state.controls)
    {
        if (EditSlotOwnsFocus(controls, focused))
        {
            snapshot.focusKind  = PluginConfigurationDialogDebugFocusKind::Edit;
            snapshot.focusLabel = controls.field.label;
            return;
        }

        if (ComboSlotOwnsFocus(controls, focused))
        {
            snapshot.focusKind  = PluginConfigurationDialogDebugFocusKind::Combo;
            snapshot.focusLabel = controls.field.label;
            return;
        }

        if (ToggleSlotOwnsFocus(controls, focused))
        {
            snapshot.focusKind  = PluginConfigurationDialogDebugFocusKind::Toggle;
            snapshot.focusLabel = controls.field.label;
            return;
        }

        for (size_t choiceIndex = 0; choiceIndex < controls.dxChoiceControls.size(); ++choiceIndex)
        {
            if (ChoiceSlotOwnsFocus(controls, choiceIndex, focused))
            {
                snapshot.focusKind  = PluginConfigurationDialogDebugFocusKind::Choice;
                snapshot.focusLabel = controls.field.label;
                return;
            }
        }
    }

    for (size_t i = 0; i < static_cast<size_t>(DxCommandButtonIndex::Count); ++i)
    {
        const auto index        = static_cast<DxCommandButtonIndex>(i);
        const HWND legacyButton = GetLegacyCommandButton(dlg, index);
        if (CommandButtonHostOwnsFocus(state.dxCommandButtons[i], legacyButton, focused))
        {
            snapshot.focusKind  = PluginConfigurationDialogDebugFocusKind::CommandButton;
            snapshot.focusLabel = GetWindowTextString(legacyButton);
            return;
        }
    }

    HWND recoveredHost = nullptr;
    if (TryRecoverPluginConfigDebugFocusedHost(state, recoveredHost))
    {
        static_cast<void>(TryFillPluginConfigFocusSnapshotFromDebugHost(dlg, state, recoveredHost, snapshot));
    }
}
#endif

[[nodiscard]] std::wstring SanitizeIntegerText(std::wstring_view text, const bool allowNegative) noexcept
{
    std::wstring sanitized;
    sanitized.reserve(text.size());

    bool sawDigit = false;
    bool sawSign  = false;
    for (const wchar_t ch : text)
    {
        if (ch >= L'0' && ch <= L'9')
        {
            sanitized.push_back(ch);
            sawDigit = true;
            continue;
        }

        if (allowNegative && ! sawSign && ! sawDigit && ch == L'-')
        {
            sanitized.push_back(ch);
            sawSign = true;
        }
    }

    return sanitized;
}

[[nodiscard]] bool GetChildWindowRect(HWND parent, HWND child, RECT& out) noexcept
{
    if (! parent || ! child || IsWindow(parent) == FALSE || IsWindow(child) == FALSE)
    {
        return false;
    }

    if (GetWindowRect(child, &out) == FALSE)
    {
        return false;
    }

    MapWindowPoints(HWND_DESKTOP, parent, reinterpret_cast<POINT*>(&out), 2);
    return true;
}

[[nodiscard]] bool GetChoiceButtonsUnionRect(HWND parent, const std::vector<HWND>& choiceButtons, RECT& out) noexcept
{
    bool first = true;
    for (const HWND choiceButton : choiceButtons)
    {
        RECT rect{};
        if (! GetChildWindowRect(parent, choiceButton, rect))
        {
            continue;
        }

        if (first)
        {
            out   = rect;
            first = false;
            continue;
        }

        out.left   = std::min(out.left, rect.left);
        out.top    = std::min(out.top, rect.top);
        out.right  = std::max(out.right, rect.right);
        out.bottom = std::max(out.bottom, rect.bottom);
    }

    return ! first;
}

[[nodiscard]] bool GetLegacyToggleState(HWND toggle) noexcept
{
    if (! toggle || IsWindow(toggle) == FALSE)
    {
        return false;
    }

    const LONG_PTR style = GetWindowLongPtrW(toggle, GWL_STYLE);
    if ((style & BS_TYPEMASK) == BS_OWNERDRAW)
    {
        return GetWindowLongPtrW(toggle, GWLP_USERDATA) != 0;
    }

    return SendMessageW(toggle, BM_GETCHECK, 0, 0) == BST_CHECKED;
}

[[nodiscard]] bool UsesLegacyOwnerDrawButton(HWND button) noexcept
{
    if (! button || IsWindow(button) == FALSE)
    {
        return false;
    }

    return (GetWindowLongPtrW(button, GWL_STYLE) & BS_TYPEMASK) == BS_OWNERDRAW;
}

void SetLegacyToggleState(HWND toggle, bool checked) noexcept
{
    if (! toggle || IsWindow(toggle) == FALSE)
    {
        return;
    }

    if (UsesLegacyOwnerDrawButton(toggle))
    {
        SetWindowLongPtrW(toggle, GWLP_USERDATA, checked ? 1 : 0);
        InvalidateRect(toggle, nullptr, TRUE);
        return;
    }

    SendMessageW(toggle, BM_SETCHECK, checked ? BST_CHECKED : BST_UNCHECKED, 0);
}

void NormalizeLegacyToggleForHiddenDxSync(HWND toggle, const bool checked) noexcept
{
    if (! UsesLegacyOwnerDrawButton(toggle))
    {
        SetLegacyToggleState(toggle, checked);
        return;
    }

    LONG_PTR style = GetWindowLongPtrW(toggle, GWL_STYLE);
    style &= ~BS_TYPEMASK;
    style |= BS_AUTOCHECKBOX;
    SetWindowLongPtrW(toggle, GWL_STYLE, style);
    SetWindowPos(toggle, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
    SetLegacyToggleState(toggle, checked);
}

void HideLegacyWindow(HWND hwnd) noexcept
{
    if (hwnd && IsWindow(hwnd) != FALSE)
    {
        ShowWindow(hwnd, SW_HIDE);
    }
}

void PostLegacyConfigCommand(HWND parent, UINT notifyCode, HWND legacyControl) noexcept
{
    if (! parent || ! legacyControl || IsWindow(parent) == FALSE || IsWindow(legacyControl) == FALSE)
    {
        return;
    }

    PostMessageW(parent, WM_COMMAND, MAKEWPARAM(0, notifyCode), reinterpret_cast<LPARAM>(legacyControl));
}

[[nodiscard]] bool MovePluginConfigDialogTabFocusFromHost(HWND hostHwnd, bool reverse) noexcept;

constexpr wchar_t kPluginConfigDxHostClassName[] = L"RedSalamander.PluginConfig.DxHost";

[[nodiscard]] bool EnsurePluginConfigDxHostClassRegistered() noexcept
{
    const HINSTANCE instance = GetModuleHandleW(nullptr);
    WNDCLASSEXW existing{};
    existing.cbSize = sizeof(existing);
    if (GetClassInfoExW(instance, kPluginConfigDxHostClassName, &existing) != 0)
    {
        return true;
    }

    WNDCLASSEXW wc{};
    wc.cbSize        = sizeof(wc);
    wc.lpfnWndProc   = DefWindowProcW;
    wc.hInstance     = instance;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = nullptr;
    wc.lpszClassName = kPluginConfigDxHostClassName;

    const ATOM atom = RegisterClassExW(&wc);
    return atom != 0 || GetLastError() == ERROR_CLASS_ALREADY_EXISTS;
}

[[nodiscard]] HWND CreatePluginConfigDxHostWindow(HWND parent, DWORD style) noexcept
{
    if (! EnsurePluginConfigDxHostClassRegistered())
    {
        return nullptr;
    }

    return CreateWindowExW(0, kPluginConfigDxHostClassName, L"", style, 0, 0, 10, 10, parent, nullptr, GetModuleHandleW(nullptr), nullptr);
}

template <typename ControlT> void DetachDxHost(PluginConfigDxHostSlot& slot, ControlT*& control) noexcept
{
    control = nullptr;
    slot.Detach();
}

template <typename ControlT, typename InitFn>
[[nodiscard]] bool EnsureDxHost(HWND parent, PluginConfigDxHostSlot& slot, ControlT*& control, const DWORD style, InitFn&& initFn) noexcept
{
    if (slot.host && slot.hostHwnd && IsWindow(slot.hostHwnd.get()) != FALSE)
    {
        return true;
    }

    DetachDxHost(slot, control);

    wil::unique_hwnd hwnd(CreatePluginConfigDxHostWindow(parent, style));
    if (! hwnd)
    {
        return false;
    }

    auto dxHost = std::make_unique<WindowHost>();
    if (! dxHost->Attach(hwnd.get()))
    {
        return false;
    }
    dxHost->SetOnTabBoundary(
        std::function<bool(bool)>{[hostHwnd = hwnd.get()](bool reverse) -> bool { return MovePluginConfigDialogTabFocusFromHost(hostHwnd, reverse); }});

    if (! SetPropW(hwnd.get(), kPluginConfigDxFieldHostStateProp, reinterpret_cast<HANDLE>(&slot)) ||
        ! InstallWndProcHook(hwnd.get(), kPluginConfigDxFieldHostOriginalWndProcProp, PluginConfigDxFieldHostWndProc))
    {
        RemovePropW(hwnd.get(), kPluginConfigDxFieldHostStateProp);
        Debug::ErrorWithLastError(L"PluginConfig: failed to install owned WNDPROC hook for DX field host.");
        dxHost->Detach();
        return false;
    }

    auto dxControl = std::make_unique<ControlT>();
    control        = dxControl.get();
    initFn(*dxControl);
    dxHost->SetRoot(std::move(dxControl));
    slot.hostHwnd = std::move(hwnd);
    slot.host     = std::move(dxHost);
    return true;
}

void HideDxHost(const PluginConfigDxHostSlot& slot) noexcept
{
    if (slot.hostHwnd)
    {
        ShowWindow(slot.hostHwnd.get(), SW_HIDE);
    }
}

template <typename ControlT> void LayoutDxHost(const PluginConfigDxHostSlot& slot, WindowHost* host, ControlT* control, const RECT& rect) noexcept
{
    if (! slot.hostHwnd || ! host || ! control || rect.right <= rect.left || rect.bottom <= rect.top)
    {
        HideDxHost(slot);
        return;
    }

    SetWindowPos(
        slot.hostHwnd.get(), nullptr, rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top, SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW);
    host->Invalidate();
}

[[maybe_unused]] [[nodiscard]] size_t CountIfVisible(const HWND hwnd) noexcept
{
    return (hwnd && IsWindowVisible(hwnd) != FALSE) ? 1u : 0u;
}

void ApplyDxFieldTheme(const PluginConfigDialogState& state, PluginConfigFieldControls& controls) noexcept
{
    const ThemePalette palette = MakeAppThemeDxPalette(state.theme, state.theme.windowBackground);

    const auto applyTheme = [&](PluginConfigDxHostSlot& slot) noexcept
    {
        if (slot.host)
        {
            slot.host->SetTheme(palette);
        }
    };

    applyTheme(controls.dxLabelSlot);
    applyTheme(controls.dxDescriptionSlot);
    applyTheme(controls.dxDefaultsSlot);
    applyTheme(controls.dxEditSlot);
    applyTheme(controls.dxComboSlot);
    applyTheme(controls.dxToggleSlot);
    for (auto& choiceControl : controls.dxChoiceControls)
    {
        applyTheme(choiceControl.slot);
    }

    if (controls.dxDescriptionControl)
    {
        controls.dxDescriptionControl->SetTextColor(ColorFromCOLORREF(state.theme.menu.shortcutText));
    }
    if (controls.dxDefaultsControl)
    {
        controls.dxDefaultsControl->SetTextColor(ColorFromCOLORREF(state.theme.menu.shortcutText));
    }
}

void AttachDxFieldHosts(
    HWND panel, PluginConfigDialogState& state, PluginConfigFieldControls& controls, const bool attachStatics, const bool attachInputs) noexcept
{
    if (! panel)
    {
        return;
    }

    const auto layoutAndHide = [&](PluginConfigDxHostSlot& slot, auto* control, HWND legacy, const bool hideLegacy = true) noexcept
    {
        RECT rect{};
        if (! GetChildWindowRect(panel, legacy, rect))
        {
            HideDxHost(slot);
            return;
        }

        LayoutDxHost(slot, slot.host.get(), control, rect);
        if (hideLegacy)
        {
            HideLegacyWindow(legacy);
        }
    };

    if (attachStatics && controls.hLabel &&
        EnsureDxHost(panel,
                     controls.dxLabelSlot,
                     controls.dxLabelControl,
                     WS_CHILD | SS_NOTIFY,
                     [&](Label& label) noexcept
    {
        label.SetText(controls.field.label);
        label.SetFontRole(FontRole::Body);
        label.SetMultiline(false);
    }))
    {
        layoutAndHide(controls.dxLabelSlot, controls.dxLabelControl, controls.hLabel);
    }

    if (attachStatics && controls.hComment)
    {
        const std::wstring description = GetWindowTextString(controls.hComment);
        if (EnsureDxHost(panel,
                         controls.dxDescriptionSlot,
                         controls.dxDescriptionControl,
                         WS_CHILD | SS_NOTIFY,
                         [&](Label& label) noexcept
        {
            label.SetText(description);
            label.SetMultiline(true);
        }))
        {
            layoutAndHide(controls.dxDescriptionSlot, controls.dxDescriptionControl, controls.hComment);
        }
    }

    if (attachStatics && controls.hDefaults)
    {
        const std::wstring defaultsText = GetWindowTextString(controls.hDefaults);
        if (EnsureDxHost(panel,
                         controls.dxDefaultsSlot,
                         controls.dxDefaultsControl,
                         WS_CHILD | SS_NOTIFY,
                         [&](Label& label) noexcept
        {
            label.SetText(defaultsText);
            label.SetMultiline(true);
        }))
        {
            layoutAndHide(controls.dxDefaultsSlot, controls.dxDefaultsControl, controls.hDefaults);
        }
    }

    if (attachInputs && controls.hEdit)
    {
        const bool numericOnly = controls.field.type == PluginConfigFieldType::Value;
        if (EnsureDxHost(panel, controls.dxEditSlot, controls.dxEditControl, WS_CHILD | WS_TABSTOP | SS_NOTIFY, [](TextField& field) noexcept {
            field.SetMultiline(false);
        }))
        {
            controls.dxEditControl->SetText(GetWindowTextString(controls.hEdit));
            controls.dxEditControl->SetOnTextChanged(
                [legacy = controls.hEdit, numericOnly, fieldControl = controls.dxEditControl](std::wstring_view text) noexcept
            {
                std::wstring normalized = numericOnly ? SanitizeIntegerText(text, false) : std::wstring(text);
                if (normalized != text)
                {
                    fieldControl->SetText(std::move(normalized));
                    return;
                }

                if (GetWindowTextString(legacy) == normalized)
                {
                    return;
                }

                SetWindowTextW(legacy, normalized.c_str());
            });
            controls.dxEditControl->SetOnSubmitted([panel, legacy = controls.hEdit]() noexcept { PostLegacyConfigCommand(panel, EN_KILLFOCUS, legacy); });
            controls.dxEditControl->SetOnBlur([panel, legacy = controls.hEdit]() noexcept { PostLegacyConfigCommand(panel, EN_KILLFOCUS, legacy); });

            const HWND anchor = controls.hEditFrame ? controls.hEditFrame : controls.hEdit;
            layoutAndHide(controls.dxEditSlot, controls.dxEditControl, anchor, false);
            HideLegacyWindow(controls.hEditFrame);
            HideLegacyWindow(controls.hEdit);
        }
    }

    if (attachInputs && controls.hCombo)
    {
        if (EnsureDxHost(panel, controls.dxComboSlot, controls.dxComboControl, WS_CHILD | WS_TABSTOP | SS_NOTIFY, [](ComboBox& combo) noexcept {
            combo.SetVariant(ComboBoxVariant::Modern);
        }))
        {
            std::vector<ComboBox::Item> items;
            items.reserve(controls.field.choices.size());
            for (const auto& choice : controls.field.choices)
            {
                items.push_back({choice.value, choice.label.empty() ? choice.value : choice.label});
            }
            controls.dxComboControl->SetItems(std::move(items));

            std::optional<size_t> selectedIndex;
            const LRESULT selection = SendMessageW(controls.hCombo, CB_GETCURSEL, 0, 0);
            if (selection >= 0 && selection <= static_cast<LRESULT>(std::numeric_limits<size_t>::max()))
            {
                selectedIndex = static_cast<size_t>(selection);
            }

            controls.dxComboControl->SetSelectedIndex(selectedIndex);
            if (selectedIndex.has_value() && selectedIndex.value() < controls.field.choices.size())
            {
                const auto& choice = controls.field.choices[selectedIndex.value()];
                controls.dxComboControl->SetText(choice.label.empty() ? choice.value : choice.label);
            }

            controls.dxComboControl->SetOnSelectionChanged([legacy = controls.hCombo](const size_t index) noexcept
            {
                if (! legacy || IsWindow(legacy) == FALSE)
                {
                    return;
                }

                SendMessageW(legacy, CB_SETCURSEL, static_cast<WPARAM>(index), 0);
            });

            const HWND anchor = controls.hComboFrame ? controls.hComboFrame : controls.hCombo;
            layoutAndHide(controls.dxComboSlot, controls.dxComboControl, anchor, false);
            HideLegacyWindow(controls.hComboFrame);
            HideLegacyWindow(controls.hCombo);
        }
    }

    const bool usesChoiceToggle =
        ! controls.hToggle && controls.choiceButtons.size() == 2u &&
        (controls.field.type == PluginConfigFieldType::Bool || (controls.field.type == PluginConfigFieldType::Option && controls.field.choices.size() == 2u));
    if (attachInputs && (controls.hToggle || usesChoiceToggle))
    {
        std::wstring checkedLabel   = L"On";
        std::wstring uncheckedLabel = L"Off";

        if (controls.hToggle)
        {
            if (controls.toggleOnChoiceIndex < controls.field.choices.size())
            {
                const auto& choice = controls.field.choices[controls.toggleOnChoiceIndex];
                checkedLabel       = choice.label.empty() ? choice.value : choice.label;
            }
            if (controls.toggleOffChoiceIndex < controls.field.choices.size())
            {
                const auto& choice = controls.field.choices[controls.toggleOffChoiceIndex];
                uncheckedLabel     = choice.label.empty() ? choice.value : choice.label;
            }
            if (controls.field.type == PluginConfigFieldType::Bool && controls.field.choices.empty())
            {
                checkedLabel   = L"True";
                uncheckedLabel = L"False";
            }
        }
        else if (controls.choiceButtons.size() == 2u)
        {
            checkedLabel   = GetWindowTextString(controls.choiceButtons[0]);
            uncheckedLabel = GetWindowTextString(controls.choiceButtons[1]);
        }

        if (EnsureDxHost(panel, controls.dxToggleSlot, controls.dxToggleControl, WS_CHILD | WS_TABSTOP | SS_NOTIFY, [&](Toggle& toggle) noexcept {
            toggle.SetStateLabels(uncheckedLabel, checkedLabel);
        }))
        {
            const bool checked = controls.hToggle ? GetLegacyToggleState(controls.hToggle)
                                                  : (controls.choiceButtons[0] && SendMessageW(controls.choiceButtons[0], BM_GETCHECK, 0, 0) == BST_CHECKED);
            controls.dxToggleControl->SetChecked(checked);
            if (controls.hToggle)
            {
                controls.dxToggleControl->SetOnToggled([legacy = controls.hToggle](const bool isChecked) noexcept
                {
                    if (! legacy || IsWindow(legacy) == FALSE)
                    {
                        return;
                    }

                    SetLegacyToggleState(legacy, isChecked);
                });
            }
            else
            {
                controls.dxToggleControl->SetOnToggled([first = controls.choiceButtons[0], second = controls.choiceButtons[1]](const bool isChecked) noexcept
                {
                    if (first && IsWindow(first) != FALSE)
                    {
                        SendMessageW(first, BM_SETCHECK, isChecked ? BST_CHECKED : BST_UNCHECKED, 0);
                    }
                    if (second && IsWindow(second) != FALSE)
                    {
                        SendMessageW(second, BM_SETCHECK, isChecked ? BST_UNCHECKED : BST_CHECKED, 0);
                    }
                });
            }

            RECT rect{};
            const bool haveRect =
                controls.hToggle ? GetChildWindowRect(panel, controls.hToggle, rect) : GetChoiceButtonsUnionRect(panel, controls.choiceButtons, rect);
            if (haveRect)
            {
                LayoutDxHost(controls.dxToggleSlot, controls.dxToggleSlot.host.get(), controls.dxToggleControl, rect);
                if (controls.hToggle)
                {
                    NormalizeLegacyToggleForHiddenDxSync(controls.hToggle, checked);
                }
                HideLegacyWindow(controls.hToggle);
                for (const HWND choiceButton : controls.choiceButtons)
                {
                    HideLegacyWindow(choiceButton);
                }
            }
        }
    }
    else if (attachInputs && ! controls.choiceButtons.empty())
    {
        controls.dxChoiceControls.clear();
        controls.dxChoiceControls.resize(controls.choiceButtons.size());

        for (size_t i = 0; i < controls.choiceButtons.size(); ++i)
        {
            const HWND legacyButton = controls.choiceButtons[i];
            if (! legacyButton)
            {
                continue;
            }

            auto& dxChoice                = controls.dxChoiceControls[i];
            const std::wstring choiceText = GetWindowTextString(legacyButton);
            if (! EnsureDxHost(panel, dxChoice.slot, dxChoice.checkbox, WS_CHILD | WS_TABSTOP | SS_NOTIFY, [&](Checkbox& checkbox) noexcept {
                checkbox.SetText(choiceText);
            }))
            {
                continue;
            }

            dxChoice.checkbox->SetChecked(SendMessageW(legacyButton, BM_GETCHECK, 0, 0) == BST_CHECKED);
            dxChoice.checkbox->SetOnToggled([legacy = legacyButton](const bool isChecked) noexcept
            {
                if (! legacy || IsWindow(legacy) == FALSE)
                {
                    return;
                }

                SendMessageW(legacy, BM_SETCHECK, isChecked ? BST_CHECKED : BST_UNCHECKED, 0);
            });

            RECT rect{};
            if (GetChildWindowRect(panel, legacyButton, rect))
            {
                LayoutDxHost(dxChoice.slot, dxChoice.slot.host.get(), dxChoice.checkbox, rect);
                HideLegacyWindow(legacyButton);
            }
        }
    }

    ApplyDxFieldTheme(state, controls);
}

[[nodiscard]] HWND GetLegacyCommandButton(const HWND dlg, const DxCommandButtonIndex index) noexcept
{
    if (! dlg)
    {
        return nullptr;
    }

    switch (index)
    {
        case DxCommandButtonIndex::Ok: return GetDlgItem(dlg, IDOK);
        case DxCommandButtonIndex::Cancel: return GetDlgItem(dlg, IDCANCEL);
        case DxCommandButtonIndex::Count: break;
    }

    return nullptr;
}

[[nodiscard]] HWND CreateDxCommandButtonHostWindow(HWND parent) noexcept
{
    return CreatePluginConfigDxHostWindow(parent, WS_CHILD | WS_TABSTOP | SS_NOTIFY);
}

[[nodiscard]] bool MovePluginConfigDialogTabFocusFromHost(HWND hostHwnd, bool reverse) noexcept
{
    if (! hostHwnd || IsWindow(hostHwnd) == FALSE)
    {
#ifdef ENABLE_TESTS
        TracePluginConfigDebug(L"plugin-config advance: invalid current host");
#endif
        return false;
    }

    const HWND dlg = GetAncestor(hostHwnd, GA_ROOT);
    if (! dlg)
    {
#ifdef ENABLE_TESTS
        TracePluginConfigDebug(std::format(L"plugin-config advance: no root dialog for host={:#x}", reinterpret_cast<uintptr_t>(hostHwnd)));
#endif
        return false;
    }

    auto* state = reinterpret_cast<PluginConfigDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));
    if (! state)
    {
#ifdef ENABLE_TESTS
        TracePluginConfigDebug(std::format(
            L"plugin-config advance: no dialog state host={:#x} dlg={:#x}", reinterpret_cast<uintptr_t>(hostHwnd), reinterpret_cast<uintptr_t>(dlg)));
#endif
        return false;
    }

    struct InteractiveHost
    {
        HWND host = nullptr;
        std::function<bool()> focus;
    };

    std::vector<InteractiveHost> hosts;
    hosts.reserve((state->controls.size() * 4u) + state->dxCommandButtons.size());

    const auto appendFieldHost = [&](PluginConfigDxHostSlot& slot, auto* control) noexcept
    {
        if (! control || ! slot.hostHwnd || ! slot.host || IsWindowVisible(slot.hostHwnd.get()) == FALSE || ! control->IsVisible() || ! control->IsEnabled())
        {
            return;
        }

        hosts.push_back(InteractiveHost{.host  = slot.hostHwnd.get(),
                                        .focus = [&slot, control]() noexcept
        {
            SetFocus(slot.hostHwnd.get());
            slot.host->SetFocusControl(control);
            slot.host->Invalidate();
            return slot.host->GetFocusControl() == control || WindowOwnsFocus(slot.hostHwnd.get(), GetFocus());
        }});
    };

    for (auto& controls : state->controls)
    {
        appendFieldHost(controls.dxEditSlot, controls.dxEditControl);
        appendFieldHost(controls.dxComboSlot, controls.dxComboControl);
        appendFieldHost(controls.dxToggleSlot, controls.dxToggleControl);
        for (auto& choiceControl : controls.dxChoiceControls)
        {
            appendFieldHost(choiceControl.slot, choiceControl.checkbox);
        }
    }

    for (DxCommandButtonHost& slot : state->dxCommandButtons)
    {
        if (! slot.button || ! slot.hostHwnd || IsWindowVisible(slot.hostHwnd.get()) == FALSE || ! slot.button->IsVisible() || ! slot.button->IsEnabled())
        {
            continue;
        }

        hosts.push_back(InteractiveHost{.host  = slot.hostHwnd.get(),
                                        .focus = [&slot]() noexcept
        {
            SetFocus(slot.hostHwnd.get());
            slot.host.SetFocusControl(slot.button);
            slot.host.Invalidate();
            return slot.host.GetFocusControl() == slot.button || WindowOwnsFocus(slot.hostHwnd.get(), GetFocus());
        }});
    }

    if (hosts.empty())
    {
#ifdef ENABLE_TESTS
        TracePluginConfigDebug(std::format(L"plugin-config advance: no visible interactive hosts dlg={:#x}", reinterpret_cast<uintptr_t>(dlg)));
#endif
        return false;
    }

    size_t currentIndex = hosts.size();
    for (size_t index = 0; index < hosts.size(); ++index)
    {
        if (hosts[index].host == hostHwnd)
        {
            currentIndex = index;
            break;
        }
    }

    const size_t nextIndex = [&]() noexcept
    {
        if (currentIndex >= hosts.size())
        {
            return reverse ? (hosts.size() - 1u) : 0u;
        }

        if (reverse)
        {
            return currentIndex == 0u ? (hosts.size() - 1u) : (currentIndex - 1u);
        }

        return (currentIndex + 1u) % hosts.size();
    }();

    const bool focused = hosts[nextIndex].focus();
#ifdef ENABLE_TESTS
    if (focused)
    {
        RememberPluginConfigDebugFocusedHost(*state, hosts[nextIndex].host, nextIndex);
    }
    TracePluginConfigDebug(
        std::format(L"plugin-config advance: currentHost={:#x} currentIndex={} nextIndex={} nextHost={:#x} reverse={} focused={} finalFocus={:#x}",
                    reinterpret_cast<uintptr_t>(hostHwnd),
                    currentIndex < hosts.size() ? static_cast<unsigned long long>(currentIndex) : ~0ull,
                    static_cast<unsigned long long>(nextIndex),
                    reinterpret_cast<uintptr_t>(hosts[nextIndex].host),
                    reverse ? 1 : 0,
                    focused ? 1 : 0,
                    reinterpret_cast<uintptr_t>(GetFocus())));
#endif
    return focused;
}

[[nodiscard]] bool AttachDxCommandButtonHost(HWND dlg, DxCommandButtonHost& slot, const UINT commandId, const bool primary) noexcept
{
    wil::unique_hwnd hwnd(CreateDxCommandButtonHostWindow(dlg));
    if (! hwnd || ! slot.host.Attach(hwnd.get()))
    {
        return false;
    }

    if (! SetPropW(hwnd.get(), kPluginConfigDxButtonHostStateProp, reinterpret_cast<HANDLE>(&slot)) ||
        ! InstallWndProcHook(hwnd.get(), kPluginConfigDxButtonHostOriginalWndProcProp, PluginConfigDxButtonHostWndProc))
    {
        RemovePropW(hwnd.get(), kPluginConfigDxButtonHostStateProp);
        Debug::ErrorWithLastError(L"PluginConfig: failed to install owned WNDPROC hook for DX command-button host.");
        slot.host.Detach();
        return false;
    }

    auto button    = std::make_unique<Button>();
    slot.button    = button.get();
    slot.commandId = commandId;
    slot.button->SetPrimary(primary);
    slot.host.SetOnTabBoundary(
        std::function<bool(bool)>{[hostHwnd = hwnd.get()](bool reverse) -> bool { return MovePluginConfigDialogTabFocusFromHost(hostHwnd, reverse); }});
    slot.button->SetOnClick([dlg, commandId]
    {
        if (dlg && IsWindow(dlg) != FALSE)
        {
            PostMessageW(dlg, WM_COMMAND, MAKEWPARAM(commandId, 0), 0);
        }
    });
    slot.host.SetRoot(std::move(button));
    slot.hostHwnd = std::move(hwnd);
    return true;
}

void ApplyDxCommandButtonTheme(PluginConfigDialogState& state) noexcept
{
    if (! state.usesDxUiCommandButtons)
    {
        return;
    }

    const ThemePalette palette = MakeAppThemeDxPalette(state.theme, state.theme.windowBackground);
    for (DxCommandButtonHost& slot : state.dxCommandButtons)
    {
        slot.host.SetTheme(palette);
    }
}

void UpdateDxCommandButtons(const HWND dlg, PluginConfigDialogState& state) noexcept
{
    if (! dlg || ! state.usesDxUiCommandButtons)
    {
        return;
    }

    const struct Entry
    {
        DxCommandButtonIndex index;
        UINT commandId;
        UINT textId;
        bool primary;
    } entries[] = {
        {DxCommandButtonIndex::Ok, IDOK, IDS_BTN_OK, true},
        {DxCommandButtonIndex::Cancel, IDCANCEL, IDS_BTN_CANCEL, false},
    };

    for (const Entry& entry : entries)
    {
        DxCommandButtonHost& slot = state.dxCommandButtons[static_cast<size_t>(entry.index)];
        if (! slot.button)
        {
            continue;
        }

        const HWND legacy = GetLegacyCommandButton(dlg, entry.index);
        slot.commandId    = entry.commandId;
        slot.button->SetPrimary(entry.primary);
        slot.button->SetText(LoadStringResource(nullptr, entry.textId));
        slot.button->SetMnemonic(entry.commandId == IDOK ? L'O' : L'C');
        slot.button->SetEnabled(! legacy || IsWindowEnabled(legacy) != FALSE);
        slot.host.Invalidate();
    }
}

[[nodiscard]] bool HandlePluginConfigDialogDxMnemonic(PluginConfigDialogState& state, const wchar_t mnemonic) noexcept
{
    const auto tryFieldSlot = [mnemonic](PluginConfigDxHostSlot& slot) noexcept
    { return slot.hostHwnd && slot.host && IsWindowVisible(slot.hostHwnd.get()) != FALSE && slot.host->HandleMnemonic(mnemonic); };

    for (PluginConfigFieldControls& controls : state.controls)
    {
        if (tryFieldSlot(controls.dxLabelSlot) || tryFieldSlot(controls.dxEditSlot) || tryFieldSlot(controls.dxComboSlot) ||
            tryFieldSlot(controls.dxToggleSlot))
        {
            return true;
        }

        for (PluginConfigChoiceDxControl& choice : controls.dxChoiceControls)
        {
            if (tryFieldSlot(choice.slot))
            {
                return true;
            }
        }
    }

    for (DxCommandButtonHost& slot : state.dxCommandButtons)
    {
        if (slot.hostHwnd && IsWindowVisible(slot.hostHwnd.get()) != FALSE && slot.host.HandleMnemonic(mnemonic))
        {
            return true;
        }
    }

    return false;
}

void PositionDxCommandButtonHost(HWND dlg, const HWND legacy, const DxCommandButtonHost& slot, const bool visible) noexcept
{
    if (! dlg || ! slot.hostHwnd)
    {
        return;
    }

    if (! legacy)
    {
        ShowWindow(slot.hostHwnd.get(), SW_HIDE);
        return;
    }

    RECT rect{};
    if (GetWindowRect(legacy, &rect) == FALSE)
    {
        ShowWindow(slot.hostHwnd.get(), SW_HIDE);
        return;
    }

    MapWindowPoints(HWND_DESKTOP, dlg, reinterpret_cast<POINT*>(&rect), 2);
    SetWindowPos(slot.hostHwnd.get(),
                 nullptr,
                 rect.left,
                 rect.top,
                 rect.right - rect.left,
                 rect.bottom - rect.top,
                 SWP_NOZORDER | SWP_NOACTIVATE | static_cast<UINT>(visible ? SWP_SHOWWINDOW : SWP_HIDEWINDOW));
}

void DetachDxCommandButtons(PluginConfigDialogState& state) noexcept
{
    for (DxCommandButtonHost& slot : state.dxCommandButtons)
    {
        slot.Detach();
    }
    state.usesDxUiCommandButtons = false;
}

void DetachDxFieldHosts(PluginConfigDialogState& state) noexcept
{
    for (PluginConfigFieldControls& controls : state.controls)
    {
        controls.dxLabelSlot.Detach();
        controls.dxDescriptionSlot.Detach();
        controls.dxDefaultsSlot.Detach();
        controls.dxEditSlot.Detach();
        controls.dxComboSlot.Detach();
        controls.dxToggleSlot.Detach();
        for (PluginConfigChoiceDxControl& choice : controls.dxChoiceControls)
        {
            choice.slot.Detach();
            choice.checkbox = nullptr;
        }

        controls.dxLabelControl       = nullptr;
        controls.dxDescriptionControl = nullptr;
        controls.dxDefaultsControl    = nullptr;
        controls.dxEditControl        = nullptr;
        controls.dxComboControl       = nullptr;
        controls.dxToggleControl      = nullptr;
    }
}

[[nodiscard]] HWND DetachPluginConfigPanel(PluginConfigDialogState& state) noexcept
{
    if (! state.panel)
    {
        return nullptr;
    }

    const HWND panel = state.panel;
    if (IsWindow(state.panel) != FALSE)
    {
        RemovePropW(state.panel, kPluginConfigPanelStateProp);
        RestoreWndProcHook(state.panel, kPluginConfigPanelOriginalWndProcProp);
    }
    state.panel = nullptr;
    return panel;
}

void LayoutPluginConfigDialog(HWND dlg, PluginConfigDialogState& state) noexcept
{
    if (! dlg)
    {
        return;
    }

    HWND panel  = state.panel ? state.panel : GetDlgItem(dlg, IDC_PLUGIN_CONFIG_PLACEHOLDER);
    HWND ok     = GetDlgItem(dlg, IDOK);
    HWND cancel = GetDlgItem(dlg, IDCANCEL);
    if (! panel || ! ok || ! cancel)
    {
        return;
    }

    RECT client{};
    GetClientRect(dlg, &client);

    const UINT dpi   = GetDpiForWindow(dlg);
    const int margin = ScaleDip(dpi, 8);
    const int gapX   = ScaleDip(dpi, 8);

    RECT okRect{};
    RECT cancelRect{};
    GetWindowRect(ok, &okRect);
    GetWindowRect(cancel, &cancelRect);
    MapWindowPoints(nullptr, dlg, reinterpret_cast<POINT*>(&okRect), 2);
    MapWindowPoints(nullptr, dlg, reinterpret_cast<POINT*>(&cancelRect), 2);

    const int okWidth      = std::max(0l, okRect.right - okRect.left);
    const int okHeight     = std::max(0l, okRect.bottom - okRect.top);
    const int cancelWidth  = std::max(0l, cancelRect.right - cancelRect.left);
    const int cancelHeight = std::max(0l, cancelRect.bottom - cancelRect.top);
    const int buttonHeight = std::max(okHeight, cancelHeight);

    const int cancelLeft = std::max(0, static_cast<int>(client.right) - margin - cancelWidth);
    const int buttonsTop = std::max(0, static_cast<int>(client.bottom) - margin - buttonHeight);
    const int okLeft     = std::max(0, cancelLeft - gapX - okWidth);

    SetWindowPos(cancel, nullptr, cancelLeft, buttonsTop, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
    SetWindowPos(ok, nullptr, okLeft, buttonsTop, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
    ShowWindow(ok, state.usesDxUiCommandButtons ? SW_HIDE : SW_SHOWNA);
    ShowWindow(cancel, state.usesDxUiCommandButtons ? SW_HIDE : SW_SHOWNA);
    PositionDxCommandButtonHost(dlg, ok, state.dxCommandButtons[static_cast<size_t>(DxCommandButtonIndex::Ok)], state.usesDxUiCommandButtons);
    PositionDxCommandButtonHost(dlg, cancel, state.dxCommandButtons[static_cast<size_t>(DxCommandButtonIndex::Cancel)], state.usesDxUiCommandButtons);

    const int panelLeft   = margin;
    const int panelTop    = margin;
    const int panelWidth  = std::max(0, static_cast<int>(client.right) - (2 * margin));
    const int panelBottom = std::max(panelTop, buttonsTop - margin);
    const int panelHeight = std::max(0, panelBottom - panelTop);

    SetWindowPos(panel, nullptr, panelLeft, panelTop, panelWidth, panelHeight, SWP_NOZORDER | SWP_NOACTIVATE);
}

LRESULT CALLBACK PluginConfigDxButtonHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    auto* host = reinterpret_cast<DxCommandButtonHost*>(GetPropW(hwnd, kPluginConfigDxButtonHostStateProp));
    if (! host)
    {
        return CallStoredWndProc(hwnd, kPluginConfigDxButtonHostOriginalWndProcProp, msg, wp, lp);
    }

    if (msg == WM_GETDLGCODE && (wp == VK_TAB || wp == VK_ESCAPE))
    {
        return CallStoredWndProc(hwnd, kPluginConfigDxButtonHostOriginalWndProcProp, msg, wp, lp) | (wp == VK_TAB ? DLGC_WANTTAB : DLGC_WANTALLKEYS);
    }

    if (msg == WM_KEYDOWN && wp == VK_TAB)
    {
        const bool reverse = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
        if (MovePluginConfigDialogTabFocusFromHost(hwnd, reverse))
        {
            return 0;
        }
    }

    if (msg == WM_KEYDOWN && wp == VK_ESCAPE && TryRoutePluginConfigDialogCommandKey(hwnd, wp))
    {
        return 0;
    }

    if (msg == WM_NCDESTROY)
    {
        const auto originalWndProc = RedSalamander::Win32Callback::GetStoredWndProc(hwnd, kPluginConfigDxButtonHostOriginalWndProcProp);
        RemovePropW(hwnd, kPluginConfigDxButtonHostStateProp);
        RedSalamander::Win32Callback::RestoreWndProcHook(hwnd, kPluginConfigDxButtonHostOriginalWndProcProp, PluginConfigDxButtonHostWndProc);
        host->host.Detach();
        host->host.ReleaseMouseCapture();
        host->hostHwnd.release();
        return originalWndProc ? RedSalamander::Win32Callback::CallWindowProcNoThrow(originalWndProc, hwnd, msg, wp, lp) : DefWindowProcW(hwnd, msg, wp, lp);
    }

    bool handled           = false;
    const LRESULT dxResult = host->host.HandleMessage(hwnd, msg, wp, lp, handled);

    return handled ? dxResult : CallStoredWndProc(hwnd, kPluginConfigDxButtonHostOriginalWndProcProp, msg, wp, lp);
}

LRESULT CALLBACK PluginConfigDxFieldHostWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    auto* slot = reinterpret_cast<PluginConfigDxHostSlot*>(GetPropW(hwnd, kPluginConfigDxFieldHostStateProp));
    if (! slot || ! slot->host)
    {
        return CallStoredWndProc(hwnd, kPluginConfigDxFieldHostOriginalWndProcProp, msg, wp, lp);
    }

    if (msg == WM_GETDLGCODE && (wp == VK_TAB || wp == VK_RETURN || wp == VK_ESCAPE))
    {
        return CallStoredWndProc(hwnd, kPluginConfigDxFieldHostOriginalWndProcProp, msg, wp, lp) | (wp == VK_TAB ? DLGC_WANTTAB : DLGC_WANTALLKEYS);
    }

    if (msg == WM_KEYDOWN && wp == VK_TAB)
    {
        const bool reverse = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
        if (MovePluginConfigDialogTabFocusFromHost(hwnd, reverse))
        {
            return 0;
        }
    }

    if (msg == WM_KEYDOWN && (wp == VK_RETURN || wp == VK_ESCAPE) && TryRoutePluginConfigDialogCommandKey(hwnd, wp))
    {
        return 0;
    }

    if (msg == WM_NCDESTROY)
    {
        const auto originalWndProc = RedSalamander::Win32Callback::GetStoredWndProc(hwnd, kPluginConfigDxFieldHostOriginalWndProcProp);
        RemovePropW(hwnd, kPluginConfigDxFieldHostStateProp);
        RedSalamander::Win32Callback::RestoreWndProcHook(hwnd, kPluginConfigDxFieldHostOriginalWndProcProp, PluginConfigDxFieldHostWndProc);
        slot->host->Detach();
        slot->host->ReleaseMouseCapture();
        if (slot->hostHwnd && slot->hostHwnd.get() == hwnd)
        {
            slot->hostHwnd.release();
        }
        return originalWndProc ? RedSalamander::Win32Callback::CallWindowProcNoThrow(originalWndProc, hwnd, msg, wp, lp) : DefWindowProcW(hwnd, msg, wp, lp);
    }

    bool handled           = false;
    const LRESULT dxResult = slot->host->HandleMessage(hwnd, msg, wp, lp, handled);

    return handled ? dxResult : CallStoredWndProc(hwnd, kPluginConfigDxFieldHostOriginalWndProcProp, msg, wp, lp);
}

void ScrollPanelTo(HWND panel, PluginConfigDialogState& state, int newScrollPosY) noexcept
{
    if (! panel)
    {
        return;
    }

    RECT client{};
    GetClientRect(panel, &client);
    const int clientHeight = std::max(0l, client.bottom - client.top);
    const int maxScroll    = std::max(0, state.contentHeight - clientHeight);

    newScrollPosY   = std::clamp(newScrollPosY, 0, maxScroll);
    const int delta = newScrollPosY - state.scrollPosY;
    if (delta == 0)
    {
        return;
    }

    state.scrollPosY = newScrollPosY;

    ScrollWindowEx(panel, 0, -delta, nullptr, nullptr, nullptr, nullptr, SW_INVALIDATE | SW_ERASE | SW_SCROLLCHILDREN);
    SetScrollPos(panel, SB_VERT, state.scrollPosY, TRUE);
    UpdateWindow(panel);
}

void UpdatePanelScrollInfo(HWND panel, PluginConfigDialogState& state) noexcept
{
    if (! panel)
    {
        return;
    }

    RECT client{};
    GetClientRect(panel, &client);
    const int clientHeight = std::max(0l, client.bottom - client.top);

    SCROLLINFO si{};
    si.cbSize = sizeof(si);
    si.fMask  = SIF_RANGE | SIF_PAGE | SIF_POS;
    si.nMin   = 0;
    si.nMax   = std::max(0, state.contentHeight - 1);
    si.nPage  = static_cast<UINT>(std::max(0, clientHeight));
    si.nPos   = std::clamp(state.scrollPosY, 0, std::max(0, state.contentHeight - clientHeight));
    SetScrollInfo(panel, SB_VERT, &si, TRUE);

    ShowScrollBar(panel, SB_VERT, state.contentHeight > clientHeight);

    if (state.scrollPosY != si.nPos)
    {
        ScrollPanelTo(panel, state, si.nPos);
    }
}

void EnsurePanelChildVisible(HWND panel, PluginConfigDialogState& state, HWND child) noexcept
{
    if (! panel || ! child)
    {
        return;
    }

    RECT client{};
    GetClientRect(panel, &client);
    const int clientHeight = std::max(0l, client.bottom - client.top);
    if (clientHeight <= 0)
    {
        return;
    }

    RECT childRect{};
    if (! GetWindowRect(child, &childRect))
    {
        return;
    }

    MapWindowPoints(nullptr, panel, reinterpret_cast<POINT*>(&childRect), 2);

    const UINT dpi       = GetDpiForWindow(panel);
    const int padY       = ScaleDip(dpi, 8);
    const int viewTop    = std::max(0, static_cast<int>(client.top) + padY);
    const int viewBottom = std::max(viewTop, static_cast<int>(client.bottom) - padY);

    if (childRect.top < viewTop)
    {
        const int delta = childRect.top - viewTop;
        ScrollPanelTo(panel, state, state.scrollPosY + delta);
        UpdatePanelScrollInfo(panel, state);
        return;
    }

    if (childRect.bottom > viewBottom)
    {
        const int delta = childRect.bottom - viewBottom;
        ScrollPanelTo(panel, state, state.scrollPosY + delta);
        UpdatePanelScrollInfo(panel, state);
        return;
    }
}

PluginConfigFieldControls* FindToggleControls(PluginConfigDialogState& state, HWND toggle) noexcept
{
    if (! toggle)
    {
        return nullptr;
    }

    for (auto& c : state.controls)
    {
        if (c.hToggle == toggle)
        {
            return std::addressof(c);
        }
    }

    return nullptr;
}

LRESULT CALLBACK PluginConfigPanelWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) noexcept
{
    auto* state = reinterpret_cast<PluginConfigDialogState*>(GetPropW(hwnd, kPluginConfigPanelStateProp));
    if (! state)
    {
        return CallStoredWndProc(hwnd, kPluginConfigPanelOriginalWndProcProp, msg, wp, lp);
    }

    switch (msg)
    {
        case WM_ERASEBKGND:
        {
            RECT rc{};
            GetClientRect(hwnd, &rc);
            FillRect(reinterpret_cast<HDC>(wp), &rc, state->backgroundBrush.get());
            return 1;
        }
        case WM_PAINT:
        {
            PAINTSTRUCT ps{};
            wil::unique_hdc_paint hdc = wil::BeginPaint(hwnd, &ps);
            if (hdc)
            {
                FillRect(hdc.get(), &ps.rcPaint, state->backgroundBrush.get());
            }
            return 0;
        }
        case WM_SETFOCUS:
        {
            // Move focus into the first input control when tabbing into the scroll panel.
            const bool backwards = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
            HWND first           = GetNextDlgTabItem(hwnd, nullptr, backwards ? TRUE : FALSE);
            if (first)
            {
                SetFocus(first);
            }
            return 0;
        }
        case WM_CTLCOLORSTATIC: return OnPluginConfigDialogCtlColorStatic(state, reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_CTLCOLORBTN: return OnPluginConfigDialogCtlColorButton(state, reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_CTLCOLOREDIT: return OnPluginConfigDialogCtlColorEdit(state, reinterpret_cast<HDC>(wp));
        case WM_CTLCOLORLISTBOX: return OnPluginConfigDialogCtlColorListBox(state, reinterpret_cast<HDC>(wp));
        case WM_COMMAND:
        {
            const UINT notify = HIWORD(wp);
            if (notify == BN_SETFOCUS || notify == EN_SETFOCUS || notify == CBN_SETFOCUS)
            {
                HWND focusedControl = reinterpret_cast<HWND>(lp);
                if (focusedControl)
                {
                    EnsurePanelChildVisible(hwnd, *state, focusedControl);
                }
            }

            if (HIWORD(wp) == BN_CLICKED)
            {
                HWND clicked = reinterpret_cast<HWND>(lp);
                if (auto* controls = FindToggleControls(*state, clicked))
                {
                    if (clicked && (IsWindowVisible(clicked) == FALSE || UsesLegacyOwnerDrawButton(clicked)))
                    {
                        SetLegacyToggleState(clicked, ! GetLegacyToggleState(clicked));
                    }
                    return 0;
                }
            }
            break;
        }
        case WM_SIZE:
        {
            UpdatePanelScrollInfo(hwnd, *state);
            break;
        }
        case WM_VSCROLL:
        {
            const UINT action  = LOWORD(wp);
            const UINT dpi     = GetDpiForWindow(hwnd);
            const int lineStep = std::max(1, ScaleDip(dpi, 18));

            SCROLLINFO si{};
            si.cbSize = sizeof(si);
            si.fMask  = SIF_ALL;
            if (! GetScrollInfo(hwnd, SB_VERT, &si))
            {
                break;
            }

            int newPos = si.nPos;
            switch (action)
            {
                case SB_LINEUP: newPos -= lineStep; break;
                case SB_LINEDOWN: newPos += lineStep; break;
                case SB_PAGEUP: newPos -= static_cast<int>(si.nPage); break;
                case SB_PAGEDOWN: newPos += static_cast<int>(si.nPage); break;
                case SB_THUMBPOSITION:
                case SB_THUMBTRACK: newPos = si.nTrackPos; break;
                case SB_TOP: newPos = si.nMin; break;
                case SB_BOTTOM: newPos = si.nMax; break;
            }

            ScrollPanelTo(hwnd, *state, newPos);
            UpdatePanelScrollInfo(hwnd, *state);
            return 0;
        }
        case WM_MOUSEWHEEL:
        {
            const int wheelDelta = GET_WHEEL_DELTA_WPARAM(wp);
            if (wheelDelta == 0)
            {
                break;
            }

            UINT lines = 3;
            SystemParametersInfoW(SPI_GETWHEELSCROLLLINES, 0, &lines, 0);
            if (lines == WHEEL_PAGESCROLL)
            {
                lines = 3;
            }

            const UINT dpi     = GetDpiForWindow(hwnd);
            const int lineStep = std::max(1, ScaleDip(dpi, 18));
            const int steps    = wheelDelta / WHEEL_DELTA;
            if (steps != 0)
            {
                ScrollPanelTo(hwnd, *state, state->scrollPosY - (steps * static_cast<int>(lines) * lineStep));
                UpdatePanelScrollInfo(hwnd, *state);
                return 0;
            }
            break;
        }
    }

    if (msg == WM_NCDESTROY)
    {
        const auto originalWndProc = RedSalamander::Win32Callback::GetStoredWndProc(hwnd, kPluginConfigPanelOriginalWndProcProp);
        RemovePropW(hwnd, kPluginConfigPanelStateProp);
        RedSalamander::Win32Callback::RestoreWndProcHook(hwnd, kPluginConfigPanelOriginalWndProcProp, PluginConfigPanelWndProc);
        return originalWndProc ? RedSalamander::Win32Callback::CallWindowProcNoThrow(originalWndProc, hwnd, msg, wp, lp) : DefWindowProcW(hwnd, msg, wp, lp);
    }

    return CallStoredWndProc(hwnd, kPluginConfigPanelOriginalWndProcProp, msg, wp, lp);
}

Common::PluginConfiguration::SchemaParseResult ParseConfigurationSchema(std::string_view schemaJsonUtf8) noexcept
{
    return Common::PluginConfiguration::ParseSchema(schemaJsonUtf8);
}
std::string BuildConfigurationJson(const std::vector<PluginConfigFieldControls>& controls, std::string_view originalConfigurationJsonUtf8)
{
    std::vector<PluginConfigField> fields;
    std::vector<Common::PluginConfiguration::FieldValue> values;
    fields.reserve(controls.size());
    values.reserve(controls.size());

    for (const PluginConfigFieldControls& controlsForField : controls)
    {
        const PluginConfigField& field = controlsForField.field;
        fields.push_back(field);

        Common::PluginConfiguration::FieldValue value;
        value.type = field.type;
        if (field.type == PluginConfigFieldType::Text)
        {
            value.text = field.defaultText;
            if (controlsForField.hEdit)
            {
                const int length = GetWindowTextLengthW(controlsForField.hEdit);
                value.text.clear();
                if (length > 0)
                {
                    value.text.resize(static_cast<size_t>(length) + 1u);
                    GetWindowTextW(controlsForField.hEdit, value.text.data(), length + 1);
                    value.text.resize(static_cast<size_t>(length));
                }
            }
        }
        else if (field.type == PluginConfigFieldType::Value)
        {
            value.integer = field.defaultInt;
            if (controlsForField.hEdit)
            {
                std::array<wchar_t, 64> buffer{};
                GetWindowTextW(controlsForField.hEdit, buffer.data(), static_cast<int>(buffer.size()));
                wchar_t* end           = nullptr;
                errno                  = 0;
                const long long parsed = wcstoll(buffer.data(), &end, 10);
                if (errno == 0 && end != buffer.data())
                {
                    value.integer = static_cast<int64_t>(parsed);
                }
            }
        }
        else if (field.type == PluginConfigFieldType::Bool)
        {
            value.boolean = field.defaultBool;
            if (controlsForField.hToggle)
            {
                value.boolean = GetLegacyToggleState(controlsForField.hToggle);
            }
            else if (! controlsForField.choiceButtons.empty())
            {
                value.boolean = SendMessageW(controlsForField.choiceButtons.front(), BM_GETCHECK, 0, 0) == BST_CHECKED;
            }
        }
        else if (field.type == PluginConfigFieldType::Option)
        {
            value.text = field.defaultOption;
            if (controlsForField.hToggle)
            {
                const bool isOn    = GetLegacyToggleState(controlsForField.hToggle);
                const size_t index = isOn ? controlsForField.toggleOnChoiceIndex : controlsForField.toggleOffChoiceIndex;
                if (index < field.choices.size())
                {
                    value.text = field.choices[index].value;
                }
            }
            else if (controlsForField.hCombo)
            {
                const LRESULT index = SendMessageW(controlsForField.hCombo, CB_GETCURSEL, 0, 0);
                if (index >= 0 && index <= static_cast<LRESULT>((std::numeric_limits<int>::max)()))
                {
                    const size_t choiceIndex = static_cast<size_t>(index);
                    if (choiceIndex < field.choices.size())
                    {
                        value.text = field.choices[choiceIndex].value;
                    }
                }
            }
            else
            {
                for (size_t index = 0; index < controlsForField.choiceButtons.size() && index < field.choices.size(); ++index)
                {
                    if (SendMessageW(controlsForField.choiceButtons[index], BM_GETCHECK, 0, 0) == BST_CHECKED)
                    {
                        value.text = field.choices[index].value;
                        break;
                    }
                }
            }
        }
        else if (field.type == PluginConfigFieldType::Selection)
        {
            value.selection = field.defaultSelection;
            if (! controlsForField.choiceButtons.empty())
            {
                value.selection.clear();
                for (size_t index = 0; index < controlsForField.choiceButtons.size() && index < field.choices.size(); ++index)
                {
                    if (SendMessageW(controlsForField.choiceButtons[index], BM_GETCHECK, 0, 0) == BST_CHECKED)
                    {
                        value.selection.push_back(field.choices[index].value);
                    }
                }
            }
        }
        values.push_back(std::move(value));
    }

    std::string serialized;
    if (FAILED(Common::PluginConfiguration::SerializeConfiguration(originalConfigurationJsonUtf8, fields, values, serialized)))
    {
        Debug::Error(L"Failed to serialize plugin configuration through the shared codec.");
        return {};
    }
    return serialized;
}

void ApplyFieldValueToControls(const PluginConfigField& field,
                               const Common::PluginConfiguration::FieldValue& value,
                               PluginConfigFieldControls& out)
{
    out.field = field;
    switch (field.type)
    {
        case PluginConfigFieldType::Text: out.field.defaultText = value.text; break;
        case PluginConfigFieldType::Value: out.field.defaultInt = value.integer; break;
        case PluginConfigFieldType::Bool: out.field.defaultBool = value.boolean; break;
        case PluginConfigFieldType::Option: out.field.defaultOption = value.text; break;
        case PluginConfigFieldType::Selection: out.field.defaultSelection = value.selection; break;
    }
}

INT_PTR OnPluginConfigDialogInit(HWND dlg, PluginConfigDialogState* state)
{
    if (! state)
    {
        return FALSE;
    }

    SetWindowLongPtrW(dlg, DWLP_USER, reinterpret_cast<LONG_PTR>(state));
    SettingsHotReload::RegisterParticipant(dlg);

    ApplyTitleBarTheme(dlg, state->theme, GetActiveWindow() == dlg);
    ApplyWindowBackdropTheme(dlg, state->theme, WindowBackdropTarget::Tool);

    state->backgroundBrush.reset(CreateSolidBrush(state->theme.windowBackground));
    state->inputBackgroundColor = GetControlSurfaceColor(state->theme);
    state->inputBrush.reset();
    if (! state->theme.highContrast)
    {
        state->inputBrush.reset(CreateSolidBrush(state->inputBackgroundColor));
    }
    state->inputFrameStyle.theme                       = &state->theme;
    state->inputFrameStyle.backdropBrush               = state->backgroundBrush.get();
    state->inputFrameStyle.inputBackgroundColor        = state->inputBackgroundColor;
    state->inputFrameStyle.inputFocusedBackgroundColor = state->inputBackgroundColor;
    state->inputFrameStyle.inputDisabledBackgroundColor =
        UiMetrics::BlendColorRefWeightedTruncate(state->theme.windowBackground, state->inputBackgroundColor, state->theme.dark ? 70 : 40, 255);
    state->contentHeight = 0;
    state->scrollPosY    = 0;

    RECT dlgRect{};
    if (GetWindowRect(dlg, &dlgRect))
    {
        state->fixedWindowWidthPx = std::max(0l, dlgRect.right - dlgRect.left);
    }

    if (! state->pluginName.empty())
    {
        SetWindowTextW(dlg, state->pluginName.c_str());
    }

    const bool dxCommandButtonsAttached =
        kEnableManagePluginsDxCommandButtons &&
        AttachDxCommandButtonHost(dlg, state->dxCommandButtons[static_cast<size_t>(DxCommandButtonIndex::Ok)], IDOK, true) &&
        AttachDxCommandButtonHost(dlg, state->dxCommandButtons[static_cast<size_t>(DxCommandButtonIndex::Cancel)], IDCANCEL, false);
    state->usesDxUiCommandButtons = dxCommandButtonsAttached;
    if (! dxCommandButtonsAttached)
    {
        DetachDxCommandButtons(*state);
    }
    else
    {
        ApplyDxCommandButtonTheme(*state);
        UpdateDxCommandButtons(dlg, *state);
    }

    state->panel = GetDlgItem(dlg, IDC_PLUGIN_CONFIG_PLACEHOLDER);
    if (state->panel)
    {
        LONG_PTR exStyle = GetWindowLongPtrW(state->panel, GWL_EXSTYLE);
        if ((exStyle & WS_EX_CONTROLPARENT) == 0)
        {
            exStyle |= WS_EX_CONTROLPARENT;
            SetWindowLongPtrW(state->panel, GWL_EXSTYLE, exStyle);
        }

        LONG_PTR style = GetWindowLongPtrW(state->panel, GWL_STYLE);
        if ((style & WS_TABSTOP) == 0 || (style & SS_NOTIFY) == 0)
        {
            style |= WS_TABSTOP | SS_NOTIFY;
            SetWindowLongPtrW(state->panel, GWL_STYLE, style);
            SetWindowPos(state->panel, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
        }

        const bool darkBackground = ChooseContrastingTextColor(state->theme.windowBackground) == RGB(255, 255, 255);
        const wchar_t* panelTheme = state->theme.highContrast ? L"" : (darkBackground ? L"DarkMode_Explorer" : L"Explorer");
        SetWindowTheme(state->panel, panelTheme, nullptr);
        SendMessageW(state->panel, WM_THEMECHANGED, 0, 0);

        if (! SetPropW(state->panel, kPluginConfigPanelStateProp, reinterpret_cast<HANDLE>(state)) ||
            ! InstallWndProcHook(state->panel, kPluginConfigPanelOriginalWndProcProp, PluginConfigPanelWndProc))
        {
            RemovePropW(state->panel, kPluginConfigPanelStateProp);
            Debug::ErrorWithLastError(L"PluginConfig: failed to install owned WNDPROC hook for scroll panel.");
        }
    }

    LayoutPluginConfigDialog(dlg, *state);

    const Common::PluginConfiguration::SchemaParseResult schema = ParseConfigurationSchema(state->schemaJsonUtf8);
    const std::vector<PluginConfigField>& fields                 = schema.fields;
    const Common::PluginConfiguration::ConfigurationParseResult configuration =
        Common::PluginConfiguration::ParseConfiguration(fields, state->configurationJsonUtf8);

    state->controls.clear();
    state->controls.reserve(fields.size());

    HWND panel = state->panel;
    if (! panel)
    {
        return TRUE;
    }

    const UINT dpi = GetDpiForWindow(dlg);

    int margin = ScaleDip(dpi, 8);

    int spacingY = ScaleDip(dpi, 10);

    int labelOffsetY = ScaleDip(dpi, 3);

    int labelGapX = ScaleDip(dpi, 10);

    int labelHeight = std::max(1, ScaleDip(dpi, 18));

    int editHeight = std::max(1, ScaleDip(dpi, 26));

    int optionHeight = std::max(1, ScaleDip(dpi, 20));

    int minControlWidth = ScaleDip(dpi, 80);

    RECT panelRect{};
    GetClientRect(panel, &panelRect);
    int panelWidth = std::max(0l, panelRect.right - panelRect.left);

    // Reserve space for the vertical scrollbar. ShowScrollBar does not shrink the client area for us
    // (see e.g. ViewerImgRaw), so controls would otherwise draw under the scrollbar when content overflows.
    const int scrollW = GetSystemMetricsForDpi(SM_CXVSCROLL, dpi);
    panelWidth        = std::max(0, panelWidth - scrollW);

    [[maybe_unused]] const int contentWidth = std::max(0, panelWidth - (2 * margin));

    const int left  = margin;
    const int top   = margin;
    const int right = std::max(0, panelWidth - margin);

    const int availableWidth = std::max(0, right - left);

    int labelWidth          = (availableWidth * 2) / 5;
    const int minLabelWidth = ScaleDip(dpi, 110);
    labelWidth              = std::clamp(labelWidth, minLabelWidth, std::max(0, availableWidth - minControlWidth));

    const int labelTextWidth = std::max(0, labelWidth - labelGapX);

    const int controlX     = left + labelWidth;
    const int controlWidth = std::max(minControlWidth, right - controlX);

    int y = top;

    for (size_t fieldIndex = 0; fieldIndex < fields.size(); ++fieldIndex)
    {
        const PluginConfigField& field = fields[fieldIndex];
        PluginConfigFieldControls controls;
        ApplyFieldValueToControls(field, configuration.values[fieldIndex], controls);

        if (controls.field.uiHidden)
        {
            state->controls.push_back(std::move(controls));
            continue;
        }

        controls.hLabel = CreateWindowExW(0,
                                          L"Static",
                                          controls.field.label.c_str(),
                                          WS_CHILD | WS_VISIBLE | SS_NOPREFIX | SS_WORDELLIPSIS,
                                          left,
                                          y + labelOffsetY,
                                          labelTextWidth,
                                          labelHeight,
                                          panel,
                                          nullptr,
                                          GetModuleHandleW(nullptr),
                                          nullptr);

        if (controls.field.type == PluginConfigFieldType::Text || controls.field.type == PluginConfigFieldType::Value)
        {
            const int valueWidth     = std::min(controlWidth, ScaleDip(dpi, 140));
            const int editFrameWidth = (controls.field.type == PluginConfigFieldType::Value) ? valueWidth : controlWidth;

            const bool customFrames = ! state->theme.highContrast;
            const int framePadding  = ScaleDip(dpi, 2);
            const int textMargin    = ScaleDip(dpi, 6);

            DWORD editStyle = WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL;
            if (controls.field.type == PluginConfigFieldType::Value)
            {
                editStyle |= ES_NUMBER;
            }

            if (customFrames)
            {
                controls.hEditFrame = CreateWindowExW(
                    0, L"Static", L"", WS_CHILD | WS_VISIBLE, controlX, y, editFrameWidth, editHeight, panel, nullptr, GetModuleHandleW(nullptr), nullptr);

                controls.hEdit = CreateWindowExW(0,
                                                 L"Edit",
                                                 L"",
                                                 editStyle,
                                                 controlX + framePadding,
                                                 y + framePadding,
                                                 std::max(1, editFrameWidth - (2 * framePadding)),
                                                 std::max(1, editHeight - (2 * framePadding)),
                                                 panel,
                                                 nullptr,
                                                 GetModuleHandleW(nullptr),
                                                 nullptr);
                if (controls.hEdit)
                {
                    SendMessageW(controls.hEdit, EM_SETMARGINS, EC_LEFTMARGIN | EC_RIGHTMARGIN, MAKELPARAM(textMargin, textMargin));
                    ThemedInputFrames::InstallFrame(controls.hEditFrame, controls.hEdit, &state->inputFrameStyle);
                }
            }
            else
            {
                controls.hEdit = CreateWindowExW(
                    WS_EX_CLIENTEDGE, L"Edit", L"", editStyle, controlX, y, editFrameWidth, editHeight, panel, nullptr, GetModuleHandleW(nullptr), nullptr);
                if (controls.hEdit)
                {
                    ThemedInputFrames::InstallControl(controls.hEdit);
                }
            }
            if (controls.hEdit)
            {
                if (controls.field.type == PluginConfigFieldType::Text)
                {
                    SetWindowTextW(controls.hEdit, controls.field.defaultText.c_str());
                }
                else
                {
                    const std::wstring text = std::to_wstring(controls.field.defaultInt);
                    SetWindowTextW(controls.hEdit, text.c_str());
                }
            }

            y += editHeight + spacingY;
        }
        else if (controls.field.type == PluginConfigFieldType::Bool)
        {
            if (! state->theme.highContrast)
            {
                controls.toggleOnChoiceIndex  = 0;
                controls.toggleOffChoiceIndex = 1;

                const int paddingX   = ScaleDip(dpi, 6);
                const int gapX       = ScaleDip(dpi, 6);
                const int trackWidth = ScaleDip(dpi, 28);

                const std::wstring_view leftLabel  = L"True";
                const std::wstring_view rightLabel = L"False";

                const int leftWidth     = Typography::MeasureSingleLineTextWidthPx(panel, FontRole::BodyStrong, leftLabel);
                const int rightWidth    = Typography::MeasureSingleLineTextWidthPx(panel, FontRole::BodyStrong, rightLabel);
                const int slackWidth    = ScaleDip(dpi, 6);
                const int measuredWidth = std::max(minControlWidth, (2 * paddingX) + leftWidth + gapX + trackWidth + gapX + rightWidth + slackWidth);
                const int toggleWidth   = std::min(controlWidth, measuredWidth);

                controls.hToggle = CreateWindowExW(0,
                                                   L"Button",
                                                   L"",
                                                   WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
                                                   controlX,
                                                   y,
                                                   toggleWidth,
                                                   editHeight,
                                                   panel,
                                                   nullptr,
                                                   GetModuleHandleW(nullptr),
                                                   nullptr);
                if (controls.hToggle)
                {
                    ThemedInputFrames::InstallControl(controls.hToggle);
                }

                SetLegacyToggleState(controls.hToggle, controls.field.defaultBool);

                y += editHeight + spacingY;
            }
            else
            {
                const int buttonHeight = std::max(1, optionHeight);
                int optionY            = y;

                for (size_t i = 0; i < 2; ++i)
                {
                    DWORD style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_MULTILINE | BS_AUTORADIOBUTTON;
                    if (i == 0)
                    {
                        style |= WS_GROUP;
                    }

                    const wchar_t* text = (i == 0) ? L"True" : L"False";
                    HWND hButton        = CreateWindowExW(
                        0, L"Button", text, style, controlX, optionY, controlWidth, buttonHeight, panel, nullptr, GetModuleHandleW(nullptr), nullptr);
                    if (hButton)
                    {
                        ThemedInputFrames::InstallControl(hButton);
                    }

                    if (hButton)
                    {
                        const bool checked = (i == 0) ? controls.field.defaultBool : (! controls.field.defaultBool);
                        SendMessageW(hButton, BM_SETCHECK, checked ? BST_CHECKED : BST_UNCHECKED, 0);
                    }

                    controls.choiceButtons.push_back(hButton);
                    optionY += buttonHeight;
                }

                y = optionY + spacingY;
            }
        }
        else if (controls.field.type == PluginConfigFieldType::Option && ! state->theme.highContrast && controls.field.choices.size() == 2)
        {
            size_t leftIndex  = 0;
            size_t rightIndex = 1;

            size_t onIndex  = 0;
            size_t offIndex = 0;
            if (Common::PluginConfiguration::TryGetBoolToggleChoiceIndices(controls.field, onIndex, offIndex))
            {
                leftIndex  = onIndex;
                rightIndex = offIndex;
            }

            controls.toggleOnChoiceIndex  = leftIndex;
            controls.toggleOffChoiceIndex = rightIndex;

            const int paddingX   = ScaleDip(dpi, 6);
            const int gapX       = ScaleDip(dpi, 6);
            const int trackWidth = ScaleDip(dpi, 28);

            std::wstring_view leftLabel;
            std::wstring_view rightLabel;
            if (leftIndex < controls.field.choices.size())
            {
                const auto& choice = controls.field.choices[leftIndex];
                leftLabel          = choice.label.empty() ? std::wstring_view(choice.value) : std::wstring_view(choice.label);
            }
            if (rightIndex < controls.field.choices.size())
            {
                const auto& choice = controls.field.choices[rightIndex];
                rightLabel         = choice.label.empty() ? std::wstring_view(choice.value) : std::wstring_view(choice.label);
            }

            const int leftWidth     = Typography::MeasureSingleLineTextWidthPx(panel, FontRole::BodyStrong, leftLabel);
            const int rightWidth    = Typography::MeasureSingleLineTextWidthPx(panel, FontRole::BodyStrong, rightLabel);
            const int slackWidth    = ScaleDip(dpi, 6);
            const int measuredWidth = std::max(minControlWidth, (2 * paddingX) + leftWidth + gapX + trackWidth + gapX + rightWidth + slackWidth);
            const int toggleWidth   = std::min(controlWidth, measuredWidth);

            controls.hToggle = CreateWindowExW(0,
                                               L"Button",
                                               L"",
                                               WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
                                               controlX,
                                               y,
                                               toggleWidth,
                                               editHeight,
                                               panel,
                                               nullptr,
                                               GetModuleHandleW(nullptr),
                                               nullptr);
            if (controls.hToggle)
            {
                ThemedInputFrames::InstallControl(controls.hToggle);
            }

            bool isLeftActive = true;
            if (rightIndex < controls.field.choices.size() && controls.field.defaultOption == controls.field.choices[rightIndex].value)
            {
                isLeftActive = false;
            }
            else if (leftIndex < controls.field.choices.size() && controls.field.defaultOption == controls.field.choices[leftIndex].value)
            {
                isLeftActive = true;
            }
            SetLegacyToggleState(controls.hToggle, isLeftActive);

            y += editHeight + spacingY;
        }
        else if (controls.field.type == PluginConfigFieldType::Option && controls.field.choices.size() > 2)
        {
            const bool customFrames = ! state->theme.highContrast;
            const int framePadding  = ScaleDip(dpi, 2);

            if (customFrames)
            {
                controls.hComboFrame = CreateWindowExW(
                    0, L"Static", L"", WS_CHILD | WS_VISIBLE, controlX, y, controlWidth, editHeight, panel, nullptr, GetModuleHandleW(nullptr), nullptr);
            }

            const int comboX      = controlX + (customFrames ? framePadding : 0);
            const int comboY      = y + (customFrames ? framePadding : 0);
            const int comboWidth  = std::max(1, controlWidth - (customFrames ? 2 * framePadding : 0));
            const int comboHeight = std::max(1, editHeight - (customFrames ? 2 * framePadding : 0));

            controls.hCombo = CreateWindowExW(0,
                                              L"ComboBox",
                                              L"",
                                              WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST | WS_VSCROLL,
                                              comboX,
                                              comboY,
                                              comboWidth,
                                              comboHeight * 8,
                                              panel,
                                              nullptr,
                                              GetModuleHandleW(nullptr),
                                              nullptr);
            if (controls.hCombo)
            {
                if (controls.hComboFrame)
                {
                    ThemedInputFrames::InstallFrame(controls.hComboFrame, controls.hCombo, &state->inputFrameStyle);
                }
                else
                {
                    ThemedInputFrames::InstallControl(controls.hCombo);
                }
            }

            if (controls.hCombo)
            {
                int selectedIndex = 0;
                for (size_t i = 0; i < controls.field.choices.size(); ++i)
                {
                    const auto& choice            = controls.field.choices[i];
                    const std::wstring_view label = choice.label.empty() ? std::wstring_view(choice.value) : std::wstring_view(choice.label);
                    SendMessageW(controls.hCombo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(label.data()));
                    if (! controls.field.defaultOption.empty() && controls.field.defaultOption == choice.value)
                    {
                        if (i <= static_cast<size_t>(std::numeric_limits<int>::max()))
                        {
                            selectedIndex = static_cast<int>(i);
                        }
                    }
                }

                SendMessageW(controls.hCombo, CB_SETCURSEL, static_cast<WPARAM>(selectedIndex), 0);
            }

            y += editHeight + spacingY;
        }
        else
        {
            const bool isRadio = controls.field.type == PluginConfigFieldType::Option;
            int optionY        = y;

            for (size_t i = 0; i < controls.field.choices.size(); ++i)
            {
                const auto& choice = controls.field.choices[i];

                DWORD style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_MULTILINE;
                style |= isRadio ? BS_AUTORADIOBUTTON : BS_AUTOCHECKBOX;
                if (i == 0)
                {
                    style |= WS_GROUP;
                }

                const int buttonHeight = std::max(1, optionHeight);

                HWND hButton = CreateWindowExW(0,
                                               L"Button",
                                               choice.label.c_str(),
                                               style,
                                               controlX,
                                               optionY,
                                               controlWidth,
                                               buttonHeight,
                                               panel,
                                               nullptr,
                                               GetModuleHandleW(nullptr),
                                               nullptr);
                if (hButton)
                {
                    ThemedInputFrames::InstallControl(hButton);
                }

                if (hButton)
                {
                    if (isRadio)
                    {
                        const bool checked = ! controls.field.defaultOption.empty() && controls.field.defaultOption == choice.value;
                        SendMessageW(hButton, BM_SETCHECK, checked ? BST_CHECKED : BST_UNCHECKED, 0);
                    }
                    else
                    {
                        const bool checked = std::find(controls.field.defaultSelection.begin(), controls.field.defaultSelection.end(), choice.value) !=
                                             controls.field.defaultSelection.end();
                        SendMessageW(hButton, BM_SETCHECK, checked ? BST_CHECKED : BST_UNCHECKED, 0);
                    }
                }

                controls.choiceButtons.push_back(hButton);
                optionY += buttonHeight;
            }

            y = optionY + spacingY;
        }

        const int infoX     = left;
        const int infoWidth = availableWidth;

        if (! controls.field.description.empty())
        {
            const int commentHeight = MeasureInfoHeight(panel, infoWidth, controls.field.description);

            controls.hComment = CreateWindowExW(0,
                                                L"Static",
                                                controls.field.description.c_str(),
                                                WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX | SS_EDITCONTROL,
                                                infoX,
                                                y,
                                                infoWidth,
                                                commentHeight,
                                                panel,
                                                nullptr,
                                                GetModuleHandleW(nullptr),
                                                nullptr);
            y += commentHeight + ScaleDip(dpi, 4);
        }

        const std::wstring defaultsText = BuildFieldDefaultsTextForDisplay(controls.field);
        if (! defaultsText.empty())
        {
            const int defaultsHeight = MeasureInfoHeight(panel, infoWidth, defaultsText);
            controls.hDefaults       = CreateWindowExW(0,
                                                       L"Static",
                                                       defaultsText.c_str(),
                                                       WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX | SS_EDITCONTROL,
                                                       infoX,
                                                       y,
                                                       infoWidth,
                                                       defaultsHeight,
                                                       panel,
                                                       nullptr,
                                                       GetModuleHandleW(nullptr),
                                                       nullptr);
            y += defaultsHeight + spacingY;
        }

        state->controls.push_back(std::move(controls));
        if (kEnableManagePluginsDxFormStatics || kEnableManagePluginsDxFormInputs)
        {
            AttachDxFieldHosts(panel, *state, state->controls.back(), kEnableManagePluginsDxFormStatics, kEnableManagePluginsDxFormInputs);
        }
    }

    state->contentHeight         = y + margin;
    state->configurationJsonUtf8 = BuildConfigurationJson(state->controls, state->configurationJsonUtf8);
    if (state->baselineConfigurationJsonUtf8.empty())
    {
        state->baselineConfigurationJsonUtf8 = state->configurationJsonUtf8;
    }
    UpdatePanelScrollInfo(panel, *state);
    g_pluginConfigurationDialogWindow.store(dlg, std::memory_order_release);
    return TRUE;
}

INT_PTR OnPluginConfigDialogCtlColorDialog(PluginConfigDialogState* state)
{
    if (! state || ! state->backgroundBrush)
    {
        return FALSE;
    }
    return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
}

INT_PTR OnPluginConfigDialogCtlColorStatic(PluginConfigDialogState* state, HDC hdc, HWND control)
{
    if (! state || ! state->backgroundBrush)
    {
        return FALSE;
    }

    COLORREF textColor = state->theme.menu.text;
    if (control)
    {
        const LONG_PTR style = GetWindowLongPtrW(control, GWL_STYLE);
        if ((style & WS_DISABLED) != 0)
        {
            textColor = state->theme.menu.disabledText;
        }
    }

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, textColor);

    // Combo box drop-down list controls often paint their selection field via a child static window; match the input background.
    HWND parent = control ? GetParent(control) : nullptr;
    if (parent)
    {
        std::array<wchar_t, 32> className{};
        const int len = GetClassNameW(parent, className.data(), static_cast<int>(className.size()));
        if (len > 0 && _wcsicmp(className.data(), L"ComboBox") == 0)
        {
            const COLORREF background = state->inputBrush ? state->inputBackgroundColor : state->theme.windowBackground;
            SetBkColor(hdc, background);
            return reinterpret_cast<INT_PTR>(state->inputBrush ? state->inputBrush.get() : state->backgroundBrush.get());
        }
    }

    return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
}

INT_PTR OnPluginConfigDialogCtlColorButton(PluginConfigDialogState* state, HDC hdc, HWND control)
{
    if (! state || ! state->backgroundBrush || ! control)
    {
        return FALSE;
    }

    const LONG_PTR style = GetWindowLongPtrW(control, GWL_STYLE);
    const LONG_PTR type  = style & BS_TYPEMASK;

    const bool themed = type == BS_CHECKBOX || type == BS_AUTOCHECKBOX || type == BS_RADIOBUTTON || type == BS_AUTORADIOBUTTON || type == BS_3STATE ||
                        type == BS_AUTO3STATE || type == BS_GROUPBOX;

    if (! themed)
    {
        return FALSE;
    }

    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, (style & WS_DISABLED) != 0 ? state->theme.menu.disabledText : state->theme.menu.text);
    return reinterpret_cast<INT_PTR>(state->backgroundBrush.get());
}

INT_PTR OnPluginConfigDialogCtlColorEdit(PluginConfigDialogState* state, HDC hdc)
{
    if (! state || ! state->backgroundBrush)
    {
        return FALSE;
    }

    const COLORREF background = state->inputBrush ? state->inputBackgroundColor : state->theme.windowBackground;
    SetBkColor(hdc, background);
    SetTextColor(hdc, state->theme.menu.text);
    return reinterpret_cast<INT_PTR>(state->inputBrush ? state->inputBrush.get() : state->backgroundBrush.get());
}

INT_PTR OnPluginConfigDialogCtlColorListBox(PluginConfigDialogState* state, HDC hdc)
{
    if (! state || ! state->backgroundBrush)
    {
        return FALSE;
    }

    const COLORREF background = state->inputBrush ? state->inputBackgroundColor : state->theme.windowBackground;
    SetBkColor(hdc, background);
    SetTextColor(hdc, state->theme.menu.text);
    return reinterpret_cast<INT_PTR>(state->inputBrush ? state->inputBrush.get() : state->backgroundBrush.get());
}

void DestroyPluginConfigChildWindows(PluginConfigDialogState& state) noexcept
{
    const HWND panel = state.panel;
    DetachDxFieldHosts(state);
    DetachDxCommandButtons(state);
    static_cast<void>(DetachPluginConfigPanel(state));

    if (! panel)
    {
        state.controls.clear();
        return;
    }

    while (const HWND child = GetWindow(panel, GW_CHILD))
    {
        DestroyWindow(child);
    }

    state.controls.clear();
}

[[nodiscard]] HRESULT LoadPluginConfigSourceData(const PluginConfigDialogState& state,
                                                 std::string& outSchemaJsonUtf8,
                                                 std::string& outConfigJsonUtf8,
                                                 AppTheme& outTheme) noexcept
{
    Common::Settings::Settings* sourceSettings = state.reloadSourceSettings ? state.reloadSourceSettings : state.settings;
    if (! sourceSettings || state.pluginId.empty())
    {
        return E_INVALIDARG;
    }

    outTheme = SettingsHotReload::ResolveDialogThemeFromSettings(*sourceSettings);

    HRESULT schemaHr = E_FAIL;
    if (state.pluginType == PluginType::FileSystem)
    {
        auto& manager = FileSystemPluginManager::GetInstance();
        schemaHr      = manager.GetConfigurationSchema(state.pluginId, *sourceSettings, outSchemaJsonUtf8);
    }
    else
    {
        auto& manager = ViewerPluginManager::GetInstance();
        schemaHr      = manager.GetConfigurationSchema(state.pluginId, *sourceSettings, outSchemaJsonUtf8);
    }
    if (FAILED(schemaHr))
    {
        return schemaHr;
    }

    outConfigJsonUtf8.clear();
    const auto it = sourceSettings->plugins.configurationByPluginId.find(state.pluginId);
    if (it != sourceSettings->plugins.configurationByPluginId.end() && ! std::holds_alternative<std::monostate>(it->second.value))
    {
        const HRESULT serializeHr = Common::Settings::SerializeJsonValue(it->second, outConfigJsonUtf8);
        if (FAILED(serializeHr))
        {
            return serializeHr;
        }
    }

    if (! outConfigJsonUtf8.empty())
    {
        return S_OK;
    }

    if (state.pluginType == PluginType::FileSystem)
    {
        auto& manager = FileSystemPluginManager::GetInstance();
        return manager.GetConfiguration(state.pluginId, *sourceSettings, outConfigJsonUtf8);
    }

    auto& manager = ViewerPluginManager::GetInstance();
    return manager.GetConfiguration(state.pluginId, *sourceSettings, outConfigJsonUtf8);
}

[[nodiscard]] HRESULT RebuildPluginConfigDialog(HWND dlg,
                                                PluginConfigDialogState& state,
                                                std::string schemaJsonUtf8,
                                                std::string displayedConfigJsonUtf8,
                                                std::string baselineConfigJsonUtf8,
                                                const AppTheme& theme) noexcept
{
    DestroyPluginConfigChildWindows(state);

    state.theme                         = theme;
    state.schemaJsonUtf8                = std::move(schemaJsonUtf8);
    state.configurationJsonUtf8         = std::move(displayedConfigJsonUtf8);
    state.baselineConfigurationJsonUtf8 = std::move(baselineConfigJsonUtf8);

    if (OnPluginConfigDialogInit(dlg, &state) == FALSE)
    {
        return E_FAIL;
    }

    if (state.baselineConfigurationJsonUtf8.empty())
    {
        state.baselineConfigurationJsonUtf8 = BuildConfigurationJson(state.controls, state.configurationJsonUtf8);
    }
    return S_OK;
}

[[nodiscard]] bool IsPluginConfigDialogDirty(const PluginConfigDialogState& state) noexcept
{
    return BuildConfigurationJson(state.controls, state.configurationJsonUtf8) != state.baselineConfigurationJsonUtf8;
}

[[nodiscard]] bool ResolvePluginConfigStaleSaveConflict(HWND dlg, PluginConfigDialogState& state) noexcept
{
    if (! state.staleFromExternalReload)
    {
        return true;
    }

    SettingsHotReload::StaleSaveChoice choice = SettingsHotReload::StaleSaveChoice::Cancel;
    const HRESULT promptHr                    = SettingsHotReload::PromptStaleSaveConflict(dlg, GetPluginConfigEditorName(state), choice);
    if (FAILED(promptHr))
    {
        Debug::Warning(L"PluginConfig: failed to prompt for stale save conflict (hr=0x{:08X})", static_cast<unsigned long>(promptHr));
        return false;
    }

    if (choice == SettingsHotReload::StaleSaveChoice::ReloadFromDisk)
    {
        std::string sourceSchema;
        std::string sourceConfig;
        AppTheme sourceTheme{};
        const HRESULT loadHr = LoadPluginConfigSourceData(state, sourceSchema, sourceConfig, sourceTheme);
        if (FAILED(loadHr))
        {
            Debug::Warning(L"PluginConfig: failed to reload configuration source (hr=0x{:08X})", static_cast<unsigned long>(loadHr));
            return false;
        }

        state.staleFromExternalReload = false;
        static_cast<void>(RebuildPluginConfigDialog(dlg, state, std::move(sourceSchema), sourceConfig, std::move(sourceConfig), sourceTheme));
        return false;
    }

    if (choice == SettingsHotReload::StaleSaveChoice::Cancel)
    {
        return false;
    }

    state.staleFromExternalReload = false;
    return true;
}

INT_PTR OnPluginConfigDialogSettingsReloadedFromDisk(HWND dlg, PluginConfigDialogState& state) noexcept
{
    std::string sourceSchema;
    std::string sourceConfig;
    AppTheme sourceTheme{};
    const HRESULT loadHr = LoadPluginConfigSourceData(state, sourceSchema, sourceConfig, sourceTheme);
    if (FAILED(loadHr))
    {
        Debug::Warning(L"PluginConfig: failed to load updated configuration source (hr=0x{:08X})", static_cast<unsigned long>(loadHr));
        return TRUE;
    }

    if (IsPluginConfigDialogDirty(state))
    {
        const std::string currentConfig = BuildConfigurationJson(state.controls, state.configurationJsonUtf8);

        SettingsHotReload::ExternalReloadChoice choice = SettingsHotReload::ExternalReloadChoice::KeepEditing;
        const HRESULT promptHr                         = SettingsHotReload::PromptExternalReloadConflict(dlg, GetPluginConfigEditorName(state), choice);
        if (FAILED(promptHr))
        {
            Debug::Warning(L"PluginConfig: failed to prompt for external reload conflict (hr=0x{:08X})", static_cast<unsigned long>(promptHr));
            return TRUE;
        }

        if (choice == SettingsHotReload::ExternalReloadChoice::KeepEditing)
        {
            state.staleFromExternalReload = true;
            static_cast<void>(RebuildPluginConfigDialog(dlg,
                                                        state,
                                                        std::move(sourceSchema),
                                                        currentConfig.empty() ? state.configurationJsonUtf8 : currentConfig,
                                                        std::move(sourceConfig),
                                                        sourceTheme));
            return TRUE;
        }
    }

    state.staleFromExternalReload = false;
    static_cast<void>(RebuildPluginConfigDialog(dlg, state, std::move(sourceSchema), sourceConfig, std::move(sourceConfig), sourceTheme));
    return TRUE;
}

INT_PTR OnPluginConfigDialogCommand(HWND dlg, PluginConfigDialogState* state, UINT commandId, UINT /*codeNotify*/, HWND /*hwndCtl*/)
{
    if (! state)
    {
        return FALSE;
    }

    if (commandId == IDOK)
    {
        if (! ResolvePluginConfigStaleSaveConflict(dlg, *state))
        {
            return TRUE;
        }

        const std::string configJson = BuildConfigurationJson(state->controls, state->configurationJsonUtf8);
        if (configJson.empty())
        {
            EndDialog(dlg, IDCANCEL);
            return TRUE;
        }

        if (state->commitMode == PluginConfigCommitMode::UpdateSettingsOnly)
        {
            if (! state->settings || state->pluginId.empty())
            {
                EndDialog(dlg, IDCANCEL);
                return TRUE;
            }

            Common::Settings::JsonValue parsedValue;
            const HRESULT parseHr = Common::Settings::ParseJsonValue(configJson, parsedValue);
            if (FAILED(parseHr))
            {
                const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
                const std::wstring message = LoadStringResource(nullptr, IDS_MSG_PLUGIN_CONFIG_APPLY_FAILED);
                ShowDialogAlert(dlg, HOST_ALERT_ERROR, title, message);
                return TRUE;
            }

            bool clearValue = std::holds_alternative<std::monostate>(parsedValue.value);
            if (! clearValue)
            {
                const auto* obj = std::get_if<Common::Settings::JsonValue::ObjectPtr>(&parsedValue.value);
                clearValue      = obj && *obj && (*obj)->members.empty();
            }

            if (clearValue)
            {
                state->settings->plugins.configurationByPluginId.erase(state->pluginId);
            }
            else
            {
                state->settings->plugins.configurationByPluginId[state->pluginId] = std::move(parsedValue);
            }

            state->configurationJsonUtf8         = configJson;
            state->baselineConfigurationJsonUtf8 = configJson;
            state->staleFromExternalReload       = false;
            EndDialog(dlg, IDOK);
            return TRUE;
        }

        HRESULT hr = E_FAIL;
        if (state->pluginType == PluginType::FileSystem)
        {
            auto& manager = FileSystemPluginManager::GetInstance();
            hr            = manager.SetConfiguration(state->pluginId, configJson, *state->settings);
        }
        else
        {
            auto& manager = ViewerPluginManager::GetInstance();
            hr            = manager.SetConfiguration(state->pluginId, configJson, *state->settings);
        }
        if (FAILED(hr))
        {
            const std::wstring title   = LoadStringResource(nullptr, IDS_CAPTION_ERROR);
            const std::wstring message = LoadStringResource(nullptr, IDS_MSG_PLUGIN_CONFIG_APPLY_FAILED);
            ShowDialogAlert(dlg, HOST_ALERT_ERROR, title, message);
            return TRUE;
        }

        if (state->settings)
        {
            const HRESULT persistHr = PersistSettings(dlg, *state->settings, state->appId);
            if (FAILED(persistHr))
            {
                return TRUE;
            }
        }

        state->configurationJsonUtf8         = configJson;
        state->baselineConfigurationJsonUtf8 = configJson;
        state->staleFromExternalReload       = false;
        EndDialog(dlg, IDOK);
        return TRUE;
    }

    if (commandId == IDCANCEL)
    {
        EndDialog(dlg, IDCANCEL);
        return TRUE;
    }

    return FALSE;
}

#ifdef ENABLE_TESTS
INT_PTR OnPluginConfigDialogDebug(HWND dlg, PluginConfigDialogState* state, WPARAM wp, LPARAM lp) noexcept
{
    if (! dlg || ! state)
    {
        return FALSE;
    }

    switch (static_cast<PluginConfigDebugCommand>(wp))
    {
        case PluginConfigDebugCommand::GetSnapshot:
        {
            auto* snapshot = reinterpret_cast<PluginConfigurationDialogDebugSnapshot*>(lp);
            if (! snapshot)
            {
                return FALSE;
            }

            snapshot->focusKind = PluginConfigurationDialogDebugFocusKind::None;
            snapshot->focusLabel.clear();

            snapshot->usesDxUiCommandButtons            = state->usesDxUiCommandButtons;
            snapshot->legacyOwnerDrawCommandButtonCount = static_cast<size_t>(UsesLegacyOwnerDrawButton(GetDlgItem(dlg, IDOK))) +
                                                          static_cast<size_t>(UsesLegacyOwnerDrawButton(GetDlgItem(dlg, IDCANCEL)));
            snapshot->visibleLegacyCommandButtonCount   = CountIfVisible(GetDlgItem(dlg, IDOK)) + CountIfVisible(GetDlgItem(dlg, IDCANCEL));
            snapshot->visibleDxCommandButtonHostCount   = 0u;
            for (const DxCommandButtonHost& slot : state->dxCommandButtons)
            {
                snapshot->visibleDxCommandButtonHostCount += CountIfVisible(slot.hostHwnd.get());
            }

            snapshot->visibleLegacyFormControlCount = 0u;
            snapshot->visibleLegacyFormStaticCount  = 0u;
            snapshot->visibleLegacyFormInputCount   = 0u;
            snapshot->legacyOwnerDrawFormInputCount = 0u;
            snapshot->visibleDxFormHostCount        = 0u;
            snapshot->visibleDxFormStaticHostCount  = 0u;
            snapshot->visibleDxFormInputHostCount   = 0u;
            for (const auto& controls : state->controls)
            {
                snapshot->visibleLegacyFormStaticCount += CountIfVisible(controls.hLabel);
                snapshot->visibleLegacyFormStaticCount += CountIfVisible(controls.hComment);
                snapshot->visibleLegacyFormStaticCount += CountIfVisible(controls.hDefaults);
                snapshot->visibleLegacyFormInputCount += CountIfVisible(controls.hEditFrame);
                snapshot->visibleLegacyFormInputCount += CountIfVisible(controls.hEdit);
                snapshot->visibleLegacyFormInputCount += CountIfVisible(controls.hComboFrame);
                snapshot->visibleLegacyFormInputCount += CountIfVisible(controls.hCombo);
                snapshot->visibleLegacyFormInputCount += CountIfVisible(controls.hToggle);
                snapshot->legacyOwnerDrawFormInputCount += static_cast<size_t>(UsesLegacyOwnerDrawButton(controls.hToggle));
                for (const HWND choiceButton : controls.choiceButtons)
                {
                    snapshot->visibleLegacyFormInputCount += CountIfVisible(choiceButton);
                }

                snapshot->visibleDxFormStaticHostCount += CountIfVisible(controls.dxLabelSlot.hostHwnd.get());
                snapshot->visibleDxFormStaticHostCount += CountIfVisible(controls.dxDescriptionSlot.hostHwnd.get());
                snapshot->visibleDxFormStaticHostCount += CountIfVisible(controls.dxDefaultsSlot.hostHwnd.get());
                snapshot->visibleDxFormInputHostCount += CountIfVisible(controls.dxEditSlot.hostHwnd.get());
                snapshot->visibleDxFormInputHostCount += CountIfVisible(controls.dxComboSlot.hostHwnd.get());
                snapshot->visibleDxFormInputHostCount += CountIfVisible(controls.dxToggleSlot.hostHwnd.get());
                for (const auto& choiceControl : controls.dxChoiceControls)
                {
                    snapshot->visibleDxFormInputHostCount += CountIfVisible(choiceControl.slot.hostHwnd.get());
                }
            }
            snapshot->visibleLegacyFormControlCount = snapshot->visibleLegacyFormStaticCount + snapshot->visibleLegacyFormInputCount;
            snapshot->visibleDxFormHostCount        = snapshot->visibleDxFormStaticHostCount + snapshot->visibleDxFormInputHostCount;
            snapshot->usesDxUiFormStatics           = snapshot->visibleDxFormStaticHostCount > 0u;
            snapshot->usesDxUiFormInputs            = snapshot->visibleDxFormInputHostCount > 0u;
            snapshot->usesDxUiFormSurface           = snapshot->visibleDxFormHostCount > 0u;
            snapshot->panelContentHeight            = std::max(0, state->contentHeight);
            snapshot->panelScrollPosY               = std::max(0, state->scrollPosY);
            snapshot->themeDark                     = state->theme.dark;
            snapshot->themeHighContrast             = state->theme.highContrast;
            snapshot->themeRainbow                  = state->theme.menu.rainbowMode;
            if (state->panel && IsWindow(state->panel) != FALSE)
            {
                RECT client{};
                if (GetClientRect(state->panel, &client))
                {
                    snapshot->panelClientHeight = std::max(0l, client.bottom - client.top);
                }
            }
            snapshot->panelHasVerticalScrollbar = snapshot->panelContentHeight > snapshot->panelClientHeight;
            FillPluginConfigFocusSnapshot(dlg, *state, *snapshot);
            SetWindowLongPtrW(dlg, DWLP_MSGRESULT, TRUE);
            return TRUE;
        }
        case PluginConfigDebugCommand::ScrollByWheelDetents:
        {
            if (! state->panel || IsWindow(state->panel) == FALSE)
            {
                return FALSE;
            }

            const int wheelDetents = static_cast<int>(lp);
            if (wheelDetents == 0)
            {
                return TRUE;
            }

            UINT lines = 3;
            SystemParametersInfoW(SPI_GETWHEELSCROLLLINES, 0, &lines, 0);
            if (lines == WHEEL_PAGESCROLL)
            {
                lines = 3;
            }

            const UINT dpi     = GetDpiForWindow(state->panel);
            const int lineStep = std::max(1, ScaleDip(dpi, 18));
            ScrollPanelTo(state->panel, *state, state->scrollPosY - (wheelDetents * static_cast<int>(lines) * lineStep));
            UpdatePanelScrollInfo(state->panel, *state);
            SetWindowLongPtrW(dlg, DWLP_MSGRESULT, TRUE);
            return TRUE;
        }
        case PluginConfigDebugCommand::FocusFirstInput:
        {
            const bool focused = FocusFirstVisibleInteractiveTarget(*state);
            SetWindowLongPtrW(dlg, DWLP_MSGRESULT, focused ? TRUE : FALSE);
            return TRUE;
        }
        case PluginConfigDebugCommand::GetFirstVisibleToggleRect:
        {
            auto* hostRect = reinterpret_cast<PluginConfigurationDialogDebugHostRect*>(lp);
            if (! hostRect)
            {
                return FALSE;
            }

            hostRect->host = nullptr;
            hostRect->rect = {};
            for (const auto& controls : state->controls)
            {
                if (! controls.dxToggleControl || ! controls.dxToggleSlot.hostHwnd || IsWindow(controls.dxToggleSlot.hostHwnd.get()) == FALSE ||
                    ! controls.dxToggleControl->IsVisible() || ! controls.dxToggleControl->IsEnabled())
                {
                    continue;
                }

                if (! controls.dxToggleSlot.host)
                {
                    continue;
                }

                const auto bounds     = controls.dxToggleControl->GetBounds();
                hostRect->host        = controls.dxToggleSlot.hostHwnd.get();
                hostRect->rect.left   = static_cast<LONG>(std::lround(controls.dxToggleSlot.host->DipsToPixels(bounds.left)));
                hostRect->rect.top    = static_cast<LONG>(std::lround(controls.dxToggleSlot.host->DipsToPixels(bounds.top)));
                hostRect->rect.right  = static_cast<LONG>(std::lround(controls.dxToggleSlot.host->DipsToPixels(bounds.right)));
                hostRect->rect.bottom = static_cast<LONG>(std::lround(controls.dxToggleSlot.host->DipsToPixels(bounds.bottom)));
                SetWindowLongPtrW(dlg, DWLP_MSGRESULT, TRUE);
                return TRUE;
            }

            SetWindowLongPtrW(dlg, DWLP_MSGRESULT, FALSE);
            return FALSE;
        }
        case PluginConfigDebugCommand::GetVisibleToggleRectByLabel:
        {
            auto* hostRect = reinterpret_cast<PluginConfigurationDialogDebugLabeledHostRect*>(lp);
            if (! hostRect || ! hostRect->label || hostRect->label[0] == L'\0')
            {
                return FALSE;
            }

            hostRect->host = nullptr;
            hostRect->rect = {};
            const std::wstring_view requestedLabel(hostRect->label);
            for (const auto& controls : state->controls)
            {
                if (controls.field.label != requestedLabel || ! controls.dxToggleControl || ! controls.dxToggleSlot.hostHwnd ||
                    IsWindow(controls.dxToggleSlot.hostHwnd.get()) == FALSE || ! controls.dxToggleControl->IsVisible() ||
                    ! controls.dxToggleControl->IsEnabled() || ! controls.dxToggleSlot.host)
                {
                    continue;
                }

                const auto metrics     = controls.dxToggleControl->GetLayoutMetrics();
                const D2D1_RECT_F rect = (metrics.trackRect.right > metrics.trackRect.left && metrics.trackRect.bottom > metrics.trackRect.top)
                                             ? metrics.trackRect
                                             : controls.dxToggleControl->GetBounds();
                hostRect->host         = controls.dxToggleSlot.hostHwnd.get();
                hostRect->rect.left    = static_cast<LONG>(std::lround(controls.dxToggleSlot.host->DipsToPixels(rect.left)));
                hostRect->rect.top     = static_cast<LONG>(std::lround(controls.dxToggleSlot.host->DipsToPixels(rect.top)));
                hostRect->rect.right   = static_cast<LONG>(std::lround(controls.dxToggleSlot.host->DipsToPixels(rect.right)));
                hostRect->rect.bottom  = static_cast<LONG>(std::lround(controls.dxToggleSlot.host->DipsToPixels(rect.bottom)));
                SetWindowLongPtrW(dlg, DWLP_MSGRESULT, TRUE);
                return TRUE;
            }

            SetWindowLongPtrW(dlg, DWLP_MSGRESULT, FALSE);
            return FALSE;
        }
        case PluginConfigDebugCommand::GetFocusedHost:
        {
            auto* focusedHost = reinterpret_cast<PluginConfigurationDialogDebugHost*>(lp);
            if (! focusedHost)
            {
                return FALSE;
            }

            const HWND focused = GetFocus();
            focusedHost->host  = nullptr;

            for (const auto& controls : state->controls)
            {
                if (controls.dxEditSlot.hostHwnd && EditSlotOwnsFocus(controls, focused))
                {
                    focusedHost->host = controls.dxEditSlot.hostHwnd.get();
#ifdef ENABLE_TESTS
                    RememberPluginConfigDebugFocusedHost(*state, focusedHost->host);
#endif
                    SetWindowLongPtrW(dlg, DWLP_MSGRESULT, TRUE);
                    return TRUE;
                }
                if (controls.dxComboSlot.hostHwnd && ComboSlotOwnsFocus(controls, focused))
                {
                    focusedHost->host = controls.dxComboSlot.hostHwnd.get();
#ifdef ENABLE_TESTS
                    RememberPluginConfigDebugFocusedHost(*state, focusedHost->host);
#endif
                    SetWindowLongPtrW(dlg, DWLP_MSGRESULT, TRUE);
                    return TRUE;
                }
                if (controls.dxToggleSlot.hostHwnd && ToggleSlotOwnsFocus(controls, focused))
                {
                    focusedHost->host = controls.dxToggleSlot.hostHwnd.get();
#ifdef ENABLE_TESTS
                    RememberPluginConfigDebugFocusedHost(*state, focusedHost->host);
#endif
                    SetWindowLongPtrW(dlg, DWLP_MSGRESULT, TRUE);
                    return TRUE;
                }
                for (size_t choiceIndex = 0; choiceIndex < controls.dxChoiceControls.size(); ++choiceIndex)
                {
                    if (ChoiceSlotOwnsFocus(controls, choiceIndex, focused))
                    {
                        focusedHost->host = controls.dxChoiceControls[choiceIndex].slot.hostHwnd.get();
#ifdef ENABLE_TESTS
                        RememberPluginConfigDebugFocusedHost(*state, focusedHost->host);
#endif
                        SetWindowLongPtrW(dlg, DWLP_MSGRESULT, TRUE);
                        return TRUE;
                    }
                }
            }

            for (size_t index = 0; index < state->dxCommandButtons.size(); ++index)
            {
                const DxCommandButtonHost& slot = state->dxCommandButtons[index];
                const HWND legacyButton         = GetLegacyCommandButton(dlg, static_cast<DxCommandButtonIndex>(index));
                if (slot.hostHwnd && CommandButtonHostOwnsFocus(slot, legacyButton, focused))
                {
                    focusedHost->host = slot.hostHwnd.get();
#ifdef ENABLE_TESTS
                    RememberPluginConfigDebugFocusedHost(*state, focusedHost->host);
#endif
                    SetWindowLongPtrW(dlg, DWLP_MSGRESULT, TRUE);
                    return TRUE;
                }
            }

            SetWindowLongPtrW(dlg, DWLP_MSGRESULT, FALSE);
            return FALSE;
        }
        case PluginConfigDebugCommand::AdvanceTab:
        {
            PluginConfigurationDialogDebugHost focusedHost{};
            const HWND focused = GetFocus();
            if (! focused)
            {
#ifdef ENABLE_TESTS
                TracePluginConfigDebug(L"plugin-config advance command: no focused HWND; attempting fallback focus recovery");
#endif
            }

            for (const auto& controls : state->controls)
            {
                if (controls.dxEditSlot.hostHwnd && EditSlotOwnsFocus(controls, focused))
                {
                    focusedHost.host = controls.dxEditSlot.hostHwnd.get();
                    break;
                }
                if (controls.dxComboSlot.hostHwnd && ComboSlotOwnsFocus(controls, focused))
                {
                    focusedHost.host = controls.dxComboSlot.hostHwnd.get();
                    break;
                }
                if (controls.dxToggleSlot.hostHwnd && ToggleSlotOwnsFocus(controls, focused))
                {
                    focusedHost.host = controls.dxToggleSlot.hostHwnd.get();
                    break;
                }
                for (size_t choiceIndex = 0; choiceIndex < controls.dxChoiceControls.size(); ++choiceIndex)
                {
                    if (ChoiceSlotOwnsFocus(controls, choiceIndex, focused))
                    {
                        focusedHost.host = controls.dxChoiceControls[choiceIndex].slot.hostHwnd.get();
                        break;
                    }
                }
                if (focusedHost.host)
                {
                    break;
                }
            }

            if (! focusedHost.host)
            {
                for (size_t index = 0; index < state->dxCommandButtons.size(); ++index)
                {
                    const DxCommandButtonHost& slot = state->dxCommandButtons[index];
                    const HWND legacyButton         = GetLegacyCommandButton(dlg, static_cast<DxCommandButtonIndex>(index));
                    if (slot.hostHwnd && CommandButtonHostOwnsFocus(slot, legacyButton, focused))
                    {
                        focusedHost.host = slot.hostHwnd.get();
                        break;
                    }
                }
            }

            if (! focusedHost.host)
            {
                if (TryRecoverPluginConfigDebugFocusedHost(*state, focusedHost.host))
                {
#ifdef ENABLE_TESTS
                    TracePluginConfigDebug(std::format(L"plugin-config advance command: reusing logical focused host={:#x} index={} after focus loss",
                                                       reinterpret_cast<uintptr_t>(focusedHost.host),
                                                       static_cast<unsigned long long>(state->lastDebugFocusedHostIndex)));
#endif
                }
            }

            if (! focusedHost.host)
            {
                if (! FocusFirstVisibleInteractiveTarget(*state))
                {
#ifdef ENABLE_TESTS
                    TracePluginConfigDebug(std::format(L"plugin-config advance command: no focused host and fallback focus failed; focused={:#x}",
                                                       reinterpret_cast<uintptr_t>(focused)));
#endif
                    return FALSE;
                }

                const HWND fallbackFocus = GetFocus();
                for (const auto& controls : state->controls)
                {
                    if (controls.dxEditSlot.hostHwnd && EditSlotOwnsFocus(controls, fallbackFocus))
                    {
                        focusedHost.host = controls.dxEditSlot.hostHwnd.get();
                        break;
                    }
                    if (controls.dxComboSlot.hostHwnd && ComboSlotOwnsFocus(controls, fallbackFocus))
                    {
                        focusedHost.host = controls.dxComboSlot.hostHwnd.get();
                        break;
                    }
                    if (controls.dxToggleSlot.hostHwnd && ToggleSlotOwnsFocus(controls, fallbackFocus))
                    {
                        focusedHost.host = controls.dxToggleSlot.hostHwnd.get();
                        break;
                    }
                    for (size_t choiceIndex = 0; choiceIndex < controls.dxChoiceControls.size(); ++choiceIndex)
                    {
                        if (ChoiceSlotOwnsFocus(controls, choiceIndex, fallbackFocus))
                        {
                            focusedHost.host = controls.dxChoiceControls[choiceIndex].slot.hostHwnd.get();
                            break;
                        }
                    }
                    if (focusedHost.host)
                    {
                        break;
                    }
                }

                if (! focusedHost.host)
                {
#ifdef ENABLE_TESTS
                    TracePluginConfigDebug(std::format(L"plugin-config advance command: fallback focus unresolved focused={:#x} fallbackFocus={:#x}",
                                                       reinterpret_cast<uintptr_t>(focused),
                                                       reinterpret_cast<uintptr_t>(fallbackFocus)));
#endif
                    return FALSE;
                }
#ifdef ENABLE_TESTS
                TracePluginConfigDebug(std::format(L"plugin-config advance command: recovered focused host={:#x} after fallback from focused={:#x}",
                                                   reinterpret_cast<uintptr_t>(focusedHost.host),
                                                   reinterpret_cast<uintptr_t>(focused)));
#endif
            }
            // The mover records the destination host after a successful focus handoff.
            // Keeping the previous destination lets debug Tab routing recover when
            // Win32 focus is transiently null after SendMessage returns.
            const bool advanced = MovePluginConfigDialogTabFocusFromHost(focusedHost.host, lp != 0);
#ifdef ENABLE_TESTS
            TracePluginConfigDebug(std::format(L"plugin-config advance command: focusedHost={:#x} reverse={} advanced={}",
                                               reinterpret_cast<uintptr_t>(focusedHost.host),
                                               lp != 0 ? 1 : 0,
                                               advanced ? 1 : 0));
#endif
            // The focus handoff pumps nested dialog traffic; preserve the intended
            // result explicitly so SendMessage() sees the final boolean.
            SetWindowLongPtrW(dlg, DWLP_MSGRESULT, advanced ? TRUE : FALSE);
            return TRUE;
        }
        case PluginConfigDebugCommand::Cancel: EndDialog(dlg, IDCANCEL); return TRUE;
    }

    return FALSE;
}
#endif

INT_PTR CALLBACK PluginConfigDialogProc(HWND dlg, UINT msg, WPARAM wp, LPARAM lp)
{
    auto* state = reinterpret_cast<PluginConfigDialogState*>(GetWindowLongPtrW(dlg, DWLP_USER));

    switch (msg)
    {
        case WM_INITDIALOG: return OnPluginConfigDialogInit(dlg, reinterpret_cast<PluginConfigDialogState*>(lp));
        case WndMsg::kSettingsReloadedFromDisk:
            if (state)
            {
                return OnPluginConfigDialogSettingsReloadedFromDisk(dlg, *state);
            }
            return TRUE;
        case WM_CTLCOLORDLG: return OnPluginConfigDialogCtlColorDialog(state);
        case WM_CTLCOLORSTATIC: return OnPluginConfigDialogCtlColorStatic(state, reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_CTLCOLORBTN: return OnPluginConfigDialogCtlColorButton(state, reinterpret_cast<HDC>(wp), reinterpret_cast<HWND>(lp));
        case WM_CTLCOLOREDIT: return OnPluginConfigDialogCtlColorEdit(state, reinterpret_cast<HDC>(wp));
        case WM_CTLCOLORLISTBOX: return OnPluginConfigDialogCtlColorListBox(state, reinterpret_cast<HDC>(wp));
        case WndMsg::kPluginConfigurationDialogApplyTheme:
        {
            if (! state)
            {
                return FALSE;
            }

            const auto* theme = reinterpret_cast<const AppTheme*>(lp);
            if (! theme)
            {
                return FALSE;
            }

            const std::string currentConfig   = BuildConfigurationJson(state->controls, state->configurationJsonUtf8);
            const std::string displayedConfig = currentConfig.empty() ? state->configurationJsonUtf8 : currentConfig;
            const std::string baselineConfig  = state->baselineConfigurationJsonUtf8.empty() ? displayedConfig : state->baselineConfigurationJsonUtf8;
            const HRESULT rebuildHr           = RebuildPluginConfigDialog(dlg, *state, state->schemaJsonUtf8, displayedConfig, baselineConfig, *theme);
            if (FAILED(rebuildHr))
            {
                return FALSE;
            }

            InvalidateRect(dlg, nullptr, FALSE);
            return TRUE;
        }
#ifdef ENABLE_TESTS
        case WndMsg::kPluginConfigurationDialogDebug: return OnPluginConfigDialogDebug(dlg, state, wp, lp);
#endif
        case WM_NCDESTROY:
            g_pluginConfigurationDialogWindow.store(nullptr, std::memory_order_release);
            if (state)
            {
                DestroyPluginConfigChildWindows(*state);
            }
            SettingsHotReload::UnregisterParticipant(dlg);
            return FALSE;
        case WM_NCACTIVATE:
            if (state)
            {
                ApplyTitleBarTheme(dlg, state->theme, wp != FALSE);
            }
            return FALSE;
        case WM_SIZING:
        {
            if (! state || state->fixedWindowWidthPx <= 0)
            {
                break;
            }

            auto* rc = reinterpret_cast<RECT*>(lp);
            if (! rc)
            {
                break;
            }

            switch (wp)
            {
                case WMSZ_LEFT:
                case WMSZ_TOPLEFT:
                case WMSZ_BOTTOMLEFT: rc->left = rc->right - state->fixedWindowWidthPx; break;
                case WMSZ_RIGHT:
                case WMSZ_TOPRIGHT:
                case WMSZ_BOTTOMRIGHT: rc->right = rc->left + state->fixedWindowWidthPx; break;
                default: rc->right = rc->left + state->fixedWindowWidthPx; break;
            }

            return TRUE;
        }
        case WM_GETMINMAXINFO:
            if (auto* info = reinterpret_cast<MINMAXINFO*>(lp))
            {
                const LONG minWidth = state ? static_cast<LONG>(state->fixedWindowWidthPx) : 0L;
                Common::WindowSizing::ApplyMinimumClientTrackSizeForDips(dlg, *info, 380, 260);
                Common::WindowSizing::ApplyMinimumTrackSize(*info, minWidth, 0L);
            }
            return TRUE;
        case WM_SIZE:
        {
            if (state)
            {
                LayoutPluginConfigDialog(dlg, *state);
                if (state->panel)
                {
                    UpdatePanelScrollInfo(state->panel, *state);
                }
            }
            return TRUE;
        }
        case WM_SYSCHAR:
            if (state)
            {
                return HandlePluginConfigDialogDxMnemonic(*state, static_cast<wchar_t>(wp)) ? TRUE : FALSE;
            }
            return FALSE;
        case WM_COMMAND: return OnPluginConfigDialogCommand(dlg, state, LOWORD(wp), HIWORD(wp), reinterpret_cast<HWND>(lp));
    }

    return FALSE;
}

HRESULT
ShowPluginConfigurationDialogInternal(HWND owner,
                                      std::wstring_view appId,
                                      PluginType pluginType,
                                      std::wstring_view pluginId,
                                      std::wstring_view pluginName,
                                      Common::Settings::Settings& settings,
                                      const AppTheme& theme)
{
    if (pluginId.empty())
    {
        return E_INVALIDARG;
    }

    std::string schema;
    HRESULT schemaHr = E_FAIL;

    if (pluginType == PluginType::FileSystem)
    {
        auto& manager = FileSystemPluginManager::GetInstance();
        schemaHr      = manager.GetConfigurationSchema(pluginId, settings, schema);
    }
    else
    {
        auto& manager = ViewerPluginManager::GetInstance();
        schemaHr      = manager.GetConfigurationSchema(pluginId, settings, schema);
    }
    if (FAILED(schemaHr))
    {
        return schemaHr;
    }

    std::string current;

    // First, try to get configuration from settings.plugins.configurationByPluginId
    const std::wstring pluginIdText(pluginId);
    const auto it = settings.plugins.configurationByPluginId.find(pluginIdText);
    if (it != settings.plugins.configurationByPluginId.end() && ! std::holds_alternative<std::monostate>(it->second.value))
    {
        const HRESULT serializeHr = Common::Settings::SerializeJsonValue(it->second, current);
        if (FAILED(serializeHr))
        {
            return serializeHr;
        }
    }

    // If no configuration in settings, fall back to plugin's current configuration
    if (current.empty())
    {
        HRESULT configHr = E_FAIL;

        if (pluginType == PluginType::FileSystem)
        {
            auto& manager = FileSystemPluginManager::GetInstance();
            configHr      = manager.GetConfiguration(pluginId, settings, current);
        }
        else
        {
            auto& manager = ViewerPluginManager::GetInstance();
            configHr      = manager.GetConfiguration(pluginId, settings, current);
        }
        if (FAILED(configHr))
        {
            return configHr;
        }
    }

    PluginConfigDialogState state;
    state.settings              = &settings;
    state.reloadSourceSettings  = &settings;
    state.appId                 = std::wstring(appId);
    state.theme                 = theme;
    state.pluginType            = pluginType;
    state.pluginId              = std::wstring(pluginId);
    state.pluginName            = pluginName.empty() ? std::wstring(pluginId) : std::wstring(pluginName);
    state.schemaJsonUtf8        = std::move(schema);
    state.configurationJsonUtf8 = std::move(current);

    const INT_PTR result = RedSalamander::Win32Callback::DialogBoxParamResourceNoThrow(
        GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_PLUGIN_CONFIG), owner, PluginConfigDialogProc, reinterpret_cast<LPARAM>(&state));

    return result == IDOK ? S_OK : S_FALSE;
}

} // namespace

HRESULT ShowPluginConfigurationDialog(HWND owner,
                                      std::wstring_view appId,
                                      PluginType pluginType,
                                      std::wstring_view pluginId,
                                      std::wstring_view pluginName,
                                      Common::Settings::Settings& settings,
                                      const AppTheme& theme)
{
    return ShowPluginConfigurationDialogInternal(owner, appId, pluginType, pluginId, pluginName, settings, theme);
}

HRESULT EditPluginConfigurationDialog(HWND owner,
                                      PluginType pluginType,
                                      std::wstring_view pluginId,
                                      std::wstring_view pluginName,
                                      Common::Settings::Settings& baselineSettings,
                                      Common::Settings::Settings& inOutWorkingSettings,
                                      const AppTheme& theme)
{
    if (pluginId.empty())
    {
        return E_INVALIDARG;
    }

    std::string schema;
    HRESULT schemaHr = E_FAIL;

    if (pluginType == PluginType::FileSystem)
    {
        auto& manager = FileSystemPluginManager::GetInstance();
        schemaHr      = manager.GetConfigurationSchema(pluginId, baselineSettings, schema);
    }
    else
    {
        auto& manager = ViewerPluginManager::GetInstance();
        schemaHr      = manager.GetConfigurationSchema(pluginId, baselineSettings, schema);
    }
    if (FAILED(schemaHr))
    {
        return schemaHr;
    }

    std::string current;
    const std::wstring pluginIdText(pluginId);
    const auto it = inOutWorkingSettings.plugins.configurationByPluginId.find(pluginIdText);
    if (it != inOutWorkingSettings.plugins.configurationByPluginId.end() && ! std::holds_alternative<std::monostate>(it->second.value))
    {
        const HRESULT serializeHr = Common::Settings::SerializeJsonValue(it->second, current);
        if (FAILED(serializeHr))
        {
            return serializeHr;
        }
    }

    if (current.empty())
    {
        HRESULT configHr = E_FAIL;

        if (pluginType == PluginType::FileSystem)
        {
            auto& manager = FileSystemPluginManager::GetInstance();
            configHr      = manager.GetConfiguration(pluginId, baselineSettings, current);
        }
        else
        {
            auto& manager = ViewerPluginManager::GetInstance();
            configHr      = manager.GetConfiguration(pluginId, baselineSettings, current);
        }
        if (FAILED(configHr))
        {
            return configHr;
        }
    }

    PluginConfigDialogState state;
    state.settings              = &inOutWorkingSettings;
    state.reloadSourceSettings  = &baselineSettings;
    state.theme                 = theme;
    state.pluginType            = pluginType;
    state.pluginId              = std::wstring(pluginId);
    state.pluginName            = pluginName.empty() ? std::wstring(pluginId) : std::wstring(pluginName);
    state.schemaJsonUtf8        = std::move(schema);
    state.configurationJsonUtf8 = std::move(current);
    state.commitMode            = PluginConfigCommitMode::UpdateSettingsOnly;

    const INT_PTR result = RedSalamander::Win32Callback::DialogBoxParamResourceNoThrow(
        GetModuleHandleW(nullptr), MAKEINTRESOURCEW(IDD_PLUGIN_CONFIG), owner, PluginConfigDialogProc, reinterpret_cast<LPARAM>(&state));

    return result == IDOK ? S_OK : S_FALSE;
}

HWND GetPluginConfigurationDialogHandle() noexcept
{
    const HWND hwnd = g_pluginConfigurationDialogWindow.load(std::memory_order_acquire);
    return hwnd && IsWindow(hwnd) != FALSE ? hwnd : nullptr;
}

void UpdatePluginConfigurationWindowsTheme(const AppTheme& theme) noexcept
{
    const HWND hwnd = GetPluginConfigurationDialogHandle();
    if (! hwnd || IsWindow(hwnd) == FALSE)
    {
        return;
    }

    static_cast<void>(SendMessageW(hwnd, WndMsg::kPluginConfigurationDialogApplyTheme, 0, reinterpret_cast<LPARAM>(&theme)));
}

#ifdef ENABLE_TESTS
bool DebugGetPluginConfigurationDialogSnapshot(PluginConfigurationDialogDebugSnapshot& out) noexcept
{
    const HWND hwnd = GetPluginConfigurationDialogHandle();
    return hwnd && SendMessageW(hwnd,
                                WndMsg::kPluginConfigurationDialogDebug,
                                static_cast<WPARAM>(PluginConfigDebugCommand::GetSnapshot),
                                reinterpret_cast<LPARAM>(&out)) != FALSE;
}

bool DebugScrollPluginConfigurationDialogByWheelDetents(int wheelDetents) noexcept
{
    const HWND hwnd = GetPluginConfigurationDialogHandle();
    return hwnd && SendMessageW(hwnd,
                                WndMsg::kPluginConfigurationDialogDebug,
                                static_cast<WPARAM>(PluginConfigDebugCommand::ScrollByWheelDetents),
                                static_cast<LPARAM>(wheelDetents)) != FALSE;
}

bool DebugFocusPluginConfigurationDialogFirstInput() noexcept
{
    const HWND hwnd = GetPluginConfigurationDialogHandle();
    return hwnd && SendMessageW(hwnd, WndMsg::kPluginConfigurationDialogDebug, static_cast<WPARAM>(PluginConfigDebugCommand::FocusFirstInput), 0) != FALSE;
}

bool DebugGetPluginConfigurationDialogFirstVisibleToggleHostAndClientRect(HWND& outHost, RECT& outRect) noexcept
{
    outHost = nullptr;
    outRect = {};

    PluginConfigurationDialogDebugHostRect hostRect{};
    const HWND hwnd = GetPluginConfigurationDialogHandle();
    if (! hwnd)
    {
        return false;
    }

    if (SendMessageW(hwnd,
                     WndMsg::kPluginConfigurationDialogDebug,
                     static_cast<WPARAM>(PluginConfigDebugCommand::GetFirstVisibleToggleRect),
                     reinterpret_cast<LPARAM>(&hostRect)) == FALSE)
    {
        return false;
    }

    outHost = hostRect.host;
    outRect = hostRect.rect;
    return outHost != nullptr;
}

bool DebugGetPluginConfigurationDialogVisibleToggleHostAndClientRectByLabel(std::wstring_view label, HWND& outHost, RECT& outRect) noexcept
{
    outHost = nullptr;
    outRect = {};

    if (label.empty())
    {
        return false;
    }

    PluginConfigurationDialogDebugLabeledHostRect hostRect{};
    hostRect.label  = label.data();
    const HWND hwnd = GetPluginConfigurationDialogHandle();
    if (! hwnd)
    {
        return false;
    }

    if (SendMessageW(hwnd,
                     WndMsg::kPluginConfigurationDialogDebug,
                     static_cast<WPARAM>(PluginConfigDebugCommand::GetVisibleToggleRectByLabel),
                     reinterpret_cast<LPARAM>(&hostRect)) == FALSE)
    {
        return false;
    }

    outHost = hostRect.host;
    outRect = hostRect.rect;
    return outHost != nullptr;
}

bool DebugGetPluginConfigurationDialogFocusedHost(HWND& outHost) noexcept
{
    outHost = nullptr;

    PluginConfigurationDialogDebugHost focusedHost{};
    const HWND hwnd = GetPluginConfigurationDialogHandle();
    if (! hwnd)
    {
        return false;
    }

    if (SendMessageW(hwnd,
                     WndMsg::kPluginConfigurationDialogDebug,
                     static_cast<WPARAM>(PluginConfigDebugCommand::GetFocusedHost),
                     reinterpret_cast<LPARAM>(&focusedHost)) == FALSE)
    {
        return false;
    }

    outHost = focusedHost.host;
    return outHost != nullptr;
}

bool DebugAdvancePluginConfigurationDialogTab(bool reverse) noexcept
{
    const HWND hwnd = GetPluginConfigurationDialogHandle();
    if (! hwnd)
    {
#ifdef ENABLE_TESTS
        TracePluginConfigDebug(L"plugin-config advance wrapper: no dialog handle");
#endif
        return false;
    }

    const LRESULT result =
        SendMessageW(hwnd, WndMsg::kPluginConfigurationDialogDebug, static_cast<WPARAM>(PluginConfigDebugCommand::AdvanceTab), reverse ? 1 : 0);
#ifdef ENABLE_TESTS
    TracePluginConfigDebug(std::format(L"plugin-config advance wrapper: hwnd={:#x} reverse={} result={} finalFocus={:#x}",
                                       reinterpret_cast<uintptr_t>(hwnd),
                                       reverse ? 1 : 0,
                                       static_cast<long long>(result),
                                       reinterpret_cast<uintptr_t>(GetFocus())));
#endif
    return result != FALSE;
}

bool DebugCancelPluginConfigurationDialog() noexcept
{
    const HWND hwnd = GetPluginConfigurationDialogHandle();
    if (! hwnd)
    {
        return false;
    }

    if (SendMessageW(hwnd, WndMsg::kPluginConfigurationDialogDebug, static_cast<WPARAM>(PluginConfigDebugCommand::Cancel), 0) != FALSE)
    {
        return true;
    }

    return PostMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDCANCEL, 0), 0) != FALSE;
}
#endif
