#pragma once

// Single-canvas DxUi Connection Manager public surface and implementation API.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "AppTheme.h"
#include "SettingsStore.h"

// Shows the Connection Manager dialog facade.
// Returns:
// - S_OK: user chose a connection (selectedConnectionNameOut set)
// - S_FALSE: user cancelled (selectedConnectionNameOut cleared)
// - failure HRESULT: unexpected error
HRESULT ShowConnectionManagerDialog(HWND owner,
                                    std::wstring_view appId,
                                    Common::Settings::Settings& settings,
                                    const AppTheme& theme,
                                    std::wstring_view filterPluginId,
                                    std::wstring& selectedConnectionNameOut) noexcept;

// Shows a modeless Connection Manager window (similar to Preferences).
// `targetPane` is an app-defined identifier (0=Left, 1=Right) used when the user clicks Connect.
[[nodiscard]] bool ShowConnectionManagerWindow(HWND owner,
                                               std::wstring_view appId,
                                               Common::Settings::Settings& settings,
                                               const AppTheme& theme,
                                               std::wstring_view filterPluginId,
                                               uint8_t targetPane) noexcept;

[[nodiscard]] HWND GetConnectionManagerDialogHandle() noexcept;
void UpdateConnectionManagerWindowsTheme(const AppTheme& theme) noexcept;

#ifdef ENABLE_TESTS
enum class ConnectionManagerDebugFocusKind : uint8_t
{
    None,
    List,
    CommandButton,
    Edit,
    Combo,
    Toggle,
    FormActionButton,
};

struct ConnectionManagerDebugSnapshot
{
    bool usesDxUiCommandButtons                 = false;
    bool usesDxUiSectionHeaders                 = false;
    bool usesDxUiFormLabels                     = false;
    bool usesDxUiFormInputs                     = false;
    bool usesDxUiFormActionButtons              = false;
    bool usesDxUiList                           = false;
    size_t legacyOwnerDrawCommandButtonCount    = 0u;
    size_t legacyOwnerDrawFormInputCount        = 0u;
    size_t legacyOwnerDrawFormActionButtonCount = 0u;
    size_t visibleLegacyCommandButtonCount      = 0u;
    size_t visibleLegacySectionHeaderCount      = 0u;
    size_t visibleLegacyFormLabelCount          = 0u;
    size_t visibleLegacyFormInputCount          = 0u;
    size_t visibleLegacyFormActionButtonCount   = 0u;
    size_t visibleLegacyListCount               = 0u;
    size_t visibleDxSectionHeaderHostCount      = 0u;
    size_t visibleDxFormInputHostCount          = 0u;
    size_t visibleDxFormActionButtonHostCount   = 0u;
    size_t visibleDxListHostCount               = 0u;
    size_t listRowCount                         = 0u;
    size_t visibleListRowCount                  = 0u;
    size_t visibleListColumnCount               = 0u;
    size_t visibleListCellCount                 = 0u;
    bool listHasVerticalScrollbar               = false;
    bool listHasHorizontalScrollbar             = false;
    bool themeDark                              = false;
    bool themeHighContrast                      = false;
    bool themeRainbow                           = false;
    uint64_t dxListRenderCount                  = 0u;
    uint64_t dxListResizeCount                  = 0u;
    uint64_t dxListResizeFailureCount           = 0u;
    UINT dxModifierState                        = 0u;
    int selectedListIndex                       = -1;
    std::wstring selectedListRowName;
    uint32_t selectedListRowFillArgb          = 0u;
    uint32_t selectedListRowTextArgb          = 0u;
    bool selectedListRowUsesRainbow           = false;
    ConnectionManagerDebugFocusKind focusKind = ConnectionManagerDebugFocusKind::None;
    std::wstring focusLabel;
    bool focusControlPresent   = false;
    bool focusControlVisible   = false;
    bool focusControlEnabled   = false;
    bool focusControlFocusable = false;
    bool nativeFocusInDialog   = false;
    bool nativeFocusIsHost     = false;
    bool nativeFocusIsTextEdit = false;
    bool newButtonVisible      = false;
    bool newButtonEnabled      = false;
    bool newButtonFocusable    = false;
    bool renameButtonVisible   = false;
    bool renameButtonEnabled   = false;
    bool renameButtonFocusable = false;
    bool removeButtonVisible   = false;
    bool removeButtonEnabled   = false;
    bool removeButtonFocusable = false;
    std::wstring currentNameText;
    std::wstring currentPluginId;
    bool secretStoredPlaceholderVisible  = false;
    size_t secretStoredPlaceholderLength = 0u;
    bool secretMasked                    = false;
    bool showSecretButtonVisible         = false;
    bool showSecretButtonEnabled         = false;
    bool nameHostPresent                 = false;
    bool nameHostVisible                 = false;
    bool nameHostEnabled                 = false;
    bool nameLegacyVisible               = false;
    bool nameTextFieldPresent            = false;
    bool nameTextFieldVisible            = false;
    bool nameTextFieldEnabled            = false;
    bool nameHostFocusControlMatches     = false;
    bool nameHostOwnsFocus               = false;
    bool layoutListButtonsMirrorTopGap   = false;
    bool layoutCardsDoNotOverlap         = false;
    bool layoutSectionTitlesOutsideCards = false;
    bool layoutEditorKeepsRightGap       = false;
    bool layoutTextFieldsAvoidClipping   = false;
    float listTopGapDip                  = 0.0f;
    float listButtonBottomGapDip         = 0.0f;
    float editorRightGapDip              = 0.0f;
    float editorScrollbarGapDip          = 0.0f;
    float minVisibleCardGapDip           = 0.0f;
    float minVisibleTextFieldHeightDip   = 0.0f;
};

[[nodiscard]] bool DebugGetConnectionManagerDialogSnapshot(ConnectionManagerDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugClickConnectionManagerListRow(size_t rowIndex) noexcept;
[[nodiscard]] bool DebugScrollConnectionManagerListByWheelDetents(int detents) noexcept;
[[nodiscard]] bool DebugFocusConnectionManagerFirstInput() noexcept;
[[nodiscard]] bool DebugFocusConnectionManagerUserInput() noexcept;
[[nodiscard]] bool DebugFocusConnectionManagerSecretInput() noexcept;
[[nodiscard]] bool DebugFocusConnectionManagerList() noexcept;
[[nodiscard]] bool DebugRouteConnectionManagerMnemonic(wchar_t mnemonic) noexcept;
[[nodiscard]] bool DebugRouteConnectionManagerCommandKey(WPARAM virtualKey) noexcept;
[[nodiscard]] bool DebugRouteConnectionManagerTab(bool reverse) noexcept;
[[nodiscard]] bool DebugSetConnectionManagerProtocolPluginId(std::wstring_view pluginId) noexcept;
[[nodiscard]] bool DebugGetConnectionManagerAlternateProtocolPluginId(std::wstring_view baselinePluginId, std::wstring& outPluginId) noexcept;
[[nodiscard]] bool DebugSetConnectionManagerNameText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugSetConnectionManagerUserText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugSetConnectionManagerSecretText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugGetConnectionManagerUserText(std::wstring& outText) noexcept;
[[nodiscard]] bool DebugGetConnectionManagerSecretText(std::wstring& outText) noexcept;
[[nodiscard]] bool DebugGetConnectionManagerTextInputHandle(HWND& outInput) noexcept;
[[nodiscard]] bool DebugGetConnectionManagerUserInputCaretClientPoint(size_t caretIndex, HWND& outHost, POINT& outPoint) noexcept;
[[nodiscard]] bool DebugGetConnectionManagerUserTextSelection(size_t& outStart, size_t& outEnd) noexcept;
[[nodiscard]] bool DebugDoubleClickConnectionManagerUserTextAtCaretIndex(size_t caretIndex) noexcept;
[[nodiscard]] bool DebugGetConnectionManagerListHostHandle(HWND& outHost) noexcept;
[[nodiscard]] bool DebugGetConnectionManagerNameHostHandle(HWND& outHost) noexcept;
[[nodiscard]] bool DebugAcknowledgeConnectionManagerS3InsecureTlsPrompt() noexcept;
[[nodiscard]] bool DebugScrollConnectionManagerSavePasswordToggleIntoView() noexcept;
[[nodiscard]] bool DebugScrollConnectionManagerS3UseHttpsToggleIntoView() noexcept;
[[nodiscard]] bool DebugGetConnectionManagerSavePasswordToggleHostAndClientRect(HWND& outHost, RECT& outRect) noexcept;
[[nodiscard]] bool DebugGetConnectionManagerSavePasswordToggleState(bool& outChecked, std::wstring& outLabel) noexcept;
[[nodiscard]] bool DebugScrollConnectionManagerCommandButtonIntoView(UINT commandId) noexcept;
[[nodiscard]] bool DebugGetConnectionManagerCommandButtonHostAndClientRect(UINT commandId, HWND& outHost, RECT& outRect, std::wstring& outLabel) noexcept;
[[nodiscard]] bool DebugGetConnectionManagerS3UseHttpsToggleHostAndClientRect(HWND& outHost, RECT& outRect) noexcept;
[[nodiscard]] bool DebugGetConnectionManagerS3UseHttpsToggleState(bool& outChecked, std::wstring& outLabel) noexcept;
#endif

namespace RedSalamander::ConnectionManager::SingleCanvas
{
// The single-canvas implementation is the live Connection Manager path.
[[nodiscard]] constexpr bool IsEnabled() noexcept
{
    return true;
}

// Modeless single-instance entry point, modeled on `ShowConnectionManagerWindow`.
// Returns false if the window cannot be created.
[[nodiscard]] bool ShowWindow(HWND owner,
                              std::wstring_view appId,
                              Common::Settings::Settings& settings,
                              const AppTheme& theme,
                              std::wstring_view filterPluginId,
                              uint8_t targetPane) noexcept;

// Synchronous facade for callers that still use `ShowConnectionManagerDialog`.
// Internally creates a modeless window, disables the owner, and runs a private
// message pump until the user picks a connection (S_OK), cancels (S_FALSE), or
// hits an error (failure HRESULT).
[[nodiscard]] HRESULT ShowDialog(HWND owner,
                                 std::wstring_view appId,
                                 Common::Settings::Settings& settings,
                                 const AppTheme& theme,
                                 std::wstring_view filterPluginId,
                                 std::wstring& selectedConnectionNameOut) noexcept;

// Returns the live HWND of the single-canvas window, or nullptr if no instance
// is currently shown. Mirrors `GetConnectionManagerDialogHandle`.
[[nodiscard]] HWND GetWindowHandle() noexcept;

void UpdateTheme(const AppTheme& theme) noexcept;

#ifdef ENABLE_TESTS
[[nodiscard]] bool DebugGetSnapshot(::ConnectionManagerDebugSnapshot& out) noexcept;
[[nodiscard]] bool DebugClickListRow(size_t rowIndex) noexcept;
[[nodiscard]] bool DebugScrollListByWheelDetents(int detents) noexcept;
[[nodiscard]] bool DebugFocusFirstInput() noexcept;
[[nodiscard]] bool DebugFocusUserInput() noexcept;
[[nodiscard]] bool DebugFocusSecretInput() noexcept;
[[nodiscard]] bool DebugFocusList() noexcept;
[[nodiscard]] bool DebugRouteMnemonic(wchar_t mnemonic) noexcept;
[[nodiscard]] bool DebugRouteCommandKey(WPARAM virtualKey) noexcept;
[[nodiscard]] bool DebugRouteTab(bool reverse) noexcept;
[[nodiscard]] bool DebugSetProtocolPluginId(std::wstring_view pluginId) noexcept;
[[nodiscard]] bool DebugGetAlternateProtocolPluginId(std::wstring_view baselinePluginId, std::wstring& outPluginId) noexcept;
[[nodiscard]] bool DebugSetNameText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugSetUserText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugSetSecretText(std::wstring_view text) noexcept;
[[nodiscard]] bool DebugGetUserText(std::wstring& outText) noexcept;
[[nodiscard]] bool DebugGetSecretText(std::wstring& outText) noexcept;
[[nodiscard]] bool DebugGetTextInputHandle(HWND& outInput) noexcept;
[[nodiscard]] bool DebugGetUserInputCaretClientPoint(size_t caretIndex, HWND& outHost, POINT& outPoint) noexcept;
[[nodiscard]] bool DebugGetUserTextSelection(size_t& outStart, size_t& outEnd) noexcept;
[[nodiscard]] bool DebugDoubleClickUserTextAtCaretIndex(size_t caretIndex) noexcept;
[[nodiscard]] bool DebugGetListHostHandle(HWND& outHost) noexcept;
[[nodiscard]] bool DebugGetNameHostHandle(HWND& outHost) noexcept;
[[nodiscard]] bool DebugAcknowledgeS3InsecureTlsPrompt() noexcept;
[[nodiscard]] bool DebugScrollS3UseHttpsToggleIntoView() noexcept;
[[nodiscard]] bool DebugGetSavePasswordToggleHostAndClientRect(HWND& outHost, RECT& outRect) noexcept;
[[nodiscard]] bool DebugGetSavePasswordToggleState(bool& outChecked, std::wstring& outLabel) noexcept;
[[nodiscard]] bool DebugScrollCommandButtonIntoView(UINT commandId) noexcept;
[[nodiscard]] bool DebugGetCommandButtonHostAndClientRect(UINT commandId, HWND& outHost, RECT& outRect, std::wstring& outLabel) noexcept;
[[nodiscard]] bool DebugGetS3UseHttpsToggleHostAndClientRect(HWND& outHost, RECT& outRect) noexcept;
[[nodiscard]] bool DebugGetS3UseHttpsToggleState(bool& outChecked, std::wstring& outLabel) noexcept;
#endif

} // namespace RedSalamander::ConnectionManager::SingleCanvas
