// Preferences.HotPaths.cpp

#include "Framework.h"

#include "Preferences.HotPaths.h"

#include <format>
#include <string>

#include <shobjidl.h>

#include <wil/com.h>

#include "Helpers.h"
#include "ThemedControls.h"
#include "resource.h"

namespace
{
using PrefsHotPaths::EnsureWorkingHotPathsSettings;
using PrefsHotPaths::GetHotPathsSettingsOrDefault;
using PrefsHotPaths::MaybeResetWorkingHotPathsSettingsIfEmpty;

constexpr int kSlotCount = 10;
} // namespace

bool HotPathsPane::EnsureCreated(HWND pageHost) noexcept
{
    return PrefsPaneHost::EnsureCreated(pageHost, _hWnd);
}

void HotPathsPane::ResizeToHostClient(HWND pageHost) noexcept
{
    PrefsPaneHost::ResizeToHostClient(pageHost, _hWnd.get());
}

void HotPathsPane::Show(bool visible) noexcept
{
    PrefsPaneHost::Show(_hWnd.get(), visible);
}

void HotPathsPane::CreateControls(HWND parent, PreferencesDialogState& state) noexcept
{
    if (! parent)
    {
        return;
    }

    const DWORD baseStaticStyle = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX;
    const DWORD wrapStaticStyle = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX | SS_EDITCONTROL;
    const bool customButtons    = ! state.theme.systemHighContrast;

    const DWORD toggleStyle = customButtons ? (WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW) : (WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX);

    const HINSTANCE instance = GetModuleHandleW(nullptr);

    state.hotPathSlotControls.resize(kSlotCount);

    for (int i = 0; i < kSlotCount; ++i)
    {
        auto& slot = state.hotPathSlotControls[static_cast<size_t>(i)];

        // Header label: "Ctrl+1" etc.
        slot.header.reset(CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));

        // Path label + framed edit + browse button
        slot.pathLabel.reset(CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));

        const int pathEditId = IDC_PREFS_HOT_PATHS_PATH_EDIT_BASE + i;
        PrefsInput::CreateFramedEditBox(state, parent, slot.pathFrame, slot.pathEdit, pathEditId, WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL);

        const int browseId = IDC_PREFS_HOT_PATHS_BROWSE_BASE + i;
        slot.browseButton.reset(CreateWindowExW(
            0, L"Button", L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON, 0, 0, 10, 10, parent,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(browseId)), instance, nullptr));

        // Label label + framed edit
        slot.labelLabel.reset(CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));

        const int labelEditId = IDC_PREFS_HOT_PATHS_LABEL_EDIT_BASE + i;
        PrefsInput::CreateFramedEditBox(state, parent, slot.labelFrame, slot.labelEdit, labelEditId, WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL);

        // Show in menu toggle
        const int showInMenuId = IDC_PREFS_HOT_PATHS_SHOW_IN_MENU_BASE + i;
        const std::wstring showLabel = customButtons ? std::wstring{} : LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_SHOW_IN_MENU);
        slot.showInMenuToggle.reset(CreateWindowExW(
            0, L"Button", showLabel.c_str(), toggleStyle, 0, 0, 10, 10, parent,
            reinterpret_cast<HMENU>(static_cast<INT_PTR>(showInMenuId)), instance, nullptr));
        PrefsInput::EnableMouseWheelForwarding(slot.showInMenuToggle);

        slot.showInMenuLabel.reset(CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));
        slot.showInMenuDescription.reset(CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));
    }

    // Open prefs on assign toggle
    const std::wstring assignLabel = customButtons ? std::wstring{} : LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_OPEN_PREFS_ON_ASSIGN);
    state.hotPathOpenPrefsOnAssignToggle.reset(CreateWindowExW(
        0, L"Button", assignLabel.c_str(), toggleStyle, 0, 0, 10, 10, parent,
        reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_HOT_PATHS_OPEN_PREFS_ON_ASSIGN)), instance, nullptr));
    PrefsInput::EnableMouseWheelForwarding(state.hotPathOpenPrefsOnAssignToggle);

    state.hotPathOpenPrefsOnAssignLabel.reset(CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));
    state.hotPathOpenPrefsOnAssignDescription.reset(
        CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));

    Refresh(parent, state);
}

void HotPathsPane::Refresh(HWND /*host*/, PreferencesDialogState& state) noexcept
{
    const auto& hp = GetHotPathsSettingsOrDefault(state.workingSettings);

    const auto setEnabledAndInvalidate = [](const auto& hwndLike, BOOL enabled) noexcept
    {
        HWND hwnd = nullptr;
        if constexpr (requires { hwndLike.get(); })
        {
            hwnd = hwndLike.get();
        }
        else
        {
            hwnd = hwndLike;
        }
        if (hwnd)
        {
            EnableWindow(hwnd, enabled);
            InvalidateRect(hwnd, nullptr, TRUE);
        }
    };

    for (int i = 0; i < kSlotCount && i < static_cast<int>(state.hotPathSlotControls.size()); ++i)
    {
        const auto& slotCtl  = state.hotPathSlotControls[static_cast<size_t>(i)];
        const auto& slotData = hp.slots[static_cast<size_t>(i)];
        const BOOL enableDependentControls = (slotData.has_value() && ! slotData.value().path.empty()) ? TRUE : FALSE;

        if (slotCtl.pathEdit)
        {
            SetWindowTextW(slotCtl.pathEdit.get(), slotData.has_value() ? slotData.value().path.c_str() : L"");
        }
        if (slotCtl.labelEdit)
        {
            SetWindowTextW(slotCtl.labelEdit.get(), slotData.has_value() ? slotData.value().label.c_str() : L"");
        }
        if (slotCtl.showInMenuToggle)
        {
            const bool checked = slotData.has_value() && slotData.value().showInMenu;
            PrefsUi::SetTwoStateToggleState(slotCtl.showInMenuToggle, state.theme.systemHighContrast, checked);
        }

        setEnabledAndInvalidate(slotCtl.labelLabel, enableDependentControls);
        setEnabledAndInvalidate(slotCtl.labelFrame, enableDependentControls);
        setEnabledAndInvalidate(slotCtl.labelEdit, enableDependentControls);
        setEnabledAndInvalidate(slotCtl.showInMenuLabel, enableDependentControls);
        setEnabledAndInvalidate(slotCtl.showInMenuToggle, enableDependentControls);
        setEnabledAndInvalidate(slotCtl.showInMenuDescription, enableDependentControls);
    }

    PrefsUi::SetTwoStateToggleState(state.hotPathOpenPrefsOnAssignToggle, state.theme.systemHighContrast, hp.openPrefsOnAssign);
}

void HotPathsPane::LayoutControls(HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, HFONT dialogFont) noexcept
{
    using namespace PrefsLayoutConstants;

    if (! host)
    {
        return;
    }

    const UINT dpi = GetDpiForWindow(host);

    const int rowHeight    = std::max(1, ThemedControls::ScaleDip(dpi, kRowHeightDip));
    const int titleHeight  = std::max(1, ThemedControls::ScaleDip(dpi, kTitleHeightDip));
    const int editHeight   = std::max(1, ThemedControls::ScaleDip(dpi, kEditHeightDip));
    const int headerHeight = std::max(1, ThemedControls::ScaleDip(dpi, kHeaderHeightDip));

    const int cardPaddingX = ThemedControls::ScaleDip(dpi, kCardPaddingXDip);
    const int cardPaddingY = ThemedControls::ScaleDip(dpi, kCardPaddingYDip);
    const int cardGapY     = ThemedControls::ScaleDip(dpi, kCardGapYDip);
    const int cardGapX     = ThemedControls::ScaleDip(dpi, kCardGapXDip);
    const int cardSpacingY = ThemedControls::ScaleDip(dpi, kCardSpacingYDip);

    const int browseWidth = std::max(1, ThemedControls::ScaleDip(dpi, 75));
    const int browseGap   = std::max(1, ThemedControls::ScaleDip(dpi, 4));
    const int innerGap    = std::max(2, gapY / 2);

    const HFONT headerFont = state.boldFont ? state.boldFont.get() : dialogFont;
    const HFONT infoFont   = state.italicFont ? state.italicFont.get() : dialogFont;

    const int minToggleWidth    = ThemedControls::ScaleDip(dpi, kMinToggleWidthDip);
    const std::wstring onLabel  = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);
    const std::wstring offLabel = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);

    const HFONT toggleMeasureFont = state.boldFont ? state.boldFont.get() : dialogFont;
    const int onWidth             = ThemedControls::MeasureTextWidth(host, toggleMeasureFont, onLabel);
    const int offWidth            = ThemedControls::MeasureTextWidth(host, toggleMeasureFont, offLabel);

    const int tPaddingX       = ThemedControls::ScaleDip(dpi, kTogglePaddingXDip);
    const int tGapX           = ThemedControls::ScaleDip(dpi, kToggleGapXDip);
    const int trackWidth      = ThemedControls::ScaleDip(dpi, kToggleTrackWidthDip);
    const int stateTextWidth  = std::max(onWidth, offWidth);

    const int measuredToggleWidth = std::max(minToggleWidth, (2 * tPaddingX) + stateTextWidth + tGapX + trackWidth);
    const int toggleWidth         = std::min(std::max(0, width - 2 * cardPaddingX), measuredToggleWidth);

    auto pushCard = [&](const RECT& card) noexcept { state.pageSettingCards.push_back(card); };

    auto layoutToggleCard = [&](int cardX, int cardWidth, HWND label, std::wstring_view labelText, HWND toggle, HWND descLabel, std::wstring_view descText) noexcept
    {
        const int textWidth  = std::max(0, cardWidth - 2 * cardPaddingX - cardGapX - toggleWidth);
        const int descHeight = descLabel ? PrefsUi::MeasureStaticTextHeight(host, infoFont, textWidth, descText) : 0;

        const int contentHeight = std::max(0, titleHeight + cardGapY + descHeight);
        const int cardHeight    = std::max(rowHeight + 2 * cardPaddingY, contentHeight + 2 * cardPaddingY);

        RECT card{};
        card.left   = cardX;
        card.top    = y;
        card.right  = cardX + cardWidth;
        card.bottom = y + cardHeight;
        pushCard(card);

        if (label)
        {
            SetWindowTextW(label, labelText.data());
            SetWindowPos(label, nullptr, card.left + cardPaddingX, card.top + cardPaddingY, textWidth, titleHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }

        if (descLabel)
        {
            const wchar_t* descTextPtr = descText.empty() ? L"" : descText.data();
            SetWindowTextW(descLabel, descTextPtr);
            SetWindowPos(descLabel,
                         nullptr,
                         card.left + cardPaddingX,
                         card.top + cardPaddingY + titleHeight + cardGapY,
                         textWidth,
                         std::max(0, descHeight),
                         SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(descLabel, WM_SETFONT, reinterpret_cast<WPARAM>(infoFont), TRUE);
        }

        if (toggle)
        {
            SetWindowPos(toggle,
                         nullptr,
                         card.right - cardPaddingX - toggleWidth,
                         card.top + (cardHeight - rowHeight) / 2,
                         toggleWidth,
                         rowHeight,
                         SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(toggle, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }

        y += cardHeight + cardSpacingY;
    };

    for (int i = 0; i < kSlotCount && i < static_cast<int>(state.hotPathSlotControls.size()); ++i)
    {
        auto& slotCtl = state.hotPathSlotControls[static_cast<size_t>(i)];

        // Slot header: "Ctrl+1" etc.
        const wchar_t digitChar = (i < 9) ? static_cast<wchar_t>(L'1' + i) : L'0';
        const std::wstring headerText = FormatStringResource(nullptr, IDS_PREFS_HOT_PATHS_SLOT_HEADER_FMT, digitChar);

        if (slotCtl.header)
        {
            SetWindowTextW(slotCtl.header.get(), headerText.c_str());
            SetWindowPos(slotCtl.header.get(), nullptr, x, y, width, headerHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(slotCtl.header.get(), WM_SETFONT, reinterpret_cast<WPARAM>(headerFont), TRUE);
            y += headerHeight + innerGap;
        }

        // Path row: label + edit + browse
        const std::wstring pathLabel = LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_PATH_LABEL);
        if (slotCtl.pathLabel)
        {
            SetWindowTextW(slotCtl.pathLabel.get(), pathLabel.c_str());
            SetWindowPos(slotCtl.pathLabel.get(), nullptr, x + margin, y, width - margin, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(slotCtl.pathLabel.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            y += rowHeight;
        }

        const int editWidth = width - margin - browseWidth - browseGap;
        if (slotCtl.pathFrame)
        {
            SetWindowPos(slotCtl.pathFrame.get(), nullptr, x + margin, y, std::max(10, editWidth), editHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        }
        if (slotCtl.pathEdit)
        {
            const int frameInset = ThemedControls::ScaleDip(dpi, 2);
            SetWindowPos(slotCtl.pathEdit.get(), nullptr, x + margin + frameInset, y + frameInset, std::max(4, editWidth - 2 * frameInset),
                         std::max(4, editHeight - 2 * frameInset), SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(slotCtl.pathEdit.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }

        const std::wstring browseText = LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_BROWSE_ELLIPSIS);
        if (slotCtl.browseButton)
        {
            SetWindowTextW(slotCtl.browseButton.get(), browseText.c_str());
            SetWindowPos(slotCtl.browseButton.get(), nullptr, x + margin + editWidth + browseGap, y, browseWidth, editHeight,
                         SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(slotCtl.browseButton.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }
        y += editHeight + innerGap;

        // Label row: label + edit
        const std::wstring labelLabel = LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_LABEL_LABEL);
        if (slotCtl.labelLabel)
        {
            SetWindowTextW(slotCtl.labelLabel.get(), labelLabel.c_str());
            SetWindowPos(slotCtl.labelLabel.get(), nullptr, x + margin, y, width - margin, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(slotCtl.labelLabel.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            y += rowHeight;
        }

        if (slotCtl.labelFrame)
        {
            SetWindowPos(slotCtl.labelFrame.get(), nullptr, x + margin, y, std::max(10, editWidth), editHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        }
        if (slotCtl.labelEdit)
        {
            const int frameInset = ThemedControls::ScaleDip(dpi, 2);
            SetWindowPos(slotCtl.labelEdit.get(), nullptr, x + margin + frameInset, y + frameInset, std::max(4, editWidth - 2 * frameInset),
                         std::max(4, editHeight - 2 * frameInset), SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(slotCtl.labelEdit.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }
        y += editHeight + innerGap;

        // Show in menu toggle card
        const std::wstring showLabel = LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_SHOW_IN_MENU);
        layoutToggleCard(x + margin, width - margin, slotCtl.showInMenuLabel.get(), showLabel,
                         slotCtl.showInMenuToggle.get(), slotCtl.showInMenuDescription.get(), std::wstring_view{});

        y += gapY;
    }

    // Open prefs on assign
    const std::wstring assignLabel = LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_OPEN_PREFS_ON_ASSIGN);
    const std::wstring assignDesc  = LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_OPEN_PREFS_ON_ASSIGN_DESC);
    layoutToggleCard(x, width, state.hotPathOpenPrefsOnAssignLabel.get(), assignLabel,
                     state.hotPathOpenPrefsOnAssignToggle.get(), state.hotPathOpenPrefsOnAssignDescription.get(), assignDesc);
}

bool HotPathsPane::HandleCommand(HWND host, PreferencesDialogState& state, UINT commandId, UINT notifyCode, HWND hwndCtl) noexcept
{
    // Handle path and label edit changes.
    const bool isPathEdit  = (commandId >= static_cast<UINT>(IDC_PREFS_HOT_PATHS_PATH_EDIT_BASE) &&
                             commandId < static_cast<UINT>(IDC_PREFS_HOT_PATHS_PATH_EDIT_BASE + kSlotCount));
    const bool isLabelEdit = (commandId >= static_cast<UINT>(IDC_PREFS_HOT_PATHS_LABEL_EDIT_BASE) &&
                              commandId < static_cast<UINT>(IDC_PREFS_HOT_PATHS_LABEL_EDIT_BASE + kSlotCount));

    if (isPathEdit || isLabelEdit)
    {
        if (notifyCode == EN_CHANGE || notifyCode == EN_KILLFOCUS)
        {
            HWND edit               = hwndCtl ? hwndCtl : GetDlgItem(host, static_cast<int>(commandId));
            const std::wstring text = PrefsUi::GetWindowTextString(edit);
            const std::wstring_view trimmed = PrefsUi::TrimWhitespace(text);
            const bool commit               = (notifyCode == EN_KILLFOCUS);

            auto* hp = EnsureWorkingHotPathsSettings(state.workingSettings);
            if (! hp)
            {
                return true;
            }

            int slotIdx = 0;
            if (isPathEdit)
            {
                slotIdx = static_cast<int>(commandId) - IDC_PREFS_HOT_PATHS_PATH_EDIT_BASE;
            }
            else
            {
                slotIdx = static_cast<int>(commandId) - IDC_PREFS_HOT_PATHS_LABEL_EDIT_BASE;
            }

            if (slotIdx < 0 || slotIdx >= kSlotCount)
            {
                return true;
            }

            const size_t idx = static_cast<size_t>(slotIdx);
            bool changed     = false;

            if (isPathEdit)
            {
                const std::wstring newValue(trimmed);
                if (newValue.empty())
                {
                    if (hp->slots[idx].has_value())
                    {
                        if (commit)
                        {
                            hp->slots[idx].reset();
                            changed = true;
                        }
                        else if (! hp->slots[idx].value().path.empty())
                        {
                            hp->slots[idx].value().path.clear();
                            changed = true;
                        }
                    }
                }
                else
                {
                    if (! hp->slots[idx].has_value())
                    {
                        hp->slots[idx] = Common::Settings::HotPathSlot{};
                    }
                    if (hp->slots[idx].value().path != newValue)
                    {
                        hp->slots[idx].value().path = newValue;
                        changed                     = true;
                    }
                }

                const bool hasPathNow = (hp->slots[idx].has_value() && ! hp->slots[idx].value().path.empty());
                const BOOL enable     = hasPathNow ? TRUE : FALSE;
                auto& slotCtl         = state.hotPathSlotControls[idx];
                EnableWindow(slotCtl.labelLabel.get(), enable);
                EnableWindow(slotCtl.labelFrame.get(), enable);
                EnableWindow(slotCtl.labelEdit.get(), enable);
                EnableWindow(slotCtl.showInMenuLabel.get(), enable);
                EnableWindow(slotCtl.showInMenuToggle.get(), enable);
                EnableWindow(slotCtl.showInMenuDescription.get(), enable);
            }
            else
            {
                const std::wstring newValue(trimmed);
                if (hp->slots[idx].has_value())
                {
                    if (hp->slots[idx].value().label != newValue)
                    {
                        hp->slots[idx].value().label = newValue;
                        changed                      = true;
                    }
                }
                else if (! newValue.empty())
                {
                    hp->slots[idx]               = Common::Settings::HotPathSlot{};
                    hp->slots[idx].value().label = newValue;
                    changed                      = true;
                }
            }

            if (changed)
            {
                MaybeResetWorkingHotPathsSettingsIfEmpty(state.workingSettings);
                SetDirty(GetParent(host), state);
            }

            if (commit)
            {
                Refresh(host, state);
            }

            return true;
        }

        return false;
    }

    // Handle browse buttons.
    const bool isBrowse = (commandId >= static_cast<UINT>(IDC_PREFS_HOT_PATHS_BROWSE_BASE) &&
                           commandId < static_cast<UINT>(IDC_PREFS_HOT_PATHS_BROWSE_BASE + kSlotCount));
    if (isBrowse && notifyCode == BN_CLICKED)
    {
        const int slotIdx = static_cast<int>(commandId) - IDC_PREFS_HOT_PATHS_BROWSE_BASE;
        if (slotIdx < 0 || slotIdx >= kSlotCount)
        {
            return true;
        }

        wil::com_ptr<IFileOpenDialog> dialog;
        HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&dialog));
        if (FAILED(hr))
        {
            return true;
        }

        DWORD options = 0;
        dialog->GetOptions(&options);
        dialog->SetOptions(options | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);

        hr = dialog->Show(host);
        if (FAILED(hr))
        {
            return true;
        }

        wil::com_ptr<IShellItem> result;
        hr = dialog->GetResult(&result);
        if (FAILED(hr) || ! result)
        {
            return true;
        }

        wil::unique_cotaskmem_string pathStr;
        hr = result->GetDisplayName(SIGDN_FILESYSPATH, &pathStr);
        if (FAILED(hr) || ! pathStr)
        {
            return true;
        }

        const size_t idx = static_cast<size_t>(slotIdx);
        auto* hp         = EnsureWorkingHotPathsSettings(state.workingSettings);
        if (! hp)
        {
            return true;
        }

        if (! hp->slots[idx].has_value())
        {
            hp->slots[idx] = Common::Settings::HotPathSlot{};
        }
        hp->slots[idx].value().path = pathStr.get();

        MaybeResetWorkingHotPathsSettingsIfEmpty(state.workingSettings);
        SetDirty(GetParent(host), state);
        Refresh(host, state);
        return true;
    }

    // Handle show-in-menu toggles.
    const bool isShowInMenu = (commandId >= static_cast<UINT>(IDC_PREFS_HOT_PATHS_SHOW_IN_MENU_BASE) &&
                               commandId < static_cast<UINT>(IDC_PREFS_HOT_PATHS_SHOW_IN_MENU_BASE + kSlotCount));
    if (isShowInMenu && notifyCode == BN_CLICKED)
    {
        const int slotIdx = static_cast<int>(commandId) - IDC_PREFS_HOT_PATHS_SHOW_IN_MENU_BASE;
        if (slotIdx < 0 || slotIdx >= kSlotCount)
        {
            return true;
        }

        const bool toggledOn = PrefsUi::GetTwoStateToggleState(hwndCtl, state.theme.systemHighContrast);
        const size_t idx     = static_cast<size_t>(slotIdx);

        auto* hp = EnsureWorkingHotPathsSettings(state.workingSettings);
        if (! hp)
        {
            return true;
        }

        if (! hp->slots[idx].has_value())
        {
            hp->slots[idx] = Common::Settings::HotPathSlot{};
        }
        hp->slots[idx].value().showInMenu = toggledOn;

        MaybeResetWorkingHotPathsSettingsIfEmpty(state.workingSettings);
        SetDirty(GetParent(host), state);
        Refresh(host, state);
        return true;
    }

    // Handle open-prefs-on-assign toggle.
    if (commandId == IDC_PREFS_HOT_PATHS_OPEN_PREFS_ON_ASSIGN && notifyCode == BN_CLICKED)
    {
        const bool toggledOn = PrefsUi::GetTwoStateToggleState(hwndCtl, state.theme.systemHighContrast);

        auto* hp = EnsureWorkingHotPathsSettings(state.workingSettings);
        if (! hp)
        {
            return true;
        }

        hp->openPrefsOnAssign = toggledOn;
        MaybeResetWorkingHotPathsSettingsIfEmpty(state.workingSettings);
        SetDirty(GetParent(host), state);
        Refresh(host, state);
        return true;
    }

    return false;
}

// Namespace helper implementations.
namespace PrefsHotPaths
{
const Common::Settings::HotPathsSettings& GetHotPathsSettingsOrDefault(const Common::Settings::Settings& settings) noexcept
{
    if (settings.hotPaths.has_value())
    {
        return settings.hotPaths.value();
    }
    static const Common::Settings::HotPathsSettings defaults{};
    return defaults;
}

Common::Settings::HotPathsSettings* EnsureWorkingHotPathsSettings(Common::Settings::Settings& settings) noexcept
{
    if (! settings.hotPaths.has_value())
    {
        settings.hotPaths = Common::Settings::HotPathsSettings{};
    }
    return &settings.hotPaths.value();
}

void MaybeResetWorkingHotPathsSettingsIfEmpty(Common::Settings::Settings& settings) noexcept
{
    if (! settings.hotPaths.has_value())
    {
        return;
    }

    const auto& hp = settings.hotPaths.value();

    bool hasAnySlot = false;
    for (const auto& slot : hp.slots)
    {
        if (slot.has_value() && (! slot.value().path.empty() || ! slot.value().label.empty() || slot.value().showInMenu))
        {
            hasAnySlot = true;
            break;
        }
    }

    if (! hasAnySlot && ! hp.openPrefsOnAssign)
    {
        settings.hotPaths.reset();
    }
}
} // namespace PrefsHotPaths
