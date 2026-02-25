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

    const DWORD baseStaticStyle         = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX;
    const DWORD captionStaticStyle      = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX | SS_CENTERIMAGE;
    const DWORD rightCaptionStaticStyle = WS_CHILD | WS_VISIBLE | SS_RIGHT | SS_NOPREFIX | SS_CENTERIMAGE;
    const DWORD wrapStaticStyle         = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX | SS_EDITCONTROL;
    const bool customButtons            = ! state.theme.systemHighContrast;

    const DWORD toggleStyle = customButtons ? (WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW) : (WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX);

    const HINSTANCE instance = GetModuleHandleW(nullptr);

    state.hotPathSlotControls.resize(kSlotCount);

    for (int i = 0; i < kSlotCount; ++i)
    {
        auto& slot = state.hotPathSlotControls[static_cast<size_t>(i)];

        // Header label: "Ctrl+1" etc.
        slot.header.reset(CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));

        // Path label + framed edit + browse button
        slot.pathLabel.reset(CreateWindowExW(0, L"Static", L"", captionStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));

        const int pathEditId = IDC_PREFS_HOT_PATHS_PATH_EDIT_BASE + i;
        PrefsInput::CreateFramedEditBox(state, parent, slot.pathFrame, slot.pathEdit, pathEditId, WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL);

        const int browseId = IDC_PREFS_HOT_PATHS_BROWSE_BASE + i;
        slot.browseButton.reset(CreateWindowExW(0,
                                                L"Button",
                                                L"",
                                                WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
                                                0,
                                                0,
                                                10,
                                                10,
                                                parent,
                                                reinterpret_cast<HMENU>(static_cast<INT_PTR>(browseId)),
                                                instance,
                                                nullptr));
        if (slot.browseButton && customButtons)
        {
            ThemedControls::EnableOwnerDrawButton(parent, browseId);
        }

        // Label label + framed edit
        slot.labelLabel.reset(CreateWindowExW(0, L"Static", L"", captionStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));

        const int labelEditId = IDC_PREFS_HOT_PATHS_LABEL_EDIT_BASE + i;
        PrefsInput::CreateFramedEditBox(state, parent, slot.labelFrame, slot.labelEdit, labelEditId, WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL);

        // Show in menu toggle
        const int showInMenuId       = IDC_PREFS_HOT_PATHS_SHOW_IN_MENU_BASE + i;
        const std::wstring showLabel = customButtons ? std::wstring{} : LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_SHOW_IN_MENU);
        slot.showInMenuToggle.reset(CreateWindowExW(0,
                                                    L"Button",
                                                    showLabel.c_str(),
                                                    toggleStyle,
                                                    0,
                                                    0,
                                                    10,
                                                    10,
                                                    parent,
                                                    reinterpret_cast<HMENU>(static_cast<INT_PTR>(showInMenuId)),
                                                    instance,
                                                    nullptr));
        PrefsInput::EnableMouseWheelForwarding(slot.showInMenuToggle);

        slot.showInMenuLabel.reset(CreateWindowExW(0, L"Static", L"", rightCaptionStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));
        slot.showInMenuDescription.reset(CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));
    }

    // Open prefs on assign toggle
    const std::wstring assignLabel = customButtons ? std::wstring{} : LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_OPEN_PREFS_ON_ASSIGN);
    state.hotPathOpenPrefsOnAssignToggle.reset(CreateWindowExW(0,
                                                               L"Button",
                                                               assignLabel.c_str(),
                                                               toggleStyle,
                                                               0,
                                                               0,
                                                               10,
                                                               10,
                                                               parent,
                                                               reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_HOT_PATHS_OPEN_PREFS_ON_ASSIGN)),
                                                               instance,
                                                               nullptr));
    PrefsInput::EnableMouseWheelForwarding(state.hotPathOpenPrefsOnAssignToggle);

    state.hotPathOpenPrefsOnAssignLabel.reset(CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));
    state.hotPathOpenPrefsOnAssignDescription.reset(CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));

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

    const auto setStaticVisuallyDisabled = [](const auto& hwndLike, bool disabled) noexcept
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
        if (! hwnd)
        {
            return;
        }

        EnableWindow(hwnd, TRUE);
        if (disabled)
        {
            SetPropW(hwnd, kPrefsVisuallyDisabledProp, reinterpret_cast<HANDLE>(1));
        }
        else
        {
            RemovePropW(hwnd, kPrefsVisuallyDisabledProp);
        }
        InvalidateRect(hwnd, nullptr, TRUE);
    };

    for (int i = 0; i < kSlotCount && i < static_cast<int>(state.hotPathSlotControls.size()); ++i)
    {
        const auto& slotCtl                = state.hotPathSlotControls[static_cast<size_t>(i)];
        const auto& slotData               = hp.slots[static_cast<size_t>(i)];
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

        setStaticVisuallyDisabled(slotCtl.labelLabel, enableDependentControls == FALSE);
        setEnabledAndInvalidate(slotCtl.labelFrame, enableDependentControls);
        setEnabledAndInvalidate(slotCtl.labelEdit, enableDependentControls);
        setStaticVisuallyDisabled(slotCtl.showInMenuLabel, enableDependentControls == FALSE);
        setEnabledAndInvalidate(slotCtl.showInMenuToggle, enableDependentControls);
        setEnabledAndInvalidate(slotCtl.showInMenuDescription, enableDependentControls);
    }

    PrefsUi::SetTwoStateToggleState(state.hotPathOpenPrefsOnAssignToggle, state.theme.systemHighContrast, hp.openPrefsOnAssign);
}

void HotPathsPane::LayoutControls(
    HWND host, PreferencesDialogState& state, int x, int& y, int width, [[maybe_unused]] int margin, int gapY, HFONT dialogFont) noexcept
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

    const int tPaddingX      = ThemedControls::ScaleDip(dpi, kTogglePaddingXDip);
    const int tGapX          = ThemedControls::ScaleDip(dpi, kToggleGapXDip);
    const int trackWidth     = ThemedControls::ScaleDip(dpi, kToggleTrackWidthDip);
    const int stateTextWidth = std::max(onWidth, offWidth);

    const int measuredToggleWidth = std::max(minToggleWidth, (2 * tPaddingX) + stateTextWidth + tGapX + trackWidth);
    const int toggleWidth         = std::min(std::max(0, width - 2 * cardPaddingX), measuredToggleWidth);

    auto pushCard = [&](const RECT& card) noexcept { state.pageSettingCards.push_back(card); };

    auto layoutToggleCard =
        [&](int cardX, int cardWidth, HWND label, std::wstring_view labelText, HWND toggle, HWND descLabel, std::wstring_view descText) noexcept
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
            SetWindowTextW(descLabel, descText.data());
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
        const wchar_t digitChar       = (i < 9) ? static_cast<wchar_t>(L'1' + i) : L'0';
        const std::wstring headerText = FormatStringResource(nullptr, IDS_PREFS_HOT_PATHS_SLOT_HEADER_FMT, digitChar);

        const int cardLeft     = x;
        const int cardWidth    = width;
        const int cardContentX = cardLeft + cardPaddingX;
        const int cardContentW = std::max(0, cardWidth - 2 * cardPaddingX);

        const int frameInset = ThemedControls::ScaleDip(dpi, kFramePaddingDip);

        const std::wstring pathLabelText  = LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_PATH_LABEL);
        const std::wstring labelLabelText = LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_LABEL_LABEL);
        const std::wstring showInMenuText = LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_SHOW_IN_MENU);

        // Slot header outside the card (matches the other "section header" style used in preferences).
        if (slotCtl.header)
        {
            SetWindowTextW(slotCtl.header.get(), headerText.c_str());
            SetWindowPos(slotCtl.header.get(), nullptr, cardLeft, y, cardWidth, headerHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(slotCtl.header.get(), WM_SETFONT, reinterpret_cast<WPARAM>(headerFont), TRUE);
            y += headerHeight + innerGap;
        }

        int slotY    = y;
        int contentY = slotY + cardPaddingY;

        // Layout: aligned caption + input + right actions.
        if (slotCtl.showInMenuDescription)
        {
            ShowWindow(slotCtl.showInMenuDescription.get(), SW_HIDE);
        }

        const int inlineGapX     = innerGap;
        const int minInlineEditW = ThemedControls::ScaleDip(dpi, kMinEditWidthDip);
        const int textPadX       = ThemedControls::ScaleDip(dpi, 6);

        const int pathCaptionW  = std::max(0, ThemedControls::MeasureTextWidth(host, dialogFont, pathLabelText) + textPadX);
        const int labelCaptionW = std::max(0, ThemedControls::MeasureTextWidth(host, dialogFont, labelLabelText) + textPadX);
        const int captionW      = std::max(pathCaptionW, labelCaptionW);

        const int showTextW = std::max(0, ThemedControls::MeasureTextWidth(host, dialogFont, showInMenuText) + textPadX);
        const int showAreaW = showTextW + inlineGapX + toggleWidth;
        const int rightEdge = cardContentX + cardContentW;
        const int browseX   = rightEdge - browseWidth;
        const int showAreaX = rightEdge - showAreaW;

        const int editX      = cardContentX + captionW + inlineGapX;
        const int pathEditW  = browseX - browseGap - editX;
        const int labelEditW = showAreaX - inlineGapX - editX;

        const bool useTableLayout =
            (cardContentW > 0 && captionW > 0 && pathEditW >= minInlineEditW && labelEditW >= minInlineEditW && showAreaW > toggleWidth);

        const std::wstring browseText = LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_BROWSE_ELLIPSIS);

        if (useTableLayout)
        {
            // Path row: [Path] [Edit] [Browse…]
            const int rowY = contentY;

            if (slotCtl.pathLabel)
            {
                SetWindowTextW(slotCtl.pathLabel.get(), pathLabelText.c_str());
                SetWindowPos(slotCtl.pathLabel.get(), nullptr, cardContentX, rowY, captionW, editHeight, SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(slotCtl.pathLabel.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            }

            if (slotCtl.pathFrame)
            {
                SetWindowPos(slotCtl.pathFrame.get(), nullptr, editX, rowY, std::max(10, pathEditW), editHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            }
            if (slotCtl.pathEdit)
            {
                SetWindowPos(slotCtl.pathEdit.get(),
                             nullptr,
                             editX + frameInset,
                             rowY + frameInset,
                             std::max(4, pathEditW - 2 * frameInset),
                             std::max(4, editHeight - 2 * frameInset),
                             SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(slotCtl.pathEdit.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            }

            if (slotCtl.browseButton)
            {
                SetWindowTextW(slotCtl.browseButton.get(), browseText.c_str());
                SetWindowPos(
                    slotCtl.browseButton.get(), nullptr, std::max(cardContentX, browseX), rowY, browseWidth, editHeight, SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(slotCtl.browseButton.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            }

            contentY += editHeight + innerGap;

            // Label row: [Label] [Edit] [Show in Change Drive menu] [On/Off toggle]
            const int rowY2 = contentY;

            if (slotCtl.labelLabel)
            {
                SetWindowTextW(slotCtl.labelLabel.get(), labelLabelText.c_str());
                SetWindowPos(slotCtl.labelLabel.get(), nullptr, cardContentX, rowY2, captionW, editHeight, SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(slotCtl.labelLabel.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            }

            if (slotCtl.labelFrame)
            {
                SetWindowPos(slotCtl.labelFrame.get(), nullptr, editX, rowY2, std::max(10, labelEditW), editHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            }
            if (slotCtl.labelEdit)
            {
                SetWindowPos(slotCtl.labelEdit.get(),
                             nullptr,
                             editX + frameInset,
                             rowY2 + frameInset,
                             std::max(4, labelEditW - 2 * frameInset),
                             std::max(4, editHeight - 2 * frameInset),
                             SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(slotCtl.labelEdit.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            }

            const int toggleY = rowY2 + (editHeight - rowHeight) / 2;

            if (slotCtl.showInMenuLabel)
            {
                SetWindowTextW(slotCtl.showInMenuLabel.get(), showInMenuText.c_str());
                const int showLabelW = std::max(0, showAreaW - inlineGapX - toggleWidth);
                SetWindowPos(slotCtl.showInMenuLabel.get(), nullptr, showAreaX, rowY2, showLabelW, editHeight, SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(slotCtl.showInMenuLabel.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            }

            if (slotCtl.showInMenuToggle)
            {
                SetWindowPos(slotCtl.showInMenuToggle.get(), nullptr, rightEdge - toggleWidth, toggleY, toggleWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(slotCtl.showInMenuToggle.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            }

            contentY += editHeight;
        }
        else
        {
            // Fallback: stack on narrow widths.
            if (slotCtl.pathLabel)
            {
                SetWindowTextW(slotCtl.pathLabel.get(), pathLabelText.c_str());
                SetWindowPos(slotCtl.pathLabel.get(), nullptr, cardContentX, contentY, cardContentW, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(slotCtl.pathLabel.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
                contentY += rowHeight;
            }

            const int stackPathEditW = std::max(0, cardContentW - browseGap - browseWidth);
            if (slotCtl.pathFrame)
            {
                SetWindowPos(slotCtl.pathFrame.get(), nullptr, cardContentX, contentY, std::max(10, stackPathEditW), editHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            }
            if (slotCtl.pathEdit)
            {
                SetWindowPos(slotCtl.pathEdit.get(),
                             nullptr,
                             cardContentX + frameInset,
                             contentY + frameInset,
                             std::max(4, stackPathEditW - 2 * frameInset),
                             std::max(4, editHeight - 2 * frameInset),
                             SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(slotCtl.pathEdit.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            }

            if (slotCtl.browseButton)
            {
                SetWindowTextW(slotCtl.browseButton.get(), browseText.c_str());
                SetWindowPos(slotCtl.browseButton.get(),
                             nullptr,
                             cardContentX + stackPathEditW + browseGap,
                             contentY,
                             browseWidth,
                             editHeight,
                             SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(slotCtl.browseButton.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            }
            contentY += editHeight + innerGap;

            if (slotCtl.labelLabel)
            {
                SetWindowTextW(slotCtl.labelLabel.get(), labelLabelText.c_str());
                SetWindowPos(slotCtl.labelLabel.get(), nullptr, cardContentX, contentY, cardContentW, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(slotCtl.labelLabel.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
                contentY += rowHeight;
            }

            if (slotCtl.labelFrame)
            {
                SetWindowPos(slotCtl.labelFrame.get(), nullptr, cardContentX, contentY, std::max(10, cardContentW), editHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            }
            if (slotCtl.labelEdit)
            {
                SetWindowPos(slotCtl.labelEdit.get(),
                             nullptr,
                             cardContentX + frameInset,
                             contentY + frameInset,
                             std::max(4, cardContentW - 2 * frameInset),
                             std::max(4, editHeight - 2 * frameInset),
                             SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(slotCtl.labelEdit.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            }
            contentY += editHeight + innerGap;

            if (slotCtl.showInMenuLabel)
            {
                SetWindowTextW(slotCtl.showInMenuLabel.get(), showInMenuText.c_str());
                SetWindowPos(slotCtl.showInMenuLabel.get(),
                             nullptr,
                             cardContentX,
                             contentY,
                             std::max(0, cardContentW - cardGapX - toggleWidth),
                             rowHeight,
                             SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(slotCtl.showInMenuLabel.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            }

            if (slotCtl.showInMenuToggle)
            {
                SetWindowPos(slotCtl.showInMenuToggle.get(),
                             nullptr,
                             cardContentX + std::max(0, cardContentW - toggleWidth),
                             contentY,
                             toggleWidth,
                             rowHeight,
                             SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(slotCtl.showInMenuToggle.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            }

            contentY += rowHeight;
        }

        const int cardBottom = std::max(slotY + 1, contentY + cardPaddingY);
        RECT card{};
        card.left   = cardLeft;
        card.top    = slotY;
        card.right  = cardLeft + cardWidth;
        card.bottom = cardBottom;
        pushCard(card);

        y = cardBottom + cardSpacingY;
    }

    // Open prefs on assign
    const std::wstring assignLabel = LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_OPEN_PREFS_ON_ASSIGN);
    const std::wstring assignDesc  = LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_OPEN_PREFS_ON_ASSIGN_DESC);
    layoutToggleCard(x,
                     width,
                     state.hotPathOpenPrefsOnAssignLabel.get(),
                     assignLabel,
                     state.hotPathOpenPrefsOnAssignToggle.get(),
                     state.hotPathOpenPrefsOnAssignDescription.get(),
                     assignDesc);
}

bool HotPathsPane::HandleCommand(HWND host, PreferencesDialogState& state, UINT commandId, UINT notifyCode, HWND hwndCtl) noexcept
{
    // Handle path and label edit changes.
    const bool isPathEdit =
        (commandId >= static_cast<UINT>(IDC_PREFS_HOT_PATHS_PATH_EDIT_BASE) && commandId < static_cast<UINT>(IDC_PREFS_HOT_PATHS_PATH_EDIT_BASE + kSlotCount));
    const bool isLabelEdit = (commandId >= static_cast<UINT>(IDC_PREFS_HOT_PATHS_LABEL_EDIT_BASE) &&
                              commandId < static_cast<UINT>(IDC_PREFS_HOT_PATHS_LABEL_EDIT_BASE + kSlotCount));

    if (isPathEdit || isLabelEdit)
    {
        if (notifyCode == EN_CHANGE || notifyCode == EN_KILLFOCUS)
        {
            HWND edit                       = hwndCtl ? hwndCtl : GetDlgItem(host, static_cast<int>(commandId));
            const std::wstring text         = PrefsUi::GetWindowTextString(edit);
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

                const auto setEnabledAndInvalidate = [](HWND hwnd, BOOL isEnabled) noexcept
                {
                    if (! hwnd)
                    {
                        return;
                    }
                    EnableWindow(hwnd, isEnabled);
                    InvalidateRect(hwnd, nullptr, TRUE);
                };

                const auto setStaticVisuallyDisabled = [](HWND hwnd, bool disabled) noexcept
                {
                    if (! hwnd)
                    {
                        return;
                    }

                    EnableWindow(hwnd, TRUE);
                    if (disabled)
                    {
                        SetPropW(hwnd, kPrefsVisuallyDisabledProp, reinterpret_cast<HANDLE>(1));
                    }
                    else
                    {
                        RemovePropW(hwnd, kPrefsVisuallyDisabledProp);
                    }
                    InvalidateRect(hwnd, nullptr, TRUE);
                };

                const bool visuallyDisabled = ! hasPathNow;
                setStaticVisuallyDisabled(slotCtl.labelLabel.get(), visuallyDisabled);
                setEnabledAndInvalidate(slotCtl.labelFrame.get(), enable);
                setEnabledAndInvalidate(slotCtl.labelEdit.get(), enable);
                setStaticVisuallyDisabled(slotCtl.showInMenuLabel.get(), visuallyDisabled);
                setEnabledAndInvalidate(slotCtl.showInMenuToggle.get(), enable);
                setEnabledAndInvalidate(slotCtl.showInMenuDescription.get(), enable);
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
    const bool isBrowse =
        (commandId >= static_cast<UINT>(IDC_PREFS_HOT_PATHS_BROWSE_BASE) && commandId < static_cast<UINT>(IDC_PREFS_HOT_PATHS_BROWSE_BASE + kSlotCount));
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

        if (! hwndCtl)
        {
            return true;
        }

        const bool ownerDraw = (GetWindowLongPtrW(hwndCtl, GWL_STYLE) & BS_TYPEMASK) == BS_OWNERDRAW;
        if (ownerDraw)
        {
            const bool toggledOn = PrefsUi::GetTwoStateToggleState(hwndCtl, false);
            PrefsUi::SetTwoStateToggleState(hwndCtl, false, ! toggledOn);
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
        if (! hwndCtl)
        {
            return true;
        }

        const bool ownerDraw = (GetWindowLongPtrW(hwndCtl, GWL_STYLE) & BS_TYPEMASK) == BS_OWNERDRAW;
        if (ownerDraw)
        {
            const bool toggledOn = PrefsUi::GetTwoStateToggleState(hwndCtl, false);
            PrefsUi::SetTwoStateToggleState(hwndCtl, false, ! toggledOn);
        }

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
