// Preferences.Panes.cpp

#include "Framework.h"

#include "Preferences.Panes.h"

#include <algorithm>
#include <string>

#include "Helpers.h"
#include "ThemedControls.h"

#include "resource.h"

bool PanesPane::EnsureCreated(HWND pageHost) noexcept
{
    return PrefsPaneHost::EnsureCreated(pageHost, _hWnd);
}

void PanesPane::ResizeToHostClient(HWND pageHost) noexcept
{
    PrefsPaneHost::ResizeToHostClient(pageHost, _hWnd.get());
}

void PanesPane::Show(bool visible) noexcept
{
    PrefsPaneHost::Show(_hWnd.get(), visible);
}

void PanesPane::LayoutControls(
    HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, int sectionY, HFONT dialogFont) noexcept
{
    using namespace PrefsLayoutConstants;

    static_cast<void>(margin);

    if (! host)
    {
        return;
    }

    const UINT dpi = GetDpiForWindow(host);

    const int rowHeight   = std::max(1, ThemedControls::ScaleDip(dpi, kRowHeightDip));
    const int labelHeight = std::max(1, ThemedControls::ScaleDip(dpi, kTitleHeightDip));
    const int gapX        = ThemedControls::ScaleDip(dpi, kToggleGapXDip);
    const int sectionX    = ThemedControls::ScaleDip(dpi, kCardPaddingXDip);

    const int headerHeight = std::max(1, ThemedControls::ScaleDip(dpi, kHeaderHeightDip));
    const HFONT headerFont = state.boldFont ? state.boldFont.get() : dialogFont;
    const HFONT infoFont   = state.italicFont ? state.italicFont.get() : dialogFont;

    const std::wstring leftHeaderText     = LoadStringResource(nullptr, IDS_PREFS_PANES_HEADER_LEFT);
    const std::wstring rightHeaderText    = LoadStringResource(nullptr, IDS_PREFS_PANES_HEADER_RIGHT);
    const std::wstring displayLabelText   = LoadStringResource(nullptr, IDS_PREFS_PANES_LABEL_DISPLAY);
    const std::wstring sortByLabelText    = LoadStringResource(nullptr, IDS_PREFS_PANES_LABEL_SORT_BY);
    const std::wstring directionLabelText = LoadStringResource(nullptr, IDS_PREFS_PANES_LABEL_DIRECTION);
    const std::wstring statusBarLabelText = LoadStringResource(nullptr, IDS_PREFS_PANES_LABEL_STATUS_BAR);
    const std::wstring statusBarDescText  = LoadStringResource(nullptr, IDS_PREFS_PANES_DESC_STATUS_BAR);
    const std::wstring historyLabelText   = LoadStringResource(nullptr, IDS_PREFS_PANES_LABEL_HISTORY_SIZE);
    const std::wstring historyDescText    = LoadStringResource(nullptr, IDS_PREFS_PANES_DESC_HISTORY_SIZE);
    const std::wstring briefText          = LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_BRIEF);
    const std::wstring detailedText       = LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_DETAILED);
    const std::wstring ascendingText      = LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_ASCENDING);
    const std::wstring descendingText     = LoadStringResource(nullptr, IDS_PREFS_PANES_OPTION_DESCENDING);

    if (state.theme.systemHighContrast)
    {
        auto placeFramedInput = [&](HWND frame, HWND control, int left, int top, int w, int h) noexcept
        {
            const int framePadding = (frame && ! state.theme.systemHighContrast) ? ThemedControls::ScaleDip(dpi, kFramePaddingDip) : 0;

            if (frame)
            {
                SetWindowPos(frame, nullptr, left, top, w, h, SWP_NOZORDER | SWP_NOACTIVATE);
            }
            if (control)
            {
                const int innerW = std::max(1, w - 2 * framePadding);
                const int innerH = std::max(1, h - 2 * framePadding);
                SetWindowPos(control, nullptr, left + framePadding, top + framePadding, innerW, innerH, SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(control, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            }
        };

        auto placeHeader = [&](HWND header, const std::wstring& text) noexcept
        {
            if (! header)
            {
                return;
            }
            SetWindowTextW(header, text.c_str());
            SetWindowPos(header, nullptr, x, y, width, headerHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(header, WM_SETFONT, reinterpret_cast<WPARAM>(headerFont), TRUE);
            y += headerHeight + gapY;
        };

        auto placeLabeledCombo = [&](HWND label, HWND frame, HWND combo, const std::wstring& labelText) noexcept
        {
            const int rowWidth   = std::max(0, width - sectionX);
            const int labelWidth = std::min(rowWidth, ThemedControls::ScaleDip(dpi, kMinComboWidthDip));
            const int available  = std::max(0, rowWidth - labelWidth - gapX);

            int desired          = combo ? ThemedControls::MeasureComboBoxPreferredWidth(combo, dpi) : 0;
            desired              = std::max(desired, ThemedControls::ScaleDip(dpi, kMinEditWidthDip + 10));
            const int comboWidth = std::min(available, desired);

            if (label)
            {
                SetWindowTextW(label, labelText.c_str());
                SetWindowPos(label, nullptr, x + sectionX, y + (rowHeight - labelHeight) / 2, labelWidth, labelHeight, SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            }
            placeFramedInput(frame, combo, x + sectionX + labelWidth + gapX, y, comboWidth, rowHeight);
            if (combo)
            {
                ThemedControls::EnsureComboBoxDroppedWidth(combo, dpi);
            }

            y += rowHeight + gapY;
        };

        auto placeStatusBarRow = [&](HWND label, HWND toggle, HWND descLabel, const std::wstring& labelText) noexcept
        {
            const int labelWidth  = std::min(width, ThemedControls::ScaleDip(dpi, kMinComboWidthDip));
            const int toggleWidth = std::min(width, ThemedControls::ScaleDip(dpi, kMediumComboWidthDip));

            if (label)
            {
                SetWindowTextW(label, labelText.c_str());
                SetWindowPos(label, nullptr, x + sectionX, y + (rowHeight - labelHeight) / 2, labelWidth, labelHeight, SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(label, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            }
            if (toggle)
            {
                SetWindowPos(toggle, nullptr, x + sectionX + labelWidth + gapX, y, toggleWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(toggle, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            }

            y += rowHeight + gapY;

            if (descLabel)
            {
                const int descHeight = PrefsUi::MeasureStaticTextHeight(host, infoFont, width - sectionX, statusBarDescText);
                SetWindowTextW(descLabel, statusBarDescText.c_str());
                SetWindowPos(descLabel, nullptr, x + sectionX, y, width - sectionX, std::max(0, descHeight), SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(descLabel, WM_SETFONT, reinterpret_cast<WPARAM>(infoFont), TRUE);
                y += std::max(0, descHeight) + sectionY;
            }
        };

        placeHeader(state.panesLeftHeader.get(), leftHeaderText);
        placeLabeledCombo(state.panesLeftDisplayLabel.get(), state.panesLeftDisplayFrame.get(), state.panesLeftDisplayCombo.get(), displayLabelText);
        placeLabeledCombo(state.panesLeftSortByLabel.get(), state.panesLeftSortByFrame.get(), state.panesLeftSortByCombo.get(), sortByLabelText);
        placeLabeledCombo(state.panesLeftSortDirLabel.get(), state.panesLeftSortDirFrame.get(), state.panesLeftSortDirCombo.get(), directionLabelText);
        placeStatusBarRow(
            state.panesLeftStatusBarLabel.get(), state.panesLeftStatusBarToggle.get(), state.panesLeftStatusBarDescription.get(), statusBarLabelText);

        placeHeader(state.panesRightHeader.get(), rightHeaderText);
        placeLabeledCombo(state.panesRightDisplayLabel.get(), state.panesRightDisplayFrame.get(), state.panesRightDisplayCombo.get(), displayLabelText);
        placeLabeledCombo(state.panesRightSortByLabel.get(), state.panesRightSortByFrame.get(), state.panesRightSortByCombo.get(), sortByLabelText);
        placeLabeledCombo(state.panesRightSortDirLabel.get(), state.panesRightSortDirFrame.get(), state.panesRightSortDirCombo.get(), directionLabelText);
        placeStatusBarRow(
            state.panesRightStatusBarLabel.get(), state.panesRightStatusBarToggle.get(), state.panesRightStatusBarDescription.get(), statusBarLabelText);

        const int rowWidth   = std::max(0, width - sectionX);
        const int labelWidth = std::min(rowWidth, ThemedControls::ScaleDip(dpi, kMinToggleWidthDip));
        const int editWidth  = std::min(std::max(0, rowWidth - labelWidth - gapX), ThemedControls::ScaleDip(dpi, 60));
        if (state.panesHistoryLabel)
        {
            SetWindowTextW(state.panesHistoryLabel.get(), historyLabelText.c_str());
            SetWindowPos(state.panesHistoryLabel.get(),
                         nullptr,
                         x + sectionX,
                         y + (rowHeight - labelHeight) / 2,
                         labelWidth,
                         labelHeight,
                         SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(state.panesHistoryLabel.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }
        if (state.panesHistoryEdit)
        {
            SetWindowPos(state.panesHistoryEdit.get(), nullptr, x + sectionX + labelWidth + gapX, y, editWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(state.panesHistoryEdit.get(), WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }
        y += rowHeight + gapY;

        if (state.panesHistoryDescription)
        {
            const int descHeight = PrefsUi::MeasureStaticTextHeight(host, infoFont, width - sectionX, historyDescText);
            SetWindowTextW(state.panesHistoryDescription.get(), historyDescText.c_str());
            SetWindowPos(
                state.panesHistoryDescription.get(), nullptr, x + sectionX, y, width - sectionX, std::max(0, descHeight), SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(state.panesHistoryDescription.get(), WM_SETFONT, reinterpret_cast<WPARAM>(infoFont), TRUE);
            y += std::max(0, descHeight) + gapY;
        }

        return;
    }

    const int titleHeight  = std::max(1, ThemedControls::ScaleDip(dpi, kTitleHeightDip));
    const int cardPaddingX = ThemedControls::ScaleDip(dpi, kCardPaddingXDip);
    const int cardPaddingY = ThemedControls::ScaleDip(dpi, kCardPaddingYDip);
    const int cardGapY     = ThemedControls::ScaleDip(dpi, kCardGapYDip);
    const int cardGapX     = ThemedControls::ScaleDip(dpi, kCardGapXDip);
    const int cardSpacingY = ThemedControls::ScaleDip(dpi, kCardSpacingYDip);

    auto pushCard = [&](const RECT& card) noexcept { state.pageSettingCards.push_back(card); };

    auto placeHeader = [&](HWND header, const std::wstring& text) noexcept
    {
        if (! header)
        {
            return;
        }
        SetWindowTextW(header, text.c_str());
        SetWindowPos(header, nullptr, x, y, width, headerHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(header, WM_SETFONT, reinterpret_cast<WPARAM>(headerFont), TRUE);
        y += headerHeight + gapY;
    };

    const std::wstring onLabel    = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);
    const std::wstring offLabel   = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);
    const HFONT toggleMeasureFont = state.boldFont ? state.boldFont.get() : dialogFont;
    const int onWidth             = ThemedControls::MeasureTextWidth(host, toggleMeasureFont, onLabel);
    const int offWidth            = ThemedControls::MeasureTextWidth(host, toggleMeasureFont, offLabel);
    const int briefWidth          = ThemedControls::MeasureTextWidth(host, toggleMeasureFont, briefText);
    const int detailedWidth       = ThemedControls::MeasureTextWidth(host, toggleMeasureFont, detailedText);
    const int ascendingWidth      = ThemedControls::MeasureTextWidth(host, toggleMeasureFont, ascendingText);
    const int descendingWidth     = ThemedControls::MeasureTextWidth(host, toggleMeasureFont, descendingText);
    const int paddingX            = ThemedControls::ScaleDip(dpi, kTogglePaddingXDip);
    const int stateGapX           = ThemedControls::ScaleDip(dpi, kToggleGapXDip);
    const int trackWidth          = ThemedControls::ScaleDip(dpi, kToggleTrackWidthDip);
    const int stateTextWidth      = (std::max)({onWidth, offWidth, briefWidth, detailedWidth, ascendingWidth, descendingWidth});
    const int measuredSwitchWidth = std::max(ThemedControls::ScaleDip(dpi, kMinToggleWidthDip), (2 * paddingX) + stateTextWidth + stateGapX + trackWidth);
    const int maxControlWidth     = std::max(0, width - 2 * cardPaddingX);
    const int switchWidth         = std::min(measuredSwitchWidth, maxControlWidth);

    auto layoutToggleCard = [&](HWND title, const std::wstring& titleText, HWND toggle, HWND descLabel, const std::wstring& descText) noexcept
    {
        const bool hasDesc = descLabel && ! descText.empty();
        const int descHeight =
            hasDesc ? PrefsUi::MeasureStaticTextHeight(host, infoFont, std::max(0, width - 2 * cardPaddingX - cardGapX - switchWidth), descText) : 0;
        const int contentHeight = hasDesc ? (titleHeight + cardGapY + descHeight) : titleHeight;
        const int cardHeight    = std::max(rowHeight + 2 * cardPaddingY, contentHeight + 2 * cardPaddingY);

        RECT card{};
        card.left   = x;
        card.top    = y;
        card.right  = x + width;
        card.bottom = y + cardHeight;
        pushCard(card);

        const int textWidth = std::max(0, width - 2 * cardPaddingX - cardGapX - switchWidth);
        const int titleY    = hasDesc ? (card.top + cardPaddingY) : (card.top + (cardHeight - titleHeight) / 2);

        if (title)
        {
            SetWindowTextW(title, titleText.c_str());
            SetWindowPos(title, nullptr, card.left + cardPaddingX, titleY, textWidth, titleHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(title, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }

        if (hasDesc)
        {
            SetWindowTextW(descLabel, descText.c_str());
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
                         card.right - cardPaddingX - switchWidth,
                         card.top + (cardHeight - rowHeight) / 2,
                         switchWidth,
                         rowHeight,
                         SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(toggle, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }

        y += cardHeight + cardSpacingY;
    };

    auto layoutFramedComboCard = [&](HWND title, const std::wstring& titleText, HWND frame, HWND combo) noexcept
    {
        int desiredWidth = combo ? ThemedControls::MeasureComboBoxPreferredWidth(combo, dpi) : 0;
        desiredWidth     = std::max(desiredWidth, ThemedControls::ScaleDip(dpi, kMinEditWidthDip + 10));
        desiredWidth     = std::min(desiredWidth, std::min(maxControlWidth, ThemedControls::ScaleDip(dpi, kMaxEditWidthDip)));

        const int textWidth  = std::max(0, width - 2 * cardPaddingX - cardGapX - desiredWidth);
        const int cardHeight = rowHeight + 2 * cardPaddingY;
        const int titleY     = cardPaddingY + (rowHeight - titleHeight) / 2;

        RECT card{};
        card.left   = x;
        card.top    = y;
        card.right  = x + width;
        card.bottom = y + cardHeight;
        pushCard(card);

        if (title)
        {
            SetWindowTextW(title, titleText.c_str());
            SetWindowPos(title, nullptr, card.left + cardPaddingX, card.top + titleY, textWidth, titleHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(title, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }

        const int inputX       = card.right - cardPaddingX - desiredWidth;
        const int inputY       = card.top + cardPaddingY;
        const int framePadding = (frame && ! state.theme.systemHighContrast) ? ThemedControls::ScaleDip(dpi, kFramePaddingDip) : 0;

        if (frame)
        {
            SetWindowPos(frame, nullptr, inputX, inputY, desiredWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        }
        if (combo)
        {
            SetWindowPos(combo,
                         nullptr,
                         inputX + framePadding,
                         inputY + framePadding,
                         std::max(1, desiredWidth - 2 * framePadding),
                         std::max(1, rowHeight - 2 * framePadding),
                         SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(combo, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            ThemedControls::EnsureComboBoxDroppedWidth(combo, dpi);
        }

        y += cardHeight + cardSpacingY;
    };

    auto layoutHistoryCard = [&](HWND title, const std::wstring& titleText, HWND frame, HWND edit, HWND descLabel, const std::wstring& descText) noexcept
    {
        const int desiredWidth = std::min(maxControlWidth, ThemedControls::ScaleDip(dpi, kMinComboWidthDip));
        const int textWidth    = std::max(0, width - 2 * cardPaddingX - cardGapX - desiredWidth);
        const int descHeight   = descLabel ? PrefsUi::MeasureStaticTextHeight(host, infoFont, textWidth, descText) : 0;

        const int contentHeight = std::max(0, titleHeight + cardGapY + descHeight);
        const int cardHeight    = std::max(rowHeight + 2 * cardPaddingY, contentHeight + 2 * cardPaddingY);

        RECT card{};
        card.left   = x;
        card.top    = y;
        card.right  = x + width;
        card.bottom = y + cardHeight;
        pushCard(card);

        if (title)
        {
            SetWindowTextW(title, titleText.c_str());
            SetWindowPos(title, nullptr, card.left + cardPaddingX, card.top + cardPaddingY, textWidth, titleHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(title, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }

        if (descLabel)
        {
            SetWindowTextW(descLabel, descText.c_str());
            SetWindowPos(descLabel,
                         nullptr,
                         card.left + cardPaddingX,
                         card.top + cardPaddingY + titleHeight + cardGapY,
                         textWidth,
                         std::max(0, descHeight),
                         SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(descLabel, WM_SETFONT, reinterpret_cast<WPARAM>(infoFont), TRUE);
        }

        const int inputX       = card.right - cardPaddingX - desiredWidth;
        const int inputY       = card.top + (cardHeight - rowHeight) / 2;
        const int framePadding = (frame && ! state.theme.systemHighContrast) ? ThemedControls::ScaleDip(dpi, kFramePaddingDip) : 0;

        if (frame)
        {
            SetWindowPos(frame, nullptr, inputX, inputY, desiredWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        }
        if (edit)
        {
            SetWindowPos(edit,
                         nullptr,
                         inputX + framePadding,
                         inputY + framePadding,
                         std::max(1, desiredWidth - 2 * framePadding),
                         std::max(1, rowHeight - 2 * framePadding),
                         SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(edit, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }

        y += cardHeight + cardSpacingY;
    };

    placeHeader(state.panesLeftHeader.get(), leftHeaderText);
    layoutToggleCard(state.panesLeftDisplayLabel.get(), displayLabelText, state.panesLeftDisplayToggle.get(), nullptr, {});
    layoutFramedComboCard(state.panesLeftSortByLabel.get(), sortByLabelText, state.panesLeftSortByFrame.get(), state.panesLeftSortByCombo.get());
    layoutToggleCard(state.panesLeftSortDirLabel.get(), directionLabelText, state.panesLeftSortDirToggle.get(), nullptr, {});
    layoutToggleCard(state.panesLeftStatusBarLabel.get(),
                     statusBarLabelText,
                     state.panesLeftStatusBarToggle.get(),
                     state.panesLeftStatusBarDescription.get(),
                     statusBarDescText);

    y += std::max(0, sectionY - cardSpacingY);

    placeHeader(state.panesRightHeader.get(), rightHeaderText);
    layoutToggleCard(state.panesRightDisplayLabel.get(), displayLabelText, state.panesRightDisplayToggle.get(), nullptr, {});
    layoutFramedComboCard(state.panesRightSortByLabel.get(), sortByLabelText, state.panesRightSortByFrame.get(), state.panesRightSortByCombo.get());
    layoutToggleCard(state.panesRightSortDirLabel.get(), directionLabelText, state.panesRightSortDirToggle.get(), nullptr, {});
    layoutToggleCard(state.panesRightStatusBarLabel.get(),
                     statusBarLabelText,
                     state.panesRightStatusBarToggle.get(),
                     state.panesRightStatusBarDescription.get(),
                     statusBarDescText);

    y += std::max(0, sectionY - cardSpacingY);

    layoutHistoryCard(state.panesHistoryLabel.get(),
                      historyLabelText,
                      state.panesHistoryFrame.get(),
                      state.panesHistoryEdit.get(),
                      state.panesHistoryDescription.get(),
                      historyDescText);
}

void PanesPane::Refresh(HWND /*host*/, PreferencesDialogState& state) noexcept
{
    const PrefsFolders::FolderPanePreferences left  = PrefsFolders::GetFolderPanePreferences(state.workingSettings, PrefsFolders::kLeftPaneSlot);
    const PrefsFolders::FolderPanePreferences right = PrefsFolders::GetFolderPanePreferences(state.workingSettings, PrefsFolders::kRightPaneSlot);
    const uint32_t historyMax                       = PrefsFolders::GetFolderHistoryMax(state.workingSettings);

    state.refreshingPanesPage = true;
    const auto reset          = wil::scope_exit([&] { state.refreshingPanesPage = false; });

    PrefsUi::SelectComboItemByData(state.panesLeftDisplayCombo.get(), static_cast<LPARAM>(left.display));
    PrefsUi::SetTwoStateToggleState(
        state.panesLeftDisplayToggle.get(), state.theme.systemHighContrast, left.display == Common::Settings::FolderDisplayMode::Brief);
    PrefsUi::SelectComboItemByData(state.panesLeftSortByCombo.get(), static_cast<LPARAM>(left.sortBy));
    PrefsUi::SelectComboItemByData(state.panesLeftSortDirCombo.get(), static_cast<LPARAM>(left.sortDirection));
    PrefsUi::SetTwoStateToggleState(
        state.panesLeftSortDirToggle.get(), state.theme.systemHighContrast, left.sortDirection == Common::Settings::FolderSortDirection::Ascending);
    PrefsUi::SetTwoStateToggleState(state.panesLeftStatusBarToggle.get(), state.theme.systemHighContrast, left.statusBarVisible);

    PrefsUi::SelectComboItemByData(state.panesRightDisplayCombo.get(), static_cast<LPARAM>(right.display));
    PrefsUi::SetTwoStateToggleState(
        state.panesRightDisplayToggle.get(), state.theme.systemHighContrast, right.display == Common::Settings::FolderDisplayMode::Brief);
    PrefsUi::SelectComboItemByData(state.panesRightSortByCombo.get(), static_cast<LPARAM>(right.sortBy));
    PrefsUi::SelectComboItemByData(state.panesRightSortDirCombo.get(), static_cast<LPARAM>(right.sortDirection));
    PrefsUi::SetTwoStateToggleState(
        state.panesRightSortDirToggle.get(), state.theme.systemHighContrast, right.sortDirection == Common::Settings::FolderSortDirection::Ascending);
    PrefsUi::SetTwoStateToggleState(state.panesRightStatusBarToggle.get(), state.theme.systemHighContrast, right.statusBarVisible);

    if (state.panesHistoryEdit)
    {
        std::wstring text;
        text = std::to_wstring(historyMax);
        SetWindowTextW(state.panesHistoryEdit.get(), text.c_str());
    }
}

bool PanesPane::HandleCommand(HWND host, PreferencesDialogState& state, UINT commandId, UINT notifyCode, HWND hwndCtl) noexcept
{
    if (! host)
    {
        return false;
    }

    HWND dlg = GetParent(host);
    if (! dlg)
    {
        return false;
    }

    if (state.refreshingPanesPage)
    {
        return false;
    }

    auto handleDisplayCombo = [&](const std::wstring_view slot, HWND combo) noexcept
    {
        const auto dataOpt = PrefsUi::TryGetSelectedComboItemData(combo);
        if (! dataOpt.has_value())
        {
            return false;
        }

        auto* pane = PrefsFolders::EnsureWorkingFolderPane(state.workingSettings, slot);
        if (! pane)
        {
            return true;
        }

        pane->view.display = static_cast<Common::Settings::FolderDisplayMode>(dataOpt.value());
        SetDirty(dlg, state);
        return true;
    };

    auto handleSortDirCombo = [&](const std::wstring_view slot, HWND combo) noexcept
    {
        const auto dataOpt = PrefsUi::TryGetSelectedComboItemData(combo);
        if (! dataOpt.has_value())
        {
            return false;
        }

        auto* pane = PrefsFolders::EnsureWorkingFolderPane(state.workingSettings, slot);
        if (! pane)
        {
            return true;
        }

        pane->view.sortDirection = static_cast<Common::Settings::FolderSortDirection>(dataOpt.value());
        SetDirty(dlg, state);
        return true;
    };

    auto handleSortByCombo = [&](const std::wstring_view slot, HWND combo) noexcept
    {
        const auto dataOpt = PrefsUi::TryGetSelectedComboItemData(combo);
        if (! dataOpt.has_value())
        {
            return false;
        }

        const auto sortBy = static_cast<Common::Settings::FolderSortBy>(dataOpt.value());
        auto* pane        = PrefsFolders::EnsureWorkingFolderPane(state.workingSettings, slot);
        if (! pane)
        {
            return true;
        }

        pane->view.sortBy        = sortBy;
        pane->view.sortDirection = PrefsFolders::DefaultFolderSortDirection(sortBy);
        SetDirty(dlg, state);
        Refresh(host, state);
        return true;
    };

    auto handleTwoStateToggle = [&](const std::wstring_view slot, HWND clicked, const bool manualFlip, auto&& apply) noexcept
    {
        if (! clicked)
        {
            return true;
        }

        const bool ownerDraw = (GetWindowLongPtrW(clicked, GWL_STYLE) & BS_TYPEMASK) == BS_OWNERDRAW;
        if (manualFlip && ownerDraw)
        {
            const LONG_PTR current = GetWindowLongPtrW(clicked, GWLP_USERDATA);
            SetWindowLongPtrW(clicked, GWLP_USERDATA, current == 0 ? 1 : 0);
            InvalidateRect(clicked, nullptr, TRUE);
        }

        const bool toggledOn = PrefsUi::GetTwoStateToggleState(clicked, state.theme.systemHighContrast);
        auto* pane           = PrefsFolders::EnsureWorkingFolderPane(state.workingSettings, slot);
        if (! pane)
        {
            return true;
        }

        apply(*pane, toggledOn);
        SetDirty(dlg, state);
        return true;
    };

    switch (commandId)
    {
        case IDC_PREFS_PANES_LEFT_DISPLAY_COMBO:
            if (notifyCode == CBN_SELCHANGE)
            {
                return handleDisplayCombo(PrefsFolders::kLeftPaneSlot, state.panesLeftDisplayCombo.get());
            }
            break;
        case IDC_PREFS_PANES_LEFT_SORTBY_COMBO:
            if (notifyCode == CBN_SELCHANGE)
            {
                return handleSortByCombo(PrefsFolders::kLeftPaneSlot, state.panesLeftSortByCombo.get());
            }
            break;
        case IDC_PREFS_PANES_LEFT_SORTDIR_COMBO:
            if (notifyCode == CBN_SELCHANGE)
            {
                return handleSortDirCombo(PrefsFolders::kLeftPaneSlot, state.panesLeftSortDirCombo.get());
            }
            break;
        case IDC_PREFS_PANES_RIGHT_DISPLAY_COMBO:
            if (notifyCode == CBN_SELCHANGE)
            {
                return handleDisplayCombo(PrefsFolders::kRightPaneSlot, state.panesRightDisplayCombo.get());
            }
            break;
        case IDC_PREFS_PANES_RIGHT_SORTBY_COMBO:
            if (notifyCode == CBN_SELCHANGE)
            {
                return handleSortByCombo(PrefsFolders::kRightPaneSlot, state.panesRightSortByCombo.get());
            }
            break;
        case IDC_PREFS_PANES_RIGHT_SORTDIR_COMBO:
            if (notifyCode == CBN_SELCHANGE)
            {
                return handleSortDirCombo(PrefsFolders::kRightPaneSlot, state.panesRightSortDirCombo.get());
            }
            break;
        case IDC_PREFS_PANES_HISTORY_MAX_EDIT:
            if (notifyCode == EN_CHANGE)
            {
                const std::wstring text = PrefsUi::GetWindowTextString(state.panesHistoryEdit.get());
                const auto valueOpt     = PrefsUi::TryParseUInt32(text);
                if (! valueOpt.has_value())
                {
                    return true;
                }

                const uint32_t value = valueOpt.value();
                if (value < 1u || value > 50u)
                {
                    return true;
                }

                auto* folders = PrefsFolders::EnsureWorkingFoldersSettings(state.workingSettings);
                if (! folders)
                {
                    return true;
                }

                folders->historyMax = value;
                SetDirty(dlg, state);
                return true;
            }
            if (notifyCode == EN_KILLFOCUS)
            {
                const std::wstring text = PrefsUi::GetWindowTextString(state.panesHistoryEdit.get());
                const auto valueOpt     = PrefsUi::TryParseUInt32(text);
                if (valueOpt.has_value())
                {
                    const uint32_t value = std::clamp(valueOpt.value(), 1u, 50u);
                    auto* folders        = PrefsFolders::EnsureWorkingFoldersSettings(state.workingSettings);
                    if (folders)
                    {
                        folders->historyMax = value;
                        SetDirty(dlg, state);
                    }
                }

                Refresh(host, state);
                return true;
            }
            break;
        case IDC_PREFS_PANES_LEFT_STATUSBAR_TOGGLE:
        case IDC_PREFS_PANES_RIGHT_STATUSBAR_TOGGLE:
            if (notifyCode == BN_CLICKED)
            {
                const std::wstring_view slot = commandId == IDC_PREFS_PANES_LEFT_STATUSBAR_TOGGLE ? PrefsFolders::kLeftPaneSlot : PrefsFolders::kRightPaneSlot;
                return handleTwoStateToggle(slot, hwndCtl, true, [](Common::Settings::FolderPane& pane, bool on) noexcept { pane.view.statusBarVisible = on; });
            }
            break;
        case IDC_PREFS_PANES_LEFT_DISPLAY_TOGGLE:
        case IDC_PREFS_PANES_RIGHT_DISPLAY_TOGGLE:
        case IDC_PREFS_PANES_LEFT_SORTDIR_TOGGLE:
        case IDC_PREFS_PANES_RIGHT_SORTDIR_TOGGLE:
            if (notifyCode == BN_CLICKED)
            {
                const bool isLeft            = commandId == IDC_PREFS_PANES_LEFT_DISPLAY_TOGGLE || commandId == IDC_PREFS_PANES_LEFT_SORTDIR_TOGGLE;
                const std::wstring_view slot = isLeft ? PrefsFolders::kLeftPaneSlot : PrefsFolders::kRightPaneSlot;

                bool changed = handleTwoStateToggle(
                    slot,
                    hwndCtl,
                    true,
                    [&](Common::Settings::FolderPane& pane, bool on) noexcept
                    {
                        if (commandId == IDC_PREFS_PANES_LEFT_DISPLAY_TOGGLE || commandId == IDC_PREFS_PANES_RIGHT_DISPLAY_TOGGLE)
                        {
                            pane.view.display = on ? Common::Settings::FolderDisplayMode::Brief : Common::Settings::FolderDisplayMode::Detailed;
                        }
                        else
                        {
                            pane.view.sortDirection = on ? Common::Settings::FolderSortDirection::Ascending : Common::Settings::FolderSortDirection::Descending;
                        }
                    });

                if (changed)
                {
                    Refresh(host, state);
                }

                return true;
            }
            break;
    }

    return false;
}

void PanesPane::CreateControls(HWND parent, PreferencesDialogState& state) noexcept
{
    if (! parent)
    {
        return;
    }

    const DWORD baseStaticStyle = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX;
    const DWORD wrapStaticStyle = WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX | SS_EDITCONTROL;

    const bool customButtons = ! state.theme.systemHighContrast;

    // Panes page (Phase 5).
    state.panesLeftHeader.reset(CreateWindowExW(0,
                                                L"Static",
                                                LoadStringResource(nullptr, IDS_PREFS_PANES_HEADER_LEFT).c_str(),
                                                baseStaticStyle,
                                                0,
                                                0,
                                                10,
                                                10,
                                                parent,
                                                nullptr,
                                                GetModuleHandleW(nullptr),
                                                nullptr));
    state.panesLeftDisplayLabel.reset(CreateWindowExW(0,
                                                      L"Static",
                                                      LoadStringResource(nullptr, IDS_PREFS_PANES_LABEL_DISPLAY).c_str(),
                                                      baseStaticStyle,
                                                      0,
                                                      0,
                                                      10,
                                                      10,
                                                      parent,
                                                      nullptr,
                                                      GetModuleHandleW(nullptr),
                                                      nullptr));
    PrefsInput::CreateFramedComboBox(state, parent, state.panesLeftDisplayFrame, state.panesLeftDisplayCombo, IDC_PREFS_PANES_LEFT_DISPLAY_COMBO);
    if (customButtons)
    {
        state.panesLeftDisplayToggle.reset(CreateWindowExW(0,
                                                           L"Button",
                                                           L"",
                                                           WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
                                                           0,
                                                           0,
                                                           10,
                                                           10,
                                                           parent,
                                                           reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_PANES_LEFT_DISPLAY_TOGGLE)),
                                                           GetModuleHandleW(nullptr),
                                                           nullptr));
    }
    state.panesLeftSortByLabel.reset(CreateWindowExW(0,
                                                     L"Static",
                                                     LoadStringResource(nullptr, IDS_PREFS_PANES_LABEL_SORT_BY).c_str(),
                                                     baseStaticStyle,
                                                     0,
                                                     0,
                                                     10,
                                                     10,
                                                     parent,
                                                     nullptr,
                                                     GetModuleHandleW(nullptr),
                                                     nullptr));
    PrefsInput::CreateFramedComboBox(state, parent, state.panesLeftSortByFrame, state.panesLeftSortByCombo, IDC_PREFS_PANES_LEFT_SORTBY_COMBO);
    state.panesLeftSortDirLabel.reset(CreateWindowExW(0,
                                                      L"Static",
                                                      LoadStringResource(nullptr, IDS_PREFS_PANES_LABEL_DIRECTION).c_str(),
                                                      baseStaticStyle,
                                                      0,
                                                      0,
                                                      10,
                                                      10,
                                                      parent,
                                                      nullptr,
                                                      GetModuleHandleW(nullptr),
                                                      nullptr));
    PrefsInput::CreateFramedComboBox(state, parent, state.panesLeftSortDirFrame, state.panesLeftSortDirCombo, IDC_PREFS_PANES_LEFT_SORTDIR_COMBO);
    if (customButtons)
    {
        state.panesLeftSortDirToggle.reset(CreateWindowExW(0,
                                                           L"Button",
                                                           L"",
                                                           WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
                                                           0,
                                                           0,
                                                           10,
                                                           10,
                                                           parent,
                                                           reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_PANES_LEFT_SORTDIR_TOGGLE)),
                                                           GetModuleHandleW(nullptr),
                                                           nullptr));
    }
    state.panesLeftStatusBarLabel.reset(CreateWindowExW(0,
                                                        L"Static",
                                                        LoadStringResource(nullptr, IDS_PREFS_PANES_LABEL_STATUS_BAR).c_str(),
                                                        baseStaticStyle,
                                                        0,
                                                        0,
                                                        10,
                                                        10,
                                                        parent,
                                                        nullptr,
                                                        GetModuleHandleW(nullptr),
                                                        nullptr));
    state.panesLeftStatusBarDescription.reset(CreateWindowExW(0,
                                                              L"Static",
                                                              LoadStringResource(nullptr, IDS_PREFS_PANES_DESC_STATUS_BAR).c_str(),
                                                              wrapStaticStyle,
                                                              0,
                                                              0,
                                                              10,
                                                              10,
                                                              parent,
                                                              nullptr,
                                                              GetModuleHandleW(nullptr),
                                                              nullptr));
    if (customButtons)
    {
        state.panesLeftStatusBarToggle.reset(CreateWindowExW(0,
                                                             L"Button",
                                                             L"",
                                                             WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
                                                             0,
                                                             0,
                                                             10,
                                                             10,
                                                             parent,
                                                             reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_PANES_LEFT_STATUSBAR_TOGGLE)),
                                                             GetModuleHandleW(nullptr),
                                                             nullptr));
    }
    else
    {
        state.panesLeftStatusBarToggle.reset(CreateWindowExW(0,
                                                             L"Button",
                                                             LoadStringResource(nullptr, IDS_PREFS_PANES_CHECK_SHOW_STATUS_BAR).c_str(),
                                                             WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
                                                             0,
                                                             0,
                                                             10,
                                                             10,
                                                             parent,
                                                             reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_PANES_LEFT_STATUSBAR_TOGGLE)),
                                                             GetModuleHandleW(nullptr),
                                                             nullptr));
    }

    state.panesRightHeader.reset(CreateWindowExW(0,
                                                 L"Static",
                                                 LoadStringResource(nullptr, IDS_PREFS_PANES_HEADER_RIGHT).c_str(),
                                                 baseStaticStyle,
                                                 0,
                                                 0,
                                                 10,
                                                 10,
                                                 parent,
                                                 nullptr,
                                                 GetModuleHandleW(nullptr),
                                                 nullptr));
    state.panesRightDisplayLabel.reset(CreateWindowExW(0,
                                                       L"Static",
                                                       LoadStringResource(nullptr, IDS_PREFS_PANES_LABEL_DISPLAY).c_str(),
                                                       baseStaticStyle,
                                                       0,
                                                       0,
                                                       10,
                                                       10,
                                                       parent,
                                                       nullptr,
                                                       GetModuleHandleW(nullptr),
                                                       nullptr));
    PrefsInput::CreateFramedComboBox(state, parent, state.panesRightDisplayFrame, state.panesRightDisplayCombo, IDC_PREFS_PANES_RIGHT_DISPLAY_COMBO);
    if (customButtons)
    {
        state.panesRightDisplayToggle.reset(CreateWindowExW(0,
                                                            L"Button",
                                                            L"",
                                                            WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
                                                            0,
                                                            0,
                                                            10,
                                                            10,
                                                            parent,
                                                            reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_PANES_RIGHT_DISPLAY_TOGGLE)),
                                                            GetModuleHandleW(nullptr),
                                                            nullptr));
    }

    state.panesRightSortByLabel.reset(CreateWindowExW(0,
                                                      L"Static",
                                                      LoadStringResource(nullptr, IDS_PREFS_PANES_LABEL_SORT_BY).c_str(),
                                                      baseStaticStyle,
                                                      0,
                                                      0,
                                                      10,
                                                      10,
                                                      parent,
                                                      nullptr,
                                                      GetModuleHandleW(nullptr),
                                                      nullptr));
    PrefsInput::CreateFramedComboBox(state, parent, state.panesRightSortByFrame, state.panesRightSortByCombo, IDC_PREFS_PANES_RIGHT_SORTBY_COMBO);
    state.panesRightSortDirLabel.reset(CreateWindowExW(0,
                                                       L"Static",
                                                       LoadStringResource(nullptr, IDS_PREFS_PANES_LABEL_DIRECTION).c_str(),
                                                       baseStaticStyle,
                                                       0,
                                                       0,
                                                       10,
                                                       10,
                                                       parent,
                                                       nullptr,
                                                       GetModuleHandleW(nullptr),
                                                       nullptr));
    PrefsInput::CreateFramedComboBox(state, parent, state.panesRightSortDirFrame, state.panesRightSortDirCombo, IDC_PREFS_PANES_RIGHT_SORTDIR_COMBO);
    if (customButtons)
    {
        state.panesRightSortDirToggle.reset(CreateWindowExW(0,
                                                            L"Button",
                                                            L"",
                                                            WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
                                                            0,
                                                            0,
                                                            10,
                                                            10,
                                                            parent,
                                                            reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_PANES_RIGHT_SORTDIR_TOGGLE)),
                                                            GetModuleHandleW(nullptr),
                                                            nullptr));
    }
    state.panesRightStatusBarLabel.reset(CreateWindowExW(0,
                                                         L"Static",
                                                         LoadStringResource(nullptr, IDS_PREFS_PANES_LABEL_STATUS_BAR).c_str(),
                                                         baseStaticStyle,
                                                         0,
                                                         0,
                                                         10,
                                                         10,
                                                         parent,
                                                         nullptr,
                                                         GetModuleHandleW(nullptr),
                                                         nullptr));
    state.panesRightStatusBarDescription.reset(CreateWindowExW(0,
                                                               L"Static",
                                                               LoadStringResource(nullptr, IDS_PREFS_PANES_DESC_STATUS_BAR).c_str(),
                                                               wrapStaticStyle,
                                                               0,
                                                               0,
                                                               10,
                                                               10,
                                                               parent,
                                                               nullptr,
                                                               GetModuleHandleW(nullptr),
                                                               nullptr));
    if (customButtons)
    {
        state.panesRightStatusBarToggle.reset(CreateWindowExW(0,
                                                              L"Button",
                                                              L"",
                                                              WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
                                                              0,
                                                              0,
                                                              10,
                                                              10,
                                                              parent,
                                                              reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_PANES_RIGHT_STATUSBAR_TOGGLE)),
                                                              GetModuleHandleW(nullptr),
                                                              nullptr));
    }
    else
    {
        state.panesRightStatusBarToggle.reset(CreateWindowExW(0,
                                                              L"Button",
                                                              LoadStringResource(nullptr, IDS_PREFS_PANES_CHECK_SHOW_STATUS_BAR).c_str(),
                                                              WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX,
                                                              0,
                                                              0,
                                                              10,
                                                              10,
                                                              parent,
                                                              reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PREFS_PANES_RIGHT_STATUSBAR_TOGGLE)),
                                                              GetModuleHandleW(nullptr),
                                                              nullptr));
    }

    state.panesHistoryLabel.reset(CreateWindowExW(0,
                                                  L"Static",
                                                  LoadStringResource(nullptr, IDS_PREFS_PANES_LABEL_HISTORY_SIZE).c_str(),
                                                  baseStaticStyle,
                                                  0,
                                                  0,
                                                  10,
                                                  10,
                                                  parent,
                                                  nullptr,
                                                  GetModuleHandleW(nullptr),
                                                  nullptr));
    PrefsInput::CreateFramedEditBox(state,
                                    parent,
                                    state.panesHistoryFrame,
                                    state.panesHistoryEdit,
                                    IDC_PREFS_PANES_HISTORY_MAX_EDIT,
                                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_NUMBER | ES_AUTOHSCROLL);
    if (state.panesHistoryEdit)
    {
        SendMessageW(state.panesHistoryEdit.get(), EM_SETLIMITTEXT, 2, 0);
    }

    state.panesHistoryDescription.reset(CreateWindowExW(0,
                                                        L"Static",
                                                        LoadStringResource(nullptr, IDS_PREFS_PANES_DESC_HISTORY_SIZE).c_str(),
                                                        wrapStaticStyle,
                                                        0,
                                                        0,
                                                        10,
                                                        10,
                                                        parent,
                                                        nullptr,
                                                        GetModuleHandleW(nullptr),
                                                        nullptr));

    PrefsInput::EnableMouseWheelForwarding(state.panesLeftDisplayToggle.get());
    PrefsInput::EnableMouseWheelForwarding(state.panesLeftSortDirToggle.get());
    PrefsInput::EnableMouseWheelForwarding(state.panesLeftStatusBarToggle.get());
    PrefsInput::EnableMouseWheelForwarding(state.panesRightDisplayToggle.get());
    PrefsInput::EnableMouseWheelForwarding(state.panesRightSortDirToggle.get());
    PrefsInput::EnableMouseWheelForwarding(state.panesRightStatusBarToggle.get());
}
