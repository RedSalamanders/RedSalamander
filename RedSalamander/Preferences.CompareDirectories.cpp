// Preferences.CompareDirectories.cpp

#include "Framework.h"

#include "Preferences.CompareDirectories.h"

#include <algorithm>
#include <string>

#include "Helpers.h"
#include "ThemedControls.h"
#include "resource.h"

// Local convenience aliases for frequently-used shared utilities
namespace
{
using PrefsCompareDirectories::EnsureWorkingCompareDirectoriesSettings;
using PrefsCompareDirectories::GetCompareDirectoriesSettingsOrDefault;
using PrefsCompareDirectories::MaybeResetWorkingCompareDirectoriesSettingsIfEmpty;
} // namespace

bool CompareDirectoriesPane::EnsureCreated(HWND pageHost) noexcept
{
    return PrefsPaneHost::EnsureCreated(pageHost, _hWnd);
}

void CompareDirectoriesPane::ResizeToHostClient(HWND pageHost) noexcept
{
    PrefsPaneHost::ResizeToHostClient(pageHost, _hWnd.get());
}

void CompareDirectoriesPane::Show(bool visible) noexcept
{
    PrefsPaneHost::Show(_hWnd.get(), visible);
}

void CompareDirectoriesPane::CreateControls(HWND parent, PreferencesDialogState& state) noexcept
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

    const auto createToggle = [&](wil::unique_hwnd& out, UINT id, UINT labelResourceId) noexcept
    {
        const std::wstring label = customButtons ? std::wstring{} : LoadStringResource(nullptr, labelResourceId);
        out.reset(CreateWindowExW(
            0, L"Button", label.c_str(), toggleStyle, 0, 0, 10, 10, parent, reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)), instance, nullptr));
        PrefsInput::EnableMouseWheelForwarding(out);
    };

    createToggle(state.advancedCompareSizeToggle, IDC_PREFS_ADV_COMPARE_SIZE_TOGGLE, IDS_COMPARE_OPTIONS_SIZE_TITLE);
    createToggle(state.advancedCompareDateTimeToggle, IDC_PREFS_ADV_COMPARE_DATETIME_TOGGLE, IDS_COMPARE_OPTIONS_DATETIME_TITLE);
    createToggle(state.advancedCompareAttributesToggle, IDC_PREFS_ADV_COMPARE_ATTRIBUTES_TOGGLE, IDS_COMPARE_OPTIONS_ATTRIBUTES_TITLE);
    createToggle(state.advancedCompareContentToggle, IDC_PREFS_ADV_COMPARE_CONTENT_TOGGLE, IDS_COMPARE_OPTIONS_CONTENT_TITLE);

    createToggle(state.advancedCompareSubdirectoriesToggle, IDC_PREFS_ADV_COMPARE_SUBDIRS_TOGGLE, IDS_COMPARE_OPTIONS_SUBDIRS_TITLE);
    createToggle(
        state.advancedCompareSubdirectoryAttributesToggle, IDC_PREFS_ADV_COMPARE_SUBDIR_ATTRIBUTES_TOGGLE, IDS_COMPARE_OPTIONS_SUBDIR_ATTRIBUTES_TITLE);
    createToggle(
        state.advancedCompareSelectSubdirsOnlyInOnePaneToggle, IDC_PREFS_ADV_COMPARE_SELECT_SUBDIRS_ONE_PANE_TOGGLE, IDS_COMPARE_OPTIONS_SELECT_SUBDIRS_TITLE);

    createToggle(state.advancedCompareShowIdenticalToggle, IDC_PREFS_ADV_COMPARE_SHOW_IDENTICAL_TOGGLE, IDS_PREFS_COMPARE_SHOW_IDENTICAL_TITLE);

    createToggle(state.advancedCompareIgnoreFilesToggle, IDC_PREFS_ADV_COMPARE_IGNORE_FILES_TOGGLE, IDS_COMPARE_OPTIONS_IGNORE_FILES_TITLE);
    createToggle(state.advancedCompareIgnoreDirectoriesToggle, IDC_PREFS_ADV_COMPARE_IGNORE_DIRECTORIES_TOGGLE, IDS_COMPARE_OPTIONS_IGNORE_DIRECTORIES_TITLE);

    state.advancedCompareDirectoriesHeader.reset(CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));
    state.advancedCompareSectionSubdirsHeader.reset(CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));
    state.advancedCompareSectionCompareHeader.reset(CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));
    state.advancedCompareSectionAdditionalHeader.reset(CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));
    state.advancedCompareSectionMoreHeader.reset(CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));

    state.advancedCompareSizeLabel.reset(CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));
    state.advancedCompareSizeDescription.reset(CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));

    state.advancedCompareDateTimeLabel.reset(CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));
    state.advancedCompareDateTimeDescription.reset(CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));

    state.advancedCompareAttributesLabel.reset(CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));
    state.advancedCompareAttributesDescription.reset(CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));

    state.advancedCompareContentLabel.reset(CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));
    state.advancedCompareContentDescription.reset(CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));

    state.advancedCompareContentWorkersLabel.reset(CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));
    PrefsInput::CreateFramedComboBox(state,
                                    parent,
                                    state.advancedCompareContentWorkersFrame,
                                    state.advancedCompareContentWorkersCombo,
                                    IDC_PREFS_ADV_COMPARE_CONTENT_WORKERS_COMBO);
    state.advancedCompareContentWorkersDescription.reset(
        CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));

    state.advancedCompareSubdirectoriesLabel.reset(CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));
    state.advancedCompareSubdirectoriesDescription.reset(CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));

    state.advancedCompareSubdirectoryAttributesLabel.reset(
        CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));
    state.advancedCompareSubdirectoryAttributesDescription.reset(
        CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));

    state.advancedCompareSelectSubdirsOnlyInOnePaneLabel.reset(
        CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));
    state.advancedCompareSelectSubdirsOnlyInOnePaneDescription.reset(
        CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));

    state.advancedCompareShowIdenticalLabel.reset(CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));
    state.advancedCompareShowIdenticalDescription.reset(CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));

    state.advancedCompareIgnoreFilesLabel.reset(CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));
    state.advancedCompareIgnoreFilesDescription.reset(CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));
    state.advancedCompareIgnoreFilesPatternsLabel.reset(CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));
    PrefsInput::CreateFramedEditBox(state,
                                    parent,
                                    state.advancedCompareIgnoreFilesPatternsFrame,
                                    state.advancedCompareIgnoreFilesPatternsEdit,
                                    IDC_PREFS_ADV_COMPARE_IGNORE_FILES_PATTERNS_EDIT,
                                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL);
    if (state.advancedCompareIgnoreFilesPatternsEdit)
    {
        SendMessageW(state.advancedCompareIgnoreFilesPatternsEdit.get(), EM_SETLIMITTEXT, 4096, 0);
    }

    state.advancedCompareIgnoreDirectoriesLabel.reset(CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));
    state.advancedCompareIgnoreDirectoriesDescription.reset(
        CreateWindowExW(0, L"Static", L"", wrapStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));
    state.advancedCompareIgnoreDirectoriesPatternsLabel.reset(
        CreateWindowExW(0, L"Static", L"", baseStaticStyle, 0, 0, 10, 10, parent, nullptr, instance, nullptr));
    PrefsInput::CreateFramedEditBox(state,
                                    parent,
                                    state.advancedCompareIgnoreDirectoriesPatternsFrame,
                                    state.advancedCompareIgnoreDirectoriesPatternsEdit,
                                    IDC_PREFS_ADV_COMPARE_IGNORE_DIRECTORIES_PATTERNS_EDIT,
                                    WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL);
    if (state.advancedCompareIgnoreDirectoriesPatternsEdit)
    {
        SendMessageW(state.advancedCompareIgnoreDirectoriesPatternsEdit.get(), EM_SETLIMITTEXT, 4096, 0);
    }

    if (state.advancedCompareContentWorkersCombo)
    {
        SendMessageW(state.advancedCompareContentWorkersCombo.get(), CB_RESETCONTENT, 0, 0);

        const std::wstring autoLabel = LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_CONTENT_WORKERS_AUTO);
        const LRESULT autoIndex =
            SendMessageW(state.advancedCompareContentWorkersCombo.get(), CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(autoLabel.c_str()));
        if (autoIndex != CB_ERR && autoIndex != CB_ERRSPACE)
        {
            SendMessageW(state.advancedCompareContentWorkersCombo.get(), CB_SETITEMDATA, static_cast<WPARAM>(autoIndex), static_cast<LPARAM>(0));
        }

        for (uint32_t i = 1; i <= 4; ++i)
        {
            const std::wstring label = std::to_wstring(i);
            const LRESULT index =
                SendMessageW(state.advancedCompareContentWorkersCombo.get(), CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(label.c_str()));
            if (index != CB_ERR && index != CB_ERRSPACE)
            {
                SendMessageW(state.advancedCompareContentWorkersCombo.get(), CB_SETITEMDATA, static_cast<WPARAM>(index), static_cast<LPARAM>(i));
            }
        }

        ThemedControls::ApplyThemeToComboBox(state.advancedCompareContentWorkersCombo, state.theme);
    }

    Refresh(parent, state);
}

void CompareDirectoriesPane::Refresh(HWND /*host*/, PreferencesDialogState& state) noexcept
{
    const auto& compare = GetCompareDirectoriesSettingsOrDefault(state.workingSettings);

    PrefsUi::SetTwoStateToggleState(state.advancedCompareSizeToggle, state.theme.systemHighContrast, compare.compareSize);
    PrefsUi::SetTwoStateToggleState(state.advancedCompareDateTimeToggle, state.theme.systemHighContrast, compare.compareDateTime);
    PrefsUi::SetTwoStateToggleState(state.advancedCompareAttributesToggle, state.theme.systemHighContrast, compare.compareAttributes);
    PrefsUi::SetTwoStateToggleState(state.advancedCompareContentToggle, state.theme.systemHighContrast, compare.compareContent);

    PrefsUi::SetTwoStateToggleState(state.advancedCompareSubdirectoriesToggle, state.theme.systemHighContrast, compare.compareSubdirectories);
    PrefsUi::SetTwoStateToggleState(state.advancedCompareSubdirectoryAttributesToggle, state.theme.systemHighContrast, compare.compareSubdirectoryAttributes);
    PrefsUi::SetTwoStateToggleState(state.advancedCompareSelectSubdirsOnlyInOnePaneToggle, state.theme.systemHighContrast, compare.selectSubdirsOnlyInOnePane);

    PrefsUi::SetTwoStateToggleState(state.advancedCompareShowIdenticalToggle, state.theme.systemHighContrast, compare.showIdenticalItems);

    PrefsUi::SetTwoStateToggleState(state.advancedCompareIgnoreFilesToggle, state.theme.systemHighContrast, compare.ignoreFiles);
    PrefsUi::SetTwoStateToggleState(state.advancedCompareIgnoreDirectoriesToggle, state.theme.systemHighContrast, compare.ignoreDirectories);

    PrefsUi::SelectComboItemByData(state.advancedCompareContentWorkersCombo, static_cast<LPARAM>(compare.contentCompareWorkerCount));

    if (state.advancedCompareIgnoreFilesPatternsEdit)
    {
        SetWindowTextW(state.advancedCompareIgnoreFilesPatternsEdit.get(), compare.ignoreFilesPatterns.c_str());
    }
    if (state.advancedCompareIgnoreDirectoriesPatternsEdit)
    {
        SetWindowTextW(state.advancedCompareIgnoreDirectoriesPatternsEdit.get(), compare.ignoreDirectoriesPatterns.c_str());
    }
}

void CompareDirectoriesPane::LayoutControls(HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, HFONT dialogFont) noexcept
{
    using namespace PrefsLayoutConstants;

    static_cast<void>(margin);

    if (! host)
    {
        return;
    }

    const UINT dpi = GetDpiForWindow(host);

    const int rowHeight   = std::max(1, ThemedControls::ScaleDip(dpi, kRowHeightDip));
    const int titleHeight = std::max(1, ThemedControls::ScaleDip(dpi, kTitleHeightDip));

    const int cardPaddingX = ThemedControls::ScaleDip(dpi, kCardPaddingXDip);
    const int cardPaddingY = ThemedControls::ScaleDip(dpi, kCardPaddingYDip);
    const int cardGapY     = ThemedControls::ScaleDip(dpi, kCardGapYDip);
    const int cardGapX     = ThemedControls::ScaleDip(dpi, kCardGapXDip);
    const int cardSpacingY = ThemedControls::ScaleDip(dpi, kCardSpacingYDip);

    const HFONT headerFont = state.boldFont ? state.boldFont.get() : dialogFont;
    const HFONT infoFont   = state.italicFont ? state.italicFont.get() : dialogFont;
    const int headerHeight = std::max(1, ThemedControls::ScaleDip(dpi, kHeaderHeightDip));

    const int minToggleWidth    = ThemedControls::ScaleDip(dpi, kMinToggleWidthDip);
    const std::wstring onLabel  = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);
    const std::wstring offLabel = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);

    const HFONT toggleMeasureFont = state.boldFont ? state.boldFont.get() : dialogFont;
    const int onWidth             = ThemedControls::MeasureTextWidth(host, toggleMeasureFont, onLabel);
    const int offWidth            = ThemedControls::MeasureTextWidth(host, toggleMeasureFont, offLabel);

    const int paddingX       = ThemedControls::ScaleDip(dpi, kTogglePaddingXDip);
    const int gapX           = ThemedControls::ScaleDip(dpi, kToggleGapXDip);
    const int trackWidth     = ThemedControls::ScaleDip(dpi, kToggleTrackWidthDip);
    const int stateTextWidth = std::max(onWidth, offWidth);

    const int measuredToggleWidth = std::max(minToggleWidth, (2 * paddingX) + stateTextWidth + gapX + trackWidth);
    const int toggleWidth         = std::min(std::max(0, width - 2 * cardPaddingX), measuredToggleWidth);

    const auto& compare = GetCompareDirectoriesSettingsOrDefault(state.workingSettings);

    auto pushCard = [&](const RECT& card) noexcept { state.pageSettingCards.push_back(card); };

    auto layoutToggleCard = [&](HWND label, std::wstring_view labelText, HWND toggle, HWND descLabel, std::wstring_view descText) noexcept
    {
        const int textWidth  = std::max(0, width - 2 * cardPaddingX - cardGapX - toggleWidth);
        const int descHeight = descLabel ? PrefsUi::MeasureStaticTextHeight(host, infoFont, textWidth, descText) : 0;

        const int contentHeight = std::max(0, titleHeight + cardGapY + descHeight);
        const int cardHeight    = std::max(rowHeight + 2 * cardPaddingY, contentHeight + 2 * cardPaddingY);

        RECT card{};
        card.left   = x;
        card.top    = y;
        card.right  = x + width;
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

    auto layoutFramedComboCard = [&](HWND label,
                                    std::wstring_view labelText,
                                    HWND frame,
                                    HWND combo,
                                    HWND descLabel,
                                    std::wstring_view descText) noexcept
    {
        int desiredWidth          = combo ? ThemedControls::MeasureComboBoxPreferredWidth(combo, dpi) : 0;
        desiredWidth              = std::max(desiredWidth, ThemedControls::ScaleDip(dpi, kMinEditWidthDip));
        const int maxControlWidth = std::max(0, width - 2 * cardPaddingX);
        desiredWidth              = std::min(desiredWidth, std::min(maxControlWidth, ThemedControls::ScaleDip(dpi, kMaxEditWidthDip)));

        const int textWidth  = std::max(0, width - 2 * cardPaddingX - cardGapX - desiredWidth);
        const int descHeight = descLabel ? PrefsUi::MeasureStaticTextHeight(host, infoFont, textWidth, descText) : 0;

        const int contentHeight = std::max(0, titleHeight + cardGapY + descHeight);
        const int cardHeight    = std::max(rowHeight + 2 * cardPaddingY, contentHeight + 2 * cardPaddingY);

        RECT card{};
        card.left   = x;
        card.top    = y;
        card.right  = x + width;
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

        const int inputX       = card.right - cardPaddingX - desiredWidth;
        const int inputY       = card.top + (cardHeight - rowHeight) / 2;
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

    auto layoutIgnoreCard = [&](HWND label,
                                std::wstring_view labelText,
                                HWND toggle,
                                HWND descLabel,
                                std::wstring_view descText,
                                HWND patternsLabel,
                                HWND patternsFrame,
                                HWND patternsEdit,
                                bool showEdit) noexcept
    {
        const int textWidth  = std::max(0, width - 2 * cardPaddingX - cardGapX - toggleWidth);
        const int descHeight = descLabel ? PrefsUi::MeasureStaticTextHeight(host, infoFont, textWidth, descText) : 0;

        int contentHeight = std::max(0, titleHeight + cardGapY + descHeight);
        if (showEdit)
        {
            contentHeight += cardGapY + rowHeight;
        }
        const int cardHeight = std::max(rowHeight + 2 * cardPaddingY, contentHeight + 2 * cardPaddingY);

        RECT card{};
        card.left   = x;
        card.top    = y;
        card.right  = x + width;
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
            // Keep the toggle aligned with the title row even when the card expands to show the pattern field.
            SetWindowPos(
                toggle, nullptr, card.right - cardPaddingX - toggleWidth, card.top + cardPaddingY, toggleWidth, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            SendMessageW(toggle, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
        }

        if (patternsLabel)
        {
            ShowWindow(patternsLabel, SW_HIDE);
        }

        if (patternsFrame)
        {
            ShowWindow(patternsFrame, showEdit ? SW_SHOW : SW_HIDE);
            EnableWindow(patternsFrame, showEdit ? TRUE : FALSE);
            InvalidateRect(patternsFrame, nullptr, TRUE);
        }
        if (patternsEdit)
        {
            ShowWindow(patternsEdit, showEdit ? SW_SHOW : SW_HIDE);
            EnableWindow(patternsEdit, showEdit ? TRUE : FALSE);
            InvalidateRect(patternsEdit, nullptr, TRUE);
        }

        if (showEdit && patternsEdit)
        {
            const int editX = card.left + cardPaddingX;
            const int editW = std::max(0, width - 2 * cardPaddingX);

            const int contentTop    = card.top + cardPaddingY;
            const int contentBottom = contentTop + titleHeight + cardGapY + descHeight;
            const int editTop       = contentBottom + cardGapY;

            const int framePadding = (patternsFrame && ! state.theme.systemHighContrast) ? ThemedControls::ScaleDip(dpi, kFramePaddingDip) : 0;

            if (patternsFrame)
            {
                SetWindowPos(patternsFrame, nullptr, editX, editTop, editW, rowHeight, SWP_NOZORDER | SWP_NOACTIVATE);
            }
            if (patternsEdit)
            {
                const int innerW = std::max(1, editW - 2 * framePadding);
                const int innerH = std::max(1, rowHeight - 2 * framePadding);
                SetWindowPos(patternsEdit, nullptr, editX + framePadding, editTop + framePadding, innerW, innerH, SWP_NOZORDER | SWP_NOACTIVATE);
                SendMessageW(patternsEdit, WM_SETFONT, reinterpret_cast<WPARAM>(dialogFont), TRUE);
            }
        }

        y += cardHeight + cardSpacingY;
    };

    if (state.advancedCompareDirectoriesHeader)
    {
        SetWindowTextW(state.advancedCompareDirectoriesHeader.get(), LoadStringResource(nullptr, IDS_PREFS_ADV_HEADER_COMPARE_DIRECTORIES).c_str());
        SetWindowPos(state.advancedCompareDirectoriesHeader.get(), nullptr, x, y, width, headerHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(state.advancedCompareDirectoriesHeader.get(), WM_SETFONT, reinterpret_cast<WPARAM>(headerFont), TRUE);
        y += headerHeight + gapY;
    }

    const auto layoutSectionHeader = [&](const wil::unique_hwnd& header, UINT textId) noexcept
    {
        if (! header)
        {
            return;
        }

        SetWindowTextW(header.get(), LoadStringResource(nullptr, textId).c_str());
        SetWindowPos(header.get(), nullptr, x, y, width, headerHeight, SWP_NOZORDER | SWP_NOACTIVATE);
        SendMessageW(header.get(), WM_SETFONT, reinterpret_cast<WPARAM>(headerFont), TRUE);
        y += headerHeight + gapY;
    };

    // 1) Subdirectories options
    layoutSectionHeader(state.advancedCompareSectionSubdirsHeader, IDS_COMPARE_OPTIONS_SECTION_SUBDIRS);
    layoutToggleCard(state.advancedCompareSubdirectoriesLabel.get(),
                     LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_SUBDIRS_TITLE),
                     state.advancedCompareSubdirectoriesToggle.get(),
                     state.advancedCompareSubdirectoriesDescription.get(),
                     LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_SUBDIRS_DESC));
    y += gapY;

    // 2) Compare files with same name by
    layoutSectionHeader(state.advancedCompareSectionCompareHeader, IDS_COMPARE_OPTIONS_SECTION_COMPARE);
    layoutToggleCard(state.advancedCompareSizeLabel.get(),
                     LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_SIZE_TITLE),
                     state.advancedCompareSizeToggle.get(),
                     state.advancedCompareSizeDescription.get(),
                     LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_SIZE_DESC));
    layoutToggleCard(state.advancedCompareDateTimeLabel.get(),
                     LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_DATETIME_TITLE),
                     state.advancedCompareDateTimeToggle.get(),
                     state.advancedCompareDateTimeDescription.get(),
                     LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_DATETIME_DESC));
    layoutToggleCard(state.advancedCompareAttributesLabel.get(),
                     LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_ATTRIBUTES_TITLE),
                     state.advancedCompareAttributesToggle.get(),
                     state.advancedCompareAttributesDescription.get(),
                     LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_ATTRIBUTES_DESC));
    layoutToggleCard(state.advancedCompareContentLabel.get(),
                     LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_CONTENT_TITLE),
                     state.advancedCompareContentToggle.get(),
                     state.advancedCompareContentDescription.get(),
                     LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_CONTENT_DESC));
    layoutFramedComboCard(state.advancedCompareContentWorkersLabel.get(),
                          LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_CONTENT_WORKERS_TITLE),
                          state.advancedCompareContentWorkersFrame.get(),
                          state.advancedCompareContentWorkersCombo.get(),
                          state.advancedCompareContentWorkersDescription.get(),
                          LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_CONTENT_WORKERS_DESC));
    y += gapY;

    // 3) Additional options
    layoutSectionHeader(state.advancedCompareSectionAdditionalHeader, IDS_COMPARE_OPTIONS_SECTION_ADVANCED);
    layoutToggleCard(state.advancedCompareSubdirectoryAttributesLabel.get(),
                     LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_SUBDIR_ATTRIBUTES_TITLE),
                     state.advancedCompareSubdirectoryAttributesToggle.get(),
                     state.advancedCompareSubdirectoryAttributesDescription.get(),
                     LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_SUBDIR_ATTRIBUTES_DESC));
    layoutToggleCard(state.advancedCompareSelectSubdirsOnlyInOnePaneLabel.get(),
                     LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_SELECT_SUBDIRS_TITLE),
                     state.advancedCompareSelectSubdirsOnlyInOnePaneToggle.get(),
                     state.advancedCompareSelectSubdirsOnlyInOnePaneDescription.get(),
                     LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_SELECT_SUBDIRS_DESC));
    layoutToggleCard(state.advancedCompareShowIdenticalLabel.get(),
                     LoadStringResource(nullptr, IDS_PREFS_COMPARE_SHOW_IDENTICAL_TITLE),
                     state.advancedCompareShowIdenticalToggle.get(),
                     state.advancedCompareShowIdenticalDescription.get(),
                     LoadStringResource(nullptr, IDS_PREFS_COMPARE_SHOW_IDENTICAL_DESC));
    y += gapY;

    // 4) More options
    layoutSectionHeader(state.advancedCompareSectionMoreHeader, IDS_COMPARE_OPTIONS_SECTION_IGNORE);
    layoutIgnoreCard(state.advancedCompareIgnoreFilesLabel.get(),
                     LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_IGNORE_FILES_TITLE),
                     state.advancedCompareIgnoreFilesToggle.get(),
                     state.advancedCompareIgnoreFilesDescription.get(),
                     LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_IGNORE_FILES_DESC),
                     state.advancedCompareIgnoreFilesPatternsLabel.get(),
                     state.advancedCompareIgnoreFilesPatternsFrame.get(),
                     state.advancedCompareIgnoreFilesPatternsEdit.get(),
                     compare.ignoreFiles);
    layoutIgnoreCard(state.advancedCompareIgnoreDirectoriesLabel.get(),
                     LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_IGNORE_DIRECTORIES_TITLE),
                     state.advancedCompareIgnoreDirectoriesToggle.get(),
                     state.advancedCompareIgnoreDirectoriesDescription.get(),
                     LoadStringResource(nullptr, IDS_COMPARE_OPTIONS_IGNORE_DIRECTORIES_DESC),
                     state.advancedCompareIgnoreDirectoriesPatternsLabel.get(),
                     state.advancedCompareIgnoreDirectoriesPatternsFrame.get(),
                     state.advancedCompareIgnoreDirectoriesPatternsEdit.get(),
                     compare.ignoreDirectories);
}

bool CompareDirectoriesPane::HandleCommand(HWND host, PreferencesDialogState& state, UINT commandId, UINT notifyCode, HWND hwndCtl) noexcept
{
    if (commandId == IDC_PREFS_ADV_COMPARE_CONTENT_WORKERS_COMBO && notifyCode == CBN_SELCHANGE)
    {
        const auto dataOpt = PrefsUi::TryGetSelectedComboItemData(state.advancedCompareContentWorkersCombo);
        if (! dataOpt.has_value())
        {
            return true;
        }

        const uint32_t newValue = std::clamp(static_cast<uint32_t>(dataOpt.value()), 0u, 4u);

        auto* compare = EnsureWorkingCompareDirectoriesSettings(state.workingSettings);
        if (! compare)
        {
            return true;
        }

        if (compare->contentCompareWorkerCount != newValue)
        {
            compare->contentCompareWorkerCount = newValue;
            MaybeResetWorkingCompareDirectoriesSettingsIfEmpty(state.workingSettings);
            SetDirty(GetParent(host), state);
            Refresh(host, state);
        }

        return true;
    }

    const bool isComparePatternEdit =
        (commandId == IDC_PREFS_ADV_COMPARE_IGNORE_FILES_PATTERNS_EDIT || commandId == IDC_PREFS_ADV_COMPARE_IGNORE_DIRECTORIES_PATTERNS_EDIT);
    if (isComparePatternEdit)
    {
        if (notifyCode == EN_CHANGE || notifyCode == EN_KILLFOCUS)
        {
            HWND edit               = hwndCtl ? hwndCtl : GetDlgItem(host, static_cast<int>(commandId));
            const std::wstring text = PrefsUi::GetWindowTextString(edit);
            const bool commit       = (notifyCode == EN_KILLFOCUS);

            const std::wstring_view trimmed = commit ? PrefsUi::TrimWhitespace(text) : std::wstring_view(text);
            const std::wstring newValue(trimmed);

            auto* compare = EnsureWorkingCompareDirectoriesSettings(state.workingSettings);
            if (! compare)
            {
                return true;
            }

            bool changed = false;
            if (commandId == IDC_PREFS_ADV_COMPARE_IGNORE_FILES_PATTERNS_EDIT)
            {
                if (compare->ignoreFilesPatterns != newValue)
                {
                    compare->ignoreFilesPatterns = newValue;
                    changed                      = true;
                }
            }
            else
            {
                if (compare->ignoreDirectoriesPatterns != newValue)
                {
                    compare->ignoreDirectoriesPatterns = newValue;
                    changed                            = true;
                }
            }

            if (changed)
            {
                MaybeResetWorkingCompareDirectoriesSettingsIfEmpty(state.workingSettings);
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

    if (notifyCode == BN_CLICKED)
    {
        const bool isCompareToggle =
            (commandId == IDC_PREFS_ADV_COMPARE_SIZE_TOGGLE || commandId == IDC_PREFS_ADV_COMPARE_DATETIME_TOGGLE ||
             commandId == IDC_PREFS_ADV_COMPARE_ATTRIBUTES_TOGGLE || commandId == IDC_PREFS_ADV_COMPARE_CONTENT_TOGGLE ||
             commandId == IDC_PREFS_ADV_COMPARE_SUBDIRS_TOGGLE || commandId == IDC_PREFS_ADV_COMPARE_SUBDIR_ATTRIBUTES_TOGGLE ||
             commandId == IDC_PREFS_ADV_COMPARE_SELECT_SUBDIRS_ONE_PANE_TOGGLE || commandId == IDC_PREFS_ADV_COMPARE_SHOW_IDENTICAL_TOGGLE ||
             commandId == IDC_PREFS_ADV_COMPARE_IGNORE_FILES_TOGGLE || commandId == IDC_PREFS_ADV_COMPARE_IGNORE_DIRECTORIES_TOGGLE);
        if (isCompareToggle)
        {
            if (hwndCtl)
            {
                const bool ownerDraw = (GetWindowLongPtrW(hwndCtl, GWL_STYLE) & BS_TYPEMASK) == BS_OWNERDRAW;
                if (ownerDraw)
                {
                    const bool current = PrefsUi::GetTwoStateToggleState(hwndCtl, false);
                    PrefsUi::SetTwoStateToggleState(hwndCtl, false, ! current);
                }
            }

            const bool toggledOn = PrefsUi::GetTwoStateToggleState(hwndCtl, state.theme.systemHighContrast);

            auto* compare = EnsureWorkingCompareDirectoriesSettings(state.workingSettings);
            if (! compare)
            {
                return true;
            }

            switch (commandId)
            {
                case IDC_PREFS_ADV_COMPARE_SIZE_TOGGLE: compare->compareSize = toggledOn; break;
                case IDC_PREFS_ADV_COMPARE_DATETIME_TOGGLE: compare->compareDateTime = toggledOn; break;
                case IDC_PREFS_ADV_COMPARE_ATTRIBUTES_TOGGLE: compare->compareAttributes = toggledOn; break;
                case IDC_PREFS_ADV_COMPARE_CONTENT_TOGGLE: compare->compareContent = toggledOn; break;
                case IDC_PREFS_ADV_COMPARE_SUBDIRS_TOGGLE: compare->compareSubdirectories = toggledOn; break;
                case IDC_PREFS_ADV_COMPARE_SUBDIR_ATTRIBUTES_TOGGLE: compare->compareSubdirectoryAttributes = toggledOn; break;
                case IDC_PREFS_ADV_COMPARE_SELECT_SUBDIRS_ONE_PANE_TOGGLE: compare->selectSubdirsOnlyInOnePane = toggledOn; break;
                case IDC_PREFS_ADV_COMPARE_SHOW_IDENTICAL_TOGGLE: compare->showIdenticalItems = toggledOn; break;
                case IDC_PREFS_ADV_COMPARE_IGNORE_FILES_TOGGLE: compare->ignoreFiles = toggledOn; break;
                case IDC_PREFS_ADV_COMPARE_IGNORE_DIRECTORIES_TOGGLE: compare->ignoreDirectories = toggledOn; break;
                default: break;
            }

            MaybeResetWorkingCompareDirectoriesSettingsIfEmpty(state.workingSettings);
            SetDirty(GetParent(host), state);
            Refresh(host, state);

            if (host && (commandId == IDC_PREFS_ADV_COMPARE_IGNORE_FILES_TOGGLE || commandId == IDC_PREFS_ADV_COMPARE_IGNORE_DIRECTORIES_TOGGLE))
            {
                RECT rc{};
                if (GetClientRect(host, &rc))
                {
                    const int cx = std::max(0l, rc.right - rc.left);
                    const int cy = std::max(0l, rc.bottom - rc.top);
                    SendMessageW(host, WM_SIZE, SIZE_RESTORED, MAKELPARAM(cx, cy));
                }
            }
            return true;
        }
    }

    return false;
}
