// Preferences.HotPaths.cpp

#include "Framework.h"

#include "Preferences.HotPaths.h"

#include <array>
#include <filesystem>
#include <format>
#include <mutex>
#include <optional>
#include <string>

#include <shobjidl.h>

#include <wil/com.h>

#include "Helpers.h"
#include "UiMetrics.h"
#include "WindowMessages.h"
#include "resource.h"

namespace
{
using PrefsHotPaths::EnsureWorkingHotPathsSettings;
using PrefsHotPaths::GetHotPathsSettingsOrDefault;
using PrefsHotPaths::MaybeResetWorkingHotPathsSettingsIfEmpty;
using RedSalamander::DxUi::Button;
using RedSalamander::DxUi::CardPanel;
using RedSalamander::DxUi::FontRole;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::TextField;
using RedSalamander::DxUi::ThemePalette;
using RedSalamander::DxUi::Toggle;

constexpr int kSlotCount = 10;

#ifdef ENABLE_TESTS
enum class DebugHotPathsBrowseResultKind
{
    Path,
    Cancel,
};

struct DebugHotPathsBrowseResult
{
    DebugHotPathsBrowseResultKind kind = DebugHotPathsBrowseResultKind::Path;
    std::filesystem::path path{};
};

std::mutex g_debugHotPathsBrowseResultMutex;
std::optional<DebugHotPathsBrowseResult> g_debugNextHotPathsBrowseResult;
#endif

struct HotPathsDxSlot
{
    Label* header            = nullptr;
    CardPanel* card          = nullptr;
    Label* pathLabel         = nullptr;
    TextField* pathEdit      = nullptr;
    Button* browseButton     = nullptr;
    Label* labelLabel        = nullptr;
    TextField* labelEdit     = nullptr;
    Label* showInMenuLabel   = nullptr;
    Toggle* showInMenuToggle = nullptr;
};

struct HotPathsDxPage
{
    HotPathsDxPage()                                 = default;
    HotPathsDxPage(const HotPathsDxPage&)            = delete;
    HotPathsDxPage& operator=(const HotPathsDxPage&) = delete;
    HotPathsDxPage(HotPathsDxPage&&)                 = delete;
    HotPathsDxPage& operator=(HotPathsDxPage&&)      = delete;

    std::array<HotPathsDxSlot, static_cast<size_t>(kSlotCount)> slots;
    CardPanel* openPrefsCard    = nullptr;
    Label* openPrefsLabel       = nullptr;
    Label* openPrefsDescription = nullptr;
    Toggle* openPrefsToggle     = nullptr;

    void Detach() noexcept
    {
        slots                = {};
        openPrefsCard        = nullptr;
        openPrefsLabel       = nullptr;
        openPrefsDescription = nullptr;
        openPrefsToggle      = nullptr;
    }
};

} // namespace

struct HotPathsPane::DxState
{
    DxState()                          = default;
    DxState(const DxState&)            = delete;
    DxState& operator=(const DxState&) = delete;
    DxState(DxState&&)                 = delete;
    DxState& operator=(DxState&&)      = delete;

    HotPathsDxPage page;

    void Detach() noexcept
    {
        page.Detach();
    }
};

HotPathsPane::HotPathsPane()  = default;
HotPathsPane::~HotPathsPane() = default;

void HotPathsPane::OnVisibilityChanged(bool visible) noexcept
{
    static_cast<void>(visible);
}

void HotPathsPane::Destroy(PreferencesDialogState& state) noexcept
{
    DetachDxHosts();
    static_cast<void>(state);
    _pageHost = nullptr;
}

bool HotPathsPane::EnsureDxHosts(HWND parent, PreferencesDialogState& state) noexcept
{
    _pageHostDx      = state.pageHostDxHost;
    _pageContentRoot = state.pageHostDxContentRootControl;
    if (! _pageHostDx || ! _pageContentRoot)
    {
        return false;
    }

    if (_dxState && PrefsUi::HasRetainedDxChildren(_pageContentRoot))
    {
        ApplyDxTheme(state);
        SyncDxControlsFromState(state);
        return true;
    }

    auto dxState = std::make_unique<DxState>();
    _pageHostDx->ResetInteractionState();
    _pageContentRoot->ClearChildren();

    auto* root = _pageContentRoot;

    for (int i = 0; i < kSlotCount; ++i)
    {
        HotPathsDxSlot& dxSlot = dxState->page.slots[static_cast<size_t>(i)];

        dxSlot.header = root->AddChild<Label>();
        dxSlot.header->SetFontRole(FontRole::Header);
        dxSlot.card      = root->AddChild<CardPanel>();
        dxSlot.pathLabel = root->AddChild<Label>();
        dxSlot.pathLabel->SetFontRole(FontRole::Body);
        dxSlot.pathEdit     = root->AddChild<TextField>();
        dxSlot.browseButton = root->AddChild<Button>(LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_BROWSE_ELLIPSIS));
        dxSlot.labelLabel   = root->AddChild<Label>();
        dxSlot.labelLabel->SetFontRole(FontRole::Body);
        dxSlot.labelEdit       = root->AddChild<TextField>();
        dxSlot.showInMenuLabel = root->AddChild<Label>();
        dxSlot.showInMenuLabel->SetFontRole(FontRole::Body);
        dxSlot.showInMenuLabel->SetAlignment(DWRITE_TEXT_ALIGNMENT_TRAILING);
        dxSlot.showInMenuToggle = root->AddChild<Toggle>();
        dxSlot.showInMenuToggle->SetStateLabels(LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF), LoadStringResource(nullptr, IDS_PREFS_COMMON_ON));

        dxSlot.pathEdit->SetOnTextChanged([this, host = parent, slotIndex = static_cast<size_t>(i)](std::wstring_view text) noexcept
        {
            if (_syncingDxPathEdits[slotIndex] || ! host || IsWindow(host) == FALSE)
            {
                return;
            }

            auto* state = PrefsUi::GetDialogState(host);
            if (! state)
            {
                return;
            }

            auto* hp = EnsureWorkingHotPathsSettings(state->workingSettings);
            if (! hp)
            {
                return;
            }

            if (slotIndex >= hp->slots.size())
            {
                return;
            }

            const std::wstring_view trimmed = PrefsUi::TrimWhitespace(text);
            const std::wstring newPath(trimmed);
            auto& slot = hp->slots[slotIndex];

            if (newPath.empty())
            {
                if (slot.has_value() && ! slot.value().path.empty())
                {
                    slot.value().path.clear();
                    MaybeResetWorkingHotPathsSettingsIfEmpty(state->workingSettings);
                    SetDirty(GetParent(host), *state);
                    SyncDxControlsFromState(*state);
                }
                return;
            }

            if (! slot.has_value())
            {
                slot = Common::Settings::HotPathSlot{};
            }

            if (slot.value().path != newPath)
            {
                slot.value().path = newPath;
                MaybeResetWorkingHotPathsSettingsIfEmpty(state->workingSettings);
                SetDirty(GetParent(host), *state);
                SyncDxControlsFromState(*state);
            }
        });
        dxSlot.pathEdit->SetOnBlur([this, host = parent]() noexcept
        {
            if (! host || IsWindow(host) == FALSE)
            {
                return;
            }

            auto* state = PrefsUi::GetDialogState(host);
            if (! state)
            {
                return;
            }
            Refresh(host, *state);
        });

        dxSlot.browseButton->SetOnClick([this, host = parent, slotIdx = i]() noexcept
        {
            if (! _dxState || ! host || IsWindow(host) == FALSE)
            {
                return;
            }

            PreferencesDialogState* state = PrefsUi::GetDialogState(host);
            if (! state)
            {
                return;
            }

            OnHotPathBrowseClicked(host, *state, slotIdx);
        });

        dxSlot.labelEdit->SetOnTextChanged([this, host = parent, slotIndex = static_cast<size_t>(i)](std::wstring_view text) noexcept
        {
            if (_syncingDxLabelEdits[slotIndex] || ! host || IsWindow(host) == FALSE)
            {
                return;
            }

            auto* state = PrefsUi::GetDialogState(host);
            if (! state)
            {
                return;
            }

            auto* hp = EnsureWorkingHotPathsSettings(state->workingSettings);
            if (! hp)
            {
                return;
            }

            if (slotIndex >= hp->slots.size())
            {
                return;
            }
            const std::wstring_view trimmed = PrefsUi::TrimWhitespace(text);
            const std::wstring newLabel(trimmed);
            auto& slot = hp->slots[slotIndex];
            if (! slot.has_value())
            {
                if (newLabel.empty())
                {
                    return;
                }

                slot = Common::Settings::HotPathSlot{};
            }

            if (slot.value().label != newLabel)
            {
                slot.value().label = newLabel;
                MaybeResetWorkingHotPathsSettingsIfEmpty(state->workingSettings);
                SetDirty(GetParent(host), *state);
                SyncDxControlsFromState(*state);
            }
        });
        dxSlot.labelEdit->SetOnBlur([this, host = parent]() noexcept
        {
            if (! host || IsWindow(host) == FALSE)
            {
                return;
            }

            auto* state = PrefsUi::GetDialogState(host);
            if (! state)
            {
                return;
            }
            Refresh(host, *state);
        });

        dxSlot.showInMenuToggle->SetOnToggled([this, host = parent, slotIndex = static_cast<size_t>(i)](bool checked) noexcept
        {
            if (! host || IsWindow(host) == FALSE)
            {
                return;
            }

            auto* state = PrefsUi::GetDialogState(host);
            if (! state)
            {
                return;
            }

            auto* hp = EnsureWorkingHotPathsSettings(state->workingSettings);
            if (! hp)
            {
                return;
            }

            if (! hp->slots[slotIndex].has_value())
            {
                hp->slots[slotIndex] = Common::Settings::HotPathSlot{};
            }
            hp->slots[slotIndex].value().showInMenu = checked;
            MaybeResetWorkingHotPathsSettingsIfEmpty(state->workingSettings);

            if (const HWND dlg = GetParent(host); dlg && IsWindow(dlg) != FALSE)
            {
                SetDirty(dlg, *state);
            }

            Refresh(host, *state);
        });
    }

    dxState->page.openPrefsCard  = root->AddChild<CardPanel>();
    dxState->page.openPrefsLabel = root->AddChild<Label>();
    dxState->page.openPrefsLabel->SetFontRole(FontRole::Body);
    dxState->page.openPrefsDescription = root->AddChild<Label>();
    dxState->page.openPrefsDescription->SetFontRole(FontRole::Small);
    dxState->page.openPrefsDescription->SetMultiline(true);
    dxState->page.openPrefsToggle = root->AddChild<Toggle>();
    dxState->page.openPrefsToggle->SetStateLabels(LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF), LoadStringResource(nullptr, IDS_PREFS_COMMON_ON));
    dxState->page.openPrefsToggle->SetOnToggled([this, host = parent](bool checked) noexcept
    {
        if (! host || IsWindow(host) == FALSE)
        {
            return;
        }

        auto* state = PrefsUi::GetDialogState(host);
        if (! state)
        {
            return;
        }

        auto* hp = EnsureWorkingHotPathsSettings(state->workingSettings);
        if (! hp)
        {
            return;
        }

        hp->openPrefsOnAssign = checked;
        MaybeResetWorkingHotPathsSettingsIfEmpty(state->workingSettings);

        if (const HWND dlg = GetParent(host); dlg && IsWindow(dlg) != FALSE)
        {
            SetDirty(dlg, *state);
        }

        Refresh(host, *state);
    });

    _dxState = std::move(dxState);
    ApplyDxTheme(state);
    SyncDxControlsFromState(state);
    return true;
}

void HotPathsPane::DetachDxHosts() noexcept
{
    if (_pageContentRoot && _pageHostDx && _pageHost && IsWindow(_pageHost) != FALSE)
    {
        _pageHostDx->ResetInteractionState();
        _pageContentRoot->ClearChildren();
    }
    _pageHostDx      = nullptr;
    _pageContentRoot = nullptr;

    if (_dxState)
    {
        _dxState->Detach();
        _dxState.reset();
    }
}

void HotPathsPane::ApplyDxTheme(const PreferencesDialogState& state) noexcept
{
    if (! _dxState || ! _pageHostDx)
    {
        return;
    }

    _pageHostDx->SetTheme(PrefsUi::MakeDxPalette(state.theme));
}

void HotPathsPane::SyncDxControlsFromState(const PreferencesDialogState& state) noexcept
{
    if (! _dxState)
    {
        return;
    }

    HotPathsDxPage& page       = _dxState->page;
    const ThemePalette palette = PrefsUi::MakeDxPalette(state.theme);
    const auto& hp             = GetHotPathsSettingsOrDefault(state.workingSettings);

    const std::wstring pathLabelText        = LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_PATH_LABEL);
    const std::wstring labelLabelText       = LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_LABEL_LABEL);
    const std::wstring showInMenuText       = LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_SHOW_IN_MENU);
    const std::wstring openPrefsLabel       = LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_OPEN_PREFS_ON_ASSIGN);
    const std::wstring openPrefsDescription = LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_OPEN_PREFS_ON_ASSIGN_DESC);

    for (int i = 0; i < kSlotCount; ++i)
    {
        HotPathsDxSlot& dxSlot        = page.slots[static_cast<size_t>(i)];
        const auto& slotData          = hp.slots[static_cast<size_t>(i)];
        const bool hasPath            = slotData.has_value() && ! slotData.value().path.empty();
        const bool hasLabel           = slotData.has_value() && ! slotData.value().label.empty();
        const bool hasPathOrLabel     = hasPath || hasLabel;
        const D2D1_COLOR_F labelColor = hasPathOrLabel ? palette.text : palette.disabledText;

        const wchar_t digitChar = (i < 9) ? static_cast<wchar_t>(L'1' + i) : L'0';
        dxSlot.header->SetText(FormatStringResource(nullptr, IDS_PREFS_HOT_PATHS_SLOT_HEADER_FMT, digitChar));
        dxSlot.pathLabel->SetText(pathLabelText);
        dxSlot.pathLabel->SetMnemonicTarget(dxSlot.pathEdit);

        _syncingDxPathEdits[static_cast<size_t>(i)] = true;
        dxSlot.pathEdit->SetText(slotData.has_value() ? slotData.value().path : std::wstring{});
        dxSlot.pathEdit->SetEnabled(true);
        _syncingDxPathEdits[static_cast<size_t>(i)] = false;

        dxSlot.labelLabel->SetText(labelLabelText);
        dxSlot.labelLabel->SetMnemonicTarget(dxSlot.labelEdit);

        _syncingDxLabelEdits[static_cast<size_t>(i)] = true;
        dxSlot.labelEdit->SetText(slotData.has_value() ? slotData.value().label : std::wstring{});
        dxSlot.labelEdit->SetEnabled(hasPathOrLabel);
        _syncingDxLabelEdits[static_cast<size_t>(i)] = false;

        dxSlot.showInMenuLabel->SetText(showInMenuText);
        dxSlot.showInMenuLabel->SetMnemonicTarget(dxSlot.showInMenuToggle);
        dxSlot.header->SetTextColor(std::nullopt);
        dxSlot.pathLabel->SetTextColor(std::nullopt);
        dxSlot.labelLabel->SetTextColor(labelColor);
        dxSlot.showInMenuLabel->SetTextColor(labelColor);
        dxSlot.showInMenuToggle->SetEnabled(hasPath);

        const bool checked = slotData.has_value() && slotData.value().showInMenu;
        dxSlot.showInMenuToggle->SetChecked(checked);
    }

    page.openPrefsLabel->SetText(openPrefsLabel);
    page.openPrefsLabel->SetMnemonicTarget(page.openPrefsToggle);
    page.openPrefsDescription->SetText(openPrefsDescription);
    page.openPrefsLabel->SetTextColor(std::nullopt);
    page.openPrefsDescription->SetTextColor(std::optional<D2D1_COLOR_F>(palette.subduedText));
    page.openPrefsToggle->SetChecked(hp.openPrefsOnAssign);
    page.openPrefsToggle->SetEnabled(true);
    if (_pageHostDx)
    {
        _pageHostDx->Invalidate();
    }
}

void HotPathsPane::InitializePage(HWND parent, PreferencesDialogState& state) noexcept
{
    if (! parent)
    {
        return;
    }

    _pageHost = parent;

    if (state.currentCategory != PrefCategory::HotPaths)
    {
        return;
    }

    if (! EnsureDxHosts(parent, state))
    {
        Debug::Error(L"Preferences.HotPaths: DxUi surface initialization failed; page will not render correctly.");
        DetachDxHosts();
        return;
    }

    Refresh(parent, state);
}

void HotPathsPane::Refresh(HWND /*host*/, PreferencesDialogState& state) noexcept
{
    if (state.currentCategory == PrefCategory::HotPaths && ! _dxState)
    {
        static_cast<void>(EnsureDxHosts(_pageHost, state));
    }

    ApplyDxTheme(state);
    SyncDxControlsFromState(state);
}

void HotPathsPane::LayoutDxPage(HWND host,
                                PreferencesDialogState& state,
                                int x,
                                int& y,
                                int width,
                                [[maybe_unused]] int margin,
                                int gapY,
                                const PreferencesTypographyContext& typography) noexcept
{
    using namespace PrefsLayoutConstants;

    static_cast<void>(host);
    static_cast<void>(state);

    Debug::Perf::Scope layoutPerf(L"preferences.ui.hot_paths_layout_us");
    layoutPerf.SetValue0(static_cast<uint64_t>(std::max(0, width)));
    layoutPerf.SetValue1(static_cast<uint64_t>(typography.dpi));

    const UINT dpi = std::max<UINT>(typography.dpi, USER_DEFAULT_SCREEN_DPI);

    const int rowHeight    = std::max(1, UiMetrics::ScaleDip(dpi, kRowHeightDip));
    const int titleHeight  = std::max(1, UiMetrics::ScaleDip(dpi, kTitleHeightDip));
    const int editHeight   = std::max(1, UiMetrics::ScaleDip(dpi, kEditHeightDip));
    const int headerHeight = std::max(1, UiMetrics::ScaleDip(dpi, kHeaderHeightDip));

    const int cardPaddingX = UiMetrics::ScaleDip(dpi, kCardPaddingXDip);
    const int cardPaddingY = UiMetrics::ScaleDip(dpi, kCardPaddingYDip);
    const int cardGapY     = UiMetrics::ScaleDip(dpi, kCardGapYDip);
    const int cardGapX     = UiMetrics::ScaleDip(dpi, kCardGapXDip);
    const int cardSpacingY = UiMetrics::ScaleDip(dpi, kCardSpacingYDip);

    const int browseWidth = std::max(1, UiMetrics::ScaleDip(dpi, 75));
    const int browseGap   = std::max(1, UiMetrics::ScaleDip(dpi, 4));
    const int innerGap    = std::max(2, gapY / 2);

    const int minToggleWidth    = UiMetrics::ScaleDip(dpi, kMinToggleWidthDip);
    const std::wstring onLabel  = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);
    const std::wstring offLabel = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);

    const int onWidth  = PrefsUi::MeasureSingleLineTextWidthPx(typography, typography.strong, onLabel);
    const int offWidth = PrefsUi::MeasureSingleLineTextWidthPx(typography, typography.strong, offLabel);

    const int tPaddingX      = UiMetrics::ScaleDip(dpi, kTogglePaddingXDip);
    const int tGapX          = UiMetrics::ScaleDip(dpi, kToggleGapXDip);
    const int trackWidth     = UiMetrics::ScaleDip(dpi, kToggleTrackWidthDip);
    const int stateTextWidth = std::max(onWidth, offWidth);

    const int measuredToggleWidth = std::max(minToggleWidth, (2 * tPaddingX) + stateTextWidth + tGapX + trackWidth);
    const int toggleWidth         = std::min(std::max(0, width - 2 * cardPaddingX - cardGapX), measuredToggleWidth);

    DxState* dxState = (_dxState && _pageHostDx && _pageContentRoot) ? _dxState.get() : nullptr;
    if (dxState)
    {
        HotPathsDxPage& dxPage = dxState->page;
        const auto pxToDip     = [dpi](const int px) noexcept { return static_cast<float>(px) * 96.0f / static_cast<float>(std::max<UINT>(1u, dpi)); };

        const int inlineGapX     = std::max(innerGap, UiMetrics::ScaleDip(dpi, 12));
        const int minInlineEditW = UiMetrics::ScaleDip(dpi, kMinEditWidthDip);
        const int textPadX       = UiMetrics::ScaleDip(dpi, 6);
        const int rowGapY        = std::max(innerGap, UiMetrics::ScaleDip(dpi, 12));

        const std::wstring pathLabelText  = LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_PATH_LABEL);
        const std::wstring labelLabelText = LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_LABEL_LABEL);
        const std::wstring showInMenuText = LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_SHOW_IN_MENU);
        const std::wstring assignLabel    = LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_OPEN_PREFS_ON_ASSIGN);
        const std::wstring assignDesc     = LoadStringResource(nullptr, IDS_PREFS_HOT_PATHS_OPEN_PREFS_ON_ASSIGN_DESC);
        const int pathCaptionW            = std::max(0, PrefsUi::MeasureSingleLineTextWidthPx(typography, typography.body, pathLabelText) + textPadX);
        const int labelCaptionW           = std::max(0, PrefsUi::MeasureSingleLineTextWidthPx(typography, typography.body, labelLabelText) + textPadX);
        const int captionW                = std::max(pathCaptionW, labelCaptionW);

        const auto layoutSlotDx = [&](HotPathsDxSlot& slot, const std::wstring& headerText) noexcept
        {
            slot.header->SetVisible(true);
            slot.header->SetText(headerText);
            slot.header->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(y), pxToDip(x + width), pxToDip(y + headerHeight)));
            y += headerHeight + innerGap;

            const int cardTop      = y;
            const int cardContentW = std::max(0, width - 2 * cardPaddingX);
            const int cardLeft     = x;
            const int contentLeft  = x + cardPaddingX;
            const int cardRight    = x + width - cardPaddingX;
            const int browseX      = cardRight - browseWidth;
            const int showTextW    = std::max(0, PrefsUi::MeasureSingleLineTextWidthPx(typography, typography.body, showInMenuText) + textPadX);
            const int showAreaW    = showTextW + inlineGapX + toggleWidth;
            const int showAreaX    = cardRight - showAreaW;
            const int editX        = contentLeft + captionW + inlineGapX;
            const int pathEditW    = browseX - browseGap - editX;
            const int labelEditW   = showAreaX - inlineGapX - editX;
            const bool useTableLayout =
                (cardContentW > 0 && captionW > 0 && pathEditW >= minInlineEditW && labelEditW >= minInlineEditW && showAreaW > toggleWidth);

            slot.card->SetVisible(true);
            slot.pathLabel->SetVisible(true);
            slot.pathEdit->SetVisible(true);
            slot.labelLabel->SetVisible(true);
            slot.labelEdit->SetVisible(true);
            slot.showInMenuLabel->SetVisible(true);
            slot.showInMenuToggle->SetVisible(true);
            slot.pathLabel->SetText(pathLabelText);
            slot.labelLabel->SetText(labelLabelText);
            slot.showInMenuLabel->SetText(showInMenuText);

            if (useTableLayout)
            {
                const int contentTop = cardTop + cardPaddingY;
                const int firstRowY  = contentTop;
                const int secondRowY = contentTop + editHeight + rowGapY;
                const int toggleY    = secondRowY + (editHeight - rowHeight) / 2;
                const int cardHeight = (2 * cardPaddingY) + (2 * editHeight) + rowGapY;
                const int showLabelW = std::max(0, showAreaW - inlineGapX - toggleWidth);

                slot.card->SetBounds(D2D1::RectF(pxToDip(cardLeft), pxToDip(cardTop), pxToDip(cardLeft + width), pxToDip(cardTop + cardHeight)));
                slot.pathLabel->SetBounds(
                    D2D1::RectF(pxToDip(contentLeft), pxToDip(firstRowY), pxToDip(contentLeft + captionW), pxToDip(firstRowY + editHeight)));
                slot.pathEdit->SetBounds(
                    D2D1::RectF(pxToDip(editX), pxToDip(firstRowY), pxToDip(editX + std::max(10, pathEditW)), pxToDip(firstRowY + editHeight)));
                slot.labelLabel->SetBounds(
                    D2D1::RectF(pxToDip(contentLeft), pxToDip(secondRowY), pxToDip(contentLeft + captionW), pxToDip(secondRowY + editHeight)));
                slot.labelEdit->SetBounds(
                    D2D1::RectF(pxToDip(editX), pxToDip(secondRowY), pxToDip(editX + std::max(10, labelEditW)), pxToDip(secondRowY + editHeight)));
                slot.showInMenuLabel->SetBounds(
                    D2D1::RectF(pxToDip(showAreaX), pxToDip(secondRowY), pxToDip(showAreaX + showLabelW), pxToDip(secondRowY + editHeight)));
                slot.showInMenuToggle->SetBounds(
                    D2D1::RectF(pxToDip(cardRight - toggleWidth), pxToDip(toggleY), pxToDip(cardRight), pxToDip(toggleY + rowHeight)));
                y += cardHeight + cardSpacingY;
            }
            else
            {
                int contentY = cardTop + cardPaddingY;
                slot.pathLabel->SetBounds(
                    D2D1::RectF(pxToDip(contentLeft), pxToDip(contentY), pxToDip(contentLeft + cardContentW), pxToDip(contentY + rowHeight)));
                contentY += rowHeight;

                const int stackPathEditW = std::max(0, cardContentW - browseGap - browseWidth);
                slot.pathEdit->SetBounds(
                    D2D1::RectF(pxToDip(contentLeft), pxToDip(contentY), pxToDip(contentLeft + std::max(10, stackPathEditW)), pxToDip(contentY + editHeight)));
                contentY += editHeight + rowGapY;

                slot.labelLabel->SetBounds(
                    D2D1::RectF(pxToDip(contentLeft), pxToDip(contentY), pxToDip(contentLeft + cardContentW), pxToDip(contentY + rowHeight)));
                contentY += rowHeight;

                slot.labelEdit->SetBounds(
                    D2D1::RectF(pxToDip(contentLeft), pxToDip(contentY), pxToDip(contentLeft + std::max(10, cardContentW)), pxToDip(contentY + editHeight)));
                contentY += editHeight + rowGapY;

                slot.showInMenuLabel->SetBounds(D2D1::RectF(pxToDip(contentLeft),
                                                            pxToDip(contentY),
                                                            pxToDip(contentLeft + std::max(0, cardContentW - cardGapX - toggleWidth)),
                                                            pxToDip(contentY + rowHeight)));
                slot.showInMenuToggle->SetBounds(D2D1::RectF(pxToDip(contentLeft + std::max(0, cardContentW - toggleWidth)),
                                                             pxToDip(contentY),
                                                             pxToDip(contentLeft + std::max(0, cardContentW - toggleWidth) + toggleWidth),
                                                             pxToDip(contentY + rowHeight)));
                contentY += rowHeight;

                const int cardHeight = std::max(1, contentY + cardPaddingY - cardTop);
                slot.card->SetBounds(D2D1::RectF(pxToDip(cardLeft), pxToDip(cardTop), pxToDip(cardLeft + width), pxToDip(cardTop + cardHeight)));
                y += cardHeight + cardSpacingY;
            }
        };

        for (int i = 0; i < kSlotCount; ++i)
        {
            const wchar_t digitChar = (i < 9) ? static_cast<wchar_t>(L'1' + i) : L'0';
            layoutSlotDx(dxPage.slots[static_cast<size_t>(i)], FormatStringResource(nullptr, IDS_PREFS_HOT_PATHS_SLOT_HEADER_FMT, digitChar));
        }

        const int openTextWidth     = std::max(0, width - 2 * cardPaddingX - cardGapX - toggleWidth);
        const int openDescHeight    = PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, openTextWidth, assignDesc);
        const int openContentHeight = std::max(0, titleHeight + cardGapY + openDescHeight);
        const int openCardHeight    = std::max(rowHeight + 2 * cardPaddingY, openContentHeight + 2 * cardPaddingY);
        const int openTop           = y;

        dxPage.openPrefsCard->SetVisible(true);
        dxPage.openPrefsLabel->SetVisible(true);
        dxPage.openPrefsDescription->SetVisible(true);
        dxPage.openPrefsToggle->SetVisible(true);
        dxPage.openPrefsLabel->SetText(assignLabel);
        dxPage.openPrefsDescription->SetText(assignDesc);
        dxPage.openPrefsCard->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(openTop), pxToDip(x + width), pxToDip(openTop + openCardHeight)));
        dxPage.openPrefsLabel->SetBounds(D2D1::RectF(pxToDip(x + cardPaddingX),
                                                     pxToDip(openTop + cardPaddingY),
                                                     pxToDip(x + cardPaddingX + openTextWidth),
                                                     pxToDip(openTop + cardPaddingY + titleHeight)));
        dxPage.openPrefsDescription->SetBounds(D2D1::RectF(pxToDip(x + cardPaddingX),
                                                           pxToDip(openTop + cardPaddingY + titleHeight + cardGapY),
                                                           pxToDip(x + cardPaddingX + openTextWidth),
                                                           pxToDip(openTop + cardPaddingY + titleHeight + cardGapY + openDescHeight)));
        dxPage.openPrefsToggle->SetBounds(D2D1::RectF(pxToDip(x + width - cardPaddingX - toggleWidth),
                                                      pxToDip(openTop + (openCardHeight - rowHeight) / 2),
                                                      pxToDip(x + width - cardPaddingX),
                                                      pxToDip(openTop + (openCardHeight - rowHeight) / 2 + rowHeight)));

        y += openCardHeight + cardSpacingY;

        _pageHostDx->Invalidate();
        return;
    }
}

void HotPathsPane::LayoutPage(
    HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, const PreferencesTypographyContext& typography) noexcept
{
    if (! host)
    {
        return;
    }

    if (EnsureDxHosts(_pageHost ? _pageHost : host, state))
    {
        LayoutDxPage(host, state, x, y, width, margin, gapY, typography);
        return;
    }

    Debug::Error(L"Preferences.HotPaths: DxUi surface initialization failed; page will not render correctly.");
}

void HotPathsPane::OnHotPathBrowseClicked(HWND host, PreferencesDialogState& state, int slotIndex) noexcept
{
    if (slotIndex < 0 || slotIndex >= kSlotCount)
    {
        return;
    }

    std::filesystem::path chosenPath;
#ifdef ENABLE_TESTS
    {
        std::scoped_lock lock(g_debugHotPathsBrowseResultMutex);
        if (g_debugNextHotPathsBrowseResult.has_value())
        {
            const DebugHotPathsBrowseResult result = *g_debugNextHotPathsBrowseResult;
            g_debugNextHotPathsBrowseResult.reset();
            if (result.kind == DebugHotPathsBrowseResultKind::Cancel)
            {
                return;
            }

            chosenPath = result.path;
        }
    }
#endif

    if (chosenPath.empty())
    {
        wil::com_ptr<IFileOpenDialog> dialog;
        HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&dialog));
        if (FAILED(hr))
        {
            return;
        }

        DWORD options = 0;
        dialog->GetOptions(&options);
        dialog->SetOptions(options | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);

        hr = dialog->Show(host);
        if (FAILED(hr))
        {
            return;
        }

        wil::com_ptr<IShellItem> result;
        hr = dialog->GetResult(&result);
        if (FAILED(hr) || ! result)
        {
            return;
        }

        wil::unique_cotaskmem_string pathStr;
        hr = result->GetDisplayName(SIGDN_FILESYSPATH, &pathStr);
        if (FAILED(hr) || ! pathStr)
        {
            return;
        }

        chosenPath = std::filesystem::path(pathStr.get());
    }

    const size_t idx = static_cast<size_t>(slotIndex);
    auto* hp         = EnsureWorkingHotPathsSettings(state.workingSettings);
    if (! hp)
    {
        return;
    }

    if (! hp->slots[idx].has_value())
    {
        hp->slots[idx] = Common::Settings::HotPathSlot{};
    }
    hp->slots[idx].value().path = chosenPath.native();

    MaybeResetWorkingHotPathsSettingsIfEmpty(state.workingSettings);
    SetDirty(GetParent(host), state);
    Refresh(host, state);
}

#ifdef ENABLE_TESTS
bool DebugSetPreferencesHotPathsNextBrowsePath(const std::wstring_view path) noexcept
{
    std::scoped_lock lock(g_debugHotPathsBrowseResultMutex);
    if (path.empty())
    {
        g_debugNextHotPathsBrowseResult.reset();
        return true;
    }

    g_debugNextHotPathsBrowseResult = DebugHotPathsBrowseResult{.kind = DebugHotPathsBrowseResultKind::Path, .path = std::filesystem::path(path)};
    return true;
}

bool DebugCancelPreferencesHotPathsNextBrowse() noexcept
{
    std::scoped_lock lock(g_debugHotPathsBrowseResultMutex);
    g_debugNextHotPathsBrowseResult = DebugHotPathsBrowseResult{.kind = DebugHotPathsBrowseResultKind::Cancel};
    return true;
}

PreferencesHotPathsDebugFocusTarget HotPathsPane::DebugGetFocusTarget() const noexcept
{
    if (! _pageHostDx || ! _dxState)
    {
        return PreferencesHotPathsDebugFocusTarget::None;
    }

    const auto* focused = _pageHostDx->GetFocusControl();
    if (! focused)
    {
        return PreferencesHotPathsDebugFocusTarget::None;
    }

    const auto& page = _dxState->page;
    if (page.slots[0].pathEdit == focused)
    {
        return PreferencesHotPathsDebugFocusTarget::FirstPathField;
    }
    if (page.slots[0].browseButton == focused)
    {
        return PreferencesHotPathsDebugFocusTarget::FirstBrowseButton;
    }
    if (page.slots[0].labelEdit == focused)
    {
        return PreferencesHotPathsDebugFocusTarget::FirstLabelField;
    }
    if (page.slots[0].showInMenuToggle == focused)
    {
        return PreferencesHotPathsDebugFocusTarget::FirstShowInMenuToggle;
    }
    if (page.slots[1].pathEdit == focused)
    {
        return PreferencesHotPathsDebugFocusTarget::SecondPathField;
    }
    if (page.openPrefsToggle == focused)
    {
        return PreferencesHotPathsDebugFocusTarget::OpenPrefsToggle;
    }

    return PreferencesHotPathsDebugFocusTarget::None;
}

bool HotPathsPane::DebugFocusFirstPathField() noexcept
{
    if (! _pageHostDx || ! _dxState)
    {
        return false;
    }

    auto* const pathField = _dxState->page.slots[0].pathEdit;
    if (! pathField || ! pathField->IsVisible() || ! pathField->IsEnabled())
    {
        return false;
    }

    _pageHostDx->SetFocusControl(pathField);
    return _pageHostDx->GetFocusControl() == pathField;
}

bool HotPathsPane::DebugGetFirstPathText(std::wstring& outText) const noexcept
{
    if (! _dxState)
    {
        return false;
    }

    const auto* const pathField = _dxState->page.slots[0].pathEdit;
    if (! pathField || ! pathField->IsVisible())
    {
        return false;
    }

    outText = pathField->GetText();
    return true;
}

bool HotPathsPane::DebugSetFirstPathText(std::wstring_view text) noexcept
{
    if (! _dxState)
    {
        return false;
    }

    auto* const pathField = _dxState->page.slots[0].pathEdit;
    if (! pathField || ! pathField->IsVisible() || ! pathField->IsEnabled())
    {
        return false;
    }

    pathField->SetText(std::wstring(text));
    if (_pageHostDx)
    {
        _pageHostDx->SyncTextInputBridge(pathField);
    }
    return true;
}

bool HotPathsPane::DebugFocusOpenPrefsToggle() noexcept
{
    if (! _pageHostDx || ! _dxState)
    {
        return false;
    }

    auto* const toggle = _dxState->page.openPrefsToggle;
    if (! toggle || ! toggle->IsVisible() || ! toggle->IsEnabled())
    {
        return false;
    }

    _pageHostDx->SetFocusControl(toggle);
    return _pageHostDx->GetFocusControl() == toggle;
}

bool HotPathsPane::DebugGetOpenPrefsToggleChecked(bool& outChecked) const noexcept
{
    if (! _dxState)
    {
        return false;
    }

    const auto* const toggle = _dxState->page.openPrefsToggle;
    if (! toggle || ! toggle->IsVisible())
    {
        return false;
    }

    outChecked = toggle->IsChecked();
    return true;
}
#endif

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
