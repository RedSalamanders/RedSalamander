// Preferences.FileOperations.cpp

#include "Framework.h"

#include "Preferences.FileOperations.h"

#include <algorithm>
#include <array>
#include <format>
#include <limits>
#include <string>
#include <string_view>
#include <vector>

#include "DxUi/DxUi.h"
#include "Helpers.h"
#include "UiMetrics.h"
#include "resource.h"

namespace
{
using PrefsFileOperations::EnsureWorkingFileOperationsSettings;
using PrefsFileOperations::GetFileOperationsSettingsOrDefault;
using PrefsFileOperations::MaybeResetWorkingFileOperationsSettingsIfEmpty;

using RedSalamander::DxUi::CardPanel;
using RedSalamander::DxUi::ComboBox;
using RedSalamander::DxUi::ComboBoxVariant;
using RedSalamander::DxUi::FontRole;
using RedSalamander::DxUi::Label;
using RedSalamander::DxUi::Panel;
using RedSalamander::DxUi::TextField;
using RedSalamander::DxUi::Toggle;

constexpr uint64_t kKiB                      = 1024ull;
constexpr uint64_t kMiB                      = 1024ull * 1024ull;
constexpr uint64_t kGiB                      = 1024ull * 1024ull * 1024ull;
constexpr size_t kBandwidthCustomPresetIndex = 7u;

struct BandwidthPreset
{
    uint64_t bytesPerSecond = 0;
    UINT labelId            = 0u;
};

constexpr std::array<BandwidthPreset, kBandwidthCustomPresetIndex> kBandwidthPresets = {{
    {0ull, IDS_PREFS_FILEOPS_BANDWIDTH_UNLIMITED},
    {1ull * kMiB, IDS_PREFS_FILEOPS_BANDWIDTH_1_MIB},
    {5ull * kMiB, IDS_PREFS_FILEOPS_BANDWIDTH_5_MIB},
    {10ull * kMiB, IDS_PREFS_FILEOPS_BANDWIDTH_10_MIB},
    {50ull * kMiB, IDS_PREFS_FILEOPS_BANDWIDTH_50_MIB},
    {100ull * kMiB, IDS_PREFS_FILEOPS_BANDWIDTH_100_MIB},
    {500ull * kMiB, IDS_PREFS_FILEOPS_BANDWIDTH_500_MIB},
}};

struct FileOperationsToggleCardDx
{
    CardPanel* card    = nullptr;
    Label* title       = nullptr;
    Label* description = nullptr;
    Toggle* toggle     = nullptr;
};

struct FileOperationsComboCardDx
{
    CardPanel* card    = nullptr;
    Label* title       = nullptr;
    Label* description = nullptr;
    ComboBox* combo    = nullptr;
};

struct FileOperationsEditCardDx
{
    CardPanel* card    = nullptr;
    Label* title       = nullptr;
    Label* description = nullptr;
    TextField* edit    = nullptr;
};

struct FileOperationsNoteCardDx
{
    CardPanel* card    = nullptr;
    Label* title       = nullptr;
    Label* description = nullptr;
};

struct FileOperationsDxPage
{
    Label* preCalcHeader = nullptr;
    FileOperationsToggleCardDx preCalcEnabled{};
    FileOperationsComboCardDx preCalcWorkers{};

    Label* bandwidthHeader = nullptr;
    FileOperationsComboCardDx bandwidthPreset{};
    FileOperationsEditCardDx customBandwidth{};

    Label* advancedHeader = nullptr;
    FileOperationsEditCardDx bridgeBuffer{};
    FileOperationsNoteCardDx pluginHint{};

    void Detach() noexcept
    {
        preCalcHeader   = nullptr;
        preCalcEnabled  = {};
        preCalcWorkers  = {};
        bandwidthHeader = nullptr;
        bandwidthPreset = {};
        customBandwidth = {};
        advancedHeader  = nullptr;
        bridgeBuffer    = {};
        pluginHint      = {};
    }
};

[[nodiscard]] std::wstring_view TrimAscii(std::wstring_view text) noexcept
{
    while (! text.empty() && text.front() <= L' ')
    {
        text.remove_prefix(1);
    }

    while (! text.empty() && text.back() <= L' ')
    {
        text.remove_suffix(1);
    }

    return text;
}

[[nodiscard]] bool EqualsIgnoreAsciiCase(std::wstring_view left, std::wstring_view right) noexcept
{
    if (left.size() != right.size())
    {
        return false;
    }

    for (size_t i = 0; i < left.size(); ++i)
    {
        const wchar_t lhs       = left[i];
        const wchar_t rhs       = right[i];
        const wchar_t lhsFolded = (lhs >= L'A' && lhs <= L'Z') ? static_cast<wchar_t>(lhs + (L'a' - L'A')) : lhs;
        const wchar_t rhsFolded = (rhs >= L'A' && rhs <= L'Z') ? static_cast<wchar_t>(rhs + (L'a' - L'A')) : rhs;
        if (lhsFolded != rhsFolded)
        {
            return false;
        }
    }

    return true;
}

[[nodiscard]] bool TryParseThroughputText(std::wstring_view text, uint64_t& outBytesPerSecond) noexcept
{
    constexpr uint64_t kTiB = 1024ull * 1024ull * 1024ull * 1024ull;
    constexpr uint64_t kPiB = 1024ull * 1024ull * 1024ull * 1024ull * 1024ull;

    outBytesPerSecond = 0;
    text              = TrimAscii(text);
    if (text.empty())
    {
        return true;
    }

    bool sawDigit          = false;
    bool sawDecimal        = false;
    double number          = 0.0;
    double fractionalScale = 0.1;
    size_t index           = 0;
    for (; index < text.size(); ++index)
    {
        const wchar_t ch = text[index];
        if (ch >= L'0' && ch <= L'9')
        {
            sawDigit                 = true;
            const unsigned int digit = static_cast<unsigned int>(ch - L'0');
            if (! sawDecimal)
            {
                number = (number * 10.0) + static_cast<double>(digit);
            }
            else
            {
                number += static_cast<double>(digit) * fractionalScale;
                fractionalScale *= 0.1;
            }
            continue;
        }

        if ((ch == L'.' || ch == L',') && ! sawDecimal)
        {
            sawDecimal = true;
            continue;
        }

        break;
    }

    if (! sawDigit)
    {
        return false;
    }

    std::wstring_view unit = TrimAscii(text.substr(index));
    if (unit.size() >= 2)
    {
        const wchar_t penultimate = unit[unit.size() - 2];
        const wchar_t last        = unit.back();
        if (penultimate == L'/' && (last == L's' || last == L'S'))
        {
            unit.remove_suffix(2);
            unit = TrimAscii(unit);
        }
    }

    uint64_t multiplier = 0;
    if (unit.empty() || EqualsIgnoreAsciiCase(unit, L"kb") || EqualsIgnoreAsciiCase(unit, L"k") || EqualsIgnoreAsciiCase(unit, L"kib"))
    {
        multiplier = kKiB;
    }
    else if (EqualsIgnoreAsciiCase(unit, L"b"))
    {
        multiplier = 1;
    }
    else if (EqualsIgnoreAsciiCase(unit, L"mb") || EqualsIgnoreAsciiCase(unit, L"m") || EqualsIgnoreAsciiCase(unit, L"mib"))
    {
        multiplier = kMiB;
    }
    else if (EqualsIgnoreAsciiCase(unit, L"gb") || EqualsIgnoreAsciiCase(unit, L"g") || EqualsIgnoreAsciiCase(unit, L"gib"))
    {
        multiplier = kGiB;
    }
    else if (EqualsIgnoreAsciiCase(unit, L"tb") || EqualsIgnoreAsciiCase(unit, L"t") || EqualsIgnoreAsciiCase(unit, L"tib"))
    {
        multiplier = kTiB;
    }
    else if (EqualsIgnoreAsciiCase(unit, L"pb") || EqualsIgnoreAsciiCase(unit, L"p") || EqualsIgnoreAsciiCase(unit, L"pib"))
    {
        multiplier = kPiB;
    }
    else
    {
        return false;
    }

    const double result = number * static_cast<double>(multiplier);
    if (result <= 0.0)
    {
        outBytesPerSecond = 0;
        return true;
    }

    constexpr double maxValue = static_cast<double>(std::numeric_limits<uint64_t>::max());
    if (result >= maxValue)
    {
        outBytesPerSecond = std::numeric_limits<uint64_t>::max();
        return true;
    }

    outBytesPerSecond = static_cast<uint64_t>(result + 0.5);
    return true;
}

[[nodiscard]] bool TryParseUnsignedDecimal(std::wstring_view text, uint32_t minimum, uint32_t maximum, uint32_t& outValue) noexcept
{
    text = TrimAscii(text);
    if (text.empty())
    {
        return false;
    }

    uint64_t value = 0;
    for (const wchar_t ch : text)
    {
        if (ch < L'0' || ch > L'9')
        {
            return false;
        }

        value = (value * 10u) + static_cast<uint64_t>(ch - L'0');
        if (value > maximum)
        {
            return false;
        }
    }

    if (value < minimum || value > maximum)
    {
        return false;
    }

    outValue = static_cast<uint32_t>(value);
    return true;
}

[[nodiscard]] size_t ResolveBandwidthPresetIndex(const uint64_t bytesPerSecond) noexcept
{
    for (size_t i = 0; i < kBandwidthPresets.size(); ++i)
    {
        if (kBandwidthPresets[i].bytesPerSecond == bytesPerSecond)
        {
            return i;
        }
    }

    return kBandwidthCustomPresetIndex;
}

[[nodiscard]] std::wstring FormatThroughputText(const uint64_t bytesPerSecond) noexcept
{
    if (bytesPerSecond == 0)
    {
        return {};
    }

    if ((bytesPerSecond % kGiB) == 0)
    {
        return std::format(L"{} GiB/s", bytesPerSecond / kGiB);
    }

    if ((bytesPerSecond % kMiB) == 0)
    {
        return std::format(L"{} MiB/s", bytesPerSecond / kMiB);
    }

    if ((bytesPerSecond % kKiB) == 0)
    {
        return std::format(L"{} KiB/s", bytesPerSecond / kKiB);
    }

    return std::format(L"{} B/s", bytesPerSecond);
}

} // namespace

struct FileOperationsPane::DxState
{
    DxState()                          = default;
    DxState(const DxState&)            = delete;
    DxState& operator=(const DxState&) = delete;
    DxState(DxState&&)                 = delete;
    DxState& operator=(DxState&&)      = delete;

    FileOperationsDxPage page;

    void Detach() noexcept
    {
        page.Detach();
    }
};

FileOperationsPane::FileOperationsPane() = default;

FileOperationsPane::~FileOperationsPane()
{
    DetachDxHosts();
}

void FileOperationsPane::OnVisibilityChanged(bool visible) noexcept
{
    static_cast<void>(visible);
}

void FileOperationsPane::Destroy(PreferencesDialogState& state) noexcept
{
    static_cast<void>(state);
    DetachDxHosts();
    _state    = nullptr;
    _pageHost = nullptr;
}

bool FileOperationsPane::EnsureDxHosts(HWND parent, PreferencesDialogState& state) noexcept
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

    dxState->page.preCalcHeader = root->AddChild<Label>();
    dxState->page.preCalcHeader->SetFontRole(FontRole::Header);

    dxState->page.preCalcEnabled.card  = root->AddChild<CardPanel>();
    dxState->page.preCalcEnabled.title = root->AddChild<Label>();
    dxState->page.preCalcEnabled.title->SetFontRole(FontRole::Body);
    dxState->page.preCalcEnabled.description = root->AddChild<Label>();
    dxState->page.preCalcEnabled.description->SetFontRole(FontRole::Small);
    dxState->page.preCalcEnabled.description->SetMultiline(true);
    dxState->page.preCalcEnabled.toggle = root->AddChild<Toggle>();
    dxState->page.preCalcEnabled.title->SetMnemonicTarget(dxState->page.preCalcEnabled.toggle);
    dxState->page.preCalcEnabled.toggle->SetStateLabels(LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF), LoadStringResource(nullptr, IDS_PREFS_COMMON_ON));
    dxState->page.preCalcEnabled.toggle->SetOnToggled([this, host = parent](bool checked) noexcept
    {
        if (! host || IsWindow(host) == FALSE)
        {
            return;
        }

        auto* dialogState = PrefsUi::GetDialogState(host);
        if (! dialogState)
        {
            return;
        }

        auto* fileOperations = EnsureWorkingFileOperationsSettings(dialogState->workingSettings);
        if (! fileOperations)
        {
            return;
        }

        if (fileOperations->preCalcEnabled != checked)
        {
            fileOperations->preCalcEnabled = checked;
            MaybeResetWorkingFileOperationsSettingsIfEmpty(dialogState->workingSettings);
            SetDirty(GetParent(host), *dialogState);
        }

        Refresh(host, *dialogState);
    });

    dxState->page.preCalcWorkers.card  = root->AddChild<CardPanel>();
    dxState->page.preCalcWorkers.title = root->AddChild<Label>();
    dxState->page.preCalcWorkers.title->SetFontRole(FontRole::Body);
    dxState->page.preCalcWorkers.description = root->AddChild<Label>();
    dxState->page.preCalcWorkers.description->SetFontRole(FontRole::Small);
    dxState->page.preCalcWorkers.description->SetMultiline(true);
    dxState->page.preCalcWorkers.combo = root->AddChild<ComboBox>();
    dxState->page.preCalcWorkers.title->SetMnemonicTarget(dxState->page.preCalcWorkers.combo);
    dxState->page.preCalcWorkers.combo->SetVariant(ComboBoxVariant::Window);
    {
        std::vector<ComboBox::Item> items;
        items.reserve(8u);
        for (uint32_t value = 1; value <= 8; ++value)
        {
            const std::wstring label = std::to_wstring(value);
            items.push_back({label, label});
        }
        dxState->page.preCalcWorkers.combo->SetItems(std::move(items));
    }
    dxState->page.preCalcWorkers.combo->SetOnSelectionChanged([this, host = parent](size_t itemIndex) noexcept
    {
        if (_syncingDxPreCalcWorkersCombo || ! host || IsWindow(host) == FALSE)
        {
            return;
        }

        auto* dialogState = PrefsUi::GetDialogState(host);
        if (! dialogState)
        {
            return;
        }

        auto* fileOperations = EnsureWorkingFileOperationsSettings(dialogState->workingSettings);
        if (! fileOperations)
        {
            return;
        }

        const uint32_t newValue = std::clamp<uint32_t>(static_cast<uint32_t>(itemIndex) + 1u, 1u, 8u);
        if (fileOperations->preCalcMaxWorkers != newValue)
        {
            fileOperations->preCalcMaxWorkers = newValue;
            MaybeResetWorkingFileOperationsSettingsIfEmpty(dialogState->workingSettings);
            SetDirty(GetParent(host), *dialogState);
        }
    });

    dxState->page.bandwidthHeader = root->AddChild<Label>();
    dxState->page.bandwidthHeader->SetFontRole(FontRole::Header);

    dxState->page.bandwidthPreset.card  = root->AddChild<CardPanel>();
    dxState->page.bandwidthPreset.title = root->AddChild<Label>();
    dxState->page.bandwidthPreset.title->SetFontRole(FontRole::Body);
    dxState->page.bandwidthPreset.description = root->AddChild<Label>();
    dxState->page.bandwidthPreset.description->SetFontRole(FontRole::Small);
    dxState->page.bandwidthPreset.description->SetMultiline(true);
    dxState->page.bandwidthPreset.combo = root->AddChild<ComboBox>();
    dxState->page.bandwidthPreset.title->SetMnemonicTarget(dxState->page.bandwidthPreset.combo);
    dxState->page.bandwidthPreset.combo->SetVariant(ComboBoxVariant::Window);
    {
        std::vector<ComboBox::Item> items;
        items.reserve(kBandwidthPresets.size() + 1u);
        for (const BandwidthPreset& preset : kBandwidthPresets)
        {
            items.push_back({std::to_wstring(preset.bytesPerSecond), LoadStringResource(nullptr, preset.labelId)});
        }
        items.push_back({L"custom", LoadStringResource(nullptr, IDS_PREFS_FILEOPS_BANDWIDTH_CUSTOM)});
        dxState->page.bandwidthPreset.combo->SetItems(std::move(items));
    }
    dxState->page.bandwidthPreset.combo->SetOnSelectionChanged([this, host = parent](size_t itemIndex) noexcept
    {
        if (_syncingDxBandwidthPresetCombo || ! host || IsWindow(host) == FALSE)
        {
            return;
        }

        auto* dialogState = PrefsUi::GetDialogState(host);
        if (! dialogState)
        {
            return;
        }

        const bool wantsCustom   = (itemIndex >= kBandwidthPresets.size());
        const bool needsRelayout = (wantsCustom != _showCustomBandwidth);
        _showCustomBandwidth     = wantsCustom;

        if (wantsCustom)
        {
            if (needsRelayout)
            {
                static_cast<void>(PrefsUi::PostDeferredAction(host, PreferencesDeferredActionKind::FileOperationsBandwidthPresetChanged));
            }
            else if (_pageHostDx && _dxState && _dxState->page.customBandwidth.edit)
            {
                _pageHostDx->SetFocusControl(_dxState->page.customBandwidth.edit);
                _pageHostDx->Invalidate();
            }
            return;
        }

        auto* fileOperations = EnsureWorkingFileOperationsSettings(dialogState->workingSettings);
        if (! fileOperations)
        {
            return;
        }

        const uint64_t newValue = kBandwidthPresets[itemIndex].bytesPerSecond;
        if (fileOperations->defaultBandwidthLimitBytesPerSecond != newValue)
        {
            fileOperations->defaultBandwidthLimitBytesPerSecond = newValue;
            MaybeResetWorkingFileOperationsSettingsIfEmpty(dialogState->workingSettings);
            SetDirty(GetParent(host), *dialogState);
        }

        if (needsRelayout)
        {
            static_cast<void>(PrefsUi::PostDeferredAction(host, PreferencesDeferredActionKind::FileOperationsBandwidthPresetChanged));
        }
        else
        {
            Refresh(host, *dialogState);
        }
    });

    dxState->page.customBandwidth.card  = root->AddChild<CardPanel>();
    dxState->page.customBandwidth.title = root->AddChild<Label>();
    dxState->page.customBandwidth.title->SetFontRole(FontRole::Body);
    dxState->page.customBandwidth.description = root->AddChild<Label>();
    dxState->page.customBandwidth.description->SetFontRole(FontRole::Small);
    dxState->page.customBandwidth.description->SetMultiline(true);
    dxState->page.customBandwidth.edit = root->AddChild<TextField>();
    dxState->page.customBandwidth.title->SetMnemonicTarget(dxState->page.customBandwidth.edit);
    dxState->page.customBandwidth.edit->SetOnTextChanged([this, host = parent](std::wstring_view text) noexcept
    {
        if (_syncingDxCustomBandwidthEdit || ! host || IsWindow(host) == FALSE || ! _state)
        {
            return;
        }

        uint64_t parsedValue = 0;
        if (! TryParseThroughputText(text, parsedValue))
        {
            return;
        }

        auto* fileOperations = EnsureWorkingFileOperationsSettings(_state->workingSettings);
        if (! fileOperations)
        {
            return;
        }

        if (fileOperations->defaultBandwidthLimitBytesPerSecond != parsedValue)
        {
            fileOperations->defaultBandwidthLimitBytesPerSecond = parsedValue;
            MaybeResetWorkingFileOperationsSettingsIfEmpty(_state->workingSettings);
            SetDirty(GetParent(host), *_state);
        }
    });
    dxState->page.customBandwidth.edit->SetOnBlur([this, host = parent]() noexcept
    {
        if (! _state || ! host || IsWindow(host) == FALSE)
        {
            return;
        }

        Refresh(host, *_state);
    });

    dxState->page.advancedHeader = root->AddChild<Label>();
    dxState->page.advancedHeader->SetFontRole(FontRole::Header);

    dxState->page.bridgeBuffer.card  = root->AddChild<CardPanel>();
    dxState->page.bridgeBuffer.title = root->AddChild<Label>();
    dxState->page.bridgeBuffer.title->SetFontRole(FontRole::Body);
    dxState->page.bridgeBuffer.description = root->AddChild<Label>();
    dxState->page.bridgeBuffer.description->SetFontRole(FontRole::Small);
    dxState->page.bridgeBuffer.description->SetMultiline(true);
    dxState->page.bridgeBuffer.edit = root->AddChild<TextField>();
    dxState->page.bridgeBuffer.title->SetMnemonicTarget(dxState->page.bridgeBuffer.edit);
    dxState->page.bridgeBuffer.edit->SetOnTextChanged([this, host = parent](std::wstring_view text) noexcept
    {
        if (_syncingDxBridgeBufferEdit || ! host || IsWindow(host) == FALSE || ! _state)
        {
            return;
        }

        uint32_t parsedValue = 0;
        if (! TryParseUnsignedDecimal(text, 512u, 16384u, parsedValue))
        {
            return;
        }

        auto* fileOperations = EnsureWorkingFileOperationsSettings(_state->workingSettings);
        if (! fileOperations)
        {
            return;
        }

        if (fileOperations->crossFsBridgeBufferSizeKB != parsedValue)
        {
            fileOperations->crossFsBridgeBufferSizeKB = parsedValue;
            MaybeResetWorkingFileOperationsSettingsIfEmpty(_state->workingSettings);
            SetDirty(GetParent(host), *_state);
        }
    });
    dxState->page.bridgeBuffer.edit->SetOnBlur([this, host = parent]() noexcept
    {
        if (! _state || ! host || IsWindow(host) == FALSE)
        {
            return;
        }

        Refresh(host, *_state);
    });

    dxState->page.pluginHint.card  = root->AddChild<CardPanel>();
    dxState->page.pluginHint.title = root->AddChild<Label>();
    dxState->page.pluginHint.title->SetFontRole(FontRole::Body);
    dxState->page.pluginHint.description = root->AddChild<Label>();
    dxState->page.pluginHint.description->SetFontRole(FontRole::Small);
    dxState->page.pluginHint.description->SetMultiline(true);

    _dxState = std::move(dxState);
    ApplyDxTheme(state);
    SyncDxControlsFromState(state);
    return true;
}

void FileOperationsPane::DetachDxHosts() noexcept
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

void FileOperationsPane::ApplyDxTheme(const PreferencesDialogState& state) noexcept
{
    if (! _dxState || ! _pageHostDx)
    {
        return;
    }

    _pageHostDx->SetTheme(PrefsUi::MakeDxPalette(state.theme));
}

void FileOperationsPane::SyncDxControlsFromState(const PreferencesDialogState& state) noexcept
{
    if (! _dxState)
    {
        return;
    }

    const auto& fileOperations = GetFileOperationsSettingsOrDefault(state.workingSettings);
    FileOperationsDxPage& page = _dxState->page;

    page.preCalcHeader->SetText(LoadStringResource(nullptr, IDS_PREFS_FILEOPS_SECTION_PRECALC));
    page.preCalcEnabled.title->SetText(LoadStringResource(nullptr, IDS_PREFS_FILEOPS_PRECALC_ENABLED_TITLE));
    page.preCalcEnabled.description->SetText(LoadStringResource(nullptr, IDS_PREFS_FILEOPS_PRECALC_ENABLED_DESC));
    page.preCalcEnabled.toggle->SetChecked(fileOperations.preCalcEnabled);
    page.preCalcEnabled.toggle->SetEnabled(true);

    page.preCalcWorkers.title->SetText(LoadStringResource(nullptr, IDS_PREFS_FILEOPS_PRECALC_WORKERS_TITLE));
    page.preCalcWorkers.description->SetText(LoadStringResource(nullptr, IDS_PREFS_FILEOPS_PRECALC_WORKERS_DESC));
    _syncingDxPreCalcWorkersCombo = true;
    page.preCalcWorkers.combo->SetSelectedIndex(std::clamp<size_t>(static_cast<size_t>(std::max<uint32_t>(1u, fileOperations.preCalcMaxWorkers) - 1u), 0u, 7u));
    page.preCalcWorkers.combo->SetEnabled(fileOperations.preCalcEnabled);
    _syncingDxPreCalcWorkersCombo = false;

    page.bandwidthHeader->SetText(LoadStringResource(nullptr, IDS_PREFS_FILEOPS_SECTION_BANDWIDTH));
    page.bandwidthPreset.title->SetText(LoadStringResource(nullptr, IDS_PREFS_FILEOPS_BANDWIDTH_PRESET_TITLE));
    page.bandwidthPreset.description->SetText(LoadStringResource(nullptr, IDS_PREFS_FILEOPS_BANDWIDTH_PRESET_DESC));
    _syncingDxBandwidthPresetCombo = true;
    const size_t savedPresetIndex  = ResolveBandwidthPresetIndex(fileOperations.defaultBandwidthLimitBytesPerSecond);
    // When already in custom mode (user picked Custom but hasn't saved a value yet),
    // keep the combo on Custom rather than reverting to the previous preset.
    const size_t bandwidthPresetIndex = _showCustomBandwidth ? kBandwidthCustomPresetIndex : savedPresetIndex;
    page.bandwidthPreset.combo->SetSelectedIndex(bandwidthPresetIndex);
    page.bandwidthPreset.combo->SetEnabled(true);
    _syncingDxBandwidthPresetCombo = false;
    _showCustomBandwidth           = (bandwidthPresetIndex == kBandwidthCustomPresetIndex);

    page.customBandwidth.title->SetText(LoadStringResource(nullptr, IDS_PREFS_FILEOPS_BANDWIDTH_CUSTOM_TITLE));
    page.customBandwidth.description->SetText(LoadStringResource(nullptr, IDS_PREFS_FILEOPS_BANDWIDTH_CUSTOM_DESC));
    _syncingDxCustomBandwidthEdit = true;
    page.customBandwidth.edit->SetText(FormatThroughputText(fileOperations.defaultBandwidthLimitBytesPerSecond));
    page.customBandwidth.edit->SetEnabled(_showCustomBandwidth);
    _syncingDxCustomBandwidthEdit = false;

    page.advancedHeader->SetText(LoadStringResource(nullptr, IDS_PREFS_FILEOPS_SECTION_ADVANCED));
    page.bridgeBuffer.title->SetText(LoadStringResource(nullptr, IDS_PREFS_FILEOPS_BRIDGE_BUFFER_TITLE));
    page.bridgeBuffer.description->SetText(LoadStringResource(nullptr, IDS_PREFS_FILEOPS_BRIDGE_BUFFER_DESC));
    _syncingDxBridgeBufferEdit = true;
    page.bridgeBuffer.edit->SetText(std::to_wstring(fileOperations.crossFsBridgeBufferSizeKB));
    page.bridgeBuffer.edit->SetEnabled(true);
    _syncingDxBridgeBufferEdit = false;

    page.pluginHint.title->SetText(LoadStringResource(nullptr, IDS_PREFS_FILEOPS_PLUGIN_HINT_TITLE));
    page.pluginHint.description->SetText(LoadStringResource(nullptr, IDS_PREFS_FILEOPS_PLUGIN_HINT_DESC));

    if (_pageHostDx)
    {
        _pageHostDx->Invalidate();
    }
}

void FileOperationsPane::InitializePage(HWND parent, PreferencesDialogState& state) noexcept
{
    if (! parent)
    {
        return;
    }

    _pageHost = parent;
    _state    = &state;

    if (state.currentCategory == PrefCategory::FileOperations)
    {
        if (! EnsureDxHosts(parent, state))
        {
            Debug::Error(L"Preferences.FileOperations: DxUi surface initialization failed; page will not render correctly.");
            DetachDxHosts();
        }
    }

    Refresh(parent, state);
}

void FileOperationsPane::Refresh(HWND host, PreferencesDialogState& state) noexcept
{
    if (state.currentCategory == PrefCategory::FileOperations && ! _dxState)
    {
        const HWND parent = _pageHost;
        if (parent)
        {
            static_cast<void>(EnsureDxHosts(parent, state));
        }
    }

    ApplyDxTheme(state);
    SyncDxControlsFromState(state);
    if (host)
    {
        InvalidateRect(host, nullptr, FALSE);
    }
}

void FileOperationsPane::LayoutDxPage(
    HWND host, PreferencesDialogState& state, int x, int& y, int width, int margin, int gapY, const PreferencesTypographyContext& typography) noexcept
{
    using namespace PrefsLayoutConstants;

    static_cast<void>(host);
    static_cast<void>(margin);

    if (! _dxState || ! _pageHostDx || ! _pageContentRoot)
    {
        return;
    }

    Debug::Perf::Scope layoutPerf(L"preferences.ui.file_operations_layout_us");
    layoutPerf.SetValue0(static_cast<uint64_t>(std::max(0, width)));
    layoutPerf.SetValue1(static_cast<uint64_t>(std::max<UINT>(typography.dpi, USER_DEFAULT_SCREEN_DPI)));

    const UINT dpi         = std::max<UINT>(typography.dpi, USER_DEFAULT_SCREEN_DPI);
    const int rowHeight    = std::max(1, UiMetrics::ScaleDip(dpi, kRowHeightDip));
    const int titleHeight  = std::max(1, UiMetrics::ScaleDip(dpi, kTitleHeightDip));
    const int headerHeight = std::max(1, UiMetrics::ScaleDip(dpi, kHeaderHeightDip));
    const int editHeight   = std::max(1, UiMetrics::ScaleDip(dpi, kEditHeightDip));
    const int comboHeight  = std::max(1, UiMetrics::ScaleDip(dpi, kComboHeightDip));

    const int cardPaddingX = UiMetrics::ScaleDip(dpi, kCardPaddingXDip);
    const int cardPaddingY = UiMetrics::ScaleDip(dpi, kCardPaddingYDip);
    const int cardGapY     = UiMetrics::ScaleDip(dpi, kCardGapYDip);
    const int cardGapX     = UiMetrics::ScaleDip(dpi, kCardGapXDip);
    const int cardSpacingY = UiMetrics::ScaleDip(dpi, kCardSpacingYDip);

    const int minToggleWidth       = UiMetrics::ScaleDip(dpi, kMinToggleWidthDip);
    const std::wstring onLabel     = LoadStringResource(nullptr, IDS_PREFS_COMMON_ON);
    const std::wstring offLabel    = LoadStringResource(nullptr, IDS_PREFS_COMMON_OFF);
    const int onWidth              = PrefsUi::MeasureSingleLineTextWidthPx(typography, typography.strong, onLabel);
    const int offWidth             = PrefsUi::MeasureSingleLineTextWidthPx(typography, typography.strong, offLabel);
    const int togglePaddingX       = UiMetrics::ScaleDip(dpi, kTogglePaddingXDip);
    const int toggleGapX           = UiMetrics::ScaleDip(dpi, kToggleGapXDip);
    const int toggleTrackWidth     = UiMetrics::ScaleDip(dpi, kToggleTrackWidthDip);
    const int toggleStateTextWidth = std::max(onWidth, offWidth);
    const int measuredToggleWidth  = std::max(minToggleWidth, (2 * togglePaddingX) + toggleStateTextWidth + toggleGapX + toggleTrackWidth);
    const int toggleWidth          = std::min(std::max(0, width - 2 * cardPaddingX - cardGapX), measuredToggleWidth);
    const int comboWidth           = std::min(std::max(0, width - 2 * cardPaddingX), UiMetrics::ScaleDip(dpi, kLargeComboWidthDip));
    const int editWidth            = std::min(std::max(0, width - 2 * cardPaddingX), UiMetrics::ScaleDip(dpi, kMaxEditWidthDip));

    const auto pxToDip = [dpi](const int px) noexcept { return static_cast<float>(px) * 96.0f / static_cast<float>(std::max<UINT>(1u, dpi)); };

    FileOperationsDxPage& page = _dxState->page;
    page.preCalcHeader->SetVisible(false);
    page.bandwidthHeader->SetVisible(false);
    page.advancedHeader->SetVisible(false);

    const auto hideToggleCard = [](FileOperationsToggleCardDx& card) noexcept
    {
        if (card.card)
        {
            card.card->SetVisible(false);
        }
        if (card.title)
        {
            card.title->SetVisible(false);
        }
        if (card.description)
        {
            card.description->SetVisible(false);
        }
        if (card.toggle)
        {
            card.toggle->SetVisible(false);
        }
    };
    const auto hideComboCard = [](FileOperationsComboCardDx& card) noexcept
    {
        if (card.card)
        {
            card.card->SetVisible(false);
        }
        if (card.title)
        {
            card.title->SetVisible(false);
        }
        if (card.description)
        {
            card.description->SetVisible(false);
        }
        if (card.combo)
        {
            card.combo->SetVisible(false);
        }
    };
    const auto hideEditCard = [](FileOperationsEditCardDx& card) noexcept
    {
        if (card.card)
        {
            card.card->SetVisible(false);
        }
        if (card.title)
        {
            card.title->SetVisible(false);
        }
        if (card.description)
        {
            card.description->SetVisible(false);
        }
        if (card.edit)
        {
            card.edit->SetVisible(false);
        }
    };
    const auto hideNoteCard = [](FileOperationsNoteCardDx& card) noexcept
    {
        if (card.card)
        {
            card.card->SetVisible(false);
        }
        if (card.title)
        {
            card.title->SetVisible(false);
        }
        if (card.description)
        {
            card.description->SetVisible(false);
        }
    };

    hideToggleCard(page.preCalcEnabled);
    hideComboCard(page.preCalcWorkers);
    hideComboCard(page.bandwidthPreset);
    hideEditCard(page.customBandwidth);
    hideEditCard(page.bridgeBuffer);
    hideNoteCard(page.pluginHint);

    const auto pushCard = [&](const RECT& card) noexcept { state.pageSettingCards.push_back(card); };

    const auto layoutHeader = [&](Label* header) noexcept
    {
        if (! header)
        {
            return;
        }

        header->SetVisible(true);
        header->SetBounds(D2D1::RectF(pxToDip(x), pxToDip(y), pxToDip(x + width), pxToDip(y + headerHeight)));
        y += headerHeight + std::max(2, gapY / 2);
    };

    const auto layoutToggleCard = [&](FileOperationsToggleCardDx& card) noexcept
    {
        const int textWidth  = std::max(0, width - 2 * cardPaddingX - cardGapX - toggleWidth);
        const int descHeight = PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, textWidth, card.description->GetText());
        const int cardHeight = std::max(rowHeight + 2 * cardPaddingY, titleHeight + cardGapY + descHeight + 2 * cardPaddingY);
        const RECT cardRect{x, y, x + width, y + cardHeight};
        pushCard(cardRect);

        card.card->SetVisible(true);
        card.title->SetVisible(true);
        card.description->SetVisible(true);
        card.toggle->SetVisible(true);

        card.card->SetBounds(D2D1::RectF(pxToDip(cardRect.left), pxToDip(cardRect.top), pxToDip(cardRect.right), pxToDip(cardRect.bottom)));
        card.title->SetBounds(
            D2D1::RectF(pxToDip(x + cardPaddingX), pxToDip(y + cardPaddingY), pxToDip(x + cardPaddingX + textWidth), pxToDip(y + cardPaddingY + titleHeight)));
        card.description->SetBounds(D2D1::RectF(pxToDip(x + cardPaddingX),
                                                pxToDip(y + cardPaddingY + titleHeight + cardGapY),
                                                pxToDip(x + cardPaddingX + textWidth),
                                                pxToDip(y + cardPaddingY + titleHeight + cardGapY + descHeight)));
        const int toggleX = x + width - cardPaddingX - toggleWidth;
        const int toggleY = y + (cardHeight - rowHeight) / 2;
        card.toggle->SetBounds(D2D1::RectF(pxToDip(toggleX), pxToDip(toggleY), pxToDip(toggleX + toggleWidth), pxToDip(toggleY + rowHeight)));
        y += cardHeight + cardSpacingY;
    };

    const auto layoutComboCard = [&](FileOperationsComboCardDx& card) noexcept
    {
        const int textWidth  = std::max(0, width - 2 * cardPaddingX - cardGapX - comboWidth);
        const int descHeight = PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, textWidth, card.description->GetText());
        const int cardHeight = std::max(comboHeight + 2 * cardPaddingY, titleHeight + cardGapY + descHeight + 2 * cardPaddingY);
        const RECT cardRect{x, y, x + width, y + cardHeight};
        pushCard(cardRect);

        card.card->SetVisible(true);
        card.title->SetVisible(true);
        card.description->SetVisible(true);
        card.combo->SetVisible(true);

        card.card->SetBounds(D2D1::RectF(pxToDip(cardRect.left), pxToDip(cardRect.top), pxToDip(cardRect.right), pxToDip(cardRect.bottom)));
        card.title->SetBounds(
            D2D1::RectF(pxToDip(x + cardPaddingX), pxToDip(y + cardPaddingY), pxToDip(x + cardPaddingX + textWidth), pxToDip(y + cardPaddingY + titleHeight)));
        card.description->SetBounds(D2D1::RectF(pxToDip(x + cardPaddingX),
                                                pxToDip(y + cardPaddingY + titleHeight + cardGapY),
                                                pxToDip(x + cardPaddingX + textWidth),
                                                pxToDip(y + cardPaddingY + titleHeight + cardGapY + descHeight)));
        const int comboX = x + width - cardPaddingX - comboWidth;
        const int comboY = y + (cardHeight - comboHeight) / 2;
        card.combo->SetBounds(D2D1::RectF(pxToDip(comboX), pxToDip(comboY), pxToDip(comboX + comboWidth), pxToDip(comboY + comboHeight)));
        y += cardHeight + cardSpacingY;
    };

    const auto layoutEditCard = [&](FileOperationsEditCardDx& card) noexcept
    {
        const int textWidth  = std::max(0, width - 2 * cardPaddingX - cardGapX - editWidth);
        const int descHeight = PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, textWidth, card.description->GetText());
        const int cardHeight = std::max(editHeight + 2 * cardPaddingY, titleHeight + cardGapY + descHeight + 2 * cardPaddingY);
        const RECT cardRect{x, y, x + width, y + cardHeight};
        pushCard(cardRect);

        card.card->SetVisible(true);
        card.title->SetVisible(true);
        card.description->SetVisible(true);
        card.edit->SetVisible(true);

        card.card->SetBounds(D2D1::RectF(pxToDip(cardRect.left), pxToDip(cardRect.top), pxToDip(cardRect.right), pxToDip(cardRect.bottom)));
        card.title->SetBounds(
            D2D1::RectF(pxToDip(x + cardPaddingX), pxToDip(y + cardPaddingY), pxToDip(x + cardPaddingX + textWidth), pxToDip(y + cardPaddingY + titleHeight)));
        card.description->SetBounds(D2D1::RectF(pxToDip(x + cardPaddingX),
                                                pxToDip(y + cardPaddingY + titleHeight + cardGapY),
                                                pxToDip(x + cardPaddingX + textWidth),
                                                pxToDip(y + cardPaddingY + titleHeight + cardGapY + descHeight)));
        const int editX = x + width - cardPaddingX - editWidth;
        const int editY = y + (cardHeight - editHeight) / 2;
        card.edit->SetBounds(D2D1::RectF(pxToDip(editX), pxToDip(editY), pxToDip(editX + editWidth), pxToDip(editY + editHeight)));
        y += cardHeight + cardSpacingY;
    };

    const auto layoutNoteCard = [&](FileOperationsNoteCardDx& card) noexcept
    {
        const int textWidth  = std::max(0, width - 2 * cardPaddingX);
        const int descHeight = PrefsUi::MeasureWrappedTextHeightPx(typography, typography.caption, textWidth, card.description->GetText());
        const int cardHeight = std::max(titleHeight + cardGapY + descHeight + 2 * cardPaddingY, rowHeight + 2 * cardPaddingY);
        const RECT cardRect{x, y, x + width, y + cardHeight};
        pushCard(cardRect);

        card.card->SetVisible(true);
        card.title->SetVisible(true);
        card.description->SetVisible(true);

        card.card->SetBounds(D2D1::RectF(pxToDip(cardRect.left), pxToDip(cardRect.top), pxToDip(cardRect.right), pxToDip(cardRect.bottom)));
        card.title->SetBounds(
            D2D1::RectF(pxToDip(x + cardPaddingX), pxToDip(y + cardPaddingY), pxToDip(x + cardPaddingX + textWidth), pxToDip(y + cardPaddingY + titleHeight)));
        card.description->SetBounds(D2D1::RectF(pxToDip(x + cardPaddingX),
                                                pxToDip(y + cardPaddingY + titleHeight + cardGapY),
                                                pxToDip(x + cardPaddingX + textWidth),
                                                pxToDip(y + cardPaddingY + titleHeight + cardGapY + descHeight)));
        y += cardHeight + cardSpacingY;
    };

    layoutHeader(page.preCalcHeader);
    layoutToggleCard(page.preCalcEnabled);
    layoutComboCard(page.preCalcWorkers);

    layoutHeader(page.bandwidthHeader);
    layoutComboCard(page.bandwidthPreset);
    if (_showCustomBandwidth)
    {
        layoutEditCard(page.customBandwidth);
    }
    else
    {
        hideEditCard(page.customBandwidth);
    }

    layoutHeader(page.advancedHeader);
    layoutEditCard(page.bridgeBuffer);
    layoutNoteCard(page.pluginHint);

    _pageHostDx->Invalidate();
}

void FileOperationsPane::LayoutPage(
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

    Debug::Error(L"Preferences.FileOperations: DxUi surface initialization failed; page will not render correctly.");
}

bool FileOperationsPane::HandleDeferredAction(HWND host, [[maybe_unused]] PreferencesDialogState& state, PreferencesDeferredActionKind action) noexcept
{
    if (action != PreferencesDeferredActionKind::FileOperationsBandwidthPresetChanged)
    {
        return false;
    }

    if (! host)
    {
        return false;
    }

    // Layout is performed by the caller (HandleDeferredPaneAction in Preferences.Dialog.cpp)
    // via LayoutPreferencesPageHost so scroll, background cards, and DxUi are all updated
    // through the proper channel.  We only handle post-layout focus here.
    return true;
}

void FileOperationsPane::PostLayoutFocusCustomBandwidthEdit() noexcept
{
    // Move keyboard focus into the custom edit after the caller's layout pass completes.
    if (_showCustomBandwidth && _pageHostDx && _dxState && _dxState->page.customBandwidth.edit && _dxState->page.customBandwidth.edit->IsVisible())
    {
        _pageHostDx->SetFocusControl(_dxState->page.customBandwidth.edit);
        _pageHostDx->Invalidate();
    }
}

#ifdef ENABLE_TESTS
PreferencesFileOperationsDebugFocusTarget FileOperationsPane::DebugGetFocusTarget() const noexcept
{
    if (! _pageHostDx || ! _dxState)
    {
        return PreferencesFileOperationsDebugFocusTarget::None;
    }

    const auto* focused = _pageHostDx->GetFocusControl();
    if (! focused)
    {
        return PreferencesFileOperationsDebugFocusTarget::None;
    }

    const auto& page = _dxState->page;
    if (page.preCalcEnabled.toggle == focused)
    {
        return PreferencesFileOperationsDebugFocusTarget::PreCalcEnabledToggle;
    }
    if (page.preCalcWorkers.combo == focused)
    {
        return PreferencesFileOperationsDebugFocusTarget::PreCalcWorkersCombo;
    }
    if (page.bandwidthPreset.combo == focused)
    {
        return PreferencesFileOperationsDebugFocusTarget::BandwidthPresetCombo;
    }
    if (page.customBandwidth.edit == focused && page.customBandwidth.edit && page.customBandwidth.edit->IsVisible())
    {
        return PreferencesFileOperationsDebugFocusTarget::CustomBandwidthEdit;
    }
    if (page.bridgeBuffer.edit == focused)
    {
        return PreferencesFileOperationsDebugFocusTarget::BridgeBufferEdit;
    }

    return PreferencesFileOperationsDebugFocusTarget::None;
}

bool FileOperationsPane::DebugFocusPreCalcEnabledToggle() noexcept
{
    if (! _pageHostDx || ! _dxState)
    {
        return false;
    }

    auto* const toggle = _dxState->page.preCalcEnabled.toggle;
    if (! toggle || ! toggle->IsVisible() || ! toggle->IsEnabled())
    {
        return false;
    }

    _pageHostDx->SetFocusControl(toggle);
    return _pageHostDx->GetFocusControl() == toggle;
}

bool FileOperationsPane::DebugGetPreCalcEnabledToggleChecked(bool& outChecked) const noexcept
{
    if (! _dxState)
    {
        return false;
    }

    const auto* const toggle = _dxState->page.preCalcEnabled.toggle;
    if (! toggle || ! toggle->IsVisible())
    {
        return false;
    }

    outChecked = toggle->IsChecked();
    return true;
}

bool FileOperationsPane::DebugSelectBandwidthPresetByText(std::wstring_view displayText) noexcept
{
    if (! _pageHostDx || ! _dxState || ! _state || ! _pageHost || IsWindow(_pageHost) == FALSE)
    {
        return false;
    }

    auto* const combo = _dxState->page.bandwidthPreset.combo;
    if (! combo || ! combo->IsVisible() || ! combo->IsEnabled())
    {
        return false;
    }

    const auto items = combo->GetItems();
    const auto it    = std::find_if(items.begin(), items.end(), [displayText](const ComboBox::Item& item) noexcept { return item.display == displayText; });
    if (it == items.end())
    {
        return false;
    }

    const size_t itemIndex   = static_cast<size_t>(std::distance(items.begin(), it));
    const bool wantsCustom   = itemIndex >= kBandwidthPresets.size();
    const bool needsRelayout = wantsCustom != _showCustomBandwidth;

    _showCustomBandwidth = wantsCustom;
    _pageHostDx->SetFocusControl(combo);
    combo->SetSelectedIndex(itemIndex);

    if (wantsCustom)
    {
        if (! needsRelayout && _dxState->page.customBandwidth.edit)
        {
            _pageHostDx->SetFocusControl(_dxState->page.customBandwidth.edit);
        }

        _pageHostDx->Invalidate();
        return combo->GetSelectedIndex().has_value() && combo->GetSelectedIndex().value() == itemIndex && combo->GetDisplayedText() == displayText;
    }

    auto* fileOperations = EnsureWorkingFileOperationsSettings(_state->workingSettings);
    if (! fileOperations)
    {
        return false;
    }

    const uint64_t newValue = kBandwidthPresets[itemIndex].bytesPerSecond;
    if (fileOperations->defaultBandwidthLimitBytesPerSecond != newValue)
    {
        fileOperations->defaultBandwidthLimitBytesPerSecond = newValue;
        MaybeResetWorkingFileOperationsSettingsIfEmpty(_state->workingSettings);
        SetDirty(GetParent(_pageHost), *_state);
    }

    if (! needsRelayout)
    {
        Refresh(_pageHost, *_state);
    }
    else
    {
        _pageHostDx->Invalidate();
    }

    return combo->GetSelectedIndex().has_value() && combo->GetSelectedIndex().value() == itemIndex && combo->GetDisplayedText() == displayText;
}

bool FileOperationsPane::DebugSetBridgeBufferText(std::wstring_view text) noexcept
{
    if (! _pageHostDx || ! _dxState || ! _state)
    {
        return false;
    }

    auto* const edit = _dxState->page.bridgeBuffer.edit;
    if (! edit || ! edit->IsVisible() || ! edit->IsEnabled() || edit->IsReadOnly())
    {
        return false;
    }

    uint32_t parsedValue = 0;
    if (! TryParseUnsignedDecimal(text, 512u, 16384u, parsedValue))
    {
        return false;
    }

    auto* fileOperations = EnsureWorkingFileOperationsSettings(_state->workingSettings);
    if (! fileOperations)
    {
        return false;
    }

    if (fileOperations->crossFsBridgeBufferSizeKB != parsedValue)
    {
        fileOperations->crossFsBridgeBufferSizeKB = parsedValue;
        MaybeResetWorkingFileOperationsSettingsIfEmpty(_state->workingSettings);
        SetDirty(GetParent(_pageHostDx->GetHwnd()), *_state);
    }

    _syncingDxBridgeBufferEdit = true;
    _pageHostDx->SetFocusControl(edit);
    edit->SetText(std::wstring{text});
    _pageHostDx->SyncTextInput(edit);
    _pageHostDx->Invalidate();
    _syncingDxBridgeBufferEdit = false;
    return edit->GetText() == text && fileOperations->crossFsBridgeBufferSizeKB == parsedValue;
}
#endif
