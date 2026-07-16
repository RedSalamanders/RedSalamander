#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <Windows.h>

#include "RedConfigureUiHelpers.h"

#include "resource.h"

#include <iterator>

namespace RedConfigure::Ui
{
std::wstring LoadAppString(HINSTANCE instance, UINT resourceId)
{
    wchar_t buffer[1024]{};
    const int length = ::LoadStringW(instance, resourceId, buffer, static_cast<int>(std::size(buffer)));
    if (length <= 0)
    {
        return {};
    }

    return std::wstring(buffer, static_cast<size_t>(length));
}

uint32_t ColorOrDefault(const RedConfigure::Themes::ThemePreviewModel& model, std::wstring_view key, uint32_t fallback)
{
    return model.GetEffectiveColor(key).value_or(fallback);
}

std::wstring PlaceholderStatusText(HINSTANCE instance, RedConfigure::Localization::PlaceholderStatus status)
{
    switch (status)
    {
        case RedConfigure::Localization::PlaceholderStatus::Ok: return LoadAppString(instance, IDS_REDCONFIGURE_STATUS_OK);
        case RedConfigure::Localization::PlaceholderStatus::BarePlaceholder: return LoadAppString(instance, IDS_REDCONFIGURE_STATUS_BARE_PLACEHOLDER);
        case RedConfigure::Localization::PlaceholderStatus::UnindexedFormatSpec: return LoadAppString(instance, IDS_REDCONFIGURE_STATUS_UNINDEXED_FORMAT);
        case RedConfigure::Localization::PlaceholderStatus::PrintfPlaceholder: return LoadAppString(instance, IDS_REDCONFIGURE_STATUS_PRINTF_FORMAT);
        case RedConfigure::Localization::PlaceholderStatus::InvalidPlaceholder: return LoadAppString(instance, IDS_REDCONFIGURE_STATUS_INVALID_PLACEHOLDER);
        case RedConfigure::Localization::PlaceholderStatus::PlaceholderMismatch: return LoadAppString(instance, IDS_REDCONFIGURE_STATUS_PLACEHOLDER_MISMATCH);
        default: return {};
    }
}

std::wstring LocalizableKindText(HINSTANCE instance, RedConfigure::Localization::RcLocalizableKind kind)
{
    switch (kind)
    {
        case RedConfigure::Localization::RcLocalizableKind::StringTable: return LoadAppString(instance, IDS_REDCONFIGURE_KIND_STRING);
        case RedConfigure::Localization::RcLocalizableKind::MenuPopup: return LoadAppString(instance, IDS_REDCONFIGURE_KIND_MENU_POPUP);
        case RedConfigure::Localization::RcLocalizableKind::MenuItem: return LoadAppString(instance, IDS_REDCONFIGURE_KIND_MENU_ITEM);
        case RedConfigure::Localization::RcLocalizableKind::DialogCaption: return LoadAppString(instance, IDS_REDCONFIGURE_KIND_DIALOG_CAPTION);
        case RedConfigure::Localization::RcLocalizableKind::DialogControl: return LoadAppString(instance, IDS_REDCONFIGURE_KIND_DIALOG_CONTROL);
        default: return {};
    }
}
} // namespace RedConfigure::Ui
